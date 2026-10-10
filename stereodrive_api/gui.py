"""Standalone StereoDrive USB controller. Starts disconnected; simulation by default."""
import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import queue
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
API_REPO = Path(__file__).resolve().parents[1]
if str(API_REPO) not in sys.path: sys.path.insert(0,str(API_REPO))
from stereodrive_api import StereoDrive, Calibration
from stereodrive_api.controller import StateStore
from stereodrive_api.protocol import AXES as CHANNELS, BACKLASH, SIGNS, STEPS
from stereodrive_api.transport import DEVICE_SERIAL
AXES = ('AP','ML','DV')

DATA = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'StereoDriveUSBController'

class Worker:
    def __init__(self, output, live, scales, states, new_reference, calibration=None,
                 speed=2, allow_dv=False, allow_piston=False, allow_drill=False):
        self.output=output; self.live=live; self.scales=scales; self.states=states
        self.new_reference=new_reference; self.calibration=calibration
        self.speed=speed; self.allow_dv=allow_dv; self.allow_piston=allow_piston; self.allow_drill=allow_drill
        self.commands=queue.Queue(); self.cancel=threading.Event()
        self.drive=None; self.session=None; self.stop_thread=None
        self.thread=threading.Thread(target=self.run,daemon=True)

    def update(self):
        p=self.drive.position()
        self.output.put(('position',dict(mm=p['axes_mm'],raw=p['raw_counts'],
            backlash=p['backlash_counts'],piston=p['piston_nl_estimate'],calibrated=p['calibrated'])))
        self.output.put(('drill_state',self.drive.drill_state()))
        self.output.put(('moving_state',self.drive.is_moving()))

    def stop_async(self):
        self.cancel.set()
        if self.stop_thread and self.stop_thread.is_alive(): return
        def stop():
            try:
                if self.drive: self.drive.stop()
                self.output.put(('stopped','Stop completed; drill OFF when enabled.'))
            except Exception as exc:self.output.put(('fault','Stop: '+str(exc)))
        self.stop_thread=threading.Thread(target=stop,daemon=True);self.stop_thread.start()

    def run(self):
        try:
            DATA.mkdir(parents=True,exist_ok=True)
            self.drive=StereoDrive(simulate=not self.live,
                state_path=DATA/('api-state-live.json' if self.live else 'api-state-simulation.json'),
                counts_per_mm={a:self.scales[a] for a in AXES},piston_counts_per_nl=self.scales['PISTON'],
                calibration=self.calibration,speed_mm_s=self.speed,allow_dv=self.allow_dv,
                allow_piston=self.allow_piston,allow_drill=self.allow_drill)
            self.drive.connect(verified_backlash=self.states,new_reference=self.new_reference)
            self.session=self.drive._session  # Diagnostic compatibility only; commands use the public API.
            self.output.put(('connected','Direct USB via stereodrive_api | '+self.session.transport.description));self.update()
            next_poll=time.monotonic()+.5
            while True:
                try:kind,value=self.commands.get(timeout=.05)
                except queue.Empty:kind,value=None,None
                if kind=='close':break
                try:
                    if kind in ('move','injector','drill'):
                        if self.cancel.is_set():raise InterruptedError('Stop requested before command')
                        done=threading.Event()
                        def watch():
                            while not done.wait(.2):
                                try:self.output.put(('moving_state',self.drive.is_moving()))
                                except Exception as exc:
                                    self.output.put(('notice','Motion query: '+str(exc)));return
                        watcher=threading.Thread(target=watch,daemon=True)
                        if kind in ('move','injector'):watcher.start()
                        try:
                            if kind=='move':self.drive.move_mm(value[0],value[1]*value[2])
                            elif kind=='injector':self.drive.piston_step(*value)
                            elif value:self.drive.drill_on()
                            else:self.drive.drill_off()
                        finally:
                            done.set()
                            if watcher.is_alive():watcher.join(3)
                        self.update();self.output.put(('arrived','Requested state verified; idle.'))
                        next_poll=time.monotonic()+.5
                    elif time.monotonic()>=next_poll and not self.session.fault:
                        self.update();next_poll=time.monotonic()+.5
                except ValueError as exc:self.output.put(('rejected',str(exc)))
                except Exception as exc:
                    self.output.put(('fault',str(exc)))
        except Exception as exc:self.output.put(('failed',str(exc)))
        finally:
            if self.stop_thread:self.stop_thread.join(5)
            if self.drive:
                try:self.drive.close()
                except Exception as exc:self.output.put(('notice','Disconnect: '+str(exc)))
            self.output.put(('disconnected',None))

class App:
    def __init__(self, root, live=False):
        self.root = root; self.worker = None; self.output = queue.Queue()
        self.connected = False; self.busy = False; self.faulted = False; self.closing = False
        self.root.title('StereoDrive | USB API Controller')
        self.root.geometry('1000x980'); self.root.minsize(960, 900)
        self.root.configure(bg='#edf1f7')
        style = ttk.Style(); style.theme_use('clam')
        style.configure('TFrame', background='#edf1f7')
        style.configure('TLabel', background='#edf1f7', foreground='#1e293b', font=('Segoe UI', 10))
        style.configure('Title.TLabel', font=('Segoe UI', 22, 'bold'))
        style.configure('Axis.TLabel', font=('Segoe UI', 17, 'bold'))
        style.configure('Position.TLabel', font=('Consolas', 23, 'bold'), foreground='#17636b')
        style.configure('TButton', font=('Segoe UI', 11), padding=8)
        style.configure('Arrow.TButton', font=('Segoe UI', 16, 'bold'), padding=8)
        self.mode = tk.StringVar(value='Live USB' if live else 'Simulation')
        self.confirm = tk.BooleanVar(value=False)
        self.allow_dv=tk.BooleanVar(value=False);self.allow_piston=tk.BooleanVar(value=False);self.allow_drill=tk.BooleanVar(value=False)
        self.speed=tk.StringVar(value='2');self.calibration=None
        self.calibration_label=tk.StringVar(value='Zero calibration missing: movement disabled')
        self.piston_position=tk.StringVar(value='Piston: -- nL');self.piston_step=tk.StringVar(value='10')
        self.drill_status=tk.StringVar(value='Drill power: unknown');self.motion_status=tk.StringVar(value='Moving: unknown')
        self.status = tk.StringVar(value='Disconnected · choose a mode, then connect.')
        self.connection = tk.StringVar(value=f'USB 0483:5743 · serial {DEVICE_SERIAL} · port detected automatically')
        self.mm = {}; self.raw = {}; self.steps = {}; self.last = {}; self.scale = {}; self.arrows = []
        outer = ttk.Frame(root, padding=14); outer.pack(fill='both', expand=True)
        ttk.Label(outer, text='StereoDrive USB controller', style='Title.TLabel').pack(anchor='w')
        ttk.Label(outer, text='AP / ML / DV | piston steps | drill power | shared stereodrive_api').pack(anchor='w', pady=(2, 15))
        bar = ttk.Frame(outer); bar.pack(fill='x')
        self.modebox = ttk.Combobox(bar, textvariable=self.mode, values=['Simulation', 'Live USB'], state='readonly', width=14)
        self.modebox.pack(side='left')
        self.modebox.bind('<<ComboboxSelected>>', lambda _: self.load_setup())
        self.connect_button = ttk.Button(bar, text='Connect / restore', command=self.connect)
        self.connect_button.pack(side='left', padx=8)
        self.disconnect_button = ttk.Button(bar, text='Disconnect', command=self.disconnect, state='disabled')
        self.disconnect_button.pack(side='left')
        ttk.Label(outer, textvariable=self.connection, wraplength=790).pack(anchor='w', pady=8)
        cards = ttk.Frame(outer); cards.pack(fill='x', pady=10)
        for i, axis in enumerate(AXES):
            card = ttk.LabelFrame(cards, text=axis, padding=8); card.grid(row=0, column=i, sticky='nsew', padx=(0, 10 if i < 2 else 0))
            cards.columnconfigure(i, weight=1)
            self.mm[axis] = tk.StringVar(value='— mm'); self.raw[axis] = tk.StringVar(value='Motor: —')
            ttk.Label(card, textvariable=self.mm[axis], style='Position.TLabel').pack(pady=(2, 3))
            ttk.Label(card, textvariable=self.raw[axis]).pack()
            ttk.Label(card, text='Controller estimate | relative or calibrated').pack(pady=(2, 8))
            row = ttk.Frame(card); row.pack(fill='x')
            for col, (direction, symbol) in enumerate(((-1, '◀ −'), (1, '+ ▶'))):
                row.columnconfigure(col, weight=1, uniform='arrows')
                button = ttk.Button(row, text=symbol, width=5, style='Arrow.TButton',
                    command=lambda a=axis, d=direction: self.move(a, d), state='disabled')
                button.grid(row=0, column=col, sticky='ew', padx=2); self.arrows.append(button)
            self.steps[axis] = tk.StringVar(value='0.01')
            row = ttk.Frame(card); row.pack(pady=(10, 2))
            ttk.Label(row, text='Step (mm)').pack(side='left', padx=(0, 7))
            ttk.Combobox(row, textvariable=self.steps[axis], values=[f'{s:g}' for s in STEPS],
                         width=6, state='readonly').pack(side='left')
        ttk.Label(outer, text='± signs mean increasing/decreasing mechanical coordinates. Axis travel: 0–40 mm; piston operating range: 500–4500 nL (nominal capacity 0–5000 nL).').pack(anchor='w')
        extra=ttk.LabelFrame(outer,text='Injection / piston steps and drill',padding=8);extra.pack(fill='x',pady=8)
        ttk.Label(extra,textvariable=self.piston_position).grid(row=0,column=0,padx=5)
        ttk.Combobox(extra,textvariable=self.piston_step,values=['10','20','50','100'],state='readonly',width=6).grid(row=0,column=1)
        self.piston_buttons=[]
        for i,d in enumerate(('up','down')):
            b=ttk.Button(extra,text='Piston '+d,command=lambda direction=d:self.inject(direction),state='disabled');b.grid(row=0,column=2+i,padx=5);self.piston_buttons.append(b)
        ttk.Label(extra,textvariable=self.drill_status).grid(row=1,column=0,columnspan=2,sticky='w',pady=5)
        self.drill_on_button=ttk.Button(extra,text='Drill ON',command=lambda:self.drill(True),state='disabled');self.drill_on_button.grid(row=1,column=2,padx=5)
        self.drill_off_button=ttk.Button(extra,text='Drill OFF',command=lambda:self.drill(False),state='disabled');self.drill_off_button.grid(row=1,column=3,padx=5)
        ttk.Label(extra,textvariable=self.motion_status).grid(row=0,column=4,padx=10)
        ttk.Label(extra,text='Piston steps are motor-derived nL; reversal may aspirate. Power OFF does not prove zero RPM.').grid(row=2,column=0,columnspan=5,sticky='w')

        settings = ttk.LabelFrame(outer, text='Connection calibration and initial direction history', padding=10)
        settings.pack(fill='x', pady=12)
        self.settings_controls = []
        for i, axis in enumerate(AXES):
            ttk.Label(settings, text=axis).grid(row=0, column=i*3, padx=5)
            self.scale[axis] = tk.StringVar(value='5225')
            entry = ttk.Entry(settings, textvariable=self.scale[axis], width=8)
            entry.grid(row=0, column=i*3+1); self.settings_controls.append(entry)
            self.last[axis] = tk.StringVar(value='+ mm')
            combo = ttk.Combobox(settings, textvariable=self.last[axis], values=['+ mm', '- mm'], state='readonly', width=7)
            combo.grid(row=0, column=i*3+2, padx=(4, 13)); self.settings_controls.append(combo)
        ttk.Label(settings, text='Counts/mm (5225 is a capture estimate) · direction of the last completed move on each axis',
                  wraplength=780).grid(row=1, column=0, columnspan=9, sticky='w', pady=7)
        self.confirmbox = ttk.Checkbutton(settings, text='Calibration / last directions verified; bench clear, drill off, physical Stop accessible', variable=self.confirm)
        self.confirmbox.grid(row=2, column=0, columnspan=9, sticky='w')
        self.new_button = ttk.Button(settings, text='Connect with new verified reference', command=lambda: self.connect(True))
        self.new_button.grid(row=3, column=0, columnspan=9, sticky='w', pady=(6, 0))
        self.scale['PISTON']=tk.StringVar(value='161.36');self.last['PISTON']=tk.StringVar(value='down')
        ttk.Label(settings,text='Piston counts/nL').grid(row=4,column=0,columnspan=2,sticky='w')
        e=ttk.Entry(settings,textvariable=self.scale['PISTON'],width=8);e.grid(row=4,column=2);self.settings_controls.append(e)
        c=ttk.Combobox(settings,textvariable=self.last['PISTON'],values=['up','down'],state='readonly',width=7);c.grid(row=4,column=3);self.settings_controls.append(c)
        ttk.Label(settings,text='Axis mm/s').grid(row=4,column=4)
        c=ttk.Combobox(settings,textvariable=self.speed,values=['1','2'],state='readonly',width=4);c.grid(row=4,column=5);self.settings_controls.append(c)
        for i,(name,var) in enumerate((('Enable DV',self.allow_dv),('Injector setup verified',self.allow_piston),('Enable drill ON',self.allow_drill))):
            c=ttk.Checkbutton(settings,text=name,variable=var);c.grid(row=5,column=i*3,columnspan=3,sticky='w');self.settings_controls.append(c)
        self.calibration_button=ttk.Button(settings,text='Load calibration JSON',command=self.load_calibration);self.calibration_button.grid(row=6,column=0,columnspan=3,sticky='w')
        self.settings_controls.append(self.calibration_button)
        ttk.Label(settings,textvariable=self.calibration_label,wraplength=600).grid(row=6,column=3,columnspan=6,sticky='w')

        footer = ttk.Frame(outer); footer.pack(fill='x', pady=(0, 8))
        self.stop_button = tk.Button(footer, text='STOP  ·  Esc', bg='#b42338', fg='white',
            activebackground='#8f1c2d', activeforeground='white', font=('Segoe UI', 14, 'bold'),
            relief='flat', padx=25, pady=9, command=self.stop, state='disabled')
        self.stop_button.pack(side='left')
        ttk.Label(footer, textvariable=self.status, wraplength=530).pack(side='left', padx=14)
        self.console = tk.Text(outer, height=3, bg='#132337', fg='#dbe7f5', relief='flat',
                               font=('Consolas', 9), padx=10, pady=8, state='disabled')
        self.console.pack(fill='both', expand=True)
        ttk.Label(outer, text=f'State and UTC logs: {DATA}', wraplength=790).pack(anchor='w', pady=(6, 0))
        root.bind('<Escape>', lambda _: self.stop())
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.load_setup()
        root.after(75, self.process_events)

    def load_setup(self):
        self.confirm.set(False)
        path = DATA / ('api-state-live.json' if self.mode.get() == 'Live USB' else 'api-state-simulation.json')
        try:
            setup=DATA/('api-setup-live.json' if self.mode.get()=='Live USB' else 'api-setup-simulation.json')
            if setup.exists():
                cfg=json.loads(setup.read_text(encoding='utf-8'))
                self.calibration=Calibration(**cfg['calibration']) if cfg.get('calibration') else None
                self.calibration_label.set('Saved absolute calibration' if self.calibration else 'Zero calibration missing: movement disabled')
                self.speed.set(str(cfg.get('speed_mm_s',2)))
            else:
                self.calibration=None;self.calibration_label.set('Zero calibration missing: movement disabled')
            data = StateStore(path).read()
            if data:
                for a in AXES:
                    self.scale[a].set(f'{data["scales"][a]:g}')
                    rawsign = 1 if data['backlash_state'][a] else -1
                    self.last[a].set('+ mm' if rawsign*SIGNS[a] > 0 else '- mm')
                self.scale['PISTON'].set(f'{data["scales"]["PISTON"]:g}')
                self.last['PISTON'].set('up' if data['backlash_state']['PISTON'] else 'down')
        except Exception as exc:
            self.note('Saved setup could not be loaded: ' + str(exc))

    def note(self, text):
        self.console.configure(state='normal')
        self.console.insert('end', datetime.now().strftime('%H:%M:%S') + '  ' + text + '\n')
        self.console.see('end'); self.console.configure(state='disabled')

    def controls(self):
        moving = self.connected and self.calibration is not None and not self.busy and not self.faulted
        for b in self.piston_buttons:b.configure(state='normal' if moving and self.allow_piston.get() else 'disabled')
        self.drill_on_button.configure(state='normal' if moving and self.allow_drill.get() else 'disabled')
        self.drill_off_button.configure(state='normal' if self.connected and not self.busy else 'disabled')
        for i,b in enumerate(self.arrows): b.configure(state='normal' if moving and (i<4 or self.allow_dv.get()) else 'disabled')
        enabled = self.worker is None
        for b in (self.connect_button, self.new_button): b.configure(state='normal' if enabled else 'disabled')
        self.disconnect_button.configure(state='normal' if self.worker else 'disabled')
        self.modebox.configure(state='readonly' if enabled else 'disabled')
        self.confirmbox.configure(state='normal' if enabled else 'disabled')
        for widget in self.settings_controls:
            widget.configure(state=('readonly' if isinstance(widget, ttk.Combobox) else 'normal') if enabled else 'disabled')
        self.stop_button.configure(state='normal' if self.connected else 'disabled')

    def connect(self, new_reference=False):
        if self.worker: return
        if self.calibration is None:
            messagebox.showwarning('Measured zero required','Load measured AP/ML/DV zero counts and piston anchor calibration before connecting for movement. Do not use GUI Bregma as a substitute.');return
        live = self.mode.get() == 'Live USB'
        if live and not self.confirm.get():
            messagebox.showinfo('Connection setup', 'Verify calibration, last completed directions and clear supervised bench, then check the setup box.'); return
        try: scales = {a: float(self.scale[a].get()) for a in CHANNELS}
        except ValueError: messagebox.showerror('Calibration', 'Enter a numeric counts/mm value for each axis.'); return
        if live and new_reference and not messagebox.askyesno('New verified reference',
                'Replace the saved position reference and backlash state using the verified directions above?'):
            return
        states = {a: BACKLASH[a] if SIGNS[a]*(1 if self.last[a].get() == '+ mm' else -1) > 0 else 0 for a in AXES}
        from dataclasses import asdict
        DATA.mkdir(parents=True,exist_ok=True)
        (DATA/('api-setup-live.json' if live else 'api-setup-simulation.json')).write_text(json.dumps({'calibration':asdict(self.calibration) if self.calibration else None,'speed_mm_s':int(self.speed.get())}),encoding='utf-8')
        states['PISTON']=BACKLASH['PISTON'] if self.last['PISTON'].get()=='up' else 0
        self.busy = True; self.faulted = False; self.status.set('Connecting and verifying idle positions…')
        self.worker = Worker(self.output, live, scales, states, new_reference,self.calibration,int(self.speed.get()),self.allow_dv.get(),self.allow_piston.get(),self.allow_drill.get()); self.controls(); self.worker.thread.start()

    def load_calibration(self):
        if self.worker:return
        filename=filedialog.askopenfilename(title='Absolute count calibration',filetypes=[('JSON','*.json')])
        if not filename:return
        try:
            value=json.loads(Path(filename).read_text(encoding='utf-8'))
            self.calibration=Calibration(**value)
            self.calibration.reference({a:float(self.scale[a].get()) for a in CHANNELS})
            self.calibration_label.set('Absolute anchors: '+Path(filename).name)
        except Exception as exc:
            self.calibration=None;messagebox.showerror('Calibration',str(exc))

    def inject(self,direction):
        if not self.connected or self.busy or self.faulted or not self.allow_piston.get():return
        self.busy=True;self.worker.cancel.clear();self.controls()
        self.status.set(f'Piston {direction} {self.piston_step.get()} nL')
        self.worker.commands.put(('injector',(direction,int(self.piston_step.get()))))

    def drill(self,enabled):
        if not self.connected or self.busy or (enabled and (self.faulted or not self.allow_drill.get())):return
        self.busy=True;self.worker.cancel.clear();self.controls()
        self.status.set('Setting drill '+('ON' if enabled else 'OFF'))
        self.worker.commands.put(('drill',enabled))

    def move(self, axis, direction):
        if not self.connected or self.busy or self.faulted or (axis=='DV' and not self.allow_dv.get()): return
        step = float(self.steps[axis].get())
        self.busy = True; self.worker.cancel.clear(); self.controls()
        self.status.set(f'Moving {axis} {direction*step:+g} mm…')
        self.note(self.status.get()); self.worker.commands.put(('move', (axis, direction, step)))

    def stop(self):
        if not self.worker or not self.connected: return
        self.busy=True;self.controls();self.worker.stop_async()
        self.status.set('Stop requested…'); self.note('Stop requested. Interrupted motion invalidates saved backlash state.')

    def disconnect(self):
        if not self.worker: return
        self.busy = True; self.controls()
        self.worker.stop_async(); self.worker.commands.put(('close', None))
        self.status.set('Disconnecting…')

    def close(self):
        self.closing = True
        if self.worker: self.disconnect()
        else: self.root.destroy()

    def process_events(self):
        while True:
            try: event, value = self.output.get_nowait()
            except queue.Empty: break
            if event == 'connected':
                self.connected = True; self.busy = False
                self.connection.set(value); self.status.set('Connected · verified idle position.'); self.note(value)
            elif event == 'position':
                for a in AXES:
                    self.mm[a].set(f'{value["mm"][a]:+.3f} mm')
                    self.raw[a].set(f'Motor: {value["raw"][a]}  ·  B: {value["backlash"][a]}')
                self.piston_position.set(f'Piston: {value["piston"]:.2f} nL | raw {value["raw"]["PISTON"]} | B {value["backlash"]["PISTON"]}')
            elif event == 'drill_state':self.drill_status.set('Drill power: '+('ON' if value else 'OFF'))
            elif event == 'moving_state':self.motion_status.set('Moving: '+('yes' if value else 'no'))
            elif event == 'raw':
                for a in AXES: self.raw[a].set(f'Motor: {value[a]} · moving / verifying')
            elif event in ('arrived', 'rejected'):
                self.busy = False; self.status.set(value); self.note(value)
            elif event == 'stopped':
                self.busy=False
                self.status.set(value if not self.faulted else 'Stopped · state verification required.'); self.note(value)
            elif event in ('fault', 'failed'):
                self.busy = False; self.faulted = True; self.status.set(value); self.note('FAULT: ' + value)
            elif event == 'notice': self.note(value)
            elif event == 'disconnected':
                self.worker = None; self.connected = False; self.busy = False
                if not self.faulted: self.status.set('Disconnected · verified state saved.')
                if self.closing: self.root.destroy(); return
            self.controls()
        self.root.after(75, self.process_events)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--live', action='store_true', help='Preselect Live USB; never auto-connect or auto-move.')
    args = parser.parse_args()
    root = tk.Tk()
    # One instance protects both the state files and command ownership.
    guard = None
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
        kernel.CreateMutexW.restype = wintypes.HANDLE
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        guard = kernel.CreateMutexW(None, False, 'Local\\StereoDriveUSBController')
        if not guard or ctypes.get_last_error() == 183:
            messagebox.showerror('Already open', 'The USB Axis Controller is already running, or its instance lock is unavailable.')
            if guard: kernel.CloseHandle(guard)
            root.destroy(); return
    try:
        App(root, live=args.live); root.mainloop()
    finally:
        if guard: kernel.CloseHandle(guard)

if __name__ == '__main__': main()

"""Rebuild a scoped evidence archive; requires installed Wireshark TShark.

Only reads completed captures. Does not capture or contact hardware.
"""
import argparse, hashlib, json, shutil, subprocess
from pathlib import Path
from decode_usb import read_capture, TSHARK

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def scrub(x):
    if isinstance(x,dict):return {k:scrub(v) for k,v in x.items() if not any(s in k.lower() for s in ('token','authorization','password'))}
    if isinstance(x,list):return [scrub(v) for v in x]
    return x

def commands(x):
    out=[]
    if isinstance(x,dict):
        if 'command_id' in x and 'kind' in x:out.append(x)
        for v in x.values():out+=commands(v)
    elif isinstance(x,list):
        for v in x:out+=commands(v)
    return out

def build(source,target):
    index=[]
    for original in sorted(source.rglob('*.pcap')):
        relative=original.relative_to(source); folder=target/'captures'/relative.parent/original.stem
        folder.mkdir(parents=True,exist_ok=True);capture=folder/'capture.pcapng'
        subprocess.run([TSHARK,'-r',str(original),'-Y',
            'usb.bus_id == 2 && (usb.device_address == 3 || usb.device_address == 7)',
            '-w',str(capture)],check=True,capture_output=True)
        rows=read_capture(capture)
        (folder/'decoded.json').write_text(json.dumps(rows,indent=2),encoding='utf-8')
        native=[];notes=[]
        for name in ('request.json','client-results.json','native-completed.json','experiment-kind.txt'):
            f=original.parent/name
            if not f.exists():continue
            if f.suffix=='.json':
                obj=scrub(json.loads(f.read_text(encoding='utf-8-sig')))
                (folder/name).write_text(json.dumps(obj,indent=2),encoding='utf-8')
                native+=commands(obj)
            else:
                text=f.read_text(encoding='utf-8-sig');(folder/name).write_text(text,encoding='utf-8');notes.append(text)
        native=list({json.dumps(c,sort_keys=True):c for c in native}.values())
        syringe_result=None
        series=original.parent.parent/'results.json'
        if series.exists():
            for result in json.loads(series.read_text(encoding='utf-8')):
                if Path(result.get('folder','')).name==original.parent.name:
                    syringe_result=result;break
        targets=[r for r in rows if r['kind']=='target'];endpoints={}
        for selector in (64,80,96,112):
            ss=[r for r in rows if r['kind']=='status' and r.get('selector')==selector]
            if ss:endpoints[hex(selector)]=dict(first_raw=ss[0]['position_raw'],last_raw=ss[-1]['position_raw'],last_moving=ss[-1]['observed_motion_byte'],samples=len(ss),last_utc=ss[-1]['utc'])
        drill=[r for r in rows if r['hex'] in ('af1101','af1100')]
        discoveries=[]
        if not targets:discoveries.append('No motor target packet recorded; not evidence of a completed requested move.')
        for t in targets:
            b=bytes.fromhex(t['hex']);discoveries.append(f"Selector {hex(t['selector'])} absolute raw target {t['target_raw']}; profile words {[int.from_bytes(b[i:i+2],'little') for i in (7,9,11)]}; tail {b[13:].hex()}.")
        if syringe_result:
            if syringe_result['passive']:discoveries.append('Passive idle observation only; no piston movement requested.')
            else:discoveries.append(f"Piston {syringe_result['direction']} {syringe_result['volume_nl']} nL: delta {syringe_result['delta_raw']} counts; last two samples independently idle at endpoint before next leg.")
        if drill:discoveries.append('Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.')
        name=str(relative).replace('\\','/')
        historical=[]
        if 'axis-series/' in name:historical.append('Historical duplicate-click incident: two outward targets and one reversal; native return faulted. Not a successful return.')
        if 'goto' in name.lower():historical.append('Historical GoTo investigation; no target packets is not arrival evidence. Inspect retained client result.')
        if 'injector-down-010-4e4a44cd' in name:historical.append('Native completion preceded final USB idle; capture ends moving. Later separate passive capture verified original count -8572 idle. Do not claim final offset.')
        if 'injector-down-010-360616e1' in name:historical.append('PASSIVE follow-up, despite directory name: original piston count -8572 verified idle; no reverse request.')
        if 'gui08-ml002' in name:historical.append('Historical incomplete ML attempt; do not combine with later successful gui09 trial.')
        metadata=dict(schema_version=1,source_relative_path=name,source_sha256=sha(original),archive_sha256=sha(capture),
          scope='Filtered bus 2, device addresses 3 and 7 only; originals remain outside repo',
          utc_start=rows[0]['utc'] if rows else None,utc_end=rows[-1]['utc'] if rows else None,
          native_commands=native,command_recovery='Native request bodies retained where available; otherwise raw targets are the observed commands; intended mm/nL cannot be inferred solely from filenames.',
          target_packets=targets,drill_writes=drill,endpoints=endpoints,syringe_result=syringe_result,
          discoveries=discoveries,historical_notes=historical,operator_notes=notes,
          limitations=['Counts are controller-reported, not independent physical displacement or fluid-volume measurements.',
                      'Filtered archive frame numbers may differ from source reports; match UTC and payload.',
                      'Native completed state alone is not proof of motor idle.'])
        (folder/'metadata.json').write_text(json.dumps(metadata,indent=2),encoding='utf-8')
        text=['# '+name,'','Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.','',
              '## Commanded','',json.dumps(native,indent=2) if native else 'Native command body unavailable; exact observed USB targets are listed below.','',
              '## Discovered','']+['- '+x for x in discoveries+historical]+['','## Final observed counts','',json.dumps(endpoints,indent=2),'',
              'Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.']
        (folder/'metadata.md').write_text('\n'.join(text),encoding='utf-8')
        index.append(dict(capture=str(capture.relative_to(target)).replace('\\','/'),metadata=str((folder/'metadata.json').relative_to(target)).replace('\\','/'),source=name,sha256=sha(capture),motor_targets=len(targets)))
    (target/'manifest.json').write_text(json.dumps(index,indent=2),encoding='utf-8')
    (target/'README.md').write_text(f"# Captured data archive\n\n{len(index)} completed capture files, filtered to authorized bus 2 devices 3 and 7. Each folder contains capture.pcapng, decoded.json, metadata.json and metadata.md, plus available scrubbed request/result evidence.\n\nmanifest.json indexes all files with SHA-256. Metadata records observed command packets, endpoint counts, discoveries and known incident/passive labels. Native requested physical units are only claimed when retained request/result evidence supports them. Filtered archive frame numbers can differ from older notes; correlate UTC and bytes.\n\nOpen capture.pcapng in Wireshark. Rebuild with `py -3 stereodrive_api/scripts/build_archive.py --source PATH_TO_LOGS`; this reads saved files only and needs TShark. Logs from the original native experiment predate this API. Archive inclusion is not evidence that the new API was hardware-tested.\n",encoding='utf-8')
    print(f'Archived {len(index)} captures',flush=True)

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--source',type=Path,required=True);ap.add_argument('--target',type=Path,default=Path(__file__).resolve().parents[1]/'data');args=ap.parse_args()
    args.target.mkdir(parents=True,exist_ok=True);build(args.source,args.target)

"""Reproduce report figures from archived decoded captures. Requires matplotlib."""
import json
from pathlib import Path
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parents[1]
DATA=ROOT/'data';OUT=ROOT/'figures';OUT.mkdir(exist_ok=True)
manifest=json.loads((DATA/'manifest.json').read_text())
plt.rcParams.update({'font.size':10,'axes.spines.top':False,'axes.spines.right':False})
def entry(part):return next(e for e in manifest if part in e['source'])
def rows(part):return json.loads((DATA/entry(part)['metadata']).with_name('decoded.json').read_text())
def save(fig,name):
    fig.savefig(OUT/(name+'.png'),dpi=180,bbox_inches='tight')
    fig.savefig(OUT/(name+'.pdf'),bbox_inches='tight');plt.close(fig)

r=rows('wide-series/AP-nudge-100-2eb263ff')
a,b=[x for x in r if x['kind']=='target']
ss=[x for x in r if x['kind']=='status' and x.get('selector')==64 and a['frame']<x['frame']<b['frame']]
pre=[x for x in r if x['kind']=='status' and x.get('selector')==64 and x['frame']<a['frame']][-1]
t=[(x['device_clock']-ss[0]['device_clock'])/1000 for x in ss];y=[(pre['position_raw']-x['position_raw'])/5225 for x in ss]
v=[(yb-ya)/(tb-ta) for ya,yb,ta,tb in zip(y,y[1:],t,t[1:])];mid=[(a+b)/2 for a,b in zip(t,t[1:])]
fig,ax=plt.subplots(2,1,figsize=(9,6),sharex=True,layout='constrained')
ax[0].plot(t,y,'o-',color='#2463aa');ax[0].set_ylabel('Motor travel equivalent (mm)');ax[0].set_title('AP +1 mm: reported position and interval speed (3 mm/s request)')
ax[1].plot(mid,v,'o-',color='#d66a20');ax[1].axhline(3,ls='--',color='gray');ax[1].set_ylabel('Interval speed equivalent (mm/s)');ax[1].set_xlabel('Seconds since first moving status')
for x in ax:x.grid(alpha=.2);x.set_xlim(-.02,.85)
fig.supxlabel('5225 counts/mm; travel includes backlash. Terminal interval can include idle dwell.',fontsize=9);save(fig,'axis-ramp')

fig,ax=plt.subplots(figsize=(8,4.5),layout='constrained')
for label,part,selector in [('AP','speed-unknown-01/AP',64),('ML','speed-unknown-01/ML',80),('DV','speed-unknown-01/DV',96)]:
 r=rows(part);a,b=[x for x in r if x['kind']=='target'];ss=[x for x in r if x['kind']=='status' and x.get('selector')==selector and a['frame']<x['frame']<b['frame']]
 t=[(x['device_clock']-ss[0]['device_clock'])/1000 for x in ss]
 v=[abs(b['position_raw']-a['position_raw'])/((b['device_clock']-a['device_clock'])/1000)/5225 for a,b in zip(ss,ss[1:])]
 ax.plot([(a+b)/2 for a,b in zip(t,t[1:])],v,'o-',label=label)
ax.axhline(1,color='gray',ls='--',label='Requested 1 mm/s');ax.set(xlabel='Seconds since first post-target status',ylabel='Interval speed equivalent (mm/s)',title='0.5 mm moves at the same requested speed');ax.legend();ax.grid(alpha=.2);ax.set_xlim(0,1);save(fig,'axis-speed-comparison')

fig,ax=plt.subplots(figsize=(7,4),layout='constrained');ax.plot([1,2,3],[653,1306,1959],'o-',color='#2463aa');ax.set(xlabel='Native configured speed (mm/s)',ylabel='Captured uint16 speed field',title='Observed axis speed-field scaling');ax.set_xticks([1,2,3]);ax.grid(alpha=.2);save(fig,'speed-field')

fig,axes=plt.subplots(1,2,figsize=(11,4.4),layout='constrained')
axes[0].plot([10,20,50,100],[7439,9052,13893,21961],'o',label='First move after reversal')
axes[0].plot([10,20,50,100],[1613,3228,8068,16137],'s',label='Continued same direction')
x=[0,100];axes[0].plot(x,[5825+161.36*v for v in x],ls='--');axes[0].plot(x,[161.36*v for v in x],ls='--');axes[0].set(xlabel='Requested piston step (nL)',ylabel='Absolute motor count increment',title='Volume term plus reversal correction');axes[0].legend()
for volume,color in [(50,'#2463aa'),(100,'#d66a20')]:
 folder=DATA/entry(f'injector-triples-05/injector-up-{volume:03d}-')['metadata'];series=[]
 for e in manifest:
  if 'injector-triples-05/injector-up-'+f'{volume:03d}-' in e['source']:
   m=json.loads((DATA/e['metadata']).read_text());series.append((m['utc_start'],m['syringe_result']['delta_raw']))
 series.sort();axes[1].plot([1,2,3],[v for _,v in series],'o-',color=color,label=f'{volume} nL')
axes[1].set(xlabel='Consecutive move number',ylabel='Motor count increment',title='Three moves in one direction');axes[1].set_xticks([1,2,3]);axes[1].legend()
for ax in axes:ax.grid(alpha=.2)
save(fig,'piston-backlash')

r=rows('drill-manual-01/drill.pcap');switches=[x for x in r if x['hex'] in ('af1101','af1100')]
def sec(x):return datetime.fromisoformat(x['utc'].replace('Z','+00:00')).timestamp()
t0=sec(switches[0]);s=[x for x in r if x['direction']=='IN' and len(x['hex'])==28 and x['hex'].startswith('12') and x['hex'][2:4] in ('00','01')]
fig,ax=plt.subplots(figsize=(9,3.7),layout='constrained');ax.step([sec(x)-t0 for x in s],[int(x['hex'][2:4],16) for x in s],where='post',label='Reported power state')
for x in switches:ax.axvline(sec(x)-t0,ls='--',color='#d66a20');ax.text(sec(x)-t0+.15,.5,'ON write' if x['hex']=='af1101' else 'OFF write')
ax.set(xlabel='Seconds relative to operator ON write',ylabel='Reported drill power',title='Drill write versus delayed reported power state',yticks=[0,1],ylim=(-.15,1.2),xlim=(-1,17));ax.grid(alpha=.2);save(fig,'drill-state')
print('Generated 5 figures and PDFs')

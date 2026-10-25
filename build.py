"""Slides only (no docx/CSV). Builds the TechVestVentures Trafficking Graphs site from focus/immigration (pptx + docx + CSVs)."""
import os,re,sys,shutil,glob
sys.path.insert(0,'/workspace/risk-watcher/graphs-site');sys.path.insert(0,'/workspace/risk-watcher');os.chdir(os.path.dirname(os.path.abspath(__file__)))
from common import *
import common,weekly_lib as WL
from pptx import Presentation
T="TechVestVentures Trafficking Graphs";F="/workspace/risk-watcher/focus/trafficking"
P=Presentation(glob.glob(F+"/TVV Focus Human Trafficking Graphs Collection 2026-10-18 mh2.pptx")[0])
shutil.rmtree("issues",ignore_errors=True);os.makedirs("issues/img");pass
secs={};heads={};summ={};allsrc=[];order=[]
for i,s in enumerate(P.slides,1):
    tx=[sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    ttl=s.shapes.title.text.strip() if s.shapes.title else (tx[0].split("\n")[0] if tx else "")
    n=s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    m=re.match(r"(Q\d)\. (.*)",ttl)
    if m:
        heads[m[1]]=m[2];summ[m[1]]=[t for t in tx if t!=ttl];continue
    for u in re.findall(r"https://[^\s)\]>,]+",n):
        if u not in allsrc: allsrc.append(u)
    m=re.match(r"(Q\d): (.*)",ttl)
    if not m: continue
    pics=[sh for sh in s.shapes if sh.shape_type==13];pic=max(pics,key=lambda x:x.width*x.height)
    f=dict(l.split(": ",1) for l in n.splitlines() if ": " in l and not l.startswith("Source"))
    urls=re.findall(r"^Source: (https://\S+)",n,re.M)
    title=f.get("Title",m[2]);key=m[1]
    cap=f"{f.get('Publisher','')}. {f.get('What it counts','')}".strip(". ")
    lst=secs.setdefault(key,[]);fn=f"{key}-{len(lst)+1:02d}.webp";save_img(pic.image.blob,"issues/img/"+fn)
    lst.append(figure("img/"+fn,title,cap,urls[:2]))
LABELS={"Q1":"Global Scale","Q2":"US Case Counts","Q3":"Labor and Visas","Q4":"Migrant Victims","Q5":"Child Migrants","Q6":"Data Quality"}
tabs=[]
def w(k,lab,body): open(f"issues/{k}.html","w").write(page(T,k,body));tabs.append((k,lab))
for q in sorted(k for k in secs if k!="TVV"):
    w(q,f"{q} {LABELS[q]}",f'<h2 class="desk">{q}. {E(heads.get(q,""))}</h2>'+"\n".join(secs[q]))
common.CSS+=WL.GRID_CSS
WG=WL.add_tabs("Trafficking",w,E,save_img,start=7,note="Weekly Focus graphs: new charts are added each week from primary data. Each caption names its sources; numbers are not copied from third-party chart images.")
allsrc_w=WL.sources("Trafficking")
for u in allsrc_w:
    if u not in allsrc: allsrc.append(u)
w("sources","Sources",f'<h2 class="desk">Sources ({len(allsrc)} links)</h2><ul class="src">'+"".join(f'<li><a href="{E(u)}">{E(u)}</a></li>' for u in allsrc)+"</ul>")
index(T,tabs,"Sections",cur="traffic");print({k:len(v) for k,v in secs.items()},len(allsrc))

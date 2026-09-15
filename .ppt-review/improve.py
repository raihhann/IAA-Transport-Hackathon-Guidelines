from pathlib import Path
from copy import deepcopy
from lxml import etree as E
import zipfile, io
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'output'
OUT.mkdir(exist_ok=True)
with zipfile.ZipFile(ROOT/'Presentation.pptx') as z:
    data = {n:z.read(n) for n in z.namelist()}
NS = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main', 'p':'http://schemas.openxmlformats.org/presentationml/2006/main', 'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
def el(tag, **attrs):
    prefix,local=tag.split(':');return E.Element('{'+NS[prefix]+'}'+local,{k:str(v) for k,v in attrs.items()})
def sub(parent,tag,**attrs):
    node=el(tag,**attrs);parent.append(node);return node
def read(n):return E.fromstring(data[n])
def save(n,r):data[n]=E.tostring(r,xml_declaration=True,encoding='UTF-8',standalone=True)
def slide(i):return read(f'ppt/slides/slide{i}.xml')
def write(i,s):save(f'ppt/slides/slide{i}.xml',s)
def tree(s):return s.find('p:cSld/p:spTree',NS)
def shape(s,id):return s.find(f'.//p:sp[p:nvSpPr/p:cNvPr[@id="{id}"]]',NS)
def findshape(s,id):return next(v for v in s.findall('.//p:sp',NS) if v.find('p:nvSpPr/p:cNvPr',NS).get('id')==str(id))
def pos(sp,x,y,w,h):
    pr=sp.find('p:spPr',NS);xf=pr.find('a:xfrm',NS)
    if xf is not None:pr.remove(xf)
    xf=el('a:xfrm');pr.insert(0,xf)
    sub(xf,'a:off',x=round(x*12700),y=round(y*12700));sub(xf,'a:ext',cx=round(w*12700),cy=round(h*12700))
def replace(s,id,text):
    sp=findshape(s,id);ts=sp.findall('.//a:t',NS);ts[0].text=text
    for t in ts[1:]:t.text=''
    return sp
def textbox(s,text,x,y,w,h,size=22,bold=False,color='30414C',font='Arial',link=None):
    ident=max([int(n.get('id')) for n in s.findall('.//p:cNvPr',NS)]+[0])+1
    sp=sub(tree(s),'p:sp');nv=sub(sp,'p:nvSpPr');sub(nv,'p:cNvPr',id=ident,name='Text '+str(ident));sub(nv,'p:cNvSpPr',txBox='1');sub(nv,'p:nvPr');sub(sp,'p:spPr');pos(sp,x,y,w,h)
    body=sub(sp,'p:txBody');sub(body,'a:bodyPr',wrap='square',lIns='0',rIns='0',tIns='0',bIns='0',anchor='t');sub(body,'a:lstStyle')
    for line in text.split('\n'):
        p=sub(body,'a:p');pp=sub(p,'a:pPr');sub(pp,'a:buNone');ln=sub(pp,'a:lnSpc');sub(ln,'a:spcPct',val='110000');after=sub(pp,'a:spcAft');sub(after,'a:spcPts',val='600')
        r=sub(p,'a:r');rp=sub(r,'a:rPr',lang='en-US',sz=round(size*100),b='1' if bold else '0');fill=sub(rp,'a:solidFill');sub(fill,'a:srgbClr',val=color);sub(rp,'a:latin',typeface=font)
        if link:sub(rp,'a:hlinkClick',**{'{'+NS['r']+'}id':link})
        sub(r,'a:t').text=line
    return sp
def blank(i,title):
    s=slide(3);t=tree(s)
    for c in list(t):
        if c.tag not in ['{'+NS['p']+'}nvGrpSpPr','{'+NS['p']+'}grpSpPr']:t.remove(c)
    textbox(s,title,41,34,740,38,28,color='003F63',font='Montserrat Medium')
    rels=E.Element('{'+REL+'}Relationships',nsmap={None:REL});E.SubElement(rels,'{'+REL+'}Relationship',Id='rId1',Type=NS['r']+'/slideLayout',Target='../slideLayouts/slideLayout9.xml');save(f'ppt/slides/_rels/slide{i}.xml.rels',rels)
    return s
def relation(i,typ,target,external=False):
    n=f'ppt/slides/_rels/slide{i}.xml.rels';r=read(n);rid='rId'+str(max([int(c.get('Id')[3:]) for c in r]+[0])+1)
    a=dict(Id=rid,Type=NS['r']+'/'+typ,Target=target)
    if external:a['TargetMode']='External'
    E.SubElement(r,'{'+REL+'}Relationship',**a);save(n,r);return rid
def linktext(s,i,label,url,x,y,w=400,size=20):
    return textbox(s,label,x,y,w,34,size,color='003F63',link=relation(i,'hyperlink',url,True))
def picture(s,i,filename,x,y,w,h,crop=None):
    im=Image.open(ROOT/filename)
    if crop:im=im.crop(crop)
    b=io.BytesIO();im.save(b,format='PNG');name='review_'+str(i)+'_'+str(len(data))+'.png';data['ppt/media/'+name]=b.getvalue()
    rid=relation(i,'image','../media/'+name)
    iw,ih=im.size;factor=min(w/iw,h/ih);pw,ph=iw*factor,ih*factor;x+=(w-pw)/2;y+=(h-ph)/2
    ident=max([int(n.get('id')) for n in s.findall('.//p:cNvPr',NS)]+[0])+1
    pic=sub(tree(s),'p:pic');nv=sub(pic,'p:nvPicPr');sub(nv,'p:cNvPr',id=ident,name=filename,descr='Screenshot from project README; example names may differ from event assignments.');locks=sub(nv,'p:cNvPicPr');sub(locks,'a:picLocks',noChangeAspect='1');sub(nv,'p:nvPr')
    fill=sub(pic,'p:blipFill');sub(fill,'a:blip',**{'{'+NS['r']+'}embed':rid});st=sub(fill,'a:stretch');sub(st,'a:fillRect')
    pr=sub(pic,'p:spPr');xf=sub(pr,'a:xfrm');sub(xf,'a:off',x=round(x*12700),y=round(y*12700));sub(xf,'a:ext',cx=round(pw*12700),cy=round(ph*12700));g=sub(pr,'a:prstGeom',prst='rect');sub(g,'a:avLst')
def notes(i,text):
    n=f'ppt/notesSlides/notesSlide{i}.xml'
    r=E.Element('{'+NS['p']+'}notes',nsmap={'a':NS['a'],'p':NS['p'],'r':NS['r']});cs=sub(r,'p:cSld');st=sub(cs,'p:spTree');nv=sub(st,'p:nvGrpSpPr');sub(nv,'p:cNvPr',id='1',name='');sub(nv,'p:cNvGrpSpPr');sub(nv,'p:nvPr');sub(st,'p:grpSpPr')
    sp=sub(st,'p:sp');nv=sub(sp,'p:nvSpPr');sub(nv,'p:cNvPr',id='2',name='Notes');sub(nv,'p:cNvSpPr');np=sub(nv,'p:nvPr');sub(np,'p:ph',type='body',idx='1');sub(sp,'p:spPr');tx=sub(sp,'p:txBody');sub(tx,'a:bodyPr');sub(tx,'a:lstStyle');p=sub(tx,'a:p');rr=sub(p,'a:r');sub(rr,'a:t').text=text;save(n,r)
    rels=E.Element('{'+REL+'}Relationships',nsmap={None:REL});E.SubElement(rels,'{'+REL+'}Relationship',Id='rId1',Type=NS['r']+'/notesMaster',Target='../notesMasters/notesMaster1.xml');E.SubElement(rels,'{'+REL+'}Relationship',Id='rId2',Type=NS['r']+'/slide',Target=f'../slides/slide{i}.xml');save(f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels',rels)
    sr=read(f'ppt/slides/_rels/slide{i}.xml.rels')
    if not any(c.get('Type').endswith('/notesSlide') for c in sr):relation(i,'notesSlide',f'../notesSlides/notesSlide{i}.xml')

REPO='https://github.com/raihhann/IAA-Transport-Hackathon-Guidelines'
PLAY='https://playground.digital.auto/'
EXAMPLE='https://playground.digital.auto/model/67d2eb3e880fe100272d033e/library/prototype/6a9684f804417b08156ceb7a/code'

# Cover: preserve campus image, translucent panel and master branding.
s=slide(1);sp=findshape(s,5);body=sp.find('p:txBody',NS)
for c in list(body):
    if c.tag=='{'+NS['a']+'}p':body.remove(c)
temp=blank(99,'unused')
new=textbox(temp,'IAA LCV Hackathon 2026\nBring-Your-Own-Device\nExternal devices. Useful vehicle applications.\ndigital.auto and prototype.club\nChris Cheng, Ferdinand-Steinbeis-Institut',0,0,1,1,20,color='003F63',font='Montserrat Medium')
for j,p in enumerate(new.findall('p:txBody/a:p',NS)):
    p.find('a:r/a:rPr',NS).set('sz',str(3000 if j==0 else 1800));body.append(deepcopy(p))
write(1,s);data.pop('ppt/slides/_rels/slide99.xml.rels',None)

s=slide(4);replace(s,3,'The challenge');sp=replace(s,40,'RESPONSE');pos(sp,460,324,170,38);replace(s,41,'VALUE');replace(s,45,'What should the device detect?');replace(s,47,'Why does it matter?');write(4,s)

s=blank(5,'Three starting points for your idea')
for x,title,desc,items in [(55,'Mobile workshop','Equipment and\noperational security','Tool detection\nMissing-item checks\nLoading completeness'),(355,'Dangerous goods','Load conditions and\nsafety awareness','Gas-cylinder detection\nStorage alerts\nSafety-equipment checks'),(655,'Animal transport','Live cargo and\ntransit conditions','Behavior monitoring\nUnusual movement\nPosition monitoring')]:
    textbox(s,title,x,140,255,40,24,True,'003F63');textbox(s,desc,x,202,255,76,21);textbox(s,items,x,310,255,120,21)
textbox(s,'Choose a user, an observable event and a useful response.',55,462,850,35,22,True,'003F63');write(5,s)

s=blank(6,'A complete device-to-vehicle loop')
textbox(s,'Technical reference: driver drowsiness',50,113,840,36,24,True)
for y,title,desc in [(174,'1  Camera input','Observe the driver.'),(242,'2  Application logic','Detect a drowsiness event.'),(310,'3  Vehicle response','Trigger headlights and horn.')]:
    textbox(s,title,55,y,360,30,21,True,'003F63');textbox(s,desc,55,y+30,360,32,20)
picture(s,6,'Imgs/dashboard.png',430,177,480,218)
textbox(s,'4  Visible evidence',460,410,400,30,21,True,'003F63');textbox(s,'Inspect Dashboard values and logs.',460,441,430,30,20)
linktext(s,6,'Open the example prototype',EXAMPLE,55,455,380,20);write(6,s)

s=slide(7);replace(s,10,'Extra sensors');sp=replace(s,13,'Radar\nBosch gas sensor\nPM sensor');pos(sp,634,192,260,125)
replace(s,15,'Limited availability. Request from a mentor; teams handle setup.')
textbox(s,'Camera access: URLs are pending in the guide.\nAsk a mentor before testing your connection.',53,400,520,70,19,color='003F63');write(7,s)

s=blank(8,'Building with digital.auto')
textbox(s,'Develop your application primarily in the Playground.',50,116,865,40,24)
for x,y,num,title,desc in [(55,188,'01','Choose the vehicle','IAA Hannover Hackathon 2026'),(510,188,'02','Create your team project','Choose a multi-files template.\nUse your team name.'),(55,330,'03','Build in SDV Code','Manage source files and dependencies.'),(510,330,'04','Configure your runtime','Get the assigned name from a mentor.\nSelect it before running your app.')]:
    textbox(s,num,x,y,60,40,30,True,'ED5025',font='Consolas');textbox(s,title,x+65,y+2,345,32,22,True,'003F63');textbox(s,desc,x+65,y+43,345,77,20)
linktext(s,8,'playground.digital.auto',PLAY,55,473,350,19);textbox(s,'Setup screenshots: appendix, slides 13–14',510,473,405,30,18);write(8,s)

s=blank(10,'What teams should bring back')
textbox(s,'Core outcome',55,135,430,40,26,True,'003F63');textbox(s,'Working prototype or proof of concept\nClear use-case description\nAccessible source code\nShort functional demo\nExplanation of the BYOD integration',55,193,455,220,22)
textbox(s,'Useful additions',550,135,355,40,26,True,'003F63');textbox(s,'Architecture diagram\nAdditional device concepts\nFuture vehicle integration ideas\nPlatform feedback',550,193,350,195,22)
textbox(s,'To be confirmed by the organizers',55,433,840,30,20,True,'003F63');textbox(s,'Submission deadline and location, source-code format, pitch and demo duration.',55,468,850,48,18);write(10,s)

s=blank(11,'What makes a strong outcome')
textbox(s,'Indicative evaluation dimensions',55,114,840,36,22)
for x,y,title,desc in [(55,183,'Problem and value','Problem relevance\nUser value'),(510,183,'Solution concept','Creativity\nBYOD integration concept'),(55,321,'Working implementation','Technical implementation\nUse of digital.auto\nDemo quality'),(510,321,'Future integration','Scalability to additional devices')]:
    textbox(s,title,x,y,405,35,24,True,'003F63');textbox(s,desc,x,y+44,405,115,22)
textbox(s,'Final judging criteria and any weighting are to be confirmed.',55,479,850,30,18);write(11,s)

s=blank(12,'Start your prototype')
textbox(s,'Choose a challenge and create your team project.',55,121,835,80,30,True,'003F63')
textbox(s,'Resources',55,223,400,37,25,True,'003F63')
linktext(s,12,'Hackathon guide and setup instructions',REPO,55,278,455,20)
linktext(s,12,'digital.auto Playground',PLAY,55,328,440,20)
linktext(s,12,'Example camera prototype',EXAMPLE,55,378,440,20)
linktext(s,12,'digital.auto documentation','https://docs.digital.auto/',55,428,440,20)
textbox(s,'Technical support',550,223,350,37,25,True,'003F63')
linktext(s,12,'Chris Cheng','https://www.linkedin.com/in/xiangwei-cheng/',550,278,350,22)
linktext(s,12,'Mohammed Raihan Soniwala','https://www.linkedin.com/in/raihan-edin/',550,328,355,21)
textbox(s,'Ask a mentor for your Runtime name,\ncamera access and optional sensors.',550,387,355,90,20)
textbox(s,'Select a resource or contact name to open its link.',55,487,850,25,16);write(12,s)

s=blank(13,'Appendix: create your team project')
textbox(s,'1  Select the hackathon vehicle',55,119,430,34,22,True,'003F63');textbox(s,'IAA Hannover Hackathon 2026',55,163,445,35,21)
textbox(s,'2  Name the prototype after your team',55,227,440,65,22,True,'003F63');textbox(s,'Choose Python Multiple Files (Beta)\nfor the Python reference workflow.',55,298,435,65,21)
picture(s,13,'Imgs/prototype name.png',520,128,390,330,crop=(925,490,1955,1205))
textbox(s,'3  Open SDV Code',55,394,435,35,22,True,'003F63');textbox(s,'Manage your source and requirements files.',55,435,435,55,21)
textbox(s,'Example interface. Use the event vehicle and your own team name.',55,494,850,26,16);write(13,s)

s=blank(14,'Appendix: connect your assigned runtime')
textbox(s,'1  Open Terminal, then Add Runtime',55,117,430,60,23,True,'003F63');textbox(s,'The Terminal toggle is at the bottom right.',55,177,430,45,19)
picture(s,14,'Imgs/runtime box.png',55,234,430,155,crop=(1790,190,2876,540))
textbox(s,'2  Enter the name supplied by a mentor',530,117,380,60,23,True,'003F63');textbox(s,'Use the pre-filled Runtime- prefix once.',530,177,380,45,19)
picture(s,14,'Imgs/runtime name.png',530,234,380,155,crop=(690,305,2195,608))
textbox(s,'3  Select the runtime and run your app',55,414,850,34,23,True,'003F63');textbox(s,'Confirm the connection in Terminal, then inspect Dashboard values and logs.',55,458,850,52,20);write(14,s)

for i in range(1,15):
    source='Source: project README.md and original Presentation.pptx. '
    extra={1:'Opening: digital.auto and prototype.club invite teams to develop BYOD solutions for lightweight commercial vehicles. A camera is the first reference device.',4:'Explain the sequence: user, event, response, value.',6:'The driver-drowsiness application is a technical reference, not an additional challenge. Dashboard screenshot: Imgs/dashboard.png. Example code: '+EXAMPLE,7:'The source README lists three cameras, but stream and snapshot URLs remain placeholders. Do not imply endpoints have been validated. Optional sensors are limited and require teams to handle setup.',8:'Use the exact model name in the README. Participants develop primarily on digital.auto Playground. Runtime names come from the supervisor.',9:'These are future BYOD design questions; do not imply discovery or permissions are implemented in the current platform.',10:'Organizers must confirm submission deadline, submission location, code format, pitch duration and demo duration before the event.',11:'The README describes these as potential evaluation dimensions. Final rules and weights are not provided.',12:'Repository URL derived from the local Git origin. Links are navigation resources, not confirmation of platform availability. Support contacts are from the README.',13:'Source screenshot: Imgs/prototype name.png, cropped to the New Prototype dialog. The screenshot opens with Python Single File selected. Choose Python Multi Files (Beta) from the visible dropdown for this Python reference workflow. Example names differ from the event assignment.',14:'Sources: Imgs/runtime box.png and Imgs/runtime name.png, cropped to the controls. Ask the supervisor for the assigned runtime name. Do not add Runtime- twice.'}.get(i,'')
    notes(i,source+extra)

# Register the two appended slides and all notes parts.
pr=read('ppt/presentation.xml');ids=pr.find('p:sldIdLst',NS);rels=read('ppt/_rels/presentation.xml.rels');maxrid=max(int(c.get('Id')[3:]) for c in rels)
for k,i in enumerate([13,14],1):
    rid='rId'+str(maxrid+k);E.SubElement(rels,'{'+REL+'}Relationship',Id=rid,Type=NS['r']+'/slide',Target=f'slides/slide{i}.xml');sub(ids,'p:sldId',id=max(int(c.get('id')) for c in ids)+1,**{'{'+NS['r']+'}id':rid})
save('ppt/presentation.xml',pr);save('ppt/_rels/presentation.xml.rels',rels)
ct=read('[Content_Types].xml');existing={c.get('PartName') for c in ct}
for i in range(1,15):
    for part,typ in [(f'/ppt/slides/slide{i}.xml','slide'),(f'/ppt/notesSlides/notesSlide{i}.xml','notesSlide')]:
        if part not in existing:E.SubElement(ct,'{'+CT+'}Override',PartName=part,ContentType='application/vnd.openxmlformats-officedocument.presentationml.'+typ+'+xml')
save('[Content_Types].xml',ct)
# English page marker throughout the reused template.
for n in list(data):
    if n.startswith('ppt/slideMasters/') and n.endswith('.xml'):
        r=read(n)
        for t in r.findall('.//a:t',NS):
            if t.text and 'Seite' in t.text:t.text=t.text.replace('Seite','Slide')
        save(n,r)
if 'docProps/app.xml' in data:
    r=read('docProps/app.xml')
    for e in r:
        if E.QName(e).localname=='Slides':e.text='14'
        if E.QName(e).localname=='Notes':e.text='14'
    save('docProps/app.xml',r)
candidate=ROOT/'.ppt-review/candidate.pptx'
with zipfile.ZipFile(candidate,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in data.items():z.writestr(n,b)
print(candidate)

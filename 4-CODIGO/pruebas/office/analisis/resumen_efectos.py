import json,sys,re
from lxml import etree
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
base=sys.argv[1]; info=json.load(open(sys.argv[2]))
filtro=sys.argv[3] if len(sys.argv)>3 else None
for x in info['diapositivas']:
    if filtro and not re.search(filtro, x['efecto']): continue
    t=etree.parse(f"{base}/ppt/slides/slide{x['diapo']}.xml")
    ef=[c for c in t.iter(P+'cTn') if c.get('presetClass')]
    for e in ef:
        def desc(n,ind=0):
            out=[]
            for c in n:
                if not isinstance(c.tag,str): continue
                ln=etree.QName(c).localname
                if ln in ('cBhvr',):
                    ctn=c.find(P+'cTn'); tg=c.find(P+'tgtEl'); an=[a.text for a in c.iter(P+'attrName')]
                    out.append(' '*ind+f"cTn{dict((k,v) for k,v in ctn.attrib.items() if k!='id')} delay={[d.get('delay') for d in ctn.iter(P+'cond')]} attr={an} {dict(c.attrib)}")
                    continue
                if ln in('tgtEl','stCondLst'): continue
                out.append(' '*ind+ln+' '+str({k:v[:70] for k,v in c.attrib.items()}))
                out+=desc(c,ind+2)
            return out
        print(f"--- {x['diapo']} {x['efecto']}{' (salida)' if x.get('salida') else ''}: id={e.get('presetID')} {e.get('presetClass')} sub={e.get('presetSubtype')} {dict((k,v) for k,v in e.attrib.items() if k not in ('id','presetID','presetClass','presetSubtype','grpId','nodeType'))}")
        print('\n'.join(desc(e.find(P+'childTnLst'),2)))
        it=e.find(P+'iterate')
        if it is not None: print('  iterate',dict(it.attrib),[ (etree.QName(c).localname,dict(c.attrib)) for c in it])
        sub=e.find(P+'subTnLst')
        if sub is not None: print('  subTnLst', etree.tostring(sub)[:300])

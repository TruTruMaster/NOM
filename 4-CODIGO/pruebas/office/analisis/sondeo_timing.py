import re,glob,collections,sys
from lxml import etree
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
tags=collections.Counter(); attrs=collections.Counter(); filters=collections.Counter(); fmlas=collections.Counter(); vals=collections.Counter(); cattrs=collections.Counter(); calc=collections.Counter(); other=collections.Counter()
for f in glob.glob(sys.argv[1]+'/ppt/slides/slide*.xml'):
    x=etree.parse(f)
    tm=x.find('.//p:timing',ns)
    if tm is None: continue
    for el in tm.iter():
        if not isinstance(el.tag,str): continue
        ln=etree.QName(el).localname
        tags[ln]+=1
        if ln=='attrName': attrs[el.text]+=1
        if ln=='animEffect': filters[(el.get('filter'),el.get('transition'),el.get('prLst'))]+=1
        if ln=='tav':
            if el.get('fmla'): fmlas[el.get('fmla')]+=1
        if ln=='strVal': vals[el.get('val')]+=1
        if ln=='cTn':
            for k in el.attrib:
                if k not in ('id',): cattrs[k]+=1
        if ln in('anim',): calc[(el.get('calcmode'),el.get('valueType'), el.get('by'), el.get('from'), el.get('to'))]+=1
        if ln=='cBhvr':
            for k,v in el.attrib.items(): other[(k,v)]+=1
        if ln in ('animMotion','animRot','animScale','animClr','iterate','cmd','set','cond','tmPct','tmAbs','by','to','from','rCtr','hsl','rgb'):
            for k,v in el.attrib.items(): other[(ln,k,v[:40])]+=1
print('TAGS',dict(tags)); print('ATTRS',dict(attrs)); print('FILTERS',filters.most_common()); print('CTN ATTRS',dict(cattrs)); print('CALC',calc.most_common(30)); print('OTHER',other.most_common(120))
print('FMLAS',fmlas.most_common()); print('VALS',vals.most_common(200))

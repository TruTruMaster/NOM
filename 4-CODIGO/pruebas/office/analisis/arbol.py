import re,sys,xml.dom.minidom
def compact(f, etiqueta='p:timing'):
    t=open(f,encoding='utf8').read()
    m=re.search(r'<%s.*</%s>'%(etiqueta,etiqueta),t,re.S).group(0)
    d=xml.dom.minidom.parseString(m.replace('<'+etiqueta,'<'+etiqueta+' xmlns:p="p" xmlns:a="a" xmlns:p14="p14" xmlns:r="r" xmlns:p15="p15" xmlns:p159="p159"',1))
    def walk(n,ind=0):
        if n.nodeType==3 and n.data.strip(): print(' '*ind+'"'+n.data.strip()+'"'); return
        if n.nodeType!=1: return
        at=' '.join(f'{k}={v}' for k,v in n.attributes.items() if not k.startswith('xmlns'))
        name=n.tagName.replace('p:','')
        print(' '*ind+name+(' '+at if at else ''))
        for c in n.childNodes: walk(c,ind+1)
    walk(d.documentElement)
for f in sys.argv[1:]:
    print('=====',f); compact(f)

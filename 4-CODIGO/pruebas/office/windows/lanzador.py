# Corre un trabajo de Office dentro de la sesión de escritorio (la tarea programada lo lanza con pythonw).
# Lee trabajo.json {id, args}, ejecuta python con esos argumentos y deja salida/estado en trabajos\<id>.*
import json, os, subprocess, sys, time
base = os.path.dirname(os.path.abspath(__file__))
t = json.load(open(os.path.join(base, 'trabajo.json'), encoding='utf-8-sig'))
d = os.path.join(base, 'trabajos'); os.makedirs(d, exist_ok=True)
log = open(os.path.join(d, t['id'] + '.log'), 'w', encoding='utf8')
py = os.path.join(os.path.dirname(sys.executable), 'python.exe')
p = subprocess.run([py, '-X', 'utf8'] + t['args'], cwd=base, stdout=log, stderr=subprocess.STDOUT, creationflags=0x08000000)
log.close()
open(os.path.join(d, t['id'] + '.fin'), 'w').write(str(p.returncode))

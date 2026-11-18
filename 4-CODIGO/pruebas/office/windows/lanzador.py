# Corre un trabajo de Office dentro de la sesión de escritorio (la tarea programada lo lanza con pythonw).
# Lee trabajos\<id>.json {id, args} (el id llega como argumento; sin él, trabajo.json como antes), ejecuta python con
# esos argumentos y deja salida y estado en trabajos\<id>.log y .fin. Así varios trabajos pueden correr a la vez
# (la regresión en paralelo pide Office desde varias pruebas al mismo tiempo).
import json, os, subprocess, sys
base = os.path.dirname(os.path.abspath(__file__))
d = os.path.join(base, 'trabajos'); os.makedirs(d, exist_ok=True)
ruta = os.path.join(d, sys.argv[1] + '.json') if len(sys.argv) > 1 else os.path.join(base, 'trabajo.json')
t = json.load(open(ruta, encoding='utf-8-sig'))
log = open(os.path.join(d, t['id'] + '.log'), 'w', encoding='utf8')
py = os.path.join(os.path.dirname(sys.executable), 'python.exe')
p = subprocess.run([py, '-X', 'utf8'] + t['args'], cwd=base, stdout=log, stderr=subprocess.STDOUT, creationflags=0x08000000)
log.close()
open(os.path.join(d, t['id'] + '.fin'), 'w').write(str(p.returncode))

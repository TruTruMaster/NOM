# Análisis de animaciones contra PowerPoint

Scripts chicos para estudiar lo que hace PowerPoint en sus videos (`CreateVideo`) y en su XML. Sus datos de trabajo van en
`~/NOM/trabajo/animaciones-17nov/` (no se suben). Los videos y .pptx están en `~/.cache/nom-pruebas/office-real/archivos/`.
Se corren con `~/.cache/nom-pruebas/venv/bin/python` (numpy, PIL, lxml, ffmpeg).

- `mosaico.py VIDEO salida.png diapo:t1,t2 …` — cuadros de varias diapositivas en un mosaico.
- `cuadros_diapo.py VIDEO DIAPO prefijo t1,t2` — cuadros sueltos a tamaño completo.
- `perfil.py` — ancho del borde suave de los filtros (10–90 %).
- `trayecto.py VIDEO DIAPO salida.json` — centro de la figura azul en cada cuadro.
- `ajuste_camino.py`, `ajuste2.py`, `ajuste3.py`, `ajuste_acel.py` — qué curva de tiempo usa PowerPoint (inicio y fin
  suaves, rebote, recorrido de trayectorias).
- `letras.py`, `cajas_letras.py` — dónde va cada letra al volar o al hacer zoom (la caja de PowerPoint por letra).
- `depurar_anim.py ARCHIVO DIAPO T 'expr JS'` — abre el .pptx en NOM, prepara el motor y evalúa algo en el instante T.
- `resumen_efectos.py CARPETA_PPTX JSON filtro`, `arbol.py slideN.xml` — el XML de los efectos en corto.
- `quietos.py`, `sondeo_timing.py` — ratos quietos del video (inicios de transiciones) y lo que trae p:timing.

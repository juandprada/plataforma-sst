# Fuente del documento (no de la interfaz)

`arimo-400.woff2` / `arimo-700.woff2` — **Arimo**, de Google (Steve Matteson), bajo
[SIL Open Font License 1.1](OFL.txt), que **permite redistribuirla** en este repo
público.

## Por qué está aquí y no simplemente `font-family: Arial`

Los 20 `.docx` de origen usan Arial, así que el documento debe medirse como Arial. Pero
el PDF **lo genera el navegador del técnico con las fuentes de SU equipo**: en Windows,
macOS e iPhone hay Arial, pero en **Android no** — cae en Roboto, que es más angosta, y
el documento sale distinto del Word (líneas de firma cortas, otros cortes de línea).

Arimo es **métricamente compatible** con Arial: cada carácter mide exactamente lo mismo.
No es "parecida", es intercambiable, así que empaquetarla da el mismo resultado en todos
los dispositivos sin redistribuir Arial (propiedad de Monotype, no redistribuible).

Verificado con `fontTools` contra `C:\Windows\Fonts\arial.ttf` y `arialbd.ttf`: **0
caracteres con ancho distinto**, incluido el guion bajo (556/1000 en ambas), que era el
síntoma que destapó todo esto.

## Cómo se generaron

Desde la fuente variable oficial (`google/fonts`, `ofl/arimo/Arimo[wght].ttf`), instanciando
los pesos 400 y 700 y recortando a los caracteres que usan los documentos (latín +
acentos + símbolos de formulario como `□ • ° º`). De ~500 KB a ~7 KB cada uno:

```python
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.subset import Subsetter, Options

f = instantiateVariableFont(TTFont("Arimo[wght].ttf"), {"wght": 400}, updateFontNames=True)
s = Subsetter(options=Options(layout_features=["*"]))
s.populate(text=CARACTERES)   # ver tools/ o el commit que agregó esta carpeta
s.subset(f)
f.flavor = "woff2"
f.save("arimo-400.woff2")
```

**Si algún formato necesita un carácter que no esté en el subconjunto, no se verá**
(el navegador caería a otra fuente para ese glifo). Al agregar formatos con símbolos
raros, regenerar ampliando `CARACTERES`.

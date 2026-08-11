# Proyectos legislativos pendientes (Senado)

Navegador de **proyectos de ley pendientes destacados** que el Senado argentino
tiene en trámite (o que tramitan en Diputados y todavía no llegaron al
Senado). Es un artefacto HTML autocontenido (sin frameworks, sin backend):
se abre `proyectos_legislativos.html` directo en el navegador.

Es hermano de otro proyecto (`pliegos_judiciales/`, navegador de pliegos
judiciales del Poder Judicial con la misma estética) pero son independientes.

## Cómo funciona

1. **`PEN_y_CD.xlsx`**: planilla madre de proyectos de ley del Senado
   (columna `TIPO == 'PL'`). Trae fechas de giro a comisión, Orden del Día,
   sanción, etc.
2. **`fuentes/*.md`**: una nota por proyecto con resúmenes generados por
   NotebookLM/Gemini, en un formato de secciones semi-fijo.
3. `data_builder.py` cruza ambas fuentes y genera `proyectos_data.json`.
4. `build_html.py` embebe ese JSON (más las fuentes Montserrat) en
   `template.html` y produce `proyectos_legislativos.html`, el entregable
   final.

## Cómo correrlo

Requiere Python 3 con `pandas` y `openpyxl` instalados.

```bash
python data_builder.py && python build_html.py
```

Después abrí `proyectos_legislativos.html` con doble clic (no necesita
servidor).

## Cómo agregar proyectos nuevos

**Un `.md` de un proyecto que ya está en la planilla del Senado:**
1. Guardalo en `fuentes/` con un nombre descriptivo.
2. Si la línea `EXPEDIENTE:` del `.md` sigue alguno de los patrones que ya
   reconoce `parse_expediente_line` (en `md_parser.py`), no hace falta tocar
   nada más.
3. Si el identificador es atípico, agregá una entrada en
   `CLAVE_OVERRIDE_POR_ARCHIVO` en `data_builder.py` con
   `("ORIGEN", nro, "AAAA")`.
4. Corré `python data_builder.py` — el log avisa si el archivo quedó sin
   matchear o si hay Orden del Día con formato raro.
5. Si el título automático no queda bien, agregale uno corto en
   `TITULOS_CURADOS`.
6. Corré `python build_html.py` y revisá el resultado.

**Un `.md` de un proyecto que tramita en Diputados (no está en el Excel):**
1. Guardalo en `fuentes/`.
2. Agregá una entrada en `PROYECTOS_SIN_EXCEL` (diccionario en
   `data_builder.py`) con: expediente, `url` (`None` si no la tenés),
   `camara_actual` (normalmente `"Diputados"`), `categoria`
   (`"en_comision"` mientras no tenga Orden del Día ni sanción), `titulo`,
   `nombre_corto` y una lista de `faltantes`.
3. Corré `data_builder.py` — al final imprime un aviso con los proyectos
   manuales incompletos y qué le falta a cada uno.

**Actualizar el Excel** (nuevas fechas, Orden del Día, sanción, etc.):
reemplazá `PEN_y_CD.xlsx` y volvé a correr `data_builder.py`. Si un proyecto
pasa a tener `LEY – FECHA` con fecha real, se excluye automáticamente del
navegador (queda listado en el log bajo "Excluidos: ya sancionados como
ley").

Para más detalle sobre el esquema del Excel, el formato de las notas `.md`,
las reglas de categorización/badges y la limpieza de citas bibliográficas,
ver los comentarios en `data_builder.py` y `md_parser.py`.

## Estructura

```
├── PEN_y_CD.xlsx                 # planilla madre (Senado)
├── data_builder.py               # Excel + fuentes/*.md -> proyectos_data.json
├── md_parser.py                  # parser de las notas .md
├── build_html.py                 # proyectos_data.json + template.html -> proyectos_legislativos.html
├── template.html                 # plantilla con placeholders
├── proyectos_data.json           # generado, se versiona
├── proyectos_legislativos.html   # generado, se versiona (entregable final)
├── fonts/                        # Montserrat subset en woff2
└── fuentes/                      # notas .md por proyecto
```

## Qué no hacer

- No mezclar este repo con el de pliegos judiciales.
- No agregar frameworks (React, Tailwind, etc.) — es HTML/JS vanilla
  autocontenido.
- No inventar URLs de expedientes no confirmadas.
- No tocar la lógica de categorización/badges sin confirmar — está ajustada
  a pedidos específicos ya validados.

"""
Parser de las notas .md por proyecto (formato con secciones
EXPEDIENTE / RESUMEN_MENSAJE_PEN / RESUMEN_EXPEDIENTE / REUNIONES_COMISION /
ORDEN_DEL_DIA / SANCION). Ver data_builder.py para el uso.
"""
import re

SECTION_HEADERS = [
    "RESUMEN_MENSAJE_PEN",
    "RESUMEN_EXPEDIENTE",
    "REUNIONES_COMISION",
    "ORDEN_DEL_DIA",
    "SANCION",
]

SIN_INFO_RE = re.compile(r'^(sin informaci[oó]n|n/?a\b)', re.IGNORECASE)

# líneas de "sugerencia" que agregan algunas herramientas de IA al final del documento
AI_SUGGESTION_RE = re.compile(r'^[📈📊🔍✨🧭📌]')


# Cita al pie tipo NotebookLM/Gemini: uno o más números/rangos ("9, 17-19", "2-4", "40, 41")
# que aparecen pegados al final de una palabra real, no dentro de un decimal (22.426) ni
# de una referencia a artículo/ley (que casi siempre va precedida de "N°", "Nº", "#", "art.").
# Dos ramas:
#  - clusters de 2+ números (con guion o coma): se aceptan como cita también cuando siguen
#    de una conjunción corta ("y", "e", "o", "u", "ni"), patrón típico de listas de citas.
#  - un solo número: solo se saca si lo que sigue es puntuación de cierre fuerte, nunca una
#    conjunción/preposición (para no comerse cantidades reales tipo "de 5 a 10 provincias").
CITATION_TAIL_RE = re.compile(
    # nota: una coma solo cuenta como separador de cita si va seguida de espacio
    # ("9, 17-19"); una coma pegada a los dígitos es un decimal en español ("0,5%")
    r'(?<=[A-Za-zÁÉÍÓÚÑÜáéíóúñü)%])\s*(?:'
    r'\d{1,3}(?:\s*[-–]\s*\d{1,3}|,\s+\d{1,3})+(?=\.(?!\d)|[;)]|,\s|\s+(?:y|e|o|u|ni)\b|\s*$)'
    r'|'
    r'\d{1,3}(?=\.(?!\d)|[;)]|,\s|\s*$)'
    r')'
)

# si justo antes del número hay "artículo/inciso/capítulo/ley..." es una enumeración real
# (p. ej. "artículos 41, 42 y 59"), nunca una cita bibliográfica -> no se toca
CONTEXTO_PROTEGIDO_RE = re.compile(
    r'(art[ií]culos?|incisos?|cap[ií]tulos?|p[aá]rrafos?|puntos?|numerales?|leyes?|decretos?)'
    r'\.?\s*(n[°ºo]\.?)?\s*$',
    re.IGNORECASE,
)


def _quitar_citas(t):
    def reemplazar(m):
        precedente = t[:m.start()]
        if CONTEXTO_PROTEGIDO_RE.search(precedente):
            return m.group(0)  # es una enumeración real (artículos/incisos/leyes), no una cita
        return ''
    return CITATION_TAIL_RE.sub(reemplazar, t)


def strip_markup(text):
    if text is None:
        return ""
    t = text
    t = t.replace("\\.", ".").replace("\\-", "-").replace("\\_", "_")
    t = t.replace("**", "")
    t = t.replace("*", "")
    t = re.sub(r'\\+', '', t)
    # artefactos de citas tipo "Image 3, Image 4"
    t = re.sub(r'(?:,?\s*Image\s*\d+)+', '', t, flags=re.IGNORECASE)
    # números de referencia bibliográfica sueltos (ver CITATION_TAIL_RE arriba)
    t = _quitar_citas(t)
    t = re.sub(r'\s+', ' ', t).strip()
    t = re.sub(r'\s+([.,;:])', r'\1', t)
    t = re.sub(r'\(\s*\)', '', t)  # paréntesis que quedaron vacíos tras limpiar
    return t.strip()


# siglas/acrónimos conocidos que deben mantenerse en mayúscula al pasar de CAPS a oración
SIGLAS_CONOCIDAS = {
    "RIGI", "BCRA", "IPC", "IVA", "CABA", "AFIP", "ANSES", "RENABAP", "MERCOSUR",
    "PYMES", "PYME", "IGN", "CPCCN", "ONG", "UIF", "SRT", "ART", "PBI", "PEN", "CD",
    "PE", "OPS", "FAM",
}


def sentence_case(text):
    """Convierte texto en MAYÚSCULA SOSTENIDA a formato de oración en español (best-effort)."""
    if not text:
        return text
    if text != text.upper():
        return text  # ya tiene minúsculas, no lo toco

    words = text.split(" ")
    out = []
    for w in words:
        core = re.sub(r'[^A-ZÁÉÍÓÚÑ]', '', w)
        if core in SIGLAS_CONOCIDAS:
            out.append(w)
        else:
            out.append(w.lower())
    result = " ".join(out)
    if result:
        result = result[0].upper() + result[1:]
    return result


# notas meta ("esto no tiene datos") que a veces vienen como un párrafo entre paréntesis
# y arrancan con un único "*" de itálica -> el parser las toma por error como viñeta
META_SIN_DATOS_RE = re.compile(
    r'^\(.*(no se registra|no consta|no aplica|sin constancias|no cuenta con|'
    r'documento contiene únicamente|al no haberse|no existen firmas|no se emiti|'
    r'no ha(n)? (sido|alcanzado)|no se ha)',
    re.IGNORECASE,
)


def split_bullets(block):
    bullets = []
    for line in block.splitlines():
        line = line.strip()
        if line.startswith("*") or line.startswith("-"):
            bullets.append(strip_markup(line.lstrip("*- ").strip()))
        elif line and bullets:
            bullets[-1] = (bullets[-1] + " " + strip_markup(line)).strip()
    bullets = [
        b for b in bullets
        if b and not SIN_INFO_RE.match(b.strip("* ")) and not META_SIN_DATOS_RE.match(b.strip())
    ]
    return bullets


def extract_labeled_field(block, label, stop_labels):
    pattern = re.escape(label) + r':\s*(.*?)(?=' + "|".join(re.escape(s) + r':' for s in stop_labels) + r'|$)'
    m = re.search(pattern, block, re.IGNORECASE | re.DOTALL)
    if not m:
        return ""
    return strip_markup(m.group(1))


def strip_trailing_stray_citation(text):
    """Saca un número de cita suelto al final, cuando queda pegado justo
    después de una fecha u otro número real (p. ej. '17/06/2026 1' -> '17/06/2026')."""
    return re.sub(r'(?<=[\d/])\s+\d+(?:\s*[-–,]\s*\d+)*\s*$', '', text).strip()


def parse_reuniones(block):
    sub_labels = ["fecha", "comision", "resumen", "posturas_principales"]
    fecha = strip_trailing_stray_citation(extract_labeled_field(block, "fecha", sub_labels))
    comision = strip_trailing_stray_citation(extract_labeled_field(block, "comision", sub_labels))
    resumen = strip_trailing_stray_citation(extract_labeled_field(block, "resumen", sub_labels))
    # posturas_principales es la lista con viñetas que sigue a ese label
    m = re.search(r'posturas_principales:\s*(.*)', block, re.IGNORECASE | re.DOTALL)
    posturas = split_bullets(m.group(1)) if m else []

    if SIN_INFO_RE.match(fecha) or not fecha:
        return None
    return {
        "fecha": fecha,
        "comision": comision,
        "resumen": resumen,
        "posturas_principales": posturas,
    }


def parse_orden_del_dia(block):
    sub_labels = ["numero", "firmantes", "diferencias_clave_vs_proyecto_original"]
    numero = strip_trailing_stray_citation(extract_labeled_field(block, "numero", sub_labels))
    if SIN_INFO_RE.match(numero) or not numero:
        return None
    m = re.search(
        r'firmantes:\s*(.*?)(?=diferencias_clave_vs_proyecto_original:|$)',
        block, re.IGNORECASE | re.DOTALL
    )
    firmantes = split_bullets(m.group(1)) if m else []
    m2 = re.search(r'diferencias_clave_vs_proyecto_original:\s*(.*)', block, re.IGNORECASE | re.DOTALL)
    diferencias = split_bullets(m2.group(1)) if m2 else []
    return {"numero": numero, "firmantes": firmantes, "diferencias_clave": diferencias}


def parse_sancion(block):
    sub_labels = ["link_taquigrafica", "posturas_principales"]
    link = extract_labeled_field(block, "link_taquigrafica", sub_labels)
    m = re.search(r'posturas_principales:\s*(.*)', block, re.IGNORECASE | re.DOTALL)
    posturas = split_bullets(m.group(1)) if m else []
    link_ok = bool(link) and bool(re.match(r'^(https?://|www\.)', link.strip(), re.IGNORECASE))
    if not link_ok and not posturas:
        return None
    return {"link_taquigrafica": link if link_ok else None, "posturas_principales": posturas}


def parse_expediente_line(line):
    """Devuelve una lista de claves candidatas (origen, nro, anio) encontradas en la línea EXPEDIENTE."""
    line = strip_markup(line)
    claves = []
    for m in re.finditer(r'\b(PE|CD)-(\d+)\s*/\s*(\d{2,4})\b', line, re.IGNORECASE):
        origen, nro, anio = m.group(1).upper(), int(m.group(2)), m.group(3)
        anio = anio if len(anio) == 4 else "20" + anio
        claves.append((origen, nro, anio))
    for m in re.finditer(r'\b(\d+)-(PE|CD|P\.E\.)\.?-(\d{4})\b', line, re.IGNORECASE):
        nro, origen, anio = int(m.group(1)), m.group(2).upper().replace(".", ""), m.group(3)
        origen = "PE" if origen.startswith("P") else origen
        claves.append((origen, nro, anio))
    return claves


def parse_md_file(path):
    raw = path.read_text(encoding="utf-8")
    raw = raw.replace("\\_", "_")
    raw = "\n".join(l for l in raw.splitlines() if not AI_SUGGESTION_RE.match(l.strip()))

    m = re.search(r'EXPEDIENTE:\s*(.*)', raw)
    expediente_line = m.group(1).splitlines()[0] if m else ""
    claves = parse_expediente_line(expediente_line)

    sections = {}
    pattern = "(" + "|".join(SECTION_HEADERS) + r"):"
    parts = re.split(pattern, raw)
    # parts alterna: [texto_antes, header, texto, header, texto, ...]
    for i in range(1, len(parts), 2):
        header = parts[i]
        body = parts[i + 1] if i + 1 < len(parts) else ""
        sections[header] = body

    resumen_mensaje = split_bullets(sections.get("RESUMEN_MENSAJE_PEN", ""))
    resumen_expediente = split_bullets(sections.get("RESUMEN_EXPEDIENTE", ""))
    reunion = parse_reuniones(sections.get("REUNIONES_COMISION", ""))
    orden_del_dia = parse_orden_del_dia(sections.get("ORDEN_DEL_DIA", ""))
    sancion = parse_sancion(sections.get("SANCION", ""))

    return {
        "archivo": path.name,
        "expediente_claves": claves,
        "expediente_raw": strip_markup(expediente_line),
        "resumen_mensaje_pen": resumen_mensaje,
        "resumen_expediente": resumen_expediente,
        "reunion_comision": reunion,
        "orden_del_dia_md": orden_del_dia,
        "sancion": sancion,
    }

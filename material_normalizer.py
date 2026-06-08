"""
Quantia AI Comparador
"""

import re
from datetime import date

_UNIT_ALIASES = {
    "lt": ["lt", "l", "litro", "litros"],
    "gal": ["gal", "galon", "galon", "galones"],
    "cubeta": ["cubeta", "cubetas", "bote"],
    "kg": ["kg", "kilo", "kilogramo", "kilogramos"],
    "ton": ["ton", "tonelada", "toneladas", "t"],
    "pza": ["pza", "pieza", "pz", "pzas", "unidad"],
    "caja": ["caja", "cajas"],
    "m": ["m", "metro", "metros"],
    "m2": ["m2", "m²", "metro cuadrado", "metros cuadrados"],
    "m3": ["m3", "m³", "metro cubico", "metros cubicos"],
    "saco": ["saco", "sacos", "bulto", "bultos"],
    "rollo": ["rollo", "rollos"],
    "tramo": ["tramo", "tramos"],
    "lote": ["lote", "lotes"],
    "serv": ["serv", "servicio", "servicios"],
}

_STOPWORDS = {"de", "la", "el", "los", "las", "para", "con", "por", "del", "en", "y", "a", "o", "un", "una", "tipo", "incluye", "material", "materiales"}

def normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip().lower()
    repl = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n","°":" grados ","/":" ","-":" ","(":" ",")":" ",",":" ",":":" ",";":" ", '"':" pulgadas "}
    for a, b in repl.items():
        s = s.replace(a, b)
    s = re.sub(r"[^a-z0-9\. ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def normalize_unit(unit):
    u = normalize_text(unit)
    for key, values in _UNIT_ALIASES.items():
        if u in values:
            return key
    return u

def tokenize_material(text):
    return [t for t in normalize_text(text).split() if len(t) > 1 and t not in _STOPWORDS]

def normalize_material_description(description, unit=""):
    norm = normalize_text(description)
    return {
        "original": description or "",
        "normalized": norm,
        "tokens": tokenize_material(description),
        "unit_normalized": normalize_unit(unit),
        "query": " ".join(tokenize_material(description)[:10]),
    }

def conversion_factor(source_unit, market_unit, market_description=""):
    src = normalize_unit(source_unit)
    mkt = normalize_unit(market_unit)
    text = normalize_text(market_description)
    if not src or not mkt or src == mkt:
        return 1.0, "misma unidad" if src == mkt else "unidad insuficiente"
    pairs = {("kg", "ton"): 1 / 1000, ("ton", "kg"): 1000, ("lt", "gal"): 1 / 3.78541, ("gal", "lt"): 3.78541}
    if (src, mkt) in pairs:
        return pairs[(src, mkt)], f"conversion directa {src}<->{mkt}"
    m = re.search(r"(\d+(?:\.\d+)?)\s*(kg|kilogramo|kilogramos|lt|litro|litros|m|metro|metros)", text)
    if m and src in {"kg", "lt", "m"} and mkt in {"saco", "cubeta", "tramo", "rollo", "caja", "pza"}:
        return 1 / float(m.group(1)), "conversion por contenido declarado en fuente"
    return None, "no convertido: falta equivalencia suficiente"

def today_iso():
    return date.today().isoformat()

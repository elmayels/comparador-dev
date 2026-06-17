# Quantia Comparador APU

> Versión v0.2 profesional: ver `README_ANALISTAS_V02_4.md` para el flujo recomendado con base Neodata/APU, Resumen PU oficial común y comparación multi-proveedor.

# 🏗️ Comparador de Cotizaciones

Sistema para comparar automáticamente cotizaciones de obra de múltiples proveedores,
generando un comparativo Excel profesional con colores, ganadores por concepto y resumen ejecutivo.

---

## 📋 Contexto del proyecto

**Cliente:** Nestlé Servicios Industriales  
**Problema:** El equipo hacía comparativos manualmente en Excel, tomando horas o días. Este sistema lo hace en segundos.  
**Primer piloto:** Ampliación de Comedor, Veracruz
**Proveedores piloto:** PROV1 ($23.9M ganador), PROV2 ($25.3M), PROV3 ($42.2M)

---

## 🏛️ Arquitectura

```
Usuario (Retool)
     │  POST /comparar  (archivos Excel + metadata)
     ▼
Backend FastAPI → processor.py → openpyxl → comparativo.xlsx
```

- **Backend:** Python 3.11 + FastAPI + openpyxl  
- **Frontend:** Retool (gratis hasta 5 usuarios) — ver GUIA_IMPLEMENTACION_DEV.docx  
- **Hosting:** Railway.app (~$5/mes)  

---

## 📁 Estructura

```
nestle-comparador/
├── backend/
│   ├── main.py           # FastAPI — /comparar y /preview
│   ├── processor.py      # Motor de extracción + generación Excel ⭐
│   ├── requirements.txt
│   └── .env.example
├── frontend/
│   └── index.html        # UI standalone (mockup visual)
├── Dockerfile            # Python 3.11 + LibreOffice
├── railway.toml          # Deploy Railway automático
├── GUIA_IMPLEMENTACION_DEV.docx
└── README.md
```

---

## 🔌 Endpoints

### POST /comparar
Recibe archivos, devuelve `.xlsx` para descarga.

```bash
curl -X POST http://localhost:8000/comparar \
  -F "archivos=@CIOC.xlsx" \
  -F "archivos=@PROIINCSA.xlsx" \
  -F 'nombres=["Prov1","Prov2"]' \
  -F "proyecto=Ampliación Comedor Parish" \
  --output Comparativo.xlsx
```

### POST /preview  
Previsualización: conceptos detectados + total estimado por archivo, sin generar Excel.

---

## 🧠 Detección automática de formatos

3 estrategias en cascada en `processor.py`:

| # | Estrategia | Detecta | Ejemplo |
|---|-----------|---------|---------|
| 1 | Presupuesto estándar | Columnas fijas: Código, Desc, Uni, Cant, P.U., Importe | PROIINCSA/RAR |
| 2 | APU (fichas multi-fila) | Etiquetas `Clave:`, `Precio unitario:`, `Total` por concepto | CIOC |
| 3 | Búsqueda por headers | Cualquier tabla con columnas reconocibles | Formatos nuevos |

Maneja duplicados (mismo concepto en distintas partidas) agrupando y promediando P.U.

---

## 🚀 Setup local

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

---

## ☁️ Deploy Railway

1. Subir repo a GitHub  
2. railway.app → New Project → Deploy from GitHub  
3. Variable: `PORT=8000`  
4. Deploy (~5 min, instala LibreOffice)  
5. Copiar URL → configurar en Retool como Resource

---

## 🗺️ Roadmap

| Fase | Feature | Esfuerzo |
|------|---------|----------|
| ✅ Versión actual | API + detección automática + Excel comparativo | Listo |
| ✅ Versión actual | Retool UI + Guía para dev | Listo |
| 🔜 Fase 2 | Auth Google OAuth en Retool | 1 día |
| 🔜 Fase 2 | Historial en Supabase | 3-5 días |
| 🔜 Fase 3 | Soporte PDF con Claude API | 1 semana |
| 🔜 Fase 3 | Catálogo Nacional de Precios como referencia | 1 semana |
| 🔜 Fase 4 | Multi-tenant (varios clientes) | 2-3 semanas |

---

## 📝 Formatos documentados

| Proveedor | Formato | Estrategia |
|-----------|---------|------------|
| Prov1 | APU fichas multi-fila (.xls) | Estrategia 2 |
| Prov2 | Presupuesto estándar (.xlsx) | Estrategia 1 |
| Prov3 | Por determinar | Auto-detect |

> Para agregar soporte a un nuevo proveedor: compartir el archivo.
> Si el auto-detect falla, se agrega `_extract_formato_[nombre]()` en processor.py.

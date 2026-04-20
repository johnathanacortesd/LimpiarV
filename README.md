<div align="center">

<br/>

```
 ██████╗  ██████╗ ██╗   ██╗    ██████╗ ██████╗  ██████╗  
██╔════╝ ██╔═══██╗██║   ██║    ██╔══██╗██╔══██╗██╔═══██╗ 
╚█████╗  ██║   ██║██║   ██║    ██████╔╝██████╔╝██║   ██║ 
 ╚═══██╗ ██║   ██║╚██╗ ██╔╝    ██╔═══╝ ██╔══██╗██║   ██║ 
██████╔╝ ╚██████╔╝ ╚████╔╝     ██║     ██║  ██║╚██████╔╝ 
╚═════╝   ╚═════╝   ╚═══╝      ╚═╝     ╚═╝  ╚═╝ ╚═════╝  
```

### 🗞️ Limpieza, deduplicación y mapeo de dossieres de prensa · Lite Edition

<br/>

![Python](https://img.shields.io/badge/Python-3.10+-1D4ED8?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.x-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Estado](https://img.shields.io/badge/Estado-Activo-4CAF50?style=for-the-badge)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://duplicas-limpiar-v.streamlit.app/)

<br/>

> Procesador de SOV es la versión **Lite** del flujo de limpieza de dossiers. Sin modelos de IA, sin dependencias pesadas — solo pandas, openpyxl y lógica de negocio sólida. Detecta duplicados exactos y consecutivos, expande menciones, mapea medios y regiones, y entrega un `.xlsx` listo para análisis de Share of Voice en segundos.

<br/>

</div>

---

## 🗺️ Tabla de contenidos

- [¿Qué hace?](#-qué-hace)
- [Diferencias con TransmiApp](#-diferencias-con-transmiapp)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Lógica de duplicados](#-lógica-de-duplicados)
- [Estructura del output](#-estructura-del-output)
- [Stack técnico](#-stack-técnico)
- [Autor](#-autor)

---

## ✨ ¿Qué hace?

En un solo clic, el procesador ejecuta cuatro pasos en cadena:

```
📂 Dossier.xlsx + Configuracion.xlsx
          ↓
  ┌───────────────────────────────┐
  │  PASO 1  Carga configuración  │  region_map + internet_map
  └───────────────┬───────────────┘
                  ↓
  ┌───────────────────────────────┐
  │  PASO 2  Lectura + expansión  │  Una fila por cada mención en
  │          por menciones        │  "Menciones - Empresa" (sep: ;)
  └───────────────┬───────────────┘
                  ↓
  ┌───────────────────────────────┐
  │  PASO 3  Limpieza y mapeo     │  Tipos de medio, regiones,
  │                               │  links, HTML entities, fechas
  └───────────────┬───────────────┘
                  ↓
  ┌───────────────────────────────┐
  │  PASO 4  Detección de         │  Exactos + consecutivos
  │          duplicados           │  Columna "Mantener" con ID origen
  └───────────────┬───────────────┘
                  ↓
       📥 Dossier_Limpio_YYYYMMDD_HHMM.xlsx
```

---

## ⚖️ Diferencias con TransmiApp

| Característica | SOV Pro (Lite) | TransmiApp |
|----------------|:--------------:|:----------:|
| Detección de duplicados | ✅ | ✅ |
| Expansión por menciones | ✅ | ✅ |
| Mapeo de medios y regiones | ✅ | ✅ |
| Limpieza HTML entities | ✅ | ✅ |
| Predicción de Tono con IA | ❌ | ✅ |
| Predicción de Tema con IA | ❌ | ✅ |
| Modelos ML (.pkl) requeridos | ❌ | ✅ |
| Columna "Mantener" con ID origen | ✅ | ❌ |
| Exportación sin límite de links | ✅ | ✅ |
| Dependencias pesadas | ❌ Ligero | ⚠️ Más pesado |

> Usa **SOV Pro (Lite)** cuando necesites velocidad y trazabilidad de duplicados sin predicciones. Usa **TransmiApp** cuando necesites análisis de tono y tema.

---

## 🚀 Instalación

### Prerrequisitos

- Python **3.10+**
- `pip`

### Pasos

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/sov-procesador.git
cd sov-procesador

# 2. Entorno virtual (recomendado)
python -m venv .venv
source .venv/bin/activate        # Linux / macOS
.venv\Scripts\activate           # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Lanzar
streamlit run app.py
```

La app abre en `http://localhost:8501` 🎉

### `requirements.txt`

```txt
streamlit>=1.32.0
pandas>=2.0.0
numpy>=1.26.0
openpyxl>=3.1.0
pyarrow>=14.0.0
```

---

## 📖 Uso

### Archivos requeridos

#### `Dossier.xlsx` (cualquier nombre sin "config")

El archivo exportado del sistema de monitoreo. Columnas clave esperadas:

```
ID Noticia · Fecha · Hora · Tipo de Medio · Medio · Región
Título · Autor - Conductor · Sección - Programa · Nro. Pagina
Dimensión · Duración - Nro. Caracteres · CPE · Tier · Audiencia
Tono · Tema · Temas Generales - Tema · Resumen - Aclaracion
Link Nota · Link (Streaming - Imagen) · Menciones - Empresa
```

#### `Configuracion.xlsx` (debe contener "config" en el nombre)

| Hoja | Columna A | Columna B |
|------|-----------|-----------|
| `Regiones` | Nombre del medio | Región asignada |
| `Internet` | Medio original | Medio normalizado |

### Flujo de uso

```
1. Arrastra ambos archivos al uploader
         ↓
2. La app detecta cuál es cuál:
   · "config" en el nombre → Configuracion.xlsx
   · El otro              → Dossier
         ↓
3. ▶ Iniciar Proceso de Limpieza
         ↓
4. 📥 Descargar resultado
```

---

## 🔍 Lógica de duplicados

El procesador implementa **dos fases** de detección:

### Fase 1 — Duplicados exactos

Agrupa por: `título normalizado + Medio + Fecha + Menciones - Empresa + Hora`

Para medios **Internet**, la hora se ignora en la comparación (`IGNORE_TIME`), ya que el mismo artículo puede aparecer con horarios distintos según el scraper.

Dentro de cada grupo de duplicados, se conserva la fila con **sección más completa** (prioridad `seccion_priority`) y las demás se marcan:

```
Estado_Duplicado = 'Eliminar'
Mantener         = 'Duplicado de: [ID Noticia original]'
Tono / Tema      = 'Duplicada'
```

### Fase 2 — Duplicados consecutivos (solo Internet)

Para noticias de Internet no marcadas en la Fase 1, detecta el mismo artículo publicado en **días consecutivos** (diferencia de fecha = 1 día) en el mismo medio y con la misma mención. Agrupa por `título + Medio + Menciones` y calcula clusters de fechas contiguas:

```python
date_diffs = group['Fecha'].diff().dt.days
cluster_ids = (date_diffs != 1).cumsum()   # nuevo cluster si hay salto de más de 1 día
```

Se conserva la primera aparición del cluster; las siguientes se marcan como duplicadas.

### Ventaja de la columna `Mantener`

A diferencia de simplemente eliminar filas, el procesador **mantiene todas las filas** en el output y añade la columna `Mantener` con el ID de la noticia original. Esto permite:

- Auditar qué se eliminó y por qué
- Filtrar en Excel por `Estado_Duplicado = 'Conservar'` para el análisis final
- Tener trazabilidad completa sin perder información

---

## 📤 Estructura del output

El archivo `Dossier_Limpio_YYYYMMDD_HHMM.xlsx` contiene las columnas en este orden:

```
ID Noticia · Fecha · Hora · Medio · Tipo de Medio · Sección - Programa
Región · Título · Autor - Conductor · Nro. Pagina · Dimensión
Duración - Nro. Caracteres · CPE · Tier · Audiencia · Tono · Tema
Temas Generales - Tema · Resumen - Aclaracion · Link Nota
Link (Streaming - Imagen) · Menciones - Empresa · Mantener  ← nueva
```

**Notas sobre el output:**

- Los links se escriben como hipervínculos activos (`openpyxl`, sin límite de 64k)
- Las fechas se formatean como `DD/MM/YYYY`
- Las columnas de título y resumen tienen ancho fijo de 50 caracteres para legibilidad
- La columna `Duración - Nro. Caracteres` se mueve a `Dimensión` para medios de Radio y Televisión

### Mapeo de tipos de medio

| Valor original | Valor normalizado |
|----------------|-------------------|
| `online` | Internet |
| `diario` | Prensa |
| `am` / `fm` | Radio |
| `aire` / `cable` | Televisión |
| `revista` | Revista |

### Lógica de links por tipo de medio

| Tipo de medio | `Link Nota` | `Link (Streaming - Imagen)` |
|---------------|------------|----------------------------|
| Internet | Recibe el link de imagen (swap) | Recibe el link de nota (swap) |
| Prensa / Revista | Hereda link de imagen si link nota está vacío | Se limpia |
| Radio / Televisión | Sin cambios | Se limpia |

---

## 🛠️ Stack técnico

| Componente | Librería |
|-----------|----------|
| UI / Web app | `streamlit` |
| Lectura Excel con links | `openpyxl` |
| Escritura Excel con hipervínculos | `openpyxl` (sin límite de 64k) |
| DataFrames | `pandas` + `pyarrow` (columnas `string[pyarrow]`) |
| Operaciones numéricas | `numpy` |
| Limpieza HTML | `html`, `re` (stdlib) |

> **¿Por qué openpyxl en lugar de xlsxwriter?** xlsxwriter tiene un límite de ~64k hipervínculos por hoja. openpyxl no tiene ese límite, lo que lo hace más adecuado para dossiers grandes.

---

## 📁 Estructura del proyecto

```
sov-procesador/
│
├── app.py              # App completa (UI + lógica en un solo archivo)
├── requirements.txt    # Dependencias
└── README.md           # Este archivo
```

---

## 👤 Autor

<div align="center">

**Johnathan Cortés** 🕵️😼

_Analista de datos · Bogotá, Colombia_

[![GitHub](https://img.shields.io/badge/GitHub-johnathanacortesd-1D4ED8?style=flat-square&logo=github)](https://github.com/johnathanacortesd)

<br/>

> _"Sin IA, sin magia — solo lógica de negocio bien ejecutada."_

<br/>

---

<sub>Parte del ecosistema de herramientas de monitoreo de prensa · 2025</sub>

</div>

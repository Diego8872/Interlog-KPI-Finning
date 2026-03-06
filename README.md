# INTERLOG · KPI Dashboard

App de gestión de KPIs de Comercio Exterior — FASA / FSM

## 🚀 Cómo deployar en Streamlit Cloud (5 pasos)

### 1. Crear cuenta en GitHub
- Ir a [github.com](https://github.com) y crear una cuenta gratuita (si no tenés)

### 2. Crear un repositorio nuevo
- Click en **"New repository"**
- Nombre: `interlog-kpi` (o el que prefieras)
- Marcarlo como **Private** (recomendado)
- Click en **"Create repository"**

### 3. Subir los archivos
Subir estos 2 archivos al repositorio:
- `app.py`
- `requirements.txt`

Podés hacerlo directo desde GitHub:
- Click en **"Add file" → "Upload files"**
- Arrastrá ambos archivos
- Click en **"Commit changes"**

### 4. Crear cuenta en Streamlit Cloud
- Ir a [share.streamlit.io](https://share.streamlit.io)
- Iniciar sesión con tu cuenta de GitHub

### 5. Deployar la app
- Click en **"New app"**
- Seleccionar tu repositorio `interlog-kpi`
- Branch: `main`
- Main file path: `app.py`
- Click en **"Deploy"**

✅ En 2-3 minutos tenés la URL de tu app lista para usar.

---

## 📋 Cómo usar la app

### Paso 1 — Cargar archivos
Subí los 4 Excel del mes:
- `LIBERADAS_MES_AÑO.xlsx`
- `OFICIALIZADOS_MES_AÑO.xlsx`
- `CM_PRESENTADOS_MES_AÑO.xlsx`
- `CM_APROBADOS_MES_AÑO.xlsx`

### Paso 2 — Revisar desvíos
- La app detecta automáticamente todas las operaciones fuera del rango
- Para cada una tenés que completar:
  - **Desvío**: descripción de qué pasó
  - **Parámetro**: quién es el responsable (ej: INTERLOG, ADUANA, OPERATIVA, FERIADO)
- Si el parámetro NO es INTERLOG → la operación se considera **IN**

### Paso 3 — Dashboard
- Ver todos los KPIs por razón social, proceso y canal
- Gráficos interactivos filtrables

### Paso 4 — Exportar
- Descargar Excel con todos los desvíos justificados
- Descargar PowerPoint del informe completo

---

## 📐 Lógica de KPIs

### Liberadas (todas las razones sociales)
Medición desde `Fecha Oficialización` hasta `Fecha Cancelada` en horas hábiles:

| Vía | Verde | Naranja | Rojo |
|-----|-------|---------|------|
| Avión | 24 hs | 72 hs | 96 hs |
| Marítimo | 72 hs | 96 hs | 120 hs |
| Camión | 24 hs | 48 hs | 72 hs |

### Oficializados
Medición desde `Fecha Oficialización` hasta `Último Evento` en horas hábiles:
- FASA (todas las vías): **24 hs**
- FSM vía aérea/camión: **24 hs**
- FSM vía marítima: **48 hs**

### CM Presentados
Medición desde `TAD Subido` hasta `Último Evento` en horas hábiles: **48 hs**

### CM Aprobados
Solo informativo. Distribución por rangos de días corridos: 0-7 / 8-15 / +15

### Target
- Todos los procesos: **95% IN**
- Una operación con desvío cuyo parámetro ≠ INTERLOG se considera **IN**

---

## 🛠️ Estructura del proyecto

```
interlog-kpi/
├── app.py           # App principal Streamlit
├── requirements.txt # Dependencias Python
└── README.md        # Este archivo
```

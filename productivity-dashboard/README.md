# KPI Dashboard — Guía de Despliegue en Netlify

## Pasos para publicar en la web (5 minutos)

### Opción A — Drag & Drop (más fácil)
1. Ve a https://app.netlify.com/drop
2. Arrastra TODA la carpeta `productivity-dashboard` al recuadro.
3. Netlify te dará una URL publica como `https://tu-nombre.netlify.app`.
4. ¡Listo! Comparte esa URL con cualquier persona.

### Opción B — Conectar con GitHub (actualización automática)
1. Crea una cuenta en https://github.com y sube la carpeta como repositorio.
2. En Netlify → "Import from Git" → selecciona tu repo.
3. Cada vez que actualices el repo, Netlify re-despliega automáticamente.

---

## ¿Cómo funciona la sincronización?

| Botón | ¿Qué hace? |
|---|---|
| **Sincronizar** | Descarga tu Excel de SharePoint y actualiza todas las gráficas |
| **Cargar Archivo** | Sube un .xlsx o .csv descargado manualmente |

El botón **Sincronizar** usa tu enlace público de SharePoint:
`https://aleprocaec-my.sharepoint.com/:x:/g/personal/...`

---

## Para actualizar los datos
1. Abre tu Excel en SharePoint y agrega los datos del día.
2. En el dashboard web, haz clic en **Sincronizar**.
3. Las gráficas se actualizan automáticamente.

---

## Filtros disponibles
- **7d / 30d / 60d / Todo** — vista rápida por período
- **Desde / Hasta** — rango de fechas personalizado

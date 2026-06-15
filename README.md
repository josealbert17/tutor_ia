# Dashboard Tutor IA — Universidad Continental

Herramienta de análisis y visualización de consultas realizadas al Tutor IA institucional.

---

## ¿Qué hace?

Carga automáticamente los datos desde Google Sheets y presenta métricas de uso del Tutor IA organizadas en pestañas: resumen general, análisis por cursos y NRC, temporalidad, temáticas, usuarios, comparativa entre periodos y control de calidad de datos.

---

## Pestañas

| Pestaña | Contenido |
|---|---|
| Resumen | KPIs principales, uso por periodo-bloque, observaciones automáticas |
| Cursos | Ranking de cursos por consultas, tasa de sin respuesta |
| NRC | Índice de dependencia por sección, consultas y usuarios por NRC |
| Temporalidad | Evolución semanal, mapa de calor hora×día, consultas por día |
| Temáticas | Clasificación automática con IA, distribución por categoría y curso |
| Usuarios | Top usuarios, indicadores de recurrencia |
| Comparativa | Métricas comparadas entre periodos, temáticas por bloque |
| Calidad | Registros sin respuesta, mal formados, fuera de rango académico |

---

## Clasificación de temáticas

Las temáticas se clasifican externamente usando OpenAI GPT-4o-mini a través de un notebook de Google Colab. El resultado se sube a Google Sheets y el dashboard lo lee directamente. Las 9 categorías son: Contenido del curso, Cálculo y métodos, Actividades y evaluaciones, Materiales y recursos, Navegación del aula, Soporte técnico, Interacción general, Sin datos, Otros.

---

## Tecnología

React 18 + Vite · Recharts · PapaParse · SheetJS · Desplegado en Vercel

---

## Configuración en Vercel

Variable de entorno requerida: `OPENAI_API_KEY` (para clasificación desde el dashboard como respaldo).

El calendario académico se carga desde una segunda hoja del Google Sheet llamada `CALENDARIO_ACADEMICO`. Para agregar nuevos periodos, solo se agrega una fila en esa hoja — sin necesidad de modificar código.

# Dashboard Tutor IA v2.0

Análisis de consultas del Tutor IA con clasificación automática de temáticas usando OpenAI.

---

## Estructura del proyecto

```
tutor-ia-dashboard/
├── api/
│   └── clasificar.js     ← Proxy seguro hacia OpenAI (la API key vive aquí)
├── src/
│   ├── main.jsx
│   └── App.jsx           ← Dashboard completo
├── index.html
├── package.json
├── vite.config.js
└── vercel.json
```

---

## Deploy en Vercel (recomendado)

### 1. Subir a GitHub
```bash
git init
git add .
git commit -m "Dashboard Tutor IA v2"
git remote add origin https://github.com/TU_USUARIO/tutor-ia-dashboard.git
git push -u origin main
```

### 2. Importar en Vercel
1. Ir a https://vercel.com → New Project
2. Importar el repositorio de GitHub
3. Clic en **Deploy** (Vercel detecta Vite automáticamente)

### 3. ⚠️ Configurar la API Key de OpenAI (IMPORTANTE)
1. En Vercel → tu proyecto → **Settings** → **Environment Variables**
2. Agregar:
   - **Name:** `OPENAI_API_KEY`
   - **Value:** `sk-proj-...` (tu clave de OpenAI)
   - **Environment:** Production, Preview, Development
3. Clic en **Save**
4. Ir a **Deployments** → clic en los tres puntos → **Redeploy**

La API key **nunca aparece en el navegador**. Solo vive en el servidor de Vercel.

---

## Desarrollo local

```bash
npm install
npm run dev
```

Para probar la clasificación con IA en local, crea un archivo `.env.local`:
```
OPENAI_API_KEY=sk-proj-...
```
Luego instala Vercel CLI para emular las API routes:
```bash
npm install -g vercel
vercel dev
```

---

## Cómo funciona la clasificación IA

1. El usuario carga su Excel o conecta Google Sheets
2. Va a la pestaña **🤖 Temáticas IA**
3. Pulsa **Clasificar con IA**
4. El dashboard envía las consultas en lotes de 30 al endpoint `/api/clasificar`
5. El servidor (Vercel) llama a OpenAI con la clave segura
6. OpenAI devuelve: temática, subtema, tipo y profundidad para cada consulta
7. El dashboard actualiza en tiempo real con barra de progreso
8. Al terminar, se puede exportar a Excel con la columna de temática incluida

### Temáticas detectadas automáticamente
OpenAI identifica las categorías que existan en los datos. Ejemplos típicos:
- Comprensión de contenido
- Evaluaciones y notas
- Navegación del aula virtual
- Tareas y entregables
- Soporte técnico
- Dudas procedimentales

---

## Columnas esperadas en el Excel
El dashboard intenta detectar automáticamente los nombres de columna. Funciona mejor si incluye:
- Consulta / Pregunta / Texto
- Respuesta (opcional)
- Nombre del curso / Curso
- Código del curso
- NRC / Sección
- Fecha / Timestamp
- Bloque, Período, Campus (opcionales)

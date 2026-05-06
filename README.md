# Dashboard Análisis Tutor IA

Dashboard interactivo para analizar registros de consultas al tutor IA del aula virtual.
Soporta carga desde archivo Excel y conexión en tiempo real con Google Sheets.

---

## Despliegue rápido

### Opción A — Vercel (recomendado)

1. Crea una cuenta en https://vercel.com (gratis)
2. Instala la CLI: `npm install -g vercel`
3. En la carpeta del proyecto:
```bash
npm install
vercel --prod
```
4. Sigue las instrucciones en pantalla (acepta los valores por defecto)
5. Vercel te dará una URL pública como `https://tutor-ia-dashboard.vercel.app`

**O desde la web:**
1. Sube la carpeta a GitHub
2. En vercel.com → New Project → importa el repositorio
3. Clic en Deploy (configuración automática detectada)

---

### Opción B — Netlify

1. Crea una cuenta en https://netlify.com (gratis)
2. Instala la CLI: `npm install -g netlify-cli`
3. En la carpeta del proyecto:
```bash
npm install
npm run build
netlify deploy --prod --dir=dist
```

**O desde la web (drag & drop):**
1. Ejecuta `npm install && npm run build`
2. En app.netlify.com → Sites → arrastra la carpeta `dist/`
3. Netlify despliega automáticamente

---

## Desarrollo local

```bash
npm install
npm run dev
```

Abre http://localhost:5173

---

## Estructura del proyecto

```
tutor-ia-dashboard/
├── src/
│   ├── main.jsx        # Punto de entrada React
│   └── App.jsx         # Dashboard completo
├── index.html          # HTML base
├── package.json        # Dependencias
├── vite.config.js      # Configuración Vite
├── vercel.json         # Configuración Vercel
└── netlify.toml        # Configuración Netlify
```

---

## Conectar Google Sheets (tiempo real)

Una vez desplegado, el dashboard se conecta automáticamente al Google Sheet configurado.
La URL está pre-configurada en `src/App.jsx`.

Para cambiarla, busca esta línea y reemplaza con tu URL:
```js
const [sheetsUrl, setSheetsUrl] = useState('TU_URL_AQUI');
```

---

## Dependencias principales

- React 18
- Recharts (gráficos)
- SheetJS / xlsx (lectura Excel)
- PapaParse (lectura CSV)
- Vite (build)

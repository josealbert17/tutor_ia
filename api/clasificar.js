export default async function handler(req, res) {
  // Solo POST
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método no permitido' });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'API key no configurada en el servidor' });
  }

  const { consultas } = req.body;
  if (!Array.isArray(consultas) || consultas.length === 0) {
    return res.status(400).json({ error: 'Se requiere un array de consultas' });
  }

  const prompt = `Eres un clasificador de consultas académicas universitarias. 
Analiza cada consulta de estudiantes dirigida a un tutor IA y devuelve una clasificación.

Para cada consulta asigna:
- "tematica": categoría principal (ej: "Comprensión de contenido", "Evaluaciones y notas", "Navegación del aula virtual", "Tareas y entregables", "Soporte técnico", "Dudas procedimentales", "Motivación y orientación", u otras que detectes)
- "subtema": descripción más específica en 3-5 palabras
- "tipo": "Conceptual" | "Operativa" | "Procedimental" | "Emocional"
- "profundidad": "Superficial" | "Intermedia" | "Profunda"

Consultas a clasificar:
${consultas.map((c, i) => `${i + 1}. "${c.texto}"`).join('\n')}

Responde ÚNICAMENTE con un JSON array con exactamente ${consultas.length} objetos en el mismo orden. 
Sin texto adicional, sin markdown, sin bloques de código. Solo el array JSON puro.
Ejemplo de formato:
[{"tematica":"Comprensión de contenido","subtema":"Concepto de variable","tipo":"Conceptual","profundidad":"Intermedia"}]`;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        temperature: 0.1,
        messages: [{ role: 'user', content: prompt }]
      })
    });

    if (!response.ok) {
      const err = await response.json();
      return res.status(response.status).json({ error: err.error?.message || 'Error de OpenAI' });
    }

    const data = await response.json();
    const text = data.choices[0].message.content.trim();

    let clasificaciones;
    try {
      clasificaciones = JSON.parse(text);
    } catch {
      // Intentar extraer JSON si viene con algo extra
      const match = text.match(/\[[\s\S]*\]/);
      if (match) {
        clasificaciones = JSON.parse(match[0]);
      } else {
        return res.status(500).json({ error: 'Respuesta de IA no parseable', raw: text });
      }
    }

    return res.status(200).json({ clasificaciones });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}

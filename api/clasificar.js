export const config = { maxDuration: 30 }

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' })

  const apiKey = process.env.OPENAI_API_KEY
  if (!apiKey) return res.status(500).json({ error: 'OPENAI_API_KEY no configurada' })

  const { consultas } = req.body
  if (!Array.isArray(consultas) || consultas.length === 0) {
    return res.status(400).json({ error: 'Se requiere array de consultas' })
  }

  const prompt = `Eres un clasificador experto de consultas académicas universitarias dirigidas a un tutor IA.

Categorías OBLIGATORIAS (elige la más apropiada):
- "Contenido del curso" → preguntas sobre conceptos, teoría, definiciones del temario
- "Cálculo y métodos" → ejercicios numéricos, fórmulas, resolución paso a paso
- "Actividades y evaluaciones" → tareas, exámenes, rúbricas, calificaciones, entregables
- "Materiales y recursos" → ubicación de PDFs, lecturas, videos, bibliografía
- "Navegación del aula" → cómo usar Moodle, dónde está algo, accesos
- "Soporte técnico" → problemas de plataforma, errores
- "Interacción general" → saludos, conversación casual, agradecimientos
- "Sin datos" → consultas vacías, sin sentido, una sola letra o símbolo
- "Otros" → no encaja en ninguna anterior

Para cada consulta también asigna:
- "subtema": 3-6 palabras descriptivas
- "tipo": "Conceptual" | "Operativa" | "Procedimental" | "Emocional"
- "profundidad": "Superficial" | "Intermedia" | "Profunda"

Consultas (${consultas.length}):
${consultas.map((c, i) => `${i + 1}. ${c.texto}`).join('\n')}

Responde con un objeto JSON con clave "resultados" que sea un array de ${consultas.length} objetos en el mismo orden.
Formato: { "resultados": [{"tematica":"...","subtema":"...","tipo":"...","profundidad":"..."}, ...] }`

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        temperature: 0,
        response_format: { type: 'json_object' },
        messages: [
          { role: 'system', content: 'Eres un clasificador. Respondes SOLO con JSON válido.' },
          { role: 'user', content: prompt }
        ]
      })
    })

    if (!response.ok) {
      const err = await response.json().catch(() => ({}))
      return res.status(response.status).json({ error: err.error?.message || `HTTP ${response.status}` })
    }

    const data = await response.json()
    const text = data.choices[0].message.content.trim()

    let parsed
    try {
      const obj = JSON.parse(text)
      parsed = obj.resultados || obj.clasificaciones || (Array.isArray(obj) ? obj : null)
    } catch {
      const match = text.match(/\[[\s\S]*\]/)
      if (match) parsed = JSON.parse(match[0])
    }

    if (!Array.isArray(parsed)) {
      return res.status(500).json({ error: 'Respuesta no parseable', raw: text.slice(0, 200) })
    }

    return res.status(200).json({ clasificaciones: parsed })
  } catch (error) {
    return res.status(500).json({ error: error.message })
  }
}

import { useState, useMemo, useCallback, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import Papa from 'papaparse'
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, Treemap
} from 'recharts'

// ─── Paleta institucional ───────────────────────────────────────────────────
const C = {
  bg:       '#080f1a',
  surface:  '#0d1b2a',
  card:     '#111e2e',
  border:   '#1a2e45',
  accent:   '#0ea5e9',
  accent2:  '#38bdf8',
  green:    '#10b981',
  amber:    '#f59e0b',
  red:      '#ef4444',
  purple:   '#8b5cf6',
  text:     '#e2e8f0',
  muted:    '#64748b',
  subtle:   '#1e3a5f',
}

const PAL = ['#0ea5e9','#10b981','#f59e0b','#8b5cf6','#ef4444','#06b6d4','#84cc16','#f97316','#ec4899','#6366f1']

// ─── Calendario académico (ajustar fechas según UC) ────────────────────────
const PERIODOS_ACAD = {
  '202520-B1': { label:'2025-20 Bloque 1', inicio: new Date('2025-03-17'), fin: new Date('2025-05-11') },
  '202520-B2': { label:'2025-20 Bloque 2', inicio: new Date('2025-05-12'), fin: new Date('2025-07-06') },
  '202520-B3': { label:'2025-20 Bloque 3', inicio: new Date('2025-07-07'), fin: new Date('2025-08-31') },
  '202610-B1': { label:'2026-10 Bloque 1', inicio: new Date('2026-03-16'), fin: new Date('2026-05-10') },
  '202610-B2': { label:'2026-10 Bloque 2', inicio: new Date('2026-05-11'), fin: new Date('2026-07-05') },
  '202610-B3': { label:'2026-10 Bloque 3', inicio: new Date('2026-07-06'), fin: new Date('2026-08-30') },
}

function getSemana(fecha) {
  for (const [key, v] of Object.entries(PERIODOS_ACAD)) {
    if (fecha >= v.inicio && fecha <= v.fin) {
      const diff = Math.floor((fecha - v.inicio) / (7 * 86400000))
      return { semana: diff + 1, periodoKey: key, periodoLabel: v.label }
    }
  }
  return { semana: null, periodoKey: null, periodoLabel: 'Sin período' }
}

// ─── Parser de Excel/CSV ────────────────────────────────────────────────────
function parseFile(file, callback) {
  const ext = file.name.split('.').pop().toLowerCase()
  if (ext === 'csv') {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: r => callback(normalizeRows(r.data))
    })
  } else {
    const reader = new FileReader()
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const raw = XLSX.utils.sheet_to_json(ws, { defval: '' })
      callback(normalizeRows(raw))
    }
    reader.readAsArrayBuffer(file)
  }
}

// Columnas exactas del Google Sheet de UC:
// User ID | USUARIO | CONSULTA | RESPUESTA DE LA IA | URL DEL CURSO |
// NOMBRE DEL CURSO | CODIGO DEL CURSO | NRC | BLOQUE | PERIODO | CAMPUS | HORA | FECHA
function normalizeRows(raw) {
  return raw.map(r => {
    // Intentar primero con nombres exactos del Sheet UC, luego fallback genérico
    const col = (exact, ...fallbacks) => {
      if (r[exact] !== undefined && r[exact] !== '') return r[exact]
      const keys = Object.keys(r)
      for (const f of fallbacks) {
        const k = keys.find(k => k.toLowerCase().includes(f.toLowerCase()))
        if (k && r[k] !== '') return r[k]
      }
      return ''
    }

    // Fecha: la columna FECHA contiene la fecha, HORA contiene la hora del día
    const fechaRaw = col('FECHA', 'fecha', 'date')
    const horaRaw  = col('HORA', 'hora', 'timestamp')
    let fecha = null
    if (fechaRaw instanceof Date) {
      fecha = fechaRaw
    } else if (typeof fechaRaw === 'string' && fechaRaw) {
      const parts = fechaRaw.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/)
      if (parts) fecha = new Date(`${parts[3].length===2?'20'+parts[3]:parts[3]}-${parts[2].padStart(2,'0')}-${parts[1].padStart(2,'0')}`)
      else fecha = new Date(fechaRaw)
    }

    // Hora del día: si HORA es un número (0-23) usarlo directamente, si no extraer de fecha
    let horaNum = null
    if (horaRaw !== '' && horaRaw !== undefined) {
      const h = parseInt(horaRaw)
      if (!isNaN(h) && h >= 0 && h <= 23) horaNum = h
    }
    if (horaNum === null && fecha && !isNaN(fecha)) horaNum = fecha.getHours()

    const { semana, periodoKey, periodoLabel } = fecha && !isNaN(fecha)
      ? getSemana(fecha)
      : { semana: null, periodoKey: null, periodoLabel: 'Sin período' }

    return {
      userId:       col('User ID', 'user id', 'userid') || '',
      usuario:      col('USUARIO', 'usuario', 'user', 'alumno', 'estudiante', 'email') || '',
      consulta:     col('CONSULTA', 'consulta', 'pregunta', 'query', 'mensaje', 'texto') || '',
      respuestaIA:  col('RESPUESTA DE LA IA', 'respuesta', 'answer', 'reply') || '',
      urlCurso:     col('URL DEL CURSO', 'url') || '',
      nombreCurso:  col('NOMBRE DEL CURSO', 'nombre del curso', 'curso', 'course', 'asignatura') || 'Sin curso',
      codigoCurso:  col('CODIGO DEL CURSO', 'codigo del curso', 'codigo', 'code', 'cod') || '',
      nrc:          String(col('NRC', 'nrc', 'section', 'seccion') || ''),
      bloque:       col('BLOQUE', 'bloque', 'block') || '',
      periodo:      col('PERIODO', 'periodo', 'period', 'ciclo') || '',
      campus:       col('CAMPUS', 'campus', 'sede') || '',
      fechaRaw,
      fecha:        fecha && !isNaN(fecha) ? fecha : null,
      fechaStr:     fecha && !isNaN(fecha) ? fecha.toISOString().slice(0,10) : '',
      hora:         horaNum,
      semana,
      periodoKey,
      periodoLabel,
      tematica:     '',
      subtema:      '',
      tipo:         '',
      profundidad:  '',
    }
  }).filter(r => r.consulta)
}

// ─── Utilidades ─────────────────────────────────────────────────────────────
function aggBy(data, fields) {
  const m = {}
  data.forEach(r => {
    const key = fields.map(f => r[f] || '—').join('||')
    if (!m[key]) { m[key] = { total: 0 }; fields.forEach(f => m[key][f] = r[f] || '—') }
    m[key].total++
  })
  return Object.values(m).sort((a, b) => b.total - a.total)
}

// ─── Componentes base ────────────────────────────────────────────────────────
const Card = ({ children, style }) => (
  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20, ...style }}>
    {children}
  </div>
)

const SLabel = ({ children }) => (
  <div style={{ fontSize: 11, fontWeight: 600, color: C.muted, textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 14 }}>
    {children}
  </div>
)

const Badge = ({ children, color = C.accent }) => (
  <span style={{ background: color + '22', color, border: `1px solid ${color}44`, borderRadius: 6, padding: '2px 8px', fontSize: 11, fontWeight: 600 }}>
    {children}
  </span>
)

const TT = () => (
  <Tooltip
    contentStyle={{ background: '#0d1b2a', border: '1px solid #2a4a6b', borderRadius: 8, fontSize: 12, color: '#e2e8f0' }}
    labelStyle={{ color: '#e2e8f0', fontWeight: 600 }}
    itemStyle={{ color: '#94c8f0' }}
  />
)

function DataTable({ headers, rows, maxH = 360 }) {
  const th = { padding: '8px 12px', textAlign: 'left', fontSize: 11, fontWeight: 600, color: C.muted, borderBottom: `1px solid ${C.border}`, whiteSpace: 'nowrap' }
  const td = { padding: '7px 12px', fontSize: 12, color: C.text, borderBottom: `1px solid ${C.border}22` }
  return (
    <div style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: maxH, borderRadius: 8, border: `1px solid ${C.border}` }}>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead style={{ position: 'sticky', top: 0, background: C.surface }}>
          <tr>{headers.map((h, i) => <th key={i} style={th}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? 'transparent' : C.surface + '88' }}>
              {row.map((cell, j) => <td key={j} style={td}>{cell}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ─── KPI Card ────────────────────────────────────────────────────────────────
function KpiCard({ label, value, sub, color = C.accent }) {
  return (
    <Card style={{ textAlign: 'center' }}>
      <div style={{ fontSize: 32, fontWeight: 700, color, fontFamily: 'JetBrains Mono, monospace' }}>{value}</div>
      <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginTop: 4 }}>{label}</div>
      {sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 3 }}>{sub}</div>}
    </Card>
  )
}

// ─── Tab: Resumen ────────────────────────────────────────────────────────────
function TabResumen({ fd }) {
  const cursos = useMemo(() => [...new Set(fd.map(r => r.nombreCurso))].length, [fd])
  const nrcs = useMemo(() => [...new Set(fd.map(r => r.nrc).filter(Boolean))].length, [fd])
  const usuarios = useMemo(() => [...new Set(fd.map(r => r.usuario).filter(Boolean))].length, [fd])
  const sinResp = useMemo(() => fd.filter(r => !r.respuestaIA || r.respuestaIA.trim() === '').length, [fd])
  const porCurso = useMemo(() => aggBy(fd, ['nombreCurso']).slice(0, 10), [fd])
  const porPeriodo = useMemo(() => aggBy(fd, ['periodoLabel']), [fd])

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 12 }}>
        <KpiCard label="Consultas totales" value={fd.length.toLocaleString()} color={C.accent} />
        <KpiCard label="Cursos" value={cursos} color={C.green} />
        <KpiCard label="NRCs" value={nrcs} color={C.purple} />
        <KpiCard label="Usuarios" value={usuarios > 0 ? usuarios : '—'} color={C.amber} />
        <KpiCard label="Sin respuesta" value={sinResp} sub={fd.length > 0 ? `${((sinResp/fd.length)*100).toFixed(1)}%` : ''} color={C.red} />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SLabel>Top 10 cursos por consultas</SLabel>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={porCurso} layout="vertical" margin={{ left: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e3a5f" />
              <XAxis type="number" tick={{ fill: '#94a3b8', fontSize: 11 }} />
              <YAxis type="category" dataKey="nombreCurso" tick={{ fill: '#cbd5e1', fontSize: 10 }} width={130} />
              <TT />
              <Bar dataKey="total" fill={C.accent} radius={[0, 4, 4, 0]} label={{ position: 'right', fill: '#cbd5e1', fontSize: 10 }} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SLabel>Consultas por período académico</SLabel>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={porPeriodo} dataKey="total" nameKey="periodoLabel" cx="50%" cy="50%" outerRadius={90}
                label={({ name, percent }) => `${(percent*100).toFixed(0)}%`}
                labelLine={{ stroke: '#94a3b8' }}>
                {porPeriodo.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
              </Pie>
              <TT />
              <Legend wrapperStyle={{ fontSize: 11, color: '#cbd5e1' }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  )
}

// ─── Tab: Cursos ─────────────────────────────────────────────────────────────
function TabCursos({ fd }) {
  const ac = useMemo(() => aggBy(fd, ['nombreCurso', 'codigoCurso']), [fd])
  const srm = useMemo(() => {
    const m = {}
    fd.filter(r => !r.respuestaIA || r.respuestaIA.trim() === '').forEach(r => { m[r.nombreCurso] = (m[r.nombreCurso] || 0) + 1 })
    return m
  }, [fd])
  return (
    <Card>
      <SLabel>Consultas por curso</SLabel>
      <DataTable
        headers={['Curso', 'Código', 'Consultas', 'Sin resp.', '% Sin resp.']}
        rows={ac.map(r => {
          const s = srm[r.nombreCurso] || 0
          return [r.nombreCurso, r.codigoCurso, r.total.toLocaleString(), s.toLocaleString(), `${r.total > 0 ? ((s/r.total)*100).toFixed(1) : 0}%`]
        })}
      />
    </Card>
  )
}

// ─── Tab: NRC ─────────────────────────────────────────────────────────────
function TabNRC({ fd }) {
  const rows = useMemo(() => aggBy(fd, ['nrc', 'nombreCurso', 'codigoCurso', 'bloque', 'periodo', 'campus']), [fd])
  return (
    <Card>
      <SLabel>Detalle por NRC</SLabel>
      <DataTable
        headers={['NRC', 'Curso', 'Código', 'Bloque', 'Periodo', 'Campus', 'Consultas']}
        rows={rows.map(r => [r.nrc, r.nombreCurso, r.codigoCurso, r.bloque, r.periodo, r.campus, r.total.toLocaleString()])}
      />
    </Card>
  )
}

// ─── Tab: Temporal ────────────────────────────────────────────────────────────
function TabTemp({ fd }) {
  const dias = useMemo(() => {
    const m = {}
    fd.forEach(r => { if (r.fechaStr) m[r.fechaStr] = (m[r.fechaStr] || 0) + 1 })
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0])).map(([f, c]) => ({ fecha: f, consultas: c }))
  }, [fd])

  const horas = useMemo(() => {
    const m = {}; for (let i = 0; i < 24; i++) m[i] = 0
    fd.forEach(r => { if (r.hora !== null) m[r.hora]++ })
    return Object.entries(m).map(([h, c]) => ({ hora: h + 'h', consultas: c }))
  }, [fd])

  const { semanalData, periodoKeys } = useMemo(() => {
    const periodos = [...new Set(fd.filter(r => r.semana).map(r => r.periodoLabel))].sort()
    const weeks = {}
    fd.forEach(r => {
      if (!r.semana) return
      const key = `Sem ${r.semana}`
      if (!weeks[key]) weeks[key] = { semana: key }
      weeks[key][r.periodoLabel] = (weeks[key][r.periodoLabel] || 0) + 1
    })
    const data = Object.values(weeks).sort((a, b) => parseInt(a.semana.replace('Sem ', '')) - parseInt(b.semana.replace('Sem ', '')))
    return { semanalData: data, periodoKeys: periodos }
  }, [fd])

  const acalInfo = Object.entries(PERIODOS_ACAD).map(([, v]) => {
    const fmt = d => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
    return [v.label, fmt(v.inicio), fmt(v.fin)]
  })

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <Card>
        <SLabel>Evolución semanal por período académico</SLabel>
        {semanalData.length > 0 ? (
          <>
            <ResponsiveContainer width="100%" height={240}>
              <LineChart data={semanalData} margin={{ right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e3a5f" />
                <XAxis dataKey="semana" tick={{ fill: '#cbd5e1', fontSize: 11 }} />
                <YAxis tick={{ fill: '#cbd5e1', fontSize: 11 }} />
                <TT />
                <Legend wrapperStyle={{ fontSize: 11, color: '#cbd5e1', paddingTop: 8 }} />
                {periodoKeys.map((p, i) => (
                  <Line key={p} type="monotone" dataKey={p} stroke={PAL[i % PAL.length]} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                ))}
              </LineChart>
            </ResponsiveContainer>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
              Sem 1 = primera semana del bloque. Las semanas se calculan automáticamente por fecha de consulta.
            </div>
          </>
        ) : (
          <div style={{ color: C.muted, fontSize: 13, padding: '16px 0' }}>No hay registros con semanas asignadas. Verifica que las fechas del Excel estén en formato reconocible.</div>
        )}
      </Card>

      <Card>
        <SLabel>Calendario académico de referencia</SLabel>
        <DataTable headers={['Período', 'Inicio', 'Fin']} rows={acalInfo} maxH={220} />
      </Card>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SLabel>Consultas por día</SLabel>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={dias}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e3a5f" />
              <XAxis dataKey="fecha" tick={{ fill: '#cbd5e1', fontSize: 10 }} interval="preserveStartEnd" />
              <YAxis tick={{ fill: '#cbd5e1', fontSize: 11 }} />
              <TT />
              <Line type="monotone" dataKey="consultas" stroke={C.accent} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SLabel>Consultas por hora del día</SLabel>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={horas}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e3a5f" />
              <XAxis dataKey="hora" tick={{ fill: '#cbd5e1', fontSize: 10 }} />
              <YAxis tick={{ fill: '#cbd5e1', fontSize: 11 }} />
              <TT />
              <Bar dataKey="consultas" fill={C.purple} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  )
}

// ─── Tab: Temáticas (con IA) ─────────────────────────────────────────────────
function TabTematicas({ data, setData }) {
  const [estado, setEstado] = useState('idle') // idle | procesando | listo | error
  const [progreso, setProgreso] = useState({ actual: 0, total: 0 })
  const [errorMsg, setErrorMsg] = useState('')
  const [filtroTematica, setFiltroTematica] = useState('todas')

  const clasificadas = useMemo(() => data.filter(r => r.tematica), [data])
  const pendientes = useMemo(() => data.filter(r => !r.tematica && r.consulta), [data])
  const yaClasificado = clasificadas.length > 0

  const tematicasDisp = useMemo(() => {
    const s = new Set(clasificadas.map(r => r.tematica))
    return ['todas', ...Array.from(s).sort()]
  }, [clasificadas])

  const fdFiltrado = useMemo(() => {
    if (filtroTematica === 'todas') return clasificadas
    return clasificadas.filter(r => r.tematica === filtroTematica)
  }, [clasificadas, filtroTematica])

  const treemapData = useMemo(() => {
    const m = {}
    clasificadas.forEach(r => { m[r.tematica] = (m[r.tematica] || 0) + 1 })
    return Object.entries(m).map(([name, size]) => ({ name, size })).sort((a,b)=>b.size-a.size)
  }, [clasificadas])

  const porTipo = useMemo(() => {
    const m = {}
    clasificadas.forEach(r => { if (r.tipo) m[r.tipo] = (m[r.tipo] || 0) + 1 })
    return Object.entries(m).map(([name, value]) => ({ name, value }))
  }, [clasificadas])

  const porProfundidad = useMemo(() => {
    const m = {}
    clasificadas.forEach(r => { if (r.profundidad) m[r.profundidad] = (m[r.profundidad] || 0) + 1 })
    return Object.entries(m).map(([name, value]) => ({ name, value }))
  }, [clasificadas])

  const iniciarClasificacion = async () => {
    const objetivos = pendientes.length > 0 ? pendientes : data
    if (objetivos.length === 0) return

    setEstado('procesando')
    setErrorMsg('')
    const LOTE = 20 // lotes más pequeños para evitar timeouts
    const total = objetivos.length
    setProgreso({ actual: 0, total })

    const indices = objetivos.map((r) => ({ r, i: data.indexOf(r) }))
    const newData = [...data]
    let procesados = 0
    let erroresConsecutivos = 0

    for (let start = 0; start < indices.length; start += LOTE) {
      const lote = indices.slice(start, start + LOTE)
      const payload = lote.map(({ r }) => ({ texto: r.consulta.slice(0, 250) }))

      let exitoso = false
      for (let intento = 0; intento < 3; intento++) {
        try {
          const res = await fetch('/api/clasificar', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ consultas: payload })
          })

          if (!res.ok) {
            const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }))
            // Rate limit: esperar más
            if (res.status === 429) {
              await new Promise(r => setTimeout(r, 3000 * (intento + 1)))
              continue
            }
            throw new Error(err.error || `Error ${res.status}`)
          }

          const { clasificaciones } = await res.json()
          lote.forEach(({ i }, j) => {
            const cl = clasificaciones?.[j]
            if (cl) {
              newData[i] = { ...newData[i], tematica: cl.tematica || 'Sin clasificar', subtema: cl.subtema || '', tipo: cl.tipo || '', profundidad: cl.profundidad || '' }
            }
          })
          procesados += lote.length
          exitoso = true
          erroresConsecutivos = 0
          break
        } catch (err) {
          if (intento === 2) {
            // Marcar este lote como "Sin clasificar" y continuar
            lote.forEach(({ i }) => {
              newData[i] = { ...newData[i], tematica: 'Sin clasificar', subtema: '', tipo: '', profundidad: '' }
            })
            procesados += lote.length
            erroresConsecutivos++
            if (erroresConsecutivos >= 5) {
              setErrorMsg(`Se detuvieron 5 lotes consecutivos con error: ${err.message}`)
              setEstado('error')
              setData([...newData])
              return
            }
          } else {
            await new Promise(r => setTimeout(r, 1500 * (intento + 1)))
          }
        }
      }

      setProgreso({ actual: procesados, total })
      setData([...newData])
      // Pausa entre lotes para no saturar la API
      await new Promise(r => setTimeout(r, 500))
    }
    setEstado('listo')
  }

  const exportarExcel = () => {
    const rows = clasificadas.map(r => ({
      'Consulta': r.consulta,
      'Temática': r.tematica,
      'Subtema': r.subtema,
      'Tipo': r.tipo,
      'Profundidad': r.profundidad,
      'Curso': r.nombreCurso,
      'NRC': r.nrc,
      'Período': r.periodoLabel,
      'Fecha': r.fechaStr,
    }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Temáticas')
    XLSX.writeFile(wb, 'clasificacion_tematicas.xlsx')
  }

  const CustomTreemapContent = ({ x, y, width, height, name, size, index }) => {
    if (width < 40 || height < 30) return null
    return (
      <g>
        <rect x={x} y={y} width={width} height={height} fill={PAL[index % PAL.length]} fillOpacity={0.85} rx={4} />
        <text x={x + width / 2} y={y + height / 2 - 6} textAnchor="middle" fill="#fff" fontSize={Math.min(12, width / 8)} fontWeight={600}>
          {name.length > 18 ? name.slice(0,16)+'…' : name}
        </text>
        <text x={x + width / 2} y={y + height / 2 + 10} textAnchor="middle" fill="#ffffffcc" fontSize={11}>
          {size}
        </text>
      </g>
    )
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>

      {/* Panel de estado */}
      <Card>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.text, marginBottom: 4 }}>Clasificación automática de temáticas</div>
            <div style={{ fontSize: 12, color: C.muted }}>
              {clasificadas.length > 0
                ? <><span style={{ color: C.green }}>✓ {clasificadas.length} consultas clasificadas</span>{pendientes.length > 0 && <span style={{ color: C.amber }}> · {pendientes.length} pendientes</span>}</>
                : <span>{data.length} consultas listas para clasificar</span>
              }
            </div>
          </div>

          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {clasificadas.length > 0 && (
              <button onClick={exportarExcel}
                style={{ background: C.subtle, border: `1px solid ${C.border}`, color: C.accent2, borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                ↓ Exportar Excel
              </button>
            )}
            <button
              onClick={iniciarClasificacion}
              disabled={estado === 'procesando' || data.length === 0}
              style={{
                background: estado === 'procesando' ? C.subtle : C.accent,
                border: 'none', color: '#fff', borderRadius: 8,
                padding: '8px 18px', fontSize: 12, fontWeight: 700,
                cursor: estado === 'procesando' ? 'not-allowed' : 'pointer',
                opacity: estado === 'procesando' ? 0.7 : 1
              }}>
              {estado === 'procesando' ? `Clasificando ${progreso.actual}/${progreso.total}...` :
               yaClasificado && pendientes.length > 0 ? `Clasificar ${pendientes.length} restantes` :
               yaClasificado ? 'Reclasificar todo' : 'Clasificar con IA'}
            </button>
          </div>
        </div>

        {estado === 'procesando' && (
          <div style={{ marginTop: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6 }}>
              <span style={{ fontSize: 12, color: '#cbd5e1' }}>Procesando en lotes de 20</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.accent, fontFamily: 'JetBrains Mono, monospace' }}>
                {progreso.actual} / {progreso.total} ({progreso.total > 0 ? Math.round((progreso.actual/progreso.total)*100) : 0}%)
              </span>
            </div>
            <div style={{ background: '#1a2e45', borderRadius: 4, height: 8, overflow: 'hidden' }}>
              <div style={{
                background: `linear-gradient(90deg, ${C.accent}, ${C.green})`,
                height: '100%',
                width: `${progreso.total > 0 ? (progreso.actual / progreso.total) * 100 : 0}%`,
                transition: 'width 0.4s ease', borderRadius: 4
              }} />
            </div>
            <div style={{ fontSize: 11, color: '#64748b', marginTop: 6 }}>
              ⏳ Tiempo estimado: ~{Math.ceil((progreso.total - progreso.actual) / 20 * 0.5)} min restantes
            </div>
          </div>
        )}

        {estado === 'error' && (
          <div style={{ marginTop: 12, background: C.red + '15', border: `1px solid ${C.red}44`, borderRadius: 8, padding: '10px 14px', fontSize: 12, color: C.red }}>
            ⚠ Error: {errorMsg}. Verifica que la variable OPENAI_API_KEY esté configurada en Vercel.
          </div>
        )}

        {estado === 'listo' && (
          <div style={{ marginTop: 12, background: C.green + '15', border: `1px solid ${C.green}44`, borderRadius: 8, padding: '10px 14px', fontSize: 12, color: C.green }}>
            ✓ Clasificación completada · {clasificadas.length} consultas procesadas
          </div>
        )}
      </Card>

      {clasificadas.length === 0 && (
        <Card>
          <div style={{ textAlign: 'center', padding: '40px 20px', color: C.muted }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>🤖</div>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.text, marginBottom: 8 }}>Sin clasificaciones aún</div>
            <div style={{ fontSize: 12 }}>Pulsa "Clasificar con IA" para que OpenAI analice automáticamente las temáticas de todas las consultas.</div>
          </div>
        </Card>
      )}

      {clasificadas.length > 0 && (
        <>
          {/* Treemap de temáticas */}
          <Card>
            <SLabel>Mapa de temáticas detectadas</SLabel>
            <ResponsiveContainer width="100%" height={280}>
              <Treemap data={treemapData} dataKey="size" content={<CustomTreemapContent />} />
            </ResponsiveContainer>
          </Card>

          {/* Tipo y Profundidad */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <Card>
              <SLabel>Tipo de consulta</SLabel>
              <ResponsiveContainer width="100%" height={200}>
                <PieChart>
                  <Pie data={porTipo} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} label={({ name, percent }) => `${name} ${(percent*100).toFixed(0)}%`}>
                    {porTipo.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
                  </Pie>
                  <TT />
                </PieChart>
              </ResponsiveContainer>
            </Card>
            <Card>
              <SLabel>Profundidad de consulta</SLabel>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={porProfundidad}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                  <XAxis dataKey="name" tick={{ fill: C.muted, fontSize: 12 }} />
                  <YAxis tick={{ fill: C.muted, fontSize: 11 }} />
                  <TT />
                  <Bar dataKey="value" radius={[4,4,0,0]}>
                    {porProfundidad.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          </div>

          {/* Tabla detalle con filtro */}
          <Card>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14, flexWrap: 'wrap', gap: 8 }}>
              <SLabel style={{ margin: 0 }}>Detalle de consultas clasificadas</SLabel>
              <select
                value={filtroTematica}
                onChange={e => setFiltroTematica(e.target.value)}
                style={{ background: C.surface, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: '5px 10px', fontSize: 12 }}>
                {tematicasDisp.map(t => <option key={t} value={t}>{t === 'todas' ? 'Todas las temáticas' : t}</option>)}
              </select>
            </div>
            <DataTable
              headers={['Consulta', 'Temática', 'Subtema', 'Tipo', 'Profundidad', 'Curso']}
              rows={fdFiltrado.slice(0, 500).map(r => [
                r.consulta.slice(0, 90) + (r.consulta.length > 90 ? '…' : ''),
                r.tematica, r.subtema, r.tipo, r.profundidad, r.nombreCurso
              ])}
              maxH={420}
            />
            {fdFiltrado.length > 500 && <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>Mostrando 500 de {fdFiltrado.length} filas. Exporta a Excel para ver todas.</div>}
          </Card>
        </>
      )}
    </div>
  )
}

// ─── URL fija del Google Sheet de UC ─────────────────────────────────────────
// URL publicada del Google Sheet de UC (formato CSV)
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSv6HB5-J_8Gqh9RctyZRLPlq8Wi8UVUl59HxdW2WiOHGV15WNr6ddLmcf0VOoLxWhhkd4Ncp6il1_g/pub?output=csv'

async function cargarDesdeSheet() {
  const res = await fetch(SHEET_CSV_URL)
  if (!res.ok) throw new Error(`Error ${res.status} al acceder al Google Sheet. Verifica que esté publicado como CSV.`)
  const text = await res.text()
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true, skipEmptyLines: true,
      complete: r => resolve(normalizeRows(r.data)),
      error: e => reject(e)
    })
  })
}

// ─── Pantalla de carga ────────────────────────────────────────────────────────
function PantallaCarga({ onLoad }) {
  const [estado, setEstado] = useState('cargando') // cargando | error | manual
  const [errorMsg, setErrorMsg] = useState('')
  const [dragging, setDragging] = useState(false)
  const [sheetsUrl, setSheetsUrl] = useState('')
  const [loadingManual, setLoadingManual] = useState(false)
  const fileRef = useRef()

  // Carga automática al montar — con timeout de 15s
  useEffect(() => {
    const timeout = setTimeout(() => {
      setErrorMsg('La carga tardó demasiado. Verifica tu conexión o carga el archivo manualmente.')
      setEstado('error')
    }, 15000)

    cargarDesdeSheet()
      .then(rows => {
        clearTimeout(timeout)
        onLoad(rows, { name: 'Google Sheets — UC', source: 'sheets' })
      })
      .catch(e => {
        clearTimeout(timeout)
        setErrorMsg(e.message)
        setEstado('error')
      })

    return () => clearTimeout(timeout)
  }, [])

  const handleFile = file => {
    parseFile(file, rows => onLoad(rows, { name: file.name, source: 'file' }))
  }

  const handleDrop = e => {
    e.preventDefault(); setDragging(false)
    const f = e.dataTransfer.files[0]
    if (f) handleFile(f)
  }

  const handleSheetsManual = async () => {
    if (!sheetsUrl.trim()) return
    setLoadingManual(true)
    try {
      const csvUrl = sheetsUrl.replace(/\/edit.*/, '/export?format=csv')
      const res = await fetch(csvUrl)
      if (!res.ok) throw new Error('No se pudo acceder. Verifica que el Sheet esté publicado.')
      const text = await res.text()
      Papa.parse(text, {
        header: true, skipEmptyLines: true,
        complete: r => onLoad(normalizeRows(r.data), { name: 'Google Sheets', source: 'sheets' })
      })
    } catch (e) {
      setErrorMsg(e.message)
    } finally {
      setLoadingManual(false)
    }
  }

  // Estado: cargando automáticamente
  if (estado === 'cargando') {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.bg }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.accent, textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 16 }}>
            Universidad Continental
          </div>
          <div style={{ fontSize: 22, fontWeight: 700, color: C.text, marginBottom: 8 }}>Dashboard Tutor IA</div>
          <div style={{ fontSize: 13, color: C.muted, marginBottom: 28 }}>Cargando datos desde Google Sheets…</div>
          <div style={{ display: 'flex', justifyContent: 'center', gap: 6 }}>
            {[0,1,2].map(i => (
              <div key={i} style={{
                width: 8, height: 8, borderRadius: '50%', background: C.accent,
                animation: `pulse 1.2s ease-in-out ${i * 0.2}s infinite`,
                opacity: 0.7
              }} />
            ))}
          </div>
          <style>{`@keyframes pulse { 0%,80%,100%{transform:scale(0.6);opacity:0.4} 40%{transform:scale(1);opacity:1} }`}</style>
        </div>
      </div>
    )
  }

  // Estado: error — mostrar opciones manuales
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, background: C.bg }}>
      <div style={{ width: '100%', maxWidth: 560 }}>
        <div style={{ textAlign: 'center', marginBottom: 28 }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.accent, textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 8 }}>
            Universidad Continental
          </div>
          <h1 style={{ fontSize: 24, fontWeight: 700, color: C.text, marginBottom: 8 }}>Dashboard Tutor IA</h1>
        </div>

        {errorMsg && (
          <div style={{ background: C.red + '15', border: `1px solid ${C.red}44`, borderRadius: 10, padding: '12px 16px', fontSize: 12, color: C.red, marginBottom: 20 }}>
            ⚠ No se pudo cargar automáticamente: {errorMsg}
            <div style={{ marginTop: 6, color: C.muted }}>
              Asegúrate de que el Sheet esté publicado: <strong style={{ color: C.text }}>Archivo → Compartir → Publicar en la web → CSV</strong>
            </div>
          </div>
        )}

        {/* Drop zone manual */}
        <div
          onDragOver={e => { e.preventDefault(); setDragging(true) }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileRef.current.click()}
          style={{
            border: `2px dashed ${dragging ? C.accent : C.border}`,
            borderRadius: 14, padding: '28px 24px', textAlign: 'center',
            cursor: 'pointer', marginBottom: 14,
            background: dragging ? C.accent + '08' : C.card, transition: 'all 0.2s'
          }}>
          <div style={{ fontSize: 28, marginBottom: 8 }}>📂</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 4 }}>Cargar archivo Excel o CSV</div>
          <div style={{ fontSize: 11, color: C.muted }}>Arrastra aquí o haz clic para seleccionar</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>

        <Card>
          <SLabel>O conectar a otro Google Sheet</SLabel>
          <div style={{ display: 'flex', gap: 8 }}>
            <input value={sheetsUrl} onChange={e => setSheetsUrl(e.target.value)}
              placeholder="URL del Google Sheet"
              style={{ flex: 1, background: C.surface, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: '8px 12px', fontSize: 12 }} />
            <button onClick={handleSheetsManual} disabled={loadingManual || !sheetsUrl.trim()}
              style={{ background: C.accent, border: 'none', color: '#fff', borderRadius: 8, padding: '8px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
              {loadingManual ? 'Cargando…' : 'Conectar'}
            </button>
          </div>
        </Card>
      </div>
    </div>
  )
}

// ─── App principal ────────────────────────────────────────────────────────────
const TABS = [
  { id: 'resumen',    label: 'Resumen' },
  { id: 'cursos',     label: 'Cursos' },
  { id: 'nrc',        label: 'NRC' },
  { id: 'temporal',   label: 'Temporal' },
  { id: 'tematicas',  label: '🤖 Temáticas IA' },
]

export default function App() {
  const [data, setData] = useState(null)
  const [fileInfo, setFileInfo] = useState(null)
  const [tab, setTab] = useState('resumen')
  const [filtroPeriodo, setFiltroPeriodo] = useState('todos')
  const [filtroCurso, setFiltroCurso] = useState('todos')

  const handleLoad = useCallback((rows, info) => {
    setData(rows)
    setFileInfo({ ...info, total: rows.length, loaded: new Date() })
    setTab('resumen')
  }, [])

  const periodos = useMemo(() => data ? ['todos', ...new Set(data.map(r => r.periodoLabel))] : [], [data])
  const cursos = useMemo(() => data ? ['todos', ...new Set(data.map(r => r.nombreCurso)).values()].slice(0, 60) : [], [data])

  const fd = useMemo(() => {
    if (!data) return []
    let r = data
    if (filtroPeriodo !== 'todos') r = r.filter(x => x.periodoLabel === filtroPeriodo)
    if (filtroCurso !== 'todos') r = r.filter(x => x.nombreCurso === filtroCurso)
    return r
  }, [data, filtroPeriodo, filtroCurso])

  if (!data) return <PantallaCarga onLoad={handleLoad} />

  const selStyle = active => ({
    padding: '7px 16px', fontSize: 13, fontWeight: active ? 700 : 500,
    color: active ? C.accent : C.muted,
    background: active ? C.accent + '15' : 'transparent',
    border: 'none', borderBottom: `2px solid ${active ? C.accent : 'transparent'}`,
    cursor: 'pointer', whiteSpace: 'nowrap', transition: 'all 0.15s'
  })

  return (
    <div style={{ minHeight: '100vh', background: C.bg }}>
      {/* Header */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: '12px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
        <div>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>Dashboard Tutor IA</div>
          <div style={{ fontSize: 11, color: C.muted }}>
            {fileInfo?.source === 'sheets' ? '🟢 Google Sheets' : `📁 ${fileInfo?.name}`} · {fileInfo?.total?.toLocaleString()} registros
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <select value={filtroPeriodo} onChange={e => setFiltroPeriodo(e.target.value)}
            style={{ background: C.card, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: '5px 10px', fontSize: 12 }}>
            {periodos.map(p => <option key={p} value={p}>{p === 'todos' ? 'Todos los períodos' : p}</option>)}
          </select>
          <select value={filtroCurso} onChange={e => setFiltroCurso(e.target.value)}
            style={{ background: C.card, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: '5px 10px', fontSize: 12 }}>
            {cursos.map(c => <option key={c} value={c}>{c === 'todos' ? 'Todos los cursos' : c}</option>)}
          </select>
          <button onClick={() => { setData(null); setFileInfo(null) }}
            style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 6, padding: '5px 12px', fontSize: 12, cursor: 'pointer' }}>
            ↺ Recargar datos
          </button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, display: 'flex', overflowX: 'auto', paddingLeft: 16 }}>
        {TABS.map(t => (
          <button key={t.id} style={selStyle(tab === t.id)} onClick={() => setTab(t.id)}>{t.label}</button>
        ))}
      </div>

      {/* Contenido */}
      <div style={{ padding: '20px 24px', maxWidth: 1200, margin: '0 auto' }}>
        {tab === 'resumen'   && <TabResumen fd={fd} />}
        {tab === 'cursos'    && <TabCursos fd={fd} />}
        {tab === 'nrc'       && <TabNRC fd={fd} />}
        {tab === 'temporal'  && <TabTemp fd={fd} />}
        {tab === 'tematicas' && <TabTematicas data={data} setData={setData} />}
      </div>
    </div>
  )
}

import { useState, useMemo, useCallback, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import Papa from 'papaparse'
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area
} from 'recharts'

// ═══════════════════════════════════════════════════════════════════════════
// PALETA DE COLORES
// ═══════════════════════════════════════════════════════════════════════════
const C = {
  bg:       '#0a1426',
  surface:  '#0f1e36',
  card:     '#13243f',
  cardAlt:  '#162a4a',
  border:   '#1e3a5f',
  borderHi: '#2a4a6b',
  accent:   '#3b82f6',
  accent2:  '#60a5fa',
  green:    '#10b981',
  amber:    '#f59e0b',
  red:      '#ef4444',
  purple:   '#a855f7',
  pink:     '#ec4899',
  cyan:     '#06b6d4',
  text:     '#f1f5f9',
  textDim:  '#cbd5e1',
  muted:    '#94a3b8',
  subtle:   '#64748b',
}

const PAL = ['#3b82f6','#10b981','#f59e0b','#ef4444','#a855f7','#06b6d4','#ec4899','#84cc16','#f97316','#6366f1','#14b8a6','#eab308']
const HEAT_PAL = ['#0a1426', '#1e3a5f', '#1e4a8a', '#2563eb', '#3b82f6', '#60a5fa']

// ═══════════════════════════════════════════════════════════════════════════
// CALENDARIO ACADÉMICO - se calcula dinámicamente desde los datos
// ═══════════════════════════════════════════════════════════════════════════
// CALENDARIO ACADEMICO UC - Fallback hardcodeado (se sobreescribe con datos del Sheet)
// Regla: el número al final del código determina el bloque (1=Bloque A, 2=Bloque B)
// Modalidades: VV=Virtual, WA=Semipresencial Arequipa, WL=Lima, WC=Cusco, WH=Huancayo
// Todos los códigos con el mismo número comparten fechas dentro del mismo periodo
const BLOQUES_UC = {
  '202520-1': { label: '202520 Bloque A', inicio: new Date('2025-08-18'), fin: new Date('2025-10-12') },
  '202520-2': { label: '202520 Bloque B', inicio: new Date('2025-10-13'), fin: new Date('2025-12-07') },
  '202600-1': { label: '202600 Bloque A', inicio: new Date('2026-01-02'), fin: new Date('2026-02-26') },
  '202610-1': { label: '202610 Bloque A', inicio: new Date('2026-03-16'), fin: new Date('2026-05-10') },
  '202610-2': { label: '202610 Bloque B', inicio: new Date('2026-05-18'), fin: new Date('2026-07-12') },
}

// Modalidades conocidas: VV, WA, WL, WC, WH — el número es siempre el último caracter
const MODALIDADES = ['VV','WA','WL','WC','WH']

function expandirCalendario(bloques) {
  const cal = {}
  Object.entries(bloques).forEach(([key, v]) => {
    const [periodo, num] = key.split('-')
    MODALIDADES.forEach(mod => {
      cal[`${periodo} ${mod}${num}`] = v
    })
  })
  return cal
}

const CALENDARIO_UC_DEFAULT = expandirCalendario(BLOQUES_UC)

// URL de la hoja CALENDARIO_ACADEMICO publicada como CSV
const CALENDARIO_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSv6HB5-J_8Gqh9RctyZRLPlq8Wi8UVUl59HxdW2WiOHGV15WNr6ddLmcf0VOoLxWhhkd4Ncp6il1_g/pub?gid=2094941470&single=true&output=csv'

// Parsea fecha DD/MM/YYYY o YYYY-MM-DD
function parseFechaUC(s) {
  if (!s) return null
  s = String(s).trim()
  // DD/MM/YYYY
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/)
  if (m1) return new Date(`${m1[3]}-${m1[2].padStart(2,'0')}-${m1[1].padStart(2,'0')}`)
  // YYYY-MM-DD
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (m2) return new Date(s)
  return null
}

async function cargarCalendarioDesdeSheet() {
  try {
    const res = await fetch(CALENDARIO_CSV_URL)
    if (!res.ok) return CALENDARIO_UC_DEFAULT
    const text = await res.text()
    const lines = text.trim().split('\n').map(l => l.split(',').map(c => c.replace(/^"|"$/g, '').trim()))
    if (lines.length < 2) return CALENDARIO_UC_DEFAULT
    // Detectar columnas
    const headers = lines[0].map(h => h.toUpperCase())
    const iPerBloque = headers.findIndex(h => h.includes('PERIODO'))
    const iCodigo    = headers.findIndex(h => h.includes('CODIGO') || h.includes('CÓDIGO'))
    const iInicio    = headers.findIndex(h => h.includes('INICIO'))
    const iFin       = headers.findIndex(h => h.includes('FIN'))
    if (iPerBloque < 0 || iInicio < 0 || iFin < 0) return CALENDARIO_UC_DEFAULT
    const bloques = {}
    lines.slice(1).forEach(cols => {
      const periodoBloque = cols[iPerBloque]  // ej: "202520-A"
      const codigo        = iCodigo >= 0 ? cols[iCodigo] : ''  // ej: "VV1"
      const inicio        = parseFechaUC(cols[iInicio])
      const fin           = parseFechaUC(cols[iFin])
      if (!periodoBloque || !inicio || !fin) return
      // Extraer numero del bloque del código (último caracter de "VV1" → "1")
      const num = codigo ? codigo.slice(-1) : ''
      const periodo = periodoBloque.replace(/-[AB]$/, '')  // "202520-A" → "202520"
      const bKey = `${periodo}-${num}`  // "202520-1"
      bloques[bKey] = { label: periodoBloque, inicio, fin }
    })
    // Expandir a todas las modalidades
    return Object.keys(bloques).length > 0 ? expandirCalendario(bloques) : CALENDARIO_UC_DEFAULT
  } catch {
    return CALENDARIO_UC_DEFAULT
  }
}

function getSemanaNum(fecha, periodoBloque, calendarioUC) {
  if (!fecha || !periodoBloque) return null
  const cal = (calendarioUC || CALENDARIO_UC_DEFAULT)[periodoBloque]
  if (!cal) return null
  if (fecha < cal.inicio || fecha > cal.fin) return null
  const diff = Math.floor((fecha - cal.inicio) / (7 * 86400000))
  return diff + 1
}

function esFueraDeRango(fecha, periodoBloque, calendarioUC) {
  if (!fecha || !periodoBloque) return true
  const cal = (calendarioUC || CALENDARIO_UC_DEFAULT)[periodoBloque]
  if (!cal) return true
  return fecha < cal.inicio || fecha > cal.fin
}

// ═══════════════════════════════════════════════════════════════════════════
// PARSER DE EXCEL/CSV - COLUMNAS UC
// ═══════════════════════════════════════════════════════════════════════════
function parseFile(file, callback) {
  const ext = file.name.split('.').pop().toLowerCase()
  if (ext === 'csv') {
    Papa.parse(file, { header: true, skipEmptyLines: true, complete: r => callback(normalizeRows(r.data)) })
  } else {
    const reader = new FileReader()
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true })
      const ws = wb.Sheets[wb.SheetNames[0]]
      callback(normalizeRows(XLSX.utils.sheet_to_json(ws, { defval: '' })))
    }
    reader.readAsArrayBuffer(file)
  }
}

// Detecta consultas mal formadas
function esConsultaInvalida(texto) {
  if (!texto || typeof texto !== 'string') return { invalido: true, motivo: 'Valor vacío o inválido' }
  const t = texto.trim()
  if (t.length === 0) return { invalido: true, motivo: 'Valor vacío o inválido' }
  if (t.length <= 3 && /^[a-zA-Z¿?¡!.,;:\s]+$/.test(t)) return { invalido: true, motivo: 'Valor vacío o inválido' }
  if (/^(si|no|ok|sí|s|n|hi|hola|gracias)$/i.test(t)) return { invalido: true, motivo: 'Valor vacío o inválido' }
  return { invalido: false, motivo: '' }
}

function normalizeRows(raw) {
  return raw.map(r => {
    const col = (exact, ...fallbacks) => {
      if (r[exact] !== undefined && r[exact] !== '') return r[exact]
      const keys = Object.keys(r)
      for (const f of fallbacks) {
        const k = keys.find(k => k.toLowerCase().includes(f.toLowerCase()))
        if (k && r[k] !== '') return r[k]
      }
      return ''
    }

    const fechaRaw = col('FECHA', 'fecha', 'date')
    const horaRaw  = col('HORA', 'hora')
    let fecha = null

    if (fechaRaw instanceof Date) {
      fecha = fechaRaw
    } else if (typeof fechaRaw === 'string' && fechaRaw.trim()) {
      const f = fechaRaw.trim()
      // Formato: "2026-06-14 00:00:00" o "2026-06-14T00:00:00"
      const mISO = f.match(/^(\d{4})-(\d{2})-(\d{2})/)
      if (mISO) {
        fecha = new Date(`${mISO[1]}-${mISO[2]}-${mISO[3]}T00:00:00`)
      } else {
        // Formato: "14/06/2026" o "14-06-2026"
        const mDMY = f.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/)
        if (mDMY) {
          fecha = new Date(`${mDMY[3]}-${mDMY[2].padStart(2,'0')}-${mDMY[1].padStart(2,'0')}T00:00:00`)
        }
      }
    }

    // Hora: viene en columna HORA separada como "21:19:00"
    let horaNum = null
    if (horaRaw && String(horaRaw).includes(':')) {
      horaNum = parseInt(String(horaRaw).split(':')[0])
    } else if (horaRaw !== '' && horaRaw !== undefined) {
      const h = parseInt(horaRaw)
      if (!isNaN(h) && h >= 0 && h <= 23) horaNum = h
    }
    // Si no hay hora separada, extraer de la fecha (pero en UC siempre viene en HORA)
    if (horaNum === null && fecha && !isNaN(fecha)) horaNum = fecha.getHours()

    const consulta = String(col('CONSULTA', 'consulta', 'pregunta', 'query', 'mensaje') || '').trim()
    const calidad = esConsultaInvalida(consulta)

    return {
      userId:       String(col('User ID', 'user id', 'userid') || ''),
      usuario:      String(col('USUARIO', 'usuario', 'alumno', 'estudiante') || ''),
      consulta,
      respuestaIA:  String(col('RESPUESTA DE LA IA', 'respuesta', 'answer') || '').trim(),
      urlCurso:     String(col('URL DEL CURSO', 'url') || ''),
      nombreCurso:  String(col('NOMBRE DEL CURSO', 'nombre del curso', 'curso', 'asignatura') || 'Sin curso'),
      codigoCurso:  String(col('CODIGO DEL CURSO', 'codigo del curso', 'codigo', 'code') || ''),
      nrc:          String(col('NRC', 'nrc', 'section', 'seccion') || ''),
      bloque:       String(col('BLOQUE', 'bloque', 'block') || ''),
      periodo:      String(col('PERIODO', 'periodo', 'period') || ''),
      campus:       String(col('CAMPUS', 'campus', 'sede') || ''),
      fechaRaw,
      fecha:        fecha && !isNaN(fecha) ? fecha : null,
      fechaStr:     fecha && !isNaN(fecha) ? fecha.toISOString().slice(0,10) : '',
      hora:         horaNum,
      diaSemana:    fecha && !isNaN(fecha) ? fecha.getDay() : null, // 0=Dom, 6=Sáb
      // Campos calculados después
      periodoBloque: '',
      semana: null,
      esInvalido: calidad.invalido,
      motivoInvalido: calidad.motivo,
      sinRespuesta: false,
      // Clasificación IA — lee del Sheet si ya existe
      tematica:    String(col("TEMATICA", "tematica", "temática") || "").trim(),
      subtema:     String(col("SUBTEMA", "subtema") || "").trim(),
      tipo:        String(col("TIPO", "tipo") || "").trim(),
      profundidad: String(col("PROFUNDIDAD", "profundidad") || "").trim(),
    }
  })
  .filter(r => r.consulta || r.userId) // mantener registros aunque sean inválidos (para calidad)
  .map(r => ({
    ...r,
    periodoBloque: r.periodo && r.bloque ? `${r.periodo} ${r.bloque}` : (r.periodo || 'Sin período'),
    sinRespuesta: !r.respuestaIA || r.respuestaIA.length === 0
  }))
}

// Calcula calendario académico dinámicamente: inicio = fecha min de cada periodo-bloque
function calcularCalendario(data, calendarioUC) {
  const base = calendarioUC || CALENDARIO_UC_DEFAULT
  const cal = {}
  Object.entries(base).forEach(([key, v]) => {
    cal[key] = { inicio: v.inicio, fin: v.fin, label: v.label, count: 0 }
  })
  data.forEach(r => {
    if (!r.periodoBloque || r.periodoBloque === 'Sin periodo') return
    if (cal[r.periodoBloque]) cal[r.periodoBloque].count++
  })
  return cal
}

// ═══════════════════════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════════════════════
function aggBy(data, field) {
  const m = {}
  data.forEach(r => {
    const k = r[field] || '—'
    m[k] = (m[k] || 0) + 1
  })
  return Object.entries(m).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value)
}

function aggByMulti(data, fields) {
  const m = {}
  data.forEach(r => {
    const key = fields.map(f => r[f] || '—').join('||')
    if (!m[key]) { m[key] = { total: 0 }; fields.forEach(f => m[key][f] = r[f] || '—') }
    m[key].total++
  })
  return Object.values(m).sort((a, b) => b.total - a.total)
}

const fmtNum = n => Number(n || 0).toLocaleString('es-PE')
const fmtPct = (n, total) => total > 0 ? `${((n / total) * 100).toFixed(1)}%` : '0%'

// ═══════════════════════════════════════════════════════════════════════════
// COMPONENTES BASE
// ═══════════════════════════════════════════════════════════════════════════
const Card = ({ children, style, borderLeft, padding = 20 }) => (
  <div style={{
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding,
    borderLeft: borderLeft ? `4px solid ${borderLeft}` : `1px solid ${C.border}`,
    ...style
  }}>{children}</div>
)

const SLabel = ({ children, style }) => (
  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: 16, ...style }}>
    {children}
  </div>
)

const TT = ({ formatter }) => (
  <Tooltip
    contentStyle={{ background: C.surface, border: `1px solid ${C.borderHi}`, borderRadius: 8, fontSize: 12, color: C.text }}
    labelStyle={{ color: C.text, fontWeight: 600 }}
    itemStyle={{ color: C.accent2 }}
    formatter={formatter}
    cursor={{ fill: C.border + '44' }}
  />
)

function DataTable({ headers, rows, maxH = 360, onRowClick, hoverable }) {
  const th = { padding: '10px 14px', textAlign: 'left', fontSize: 11, fontWeight: 700, color: C.muted, borderBottom: `1px solid ${C.border}`, whiteSpace: 'nowrap', textTransform: 'uppercase', letterSpacing: '0.05em', background: C.surface }
  const td = { padding: '10px 14px', fontSize: 12.5, color: C.textDim, borderBottom: `1px solid ${C.border}55` }
  return (
    <div style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: maxH, borderRadius: 8, border: `1px solid ${C.border}` }}>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead style={{ position: 'sticky', top: 0, zIndex: 1 }}>
          <tr>{headers.map((h, i) => <th key={i} style={th}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i}
                onClick={onRowClick ? () => onRowClick(row, i) : undefined}
                style={{
                  background: i % 2 === 0 ? 'transparent' : C.surface + '77',
                  cursor: hoverable || onRowClick ? 'pointer' : 'default',
                  transition: 'background 0.15s'
                }}
                onMouseEnter={e => { if (hoverable || onRowClick) e.currentTarget.style.background = C.borderHi + '55' }}
                onMouseLeave={e => { if (hoverable || onRowClick) e.currentTarget.style.background = i % 2 === 0 ? 'transparent' : C.surface + '77' }}>
              {row.map((cell, j) => <td key={j} style={td}>{cell}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// KPI grande con borde lateral de color
function KPI({ label, value, sub, subColor, color = C.accent }) {
  return (
    <div style={{
      background: C.card, border: `1px solid ${C.border}`, borderLeft: `4px solid ${color}`,
      borderRadius: 10, padding: '18px 22px', minWidth: 0
    }}>
      <div style={{ fontSize: 36, fontWeight: 700, color: C.text, lineHeight: 1, letterSpacing: '-0.02em' }}>
        {typeof value === 'number' ? fmtNum(value) : value}
      </div>
      <div style={{ fontSize: 12, color: C.muted, marginTop: 8, fontWeight: 500 }}>{label}</div>
      {sub && <div style={{ fontSize: 12, fontWeight: 700, color: subColor || color, marginTop: 6 }}>{sub}</div>}
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: RESUMEN
// ═══════════════════════════════════════════════════════════════════════════
function TabResumen({ fd, calendario }) {
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])
  const usuarios = useMemo(() => new Set(validas.map(r => r.userId).filter(Boolean)).size, [validas])
  const cursos = useMemo(() => new Set(validas.map(r => r.nombreCurso).filter(c => c && c !== 'Sin curso')).size, [validas])
  const sinResp = useMemo(() => validas.filter(r => r.sinRespuesta).length, [validas])

  const porPeriodoBloque = useMemo(() => {
    const m = {}
    validas.forEach(r => { m[r.periodoBloque] = (m[r.periodoBloque] || 0) + 1 })
    return Object.entries(m).sort((a,b)=>a[0].localeCompare(b[0])).map(([name, value]) => ({ name, value }))
  }, [validas])

  const cursoTop = useMemo(() => {
    const a = aggBy(validas, 'nombreCurso').filter(x => x.name && x.name !== 'Sin curso' && x.name !== '—')
    return a[0]
  }, [validas])
  const bloqueTop = useMemo(() => {
    const a = aggBy(validas, 'periodoBloque').filter(x => x.name && x.name !== '—' && x.name !== 'Sin período')
    return a[0]
  }, [validas])

  const periodosUnicos = [...new Set(validas.map(r => r.periodo).filter(Boolean))]
  const bloquesUnicos = [...new Set(validas.map(r => r.bloque).filter(Boolean))]
  const nrcsUnicos = new Set(validas.map(r => r.nrc).filter(Boolean)).size

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
        <KPI label="Consultas válidas" value={validas.length} color={C.accent} />
        <KPI label="Usuarios únicos" value={usuarios} color={C.green} />
        <KPI label="Cursos visibles" value={cursos} color={C.amber} />
        <KPI label="Sin respuesta" value={sinResp}
          sub={fmtPct(sinResp, validas.length)}
          subColor={C.red} color={C.red} />
      </div>

      {/* Uso por periodo + Observaciones */}
      <div style={{ display: 'grid', gridTemplateColumns: '1.4fr 1fr', gap: 16 }}>
        <Card>
          <SLabel>Uso por periodo y bloque</SLabel>
          <ResponsiveContainer width="100%" height={290}>
            <BarChart data={porPeriodoBloque} margin={{ top: 10, right: 10, bottom: 10, left: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
              <XAxis dataKey="name" tick={{ fill: C.textDim, fontSize: 11 }} />
              <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
              <TT />
              <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                {porPeriodoBloque.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <SLabel>Observaciones automáticas</SLabel>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            {cursoTop && (
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
                <div style={{ width: 3, alignSelf: 'stretch', background: C.accent, borderRadius: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 13, color: C.textDim, lineHeight: 1.5 }}>
                  Curso más activo: <span style={{ color: C.accent, fontWeight: 700 }}>{cursoTop.name}</span> con{' '}
                  <span style={{ color: C.accent, fontWeight: 700 }}>{fmtNum(cursoTop.value)}</span> consultas.
                </div>
              </div>
            )}
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
              <div style={{ width: 3, alignSelf: 'stretch', background: C.amber, borderRadius: 2, flexShrink: 0 }} />
              <div style={{ fontSize: 13, color: C.textDim, lineHeight: 1.5 }}>
                El <span style={{ color: C.amber, fontWeight: 700 }}>{fmtPct(sinResp, validas.length)}</span> ({fmtNum(sinResp)}) de las consultas no recibieron respuesta.
              </div>
            </div>
            {bloqueTop && (
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
                <div style={{ width: 3, alignSelf: 'stretch', background: C.accent, borderRadius: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 13, color: C.textDim, lineHeight: 1.5 }}>
                  Bloque más activo: <span style={{ color: C.accent, fontWeight: 700 }}>{bloqueTop.name}</span> con {fmtNum(bloqueTop.value)} consultas.
                </div>
              </div>
            )}

            <div style={{ marginTop: 6, padding: '14px 16px', background: C.surface, borderRadius: 8, border: `1px solid ${C.border}` }}>
              <div style={{ fontSize: 11, color: C.muted, fontWeight: 700, marginBottom: 8, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Contexto académico</div>
              <div style={{ fontSize: 12.5, color: C.textDim, lineHeight: 1.7 }}>
                Periodos: <strong style={{ color: C.accent }}>{periodosUnicos.join(', ') || '—'}</strong><br />
                Bloques: <strong style={{ color: C.accent }}>{bloquesUnicos.join(', ') || '—'}</strong><br />
                NRCs únicos: <strong style={{ color: C.accent }}>{nrcsUnicos}</strong>
              </div>
            </div>
          </div>
        </Card>
      </div>

      {/* Distribución por bloque (pie + tabla) */}
      <Card>
        <SLabel>Distribución por bloque</SLabel>
        <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 24, alignItems: 'center' }}>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={porPeriodoBloque} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} innerRadius={45}>
                {porPeriodoBloque.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} stroke={C.bg} strokeWidth={2} />)}
              </Pie>
              <TT />
            </PieChart>
          </ResponsiveContainer>
          <DataTable
            headers={['Periodo', 'Bloque', 'Consultas']}
            rows={validas.length > 0 ? Object.entries(
              validas.reduce((acc, r) => {
                const k = `${r.periodo || '—'}||${r.bloque || '—'}`
                acc[k] = (acc[k] || 0) + 1
                return acc
              }, {})
            ).sort((a,b)=>b[1]-a[1]).map(([k, v]) => {
              const [p, b] = k.split('||')
              return [p, b, fmtNum(v)]
            }) : []}
            maxH={220}
          />
        </div>
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: CURSOS
// ═══════════════════════════════════════════════════════════════════════════
function TabCursos({ fd }) {
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])
  const ranking = useMemo(() => {
    const m = {}
    const sinResp = {}
    validas.forEach(r => {
      const k = r.nombreCurso
      m[k] = (m[k] || 0) + 1
      if (r.sinRespuesta) sinResp[k] = (sinResp[k] || 0) + 1
    })
    return Object.entries(m)
      .filter(([k]) => k && k !== 'Sin curso' && k !== '—')
      .map(([k, v]) => ({ curso: k, total: v, sinResp: sinResp[k] || 0, pctSinResp: v > 0 ? ((sinResp[k]||0)/v)*100 : 0 }))
      .sort((a, b) => b.total - a.total)
  }, [validas])

  const top15 = ranking.slice(0, 15).map(r => ({ ...r, cursoCorto: r.curso.length > 38 ? r.curso.slice(0, 36) + '…' : r.curso }))

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <Card>
        <SLabel>Ranking de cursos por consultas (top 15)</SLabel>
        <ResponsiveContainer width="100%" height={Math.max(360, top15.length * 28)}>
          <BarChart data={top15} layout="vertical" margin={{ top: 10, right: 60, bottom: 10, left: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} horizontal={false} />
            <XAxis type="number" tick={{ fill: C.textDim, fontSize: 11 }} />
            <YAxis type="category" dataKey="cursoCorto" tick={{ fill: C.textDim, fontSize: 11 }} width={260} />
            <TT />
            <Bar dataKey="total" radius={[0, 6, 6, 0]} label={{ position: 'right', fill: C.textDim, fontSize: 11, fontWeight: 600 }}>
              {top15.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <SLabel>Detalle de todos los cursos</SLabel>
        <DataTable
          headers={['#', 'Curso', 'Consultas', 'Sin respuesta', '% Sin resp.']}
          rows={ranking.map((r, i) => [
            i + 1,
            r.curso,
            fmtNum(r.total),
            r.sinResp > 0 ? <span style={{ color: C.red }}>{fmtNum(r.sinResp)}</span> : '0',
            <span style={{ color: r.pctSinResp > 10 ? C.red : r.pctSinResp > 5 ? C.amber : C.muted }}>{r.pctSinResp.toFixed(1)}%</span>
          ])}
          maxH={500}
        />
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: NRC
// ═══════════════════════════════════════════════════════════════════════════
function TabNRC({ fd }) {
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])

  const indiceNRC = useMemo(() => {
    const m = {}
    validas.forEach(r => {
      if (!r.nrc) return
      if (!m[r.nrc]) m[r.nrc] = {
        nrc: r.nrc, curso: r.nombreCurso, codigo: r.codigoCurso, bloque: r.bloque,
        periodo: r.periodo, total: 0, usuarios: new Set(), sinResp: 0
      }
      m[r.nrc].total++
      if (r.userId) m[r.nrc].usuarios.add(r.userId)
      if (r.sinRespuesta) m[r.nrc].sinResp++
    })
    return Object.values(m).map(r => ({
      ...r,
      usuarios: r.usuarios.size,
      consPorUser: r.usuarios.size > 0 ? r.total / r.usuarios.size : 0,
      pctSinResp: r.total > 0 ? (r.sinResp / r.total) * 100 : 0,
      dependencia: r.usuarios.size > 0 ? r.total / r.usuarios.size : 0
    })).sort((a, b) => b.total - a.total)
  }, [validas])

  const promedio = indiceNRC.length > 0 ? (validas.length / indiceNRC.length).toFixed(1) : 0
  const masActivo = indiceNRC[0]?.total || 0

  const nivelDep = v => v >= 7 ? { label: 'Alta', color: C.red } : v >= 5 ? { label: 'Med', color: C.amber } : { label: 'Baja', color: C.green }
  const maxDep = Math.max(...indiceNRC.map(r => r.dependencia), 1)

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
        <KPI label="NRCs activos" value={indiceNRC.length} color={C.accent} />
        <KPI label="Consultas/NRC promedio" value={promedio} color={C.green} />
        <KPI label="NRC más activo" value={masActivo} color={C.amber} />
      </div>

      <Card>
        <SLabel>Índice de dependencia por NRC</SLabel>
        {/* Leyenda explicativa */}
        <div style={{ background: C.surface, borderRadius: 8, padding: '10px 14px', marginBottom: 14, fontSize: 12, color: C.muted, display: 'flex', gap: 24, flexWrap: 'wrap' }}>
          <span>📌 <strong style={{color:C.textDim}}>Cons./usr:</strong> promedio de consultas por estudiante en ese NRC</span>
          <span>📌 <strong style={{color:C.textDim}}>Dependencia:</strong> nivel de uso del Tutor IA. <strong style={{color:C.red}}>Alta</strong> ≥7 cons/usr · <strong style={{color:C.amber}}>Med</strong> 5-7 · <strong style={{color:C.green}}>Baja</strong> &lt;5</span>
          <span>📌 <strong style={{color:C.textDim}}>% Sin resp.:</strong> porcentaje de consultas que no recibieron respuesta del tutor</span>
        </div>
        <DataTable
          headers={['Curso', 'Código', 'NRC', 'Bloque', 'Periodo', 'Consultas', 'Usuarios', 'Cons./usr', 'Sin resp.', 'Dependencia']}
          rows={indiceNRC.map(r => {
            const dep = nivelDep(r.dependencia)
            const widthPct = (r.dependencia / maxDep) * 100
            return [
              r.curso, r.codigo, r.nrc, r.bloque, r.periodo,
              <span style={{ color: C.accent, fontWeight: 600 }}>{fmtNum(r.total)}</span>,
              fmtNum(r.usuarios),
              <span style={{ color: C.amber, fontWeight: 600 }}>{r.consPorUser.toFixed(1)}</span>,
              <span style={{ color: r.pctSinResp > 10 ? C.red : C.muted }}>{r.pctSinResp.toFixed(1)}%</span>,
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 130 }}>
                <div style={{ flex: 1, background: C.border, borderRadius: 4, height: 6 }}>
                  <div style={{ width: `${widthPct}%`, background: dep.color, height: '100%', borderRadius: 4 }} />
                </div>
                <span style={{ color: dep.color, fontWeight: 700, fontSize: 11 }}>{dep.label}</span>
              </div>
            ]
          })}
          maxH={600}
        />
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: TEMPORALIDAD - con heatmap día × hora
// ═══════════════════════════════════════════════════════════════════════════
function TabTemporalidad({ fd, calendario }) {
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])

  // Evolución semanal por periodo-bloque
  const { semanalData, periodoKeys } = useMemo(() => {
    // Calcular semana relativa al inicio de cada periodo-bloque
    const conSemana = validas.map(r => {
      if (!r.fecha || !r.periodoBloque) return null
      const sem = getSemanaNum(r.fecha, r.periodoBloque, calendario)
      if (!sem) return null
      return { ...r, semana: sem }
    }).filter(Boolean)

    const periodos = [...new Set(conSemana.map(r => r.periodoBloque))].sort()
    const weeks = {}
    conSemana.forEach(r => {
      if (!r.semana || r.semana < 1 || r.semana > 20) return
      const key = `Sem ${r.semana}`
      if (!weeks[key]) weeks[key] = { semana: key, _num: r.semana }
      weeks[key][r.periodoBloque] = (weeks[key][r.periodoBloque] || 0) + 1
    })
    return {
      semanalData: Object.values(weeks).sort((a, b) => a._num - b._num),
      periodoKeys: periodos
    }
  }, [validas, calendario])

  // Heatmap día x hora
  const { heatmap, maxHeat } = useMemo(() => {
    const grid = Array.from({ length: 7 }, () => Array(24).fill(0))
    validas.forEach(r => {
      if (r.diaSemana !== null && r.hora !== null && r.hora >= 0 && r.hora <= 23) {
        grid[r.diaSemana][r.hora]++
      }
    })
    let max = 0
    grid.forEach(row => row.forEach(v => { if (v > max) max = v }))
    return { heatmap: grid, maxHeat: max }
  }, [validas])

  const colorHeat = v => {
    if (v === 0) return C.surface
    const ratio = v / maxHeat
    const idx = Math.min(Math.floor(ratio * (HEAT_PAL.length - 1)) + 1, HEAT_PAL.length - 1)
    return HEAT_PAL[idx]
  }

  const dias = useMemo(() => {
    const m = {}
    validas.forEach(r => { if (r.fechaStr) m[r.fechaStr] = (m[r.fechaStr] || 0) + 1 })
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0])).map(([f, c]) => ({ fecha: f, consultas: c }))
  }, [validas])

  const diasNombre = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb']

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <Card>
        <SLabel>Evolución semanal por periodo académico</SLabel>
        {semanalData.length > 0 ? (
          <>
            <ResponsiveContainer width="100%" height={280}>
              <LineChart data={semanalData} margin={{ top: 10, right: 30, bottom: 10, left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                <XAxis dataKey="semana" tick={{ fill: C.textDim, fontSize: 12 }} />
                <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
                <TT />
                <Legend wrapperStyle={{ fontSize: 12, color: C.textDim, paddingTop: 12 }} />
                {periodoKeys.map((p, i) => (
                  <Line key={p} type="monotone" dataKey={p} stroke={PAL[i % PAL.length]} strokeWidth={2.5}
                    dot={{ r: 4, fill: PAL[i % PAL.length], stroke: C.bg, strokeWidth: 1.5 }}
                    activeDot={{ r: 6 }} connectNulls />
                ))}
              </LineChart>
            </ResponsiveContainer>
            <div style={{ background: C.surface, borderRadius: 8, padding: '10px 14px', marginTop: 12, fontSize: 12, color: C.muted }}>
              📅 <strong style={{color:C.textDim}}>Cómo se calculan las semanas:</strong> Sem 1 corresponde a la primera semana oficial de clases según el calendario académico de UC cargado desde Google Sheets. Las consultas fuera del rango oficial no se asignan a ninguna semana y no aparecen en este gráfico.
            </div>
          </>
        ) : <div style={{ color: C.muted, padding: 30, textAlign: 'center' }}>Sin datos temporales suficientes</div>}
      </Card>

      <Card>
        <SLabel>Mapa de calor — hora × día de la semana</SLabel>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ borderCollapse: 'separate', borderSpacing: 3, margin: '0 auto' }}>
            <thead>
              <tr>
                <th style={{ width: 36 }}></th>
                {Array.from({ length: 24 }, (_, i) => (
                  <th key={i} style={{ fontSize: 10, color: C.muted, fontWeight: 600, width: 30, height: 22 }}>{i}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {heatmap.map((row, dia) => (
                <tr key={dia}>
                  <td style={{ fontSize: 11, color: C.textDim, fontWeight: 600, paddingRight: 8, textAlign: 'right' }}>{diasNombre[dia]}</td>
                  {row.map((v, h) => (
                    <td key={h} title={`${diasNombre[dia]} ${h}:00 — ${v} consultas`}
                        style={{
                          width: 30, height: 24, background: colorHeat(v), borderRadius: 4,
                          textAlign: 'center', fontSize: 9, color: v > maxHeat * 0.5 ? '#fff' : C.muted,
                          fontWeight: 600, cursor: 'pointer', transition: 'transform 0.1s'
                        }}>
                      {v > maxHeat * 0.3 ? v : ''}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 14, fontSize: 11, color: C.muted }}>
          <span>Menor</span>
          {HEAT_PAL.map((c, i) => <div key={i} style={{ width: 24, height: 12, background: c, borderRadius: 3, border: `1px solid ${C.border}` }} />)}
          <span>Mayor actividad</span>
          <span style={{ marginLeft: 'auto', fontFamily: 'JetBrains Mono, monospace' }}>Máx: {fmtNum(maxHeat)}</span>
        </div>
      </Card>

      <Card>
        <SLabel>Consultas por día</SLabel>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={dias} margin={{ top: 10, right: 20, bottom: 10, left: 0 }}>
            <defs>
              <linearGradient id="gradDia" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={C.accent} stopOpacity={0.5}/>
                <stop offset="100%" stopColor={C.accent} stopOpacity={0}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="fecha" tick={{ fill: C.textDim, fontSize: 10 }} interval="preserveStartEnd" />
            <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
            <TT />
            <Area type="monotone" dataKey="consultas" stroke={C.accent} strokeWidth={2} fill="url(#gradDia)" />
          </AreaChart>
        </ResponsiveContainer>
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: USUARIOS
// ═══════════════════════════════════════════════════════════════════════════
function TabUsuarios({ fd }) {
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])

  const usuarios = useMemo(() => {
    const m = {}
    validas.forEach(r => {
      const k = r.userId || r.usuario
      if (!k) return
      if (!m[k]) m[k] = {
        userId: r.userId, usuario: r.usuario, total: 0,
        cursos: {}, sinResp: 0
      }
      m[k].total++
      m[k].cursos[r.nombreCurso] = (m[k].cursos[r.nombreCurso] || 0) + 1
      if (r.sinRespuesta) m[k].sinResp++
    })
    return Object.values(m).map(u => {
      const cursoTop = Object.entries(u.cursos).sort((a, b) => b[1] - a[1])[0]
      return { ...u, cursoPrincipal: cursoTop ? cursoTop[0] : '—' }
    }).sort((a, b) => b.total - a.total)
  }, [validas])

  const totalConsultas = validas.length
  const promedio = usuarios.length > 0 ? (totalConsultas / usuarios.length).toFixed(1) : 0
  const maxConsultas = usuarios[0]?.total || 0
  const unaSola = usuarios.filter(u => u.total === 1).length
  const recurrentes = usuarios.filter(u => u.total > 5).length
  const consRecurrentes = usuarios.filter(u => u.total > 5).reduce((sum, u) => sum + u.total, 0)
  const pctRecurrentes = totalConsultas > 0 ? ((consRecurrentes / totalConsultas) * 100).toFixed(1) : 0

  const top15 = usuarios.slice(0, 15).map(u => ({
    ...u,
    nombreCorto: (u.usuario || u.userId).length > 22 ? (u.usuario || u.userId).slice(0, 20) + '…' : (u.usuario || u.userId)
  }))

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1.4fr 1fr', gap: 16 }}>
        <Card>
          <SLabel>Top 15 usuarios por consultas</SLabel>
          <ResponsiveContainer width="100%" height={460}>
            <BarChart data={top15} layout="vertical" margin={{ top: 10, right: 40, bottom: 10, left: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} horizontal={false} />
              <XAxis type="number" tick={{ fill: C.textDim, fontSize: 11 }} />
              <YAxis type="category" dataKey="nombreCorto" tick={{ fill: C.textDim, fontSize: 10 }} width={160} />
              <TT />
              <Bar dataKey="total" radius={[0, 6, 6, 0]} label={{ position: 'right', fill: C.textDim, fontSize: 11, fontWeight: 600 }}>
                {top15.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <SLabel>Indicadores de uso</SLabel>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            {[
              { label: 'Usuarios únicos', value: fmtNum(usuarios.length), color: C.accent },
              { label: 'Promedio consultas / usuario', value: promedio, color: C.green },
              { label: 'Máx. consultas (1 usuario)', value: fmtNum(maxConsultas), color: C.amber },
              { label: 'Con 1 sola consulta', value: fmtNum(unaSola), color: C.purple },
              { label: 'Usuarios recurrentes (> 5)', value: fmtNum(recurrentes), color: C.cyan },
            ].map((m, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', paddingBottom: 12, borderBottom: i < 4 ? `1px solid ${C.border}` : 'none' }}>
                <span style={{ fontSize: 13, color: C.textDim }}>{m.label}</span>
                <span style={{ fontSize: 22, fontWeight: 700, color: m.color, fontFamily: 'JetBrains Mono, monospace' }}>{m.value}</span>
              </div>
            ))}
            <div style={{ padding: '12px 14px', background: C.surface, borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12.5, color: C.textDim, lineHeight: 1.5 }}>
              Los recurrentes concentran <strong style={{ color: C.accent }}>{pctRecurrentes}%</strong> del total de consultas.
            </div>
          </div>
        </Card>
      </div>

      <Card>
        <SLabel>Top 20 usuarios</SLabel>
        <DataTable
          headers={['#', 'Usuario', 'ID', 'Consultas', '% total', 'Curso principal', 'Sin resp.']}
          rows={usuarios.slice(0, 20).map((u, i) => [
            i + 1,
            <span style={{ color: C.accent, fontWeight: 600 }}>{u.usuario || '—'}</span>,
            <span style={{ color: C.muted, fontFamily: 'JetBrains Mono, monospace' }}>{u.userId}</span>,
            <span style={{ color: C.accent, fontWeight: 700 }}>{fmtNum(u.total)}</span>,
            <span style={{ color: C.muted }}>{((u.total / totalConsultas) * 100).toFixed(1)}%</span>,
            u.cursoPrincipal,
            u.sinResp > 0 ? <span style={{ color: C.red }}>{fmtNum(u.sinResp)}</span> : <span style={{ color: C.muted }}>0</span>
          ])}
          maxH={500}
        />
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: CALIDAD
// ═══════════════════════════════════════════════════════════════════════════
function TabCalidad({ fd, totalCargados, calendarioUC }) {
  const [mostrarFueraRango, setMostrarFueraRango] = useState(false)
  const sinResp = useMemo(() => fd.filter(r => !r.esInvalido && r.sinRespuesta), [fd])
  const malFormados = useMemo(() => fd.filter(r => r.esInvalido), [fd])
  const validas = useMemo(() => fd.filter(r => !r.esInvalido), [fd])
  const fueraRango = useMemo(() => validas.filter(r =>
    r.fecha && r.periodoBloque ? esFueraDeRango(r.fecha, r.periodoBloque, calendarioUC) : false
  ), [validas, calendarioUC])

  const usuariosUnicos = new Set(validas.map(r => r.userId).filter(Boolean)).size
  const nrcsUnicos = new Set(validas.map(r => r.nrc).filter(Boolean)).size
  const pctSinResp = validas.length > 0 ? (sinResp.length / validas.length * 100).toFixed(1) : 0
  const pctFueraRango = validas.length > 0 ? (fueraRango.length / validas.length * 100).toFixed(1) : 0

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SLabel>Registros sin respuesta</SLabel>
          <div style={{ display: 'flex', gap: 28, alignItems: 'baseline', marginBottom: 14 }}>
            <div>
              <div style={{ fontSize: 36, fontWeight: 700, color: C.red, fontFamily: 'JetBrains Mono, monospace', lineHeight: 1 }}>{fmtNum(sinResp.length)}</div>
              <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Total</div>
            </div>
            <div>
              <div style={{ fontSize: 36, fontWeight: 700, color: C.amber, fontFamily: 'JetBrains Mono, monospace', lineHeight: 1 }}>{pctSinResp}%</div>
              <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Del filtrado</div>
            </div>
          </div>
          <DataTable
            headers={['Usuario', 'Curso', 'NRC', 'Consulta']}
            rows={sinResp.slice(0, 50).map(r => [
              r.usuario || '—',
              r.nombreCurso,
              r.nrc,
              (r.consulta || '').slice(0, 60) + (r.consulta.length > 60 ? '…' : '')
            ])}
            maxH={280}
          />
        </Card>

        <Card>
          <SLabel>Registros mal formados</SLabel>
          <DataTable
            headers={['User ID', 'Campo', 'Valor', 'Motivo']}
            rows={malFormados.slice(0, 100).map(r => [
              <span style={{ fontFamily: 'JetBrains Mono, monospace' }}>{r.userId}</span>,
              'CONSULTA',
              <span style={{ color: C.amber, fontFamily: 'JetBrains Mono, monospace' }}>{r.consulta || '(vacío)'}</span>,
              <span style={{ color: C.muted }}>{r.motivoInvalido}</span>
            ])}
            maxH={420}
          />
        </Card>
      </div>

      {fueraRango.length > 0 && (
        <Card>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <div>
              <SLabel style={{ margin: 0 }}>Consultas fuera de rango académico</SLabel>
              <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
                Consultas con fecha fuera de los bloques registrados en el calendario UC. Pueden ser registros de prueba, errores de fecha o consultas en periodos no configurados aún.
              </div>
            </div>
            <button
              onClick={() => setMostrarFueraRango(!mostrarFueraRango)}
              style={{ background: mostrarFueraRango ? C.amber+'22' : C.surface, border: `1px solid ${mostrarFueraRango ? C.amber : C.border}`, color: mostrarFueraRango ? C.amber : C.textDim, borderRadius: 8, padding: '7px 16px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
              {mostrarFueraRango ? '▲ Ocultar detalle' : `▼ Ver ${fmtNum(fueraRango.length)} registros`}
            </button>
          </div>
          {mostrarFueraRango && (
            <>
              <DataTable
                headers={['Usuario', 'Curso', 'Fecha', 'Periodo-Bloque', 'Consulta']}
                rows={fueraRango.slice(0, 200).map(r => [
                  r.usuario || '—',
                  r.nombreCurso,
                  r.fechaStr,
                  r.periodoBloque,
                  (r.consulta || '').slice(0, 80) + (r.consulta.length > 80 ? '…' : '')
                ])}
                maxH={360}
              />
              {fueraRango.length > 200 && (
                <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>
                  Mostrando 200 de {fmtNum(fueraRango.length)} registros. Exporta a Excel para ver todos.
                </div>
              )}
            </>
          )}
        </Card>
      )}

      <Card>
        <SLabel>Resumen de calidad del dataset</SLabel>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 14 }}>
          <KPI label="Total cargados" value={totalCargados} color={C.accent} />
          <KPI label="Filtrados visibles" value={validas.length} color={C.green} />
          <KPI label="Sin respuesta" value={sinResp.length} color={C.red} />
          <KPI label="Mal formados" value={malFormados.length} color={C.amber} />
          <KPI label="Fuera de rango académico" value={fueraRango.length}
            sub={`${pctFueraRango}% del total`} subColor={C.amber} color={C.amber} />
          <KPI label="Usuarios únicos" value={usuariosUnicos} color={C.purple} />
          <KPI label="NRCs únicos" value={nrcsUnicos} color={C.cyan} />
        </div>
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: TEMÁTICAS IA - con paralelización, persistencia y filtrado previo
// ═══════════════════════════════════════════════════════════════════════════
const STORAGE_KEY = 'tutor_ia_clasificaciones_v1'

function TabTematicas({ data, fd, setData }) {
  const [estado, setEstado] = useState('idle') // idle | procesando | listo | error
  const [progreso, setProgreso] = useState({ actual: 0, total: 0, errores: 0 })
  const [errorMsg, setErrorMsg] = useState('')
  const [subtemaActivo, setSubtemaActivo] = useState(null) // para filtrar tabla al hacer clic

  // Recuperar clasificaciones guardadas al montar
  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}')
      if (Object.keys(saved).length > 0) {
        const newData = data.map(r => {
          if (r.tematica) return r
          const key = r.userId + '||' + r.consulta.slice(0, 100)
          const cached = saved[key]
          if (cached) return { ...r, ...cached }
          return r
        })
        const hayCambios = newData.some((r, i) => r.tematica !== data[i].tematica)
        if (hayCambios) setData(newData)
      }
    } catch (e) {
      console.warn('No se pudieron recuperar clasificaciones guardadas')
    }
  }, [])

  const clasificadas = useMemo(() => data.filter(r => r.tematica && !r.esInvalido), [data])
  const validasSinClasificar = useMemo(() => data.filter(r => !r.tematica && !r.esInvalido), [data])
  // Versión filtrada para visualizaciones (respeta filtros globales)
  const clasificadasFd = useMemo(() => (fd || data).filter(r => r.tematica && !r.esInvalido), [fd, data])

  // Visualizaciones usan clasificadasFd (respeta filtros globales)
  const temasMap = useMemo(() => {
    const m = {}
    clasificadasFd.forEach(r => {
      if (!m[r.subtema || 'Sin subtema']) m[r.subtema || 'Sin subtema'] = { subtema: r.subtema, tematica: r.tematica, count: 0 }
      m[r.subtema || 'Sin subtema'].count++
    })
    return Object.values(m).sort((a, b) => b.count - a.count)
  }, [clasificadasFd])

  const distribucion = useMemo(() => {
    const m = {}
    clasificadasFd.forEach(r => { m[r.tematica] = (m[r.tematica] || 0) + 1 })
    return Object.entries(m).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value)
  }, [clasificadasFd])

  const tematicaPorCurso = useMemo(() => {
    const m = {}
    clasificadasFd.forEach(r => {
      const k = `${r.tematica}||${r.nombreCurso}`
      if (!m[k]) m[k] = { tematica: r.tematica, curso: r.nombreCurso, count: 0, ejemplos: [] }
      m[k].count++
      if (m[k].ejemplos.length < 3 && r.consulta.length > 20) m[k].ejemplos.push(r.consulta)
    })
    return Object.values(m).sort((a, b) => b.count - a.count)
  }, [clasificadasFd])

  // Clasificación paralela (3 lotes simultáneos de 25 = 75 consultas en paralelo)
  const iniciarClasificacion = async () => {
    const objetivos = validasSinClasificar
    if (objetivos.length === 0) return

    setEstado('procesando')
    setErrorMsg('')
    const LOTE = 25
    const CONCURRENCIA = 3
    const total = objetivos.length
    let procesados = 0
    let errores = 0
    setProgreso({ actual: 0, total, errores: 0 })

    const newData = [...data]
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}')

    // Construir lotes
    const lotes = []
    for (let i = 0; i < objetivos.length; i += LOTE) {
      lotes.push(objetivos.slice(i, i + LOTE))
    }

    const procesarLote = async (lote) => {
      const payload = lote.map(r => ({ texto: r.consulta.slice(0, 250) }))
      try {
        const res = await fetch('/api/clasificar', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ consultas: payload })
        })
        if (!res.ok) {
          const err = await res.json().catch(() => ({}))
          throw new Error(err.error || `HTTP ${res.status}`)
        }
        const { clasificaciones } = await res.json()
        lote.forEach((r, j) => {
          const cl = clasificaciones?.[j]
          const idx = data.indexOf(r)
          if (cl && idx >= 0) {
            const update = {
              tematica: cl.tematica || 'Otros',
              subtema: cl.subtema || '',
              tipo: cl.tipo || '',
              profundidad: cl.profundidad || ''
            }
            newData[idx] = { ...newData[idx], ...update }
            // Guardar en cache
            const cacheKey = r.userId + '||' + r.consulta.slice(0, 100)
            saved[cacheKey] = update
          }
        })
        procesados += lote.length
      } catch (err) {
        errores += lote.length
        // marcar como Otros para continuar
        lote.forEach(r => {
          const idx = data.indexOf(r)
          if (idx >= 0) newData[idx] = { ...newData[idx], tematica: 'Sin clasificar', subtema: '', tipo: '', profundidad: '' }
        })
        procesados += lote.length
      }
    }

    // Procesar en grupos de CONCURRENCIA
    for (let i = 0; i < lotes.length; i += CONCURRENCIA) {
      const grupo = lotes.slice(i, i + CONCURRENCIA)
      await Promise.all(grupo.map(procesarLote))
      setProgreso({ actual: procesados, total, errores })
      setData([...newData])
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(saved)) } catch {}
      // Pausa entre grupos para respetar rate limits
      if (i + CONCURRENCIA < lotes.length) await new Promise(r => setTimeout(r, 600))
    }

    if (errores > total * 0.5) {
      setErrorMsg(`Demasiados errores: ${errores}/${total}. Verifica OPENAI_API_KEY en Vercel.`)
      setEstado('error')
    } else {
      setEstado('listo')
    }
  }

  const limpiarClasificaciones = () => {
    if (!confirm('¿Borrar todas las clasificaciones guardadas?')) return
    localStorage.removeItem(STORAGE_KEY)
    const newData = data.map(r => ({ ...r, tematica: '', subtema: '', tipo: '', profundidad: '' }))
    setData(newData)
    setEstado('idle')
    setProgreso({ actual: 0, total: 0, errores: 0 })
  }

  const exportarExcel = () => {
    const rows = clasificadas.map(r => ({
      'User ID': r.userId, 'Usuario': r.usuario, 'Consulta': r.consulta,
      'Temática': r.tematica, 'Subtema': r.subtema, 'Tipo': r.tipo, 'Profundidad': r.profundidad,
      'Curso': r.nombreCurso, 'NRC': r.nrc, 'Bloque': r.bloque, 'Periodo': r.periodo, 'Fecha': r.fechaStr
    }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Temáticas')
    XLSX.writeFile(wb, 'clasificacion_tematicas.xlsx')
  }

  // Si aún no hay clasificaciones, mostrar pantalla de inicio del proceso
  if (clasificadas.length === 0 && estado !== 'procesando') {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <Card style={{ padding: 36 }}>
          <div style={{ textAlign: 'center', maxWidth: 500, margin: '0 auto' }}>
            <div style={{ fontSize: 48, marginBottom: 14 }}>🤖</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: C.text, marginBottom: 8 }}>Clasificación automática de temáticas</div>
            <div style={{ fontSize: 13, color: C.textDim, marginBottom: 20, lineHeight: 1.6 }}>
              OpenAI analizará las <strong style={{ color: C.accent }}>{fmtNum(validasSinClasificar.length)}</strong> consultas válidas
              y las clasificará en categorías como Contenido del curso, Cálculo y métodos, Navegación del aula, etc.
              <br /><br />
              Las consultas mal formadas se omiten automáticamente. Tiempo estimado: <strong>{Math.ceil(validasSinClasificar.length / 75 * 4)} min</strong>.
            </div>
            {errorMsg && (
              <div style={{ background: C.red + '15', border: `1px solid ${C.red}55`, borderRadius: 8, padding: 12, marginBottom: 16, fontSize: 12, color: C.red }}>
                ⚠ {errorMsg}
              </div>
            )}
            <button onClick={iniciarClasificacion}
              style={{
                background: C.accent, border: 'none', color: '#fff', borderRadius: 10,
                padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer'
              }}>
              Iniciar clasificación con IA →
            </button>
          </div>
        </Card>
      </div>
    )
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {/* Panel de estado */}
      <Card>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>
              {estado === 'procesando' ? 'Clasificando con IA…' : 'Clasificación de temáticas'}
            </div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
              <span style={{ color: C.green }}>✓ {fmtNum(clasificadasFd.length)} clasificadas{clasificadasFd.length !== clasificadas.length ? <span style={{color:C.amber}}> (filtradas)</span> : ''}</span>
              {validasSinClasificar.length > 0 && <span style={{ color: C.amber }}> · {fmtNum(validasSinClasificar.length)} pendientes</span>}
              {progreso.errores > 0 && <span style={{ color: C.red }}> · {fmtNum(progreso.errores)} errores</span>}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            {clasificadas.length > 0 && (
              <>
                <button onClick={exportarExcel}
                  style={{ background: C.surface, border: `1px solid ${C.border}`, color: C.accent2, borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                  ↓ Exportar Excel
                </button>
                <button onClick={limpiarClasificaciones}
                  style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 8, padding: '8px 14px', fontSize: 12, cursor: 'pointer' }}>
                  Limpiar
                </button>
              </>
            )}
            {validasSinClasificar.length > 0 && estado !== 'procesando' && (
              <button onClick={iniciarClasificacion}
                style={{ background: C.accent, border: 'none', color: '#fff', borderRadius: 8, padding: '8px 18px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                Clasificar {fmtNum(validasSinClasificar.length)} pendientes
              </button>
            )}
          </div>
        </div>

        {estado === 'procesando' && (
          <div style={{ marginTop: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 8 }}>
              <span style={{ color: C.textDim }}>Procesando en paralelo (3 lotes simultáneos)</span>
              <span style={{ color: C.accent, fontWeight: 700, fontFamily: 'JetBrains Mono, monospace' }}>
                {fmtNum(progreso.actual)} / {fmtNum(progreso.total)} ({Math.round((progreso.actual/progreso.total)*100)}%)
              </span>
            </div>
            <div style={{ background: C.border, borderRadius: 4, height: 8, overflow: 'hidden' }}>
              <div style={{ background: `linear-gradient(90deg, ${C.accent}, ${C.green})`, height: '100%', width: `${(progreso.actual/progreso.total)*100}%`, transition: 'width 0.4s ease' }} />
            </div>
          </div>
        )}
      </Card>

      {clasificadas.length > 0 && (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <Card>
              <SLabel>Distribución temática general</SLabel>
              <ResponsiveContainer width="100%" height={320}>
                <PieChart>
                  <Pie data={distribucion} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={100}
                    label={({ value, percent }) => `${(percent*100).toFixed(0)}%`}
                    labelLine={{ stroke: C.muted, strokeWidth: 1 }}>
                    {distribucion.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} stroke={C.bg} strokeWidth={2} />)}
                  </Pie>
                  <TT />
                </PieChart>
              </ResponsiveContainer>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 12 }}>
                {distribucion.map((d, i) => (
                  <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: C.textDim }}>
                    <div style={{ width: 10, height: 10, background: PAL[i % PAL.length], borderRadius: 2 }} />
                    {d.name}: <strong style={{ color: C.text }}>{fmtNum(d.value)}</strong>
                  </div>
                ))}
              </div>
            </Card>

            <Card>
              <SLabel>Temas más consultados</SLabel>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 8 }}>
                Haz clic en un tema para filtrar la tabla de abajo
              </div>
              <div style={{ maxHeight: 360, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 6 }}>
                {temasMap.slice(0, 30).map((t, i) => {
                  const activo = subtemaActivo === (t.subtema || 'Sin subtema')
                  return (
                    <div key={i}
                      onClick={() => setSubtemaActivo(activo ? null : (t.subtema || 'Sin subtema'))}
                      style={{
                        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                        padding: '10px 14px', borderRadius: 8, cursor: 'pointer', transition: 'all 0.15s',
                        background: activo ? C.accent + '22' : C.surface,
                        border: `1px solid ${activo ? C.accent : C.border}`,
                      }}
                      onMouseEnter={e => { if (!activo) e.currentTarget.style.background = C.borderHi + '44' }}
                      onMouseLeave={e => { if (!activo) e.currentTarget.style.background = C.surface }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: activo ? C.accent2 : C.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {t.subtema || 'Sin subtema'}
                        </div>
                        <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{t.tematica}</div>
                      </div>
                      <div style={{ background: activo ? C.accent : C.accent + '22', color: activo ? '#fff' : C.accent2, padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 700, transition: 'all 0.15s' }}>
                        {fmtNum(t.count)}
                      </div>
                    </div>
                  )
                })}
              </div>
              {subtemaActivo && (
                <button onClick={() => setSubtemaActivo(null)}
                  style={{ marginTop: 8, background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 6, padding: '5px 12px', fontSize: 11, cursor: 'pointer' }}>
                  ✕ Quitar filtro
                </button>
              )}
            </Card>
          </div>

          <Card>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
              <div>
                <SLabel style={{ margin: 0 }}>Tabla temática × curso</SLabel>
                {subtemaActivo && (
                  <div style={{ fontSize: 12, color: C.accent, marginTop: 4 }}>
                    🔍 Filtrando por subtema: <strong>"{subtemaActivo}"</strong>
                    <button onClick={() => setSubtemaActivo(null)}
                      style={{ marginLeft: 10, background: 'transparent', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 11 }}>
                      ✕ quitar
                    </button>
                  </div>
                )}
              </div>
            </div>
            <DataTable
              headers={['Temática', 'Curso', 'Consultas', 'Ejemplo']}
              rows={(() => {
                // Si hay subtema activo, filtrar clasificadas por ese subtema y agrupar por temática+curso
                if (subtemaActivo) {
                  const filtradas = clasificadasFd.filter(r => (r.subtema || 'Sin subtema') === subtemaActivo)
                  const m = {}
                  filtradas.forEach(r => {
                    const k = `${r.tematica}||${r.nombreCurso}`
                    if (!m[k]) m[k] = { tematica: r.tematica, curso: r.nombreCurso, count: 0, ejemplos: [] }
                    m[k].count++
                    if (m[k].ejemplos.length < 2 && r.consulta.length > 10) m[k].ejemplos.push(r.consulta)
                  })
                  return Object.values(m).sort((a,b) => b.count - a.count).map(r => [
                    <span style={{ color: C.accent, fontWeight: 600 }}>{r.tematica}</span>,
                    r.curso,
                    <span style={{ color: C.accent2, fontWeight: 700 }}>{fmtNum(r.count)}</span>,
                    <span style={{ color: C.muted, fontStyle: 'italic' }}>"{(r.ejemplos[0] || '').slice(0, 80)}{(r.ejemplos[0] || '').length > 80 ? '…' : ''}"</span>
                  ])
                }
                return tematicaPorCurso.slice(0, 100).map(r => [
                  <span style={{ color: C.accent, fontWeight: 600 }}>{r.tematica}</span>,
                  r.curso,
                  <span style={{ color: C.accent2, fontWeight: 700 }}>{fmtNum(r.count)}</span>,
                  <span style={{ color: C.muted, fontStyle: 'italic' }}>"{(r.ejemplos[0] || '').slice(0, 80)}{r.ejemplos[0]?.length > 80 ? '…' : ''}"</span>
                ])
              })()}
              maxH={500}
            />
          </Card>
        </>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// TAB: COMPARATIVA — comparación entre periodos y bloques
// ═══════════════════════════════════════════════════════════════════════════
function TabComparativa({ data, fd, calendarioUC }) {
  const [ejeX, setEjeX] = useState('periodoBloque')
  const [metrica, setMetrica] = useState('consultas')

  // Usa fd (filtrado) para visualizaciones, data completo para clasificación
  const validas = useMemo(() => (fd || data || []).filter(r => !r.esInvalido), [fd, data])

  // Datos por periodo-bloque
  const porPeriodo = useMemo(() => {
    const m = {}
    validas.forEach(r => {
      const k = r.periodoBloque || 'Sin período'
      if (!m[k]) m[k] = { label: k, consultas: 0, usuarios: new Set(), sinResp: 0, tematicas: {}, cursos: new Set() }
      m[k].consultas++
      if (r.userId) m[k].usuarios.add(r.userId)
      if (r.sinRespuesta) m[k].sinResp++
      if (r.tematica) m[k].tematicas[r.tematica] = (m[k].tematicas[r.tematica] || 0) + 1
      if (r.nombreCurso) m[k].cursos.add(r.nombreCurso)
    })
    return Object.values(m).map(r => ({
      ...r,
      usuarios: r.usuarios.size,
      cursos: r.cursos.size,
      pctSinResp: r.consultas > 0 ? +((r.sinResp / r.consultas) * 100).toFixed(1) : 0,
      promConsUser: r.usuarios.size > 0 ? +(r.consultas / r.usuarios.size).toFixed(1) : 0,
      tematicaTop: Object.entries(r.tematicas).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—'
    })).sort((a,b) => a.label.localeCompare(b.label))
  }, [validas])

  // Datos semana a semana por periodo (para gráfico)
  const semanalComp = useMemo(() => {
    const periodos = [...new Set(validas.map(r => r.periodoBloque).filter(Boolean))].sort()
    const weeks = {}
    validas.forEach(r => {
      if (!r.fecha || !r.periodoBloque) return
      const sem = getSemanaNum(r.fecha, r.periodoBloque, calendarioUC)
      if (!sem || sem < 1 || sem > 10) return
      const key = `Sem ${sem}`
      if (!weeks[key]) weeks[key] = { semana: key, _num: sem }
      weeks[key][r.periodoBloque] = (weeks[key][r.periodoBloque] || 0) + 1
    })
    return { data: Object.values(weeks).sort((a,b) => a._num - b._num), periodos }
  }, [validas, calendarioUC])

  // Temáticas por periodo
  const tematicasComp = useMemo(() => {
    const tematicas = [...new Set(validas.map(r => r.tematica).filter(Boolean))].sort()
    const periodos = [...new Set(validas.map(r => r.periodoBloque).filter(Boolean))].sort()
    const m = {}
    validas.forEach(r => {
      if (!r.tematica || !r.periodoBloque) return
      if (!m[r.tematica]) m[r.tematica] = { tematica: r.tematica }
      m[r.tematica][r.periodoBloque] = (m[r.tematica][r.periodoBloque] || 0) + 1
    })
    return { data: Object.values(m).sort((a,b) => {
      const totalA = periodos.reduce((s,p) => s + (a[p]||0), 0)
      const totalB = periodos.reduce((s,p) => s + (b[p]||0), 0)
      return totalB - totalA
    }), periodos }
  }, [validas])

  const metricas = [
    { id: 'consultas', label: 'Consultas totales' },
    { id: 'usuarios', label: 'Usuarios únicos' },
    { id: 'pctSinResp', label: '% Sin respuesta' },
    { id: 'promConsUser', label: 'Promedio cons./usuario' },
    { id: 'cursos', label: 'Cursos activos' },
  ]

  const btnStyle = (active) => ({
    padding: '6px 14px', fontSize: 12, fontWeight: active ? 700 : 500,
    background: active ? C.accent+'22' : 'transparent',
    border: `1px solid ${active ? C.accent : C.border}`,
    color: active ? C.accent2 : C.muted,
    borderRadius: 6, cursor: 'pointer'
  })

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>

      {/* Tabla resumen comparativo */}
      <Card>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 14, flexWrap: 'wrap', gap: 10 }}>
          <div>
            <SLabel style={{ margin: 0 }}>Comparativa por periodo-bloque</SLabel>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>Métricas clave de cada bloque académico para identificar tendencias y variaciones entre periodos.</div>
          </div>
        </div>
        <DataTable
          headers={['Periodo-Bloque', 'Consultas', 'Usuarios', 'Prom. cons./usr', '% Sin resp.', 'Cursos activos', 'Temática principal']}
          rows={porPeriodo.map(r => [
            <strong style={{ color: C.accent }}>{r.label}</strong>,
            <span style={{ color: C.text, fontWeight: 700, fontFamily: 'JetBrains Mono, monospace' }}>{fmtNum(r.consultas)}</span>,
            fmtNum(r.usuarios),
            <span style={{ color: r.promConsUser >= 7 ? C.red : r.promConsUser >= 5 ? C.amber : C.green, fontWeight: 600 }}>{r.promConsUser}</span>,
            <span style={{ color: r.pctSinResp > 10 ? C.red : r.pctSinResp > 5 ? C.amber : C.muted }}>{r.pctSinResp}%</span>,
            fmtNum(r.cursos),
            <span style={{ color: C.muted, fontSize: 11 }}>{r.tematicaTop}</span>
          ])}
          maxH={300}
        />
      </Card>

      {/* Gráfico de barras comparativo por métrica */}
      <Card>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
          <SLabel style={{ margin: 0 }}>Comparativa visual por métrica</SLabel>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {metricas.map(m => (
              <button key={m.id} style={btnStyle(metrica === m.id)} onClick={() => setMetrica(m.id)}>{m.label}</button>
            ))}
          </div>
        </div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={porPeriodo} margin={{ top: 10, right: 20, bottom: 10, left: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
            <XAxis dataKey="label" tick={{ fill: C.textDim, fontSize: 11 }} />
            <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
            <TT />
            <Bar dataKey={metrica} radius={[6,6,0,0]} label={{ position: 'top', fill: C.textDim, fontSize: 11, fontWeight: 600 }}>
              {porPeriodo.map((_, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>
          📌 Selecciona una métrica arriba para cambiar el gráfico. El promedio cons./usuario indica el nivel de dependencia al Tutor IA por bloque.
        </div>
      </Card>

      {/* Evolución semanal comparativa */}
      <Card>
        <SLabel>Evolución semanal — comparativa entre periodos</SLabel>
        <div style={{ fontSize: 12, color: C.muted, marginBottom: 14 }}>
          Muestra cómo varía el volumen de consultas semana a semana en cada periodo-bloque. Útil para identificar picos de demanda (semanas de exámenes, entregas, etc.).
        </div>
        {semanalComp.data.length > 0 ? (
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={semanalComp.data} margin={{ top: 10, right: 30, bottom: 10, left: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
              <XAxis dataKey="semana" tick={{ fill: C.textDim, fontSize: 12 }} />
              <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
              <TT />
              <Legend wrapperStyle={{ fontSize: 11, color: C.textDim, paddingTop: 12 }} />
              {semanalComp.periodos.map((p, i) => (
                <Line key={p} type="monotone" dataKey={p} stroke={PAL[i % PAL.length]}
                  strokeWidth={2.5} dot={{ r: 4, stroke: C.bg, strokeWidth: 1.5 }}
                  activeDot={{ r: 6 }} connectNulls />
              ))}
            </LineChart>
          </ResponsiveContainer>
        ) : (
          <div style={{ color: C.muted, padding: 30, textAlign: 'center', fontSize: 13 }}>Sin datos semanales disponibles</div>
        )}
      </Card>

      {/* Temáticas por periodo */}
      <Card>
        <SLabel>Temáticas por periodo-bloque</SLabel>
        <div style={{ fontSize: 12, color: C.muted, marginBottom: 14 }}>
          Distribución de temáticas de consulta en cada bloque académico. Permite detectar si los temas varían entre periodos o si hay temáticas recurrentes.
        </div>
        <DataTable
          headers={['Temática', ...tematicasComp.periodos, 'Total']}
          rows={tematicasComp.data.slice(0, 20).map(r => {
            const total = tematicasComp.periodos.reduce((s,p) => s + (r[p]||0), 0)
            return [
              <span style={{ color: C.accent, fontWeight: 600 }}>{r.tematica}</span>,
              ...tematicasComp.periodos.map(p => (
                r[p] ? <span style={{ color: C.textDim, fontFamily: 'JetBrains Mono, monospace' }}>{fmtNum(r[p])}</span> : <span style={{ color: C.border }}>—</span>
              )),
              <strong style={{ color: C.text }}>{fmtNum(total)}</strong>
            ]
          })}
          maxH={460}
        />
      </Card>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// CARGA DESDE GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════════════════════
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSv6HB5-J_8Gqh9RctyZRLPlq8Wi8UVUl59HxdW2WiOHGV15WNr6ddLmcf0VOoLxWhhkd4Ncp6il1_g/pub?output=csv'

async function cargarDesdeSheet() {
  const res = await fetch(SHEET_CSV_URL)
  if (!res.ok) throw new Error(`HTTP ${res.status} al acceder al Sheet`)
  const text = await res.text()
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true, skipEmptyLines: true,
      complete: r => resolve(normalizeRows(r.data)),
      error: e => reject(e)
    })
  })
}

function PantallaCarga({ onLoad }) {
  const [estado, setEstado] = useState('cargando')
  const [errorMsg, setErrorMsg] = useState('')
  const [dragging, setDragging] = useState(false)
  const fileRef = useRef()

  useEffect(() => {
    const timeout = setTimeout(() => { setErrorMsg('La carga tardó demasiado'); setEstado('error') }, 20000)
    cargarDesdeSheet()
      .then(rows => { clearTimeout(timeout); onLoad(rows, { name: 'Google Sheets — UC', source: 'sheets' }) })
      .catch(e => { clearTimeout(timeout); setErrorMsg(e.message); setEstado('error') })
    return () => clearTimeout(timeout)
  }, [])

  const handleFile = file => parseFile(file, rows => onLoad(rows, { name: file.name, source: 'file' }))

  if (estado === 'cargando') {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.bg }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.accent, textTransform: 'uppercase', letterSpacing: '0.15em', marginBottom: 16 }}>Universidad Continental</div>
          <div style={{ fontSize: 26, fontWeight: 700, color: C.text, marginBottom: 8 }}>Dashboard Tutor IA</div>
          <div style={{ fontSize: 13, color: C.muted, marginBottom: 32 }}>Cargando datos desde Google Sheets…</div>
          <div style={{ display: 'flex', justifyContent: 'center', gap: 8 }}>
            {[0,1,2].map(i => (
              <div key={i} style={{ width: 10, height: 10, borderRadius: '50%', background: C.accent,
                animation: `pulse 1.4s ease-in-out ${i * 0.2}s infinite` }} />
            ))}
          </div>
          <style>{`@keyframes pulse { 0%,80%,100%{transform:scale(0.4);opacity:0.3} 40%{transform:scale(1);opacity:1} }`}</style>
        </div>
      </div>
    )
  }

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, background: C.bg }}>
      <div style={{ width: '100%', maxWidth: 520 }}>
        <div style={{ textAlign: 'center', marginBottom: 24 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.accent, textTransform: 'uppercase', letterSpacing: '0.15em', marginBottom: 8 }}>Universidad Continental</div>
          <h1 style={{ fontSize: 24, fontWeight: 700, color: C.text }}>Dashboard Tutor IA</h1>
        </div>
        {errorMsg && (
          <div style={{ background: C.red + '15', border: `1px solid ${C.red}55`, borderRadius: 10, padding: 14, fontSize: 12, color: C.red, marginBottom: 20 }}>
            ⚠ {errorMsg}
            <div style={{ marginTop: 6, color: C.muted }}>Verifica que el Sheet esté publicado como CSV.</div>
          </div>
        )}
        <div onDragOver={e => { e.preventDefault(); setDragging(true) }} onDragLeave={() => setDragging(false)}
          onDrop={e => { e.preventDefault(); setDragging(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]) }}
          onClick={() => fileRef.current.click()}
          style={{ border: `2px dashed ${dragging ? C.accent : C.border}`, borderRadius: 14, padding: '32px 24px', textAlign: 'center', cursor: 'pointer', background: dragging ? C.accent + '08' : C.card }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📂</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>Cargar archivo Excel o CSV</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════════════
// APP PRINCIPAL CON FILTROS GLOBALES
// ═══════════════════════════════════════════════════════════════════════════
const TABS = [
  { id: 'resumen',      label: 'Resumen' },
  { id: 'cursos',       label: 'Cursos' },
  { id: 'nrc',          label: 'NRC' },
  { id: 'temporalidad', label: 'Temporalidad' },
  { id: 'tematicas',    label: 'Temáticas' },
  { id: 'usuarios',     label: 'Usuarios' },
  { id: 'comparativa',  label: '📊 Comparativa' },
  { id: 'calidad',      label: 'Calidad' },
]

export default function App() {
  const [data, setData] = useState(null)
  const [fileInfo, setFileInfo] = useState(null)
  const [tab, setTab] = useState('resumen')
  const [calendarioUC, setCalendarioUC] = useState(CALENDARIO_UC_DEFAULT)
  const [calendarioCargado, setCalendarioCargado] = useState(false)

  // Cargar calendario desde Google Sheets al montar
  useEffect(() => {
    cargarCalendarioDesdeSheet().then(cal => {
      setCalendarioUC(cal)
      setCalendarioCargado(true)
    })
  }, [])

  // Filtros globales
  const [fPeriodo, setFPeriodo] = useState('todos')
  const [fBloque, setFBloque] = useState('todos')
  const [fSemana, setFSemana] = useState('todas')
  const [fCurso, setFCurso] = useState('todos')
  const [fNrc, setFNrc] = useState('')

  const handleLoad = useCallback((rows, info) => {
    setData(rows)
    setFileInfo({ ...info, total: rows.length, loaded: new Date() })
  }, [])

  // Calendario académico con fechas reales de UC
  const calendario = useMemo(() => data ? calcularCalendario(data, calendarioUC) : {}, [data, calendarioUC])

  // Opciones de filtros
  const periodos = useMemo(() => data ? [...new Set(data.map(r => r.periodo).filter(Boolean))].sort() : [], [data])
  const bloques = useMemo(() => data ? [...new Set(data.map(r => r.bloque).filter(Boolean))].sort() : [], [data])
  const cursos = useMemo(() => data ? [...new Set(data.map(r => r.nombreCurso).filter(c => c && c !== 'Sin curso'))].sort() : [], [data])

  // Aplicar filtros
  const fd = useMemo(() => {
    if (!data) return []
    let r = data
    if (fPeriodo !== 'todos') r = r.filter(x => x.periodo === fPeriodo)
    if (fBloque !== 'todos')  r = r.filter(x => x.bloque === fBloque)
    if (fCurso !== 'todos')   r = r.filter(x => x.nombreCurso === fCurso)
    if (fNrc.trim())          r = r.filter(x => x.nrc.includes(fNrc.trim()) || x.codigoCurso.toLowerCase().includes(fNrc.trim().toLowerCase()))
    if (fSemana !== 'todas') {
      const semNum = parseInt(fSemana)
      r = r.filter(x => {
        if (!x.fecha || !x.periodoBloque) return false
        return getSemanaNum(x.fecha, x.periodoBloque, calendarioUC) === semNum
      })
    }
    return r
  }, [data, fPeriodo, fBloque, fCurso, fNrc, fSemana, calendario])

  const totalCargados = data?.length || 0

  if (!data) return <PantallaCarga onLoad={handleLoad} />

  const limpiar = () => {
    setFPeriodo('todos'); setFBloque('todos'); setFSemana('todas'); setFCurso('todos'); setFNrc('')
  }

  const selStyle = (a) => ({
    padding: '12px 18px', fontSize: 13, fontWeight: a ? 700 : 500,
    color: a ? C.accent2 : C.muted,
    background: 'transparent', border: 'none',
    borderBottom: `2px solid ${a ? C.accent : 'transparent'}`,
    cursor: 'pointer', whiteSpace: 'nowrap', transition: 'all 0.15s'
  })

  const inputStyle = { background: C.surface, border: `1px solid ${C.border}`, color: C.text, borderRadius: 6, padding: '6px 10px', fontSize: 12, minWidth: 0 }

  return (
    <div style={{ minHeight: '100vh', background: C.bg }}>
      {/* Header */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: '14px 24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14 }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>Análisis de consultas del Tutor IA</div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
              {fileInfo?.source === 'sheets' ? '🟢' : '📁'} {fileInfo?.name} · {fmtNum(totalCargados)} registros · {new Date().toLocaleTimeString('es-PE')}
          {calendarioCargado && <span style={{ marginLeft: 12, color: C.green }}>📅 Calendario UC cargado ({Object.keys(calendarioUC).length} bloques)</span>}
            </div>
          </div>
          <button onClick={() => { setData(null); setFileInfo(null) }}
            style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.textDim, borderRadius: 6, padding: '6px 14px', fontSize: 12, cursor: 'pointer' }}>
            ↺ Recargar fuente
          </button>
        </div>

        {/* Filtros */}
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <label style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>Periodo</label>
          <select value={fPeriodo} onChange={e => setFPeriodo(e.target.value)} style={inputStyle}>
            <option value="todos">Todos</option>
            {periodos.map(p => <option key={p} value={p}>{p}</option>)}
          </select>

          <label style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>Bloque</label>
          <select value={fBloque} onChange={e => setFBloque(e.target.value)} style={inputStyle}>
            <option value="todos">Todos</option>
            {bloques.map(b => <option key={b} value={b}>{b}</option>)}
          </select>

          <label style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>Semana</label>
          <select value={fSemana} onChange={e => setFSemana(e.target.value)} style={inputStyle}>
            <option value="todas">Todas</option>
            {(() => {
              // Calcular máximo de semanas según bloque seleccionado
              let maxSem = 8
              if (fBloque !== 'todos' && fPeriodo !== 'todos') {
                const key = `${fPeriodo} ${fBloque}`
                const cal = calendarioUC[key]
                if (cal) maxSem = Math.ceil((cal.fin - cal.inicio) / (7 * 86400000))
              }
              return Array.from({ length: Math.min(maxSem, 10) }, (_, i) => i + 1).map(s =>
                <option key={s} value={s}>Sem {s}</option>
              )
            })()}
          </select>

          <label style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>Curso</label>
          <select value={fCurso} onChange={e => setFCurso(e.target.value)} style={{ ...inputStyle, minWidth: 180 }}>
            <option value="todos">Todos</option>
            {cursos.map(c => <option key={c} value={c}>{c.length > 40 ? c.slice(0, 38) + '…' : c}</option>)}
          </select>

          <input value={fNrc} onChange={e => setFNrc(e.target.value)} placeholder="NRC o código…" style={{ ...inputStyle, width: 130 }} />

          <button onClick={limpiar}
            style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.textDim, borderRadius: 6, padding: '6px 14px', fontSize: 12, cursor: 'pointer', marginLeft: 'auto' }}>
            Limpiar
          </button>
          <div style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>{fmtNum(fd.length)} registros</div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, display: 'flex', overflowX: 'auto', paddingLeft: 18 }}>
        {TABS.map(t => (
          <button key={t.id} style={selStyle(tab === t.id)} onClick={() => setTab(t.id)}>{t.label}</button>
        ))}
      </div>

      {/* Contenido */}
      <div style={{ padding: '20px 24px', maxWidth: 1400, margin: '0 auto' }}>
        {tab === 'resumen'       && <TabResumen fd={fd} calendario={calendario} />}
        {tab === 'cursos'        && <TabCursos fd={fd} />}
        {tab === 'nrc'           && <TabNRC fd={fd} />}
        {tab === 'temporalidad'  && <TabTemporalidad fd={fd} calendario={calendario} />}
        {tab === 'tematicas'     && <TabTematicas data={data} fd={fd} setData={setData} />}
        {tab === 'usuarios'      && <TabUsuarios fd={fd} />}
        {tab === 'comparativa'   && <TabComparativa data={data} fd={fd} calendarioUC={calendarioUC} />}
        {tab === 'calidad'       && <TabCalidad fd={fd} totalCargados={totalCargados} calendarioUC={calendarioUC} />}
      </div>
    </div>
  )
}

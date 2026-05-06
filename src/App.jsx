import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, CartesianGrid, Legend
} from "recharts";

// ─── Constantes ───────────────────────────────────────────────────────────────
const PAL = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899','#84cc16','#f97316','#a855f7'];
const DAYS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
const ACAD = {
  '202520-VV1':{ label:'202520-A', inicio:new Date(2025,7,18), fin:new Date(2025,9,12) },
  '202520-VV2':{ label:'202520-B', inicio:new Date(2025,9,13), fin:new Date(2025,11,7) },
  '202600-VV1':{ label:'202600-A', inicio:new Date(2026,0,2),  fin:new Date(2026,1,26) },
  '202610-VV1':{ label:'202610-A', inicio:new Date(2026,2,16), fin:new Date(2026,4,10) },
  '202610-VV2':{ label:'202610-B', inicio:new Date(2026,4,18), fin:new Date(2026,6,12) },
};

// ─── Helpers de datos ─────────────────────────────────────────────────────────
function parseFecha(s) {
  const p = String(s).split('-');
  return p.length === 3 ? new Date(+p[0], +p[1]-1, +p[2]) : null;
}
function getSemana(periodo, bloque, fechaStr) {
  const ac = ACAD[`${periodo}-${bloque}`];
  if (!ac) return null;
  const d = parseFecha(fechaStr);
  if (!d) return null;
  const w = Math.floor((d - ac.inicio) / (7*24*3600*1000)) + 1;
  return w >= 1 && w <= 10 ? w : null;
}
function classify(text) {
  if (!text || text.length < 3) return { g:'Sin datos', e:'Sin datos' };
  const t = text.toLowerCase();
  if (/^(hola|buenos|buenas|gracias|ok\b|bien\b|saludos)/.test(t.trim()) && t.length < 45) return { g:'Interacción general', e:'Saludo/Agradecimiento' };
  if (/error|no funciona|no carga|tecnico|falla/.test(t)) return { g:'Soporte técnico', e:'Problema técnico' };
  if (/(d[oó]nde|como).*(encuentro|est[aá]|queda|accedo|ver)|no encuentro|no aparece/.test(t)) return { g:'Navegación del aula', e:'Ubicación de contenidos' };
  if (/actividad|autoevaluaci|examen|evaluaci[oó]|tarea|entrega|nota\b|calificaci|quiz|plazo/.test(t)) return { g:'Actividades y evaluaciones', e:'Evaluaciones y tareas' };
  if (/material|lectura|video|recurso|archivo|gu[ií]a|semana \d|unidad \d|rise/.test(t)) return { g:'Materiales y recursos', e:'Recursos del curso' };
  if (/calcul|m[eé]todo|f[oó]rmula|ejercicio|punto medio|resolver|ejemplo/.test(t)) return { g:'Cálculo y métodos', e:'Resolución de ejercicios' };
  if (/oferta|demanda|precio|mercado|bien inferior|inflac|elasticidad|monopolio|equilibrio/.test(t)) return { g:'Contenido del curso', e:'Economía y mercados' };
  if (/redacc|comprens|escritura|p[aá]rrafo|argumentaci|ensayo|ortogr/.test(t)) return { g:'Contenido del curso', e:'Comprensión y textos' };
  if (/psicolog|conducta|cognitiv|emoci[oó]|terapia|empat|cl[ií]nic/.test(t)) return { g:'Contenido del curso', e:'Psicología' };
  if (/liderazgo|innovaci[oó]|emprendimiento|gesti[oó]n|l[ií]der/.test(t)) return { g:'Contenido del curso', e:'Liderazgo e innovación' };
  if (/matem[aá]tic|algebra|funci[oó]n|ecuaci[oó]|vector|matriz/.test(t)) return { g:'Contenido del curso', e:'Matemática' };
  if (/derecho|jur[ií]dic|ley\b|norma\b|contrato|legal|constituc/.test(t)) return { g:'Contenido del curso', e:'Derecho' };
  if (/ingenier[ií]a|industrial|proceso|producci[oó]n/.test(t)) return { g:'Contenido del curso', e:'Ingeniería' };
  if (/sociolog|antropolog|cultura|sociedad|comunidad/.test(t)) return { g:'Contenido del curso', e:'Sociología/Antropología' };
  if (/digital|tecnolog|herramienta|software|plataforma/.test(t)) return { g:'Contenido del curso', e:'Tecnología digital' };
  return { g:'Otros', e:'Otros' };
}
function processRows(rows) {
  return rows.map(r => {
    const q  = String(r['CONSULTA'] || '');
    const rp = String(r['RESPUESTA DE LA IA'] || '');
    const sinR = !rp.trim() || rp === 'undefined' || rp === 'null';
    const cl = classify(q);
    const per = String(r['PERIODO'] || '');
    const bloque = String(r['BLOQUE'] || '').trim();
    const fechaStr = String(r['FECHA'] || '');
    const semana = getSemana(per, bloque, fechaStr);
    const acadKey = `${per}-${bloque}`;
    const periodoLabel = ACAD[acadKey]?.label || `${per} ${bloque}`;
    const dObj = parseFecha(fechaStr);
    const lenCat = q.length < 100 ? 'Corta (<100)' : q.length < 400 ? 'Media (100–400)' : 'Larga (>400)';
    return {
      userId: r['User ID'],
      usuario: String(r['USUARIO'] || '').trim(),
      consulta: q, sinR,
      nombreCurso: String(r['NOMBRE DEL CURSO'] || '').trim(),
      codigoCurso: String(r['CODIGO DEL CURSO'] || '').trim(),
      nrc: String(r['NRC'] || '').trim(),
      bloque, periodo: per, año: per.substring(0,4),
      campus: String(r['CAMPUS'] || '').trim(),
      hora: parseInt(String(r['HORA'] || '0').split(':')[0]) || 0,
      fechaStr, semana, periodoLabel,
      dayOfWeek: dObj ? dObj.getDay() : null,
      consultaLen: q.length, lenCat,
      tg: cl.g, te: cl.e,
    };
  }).filter(r => r.consulta.length > 0 && r.consulta !== 'undefined');
}
function aggBy(data, keys) {
  const m = {};
  data.forEach(r => {
    const k = keys.map(k2 => r[k2]).join('||');
    if (!m[k]) { m[k] = { total:0 }; keys.forEach(k2 => { m[k][k2] = r[k2]; }); }
    m[k].total++;
  });
  return Object.values(m).sort((a,b) => b.total - a.total);
}

// ─── UI Primitivos ────────────────────────────────────────────────────────────
const Card = ({ children, style={} }) => (
  <div style={{ background:'#1e293b', borderRadius:12, padding:'16px 20px', ...style }}>{children}</div>
);
const SLabel = ({ children }) => (
  <div style={{ fontSize:11, fontWeight:600, color:'#64748b', textTransform:'uppercase', letterSpacing:'0.07em', marginBottom:12 }}>{children}</div>
);
const KPI = ({ label, value, sub, color }) => (
  <div style={{ background:'#1e293b', borderRadius:10, padding:'14px 16px', borderLeft:`3px solid ${color}` }}>
    <div style={{ fontSize:28, fontWeight:700, color:'#f1f5f9', lineHeight:1 }}>{value}</div>
    {sub && <div style={{ fontSize:12, color, marginTop:2 }}>{sub}</div>}
    <div style={{ fontSize:11, color:'#64748b', marginTop:3 }}>{label}</div>
  </div>
);
const TT = (p) => <Tooltip {...p} contentStyle={{ background:'#0f172a', border:'1px solid #334155', borderRadius:8, color:'#e2e8f0', fontSize:12 }} />;
const TH = { padding:'7px 10px', background:'#0f172a', color:'#64748b', textAlign:'left', fontWeight:600, fontSize:11, borderBottom:'1px solid #1e293b' };
const TD = { padding:'6px 10px', borderBottom:'1px solid #0f172a', color:'#cbd5e1', verticalAlign:'top' };

function DTable({ headers, rows }) {
  return (
    <div style={{ overflowX:'auto' }}>
      <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
        <thead><tr>{headers.map(h => <th key={h} style={TH}>{h}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ background: i%2===0 ? '#141c2b' : '#1e293b' }}>
              {row.map((cell, j) => <td key={j} style={TD}>{cell}</td>)}
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan={headers.length} style={{ ...TD, textAlign:'center', color:'#475569' }}>Sin datos</td></tr>}
        </tbody>
      </table>
    </div>
  );
}

// ─── Alertas automáticas ──────────────────────────────────────────────────────
function Alerts({ fd, filters }) {
  const alerts = useMemo(() => {
    if (!fd.length) return [];
    const list = [];
    const sinRn = fd.filter(r => r.sinR).length;
    const pct = (sinRn / fd.length) * 100;
    if (pct > 10) list.push({ type:'warning', msg:`El ${pct.toFixed(1)}% de las consultas (${sinRn}) no tiene respuesta de la IA.` });
    const nrcMap = {};
    fd.forEach(r => {
      if (!nrcMap[r.nrc]) nrcMap[r.nrc] = { total:0, sinR:0 };
      nrcMap[r.nrc].total++;
      if (r.sinR) nrcMap[r.nrc].sinR++;
    });
    const badNRCs = Object.entries(nrcMap).filter(([,v]) => v.total >= 5 && v.sinR/v.total > 0.25);
    if (badNRCs.length > 0) list.push({ type:'warning', msg:`${badNRCs.length} NRC(s) con más del 25% sin respuesta: ${badNRCs.slice(0,3).map(([n])=>n).join(', ')}.` });
    if (filters.periodo !== 'Todos' || filters.bloque !== 'Todos') {
      const weeksWithData = new Set(fd.filter(r => r.semana).map(r => r.semana));
      if (weeksWithData.size > 0) {
        const missing = [1,2,3,4,5,6,7,8].filter(w => !weeksWithData.has(w));
        if (missing.length > 0) list.push({ type:'info', msg:`Semanas sin actividad: ${missing.map(w=>'Sem '+w).join(', ')}.` });
      }
    }
    return list;
  }, [fd, filters]);
  if (!alerts.length) return null;
  return (
    <div style={{ padding:'8px 24px', display:'flex', flexDirection:'column', gap:6, background:'#0f172a', borderBottom:'1px solid #1e293b' }}>
      {alerts.map((a,i) => (
        <div key={i} style={{ display:'flex', alignItems:'flex-start', gap:8, padding:'7px 12px', borderRadius:8,
          background: a.type==='warning' ? 'rgba(245,158,11,0.1)' : 'rgba(59,130,246,0.1)',
          border: `1px solid ${a.type==='warning' ? 'rgba(245,158,11,0.3)' : 'rgba(59,130,246,0.3)'}` }}>
          <span style={{ fontSize:14, flexShrink:0 }}>{a.type==='warning' ? '⚠️' : 'ℹ️'}</span>
          <span style={{ fontSize:12, color: a.type==='warning' ? '#fbbf24' : '#93c5fd', lineHeight:1.5 }}>{a.msg}</span>
        </div>
      ))}
    </div>
  );
}

// ─── Mapa de calor ────────────────────────────────────────────────────────────
function Heatmap({ fd }) {
  const { heatMap, maxVal } = useMemo(() => {
    const m = {}; let max = 0;
    fd.forEach(r => {
      if (r.dayOfWeek === null) return;
      const k = `${r.dayOfWeek}-${r.hora}`;
      m[k] = (m[k]||0) + 1;
      if (m[k] > max) max = m[k];
    });
    return { heatMap:m, maxVal:max };
  }, [fd]);
  const hours = Array.from({ length:24 }, (_,i) => i);
  const getColor = (d,h) => {
    const c = heatMap[`${d}-${h}`] || 0;
    if (c === 0) return '#0f172a';
    return `rgba(59,130,246,${0.15 + (c/maxVal)*0.85})`;
  };
  return (
    <div style={{ overflowX:'auto' }}>
      <div style={{ minWidth:660 }}>
        <div style={{ display:'flex', alignItems:'center', gap:2, marginBottom:4, paddingLeft:40 }}>
          {hours.map(h => <div key={h} style={{ width:24, fontSize:9, color:'#475569', textAlign:'center', flexShrink:0 }}>{h}</div>)}
        </div>
        {DAYS.map((day,d) => (
          <div key={d} style={{ display:'flex', alignItems:'center', gap:2, marginBottom:2 }}>
            <div style={{ width:36, fontSize:11, color:'#94a3b8', flexShrink:0, textAlign:'right', paddingRight:4 }}>{day}</div>
            {hours.map(h => (
              <div key={h} title={`${day} ${h}h: ${heatMap[`${d}-${h}`]||0} consultas`}
                style={{ width:24, height:20, borderRadius:3, background:getColor(d,h), flexShrink:0 }} />
            ))}
          </div>
        ))}
        <div style={{ display:'flex', alignItems:'center', gap:8, marginTop:10, paddingLeft:40 }}>
          <span style={{ fontSize:10, color:'#475569' }}>Menor</span>
          {[0.1,0.3,0.55,0.75,0.95].map(i => (
            <div key={i} style={{ width:16, height:12, borderRadius:2, background:`rgba(59,130,246,${i})` }} />
          ))}
          <span style={{ fontSize:10, color:'#475569' }}>Mayor actividad</span>
          {maxVal > 0 && <span style={{ fontSize:10, color:'#475569', marginLeft:8 }}>Máx: {maxVal}</span>}
        </div>
      </div>
    </div>
  );
}

// ─── Perfil de usuario ────────────────────────────────────────────────────────
function UserProfile({ user, fd }) {
  const userRows = useMemo(() => fd.filter(r => r.userId === user.userId), [user, fd]);
  const horaData = useMemo(() => {
    const m = {}; for (let i=0;i<24;i++) m[i]=0;
    userRows.forEach(r => { m[r.hora]++; });
    return Object.entries(m).filter(([,v])=>v>0).map(([h,c])=>({ hora:h+'h', consultas:c }));
  }, [userRows]);
  const tgData = useMemo(() => {
    const m = {};
    userRows.forEach(r => { m[r.tg]=(m[r.tg]||0)+1; });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({ name,value }));
  }, [userRows]);
  const cursosData = useMemo(() => aggBy(userRows, ['nombreCurso']), [userRows]);
  const avgLen = userRows.length > 0 ? Math.round(userRows.reduce((s,r)=>s+r.consultaLen,0)/userRows.length) : 0;
  return (
    <Card style={{ borderLeft:'3px solid #10b981', marginTop:12 }}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:14 }}>
        <div>
          <div style={{ fontSize:15, fontWeight:700, color:'#fff', marginBottom:2 }}>{user.usuario}</div>
          <div style={{ fontSize:11, color:'#475569' }}>ID: {user.userId} · {userRows.length} consultas · {cursosData.length} curso(s)</div>
        </div>
        <div style={{ display:'flex', gap:16 }}>
          {[{ v:userRows.length, l:'Consultas', c:'#3b82f6' },{ v:userRows.filter(r=>r.sinR).length, l:'Sin resp.', c:'#ef4444' },{ v:avgLen, l:'Chars prom.', c:'#f59e0b' }].map(it=>(
            <div key={it.l} style={{ textAlign:'center' }}>
              <div style={{ fontSize:20, fontWeight:700, color:it.c }}>{it.v}</div>
              <div style={{ fontSize:10, color:'#64748b' }}>{it.l}</div>
            </div>
          ))}
        </div>
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:16, marginBottom:14 }}>
        <div>
          <SLabel>Actividad por hora</SLabel>
          <ResponsiveContainer width="100%" height={120}>
            <BarChart data={horaData}>
              <XAxis dataKey="hora" tick={{ fill:'#64748b', fontSize:9 }} />
              <YAxis tick={{ fill:'#64748b', fontSize:9 }} />
              <TT />
              <Bar dataKey="consultas" fill="#3b82f6" radius={[2,2,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div>
          <SLabel>Temáticas</SLabel>
          <ResponsiveContainer width="100%" height={120}>
            <PieChart>
              <Pie data={tgData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={48}>
                {tgData.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Pie>
              <TT />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div>
          <SLabel>Cursos</SLabel>
          {cursosData.map((c,i) => (
            <div key={c.nombreCurso} style={{ display:'flex', justifyContent:'space-between', padding:'4px 0', borderBottom:'1px solid #1e293b', fontSize:11 }}>
              <span style={{ color:'#cbd5e1', maxWidth:130, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{c.nombreCurso}</span>
              <span style={{ color:PAL[i%PAL.length], fontWeight:600, flexShrink:0, marginLeft:6 }}>{c.total}</span>
            </div>
          ))}
        </div>
      </div>
      <SLabel>Últimas consultas</SLabel>
      <DTable
        headers={['Fecha','Hora','Curso','Temática','Consulta','Resp.']}
        rows={userRows.slice(0,10).map(r => [
          r.fechaStr, r.hora+'h', r.nombreCurso.substring(0,20), r.tg,
          r.consulta.substring(0,55)+'…',
          r.sinR ? <span style={{ color:'#ef4444' }}>✗</span> : <span style={{ color:'#10b981' }}>✓</span>
        ])}
      />
    </Card>
  );
}

// ─── Tabs ─────────────────────────────────────────────────────────────────────
function TabResumen({ fd }) {
  const pbRows = useMemo(() => aggBy(fd, ['periodo','bloque']).sort((a,b)=>a.periodo.localeCompare(b.periodo)), [fd]);
  const pbChart = pbRows.map(r => ({ name:`${r.periodo} ${r.bloque}`, total:r.total }));
  const topCurso = useMemo(() => aggBy(fd, ['nombreCurso']), [fd]);
  const sinRn = fd.filter(r=>r.sinR).length;
  const sinRpct = fd.length > 0 ? ((sinRn/fd.length)*100).toFixed(1) : 0;
  const topPB = [...pbRows].sort((a,b)=>b.total-a.total)[0];
  const pers = [...new Set(fd.map(r=>r.periodo))].sort();
  const blqs = [...new Set(fd.map(r=>r.bloque))].sort();
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
        <Card>
          <SLabel>Uso por periodo y bloque</SLabel>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={pbChart}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
              <XAxis dataKey="name" tick={{ fill:'#64748b', fontSize:11 }} />
              <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
              <TT />
              <Bar dataKey="total" radius={[4,4,0,0]}>
                {pbChart.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SLabel>Observaciones automáticas</SLabel>
          {topCurso.length > 0 && (
            <p style={{ margin:'0 0 10px', paddingLeft:10, borderLeft:'2px solid #3b82f6', fontSize:13, color:'#cbd5e1', lineHeight:1.7 }}>
              Curso más activo: <strong style={{ color:'#60a5fa' }}>{topCurso[0].nombreCurso}</strong> con <strong style={{ color:'#60a5fa' }}>{topCurso[0].total.toLocaleString()}</strong> consultas.
            </p>
          )}
          <p style={{ margin:'0 0 10px', paddingLeft:10, borderLeft:'2px solid #3b82f6', fontSize:13, color:'#cbd5e1', lineHeight:1.7 }}>
            El <strong style={{ color:'#f59e0b' }}>{sinRpct}%</strong> ({sinRn.toLocaleString()}) de las consultas no recibieron respuesta.
          </p>
          {topPB && (
            <p style={{ margin:'0 0 10px', paddingLeft:10, borderLeft:'2px solid #3b82f6', fontSize:13, color:'#cbd5e1', lineHeight:1.7 }}>
              Bloque más activo: <strong style={{ color:'#60a5fa' }}>{topPB.periodo} {topPB.bloque}</strong> con {topPB.total.toLocaleString()} consultas.
            </p>
          )}
          <div style={{ marginTop:14, padding:'10px 12px', background:'#0f172a', borderRadius:8 }}>
            <div style={{ fontSize:11, color:'#475569', marginBottom:6 }}>Contexto académico</div>
            <div style={{ fontSize:12, color:'#64748b', lineHeight:1.9 }}>
              Periodos: <span style={{ color:'#93c5fd' }}>{pers.join(', ') || '-'}</span><br />
              Bloques: <span style={{ color:'#93c5fd' }}>{blqs.join(', ') || '-'}</span><br />
              NRCs únicos: <span style={{ color:'#93c5fd' }}>{new Set(fd.map(r=>r.nrc)).size}</span>
            </div>
          </div>
        </Card>
      </div>
      <Card>
        <SLabel>Distribución por bloque</SLabel>
        <div style={{ display:'flex', gap:20, alignItems:'flex-start', flexWrap:'wrap' }}>
          <ResponsiveContainer width={200} height={160}>
            <PieChart>
              <Pie data={pbChart} dataKey="total" nameKey="name" cx="50%" cy="50%" outerRadius={70}>
                {pbChart.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Pie>
              <TT />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ flex:1, minWidth:180 }}>
            <DTable headers={['Periodo','Bloque','Consultas']} rows={pbRows.map(r=>[r.periodo,r.bloque,r.total.toLocaleString()])} />
          </div>
        </div>
      </Card>
    </div>
  );
}

function TabCursos({ fd }) {
  const ac = useMemo(() => aggBy(fd, ['nombreCurso','codigoCurso']), [fd]);
  const srMap = useMemo(() => {
    const m = {};
    fd.filter(r=>r.sinR).forEach(r => { m[r.nombreCurso]=(m[r.nombreCurso]||0)+1; });
    return m;
  }, [fd]);
  const chartData = ac.map(r => ({ name: r.nombreCurso.length>26 ? r.nombreCurso.substring(0,26)+'…' : r.nombreCurso, total:r.total }));
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <Card>
        <SLabel>Ranking de cursos</SLabel>
        <ResponsiveContainer width="100%" height={Math.max(260, ac.length*32+60)}>
          <BarChart data={chartData} layout="vertical" margin={{ left:10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
            <XAxis type="number" tick={{ fill:'#64748b', fontSize:11 }} />
            <YAxis dataKey="name" type="category" tick={{ fill:'#cbd5e1', fontSize:11 }} width={200} />
            <TT />
            <Bar dataKey="total" radius={[0,4,4,0]}>
              {chartData.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </Card>
      <Card>
        <SLabel>Tabla comparativa</SLabel>
        <DTable
          headers={['Curso','Código','Consultas','Sin resp.','% Sin resp.']}
          rows={ac.map(r => {
            const s = srMap[r.nombreCurso] || 0;
            return [r.nombreCurso, r.codigoCurso, r.total.toLocaleString(), s.toLocaleString(), `${r.total>0?((s/r.total)*100).toFixed(1):0}%`];
          })}
        />
      </Card>
    </div>
  );
}

function TabNRC({ fd }) {
  const rows = useMemo(() => {
    const m = {};
    fd.forEach(r => {
      if (!m[r.nrc]) m[r.nrc] = { nrc:r.nrc, curso:r.nombreCurso, codigo:r.codigoCurso, bloque:r.bloque, periodo:r.periodo, campus:r.campus, total:0, usuarios:new Set(), sinR:0 };
      m[r.nrc].total++;
      m[r.nrc].usuarios.add(r.userId);
      if (r.sinR) m[r.nrc].sinR++;
    });
    const allRows = Object.values(m).map(r => ({
      ...r,
      uCount: r.usuarios.size,
      ratio: r.usuarios.size > 0 ? (r.total/r.usuarios.size).toFixed(1) : '0',
      sinRPct: r.total > 0 ? ((r.sinR/r.total)*100).toFixed(1) : '0',
    })).sort((a,b) => b.total - a.total);
    const maxRatio = Math.max(...allRows.map(r => parseFloat(r.ratio)), 1);
    return allRows.map(r => ({ ...r, maxRatio }));
  }, [fd]);

  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:12 }}>
        {[
          { label:'NRCs activos', value:rows.length, color:'#3b82f6' },
          { label:'Consultas/NRC promedio', value:rows.length>0?(fd.length/rows.length).toFixed(1):0, color:'#10b981' },
          { label:'NRC más activo', value:rows[0]?.total||0, color:'#f59e0b' },
        ].map(it => (
          <div key={it.label} style={{ background:'#0f172a', borderRadius:8, padding:'10px 14px', borderLeft:`3px solid ${it.color}` }}>
            <div style={{ fontSize:20, fontWeight:700, color:'#f1f5f9' }}>{it.value}</div>
            <div style={{ fontSize:11, color:'#64748b', marginTop:2 }}>{it.label}</div>
          </div>
        ))}
      </div>
      <Card>
        <SLabel>Índice de dependencia por NRC</SLabel>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
            <thead><tr>{['Curso','Código','NRC','Bloque','Periodo','Consultas','Usuarios','Cons./usr','Sin resp.','Dependencia'].map(h => <th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map((r,i) => {
                const dep = parseFloat(r.ratio);
                const pct = r.maxRatio > 0 ? dep/r.maxRatio : 0;
                const col = pct > 0.7 ? '#ef4444' : pct > 0.4 ? '#f59e0b' : '#3b82f6';
                return (
                  <tr key={i} style={{ background: i%2===0 ? '#141c2b' : '#1e293b' }}>
                    <td style={TD}>{r.curso}</td>
                    <td style={TD}>{r.codigo}</td>
                    <td style={TD}>{r.nrc}</td>
                    <td style={TD}>{r.bloque}</td>
                    <td style={TD}>{r.periodo}</td>
                    <td style={{ ...TD, color:'#60a5fa', fontWeight:600 }}>{r.total}</td>
                    <td style={TD}>{r.uCount}</td>
                    <td style={{ ...TD, color: dep>3?'#f59e0b':dep>1.5?'#10b981':'#64748b', fontWeight:600 }}>{r.ratio}</td>
                    <td style={{ ...TD, color: parseFloat(r.sinRPct)>20?'#ef4444':'#64748b' }}>{r.sinRPct}%</td>
                    <td style={TD}>
                      <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                        <div style={{ flex:1, height:8, background:'#0f172a', borderRadius:4, overflow:'hidden', minWidth:50 }}>
                          <div style={{ height:'100%', width:`${pct*100}%`, background:col, borderRadius:4 }} />
                        </div>
                        <span style={{ fontSize:10, color:col, width:28 }}>{pct>0.7?'Alta':pct>0.4?'Med':'Baja'}</span>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop:10, fontSize:11, color:'#475569' }}>Alta (&gt;3 cons/usr) · Media (1.5–3) · Baja (&lt;1.5)</div>
      </Card>
    </div>
  );
}

function TabTemp({ fd }) {
  const diasData = useMemo(() => {
    const m = {};
    fd.forEach(r => { m[r.fechaStr]=(m[r.fechaStr]||0)+1; });
    return Object.entries(m).sort((a,b)=>a[0].localeCompare(b[0])).map(([f,c])=>({ fecha:f, consultas:c }));
  }, [fd]);
  const horasData = useMemo(() => {
    const m = {}; for (let i=0;i<24;i++) m[i]=0;
    fd.forEach(r => { m[r.hora]++; });
    return Object.entries(m).sort((a,b)=>+a[0]-+b[0]).map(([h,c])=>({ hora:h+'h', consultas:c }));
  }, [fd]);
  const { semanalData, periodoKeys } = useMemo(() => {
    const keys = [...new Set(fd.filter(r=>r.semana).map(r=>r.periodoLabel))].sort();
    const weeks = {};
    fd.forEach(r => {
      if (!r.semana) return;
      if (!weeks[r.semana]) weeks[r.semana] = { semana:`Sem ${r.semana}` };
      weeks[r.semana][r.periodoLabel] = (weeks[r.semana][r.periodoLabel]||0)+1;
    });
    return {
      semanalData: Object.values(weeks).sort((a,b)=>+a.semana.split(' ')[1]-+b.semana.split(' ')[1]),
      periodoKeys: keys,
    };
  }, [fd]);
  const lenData = useMemo(() => {
    const m = { 'Corta (<100)':0, 'Media (100–400)':0, 'Larga (>400)':0 };
    fd.forEach(r => { m[r.lenCat]=(m[r.lenCat]||0)+1; });
    return Object.entries(m).map(([name,value])=>({ name,value }));
  }, [fd]);
  const lenByCurso = useMemo(() => {
    const m = {};
    fd.forEach(r => {
      if (!m[r.nombreCurso]) m[r.nombreCurso]={ sum:0, n:0 };
      m[r.nombreCurso].sum += r.consultaLen;
      m[r.nombreCurso].n++;
    });
    return Object.entries(m).map(([nombre,d])=>({ nombre:nombre.length>24?nombre.substring(0,24)+'…':nombre, avg:Math.round(d.sum/d.n) })).sort((a,b)=>b.avg-a.avg);
  }, [fd]);
  const avgLen = fd.length > 0 ? Math.round(fd.reduce((s,r)=>s+r.consultaLen,0)/fd.length) : 0;
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <Card>
        <SLabel>Evolución semanal por periodo académico</SLabel>
        {semanalData.length > 0 ? (
          <>
            <ResponsiveContainer width="100%" height={240}>
              <LineChart data={semanalData} margin={{ right:20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                <XAxis dataKey="semana" tick={{ fill:'#64748b', fontSize:11 }} />
                <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
                <TT />
                <Legend wrapperStyle={{ fontSize:11, color:'#94a3b8', paddingTop:8 }} />
                {periodoKeys.map((p,i) => <Line key={p} type="monotone" dataKey={p} stroke={PAL[i%PAL.length]} strokeWidth={2} dot={{ r:3 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
            <div style={{ fontSize:11, color:'#475569', marginTop:8 }}>Sem 1 = primera semana de cada bloque según el calendario académico.</div>
          </>
        ) : (
          <div style={{ color:'#475569', fontSize:13, padding:'16px 0' }}>Aplica filtros de periodo o bloque para ver la evolución semanal.</div>
        )}
      </Card>
      <Card>
        <SLabel>Mapa de calor — hora × día de la semana</SLabel>
        <Heatmap fd={fd} />
      </Card>
      <Card>
        <SLabel>Consultas por día</SLabel>
        <ResponsiveContainer width="100%" height={190}>
          <LineChart data={diasData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
            <XAxis dataKey="fecha" tick={{ fill:'#64748b', fontSize:10 }} interval="preserveStartEnd" />
            <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
            <TT />
            <Line type="monotone" dataKey="consultas" stroke="#3b82f6" strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </Card>
      <Card>
        <SLabel>Consultas por hora del día</SLabel>
        <ResponsiveContainer width="100%" height={190}>
          <BarChart data={horasData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
            <XAxis dataKey="hora" tick={{ fill:'#64748b', fontSize:10 }} />
            <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
            <TT />
            <Bar dataKey="consultas" fill="#10b981" radius={[3,3,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
        <Card>
          <SLabel>Distribución por longitud de consulta</SLabel>
          <div style={{ display:'flex', gap:16, marginBottom:12, alignItems:'center' }}>
            <div style={{ textAlign:'center' }}>
              <div style={{ fontSize:24, fontWeight:700, color:'#f1f5f9' }}>{avgLen}</div>
              <div style={{ fontSize:11, color:'#64748b' }}>chars promedio</div>
            </div>
            <div style={{ flex:1 }}>
              {lenData.map((d,i) => (
                <div key={d.name} style={{ marginBottom:8 }}>
                  <div style={{ display:'flex', justifyContent:'space-between', fontSize:11, color:'#94a3b8', marginBottom:3 }}>
                    <span>{d.name}</span>
                    <span style={{ color:PAL[i] }}>{d.value} ({fd.length>0?((d.value/fd.length)*100).toFixed(0):0}%)</span>
                  </div>
                  <div style={{ height:8, background:'#0f172a', borderRadius:4, overflow:'hidden' }}>
                    <div style={{ height:'100%', width:`${fd.length>0?(d.value/fd.length)*100:0}%`, background:PAL[i], borderRadius:4 }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div style={{ fontSize:11, color:'#475569', lineHeight:1.6 }}>
            Consultas <strong style={{ color:'#10b981' }}>largas</strong> (&gt;400): preguntas conceptuales.<br />
            Consultas <strong style={{ color:'#3b82f6' }}>cortas</strong> (&lt;100): operativas o de navegación.
          </div>
        </Card>
        <Card>
          <SLabel>Longitud promedio por curso</SLabel>
          <ResponsiveContainer width="100%" height={Math.max(180, lenByCurso.length*28+40)}>
            <BarChart data={lenByCurso} layout="vertical" margin={{ left:5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
              <XAxis type="number" tick={{ fill:'#64748b', fontSize:10 }} />
              <YAxis dataKey="nombre" type="category" tick={{ fill:'#cbd5e1', fontSize:10 }} width={165} />
              <TT />
              <Bar dataKey="avg" name="Chars promedio" radius={[0,4,4,0]}>
                {lenByCurso.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  );
}

function TabTemas({ fd }) {
  const [selTema, setSelTema] = useState(null);
  const tgData = useMemo(() => {
    const m = {};
    fd.forEach(r => { m[r.tg]=(m[r.tg]||0)+1; });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({ name,value }));
  }, [fd]);
  const teData = useMemo(() => {
    const m = {};
    fd.forEach(r => {
      if (!m[r.te]) m[r.te] = { name:r.te, general:r.tg, value:0, cursos:{}, ejs:[] };
      m[r.te].value++;
      m[r.te].cursos[r.nombreCurso] = (m[r.te].cursos[r.nombreCurso]||0)+1;
      if (m[r.te].ejs.length < 5) m[r.te].ejs.push(r.consulta.substring(0,110));
    });
    return Object.values(m).sort((a,b)=>b.value-a.value);
  }, [fd]);
  const txc = useMemo(() => {
    const m = {};
    fd.forEach(r => {
      const k = `${r.tg}||${r.nombreCurso}`;
      if (!m[k]) m[k] = { tema:r.tg, curso:r.nombreCurso, total:0, ej:'' };
      m[k].total++;
      if (!m[k].ej) m[k].ej = r.consulta.substring(0,85);
    });
    return Object.values(m).sort((a,b)=>b.total-a.total).slice(0,30);
  }, [fd]);
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
        <Card>
          <SLabel>Distribución temática general</SLabel>
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={tgData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={88} label={({ percent }) => `${(percent*100).toFixed(0)}%`} labelLine={false}>
                {tgData.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Pie>
              <TT />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginTop:8 }}>
            {tgData.map((t,i) => (
              <div key={t.name} style={{ display:'flex', alignItems:'center', gap:5, fontSize:11 }}>
                <span style={{ width:8, height:8, borderRadius:2, background:PAL[i%PAL.length], display:'inline-block' }} />
                <span style={{ color:'#64748b' }}>{t.name}: </span>
                <span style={{ color:'#f1f5f9', fontWeight:600 }}>{t.value}</span>
              </div>
            ))}
          </div>
        </Card>
        <Card>
          <SLabel>Temas más consultados</SLabel>
          <div style={{ maxHeight:320, overflowY:'auto' }}>
            {teData.slice(0,15).map(t => (
              <div key={t.name} onClick={() => setSelTema(selTema?.name===t.name ? null : t)}
                style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'7px 10px', marginBottom:3, borderRadius:8, cursor:'pointer',
                  background: selTema?.name===t.name ? '#1e3a5f' : '#0f172a', border:'1px solid #1e293b' }}>
                <div>
                  <div style={{ fontSize:12, color:'#e2e8f0', fontWeight:500 }}>{t.name}</div>
                  <div style={{ fontSize:10, color:'#475569' }}>{t.general}</div>
                </div>
                <span style={{ background:'#1e3a5f', color:'#60a5fa', borderRadius:20, padding:'2px 9px', fontSize:11, fontWeight:700, flexShrink:0 }}>{t.value}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>
      {selTema && (
        <Card style={{ borderLeft:'3px solid #3b82f6' }}>
          <SLabel>Detalle: {selTema.name}</SLabel>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20 }}>
            <div>
              <div style={{ fontSize:11, color:'#475569', marginBottom:8 }}>Cursos relacionados</div>
              {Object.entries(selTema.cursos).sort((a,b)=>b[1]-a[1]).map(([c,n]) => (
                <div key={c} style={{ display:'flex', justifyContent:'space-between', padding:'4px 0', borderBottom:'1px solid #1e293b', fontSize:12 }}>
                  <span style={{ color:'#cbd5e1' }}>{c}</span>
                  <span style={{ color:'#60a5fa', fontWeight:600 }}>{n}</span>
                </div>
              ))}
            </div>
            <div>
              <div style={{ fontSize:11, color:'#475569', marginBottom:8 }}>Ejemplos{selTema.name==='Otros'?' (varios)':''}</div>
              {selTema.ejs.map((ej,i) => (
                <div key={i} style={{ padding:'6px 8px', background:'#0f172a', borderRadius:6, marginBottom:5, fontSize:11, color:'#94a3b8', fontStyle:'italic', lineHeight:1.5 }}>"{ej}…"</div>
              ))}
            </div>
          </div>
        </Card>
      )}
      <Card>
        <SLabel>Tabla temática × curso</SLabel>
        <DTable
          headers={['Temática','Curso','Consultas','Ejemplo']}
          rows={txc.map(r => [r.tema, r.curso, r.total.toLocaleString(), <span style={{ color:'#475569', fontStyle:'italic', fontSize:11 }}>"{r.ej}…"</span>])}
        />
      </Card>
    </div>
  );
}

function TabUsuarios({ fd }) {
  const [selUser, setSelUser] = useState(null);
  const total = fd.length;
  const topU = useMemo(() => {
    const m = {};
    fd.forEach(r => {
      if (!m[r.userId]) m[r.userId] = { userId:r.userId, usuario:r.usuario, total:0, cursos:{}, sinR:0 };
      m[r.userId].total++;
      m[r.userId].cursos[r.nombreCurso] = (m[r.userId].cursos[r.nombreCurso]||0)+1;
      if (r.sinR) m[r.userId].sinR++;
    });
    return Object.values(m).sort((a,b) => b.total - a.total);
  }, [fd]);
  const top15 = topU.slice(0,15);
  const chartData = top15.map(u => ({ name: u.usuario.length>20?u.usuario.substring(0,20)+'…':u.usuario, consultas:u.total }));
  const recurrentes = topU.filter(u=>u.total>5);
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
        <Card>
          <SLabel>Top 15 usuarios por consultas</SLabel>
          <ResponsiveContainer width="100%" height={Math.max(240, top15.length*30+60)}>
            <BarChart data={chartData} layout="vertical" margin={{ left:10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
              <XAxis type="number" tick={{ fill:'#64748b', fontSize:11 }} />
              <YAxis dataKey="name" type="category" tick={{ fill:'#cbd5e1', fontSize:10 }} width={170} />
              <TT />
              <Bar dataKey="consultas" radius={[0,4,4,0]}>
                {chartData.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SLabel>Indicadores de uso</SLabel>
          {[
            { label:'Usuarios únicos', value:topU.length.toLocaleString(), color:'#3b82f6' },
            { label:'Promedio consultas / usuario', value:topU.length>0?(total/topU.length).toFixed(1):0, color:'#10b981' },
            { label:'Máx. consultas (1 usuario)', value:top15[0]?.total||0, color:'#f59e0b' },
            { label:'Con 1 sola consulta', value:topU.filter(u=>u.total===1).length, color:'#8b5cf6' },
            { label:'Usuarios recurrentes (> 5)', value:recurrentes.length, color:'#06b6d4' },
          ].map(it => (
            <div key={it.label} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'8px 0', borderBottom:'1px solid #1e293b' }}>
              <span style={{ fontSize:13, color:'#94a3b8' }}>{it.label}</span>
              <span style={{ fontSize:16, fontWeight:700, color:it.color }}>{it.value}</span>
            </div>
          ))}
          <div style={{ padding:'10px 12px', background:'#0f172a', borderRadius:8, fontSize:12, color:'#64748b', lineHeight:1.7, marginTop:12 }}>
            Los recurrentes concentran <strong style={{ color:'#93c5fd' }}>
              {total>0?((recurrentes.reduce((s,u)=>s+u.total,0)/total)*100).toFixed(1):0}%
            </strong> del total de consultas.
          </div>
        </Card>
      </div>
      <Card>
        <SLabel>Top 20 usuarios — haz clic en una fila para ver el perfil</SLabel>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
            <thead><tr>{['#','Usuario','ID','Consultas','% total','Curso principal','Sin resp.'].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>
              {topU.slice(0,20).map((u,i) => {
                const cp = Object.entries(u.cursos).sort((a,b)=>b[1]-a[1])[0]?.[0]||'-';
                const isSel = selUser?.userId === u.userId;
                return (
                  <tr key={i} onClick={() => setSelUser(isSel ? null : u)}
                    style={{ background: isSel ? '#1e3a5f' : i%2===0 ? '#141c2b' : '#1e293b', cursor:'pointer' }}>
                    <td style={TD}>{i+1}</td>
                    <td style={{ ...TD, color:'#93c5fd', fontWeight:500 }}>{u.usuario}</td>
                    <td style={{ ...TD, color:'#475569' }}>{u.userId}</td>
                    <td style={{ ...TD, color:'#60a5fa', fontWeight:600 }}>{u.total}</td>
                    <td style={TD}>{total>0?((u.total/total)*100).toFixed(1):0}%</td>
                    <td style={TD}>{cp.length>28?cp.substring(0,28)+'…':cp}</td>
                    <td style={{ ...TD, color:u.sinR>0?'#ef4444':'#475569' }}>{u.sinR}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {selUser && <UserProfile user={selUser} fd={fd} />}
      </Card>
    </div>
  );
}

function TabCalidad({ fd, raw }) {
  const sinR = fd.filter(r=>r.sinR);
  const pct = fd.length>0?((sinR.length/fd.length)*100).toFixed(1):0;
  const mal = raw.filter(r=>!r.consulta||r.consulta.length<3||!r.nombreCurso||r.nombreCurso==='undefined');
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:16 }}>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
        <Card>
          <SLabel>Registros sin respuesta</SLabel>
          <div style={{ display:'flex', gap:24, marginBottom:14 }}>
            <div style={{ textAlign:'center' }}>
              <div style={{ fontSize:32, fontWeight:700, color:'#ef4444' }}>{sinR.length.toLocaleString()}</div>
              <div style={{ fontSize:11, color:'#64748b' }}>Total</div>
            </div>
            <div style={{ textAlign:'center' }}>
              <div style={{ fontSize:32, fontWeight:700, color:'#f59e0b' }}>{pct}%</div>
              <div style={{ fontSize:11, color:'#64748b' }}>Del filtrado</div>
            </div>
          </div>
          <div style={{ maxHeight:220, overflowY:'auto' }}>
            <DTable headers={['Usuario','Curso','NRC','Consulta']}
              rows={sinR.slice(0,50).map(r=>[r.usuario.substring(0,18),r.nombreCurso.substring(0,18),r.nrc,r.consulta.substring(0,45)+'…'])} />
          </div>
        </Card>
        <Card>
          <SLabel>Registros mal formados</SLabel>
          {mal.length===0
            ? <div style={{ color:'#10b981', textAlign:'center', padding:20, fontSize:13 }}>✓ Sin registros mal formados</div>
            : <DTable headers={['User ID','Campo','Valor','Motivo']}
                rows={mal.slice(0,25).map(r=>{const c=!r.consulta||r.consulta.length<3?'CONSULTA':'NOMBRE DEL CURSO';return[r.userId,c,String(c==='CONSULTA'?r.consulta:r.nombreCurso).substring(0,25),'Valor vacío o inválido'];})} />
          }
        </Card>
      </div>
      <Card>
        <SLabel>Resumen de calidad del dataset</SLabel>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:12 }}>
          {[
            { label:'Total cargados', value:raw.length, color:'#3b82f6' },
            { label:'Filtrados visibles', value:fd.length, color:'#10b981' },
            { label:'Sin respuesta', value:sinR.length, color:'#ef4444' },
            { label:'Mal formados', value:mal.length, color:'#f59e0b' },
            { label:'Usuarios únicos', value:new Set(fd.map(r=>r.userId)).size, color:'#8b5cf6' },
            { label:'NRCs únicos', value:new Set(fd.map(r=>r.nrc)).size, color:'#06b6d4' },
          ].map(it => (
            <div key={it.label} style={{ background:'#0f172a', borderRadius:8, padding:'10px 14px', borderLeft:`3px solid ${it.color}` }}>
              <div style={{ fontSize:22, fontWeight:700, color:'#f1f5f9' }}>{it.value.toLocaleString()}</div>
              <div style={{ fontSize:11, color:'#64748b', marginTop:2 }}>{it.label}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ─── Vista Comparaciones — TODOS los hooks antes de cualquier return ──────────
function Comparaciones({ raw, onBack }) {
  // ⚠️ TODOS los useMemo deben ir ANTES de cualquier return condicional
  const pbChart = useMemo(() => {
    if (!raw || !raw.length) return [];
    const m = {};
    raw.forEach(r => { const k=`${r.periodo} ${r.bloque}`; m[k]=(m[k]||0)+1; });
    return Object.entries(m).sort((a,b)=>a[0].localeCompare(b[0])).map(([name,total])=>({ name,total }));
  }, [raw]);

  const periodos = useMemo(() => {
    if (!raw || !raw.length) return [];
    return [...new Set(raw.map(r=>r.periodo))].filter(Boolean).sort();
  }, [raw]);

  const cursos = useMemo(() => {
    if (!raw || !raw.length) return [];
    return [...new Set(raw.map(r=>r.nombreCurso))].filter(Boolean).sort();
  }, [raw]);

  const cxp = useMemo(() => {
    if (!raw || !raw.length) return {};
    const m = {};
    raw.forEach(r => {
      if (!m[r.nombreCurso]) m[r.nombreCurso] = {};
      m[r.nombreCurso][r.periodo] = (m[r.nombreCurso][r.periodo]||0)+1;
    });
    return m;
  }, [raw]);

  const blqChart = useMemo(() => {
    if (!raw || !raw.length) return [];
    const m = {};
    raw.forEach(r => { if(r.bloque) m[r.bloque]=(m[r.bloque]||0)+1; });
    return Object.entries(m).sort((a,b)=>a[0].localeCompare(b[0])).map(([name,total])=>({ name,total }));
  }, [raw]);

  // Temáticas × periodo: usamos periodos como valor calculado, no como dep de useMemo encadenado
  const temaTabla = useMemo(() => {
    if (!raw || !raw.length) return { rows:[], pers:[] };
    const pers = [...new Set(raw.map(r=>r.periodo))].filter(Boolean).sort();
    const m = {};
    raw.forEach(r => {
      if (!r.tg || !r.periodo) return;
      if (!m[r.tg]) m[r.tg] = { tg:r.tg };
      m[r.tg][r.periodo] = (m[r.tg][r.periodo]||0)+1;
    });
    const rows = Object.values(m).sort((a,b) => {
      const sa = pers.reduce((s,p)=>s+(a[p]||0),0);
      const sb = pers.reduce((s,p)=>s+(b[p]||0),0);
      return sb-sa;
    });
    return { rows, pers };
  }, [raw]);

  // ─── A partir de aquí sí podemos renderizar condicionalmente ───────────────
  if (!raw || !raw.length) {
    return (
      <div style={{ fontFamily:'system-ui,sans-serif', background:'#0d1b2a', minHeight:'100vh', color:'#e2e8f0', padding:'20px 24px' }}>
        <button onClick={onBack} style={{ background:'#1e293b', border:'1px solid #3b82f6', color:'#93c5fd', borderRadius:8, padding:'6px 14px', fontSize:12, cursor:'pointer' }}>← Volver</button>
        <p style={{ color:'#475569', marginTop:20 }}>Sin datos disponibles.</p>
      </div>
    );
  }

  return (
    <div style={{ fontFamily:'system-ui,sans-serif', background:'#0d1b2a', minHeight:'100vh', color:'#e2e8f0', padding:'20px 24px' }}>
      <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:20 }}>
        <button onClick={onBack} style={{ background:'#1e293b', border:'1px solid #3b82f6', color:'#93c5fd', borderRadius:8, padding:'6px 14px', fontSize:12, cursor:'pointer' }}>← Volver</button>
        <div style={{ fontSize:18, fontWeight:700, color:'#fff' }}>Comparaciones</div>
      </div>
      <div style={{ display:'flex', flexDirection:'column', gap:16 }}>

        {pbChart.length > 0 && (
          <Card>
            <SLabel>Volumen por periodo y bloque</SLabel>
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={pbChart}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                <XAxis dataKey="name" tick={{ fill:'#64748b', fontSize:11 }} />
                <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
                <TT />
                <Bar dataKey="total" radius={[4,4,0,0]}>
                  {pbChart.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>
        )}

        {cursos.length > 0 && periodos.length > 0 && (
          <Card>
            <SLabel>Consultas por curso entre periodos</SLabel>
            <div style={{ overflowX:'auto' }}>
              <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
                <thead>
                  <tr>
                    <th style={TH}>Curso</th>
                    {periodos.map(p => <th key={p} style={{ ...TH, textAlign:'center' }}>{p}</th>)}
                    <th style={{ ...TH, textAlign:'center' }}>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {cursos.map((c,i) => {
                    const rowTotal = periodos.reduce((s,p)=>s+(cxp[c]?.[p]||0),0);
                    return (
                      <tr key={c} style={{ background: i%2===0?'#141c2b':'#1e293b' }}>
                        <td style={TD}>{c}</td>
                        {periodos.map(p => (
                          <td key={p} style={{ ...TD, textAlign:'center', color:cxp[c]?.[p]?'#60a5fa':'#334155', fontWeight:cxp[c]?.[p]?600:400 }}>
                            {cxp[c]?.[p]||'-'}
                          </td>
                        ))}
                        <td style={{ ...TD, textAlign:'center', color:'#10b981', fontWeight:700 }}>{rowTotal||'-'}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        {temaTabla.rows.length > 0 && (
          <Card>
            <SLabel>Comparación de temáticas entre periodos</SLabel>
            <div style={{ overflowX:'auto', marginBottom:10 }}>
              <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
                <thead>
                  <tr>
                    <th style={TH}>Temática</th>
                    {temaTabla.pers.map(p => <th key={p} style={{ ...TH, textAlign:'center' }}>{p}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {temaTabla.rows.map((t,i) => (
                    <tr key={t.tg} style={{ background: i%2===0?'#141c2b':'#1e293b' }}>
                      <td style={TD}>{t.tg}</td>
                      {temaTabla.pers.map((p,pi) => {
                        const val = t[p]||0;
                        const maxV = Math.max(...temaTabla.rows.map(x=>x[p]||0),1);
                        return (
                          <td key={p} style={{ ...TD, textAlign:'center' }}>
                            <div style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:3 }}>
                              <span style={{ color:val>0?PAL[pi%PAL.length]:'#334155', fontWeight:val>0?600:400 }}>{val||'-'}</span>
                              {val>0 && (
                                <div style={{ width:44, height:4, background:'#0f172a', borderRadius:2, overflow:'hidden' }}>
                                  <div style={{ height:'100%', width:`${(val/maxV)*100}%`, background:PAL[pi%PAL.length], borderRadius:2 }} />
                                </div>
                              )}
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ fontSize:11, color:'#475569' }}>
              La barra muestra el peso relativo de cada temática dentro del periodo. Detecta cambios en el tipo de uso del tutor IA entre periodos.
            </div>
          </Card>
        )}

        {blqChart.length > 0 && (
          <Card>
            <SLabel>Comparación entre bloques</SLabel>
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={blqChart}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                <XAxis dataKey="name" tick={{ fill:'#64748b', fontSize:13 }} />
                <YAxis tick={{ fill:'#64748b', fontSize:11 }} />
                <TT />
                <Bar dataKey="total" radius={[4,4,0,0]}>
                  {blqChart.map((_,i) => <Cell key={i} fill={PAL[i%PAL.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>
        )}

      </div>
    </div>
  );
}

// Helpers Google Sheets
function extractSheetId(url) {
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : null;
}
// Builds the most reliable CSV URL from any Google Sheets link
function buildCsvUrl(url) {
  // Already a complete pub CSV URL — use as-is
  if (url.includes('/pub?') && url.includes('output=csv')) return url;
  const id = extractSheetId(url);
  if (!id) return null;
  // Pub URL without output param
  if (url.includes('/pub')) {
    const gidMatch = url.match(/gid=(\d+)/);
    const gid = gidMatch ? gidMatch[1] : '0';
    return `https://docs.google.com/spreadsheets/d/${id}/pub?output=csv&gid=${gid}&single=true`;
  }
  // Standard share/edit URL -> export CSV
  const gidMatch = url.match(/gid=(\d+)/);
  const gid = gidMatch ? gidMatch[1] : '0';
  return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
}
function fmtCountdown(s) {
  const m = Math.floor(s/60), sec = s%60;
  return `${m}:${String(sec).padStart(2,'0')}`;
}

// App principal
export default function App() {
  const [raw, setRaw] = useState(null);
  const [fileInfo, setFileInfo] = useState(null);
  const [loading, setLoading] = useState(false);
  const [filters, setFilters] = useState({ periodo:'Todos', bloque:'Todos', curso:'Todos', semana:'Todas', busqueda:'' });
  const [activeTab, setActiveTab] = useState('resumen');
  const [view, setView] = useState('dashboard');
  const fileRef = useRef();

  // Google Sheets
  const [dataSource, setDataSource] = useState('');
  const [sheetsUrl, setSheetsUrl] = useState('https://docs.google.com/spreadsheets/d/e/2PACX-1vSv6HB5-J_8Gqh9RctyZRLPlq8Wi8UVUl59HxdW2WiOHGV15WNr6ddLmcf0VOoLxWhhkd4Ncp6il1_g/pub?gid=0&single=true&output=csv');
  const [sheetsError, setSheetsError] = useState('');
  const [lastUpdated, setLastUpdated] = useState(null);
  const [countdown, setCountdown] = useState(300);
  const REFRESH_SEC = 300;

  const fetchFromSheets = useCallback(async (url) => {
    const csvUrl = url.includes('output=csv') ? url : buildCsvUrl(url);
    if (!csvUrl) { setSheetsError('URL invalida. Usa el enlace de "Publicar en la web" con formato CSV.'); return; }
    setSheetsError(''); setLoading(true);

    // Proxies CORS en orden de preferencia
    const PROXIES = [
      (u) => u,                                                          // directo
      (u) => `https://corsproxy.io/?${encodeURIComponent(u)}`,          // corsproxy.io
      (u) => `https://api.allorigins.win/raw?url=${encodeURIComponent(u)}`, // allorigins
      (u) => `https://thingproxy.freeboard.io/fetch/${u}`,              // thingproxy
    ];

    let lastErr = '';
    for (const proxy of PROXIES) {
      try {
        const proxyUrl = proxy(csvUrl);
        const res = await fetch(proxyUrl, { signal: AbortSignal.timeout(8000) });
        if (!res.ok) {
          lastErr = `HTTP ${res.status}`;
          continue;
        }
        const text = await res.text();
        if (!text || text.trim().startsWith('<') || text.includes('Host not in allowlist')) {
          lastErr = 'Respuesta invalida del servidor';
          continue;
        }
        const result = Papa.parse(text, { header:true, skipEmptyLines:true, transformHeader: h => h.trim() });
        if (!result.data || result.data.length === 0) {
          lastErr = 'Sin datos o encabezados incorrectos';
          continue;
        }
        const processed = processRows(result.data);
        setRaw(processed);
        const now = new Date();
        setFileInfo({ name:'Google Sheets', total:result.data.length, loaded:now });
        setLastUpdated(now); setCountdown(REFRESH_SEC);
        setLoading(false);
        return; // exito
      } catch(e) {
        lastErr = e.message || 'Error de red';
      }
    }

    // Si todos los proxies fallaron
    setSheetsError(`No se pudo conectar con Google Sheets. Ultimo error: ${lastErr}. Verifica que la hoja este publicada como CSV (Archivo → Compartir → Publicar en la web).`);
    setLoading(false);
  }, []);

  // Auto-connect on mount with the preconfigured sheet
  useEffect(() => {
    setDataSource('sheets');
    fetchFromSheets('https://docs.google.com/spreadsheets/d/e/2PACX-1vSv6HB5-J_8Gqh9RctyZRLPlq8Wi8UVUl59HxdW2WiOHGV15WNr6ddLmcf0VOoLxWhhkd4Ncp6il1_g/pub?gid=0&single=true&output=csv');
  }, []);

  useEffect(() => {
    if (dataSource !== 'sheets' || !sheetsUrl) return;
    const timer = setInterval(() => {
      setCountdown(prev => {
        if (prev <= 1) { fetchFromSheets(sheetsUrl); return REFRESH_SEC; }
        return prev - 1;
      });
    }, 1000);
    return () => clearInterval(timer);
  }, [dataSource, sheetsUrl, fetchFromSheets]);

  const handleFile = useCallback(e => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type:'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        setRaw(processRows(rows));
        setFileInfo({ name:file.name, total:rows.length, loaded:new Date() });
        setDataSource('file');
      } catch(err) { alert('Error al leer el archivo: ' + err.message); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const opts = useMemo(() => {
    if (!raw) return { periodos:[], bloques:[], cursos:[], semanas:[] };
    const semanas = [...new Set(raw.filter(r=>r.semana).map(r=>r.semana))].sort((a,b)=>a-b);
    return {
      periodos: ['Todos',...[...new Set(raw.map(r=>r.periodo))].sort().reverse()],
      bloques:  ['Todos',...[...new Set(raw.map(r=>r.bloque))].sort()],
      cursos:   ['Todos',...[...new Set(raw.map(r=>r.nombreCurso))].sort()],
      semanas:  ['Todas',...semanas.map(s=>`Sem ${s}`)],
    };
  }, [raw]);

  const fd = useMemo(() => {
    if (!raw) return [];
    return raw.filter(r => {
      if (filters.periodo !== 'Todos' && r.periodo !== filters.periodo) return false;
      if (filters.bloque  !== 'Todos' && r.bloque  !== filters.bloque)  return false;
      if (filters.curso   !== 'Todos' && r.nombreCurso !== filters.curso) return false;
      if (filters.semana  !== 'Todas') {
        const w = parseInt(filters.semana.replace('Sem ',''));
        if (r.semana !== w) return false;
      }
      if (filters.busqueda) {
        const b = filters.busqueda.toLowerCase();
        if (!r.nrc.toLowerCase().includes(b) && !r.codigoCurso.toLowerCase().includes(b)) return false;
      }
      return true;
    });
  }, [raw, filters]);

  const kpis = useMemo(() => ({
    total:    fd.length,
    usuarios: new Set(fd.map(r=>r.userId)).size,
    cursos:   new Set(fd.map(r=>r.nombreCurso)).size,
    sinR:     fd.filter(r=>r.sinR).length,
  }), [fd]);

  const TABS = [
    { id:'resumen',   label:'Resumen'      },
    { id:'cursos',    label:'Cursos'       },
    { id:'nrc',       label:'NRC'          },
    { id:'temp',      label:'Temporalidad' },
    { id:'temas',     label:'Temáticas'    },
    { id:'usuarios',  label:'Usuarios'     },
    { id:'calidad',   label:'Calidad'      },
  ];

  const SEL = { background:'#0f172a', border:'1px solid #1e293b', borderRadius:6, padding:'5px 8px', color:'#e2e8f0', fontSize:12 };
  const resetF = () => setFilters({ periodo:'Todos', bloque:'Todos', curso:'Todos', semana:'Todas', busqueda:'' });

  // Pantalla de carga
  if (!raw) return (
    <div style={{ fontFamily:'system-ui,sans-serif', background:'#0d1b2a', minHeight:'100vh', display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
      <div style={{ width:'100%', maxWidth:520 }}>
        <div style={{ textAlign:'center', marginBottom:28 }}>
          <div style={{ width:56, height:56, borderRadius:'50%', background:'#1e3a5f', margin:'0 auto 14px', display:'flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 12h8M12 8v8"/></svg>
          </div>
          <h2 style={{ margin:'0 0 6px', fontSize:20, fontWeight:700, color:'#fff' }}>Análisis de consultas del tutor IA</h2>
          <p style={{ color:'#64748b', fontSize:13, lineHeight:1.6, margin:0 }}>Conectando con Google Sheets...</p>
        </div>

        {/* Opcion 1: Excel */}
        <div style={{ background:'#1e293b', borderRadius:14, padding:'20px 24px', marginBottom:14, border:'1px solid #1e3a5f' }}>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12 }}>
            <div style={{ width:32, height:32, borderRadius:8, background:'rgba(59,130,246,0.15)', display:'flex', alignItems:'center', justifyContent:'center' }}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            </div>
            <div>
              <div style={{ fontSize:14, fontWeight:600, color:'#fff' }}>Archivo Excel</div>
              <div style={{ fontSize:11, color:'#475569' }}>Carga manual · Sin actualizacion automatica</div>
            </div>
          </div>
          <div onClick={() => fileRef.current?.click()}
            style={{ border:'2px dashed #334155', borderRadius:10, padding:'18px', cursor:'pointer', textAlign:'center', background:'#0f172a' }}
            onMouseEnter={e => e.currentTarget.style.borderColor='#3b82f6'}
            onMouseLeave={e => e.currentTarget.style.borderColor='#334155'}>
            {loading && dataSource === 'file' ? (
              <div style={{ color:'#3b82f6', fontSize:13 }}>Procesando...</div>
            ) : (
              <>
                <div style={{ color:'#93c5fd', fontWeight:600, fontSize:13, marginBottom:3 }}>Haz clic para seleccionar</div>
                <div style={{ color:'#475569', fontSize:12 }}>REPORTE_TUTOR_IA.xlsx</div>
              </>
            )}
          </div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:'none' }} />
        </div>

        {/* Opcion 2: Google Sheets */}
        <div style={{ background:'#1e293b', borderRadius:14, padding:'20px 24px', border:'1px solid #1e3a5f' }}>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12 }}>
            <div style={{ width:32, height:32, borderRadius:8, background:'rgba(16,185,129,0.15)', display:'flex', alignItems:'center', justifyContent:'center' }}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            </div>
            <div>
              <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                <div style={{ fontSize:14, fontWeight:600, color:'#fff' }}>Google Sheets</div>
                <span style={{ background:'rgba(16,185,129,0.2)', color:'#10b981', borderRadius:20, padding:'1px 8px', fontSize:10, fontWeight:600 }}>TIEMPO REAL</span>
              </div>
              <div style={{ fontSize:11, color:'#475569' }}>Actualizacion automatica cada 5 minutos</div>
            </div>
          </div>

          <div style={{ marginBottom:10 }}>
            <input
              value={sheetsUrl}
              onChange={e => { setSheetsUrl(e.target.value); setSheetsError(''); }}
              placeholder="https://docs.google.com/spreadsheets/d/.../pub?output=csv"
              style={{ width:'100%', background:'#0f172a', border:'1px solid #334155', borderRadius:8, padding:'9px 12px', color:'#e2e8f0', fontSize:12, boxSizing:'border-box' }}
            />
          </div>

          <button
            onClick={() => { if (sheetsUrl.trim()) { setDataSource('sheets'); fetchFromSheets(sheetsUrl.trim()); } }}
            disabled={!sheetsUrl.trim() || loading}
            style={{ width:'100%', background: sheetsUrl.trim() && !loading ? '#10b981' : '#1e3a5f', color:'#fff', border:'none', borderRadius:8, padding:'9px', fontSize:13, fontWeight:600, cursor: sheetsUrl.trim() && !loading ? 'pointer' : 'default' }}>
            {loading && dataSource === 'sheets' ? 'Conectando...' : 'Conectar Google Sheets'}
          </button>

          {sheetsError && (
            <div style={{ marginTop:10, padding:'8px 12px', background:'rgba(239,68,68,0.1)', border:'1px solid rgba(239,68,68,0.3)', borderRadius:8, fontSize:12, color:'#fca5a5', lineHeight:1.5 }}>
              {sheetsError}
            </div>
          )}

          <div style={{ marginTop:14, padding:'12px', background:'#0f172a', borderRadius:8 }}>
            <div style={{ fontSize:11, color:'#475569', marginBottom:8, fontWeight:600 }}>Pasos para conectar:</div>
            {['Archivo → Compartir → Publicar en la web', 'Selecciona la hoja → formato CSV → Publicar', 'Copia la URL generada y pegala arriba'].map((step, i) => (
              <div key={i} style={{ display:'flex', gap:8, marginBottom:5, alignItems:'flex-start' }}>
                <span style={{ width:18, height:18, borderRadius:'50%', background:'#1e3a5f', color:'#60a5fa', fontSize:10, fontWeight:700, display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0 }}>{i+1}</span>
                <span style={{ fontSize:11, color:'#64748b', lineHeight:1.5 }}>{step}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );

  // ── Vista Comparaciones ──
  if (view === 'comp') return <Comparaciones raw={raw} onBack={() => setView('dashboard')} />;

  // ── Dashboard principal ──
  return (
    <div style={{ fontFamily:'system-ui,sans-serif', background:'#0d1b2a', minHeight:'100vh', color:'#e2e8f0' }}>
      {/* Header */}
      <div style={{ background:'linear-gradient(135deg,#1a2f4e,#0d1b2a)', padding:'16px 24px', borderBottom:'1px solid #1e293b' }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', flexWrap:'wrap', gap:8 }}>
          <div>
            <div style={{ fontSize:18, fontWeight:700, color:'#fff', marginBottom:3 }}>Análisis de consultas del tutor IA</div>
            <div style={{ display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
              <span style={{ fontSize:11, color:'#475569' }}>
                {dataSource === 'sheets' ? '🟢 Google Sheets' : '📁 ' + fileInfo.name} · {fileInfo.total.toLocaleString()} registros
              </span>
              {dataSource === 'sheets' && lastUpdated && (
                <span style={{ fontSize:11, color:'#64748b' }}>
                  · Actualizado: {lastUpdated.toLocaleTimeString()} · Próxima actualización: <strong style={{ color:'#10b981' }}>{fmtCountdown(countdown)}</strong>
                </span>
              )}
              {dataSource === 'file' && (
                <span style={{ fontSize:11, color:'#64748b' }}>· {fileInfo.loaded.toLocaleTimeString()}</span>
              )}
            </div>
          </div>
          <div style={{ display:'flex', gap:8, alignItems:'center' }}>
            {dataSource === 'sheets' && (
              <button onClick={() => fetchFromSheets(sheetsUrl)}
                disabled={loading}
                style={{ background:'#0f2744', border:'1px solid #10b981', color:'#10b981', borderRadius:8, padding:'7px 12px', fontSize:12, fontWeight:600, cursor:'pointer', display:'flex', alignItems:'center', gap:5 }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>
                {loading ? 'Actualizando...' : 'Actualizar ahora'}
              </button>
            )}
            <button onClick={() => { setRaw(null); setFileInfo(null); setDataSource(''); setSheetsUrl(''); }}
              style={{ background:'#0f172a', border:'1px solid #334155', color:'#64748b', borderRadius:8, padding:'7px 12px', fontSize:12, cursor:'pointer' }}>
              Cambiar fuente
            </button>
            <button onClick={() => setView('comp')}
              style={{ background:'#1e3a5f', border:'1px solid #3b82f6', color:'#93c5fd', borderRadius:8, padding:'7px 14px', fontSize:12, fontWeight:600, cursor:'pointer' }}>
              📊 Ver comparaciones
            </button>
          </div>
        </div>
        {/* Badges de filtros activos */}
        <div style={{ display:'flex', gap:6, marginTop:8, flexWrap:'wrap' }}>
          {filters.periodo!=='Todos' && <span style={{ background:'#1e3a5f', color:'#93c5fd', borderRadius:20, padding:'2px 10px', fontSize:11 }}>Periodo: {filters.periodo} <span onClick={()=>setFilters(f=>({...f,periodo:'Todos'}))} style={{ cursor:'pointer', marginLeft:4 }}>×</span></span>}
          {filters.bloque!=='Todos'  && <span style={{ background:'#1e3a5f', color:'#93c5fd', borderRadius:20, padding:'2px 10px', fontSize:11 }}>Bloque: {filters.bloque} <span onClick={()=>setFilters(f=>({...f,bloque:'Todos'}))} style={{ cursor:'pointer', marginLeft:4 }}>×</span></span>}
          {filters.curso!=='Todos'   && <span style={{ background:'#1e3a5f', color:'#93c5fd', borderRadius:20, padding:'2px 10px', fontSize:11 }}>Curso: {filters.curso.substring(0,22)}… <span onClick={()=>setFilters(f=>({...f,curso:'Todos'}))} style={{ cursor:'pointer', marginLeft:4 }}>×</span></span>}
          {filters.semana!=='Todas'  && <span style={{ background:'#1e3a5f', color:'#93c5fd', borderRadius:20, padding:'2px 10px', fontSize:11 }}>{filters.semana} <span onClick={()=>setFilters(f=>({...f,semana:'Todas'}))} style={{ cursor:'pointer', marginLeft:4 }}>×</span></span>}
          {filters.busqueda          && <span style={{ background:'#1e3a5f', color:'#93c5fd', borderRadius:20, padding:'2px 10px', fontSize:11 }}>NRC: {filters.busqueda} <span onClick={()=>setFilters(f=>({...f,busqueda:''}))} style={{ cursor:'pointer', marginLeft:4 }}>×</span></span>}
        </div>
      </div>

      {/* Filtros */}
      <div style={{ background:'#0f172a', padding:'10px 24px', borderBottom:'1px solid #1e293b', display:'flex', gap:10, flexWrap:'wrap', alignItems:'center' }}>
        <label style={{ fontSize:12, color:'#64748b' }}>Periodo
          <select style={{ ...SEL, marginLeft:5 }} value={filters.periodo} onChange={e=>setFilters(f=>({...f,periodo:e.target.value}))}>
            {opts.periodos.map(o=><option key={o}>{o}</option>)}
          </select>
        </label>
        <label style={{ fontSize:12, color:'#64748b' }}>Bloque
          <select style={{ ...SEL, marginLeft:5 }} value={filters.bloque} onChange={e=>setFilters(f=>({...f,bloque:e.target.value}))}>
            {opts.bloques.map(o=><option key={o}>{o}</option>)}
          </select>
        </label>
        <label style={{ fontSize:12, color:'#64748b' }}>Semana
          <select style={{ ...SEL, marginLeft:5 }} value={filters.semana} onChange={e=>setFilters(f=>({...f,semana:e.target.value}))}>
            {opts.semanas.map(o=><option key={o}>{o}</option>)}
          </select>
        </label>
        <label style={{ fontSize:12, color:'#64748b' }}>Curso
          <select style={{ ...SEL, marginLeft:5, maxWidth:200 }} value={filters.curso} onChange={e=>setFilters(f=>({...f,curso:e.target.value}))}>
            {opts.cursos.map(o=><option key={o}>{o}</option>)}
          </select>
        </label>
        <input placeholder="NRC o código…" value={filters.busqueda}
          onChange={e=>setFilters(f=>({...f,busqueda:e.target.value}))}
          style={{ ...SEL, padding:'5px 9px', width:145 }} />
        <button onClick={resetF}
          style={{ marginLeft:'auto', background:'transparent', border:'1px solid #1e293b', color:'#64748b', borderRadius:6, padding:'5px 10px', fontSize:11, cursor:'pointer' }}>
          Limpiar
        </button>
        <span style={{ fontSize:11, color:'#475569' }}>{fd.length.toLocaleString()} registros</span>
      </div>

      {/* Alertas */}
      <Alerts fd={fd} filters={filters} />

      {/* KPIs */}
      <div style={{ padding:'14px 24px', display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12 }}>
        <KPI label="Consultas válidas" value={kpis.total.toLocaleString()} color="#3b82f6" />
        <KPI label="Usuarios únicos"   value={kpis.usuarios.toLocaleString()} color="#10b981" />
        <KPI label="Cursos visibles"   value={kpis.cursos.toLocaleString()} color="#f59e0b" />
        <KPI label="Sin respuesta"     value={kpis.sinR.toLocaleString()} color="#ef4444"
          sub={`${fd.length>0?((kpis.sinR/fd.length)*100).toFixed(1):0}%`} />
      </div>

      {/* Tabs */}
      <div style={{ padding:'0 24px', display:'flex', borderBottom:'1px solid #1e293b', overflowX:'auto' }}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => setActiveTab(t.id)}
            style={{ background:'transparent', border:'none',
              borderBottom: activeTab===t.id ? '2px solid #3b82f6' : '2px solid transparent',
              color: activeTab===t.id ? '#93c5fd' : '#475569',
              padding:'11px 16px', fontSize:13, fontWeight:600, cursor:'pointer', whiteSpace:'nowrap' }}>
            {t.label}
          </button>
        ))}
      </div>

      {/* Contenido */}
      <div style={{ padding:'20px 24px' }}>
        {activeTab==='resumen'  && <TabResumen  fd={fd} />}
        {activeTab==='cursos'   && <TabCursos   fd={fd} />}
        {activeTab==='nrc'      && <TabNRC      fd={fd} />}
        {activeTab==='temp'     && <TabTemp     fd={fd} />}
        {activeTab==='temas'    && <TabTemas    fd={fd} />}
        {activeTab==='usuarios' && <TabUsuarios fd={fd} />}
        {activeTab==='calidad'  && <TabCalidad  fd={fd} raw={raw} />}
      </div>
    </div>
  );
}

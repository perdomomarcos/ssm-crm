import { useState, useEffect, useRef, useCallback, useMemo, createContext, useContext } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const STORAGE_KEY = "ssm_crm_v3";
const SUPABASE_URL = "https://pdapxdlnbbhyyizbmtyi.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBkYXB4ZGxuYmJoeXlpemJtdHlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk3OTM2ODgsImV4cCI6MjA5NTM2OTY4OH0.N5YM1jl2cUbpANfeyT9y59RWSgxuPQC-ip_y6F1Jfio";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const sb = {
  async getAll() {
    const { data, error } = await supabase.from("clients").select("*").order("name");
    if (error) throw error;
    return (data||[]).map(fromDB);
  },
  async upsert(c) {
    const { error } = await supabase.from("clients").upsert(toDB(c), { onConflict: "id" });
    if (error) throw error;
  },
  async delete(id) {
    const { error } = await supabase.from("clients").delete().eq("id", id);
    if (error) throw error;
  },
};

// Converte qualquer formato de data para AAAA-MM-DD — retorna null se inválido
function toISODate(str) {
  try {
    if (!str || str.trim() === "") return null;
    const s = str.trim();
    // DD/MM/AAAA ou D/M/AAAA
    const dmY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (dmY) {
      const iso = `${dmY[3]}-${dmY[2].padStart(2,"0")}-${dmY[1].padStart(2,"0")}`;
      // Valida a data gerada
      const d = new Date(iso);
      if (isNaN(d.getTime())) return null;
      return iso;
    }
    // AAAA-MM-DD
    const ymd = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (ymd) {
      const d = new Date(s);
      if (isNaN(d.getTime())) return null;
      return s;
    }
    return null;
  } catch(e) { return null; }
}

function toDB(c) {
  return {
    id:c.id, name:c.name, tech:c.tech, n_contracts:c.nContracts,
    contract_active:c.contractActive, category:c.category,
    contract_end:toISODate(c.contractEnd),
    last_contact_log:c.lastContactLog,
    last_contact_ssm:toISODate(c.lastContactSSM),
    am:c.am, am_knows:c.amKnows, tickets_90d:Number(c.tickets90d)||0, ticket_type:c.ticketType,
    commercial_contact:c.commercialContact, renewal_history:c.renewalHistory,
    inactivity_reason:c.inactivityReason, observations:c.observations||"",
    folder_path:c.folderPath||"", br_frequency:c.brFrequency||"Trimestral",
    aws_history:c.awsHistory||[], azure_history:c.azureHistory||[],
    m365_history:c.m365History||[], br_reviews:c.brReviews||[],
  };
}

function fromDB(r) {
  return {
    id:r.id, name:r.name, tech:r.tech||"Desconhecido", nContracts:r.n_contracts||"1",
    contractActive:r.contract_active||"Desconhecido", category:r.category||"Desconhecido",
    contractEnd:r.contract_end||"", lastContactLog:r.last_contact_log||"",
    lastContactSSM:r.last_contact_ssm||"", am:r.am||"", amKnows:r.am_knows||"Desconhecido",
    tickets90d:r.tickets_90d||0, ticketType:r.ticket_type||"Nenhum",
    commercialContact:r.commercial_contact||"Desconhecido", renewalHistory:r.renewal_history||"Desconhecido",
    inactivityReason:r.inactivity_reason||"Desconhecido", observations:r.observations||"",
    folderPath:r.folder_path||"", brFrequency:r.br_frequency||"Trimestral",
    awsHistory:r.aws_history||[], azureHistory:r.azure_history||[],
    m365History:r.m365_history||[], brReviews:r.br_reviews||[],
  };
}

function getClosedMonths() {
  const now = new Date();
  const months = [];
  for (let i = 6; i >= 1; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(d.toLocaleDateString("pt-BR", { month: "short", year: "2-digit" })
      .replace(".", "").replace(/^\w/, c => c.toUpperCase()));
  }
  return months;
}
const MONTHS = getClosedMonths();

const TECH_OPTIONS = ["AWS","Azure","M365-CSP","M365-EA","AWS+Azure","AWS+M365-CSP","Azure+M365-CSP","Misto","Outro"];
const CATEGORY_OPTIONS = ["Premium","Advanced","Basic","Desconhecido"];
const SIM_NAO = ["Sim","Não","Desconhecido"];
const TICKET_TYPES = ["Nenhum","Dúvida","Incidente","Crítico","Misto"];
const RENEWAL_OPTIONS = ["Tranquila","Com desconto","Quase cancelou","Primeira renovação","Desconhecido"];
const INACTIVITY_OPTIONS = ["Troca de gestor","Corte de budget","Insatisfação declarada","Projeto pausado","Sem uso claro","Esquecimento mútuo","Desconhecido","Outro"];

// ─── Theme ────────────────────────────────────────────────────────────────────
const ThemeCtx = createContext({});
const useT = () => useContext(ThemeCtx);

const LIGHT = {
  bg:"#f1f5f9", card:"#ffffff", cardAlt:"#f8fafc",
  border:"#e5e7eb", borderLight:"#f3f4f6",
  text:"#111827", textMuted:"#6b7280", textSub:"#9ca3af",
  inputBg:"#ffffff", inputBorder:"#d1d5db", inputText:"#111827",
  sidebar:"#0f172a", sidebarBorder:"#1e293b", sidebarText:"#64748b", sidebarActive:"#93c5fd",
  tableHead:"#f8fafc", tableHover:"#f8fafc",
  btnSecBg:"#f3f4f6", btnSecText:"#374151", btnSecBorder:"#e5e7eb",
  chipBg:"rgba(255,255,255,0.7)",
};
const DARK = {
  bg:"#0f172a", card:"#1e293b", cardAlt:"#172033",
  border:"#334155", borderLight:"#334155",
  text:"#f1f5f9", textMuted:"#94a3b8", textSub:"#64748b",
  inputBg:"#0f172a", inputBorder:"#475569", inputText:"#f1f5f9",
  sidebar:"#020617", sidebarBorder:"#1e293b", sidebarText:"#64748b", sidebarActive:"#93c5fd",
  tableHead:"#1e293b", tableHover:"#263248",
  btnSecBg:"#334155", btnSecText:"#e2e8f0", btnSecBorder:"#475569",
  chipBg:"rgba(30,41,59,0.85)",
};

const TECH_COLORS = {
  "AWS":            { bg:"#fff7ed", border:"#f97316", text:"#c2410c" },
  "Azure":          { bg:"#eff6ff", border:"#3b82f6", text:"#1d4ed8" },
  "M365-CSP":       { bg:"#fff1f2", border:"#f43f5e", text:"#be123c" },
  "M365-EA":        { bg:"#fdf4ff", border:"#a855f7", text:"#7e22ce" },
  "AWS+Azure":      { bg:"#f0fdf4", border:"#16a34a", text:"#15803d" },
  "AWS+M365-CSP":   { bg:"#fff7ed", border:"#ea580c", text:"#9a3412" },
  "Azure+M365-CSP": { bg:"#f0f9ff", border:"#0284c7", text:"#0369a1" },
  "Misto":          { bg:"#f8fafc", border:"#64748b", text:"#334155" },
  "Outro":          { bg:"#f9fafb", border:"#9ca3af", text:"#374151" },
};

// ─── Score engines ────────────────────────────────────────────────────────────
function calcTrend(arr) {
  if (!arr || arr.length < 2) return null;
  const valid = arr.filter(v => v != null && v !== "" && !isNaN(v) && Number(v) > 0).map(Number);
  if (valid.length < 2) return "sem_dados";
  const last = valid[valid.length - 1];
  const peak = Math.max(...valid);
  const avg1 = valid.slice(0, Math.ceil(valid.length / 2)).reduce((a,b)=>a+b,0) / Math.ceil(valid.length/2);
  const avg2 = valid.slice(Math.floor(valid.length / 2)).reduce((a,b)=>a+b,0) / Math.ceil(valid.length/2);
  if (last === 0) return "zerado";
  if (peak > 0 && last / peak < 0.55) return "queda_abrupta";
  if (avg2 > avg1 * 1.1) return "crescente";
  if (avg2 < avg1 * 0.9) return "decrescente";
  return "estavel";
}

function parseDate(str) {
  if (!str) return null;
  // Try DD/MM/YYYY first
  const dmY = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmY) return new Date(Number(dmY[3]), Number(dmY[2])-1, Number(dmY[1]));
  // Try YYYY-MM-DD
  const ymd = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) return new Date(Number(ymd[1]), Number(ymd[2])-1, Number(ymd[3]));
  // Fallback
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

function daysSince(dateStr) {
  if (!dateStr) return null;
  const d = parseDate(dateStr);
  if (!d || isNaN(d)) return null;
  return Math.floor((Date.now() - d) / 86400000);
}

// ─── Business Review helpers ──────────────────────────────────────────────────
function brThisYear(brList) {
  const year = new Date().getFullYear();
  return (brList||[]).filter(br => {
    const d = parseDate(br.date);
    return d && d.getFullYear() === year;
  });
}

function calcExpectedBRs(freq) {
  if (freq === "Mensal")    return 12;
  if (freq === "Bimestral") return 6;
  if (freq === "Trimestral")return 4;
  if (freq === "Semestral") return 2;
  if (freq === "Anual")     return 1;
  return null;
}

function calcChurnRisk(c) {
  // Score 0-100: quanto maior, maior risco de churn
  let score = 0;

  const cloudAll = [...(c.awsHistory||[]),...(c.azureHistory||[])].map(Number).filter(v=>!isNaN(v)&&v>=0);
  const m365 = (c.m365History||[]).map(Number).filter(v=>!isNaN(v)&&v>=0);
  const trend = calcTrend(cloudAll.length ? [...(c.awsHistory||[]),...(c.azureHistory||[])] : c.m365History);

  // Consumo caindo = risco alto
  if (trend === "zerado")        score += 35;
  else if (trend === "queda_abrupta") score += 30;
  else if (trend === "decrescente")   score += 20;
  else if (trend === "sem_dados")     score += 15; // sem dados = sinal ruim
  else if (trend === "estavel")       score += 5;
  else if (trend === "crescente")     score += 0;

  // Tempo sem contato SSM
  const dSSM = daysSince(c.lastContactSSM);
  if (dSSM === null)    score += 20;
  else if (dSSM > 730)  score += 25;
  else if (dSSM > 365)  score += 18;
  else if (dSSM > 180)  score += 10;
  else if (dSSM > 90)   score += 4;

  // Insatisfação declarada
  if (c.inactivityReason === "Insatisfação declarada") score += 20;

  // Contrato ativo?
  if (c.contractActive === "Não") score += 15;

  // Sem contato comercial
  if (c.commercialContact === "Não") score += 5;
  else if (c.commercialContact === "Desconhecido") score += 8;

  // Histórico de renovação ruim
  if (c.renewalHistory === "Quase cancelou") score += 10;
  else if (c.renewalHistory === "Com desconto") score += 5;

  // Redução de licenças M365
  if (m365.length >= 3) {
    const first = m365[0], last = m365[m365.length-1];
    if (last < first * 0.8) score += 15;
  }

  return Math.min(100, Math.round(score));
}

function calcRescuePotential(c) {
  // Score 0-100: quanto maior, mais vale resgatar
  let score = 0;

  const cloudAll = [...(c.awsHistory||[]),...(c.azureHistory||[])].map(Number).filter(v=>!isNaN(v)&&v>0);
  const m365 = (c.m365History||[]).map(Number).filter(v=>!isNaN(v)&&v>0);
  const trend = calcTrend(cloudAll.length ? [...(c.awsHistory||[]),...(c.azureHistory||[])] : c.m365History);

  // Porte estimado pelo pico de consumo
  const peak = cloudAll.length ? Math.max(...cloudAll) : 0;
  const maxL = m365.length ? Math.max(...m365) : 0;
  if (peak >= 10000)     score += 25;
  else if (peak >= 3000) score += 18;
  else if (peak >= 500)  score += 10;
  else if (maxL >= 100)  score += 20;
  else if (maxL >= 20)   score += 12;
  else if (maxL > 0)     score += 6;
  else                   score += 5; // sem dados mas contrato ativo = potencial desconhecido

  // Categoria
  if (c.category === "Premium")  score += 15;
  else if (c.category === "Advanced") score += 8;

  // Motivo reversível
  const reversivel = ["Troca de gestor","Projeto pausado","Corte de budget","Esquecimento mútuo"];
  if (reversivel.includes(c.inactivityReason)) score += 20;
  else if (c.inactivityReason === "Insatisfação declarada") score += 8;

  // Engajamento residual
  const tix = parseInt(c.tickets90d) || 0;
  if (tix > 0) score += 12;
  if (c.commercialContact === "Sim") score += 8;

  // AM conhece = facilita abordagem
  if (c.amKnows === "Sim") score += 8;

  // Histórico positivo
  if (c.renewalHistory === "Tranquila") score += 7;

  // Tendência não completamente morta
  if (trend === "crescente") score += 5;
  else if (trend === "estavel") score += 3;

  // Penalizar se contrato inativo
  if (c.contractActive === "Não") score -= 15;

  return Math.min(100, Math.max(0, Math.round(score)));
}

function churnLabel(s) {
  if (s >= 70) return { label:"Crítico",   color:"#dc2626", bg:"#fef2f2", border:"#fecaca", emoji:"🔴" };
  if (s >= 45) return { label:"Alto",      color:"#ea580c", bg:"#fff7ed", border:"#fed7aa", emoji:"🟠" };
  if (s >= 25) return { label:"Médio",     color:"#d97706", bg:"#fffbeb", border:"#fde68a", emoji:"🟡" };
  return              { label:"Baixo",     color:"#16a34a", bg:"#f0fdf4", border:"#bbf7d0", emoji:"🟢" };
}

function rescueLabel(s) {
  if (s >= 65) return { label:"Alto",      color:"#1d4ed8", bg:"#eff6ff", border:"#bfdbfe", emoji:"⭐" };
  if (s >= 40) return { label:"Médio",     color:"#7c3aed", bg:"#f5f3ff", border:"#ddd6fe", emoji:"🔵" };
  if (s >= 20) return { label:"Baixo",     color:"#6b7280", bg:"#f9fafb", border:"#e5e7eb", emoji:"⚪" };
  return              { label:"Mínimo",    color:"#9ca3af", bg:"#f9fafb", border:"#f3f4f6", emoji:"—"  };
}

function matrixQuadrant(churn, rescue) {
  const highChurn = churn >= 45, highRescue = rescue >= 40;
  if (highChurn && highRescue)  return { label:"🔴 Resgatar Agora",  color:"#dc2626", bg:"#fef2f2" };
  if (highChurn && !highRescue) return { label:"🟠 Monitorar",       color:"#ea580c", bg:"#fff7ed" };
  if (!highChurn && highRescue) return { label:"⭐ Expandir",         color:"#1d4ed8", bg:"#eff6ff" };
  return                               { label:"⚪ Arquivar",          color:"#9ca3af", bg:"#f9fafb" };
}

function trendLabel(t) {
  const map = { crescente:"📈 Crescente", estavel:"➡️ Estável", decrescente:"📉 Decrescente", queda_abrupta:"⚠️ Queda abrupta", zerado:"🚫 Zerado", sem_dados:"— Sem dados" };
  return map[t] || "—";
}

// ─── Storage ──────────────────────────────────────────────────────────────────
function loadClients() {
  try { const r = localStorage.getItem(STORAGE_KEY); if (r) return JSON.parse(r); } catch {}
  return [];
}
function saveClients(list) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); } catch {}
}

// ─── Excel Import ─────────────────────────────────────────────────────────────
function parseExcelClients(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type:"array" });
        const ws = wb.Sheets["Triagem Brasil"];
        if (!ws) { reject(new Error("Aba 'Triagem Brasil' não encontrada")); return; }
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        const clients = [];
        for (const row of rows.slice(4)) {
          if (!row[0] || isNaN(Number(row[0]))) continue;
          const name = String(row[1]||"").trim(); if (!name) continue;
          const pn = v => { const n=parseFloat(String(v).replace(",",".")); return isNaN(n)?"":n; };
          clients.push({
            id: String(row[2]||`BR-SCU-${Date.now()}`).trim(),
            name,
            tech: String(row[3]||"").trim()||"Desconhecido",
            nContracts: String(row[4]||"1").trim(),
            lastContactLog: String(row[5]||"").trim(),
            contractActive: String(row[6]||"Desconhecido").trim(),
            category: String(row[7]||"Desconhecido").trim(),
            contractEnd: String(row[8]||"").trim(),
            awsHistory:   [row[9],row[10],row[11],row[12],row[13],row[14]].map(pn),
            azureHistory: [row[15],row[16],row[17],row[18],row[19],row[20]].map(pn),
            m365History:  [row[21],row[22],row[23],row[24],row[25],row[26]].map(pn),
            tickets90d: pn(row[27])||0,
            ticketType: String(row[28]||"Nenhum").trim(),
            commercialContact: String(row[29]||"Desconhecido").trim(),
            lastContactSSM: String(row[30]||"").trim(),
            renewalHistory: String(row[31]||"Desconhecido").trim(),
            amKnows: String(row[32]||"Desconhecido").trim(),
            inactivityReason: String(row[33]||"Desconhecido").trim(),
            am: String(row[34]||"").trim(),
            observations: String(row[35]||"").trim(),
            folderPath: "",
            brReviews: [],
            brFrequency: "Trimestral",
          });
        }
        resolve(clients);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── UI primitives ────────────────────────────────────────────────────────────
function mkCss(T) { return {
  input:  { width:"100%", padding:"8px 10px", borderRadius:7, border:`1px solid ${T.inputBorder}`, fontSize:13, fontFamily:"inherit", boxSizing:"border-box", outline:"none", background:T.inputBg, color:T.inputText },
  select: { width:"100%", padding:"8px 10px", borderRadius:7, border:`1px solid ${T.inputBorder}`, fontSize:13, background:T.inputBg, boxSizing:"border-box", color:T.inputText },
  label:  { fontSize:11, fontWeight:600, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.05em", display:"block", marginBottom:4 },
  card:   { background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 },
}; }
// Legacy css for non-theme components (overridden in themed ones)
const css = mkCss(LIGHT);

function TechBadge({ tech }) {
  const s = TECH_COLORS[tech]||TECH_COLORS["Outro"];
  return <span style={{ background:s.bg, border:`1px solid ${s.border}`, color:s.text, borderRadius:6, padding:"2px 8px", fontSize:11, fontWeight:600, whiteSpace:"nowrap" }}>{tech}</span>;
}

function Sparkline({ data, color="#3b82f6" }) {
  const valid = (data||[]).map(Number).filter(v=>!isNaN(v)&&v>0);
  if (!valid.length) return <span style={{ color:"#d1d5db", fontSize:11 }}>—</span>;
  const max = Math.max(...valid,1), w=72, h=24;
  const pts = valid.map((v,i)=>`${(i/(valid.length-1||1))*w},${h-(v/max)*(h-4)-2}`).join(" ");
  const last = pts.split(" ").pop().split(",");
  return <svg width={w} height={h}><polyline points={pts} fill="none" stroke={color} strokeWidth={1.8} strokeLinejoin="round" strokeLinecap="round"/><circle cx={last[0]} cy={last[1]} r={2.5} fill={color}/></svg>;
}

function ScorePill({ score, labelFn, title }) {
  const { label, color, bg, border, emoji } = labelFn(score);
  return (
    <div title={title} style={{ display:"flex", flexDirection:"column", alignItems:"center", background:bg, border:`1px solid ${border}`, borderRadius:8, padding:"6px 10px", minWidth:80 }}>
      <span style={{ fontSize:10, color:"#9ca3af", fontWeight:600, textTransform:"uppercase", letterSpacing:"0.04em", marginBottom:2 }}>{title}</span>
      <span style={{ fontSize:18, fontWeight:900, color, fontFamily:"monospace", lineHeight:1 }}>{score}</span>
      <span style={{ fontSize:10, color, fontWeight:600 }}>{emoji} {label}</span>
    </div>
  );
}

// ─── Matrix 2x2 ───────────────────────────────────────────────────────────────
function Matrix({ clients }) {
  const T = useT();
  const data = clients.map(c => ({
    ...c,
    churn: calcChurnRisk(c),
    rescue: calcRescuePotential(c),
  }));

  const quadrants = [
    { key:"RR", label:"🔴 Resgatar Agora",  desc:"Alto risco + Alto potencial", color:"#dc2626", bg:"#fef2f2", border:"#fecaca", filter: d => d.churn>=45 && d.rescue>=40 },
    { key:"MN", label:"🟠 Monitorar",        desc:"Alto risco + Baixo potencial", color:"#ea580c", bg:"#fff7ed", border:"#fed7aa", filter: d => d.churn>=45 && d.rescue<40 },
    { key:"EX", label:"⭐ Expandir",          desc:"Baixo risco + Alto potencial", color:"#1d4ed8", bg:"#eff6ff", border:"#bfdbfe", filter: d => d.churn<45 && d.rescue>=40 },
    { key:"AR", label:"⚪ Arquivar",           desc:"Baixo risco + Baixo potencial", color:"#9ca3af", bg:"#f9fafb", border:"#e5e7eb", filter: d => d.churn<45 && d.rescue<40 },
  ];

  return (
    <div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:16 }}>
        {/* Axis labels */}
        <div style={{ gridColumn:"span 2", display:"flex", justifyContent:"center", fontSize:11, color:"#6b7280", fontWeight:600, textTransform:"uppercase", letterSpacing:"0.05em" }}>
          ← Risco de Churn →
        </div>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        {quadrants.map(q => {
          const items = data.filter(q.filter).sort((a,b) => (b.churn+b.rescue)-(a.churn+a.rescue));
          return (
            <div key={q.key} style={{ background:q.bg, border:`1px solid ${q.border}`, borderRadius:12, padding:14 }}>
              <div style={{ fontSize:13, fontWeight:700, color:q.color, marginBottom:4 }}>{q.label}</div>
              <div style={{ fontSize:11, color:"#9ca3af", marginBottom:10 }}>{q.desc} · {items.length} clientes</div>
              {items.slice(0,5).map(c => (
                <div key={c.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6, padding:"4px 8px", background:T.chipBg, borderRadius:6 }}>
                  <span style={{ fontSize:12, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:140 }}>{c.name.split(" ").slice(0,2).join(" ")}</span>
                  <div style={{ display:"flex", gap:4, flexShrink:0 }}>
                    <span style={{ fontSize:10, background:"#fee2e2", color:"#dc2626", borderRadius:4, padding:"1px 5px", fontWeight:700 }}>C:{c.churn}</span>
                    <span style={{ fontSize:10, background:"#dbeafe", color:"#1d4ed8", borderRadius:4, padding:"1px 5px", fontWeight:700 }}>R:{c.rescue}</span>
                  </div>
                </div>
              ))}
              {items.length > 5 && <div style={{ fontSize:11, color:q.color, marginTop:4 }}>+{items.length-5} mais</div>}
              {items.length === 0 && <div style={{ fontSize:12, color:T.textSub }}>Nenhum cliente neste quadrante.</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({ clients, onNavigate }) {
  const T = useT();
  const scored = useMemo(() => clients.map(c => ({
    ...c,
    churn: calcChurnRisk(c),
    rescue: calcRescuePotential(c),
  })), [clients]);
  const critical = scored.filter(c => c.churn>=70).length;
  const highRisk = scored.filter(c => c.churn>=45&&c.churn<70).length;
  const highRescue = scored.filter(c => c.rescue>=65).length;
  const techDist = useMemo(() => clients.reduce((a,c)=>{a[c.tech]=(a[c.tech]||0)+1;return a;},{}), [clients]);
  // Painel Churn: churn >= 45, sem contato SSM nos últimos 90 dias
  const churnList = scored
    .filter(c => c.churn >= 45 && (daysSince(c.lastContactSSM) === null || daysSince(c.lastContactSSM) > 90))
    .sort((a,b) => b.churn - a.churn)
    .slice(0, 8);
  // Painel Resgate: rescue >= 40, sem contato SSM nos últimos 90 dias
  const rescueList = scored
    .filter(c => c.rescue >= 40 && (daysSince(c.lastContactSSM) === null || daysSince(c.lastContactSSM) > 90))
    .sort((a,b) => b.rescue - a.rescue)
    .slice(0, 8);

  const StatCard = ({ label, value, color, sub }) => (
    <div style={{ background:T.card, borderRadius:12, padding:"18px 20px", border:`1px solid ${T.border}`, borderTop:`3px solid ${color}`, flex:1, minWidth:120 }}>
      <div style={{ fontSize:34, fontWeight:900, color, fontFamily:"monospace", lineHeight:1.1 }}>{value}</div>
      <div style={{ fontSize:13, fontWeight:600, color:T.text, marginTop:4 }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:T.textSub, marginTop:2 }}>{sub}</div>}
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
      <div style={{ display:"flex", gap:12, flexWrap:"wrap" }}>
        <StatCard label="Total Clientes"    value={clients.length} color="#1d4ed8" sub="carteira Brasil" />
        <StatCard label="Churn Crítico"     value={critical}       color="#dc2626" sub="risco ≥ 70" />
        <StatCard label="Churn Alto"        value={highRisk}       color="#ea580c" sub="risco 45–69" />
        <StatCard label="Alto Potencial"    value={highRescue}     color="#7c3aed" sub="resgate ≥ 65" />
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
        {/* Painel Risco de Churn */}
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:14, fontWeight:700, color:"#dc2626", marginBottom:4 }}>🔴 Risco de Churn</div>
          <div style={{ fontSize:11, color:T.textMuted, marginBottom:14 }}>Churn ≥ 45 · sem contato nos últimos 90 dias</div>
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {churnList.map((c,i) => {
              const cL = churnLabel(c.churn);
              const d = daysSince(c.lastContactSSM);
              return (
                <div key={c.id} onClick={()=>onNavigate("detail",c)} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:8, background:T.cardAlt, cursor:"pointer", border:`1px solid ${T.border}` }}>
                  <div style={{ width:36, height:36, borderRadius:8, background:cL.bg, border:`1px solid ${cL.border}`, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:900, fontSize:14, color:cL.color, fontFamily:"monospace", flexShrink:0 }}>{c.churn}</div>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:13, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{c.name}</div>
                    <div style={{ display:"flex", gap:6, marginTop:2, alignItems:"center" }}>
                      <TechBadge tech={c.tech}/>
                      <span style={{ fontSize:10, color:T.textSub }}>{c.inactivityReason||"—"}</span>
                    </div>
                  </div>
                  <div style={{ fontSize:10, color: d===null?"#dc2626": d>365?"#dc2626": d>180?"#ea580c":"#d97706", fontWeight:700, textAlign:"right", flexShrink:0 }}>
                    {d===null?"sem data":`${d}d`}
                  </div>
                </div>
              );
            })}
            {churnList.length===0 && <div style={{ fontSize:12, color:T.textSub, padding:"12px 0" }}>Nenhum cliente em risco com contato pendente.</div>}
          </div>
        </div>

        {/* Painel Potencial de Resgate */}
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:14, fontWeight:700, color:"#1d4ed8", marginBottom:4 }}>⭐ Potencial de Resgate</div>
          <div style={{ fontSize:11, color:T.textMuted, marginBottom:14 }}>Resgate ≥ 40 · sem contato nos últimos 90 dias</div>
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {rescueList.map((c,i) => {
              const rL = rescueLabel(c.rescue);
              const cloudH = [...(c.azureHistory||[]),...(c.awsHistory||[])].map(Number).filter(v=>!isNaN(v)&&v>0);
              return (
                <div key={c.id} onClick={()=>onNavigate("detail",c)} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:8, background:T.cardAlt, cursor:"pointer", border:`1px solid ${T.border}` }}>
                  <div style={{ width:36, height:36, borderRadius:8, background:rL.bg, border:`1px solid ${rL.border}`, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:900, fontSize:14, color:rL.color, fontFamily:"monospace", flexShrink:0 }}>{c.rescue}</div>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:13, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{c.name}</div>
                    <div style={{ display:"flex", gap:6, marginTop:2, alignItems:"center" }}>
                      <TechBadge tech={c.tech}/>
                      <span style={{ fontSize:10, color:T.textSub }}>{c.category}</span>
                    </div>
                  </div>
                  <Sparkline data={cloudH.length?cloudH:(c.m365History||[])} color={cloudH.length?"#3b82f6":"#f43f5e"}/>
                </div>
              );
            })}
            {rescueList.length===0 && <div style={{ fontSize:12, color:T.textSub, padding:"12px 0" }}>Nenhum cliente com alto potencial de resgate pendente.</div>}
          </div>
        </div>

        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
            <div style={{ fontSize:14, fontWeight:700, color:T.text, marginBottom:12 }}>☁️ Por Tecnologia</div>
            {Object.entries(techDist).sort((a,b)=>b[1]-a[1]).map(([tech,count])=>{
              const s = TECH_COLORS[tech]||TECH_COLORS["Outro"];
              const pct = Math.round((count/clients.length)*100);
              return (
                <div key={tech} style={{ marginBottom:9 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                    <span style={{ fontSize:12, color:s.text, fontWeight:600 }}>{tech}</span>
                    <span style={{ fontSize:12, color:T.textMuted }}>{count} ({pct}%)</span>
                  </div>
                  <div style={{ height:5, background:"#f3f4f6", borderRadius:3 }}>
                    <div style={{ height:"100%", width:`${pct}%`, background:s.border, borderRadius:3 }}/>
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
            <div style={{ fontSize:14, fontWeight:700, color:T.text, marginBottom:12 }}>⚡ Alertas Críticos</div>
            {scored.filter(c=>c.churn>=70).slice(0,4).map(c=>(
              <div key={c.id} onClick={()=>onNavigate("detail",c)} style={{ fontSize:12, color:T.text, marginBottom:8, cursor:"pointer", display:"flex", gap:6 }}>
                <span>🔴</span><span><strong>{c.name.split(" ")[0]}</strong> · Churn {c.churn} · {c.inactivityReason||"—"}</span>
              </div>
            ))}
            {scored.filter(c=>c.churn>=70).length===0 && <div style={{ fontSize:12, color:T.textSub }}>Nenhum cliente em churn crítico.</div>}
          </div>
        </div>
      </div>

      {/* Matrix */}
      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
        <div style={{ fontSize:14, fontWeight:700, color:T.text, marginBottom:4 }}>📊 Matriz Risco de Churn × Potencial de Resgate</div>
        <div style={{ fontSize:12, color:T.textMuted, marginBottom:16 }}>Cada cliente posicionado pelos dois scores simultaneamente</div>
        <Matrix clients={clients}/>
      </div>
    </div>
  );
}

// ─── Client List ──────────────────────────────────────────────────────────────
function ClientList({ clients, onSelect, onAdd, onImport, importing }) {
  const T = useT();
  const css = mkCss(T);
  const [search, setSearch] = useState("");
  const [filterTech, setFilterTech] = useState("Todos");
  const [filterQ, setFilterQ] = useState("Todos");
  const [sort, setSort] = useState("churn_desc");
  const fileRef = useRef();

  const techs = ["Todos",...new Set(clients.map(c=>c.tech))];
  const qFilters = ["Todos","Resgatar Agora","Monitorar","Expandir","Arquivar"];

  const processed = useMemo(() => {
    let list = clients.map(c=>({ ...c, churn:calcChurnRisk(c), rescue:calcRescuePotential(c) }));
    list = list.map(c=>({ ...c, quadrant:matrixQuadrant(c.churn,c.rescue).label }));
    if (search) list = list.filter(c=>c.name.toLowerCase().includes(search.toLowerCase())||(c.id||"").toLowerCase().includes(search.toLowerCase()));
    if (filterTech!=="Todos") list = list.filter(c=>c.tech===filterTech);
    if (filterQ!=="Todos") list = list.filter(c=>c.quadrant.includes(filterQ.split(" ").pop()));
    if (sort==="churn_desc")  list.sort((a,b)=>b.churn-a.churn);
    else if (sort==="rescue_desc") list.sort((a,b)=>b.rescue-a.rescue);
    else if (sort==="name")   list.sort((a,b)=>a.name.localeCompare(b.name));
    return list;
  }, [clients,search,filterTech,filterQ,sort]);

  return (
    <div>
      <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por nome ou SCU…" style={{ ...css.input, width:230 }}/>
        <select value={filterTech} onChange={e=>setFilterTech(e.target.value)} style={{ ...css.select, width:"auto", padding:"8px 10px" }}>
          {techs.map(t=><option key={t}>{t}</option>)}
        </select>
        <select value={filterQ} onChange={e=>setFilterQ(e.target.value)} style={{ ...css.select, width:"auto", padding:"8px 10px" }}>
          {qFilters.map(q=><option key={q}>{q}</option>)}
        </select>
        <select value={sort} onChange={e=>setSort(e.target.value)} style={{ ...css.select, width:"auto", padding:"8px 10px" }}>
          <option value="churn_desc">Churn ↓</option>
          <option value="rescue_desc">Potencial ↓</option>
          <option value="name">Nome A–Z</option>
        </select>
        <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
          <input ref={fileRef} type="file" accept=".xlsx" style={{ display:"none" }} onChange={e=>e.target.files[0]&&onImport(e.target.files[0])}/>
          <button onClick={()=>fileRef.current.click()} disabled={importing} style={{ padding:"8px 16px", borderRadius:8, background:"#f0fdf4", color:"#16a34a", border:"1px solid #bbf7d0", fontSize:13, fontWeight:600, cursor:"pointer" }}>
            {importing?"⏳ Importando…":"📥 Importar Excel"}
          </button>
          <button onClick={onAdd} style={{ padding:"8px 16px", borderRadius:8, background:"#1d4ed8", color:"#fff", border:"none", fontSize:13, fontWeight:700, cursor:"pointer" }}>
            + Novo Cliente
          </button>
        </div>
      </div>
      <div style={{ fontSize:12, color:T.textMuted, marginBottom:8 }}>{processed.length} clientes</div>
      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:0, overflow:"hidden" }}>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead>
            <tr style={{ background:T.tableHead }}>
              {["Cliente","SCU","Tecnologia","Risco Churn","Pot. Resgate","Quadrante","Consumo (6m)","Último Contato"].map(h=>(
                <th key={h} style={{ padding:"10px 12px", textAlign:"left", fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.05em", borderBottom:`1px solid ${T.border}`, whiteSpace:"nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {processed.map(c=>{
              const cL=churnLabel(c.churn), rL=rescueLabel(c.rescue), q=matrixQuadrant(c.churn,c.rescue);
              const cloudH=[...(c.azureHistory||[]),...(c.awsHistory||[])].map(Number).filter(v=>!isNaN(v)&&v>0);
              return (
                <tr key={c.id} onClick={()=>onSelect(c)} style={{ cursor:"pointer", borderBottom:`1px solid ${T.borderLight}` }}
                  onMouseEnter={e=>e.currentTarget.style.background=T.tableHover}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{ padding:"11px 12px" }}>
                    <div style={{ fontWeight:600, fontSize:13, color:T.text }}>{c.name}</div>
                    <div style={{ fontSize:11, color:T.textSub }}>{c.category}</div>
                  </td>
                  <td style={{ padding:"11px 12px" }}><span style={{ fontFamily:"monospace", fontSize:11, color:T.textMuted }}>{c.id}</span></td>
                  <td style={{ padding:"11px 12px" }}><TechBadge tech={c.tech}/></td>
                  <td style={{ padding:"11px 12px" }}>
                    <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                      <div style={{ width:32, height:32, borderRadius:"50%", border:`2px solid ${cL.color}`, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:12, color:cL.color }}>{c.churn}</div>
                      <span style={{ fontSize:11, color:cL.color, fontWeight:600 }}>{cL.emoji}</span>
                    </div>
                  </td>
                  <td style={{ padding:"11px 12px" }}>
                    <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                      <div style={{ width:32, height:32, borderRadius:"50%", border:`2px solid ${rL.color}`, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:12, color:rL.color }}>{c.rescue}</div>
                      <span style={{ fontSize:11, color:rL.color, fontWeight:600 }}>{rL.emoji}</span>
                    </div>
                  </td>
                  <td style={{ padding:"11px 12px" }}>
                    <span style={{ fontSize:11, fontWeight:700, color:q.color, background:q.bg, borderRadius:6, padding:"2px 8px", whiteSpace:"nowrap" }}>{q.label}</span>
                  </td>
                  <td style={{ padding:"11px 12px" }}>
                    <Sparkline data={cloudH.length?cloudH:c.m365History} color={cloudH.length?"#3b82f6":"#f43f5e"}/>
                  </td>
                  <td style={{ padding:"11px 12px" }}>
                    <span style={{ fontSize:12, color:T.text }}>{c.lastContactSSM||c.lastContactLog||"—"}</span>
                  </td>
                </tr>
              );
            })}
            {processed.length===0&&(
              <tr><td colSpan={8} style={{ padding:40, textAlign:"center", color:T.textSub, fontSize:14, background:T.card }}>
                {clients.length===0?"Nenhum cliente cadastrado. Importe o Excel ou adicione manualmente.":"Nenhum resultado."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Client Form ──────────────────────────────────────────────────────────────
function FormSection({ title, children }) {
  const T = useT();
  const css = mkCss(T);
  return (
    <div style={{ ...css.card, marginBottom:14 }}>
      <div style={{ fontSize:12, fontWeight:700, color:T.text, paddingBottom:12, marginBottom:14, borderBottom:`1px solid ${T.borderLight}` }}>{title}</div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>{children}</div>
    </div>
  );
}

// ─── Form field components (fora do ClientForm para evitar re-mount no re-render) ──
function FInp({ label, field, value, onChange, type="text", span=1, placeholder="", css, T }) {
  return (
    <div style={{ gridColumn:`span ${span}` }}>
      <label style={css.label}>{label}</label>
      <input type={type} value={value ?? ""} placeholder={placeholder}
        onChange={e => onChange(field, e.target.value)}
        style={css.input} />
    </div>
  );
}

function FSel({ label, field, value, onChange, options, span=1, css }) {
  return (
    <div style={{ gridColumn:`span ${span}` }}>
      <label style={css.label}>{label}</label>
      <select value={value ?? ""} onChange={e => onChange(field, e.target.value)} style={css.select}>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

function FHistRow({ label, field, values, onChange, color, css, T }) {
  return (
    <div>
      <label style={css.label}>{label}</label>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:6 }}>
        {MONTHS.map((m,i) => (
          <div key={m}>
            <div style={{ fontSize:10, color:T.textSub, textAlign:"center", marginBottom:3 }}>{m}</div>
            <input type="number" value={values?.[i] ?? ""}
              onChange={e => onChange(field, i, e.target.value)}
              style={{ ...css.input, textAlign:"center", padding:"6px 4px", borderColor:`${color}60` }}/>
          </div>
        ))}
      </div>
    </div>
  );
}

function ClientForm({ initial, onSave, onCancel }) {
  const T = useT();
  const css = mkCss(T);

  const [form, setForm] = useState(() => ({
    id: initial?.id || "",
    name: initial?.name || "",
    tech: initial?.tech || "Azure",
    nContracts: initial?.nContracts || "1",
    contractActive: initial?.contractActive || "Desconhecido",
    category: initial?.category || "Desconhecido",
    contractEnd: initial?.contractEnd || "",
    lastContactLog: initial?.lastContactLog || "",
    awsHistory: initial?.awsHistory || Array(6).fill(""),
    azureHistory: initial?.azureHistory || Array(6).fill(""),
    m365History: initial?.m365History || Array(6).fill(""),
    tickets90d: initial?.tickets90d || "",
    ticketType: initial?.ticketType || "Nenhum",
    commercialContact: initial?.commercialContact || "Desconhecido",
    lastContactSSM: initial?.lastContactSSM || "",
    renewalHistory: initial?.renewalHistory || "Desconhecido",
    amKnows: initial?.amKnows || "Desconhecido",
    inactivityReason: initial?.inactivityReason || "Desconhecido",
    am: initial?.am || "",
    observations: initial?.observations || "",
    folderPath: initial?.folderPath || "",
    brFrequency: initial?.brFrequency || "Trimestral",
    brReviews: initial?.brReviews || [],
  }));

  const set = useCallback((field, val) =>
    setForm(f => ({...f, [field]: val})), []);

  const setHist = useCallback((field, idx, val) =>
    setForm(f => {
      const arr = [...(f[field] || Array(6).fill(""))];
      arr[idx] = val === "" ? "" : isNaN(Number(val)) ? val : Number(val);
      return {...f, [field]: arr};
    }), []);

  const collect = () => ({
    ...form,
    id: form.id || `BR-SCU-${Date.now().toString().slice(-6)}`,
  });

  const p = { onChange: set, css, T };
  const ph = { onChange: setHist, css, T };

  return (
    <div>
      <FormSection title="📋 Identificação & Contrato">
        <FInp {...p} label="Nome do Cliente *" field="name" value={form.name} span={2}/>
        <FInp {...p} label="SCU / ID Interno" field="id" value={form.id}/>
        <FSel {...p} label="Tecnologia Principal" field="tech" value={form.tech} options={TECH_OPTIONS}/>
        <FSel {...p} label="Categoria" field="category" value={form.category} options={CATEGORY_OPTIONS}/>
        <FInp {...p} label="Nº de Contratos" field="nContracts" value={form.nContracts} type="number"/>
        <FSel {...p} label="Contrato Realmente Ativo?" field="contractActive" value={form.contractActive} options={SIM_NAO}/>
        <FInp {...p} label="Vencimento do Contrato (DD/MM/AAAA)" field="contractEnd" value={form.contractEnd} placeholder="Ex: 31/12/2026"/>
        <FInp {...p} label="Account Manager" field="am" value={form.am}/>
      </FormSection>

      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20, marginBottom:14 }}>
        <div style={{ fontSize:12, fontWeight:700, color:T.text, paddingBottom:12, marginBottom:14, borderBottom:`1px solid ${T.borderLight}` }}>
          📊 Histórico de Consumo — Últimos 6 meses fechados ({MONTHS[0]} → {MONTHS[5]})
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
          <FHistRow {...ph} label="AWS — Valor mensal (moeda do contrato)" field="awsHistory" values={form.awsHistory} color="#f97316"/>
          <FHistRow {...ph} label="Azure — Valor mensal (moeda do contrato)" field="azureHistory" values={form.azureHistory} color="#3b82f6"/>
          <FHistRow {...ph} label="M365-CSP — Licenças ativas por mês" field="m365History" values={form.m365History} color="#f43f5e"/>
        </div>
      </div>

      <FormSection title="🎫 Suporte & Comercial">
        <FInp {...p} label="Tickets (últimos 90d)" field="tickets90d" value={form.tickets90d} type="number"/>
        <FSel {...p} label="Tipo do Último Ticket" field="ticketType" value={form.ticketType} options={TICKET_TYPES}/>
        <FSel {...p} label="Contato Comercial (CRM)?" field="commercialContact" value={form.commercialContact} options={SIM_NAO}/>
      </FormSection>

      <FormSection title="🤝 Relacionamento SSM">
        <FInp {...p} label="Data Último Contato SSM (DD/MM/AAAA)" field="lastContactSSM" value={form.lastContactSSM} placeholder="Ex: 15/05/2026"/>
        <FSel {...p} label="Histórico de Renovação" field="renewalHistory" value={form.renewalHistory} options={RENEWAL_OPTIONS}/>
        <FSel {...p} label="AM Conhece o Cliente?" field="amKnows" value={form.amKnows} options={SIM_NAO}/>
        <FSel {...p} label="Motivo de Inatividade" field="inactivityReason" value={form.inactivityReason} options={INACTIVITY_OPTIONS} span={2}/>
        <FInp {...p} label="Último Contato (log tracker)" field="lastContactLog" value={form.lastContactLog}/>
        <FSel {...p} label="Freq. de Business Review" field="brFrequency" value={form.brFrequency} options={["Mensal","Bimestral","Trimestral","Semestral","Anual","Não definida"]}/>
      </FormSection>

      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20, marginBottom:14 }}>
        <div style={{ fontSize:12, fontWeight:700, color:T.text, paddingBottom:12, marginBottom:14, borderBottom:`1px solid ${T.borderLight}` }}>📁 Diretório do Cliente</div>
        <div>
          <label style={css.label}>Caminho da Pasta</label>
          <div style={{ display:"flex", gap:8 }}>
            <input value={form.folderPath ?? ""} placeholder="/Users/perdomo/Documents/Clientes/NomeDoCliente"
              onChange={e => set("folderPath", e.target.value)}
              style={{ ...css.input, fontFamily:"monospace", fontSize:12 }}/>
            <button onClick={()=>{ if(form.folderPath) navigator.clipboard.writeText(form.folderPath); }}
              style={{ padding:"8px 12px", borderRadius:7, background:T.btnSecBg, border:`1px solid ${T.btnSecBorder}`, fontSize:12, cursor:"pointer", whiteSpace:"nowrap", color:T.btnSecText }}>
              📋 Copiar
            </button>
          </div>
          <div style={{ fontSize:11, color:T.textSub, marginTop:4 }}>
            Cole o caminho completo da pasta onde ficam os arquivos deste cliente.
          </div>
        </div>
      </div>

      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20, marginBottom:20 }}>
        <label style={css.label}>🧩 Observações & Contexto</label>
        <textarea value={form.observations ?? ""} rows={3}
          onChange={e => set("observations", e.target.value)}
          placeholder="Notas relevantes, notícias recentes, contexto do cliente..."
          style={{ ...css.input, resize:"vertical" }}/>
      </div>

      <div style={{ display:"flex", gap:10 }}>
        <button onClick={()=>{ const d=collect(); if(d.name.trim()) onSave(d); }}
          style={{ padding:"10px 24px", borderRadius:8, background:"#1d4ed8", color:"#fff", border:"none", fontSize:14, fontWeight:700, cursor:"pointer" }}>
          Salvar Cliente
        </button>
        <button onClick={onCancel} style={{ padding:"10px 18px", borderRadius:8, background:T.btnSecBg, color:T.btnSecText, border:"none", fontSize:14, cursor:"pointer" }}>
          Cancelar
        </button>
      </div>
    </div>
  );
}


// ─── Business Review Section ──────────────────────────────────────────────────
const BR_TYPES = ["Online","Presencial","Híbrido","Email","Outro"];

function BRSection({ client, onSave }) {
  const T = useT();
  const css = mkCss(T);
  const [showForm, setShowForm] = useState(false);
  const [newBR, setNewBR] = useState({ date:"", type:"Online", notes:"" });
  const brList = client.brReviews || [];
  const thisYear = brThisYear(brList);
  const expected = calcExpectedBRs(client.brFrequency);
  const pct = expected ? Math.min(100, Math.round((thisYear.length / expected) * 100)) : null;
  const yearNow = new Date().getFullYear();

  const handleAdd = () => {
    if (!newBR.date) return;
    const updated = { ...client, brReviews: [...brList, { ...newBR, id: Date.now() }] };
    onSave(updated);
    setNewBR({ date:"", type:"Online", notes:"" });
    setShowForm(false);
  };

  const handleDelete = (id) => {
    const updated = { ...client, brReviews: brList.filter(br => br.id !== id) };
    onSave(updated);
  };

  return (
    <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20, marginBottom:14 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, paddingBottom:10, borderBottom:`1px solid ${T.borderLight}` }}>
        <div>
          <div style={{ fontSize:12, fontWeight:700, color:T.text }}>📅 Business Reviews</div>
          <div style={{ fontSize:11, color:T.textSub, marginTop:2 }}>
            Freq. definida: <strong>{client.brFrequency||"Não definida"}</strong>
            {expected && <span> · Meta {yearNow}: <strong>{expected} BRs</strong></span>}
          </div>
        </div>
        <button onClick={() => setShowForm(s => !s)}
          style={{ padding:"6px 14px", borderRadius:7, background:"#1d4ed8", color:"#fff", border:"none", fontSize:12, fontWeight:600, cursor:"pointer" }}>
          {showForm ? "Cancelar" : "+ Registrar BR"}
        </button>
      </div>

      {/* Progress bar */}
      {expected && (
        <div style={{ marginBottom:16 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
            <span style={{ fontSize:12, color:T.text, fontWeight:600 }}>
              {thisYear.length} de {expected} BRs realizados em {yearNow}
            </span>
            <span style={{ fontSize:12, fontWeight:700, color: pct>=100?"#16a34a":pct>=50?"#d97706":"#dc2626" }}>{pct}%</span>
          </div>
          <div style={{ height:8, background:"#f3f4f6", borderRadius:4, overflow:"hidden" }}>
            <div style={{ height:"100%", width:`${pct}%`, borderRadius:4, transition:"width 0.4s",
              background: pct>=100?"#16a34a":pct>=50?"#f59e0b":"#ef4444" }}/>
          </div>
        </div>
      )}

      {/* Add form */}
      {showForm && (
        <div style={{ background:T.cardAlt, borderRadius:10, padding:14, marginBottom:14, border:`1px solid ${T.border}` }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 2fr auto", gap:10, alignItems:"end" }}>
            <div>
              <label style={css.label}>Data do BR</label>
              <input type="text" placeholder="DD/MM/AAAA" value={newBR.date}
                onChange={e => setNewBR(b => ({...b, date:e.target.value}))}
                style={css.input}/>
            </div>
            <div>
              <label style={css.label}>Formato</label>
              <select value={newBR.type} onChange={e => setNewBR(b => ({...b, type:e.target.value}))} style={css.select}>
                {BR_TYPES.map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
            <div>
              <label style={css.label}>Observações</label>
              <input type="text" placeholder="Temas abordados, próximos passos…" value={newBR.notes}
                onChange={e => setNewBR(b => ({...b, notes:e.target.value}))}
                style={css.input}/>
            </div>
            <button onClick={handleAdd}
              style={{ padding:"8px 16px", borderRadius:7, background:"#16a34a", color:"#fff", border:"none", fontSize:13, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap" }}>
              Salvar
            </button>
          </div>
        </div>
      )}

      {/* BR list */}
      {brList.length === 0 ? (
        <div style={{ fontSize:13, color:T.textSub, textAlign:"center", padding:"20px 0" }}>
          Nenhum Business Review registrado ainda.
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {[...brList].sort((a,b) => {
            const da = parseDate(a.date), db = parseDate(b.date);
            return (db||0) - (da||0);
          }).map(br => {
            const d = parseDate(br.date);
            const isThisYear = d && d.getFullYear() === yearNow;
            return (
              <div key={br.id} style={{ display:"flex", alignItems:"center", gap:12, padding:"10px 12px", borderRadius:8, background: isThisYear?"#f0fdf4":T.cardAlt, border:`1px solid ${isThisYear?"#bbf7d0":T.borderLight}` }}>
                <div style={{ width:36, height:36, borderRadius:8, background: isThisYear?"#16a34a":"#6b7280", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                  <span style={{ fontSize:14 }}>📋</span>
                </div>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                    <span style={{ fontSize:13, fontWeight:700, color:"#111827" }}>{br.date}</span>
                    <span style={{ fontSize:11, background: isThisYear?"#dcfce7":"#f3f4f6", color: isThisYear?"#15803d":"#6b7280", borderRadius:4, padding:"1px 6px", fontWeight:600 }}>{br.type}</span>
                    {isThisYear && <span style={{ fontSize:10, color:"#16a34a", fontWeight:600 }}>✓ {yearNow}</span>}
                  </div>
                  {br.notes && <div style={{ fontSize:12, color:T.textMuted, marginTop:2, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{br.notes}</div>}
                </div>
                <button onClick={() => handleDelete(br.id)}
                  style={{ background:"none", border:"none", color:"#d1d5db", cursor:"pointer", fontSize:16, padding:"0 4px", flexShrink:0 }}
                  onMouseEnter={e=>e.target.style.color="#ef4444"}
                  onMouseLeave={e=>e.target.style.color="#d1d5db"}>
                  ×
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Client Detail ────────────────────────────────────────────────────────────
function ClientDetail({ client, onSave, onBack, onDelete }) {
  const T = useT();
  const css = mkCss(T);
  const [editing, setEditing] = useState(false);
  const churn = calcChurnRisk(client), rescue = calcRescuePotential(client);
  const q = matrixQuadrant(churn, rescue);
  const cL = churnLabel(churn), rL = rescueLabel(rescue);
  const cloudH = [...(client.azureHistory||[]),...(client.awsHistory||[])].map(Number).filter(v=>!isNaN(v)&&v>0);
  const trend = calcTrend(cloudH.length?[...(client.awsHistory||[]),...(client.azureHistory||[])]:client.m365History);
  const dSSM = daysSince(client.lastContactSSM);

  if (editing) return (
    <div>
      <button onClick={()=>setEditing(false)} style={{ background:"none", border:"none", color:T.textMuted, cursor:"pointer", fontSize:13, marginBottom:16, padding:0 }}>← Cancelar edição</button>
      <ClientForm initial={client} onSave={c=>{onSave(c);setEditing(false);}} onCancel={()=>setEditing(false)}/>
    </div>
  );

  const Row = ({ label, value }) => (
    <div>
      <div style={{ fontSize:11, fontWeight:600, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.05em", display:"block", marginBottom:4 }}>{label}</div>
      <div style={{ fontSize:13, color:T.text, padding:"6px 0" }}>{value||"—"}</div>
    </div>
  );

  const HistViz = ({ label, data, color }) => {
    const vals = (data||[]).map(v=>v===""||v==null?null:Number(v));
    if (!vals.some(v=>v!=null&&!isNaN(v)&&v>0)) return null;
    return (
      <div>
        <label style={css.label}>{label}</label>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:6 }}>
          {MONTHS.map((m,i)=>(
            <div key={m} style={{ textAlign:"center" }}>
              <div style={{ fontSize:10, color:T.textSub, marginBottom:3 }}>{m}</div>
              <div style={{ padding:"6px 4px", borderRadius:6, background:vals[i]?`${color}15`:T.cardAlt, fontSize:12, fontWeight:vals[i]?700:400, color:vals[i]?color:T.textSub }}>
                {vals[i]!=null&&!isNaN(vals[i])&&vals[i]>0?vals[i].toLocaleString("pt-BR"):"—"}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
        <div>
          <button onClick={onBack} style={{ background:"none", border:"none", color:T.textMuted, cursor:"pointer", fontSize:13, marginBottom:8, padding:0 }}>← Voltar</button>
          <h2 style={{ margin:0, fontSize:20, fontWeight:800, color:T.text }}>{client.name}</h2>
          <div style={{ display:"flex", gap:8, marginTop:6, alignItems:"center" }}>
            <TechBadge tech={client.tech}/>
            <span style={{ fontFamily:"monospace", fontSize:11, color:T.textSub }}>{client.id}</span>
            <span style={{ fontSize:11, color:T.textSub }}>· {client.category}</span>
          </div>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={()=>setEditing(true)} style={{ padding:"8px 18px", borderRadius:8, background:T.btnSecBg, color:T.btnSecText, border:"none", fontSize:13, fontWeight:600, cursor:"pointer" }}>✏️ Editar</button>
          <button onClick={()=>onDelete(client.id)} style={{ padding:"8px 14px", borderRadius:8, background:"#fef2f2", color:"#ef4444", border:"1px solid #fecaca", fontSize:13, cursor:"pointer" }}>🗑️</button>
        </div>
      </div>

      {/* Score cards */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:18 }}>
        <div style={{ background:q.bg, border:`1px solid ${q.color}30`, borderRadius:12, padding:"14px 16px" }}>
          <div style={{ fontSize:11, color:"#9ca3af", fontWeight:600, textTransform:"uppercase", marginBottom:6 }}>Quadrante</div>
          <div style={{ fontSize:16, fontWeight:800, color:q.color }}>{q.label}</div>
          <div style={{ fontSize:11, color:T.textMuted, marginTop:4 }}>Tendência: {trendLabel(trend)}</div>
          {dSSM!==null&&<div style={{ fontSize:11, color:T.textMuted }}>{dSSM}d sem contato SSM</div>}
        </div>
        <div style={{ background:cL.bg, border:`1px solid ${cL.border}`, borderRadius:12, padding:"14px 16px", display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ textAlign:"center" }}>
            <div style={{ fontSize:44, fontWeight:900, color:cL.color, fontFamily:"monospace", lineHeight:1 }}>{churn}</div>
            <div style={{ fontSize:10, color:cL.color, fontWeight:600 }}>/ 100</div>
          </div>
          <div>
            <div style={{ fontSize:11, color:"#9ca3af", fontWeight:600, textTransform:"uppercase" }}>Risco de Churn</div>
            <div style={{ fontSize:15, fontWeight:700, color:cL.color }}>{cL.emoji} {cL.label}</div>
            {client.inactivityReason&&<div style={{ fontSize:11, color:"#6b7280", marginTop:2 }}>{client.inactivityReason}</div>}
          </div>
        </div>
        <div style={{ background:rL.bg, border:`1px solid ${rL.border}`, borderRadius:12, padding:"14px 16px", display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ textAlign:"center" }}>
            <div style={{ fontSize:44, fontWeight:900, color:rL.color, fontFamily:"monospace", lineHeight:1 }}>{rescue}</div>
            <div style={{ fontSize:10, color:rL.color, fontWeight:600 }}>/ 100</div>
          </div>
          <div>
            <div style={{ fontSize:11, color:"#9ca3af", fontWeight:600, textTransform:"uppercase" }}>Potencial de Resgate</div>
            <div style={{ fontSize:15, fontWeight:700, color:rL.color }}>{rL.emoji} {rL.label}</div>
            {client.category&&<div style={{ fontSize:11, color:"#6b7280", marginTop:2 }}>{client.category}</div>}
          </div>
        </div>
      </div>

      {/* Info grid */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14, marginBottom:14 }}>
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:12, paddingBottom:8, borderBottom:`1px solid ${T.borderLight}` }}>📋 Contrato</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <Row label="Ativo?" value={client.contractActive}/>
            <Row label="Categoria" value={client.category}/>
            <Row label="Vencimento" value={client.contractEnd}/>
            <Row label="Nº Contratos" value={client.nContracts}/>
          </div>
        </div>
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:12, paddingBottom:8, borderBottom:`1px solid ${T.borderLight}` }}>🤝 Relacionamento SSM</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <Row label="Último Contato SSM" value={client.lastContactSSM}/>
            <Row label="AM Conhece?" value={client.amKnows}/>
            <Row label="Account Manager" value={client.am}/>
            <Row label="Renovação" value={client.renewalHistory}/>
          </div>
        </div>
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:12, paddingBottom:8, borderBottom:`1px solid ${T.borderLight}` }}>🎫 Suporte & Comercial</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <Row label="Tickets (90d)" value={client.tickets90d}/>
            <Row label="Tipo Ticket" value={client.ticketType}/>
            <Row label="Contato Comercial?" value={client.commercialContact}/>
            <Row label="Motivo Inatividade" value={client.inactivityReason}/>
          </div>
        </div>
      </div>

      {/* Consumption */}
      <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20, marginBottom:14 }}>
        <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:16, paddingBottom:8, borderBottom:`1px solid ${T.borderLight}` }}>
          📊 Histórico de Consumo — {MONTHS[0]} a {MONTHS[5]}
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <HistViz label="AWS — Valor mensal" data={client.awsHistory} color="#f97316"/>
          <HistViz label="Azure — Valor mensal" data={client.azureHistory} color="#3b82f6"/>
          <HistViz label="M365-CSP — Licenças ativas" data={client.m365History} color="#f43f5e"/>
          {!cloudH.length&&!(client.m365History||[]).some(v=>v>0)&&(
            <div style={{ fontSize:13, color:T.textSub }}>Nenhum histórico de consumo cadastrado.</div>
          )}
        </div>
      </div>

      {/* Business Reviews */}
      <BRSection client={client} onSave={onSave}/>

      {/* Folder + Observations */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
          <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:8 }}>📁 Diretório do Cliente</div>
          {client.folderPath ? (
            <div>
              <div style={{ fontFamily:"monospace", fontSize:12, color:T.text, background:T.cardAlt, border:`1px solid ${T.border}`, borderRadius:6, padding:"8px 10px", wordBreak:"break-all" }}>{client.folderPath}</div>
              <button onClick={()=>navigator.clipboard.writeText(client.folderPath)}
                style={{ marginTop:8, padding:"6px 12px", borderRadius:6, background:T.btnSecBg, border:`1px solid ${T.btnSecBorder}`, fontSize:12, cursor:"pointer", color:T.btnSecText }}>
                📋 Copiar caminho
              </button>
            </div>
          ) : (
            <div style={{ fontSize:12, color:T.textSub }}>Nenhum diretório configurado. Edite o cliente para adicionar.</div>
          )}
        </div>
        {client.observations&&(
          <div style={{ background:T.card, borderRadius:12, border:`1px solid ${T.border}`, padding:20 }}>
            <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:8 }}>🧩 Observações</div>
            <div style={{ fontSize:13, color:T.text, lineHeight:1.6 }}>{client.observations}</div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [clients, setClients] = useState(loadClients());
  const [view, setView] = useState("dashboard");
  const [selected, setSelected] = useState(null);
  const [importing, setImporting] = useState(false);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [dark, setDark] = useState(() => localStorage.getItem("ssm_dark")==="1");
  const T = dark ? DARK : LIGHT;

  useEffect(() => { localStorage.setItem("ssm_dark", dark?"1":"0"); }, [dark]);

  const showToast = (msg,type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),3000); };

  useEffect(()=>{
    sb.getAll()
      .then(list=>{
        if(list && list.length > 0) {
          setClients(list);
          saveClients(list);
        }
        // Se Supabase retornar vazio mas temos dados locais, manter os locais
      })
      .catch(()=>{ /* Modo offline — mantém dados do localStorage */ })
      .finally(()=>setLoading(false));
  },[]);

  const persist = useCallback(async list=>{ setClients(list); saveClients(list); },[]);

  const handleSave = async updated => {
    try {
      // Atualiza local PRIMEIRO — nunca perde o dado
      const idx = clients.findIndex(c=>c.id===updated.id);
      const next = idx>=0 ? clients.map(c=>c.id===updated.id?updated:c) : [...clients,updated];
      await persist(next);
      setSelected(updated);
      setView("detail");
      showToast(idx>=0?"Cliente atualizado.":"Cliente adicionado.");
      // Envia ao Supabase em background — falha silenciosa não afeta o usuário
      sb.upsert(updated).catch(err => console.warn("Supabase sync error:", err));
    } catch(err) {
      showToast(`Erro ao salvar: ${err.message}`,"error");
    }
  };

  const handleDelete = async id => {
    if(window.confirm("Remover este cliente?")) {
      // Remove local PRIMEIRO
      await persist(clients.filter(c=>c.id!==id));
      setView("list");
      showToast("Cliente removido.", "error");
      // Sincroniza com Supabase em background
      sb.delete(id).catch(err => console.warn("Supabase delete error:", err));
    }
  };

  const handleImport = async file => {
    setImporting(true);
    try {
      const imported = await parseExcelClients(file);
      // Salva local PRIMEIRO
      const map = Object.fromEntries(clients.map(c=>[c.id,c]));
      imported.forEach(c=>{ map[c.id]={folderPath:"",brReviews:[],brFrequency:"Trimestral",...c}; });
      const merged = Object.values(map);
      await persist(merged);
      setView("list");
      showToast(`${imported.length} clientes importados.`);
      // Sincroniza Supabase em background
      (async () => {
        try {
          for(let i=0;i<imported.length;i+=5) {
            await Promise.all(imported.slice(i,i+5).map(c=>sb.upsert({folderPath:"",brReviews:[],brFrequency:"Trimestral",...c})));
            await new Promise(res=>setTimeout(res,400));
          }
        } catch(e) { console.warn("Supabase import sync error:", e); }
      })();
    } catch(err) { showToast(`Erro: ${err.message}`,"error"); }
    finally { setImporting(false); }
  };

  const navTo = (v,client=null) => { setView(v); if(client) setSelected(client); };
  const scored = useMemo(()=>clients.map(c=>({...c,churn:calcChurnRisk(c),rescue:calcRescuePotential(c)})),[clients]);

  return (
    <ThemeCtx.Provider value={T}>
    <div style={{ fontFamily:"'DM Sans','Segoe UI',sans-serif", background:T.bg, minHeight:"100vh", display:"flex", transition:"background 0.2s" }}>
      {toast&&(
        <div style={{ position:"fixed", top:20, right:20, zIndex:999, background:toast.type==="error"?"#fef2f2":"#f0fdf4", border:`1px solid ${toast.type==="error"?"#fecaca":"#bbf7d0"}`, color:toast.type==="error"?"#dc2626":"#16a34a", padding:"10px 18px", borderRadius:10, fontSize:13, fontWeight:600, boxShadow:"0 4px 12px rgba(0,0,0,0.1)" }}>
          {toast.msg}
        </div>
      )}

      {loading&&<div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.8)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12}}><div style={{width:36,height:36,border:"4px solid #1e3a5f",borderTop:"4px solid #93c5fd",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/><div style={{color:"#93c5fd",fontSize:14,fontWeight:600}}>Conectando…</div><style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style></div>}
      {/* Sidebar */}
      <div style={{ width:210, background:"#0f172a", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ padding:"22px 20px 18px", borderBottom:"1px solid #1e293b" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div style={{ fontSize:15, fontWeight:800, color:"#f8fafc", letterSpacing:"-0.02em" }}>SSM CRM</div>
            <button onClick={()=>setDark(d=>!d)} style={{ background:"none", border:"none", cursor:"pointer", fontSize:16, padding:"2px 6px", color:"#94a3b8", lineHeight:1 }} title={dark?"Modo claro":"Modo escuro"}>
              {dark?"☀️":"🌙"}
            </button>
          </div>
          <div style={{ fontSize:11, color:"#475569", marginTop:2 }}>Portfolio Intelligence</div>
        </div>
        <nav style={{ padding:"14px 10px", display:"flex", flexDirection:"column", gap:3 }}>
          {[{id:"dashboard",label:"Dashboard",icon:"📊"},{id:"list",label:"Clientes",icon:"👥"}].map(item=>(
            <button key={item.id} onClick={()=>setView(item.id)}
              style={{ display:"flex", alignItems:"center", gap:9, padding:"9px 12px", borderRadius:8, background:view===item.id?"#1e3a5f":"transparent", color:view===item.id?"#93c5fd":"#64748b", border:"none", cursor:"pointer", fontSize:13, fontWeight:view===item.id?600:400, textAlign:"left", width:"100%" }}>
              <span>{item.icon}</span><span>{item.label}</span>
            </button>
          ))}
        </nav>
        <div style={{ marginTop:"auto", padding:"16px 20px", borderTop:"1px solid #1e293b" }}>
          <div style={{ fontSize:10, color:"#475569", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.05em" }}>Carteira</div>
          {[
            { label:"Total",         val:clients.length,                                        color:"#93c5fd" },
            { label:"🔴 Churn Crítico", val:scored.filter(c=>c.churn>=70).length,               color:"#fca5a5" },
            { label:"🟠 Churn Alto",    val:scored.filter(c=>c.churn>=45&&c.churn<70).length,   color:"#fdba74" },
            { label:"⭐ Alto Potencial", val:scored.filter(c=>c.rescue>=65).length,              color:"#c4b5fd" },
          ].map(s=>(
            <div key={s.label} style={{ display:"flex", justifyContent:"space-between", marginBottom:5 }}>
              <span style={{ fontSize:12, color:"#94a3b8" }}>{s.label}</span>
              <span style={{ fontSize:12, fontWeight:700, color:s.color, fontFamily:"monospace" }}>{s.val}</span>
            </div>
          ))}
          <div style={{ marginTop:10, fontSize:10, color:"#64748b" }}>{MONTHS[0]} → {MONTHS[5]}</div>
        </div>
      </div>

      {/* Content */}
      <div style={{ flex:1, overflow:"auto", background:T.bg }}>
        <div style={{ padding:28, maxWidth:1300 }}>
          {view!=="detail"&&view!=="new"&&(
            <div style={{ marginBottom:22 }}>
              <h1 style={{ margin:0, fontSize:20, fontWeight:800, color:T.text }}>
                {view==="dashboard"?"📊 Dashboard":"👥 Clientes"}
              </h1>
              <div style={{ fontSize:13, color:T.textMuted, marginTop:3 }}>
                {view==="dashboard"?"Visão geral · "+MONTHS[0]+" a "+MONTHS[5]:clients.length+" clientes cadastrados"}
              </div>
            </div>
          )}
          {view==="dashboard"&&<Dashboard clients={clients} onNavigate={navTo}/>}
          {view==="list"&&<ClientList clients={clients} onSelect={c=>navTo("detail",c)} onAdd={()=>setView("new")} onImport={handleImport} importing={importing}/>}
          {view==="detail"&&selected&&<ClientDetail client={clients.find(c=>c.id===selected.id)||selected} onSave={handleSave} onBack={()=>setView("list")} onDelete={handleDelete}/>}
          {view==="new"&&(
            <div>
              <button onClick={()=>setView("list")} style={{ background:"none", border:"none", color:T.textMuted, cursor:"pointer", fontSize:13, marginBottom:16, padding:0 }}>← Voltar</button>
              <h2 style={{ margin:"0 0 20px", fontSize:18, fontWeight:800, color:T.text }}>+ Novo Cliente</h2>
              <ClientForm onSave={handleSave} onCancel={()=>setView("list")}/>
            </div>
          )}
        </div>
      </div>
    </div>
    </ThemeCtx.Provider>
  );
}

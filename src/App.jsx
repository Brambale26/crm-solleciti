import { useState, useEffect, useMemo } from "react";

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxM-HnsnxAehTsVTp_YxmVOKb4xuVAMSc33wzRH0mVhRclFkRrxG7-vxIDb7SQiLgg/exec";

const THEMES = {
  dark: {
    bg: "#0a0a0f", bgCard: "#13131a", bgAlt: "#0d0d14", bgInput: "#13131a",
    border: "#1e1e2a", borderInput: "#2a2a38",
    text: "#e8e6df", textSub: "#6a6878", textMuted: "#4a4858",
    navBg: "#0d0d14", navBorder: "#1a1a25",
    warnBg: "#1f1808", warnBorder: "#3a2e08",
    infoBg: "#0a1f2a", infoBorder: "#1a3a4a",
    dangerBg: "#2a0a0a", dangerBorder: "#4a1a1a",
    successBg: "#0f2a1a", successBorder: "#1a4a2a",
    diario: {
      "Mio contatto":  { border: "#378ADD", bg: "#0a1828", label: "#378ADD" },
      "Via agente":    { border: "#4ecb8d", bg: "#0a2018", label: "#4ecb8d" },
      "Nota personale":{ border: "#5a5868", bg: "#0d0d14", label: "#7a7888" },
      "Cliente":       { border: "#9a98a0", bg: "#121218", label: "#9a98a0" },
    }
  },
  light: {
    bg: "#f4f5f7", bgCard: "#ffffff", bgAlt: "#f0f1f3", bgInput: "#ffffff",
    border: "#e2e4e9", borderInput: "#d0d3db",
    text: "#1a1a2e", textSub: "#6b6f7d", textMuted: "#9a9fad",
    navBg: "#ffffff", navBorder: "#e2e4e9",
    warnBg: "#fffbf0", warnBorder: "#f5d87a",
    infoBg: "#f0f7ff", infoBorder: "#90c4f0",
    dangerBg: "#fff5f5", dangerBorder: "#f0a0a0",
    successBg: "#f0fff8", successBorder: "#90e0b8",
    diario: {
      "Mio contatto":  { border: "#378ADD", bg: "#eaf3fd", label: "#2265b0" },
      "Via agente":    { border: "#1D9E75", bg: "#e6f7f2", label: "#0f6e56" },
      "Nota personale":{ border: "#9a9fad", bg: "#f4f5f7", label: "#6b6f7d" },
      "Cliente":       { border: "#c0c3cb", bg: "#f9f9fb", label: "#9a9fad" },
    }
  }
};

const STATI_OP = ["CONTROLLA", "AGENTE", "AGENZIA", "L/L", "RICHIAMA", "PAGATO", "PROBLEMA"];
const PRIORITA = ["ALTA", "MEDIA", "BASSA"];
const TIPI_CONTATTO = ["Telefono", "Email", "WhatsApp", "Altro"];
const ESITI_RACC = ["In attesa", "Consegnata", "Non consegnata", "Rifiutata", "Indirizzo errato"];
const TIPI_TITOLO = ["Assegno a vista", "Assegno postdatato", "Cambiale", "Contanti", "Bonifico"];
const TIPI_DIARIO = ["Mio contatto", "Via agente", "Nota personale", "Cliente"];
const STORAGE_KEY = "crm_solleciti_v3";

const STATO_COLORS = {
  "CONTROLLA": "#378ADD", "AGENTE": "#EF9F27", "AGENZIA": "#7F77DD",
  "L/L": "#e24b4a", "RICHIAMA": "#4ecb8d", "PAGATO": "#1D9E75", "PROBLEMA": "#FF4A8D"
};
const PRIORITA_COLORS = { "ALTA": "#e24b4a", "MEDIA": "#EF9F27", "BASSA": "#4ecb8d" };
const IBAN = "IT79G0306909496100000011059";
const AZIENDA = "Saratoga Int. Sforza SPA";
const BANCA = "Banca Intesa San Paolo - Via Lorenteggio 70";

const NAV_SECTIONS = [
  { section: "Operativo", items: [
    { id: "oggi",        label: "Oggi",       icon: "ti-calendar-event" },
    { id: "clienti",     label: "Clienti",    icon: "ti-users" },
    { id: "agenti",      label: "Agenti",     icon: "ti-briefcase" },
  ]},
  { section: "Gestione", items: [
    { id: "ll",          label: "L/L",        icon: "ti-file-text" },
    { id: "incassi",     label: "Incassi",    icon: "ti-coin" },
  ]},
  { section: "Analytics", items: [
    { id: "dashboard",   label: "Dashboard",  icon: "ti-chart-bar" },
  ]},
  { section: "Anagrafica", items: [
    { id: "ana-clienti", label: "Clienti",    icon: "ti-user" },
    { id: "ana-agenti",  label: "Agenti",     icon: "ti-id-badge" },
  ]},
];

function loadLocal() {
  try {
    const r = localStorage.getItem(STORAGE_KEY);
    if (r) return JSON.parse(r);
    // migrate from v2
    const v2 = localStorage.getItem("crm_solleciti_v2") || localStorage.getItem("crm_solleciti_v1");
    if (v2) return { ...JSON.parse(v2), agenti: JSON.parse(v2).agenti || [], insoluti: JSON.parse(v2).insoluti || [] };
    return { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [] };
  } catch { return { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [] }; }
}
function saveLocal(d) { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); }
async function loadFromSheets() {
  try {
    const url = SCRIPT_URL + "?action=load";
    const r = await fetch("https://api.allorigins.win/get?url=" + encodeURIComponent(url));
    const json = await r.json();
    return JSON.parse(json.contents);
  } catch(e) { return {}; }
}
async function saveToSheets(d) {
  try {
    const url = SCRIPT_URL + "?action=save&data=" + encodeURIComponent(JSON.stringify(d));
    await fetch("https://api.allorigins.win/get?url=" + encodeURIComponent(url));
  } catch(e) { console.error("Sync error", e); }
}

function oggi() { return new Date().toISOString().slice(0, 10); }
function oraOra() { return new Date().toTimeString().slice(0, 5); }
function fmtEur(n) {
  if (n == null || isNaN(n)) return "—";
  return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(n);
}
function fmtData(d) {
  if (!d) return "—";
  if (typeof d === "number") return new Date((d - 25569) * 86400000).toLocaleDateString("it-IT");
  if (String(d).includes("-")) { const [y, m, dd] = String(d).split("-"); return dd + "/" + m + "/" + y; }
  return d;
}
function giorniDa(ds) { if (!ds) return 9999; return Math.floor((new Date() - new Date(ds)) / 86400000); }
function addGiorni(ds, n) { if (!ds) return ""; const d = new Date(ds); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); }

async function parseSpaccatureExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const sn = wb.SheetNames.find(s => s.toUpperCase().includes("SCADUT") || s.toUpperCase().includes("SPACCATURE")) || wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
        if (raw.length < 2) { resolve([]); return; }
        const headers = raw[0];
        const si = headers.findIndex(h => h instanceof Date || (typeof h === "string" && /\d{2}\/\d{2}\/\d{4}/.test(h)));
        resolve(raw.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[String(h)] = row[i]; });
          if (si >= 0) obj["__SCADENZA__"] = row[si];
          return obj;
        }));
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function elaboraSpaccature(rows) {
  const map = {};
  rows.forEach(row => {
    const cod = String(row["COD."] || row["CODICE"] || "").trim().padStart(5, "0");
    if (!cod || cod === "00000") return;
    const ragione = String(row["RAGIONE SOCIALE"] || "").trim();
    const agente = String(row["AGENTE"] || "").trim();
    const localita = String(row["LOCALITA\'"] || row["LOCALITA"] || "").trim();
    const prov = String(row["PROV."] || row["PROV"] || "").trim();
    const doc = String(row["DOC."] || row["DOC"] || "").trim();
    const sr = row["__SCADENZA__"] || row["SCADENZE"] || row["SCADENZA"] || "";
    let scadenza = "";
    if (sr instanceof Date) scadenza = sr.toISOString().slice(0, 10);
    else if (typeof sr === "number" && sr > 40000) scadenza = new Date((sr - 25569) * 86400000).toISOString().slice(0, 10);
    else scadenza = String(sr);
    const importo = parseFloat(row["IMPORTO"] || 0);
    const giorni = parseInt(row["GIORNI"] || 0);
    const tag = cod + "|" + doc + "|" + scadenza + "|" + importo.toFixed(2);
    if (!map[cod]) {
      map[cod] = { id: cod, codice: cod, ragione, agente, localita, prov, email: "", telefono: "", cellulare: "", orari: "", personeRif: "", statoOp: "CONTROLLA", statoCl: "ATTIVO", priorita: "ALTA", esito: "", ultimoContatto: null, dataRichiamo: null, note: "", diario: [], scadenze: [], totaleScaduto: 0, giorniMaxRitardo: 0 };
    }
    const cl = map[cod];
    if (ragione && !cl.ragione) cl.ragione = ragione;
    if (agente && !cl.agente) cl.agente = agente;
    if (!cl.scadenze.find(s => s.tag === tag)) cl.scadenze.push({ tag, doc, scadenza, importo, giorni, residuo: importo, stato: "APERTA", accorpata: false, fatturaAccorpante: null });
    cl.totaleScaduto += importo;
    if (giorni > cl.giorniMaxRitardo) cl.giorniMaxRitardo = giorni;
  });
  return Object.values(map);
}

function mergeClienti(esistenti, nuovi) {
  const result = [...esistenti];
  nuovi.forEach(nc => {
    const es = result.find(c => c.codice === nc.codice);
    if (es) {
      es.ragione = nc.ragione || es.ragione; es.agente = nc.agente || es.agente; es.localita = nc.localita || es.localita;
      nc.scadenze.forEach(ns => { if (!es.scadenze.find(s => s.tag === ns.tag)) es.scadenze.push(ns); });
      es.totaleScaduto = es.scadenze.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
    } else { result.push(nc); }
  });
  return result;
}

function generaMailSollecito(cliente) {
  const sc = (cliente.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  if (!sc.length || !cliente.email) return null;
  let corpo = "Gentile Cliente,\n\na seguito di controlli contabili risultano ancora in sospeso i pagamenti relativi alle seguenti fatture:\n\n";
  sc.forEach(s => { corpo += "- Fattura n. " + s.doc + " scaduta il " + fmtData(s.scadenza) + " - EUR " + (s.residuo || s.importo).toFixed(2) + "\n\n"; });
  corpo += "Al fine di agevolare la procedura, riportiamo le coordinate per il pagamento.\n\n" + IBAN + "\n" + AZIENDA + "\n" + BANCA + "\n\nQualora abbiate gia provveduto ritenete nullo il presente sollecito.\n\nCordiali saluti.";
  return "mailto:" + cliente.email + "?subject=" + encodeURIComponent("Verifica stato pagamento fatture") + "&body=" + encodeURIComponent(corpo);
}
function generaMailInsolutoCliente(cliente, ins) {
  if (!cliente.email) return null;
  const corpo = "Gentile Cliente,\n\nLa informiamo che l\'assegno n. " + (ins.riferimento || "-") + " di EUR " + (ins.importo || 0).toFixed(2) + " del " + fmtData(ins.dataEmissione) + " e tornato insoluto per " + (ins.motivo || "mancanza fondi") + ".\n\nLa preghiamo di provvedere entro 60 giorni.\n\n" + IBAN + "\n" + AZIENDA + "\n\nCordiali saluti.";
  return "mailto:" + cliente.email + "?subject=" + encodeURIComponent("Assegno insoluto - regolarizzazione entro 60 giorni") + "&body=" + encodeURIComponent(corpo);
}
function generaMailInsolutoAgente(agente, cliente, ins) {
  if (!agente?.email) return null;
  const corpo = "Gentile " + agente.nome + ",\n\nTi informo che l\'assegno del cliente " + cliente.ragione + " (cod. " + cliente.codice + ") di EUR " + (ins.importo || 0).toFixed(2) + " del " + fmtData(ins.dataEmissione) + " e tornato insoluto per " + (ins.motivo || "mancanza fondi") + ".\n\nTi chiedo di non accettare ulteriori titoli da questo cliente.\n\nCordiali saluti.";
  return "mailto:" + agente.email + "?subject=" + encodeURIComponent("Segnalazione insoluto " + cliente.ragione + " " + cliente.codice) + "&body=" + encodeURIComponent(corpo);
}

function getGlobalStyles(T) {
  return `
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
    @import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');
    :root{--bg:${T.bg};--bgc:${T.bgCard};--bga:${T.bgAlt};--bgi:${T.bgInput};--br:${T.border};--bri:${T.borderInput};--tx:${T.text};--txs:${T.textSub};--txm:${T.textMuted};--wb:${T.warnBg};--wbr:${T.warnBorder};--ib:${T.infoBg};--ibr:${T.infoBorder};--db:${T.dangerBg};--dbr:${T.dangerBorder};--sb:${T.successBg};--sbr:${T.successBorder};}
    *{box-sizing:border-box;margin:0;padding:0}
    body{background:var(--bg);color:var(--tx);font-family:'DM Sans','Segoe UI',sans-serif}
    input,select,textarea{background:var(--bgi);border:1px solid var(--bri);border-radius:10px;color:var(--tx);padding:10px 14px;font-family:inherit;font-size:15px;width:100%;outline:none;-webkit-appearance:none}
    input:focus,select:focus,textarea:focus{border-color:#4a7fd4}
    select option{background:var(--bgi);color:var(--tx)}
    textarea{resize:vertical;min-height:80px}
    button{cursor:pointer;font-family:inherit;font-size:14px;border:none;border-radius:10px;padding:10px 18px;transition:all .15s;-webkit-tap-highlight-color:transparent}
    .btn-p{background:#4a7fd4;color:#fff;font-weight:500}
    .btn-p:active{transform:scale(.98)}
    .btn-g{background:transparent;color:var(--txs);border:1px solid var(--bri)}
    .btn-g:hover{background:var(--bga)}
    .btn-d{background:transparent;color:#e24b4a;border:1px solid var(--dbr)}
    .btn-ok{background:var(--sb);color:#4ecb8d;border:1px solid var(--sbr)}
    .card{background:var(--bgc);border:1px solid var(--br);border-radius:16px;padding:16px}
    .tag{display:inline-block;font-size:11px;font-weight:500;padding:3px 9px;border-radius:20px}
    label{font-size:12px;color:var(--txs);margin-bottom:5px;display:block;letter-spacing:.04em}
    .mbg{position:fixed;inset:0;background:rgba(0,0,0,.7);display:flex;align-items:flex-end;justify-content:center;z-index:300}
    .mo{background:var(--bgc);border:1px solid var(--bri);border-radius:20px 20px 0 0;padding:24px 20px 40px;width:100%;max-width:520px;max-height:92vh;overflow-y:auto}
    .mono{font-family:'DM Mono',monospace}
    @keyframes spin{to{transform:rotate(360deg)}}
    .spin{animation:spin 1s linear infinite;display:inline-block}
    .fg{margin-bottom:14px}
    .r2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .tabs{display:flex;gap:8px;margin-bottom:16px;overflow-x:auto;padding-bottom:2px}
    .tab{padding:8px 16px;border-radius:20px;font-size:13px;border:1px solid var(--bri);background:transparent;color:var(--txs);white-space:nowrap;cursor:pointer}
    .tab.on{background:#4a7fd4;color:#fff;border-color:transparent}
    .cc{background:var(--bgc);border:1px solid var(--br);border-radius:14px;padding:14px;margin-bottom:10px;cursor:pointer;transition:border-color .15s}
    .cc:hover{border-color:#4a7fd480}
    .kv{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--br);font-size:13px}
    .kv:last-child{border-bottom:none}
    .wb{background:var(--wb);border:1px solid var(--wbr);border-radius:10px;padding:10px 14px;font-size:13px;color:#EF9F27;margin-bottom:14px}
    .ib{background:var(--ib);border:1px solid var(--ibr);border-radius:10px;padding:10px 14px;font-size:13px;color:#378ADD;margin-bottom:14px}
    .db{background:var(--db);border:1px solid var(--dbr);border-radius:10px;padding:10px 14px;font-size:13px;color:#e24b4a;margin-bottom:14px}
    .st{font-size:22px;font-weight:600;letter-spacing:-.02em;margin-bottom:4px;color:var(--tx)}
    .sw{position:relative;margin-bottom:16px}
    .si{background:var(--bga);border:1px solid var(--br);border-radius:12px;color:var(--tx);padding:10px 14px 10px 38px;font-size:15px;width:100%;outline:none}
    .sic{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--txm);font-size:18px}
    .sr{background:var(--bga);border-radius:10px;padding:10px 12px;margin-bottom:8px}
    .pb{width:4px;border-radius:2px;align-self:stretch;flex-shrink:0}
    .nav-item{width:100%;display:flex;align-items:center;border:none;padding:9px 10px;border-radius:10px;background:transparent;color:var(--txs);font-weight:400;font-size:14px;transition:all .15s;cursor:pointer;margin-bottom:2px}
    .nav-item.on{background:#4a7fd420;color:#4a7fd4;font-weight:500}
    .nav-item:hover:not(.on){background:var(--bga)}
  `;
}

export default function App() {
  const [page, setPage] = useState("oggi");
  const [data, setData] = useState(loadLocal);
  const [toast, setToast] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [loadingCloud, setLoadingCloud] = useState(true);
  const [themeMode, setThemeMode] = useState(() => localStorage.getItem("crm_theme") || "dark");
  const [navOpen, setNavOpen] = useState(() => localStorage.getItem("crm_nav") !== "closed");
  const [clienteSel, setClienteSel] = useState(null);
  const [agenteSel, setAgenteSel] = useState(null);
  const [navStack, setNavStack] = useState([]);
  const T = THEMES[themeMode];

  function toggleTheme() { const n = themeMode === "dark" ? "light" : "dark"; setThemeMode(n); localStorage.setItem("crm_theme", n); }
  function toggleNav() { const n = !navOpen; setNavOpen(n); localStorage.setItem("crm_nav", n ? "open" : "closed"); }

  function goTo(newPage) { setNavStack([]); setClienteSel(null); setAgenteSel(null); setPage(newPage); }
  function openCliente(id) { setNavStack(s => [...s, { page, clienteSel, agenteSel }]); setClienteSel(id); setAgenteSel(null); }
  function openAgente(id) { setNavStack(s => [...s, { page, clienteSel, agenteSel }]); setAgenteSel(id); setClienteSel(null); }
  function goBack() {
    const prev = navStack[navStack.length - 1];
    if (!prev) { setClienteSel(null); setAgenteSel(null); return; }
    setNavStack(s => s.slice(0, -1));
    setClienteSel(prev.clienteSel); setAgenteSel(prev.agenteSel);
    if (!prev.clienteSel && !prev.agenteSel) setPage(prev.page);
  }

  useEffect(() => {
    loadFromSheets().then(remote => {
      if (remote && !remote.error && (remote.clienti?.length > 0)) {
        const merged = { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [], ...remote };
        setData(merged); saveLocal(merged);
      }
    }).catch(() => {}).finally(() => setLoadingCloud(false));
  }, []);

  useEffect(() => { saveLocal(data); }, [data]);

  function showToast(msg, type = "ok") { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); }
  async function sync(nd) { setSyncing(true); try { await saveToSheets(nd); } catch { showToast("Errore sync", "warn"); } setSyncing(false); }

  function updCliente(id, u) { setData(d => { const nd = { ...d, clienti: d.clienti.map(c => c.id === id ? { ...c, ...u } : c) }; sync(nd); return nd; }); }
  function addDiario(id, v) { setData(d => { const nd = { ...d, clienti: d.clienti.map(c => c.id === id ? { ...c, diario: [v, ...(c.diario || [])], ultimoContatto: v.tipo !== "Nota personale" ? v.data : c.ultimoContatto } : c) }; sync(nd); return nd; }); }

  async function importa(file) {
    try {
      showToast("Importazione...", "info");
      const rows = await parseSpaccatureExcel(file);
      const nuovi = elaboraSpaccature(rows);
      setData(d => {
        const cl = mergeClienti(d.clienti || [], nuovi);
        const agE = d.agenti || [];
        const nomiN = [...new Set(nuovi.map(c => c.agente).filter(Boolean))];
        const nuoviAg = nomiN.filter(n => !agE.find(a => a.nome === n)).map(n => ({ id: Date.now() + Math.random(), nome: n, email: "", telefono: "", zona: "" }));
        const tot = cl.filter(c => c.statoCl !== "PAGATO").reduce((s, c) => s + (c.totaleScaduto || 0), 0);
        const nd = { ...d, clienti: cl, agenti: [...agE, ...nuoviAg], ultimoImport: { data: oggi(), scaduto: tot } };
        sync(nd); showToast(cl.length + " clienti - " + nuoviAg.length + " nuovi agenti"); return nd;
      });
    } catch (e) { showToast("Errore: " + e.message, "warn"); }
  }

  function addLL(cid) {
    const cl = data.clienti.find(c => c.id === cid); if (!cl) return;
    setData(d => { const nd = { ...d, praticheLl: [...(d.praticheLl || []), { id: Date.now(), clienteId: cid, ragione: cl.ragione, agente: cl.agente, stato: "IN CODA", lettera: 0, dataDecisione: oggi(), dataUltimaLettera: null, esito: "" }] }; showToast("Pratica L/L aperta"); sync(nd); return nd; });
  }
  function updLL(id, u) { setData(d => { const nd = { ...d, praticheLl: d.praticheLl.map(p => p.id === id ? { ...p, ...u } : p) }; sync(nd); return nd; }); }
  function addIncasso(inc) {
    const stato = inc.tipo === "Bonifico" ? "REGISTRATO" : "IN MANO AGENTE";
    setData(d => { const nd = { ...d, incassi: [{ ...inc, id: Date.now(), dataRicezione: oggi(), dataRegistrazione: stato === "REGISTRATO" ? oggi() : null, stato }, ...(d.incassi || [])] }; showToast(stato === "REGISTRATO" ? "Bonifico registrato" : "Incasso registrato"); sync(nd); return nd; });
  }
  function updIncasso(id, u) { setData(d => { const nd = { ...d, incassi: d.incassi.map(i => i.id === id ? { ...i, ...u } : i) }; sync(nd); return nd; }); }
  function delIncasso(id) { setData(d => { const nd = { ...d, incassi: d.incassi.filter(i => i.id !== id) }; showToast("Eliminato"); sync(nd); return nd; }); }
  function addAgente(ag) { setData(d => { const nd = { ...d, agenti: [...(d.agenti || []), { ...ag, id: Date.now() }] }; showToast("Agente aggiunto"); sync(nd); return nd; }); }
  function updAgente(id, u) { setData(d => { const nd = { ...d, agenti: d.agenti.map(a => a.id === id ? { ...a, ...u } : a) }; sync(nd); return nd; }); }
  function addInsoluto(ins) {
    setData(d => { const nd = { ...d, clienti: d.clienti.map(c => c.id === ins.clienteId ? { ...c, statoOp: "PROBLEMA" } : c), insoluti: [{ ...ins, id: Date.now(), data: oggi(), stato: "APERTO", scadenza60gg: addGiorni(oggi(), 60) }, ...(d.insoluti || [])] }; showToast("Insoluto - cliente PROBLEMA"); sync(nd); return nd; });
  }

  const listaOggi = useMemo(() => {
    const attivi = (data.clienti || []).filter(c => c.statoCl !== "PAGATO" && c.statoOp !== "PAGATO");
    return [...attivi].sort((a, b) => {
      const aR = a.dataRichiamo === oggi() ? 0 : 1, bR = b.dataRichiamo === oggi() ? 0 : 1;
      if (aR !== bR) return aR - bR;
      const ag = giorniDa(a.ultimoContatto), bg = giorniDa(b.ultimoContatto);
      if (ag !== bg) return bg - ag;
      return (a.agente || "").localeCompare(b.agente || "");
    });
  }, [data.clienti]);

  const kpi = useMemo(() => {
    const cl = data.clienti || [], inc = data.incassi || [], mese = oggi().slice(0, 7);
    const totS = cl.filter(c => c.statoCl !== "PAGATO").reduce((s, c) => s + (c.totaleScaduto || 0), 0);
    const confT = inc.filter(i => i.stato === "REGISTRATO").reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const confM = inc.filter(i => i.stato === "REGISTRATO" && (i.dataRegistrazione || "").startsWith(mese)).reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const attT = inc.filter(i => i.stato === "IN MANO AGENTE" || i.stato === "SPEDITO").reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const attM = inc.filter(i => (i.stato === "IN MANO AGENTE" || i.stato === "SPEDITO") && (i.dataRicezione || "").startsWith(mese)).reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const si = data.ultimoImport?.scaduto || totS;
    return { totS, si, confT, confM, attT, attM, tot: confT + attT, totM: confM + attM, pctC: totS > 0 ? (confT / totS) * 100 : 0, pctM: si > 0 ? ((confM + attM) / si) * 100 : 0, ll: (data.praticheLl || []).filter(p => p.stato !== "AGENZIA" && p.stato !== "CHIUSA").length, tm: inc.filter(i => i.stato === "IN MANO AGENTE").length, rc: listaOggi.filter(c => c.dataRichiamo === oggi()).length, ins: (data.insoluti || []).filter(i => i.stato === "APERTO").length, di: data.ultimoImport?.data || null };
  }, [data, listaOggi]);

  const SW = navOpen ? 220 : 56;
  const cl = clienteSel ? data.clienti.find(c => c.id === clienteSel) : null;
  const ag = agenteSel ? (data.agenti.find(a => a.id === agenteSel) || { id: agenteSel, nome: String(agenteSel), email: "", telefono: "", zona: "" }) : null;
  const canBack = navStack.length > 0 || clienteSel || agenteSel;

  return (
    <div style={{ minHeight: "100vh", background: T.bg, color: T.text, fontFamily: "'DM Sans','Segoe UI',sans-serif", display: "flex" }}>
      <style>{getGlobalStyles(T)}</style>

      {/* SIDEBAR */}
      <aside style={{ width: SW, minHeight: "100vh", background: T.navBg, borderRight: "1px solid " + T.navBorder, display: "flex", flexDirection: "column", flexShrink: 0, transition: "width .2s ease", overflow: "hidden", position: "sticky", top: 0, height: "100vh" }}>
        <div style={{ padding: "18px 12px 14px", display: "flex", alignItems: "center", gap: 10, borderBottom: "1px solid " + T.navBorder }}>
          <button onClick={toggleNav} style={{ background: "transparent", border: "none", color: T.textSub, fontSize: 20, padding: 4, flexShrink: 0, cursor: "pointer" }}>
            <i className="ti ti-menu-2" />
          </button>
          {navOpen && <div><div style={{ fontSize: 14, fontWeight: 600 }}>CRM Solleciti</div><div style={{ fontSize: 10, color: T.textMuted, letterSpacing: ".06em" }}>SARATOGA</div></div>}
        </div>
        <nav style={{ flex: 1, padding: "8px", overflowY: "auto" }}>
          {NAV_SECTIONS.map(sec => (
            <div key={sec.section} style={{ marginBottom: 6 }}>
              {navOpen && <div style={{ fontSize: 10, color: T.textMuted, letterSpacing: ".08em", padding: "6px 8px 3px", textTransform: "uppercase" }}>{sec.section}</div>}
              {sec.items.map(item => {
                const isOn = page === item.id && !clienteSel && !agenteSel;
                return (
                  <button key={item.id} onClick={() => goTo(item.id)} className={"nav-item" + (isOn ? " on" : "")} style={{ justifyContent: navOpen ? "flex-start" : "center", gap: navOpen ? 10 : 0 }}>
                    <i className={"ti " + item.icon} style={{ fontSize: 18, flexShrink: 0 }} />
                    {navOpen && <span style={{ whiteSpace: "nowrap" }}>{item.label}</span>}
                  </button>
                );
              })}
            </div>
          ))}
        </nav>
        <div style={{ padding: "10px 8px", borderTop: "1px solid " + T.navBorder }}>
          <button onClick={toggleTheme} style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: navOpen ? "flex-start" : "center", gap: 10, padding: navOpen ? "8px 10px" : "8px 0", background: "transparent", border: "none", color: T.textSub, fontSize: 14, borderRadius: 8, cursor: "pointer" }}>
            <span style={{ fontSize: 18 }}>{themeMode === "dark" ? "☀️" : "🌙"}</span>
            {navOpen && <span>{themeMode === "dark" ? "Tema chiaro" : "Tema scuro"}</span>}
          </button>
          {navOpen && <div style={{ fontSize: 10, color: T.textMuted, padding: "4px 10px" }}>
            {syncing ? <><span className="spin">↻</span> Sync...</> : kpi.di ? "Import: " + fmtData(kpi.di) : "Nessun import"}
          </div>}
        </div>
      </aside>

      {/* MAIN */}
      <main style={{ flex: 1, padding: "24px 28px 40px", overflowY: "auto", minWidth: 0 }}>
        {/* TOP BAR */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {canBack && <button className="btn-g" style={{ padding: "7px 14px", fontSize: 13 }} onClick={goBack}><i className="ti ti-arrow-left" /> Indietro</button>}
            {(cl || ag) && <div style={{ fontSize: 13, color: T.textSub, fontWeight: 500 }}>{cl ? cl.ragione : ag ? ag.nome : ""}</div>}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {kpi.rc > 0 && <span style={{ background: T.warnBg, border: "1px solid " + T.warnBorder, color: "#EF9F27", borderRadius: 8, padding: "3px 10px", fontSize: 12 }}>📅 {kpi.rc}</span>}
            {kpi.ins > 0 && <span style={{ background: T.dangerBg, border: "1px solid " + T.dangerBorder, color: "#e24b4a", borderRadius: 8, padding: "3px 10px", fontSize: 12 }}>⚠ {kpi.ins}</span>}
            {kpi.tm > 0 && <span style={{ background: T.infoBg, border: "1px solid " + T.infoBorder, color: "#378ADD", borderRadius: 8, padding: "3px 10px", fontSize: 12 }}>💼 {kpi.tm}</span>}
          </div>
        </div>

        {loadingCloud ? (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "50vh", gap: 16, color: T.textMuted }}>
            <span className="spin" style={{ fontSize: 36 }}>↻</span>
            <div>Carico dal cloud...</div>
          </div>
        ) : cl ? (
          <SchedaCliente cliente={cl} agenti={data.agenti || []} T={T} showToast={showToast} onUpdate={u => updCliente(cl.id, u)} onAddDiario={v => addDiario(cl.id, v)} onApriLL={() => { addLL(cl.id); goTo("ll"); }} onAddIncasso={addIncasso} onAddInsoluto={addInsoluto} />
        ) : ag ? (
          <SchedaAgente agente={ag} clienti={(data.clienti || []).filter(c => c.agente === ag.nome)} incassi={(data.incassi || []).filter(i => i.agente === ag.nome && i.stato === "IN MANO AGENTE")} T={T} showToast={showToast} onUpdate={u => updAgente(ag.id, u)} onAddDiario={(cid, v) => addDiario(cid, v)} onUpdCliente={updCliente} onUpdIncasso={updIncasso} onDelIncasso={delIncasso} onSelectCliente={openCliente} />
        ) : (
          <>
            {page === "oggi" && <Oggi lista={listaOggi} insoluti={data.insoluti || []} T={T} onSel={openCliente} />}
            {page === "clienti" && <Clienti clienti={data.clienti || []} T={T} onSel={openCliente} onImporta={importa} />}
            {page === "agenti" && <Agenti agenti={data.agenti || []} clienti={data.clienti || []} incassi={data.incassi || []} T={T} onAdd={addAgente} onSel={openAgente} />}
            {page === "ll" && <PraticheLl pratiche={data.praticheLl || []} clienti={data.clienti || []} T={T} onUpd={updLL} onSelCl={openCliente} showToast={showToast} />}
            {page === "incassi" && <Incassi incassi={data.incassi || []} clienti={data.clienti || []} T={T} onAdd={addIncasso} onUpd={updIncasso} onDel={delIncasso} onSelCl={openCliente} showToast={showToast} />}
            {page === "dashboard" && <Dashboard kpi={kpi} clienti={data.clienti || []} agenti={data.agenti || []} insoluti={data.insoluti || []} T={T} onSelAg={openAgente} />}
            {page === "ana-clienti" && <AnagraficaClienti clienti={data.clienti || []} T={T} onUpd={updCliente} showToast={showToast} />}
            {page === "ana-agenti" && <AnagraficaAgenti agenti={data.agenti || []} T={T} onUpd={updAgente} onAdd={addAgente} showToast={showToast} />}
          </>
        )}
      </main>

      {toast && (
        <div style={{ position: "fixed", bottom: 24, right: 24, background: toast.type === "warn" ? T.warnBg : toast.type === "info" ? T.infoBg : T.successBg, border: "1px solid " + (toast.type === "warn" ? T.warnBorder : toast.type === "info" ? T.infoBorder : T.successBorder), color: toast.type === "warn" ? "#EF9F27" : toast.type === "info" ? "#378ADD" : "#4ecb8d", padding: "12px 20px", borderRadius: 12, fontSize: 14, zIndex: 400, boxShadow: "0 4px 20px rgba(0,0,0,.15)" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

/* --- OGGI --- */
function Oggi({ lista, insoluti, T, onSel }) {
  const o = oggi();
  const ins = insoluti.filter(i => i.stato === "APERTO");
  const G = ({ label, items, col }) => !items.length ? null : (
    <div style={{ marginBottom: 20 }}>
      <div style={{ fontSize: 11, color: col, letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 10, fontWeight: 500 }}>{label} ({items.length})</div>
      {items.map(c => <CC key={c.id} c={c} T={T} onClick={() => onSel(c.id)} />)}
    </div>
  );
  return (
    <div>
      <div className="st">Oggi</div>
      <div style={{ color: T.textSub, fontSize: 13, marginBottom: 20 }}>{lista.length} clienti attivi · {new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long" })}</div>
      {ins.length > 0 && <div className="db">⚠ {ins.length} insoluti aperti</div>}
      <G label="📅 Richiami oggi" items={lista.filter(c => c.dataRichiamo === o)} col="#EF9F27" />
      <G label="⚠ Problema" items={lista.filter(c => c.statoOp === "PROBLEMA" && c.dataRichiamo !== o)} col="#e24b4a" />
      <G label="🆕 Mai contattati" items={lista.filter(c => !c.ultimoContatto && c.dataRichiamo !== o && c.statoOp !== "PROBLEMA")} col="#e24b4a" />
      <G label="📞 Da richiamare" items={lista.filter(c => c.ultimoContatto && c.dataRichiamo !== o && c.statoOp !== "PROBLEMA")} col={T.textSub} />
      {!lista.length && <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}><i className="ti ti-calendar-off" style={{ fontSize: 40, display: "block", marginBottom: 12 }} />Importa le spaccature dalla sezione Clienti</div>}
    </div>
  );
}

function CC({ c, T, onClick }) {
  const gg = giorniDa(c.ultimoContatto);
  return (
    <div className="cc" onClick={onClick} style={{ display: "flex", gap: 12, alignItems: "stretch" }}>
      <div className="pb" style={{ background: PRIORITA_COLORS[c.priorita] || "#888" }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2 }}>{c.ragione}</div>
            <div style={{ fontSize: 11, color: T.textSub }}>{c.agente} · {c.codice}</div>
          </div>
          <div style={{ textAlign: "right", flexShrink: 0 }}>
            <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
            <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888", display: "inline-block", marginTop: 4 }}>{c.statoOp}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          {c.ultimoContatto ? <span style={{ fontSize: 11, color: T.textSub }}>{gg < 9999 ? gg + "gg fa" : "—"}</span> : <span style={{ fontSize: 11, color: "#e24b4a" }}>Mai contattato</span>}
          {c.dataRichiamo === oggi() && <span style={{ background: T.warnBg, border: "1px solid " + T.warnBorder, color: "#EF9F27", borderRadius: 8, padding: "1px 7px", fontSize: 11 }}>⏰ oggi</span>}
          {c.esito && <span style={{ fontSize: 11, color: T.textMuted, fontStyle: "italic" }}>"{c.esito.slice(0, 50)}{c.esito.length > 50 ? "…" : ""}"</span>}
        </div>
      </div>
    </div>
  );
}

/* --- CLIENTI --- */
function Clienti({ clienti, T, onSel, onImporta }) {
  const [search, setSearch] = useState("");
  const [fs, setFs] = useState("TUTTI");
  const [fa, setFa] = useState("TUTTI");
  const agenti = useMemo(() => ["TUTTI", ...new Set(clienti.map(c => c.agente).filter(Boolean))].sort(), [clienti]);
  const filtrati = useMemo(() => clienti.filter(c => {
    if (fs !== "TUTTI" && c.statoOp !== fs) return false;
    if (fa !== "TUTTI" && c.agente !== fa) return false;
    if (search && !c.ragione?.toLowerCase().includes(search.toLowerCase()) && !c.codice?.includes(search)) return false;
    return true;
  }).sort((a, b) => (a.agente || "").localeCompare(b.agente || "") || (b.totaleScaduto || 0) - (a.totaleScaduto || 0)), [clienti, search, fs, fa]);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div><div className="st">Clienti</div><div style={{ color: T.textSub, fontSize: 13 }}>{clienti.length} totali · {filtrati.length} mostrati</div></div>
        <label style={{ background: "#4a7fd4", color: "#fff", borderRadius: 10, padding: "10px 14px", fontSize: 14, fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>
          <i className="ti ti-upload" /> Importa
          <input type="file" accept=".xlsx,.xls,.xlsm" style={{ display: "none" }} onChange={e => e.target.files[0] && onImporta(e.target.files[0])} />
        </label>
      </div>
      <div className="sw"><i className="ti ti-search sic" /><input className="si" placeholder="Cerca nome o codice..." value={search} onChange={e => setSearch(e.target.value)} /></div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12, overflowX: "auto", paddingBottom: 4 }}>
        {["TUTTI", ...STATI_OP].map(s => <button key={s} onClick={() => setFs(s)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: fs === s ? "#4a7fd4" : "transparent", color: fs === s ? "#fff" : T.textSub, border: fs === s ? "none" : "1px solid " + T.borderInput }}>{s}</button>)}
      </div>
      <div className="fg"><select value={fa} onChange={e => setFa(e.target.value)}>{agenti.map(a => <option key={a}>{a}</option>)}</select></div>
      {!filtrati.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun cliente trovato</div> : filtrati.map(c => <CC key={c.id} c={c} T={T} onClick={() => onSel(c.id)} />)}
    </div>
  );
}

/* --- SCHEDA CLIENTE --- */
function SchedaCliente({ cliente: c, agenti, T, showToast, onUpdate, onAddDiario, onApriLL, onAddIncasso, onAddInsoluto }) {
  const [tab, setTab] = useState("info");
  const [showC, setShowC] = useState(false);
  const [showI, setShowI] = useState(false);
  const [showA, setShowA] = useState(false);
  const [showIns, setShowIns] = useState(false);
  const [showAna, setShowAna] = useState(false);
  const [ft, setFt] = useState("TUTTI");
  const sc = (c.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  const mailLink = generaMailSollecito(c);
  const df = useMemo(() => { const d = c.diario || []; return ft === "TUTTI" ? d : d.filter(v => v.tipo === ft); }, [c.diario, ft]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 700, marginBottom: 4 }}>{c.ragione}</div>
          <div style={{ fontSize: 12, color: T.textSub }}>{c.codice} · {c.agente}</div>
          <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
            <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888" }}>{c.statoOp}</span>
            <span className="tag" style={{ background: (PRIORITA_COLORS[c.priorita] || "#888") + "22", color: PRIORITA_COLORS[c.priorita] || "#888" }}>{c.priorita}</span>
          </div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div className="mono" style={{ fontSize: 24, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
          <div style={{ fontSize: 11, color: T.textSub, marginTop: 2 }}>scaduto</div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 16 }}>
        <button className="btn-p" onClick={() => setShowC(true)}><i className="ti ti-phone" /> Contatto</button>
        {mailLink ? <a href={mailLink} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: T.infoBg, color: "#378ADD", border: "1px solid " + T.infoBorder, borderRadius: 10, padding: 10, fontSize: 14, textDecoration: "none" }}><i className="ti ti-mail" /> Mail</a> : <button className="btn-g" disabled style={{ opacity: .4 }}>✉ Mail</button>}
        <button className="btn-g" onClick={() => setShowI(true)}><i className="ti ti-coin" /> Incasso</button>
        <button className="btn-d" onClick={() => setShowIns(true)}><i className="ti ti-alert-triangle" /> Insoluto</button>
      </div>

      {c.dataRichiamo && <div className="wb">⏰ Richiamo: {fmtData(c.dataRichiamo)}</div>}
      {c.esito && <div style={{ background: T.bgAlt, border: "1px solid " + T.border, borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: T.textSub }}><span style={{ fontSize: 11 }}>ULTIMO ESITO · </span>{c.esito}{c.ultimoContatto && <span style={{ fontSize: 11 }}> · {fmtData(c.ultimoContatto)}</span>}</div>}

      <div className="tabs">
        {[["info", "Info"], ["scadenze", "Scad. (" + sc.length + ")"], ["diario", "Diario (" + (c.diario || []).length + ")"]].map(([k, l]) => (
          <button key={k} className={"tab" + (tab === k ? " on" : "")} onClick={() => setTab(k)}>{l}</button>
        ))}
      </div>

      {tab === "info" && (
        <div>
          <div className="card" style={{ marginBottom: 12 }}>
            {[["Email", c.email || "—"], ["Telefono", c.telefono || "—"], ["Cellulare", c.cellulare || "—"], ["Orari", c.orari || "—"], ["Riferimento", c.personeRif || "—"], ["Agente", c.agente || "—"], ["Localita", (c.localita || "—") + " (" + (c.prov || "—") + ")"], ["Max ritardo", c.giorniMaxRitardo ? c.giorniMaxRitardo + "gg" : "—"], ["Note", c.note || "—"]].map(([k, v]) => (
              <div key={k} className="kv"><span style={{ color: T.textSub }}>{k}</span><span style={{ fontWeight: 500 }}>{v}</span></div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="btn-g" style={{ flex: 1 }} onClick={() => setShowAna(true)}><i className="ti ti-edit" /> Modifica</button>
            <button className="btn-d" style={{ flex: 1 }} onClick={onApriLL}><i className="ti ti-file-text" /> Apri L/L</button>
          </div>
        </div>
      )}

      {tab === "scadenze" && (
        <div>
          <button className="btn-g" style={{ width: "100%", marginBottom: 12 }} onClick={() => setShowA(true)}><i className="ti ti-arrows-join" /> Gestisci pagamento / accorpa</button>
          {!sc.length ? <div className="card" style={{ textAlign: "center", padding: 30, color: T.textMuted }}>Nessuna scadenza aperta</div> : sc.map((s, i) => (
            <div key={i} className="sr">
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                <span style={{ fontWeight: 500 }}>Ft. {s.doc}</span>
                <span className="mono" style={{ fontWeight: 600 }}>{fmtEur(s.residuo || s.importo)}</span>
              </div>
              <div style={{ fontSize: 12, color: T.textSub }}>Scad: {fmtData(s.scadenza)} · {s.giorni || 0}gg {s.residuo < s.importo && <span style={{ color: "#EF9F27" }}>(parziale)</span>}</div>
            </div>
          ))}
        </div>
      )}

      {tab === "diario" && (
        <div>
          <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto" }}>
            {["TUTTI", ...TIPI_DIARIO].map(t => (
              <button key={t} onClick={() => setFt(t)} style={{ padding: "6px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: ft === t ? "#4a7fd4" : "transparent", color: ft === t ? "#fff" : T.textSub, border: ft === t ? "none" : "1px solid " + T.borderInput }}>{t}</button>
            ))}
          </div>
          {!df.length ? <div className="card" style={{ textAlign: "center", padding: 30, color: T.textMuted }}>Nessuna voce</div> : df.map((v, i) => {
            const dc = T.diario[v.tipo] || T.diario["Nota personale"];
            return (
              <div key={i} style={{ borderLeft: "3px solid " + dc.border, padding: "10px 12px", background: dc.bg, borderRadius: "0 10px 10px 0", marginBottom: 8 }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 6, flexWrap: "wrap", alignItems: "center" }}>
                  <span style={{ fontSize: 11, color: T.textSub }}>{fmtData(v.data)} {v.ora}</span>
                  <span style={{ fontSize: 11, fontWeight: 500, color: dc.label, border: "1px solid " + dc.border, background: dc.bg, padding: "1px 8px", borderRadius: 20 }}>{v.tipo}</span>
                  {v.canale && <span className="tag" style={{ background: T.bgAlt, color: T.textSub }}>{v.canale}</span>}
                  {v.stato && <span className="tag" style={{ background: (STATO_COLORS[v.stato] || "#888") + "22", color: STATO_COLORS[v.stato] || "#888" }}>{v.stato}</span>}
                </div>
                <div style={{ fontSize: 13, lineHeight: 1.5, color: T.text }}>{v.testo || v.esito}</div>
              </div>
            );
          })}
        </div>
      )}

      {showC && <MContatto c={c} T={T} onClose={() => setShowC(false)} onSave={(v, u) => { onAddDiario(v); onUpdate(u); setShowC(false); showToast("Contatto registrato"); }} />}
      {showI && <MIncasso c={c} T={T} onClose={() => setShowI(false)} onSave={inc => { onAddIncasso(inc); setShowI(false); }} />}
      {showA && <MAccorpa c={c} T={T} onClose={() => setShowA(false)} onSave={u => { onUpdate(u); setShowA(false); showToast("Aggiornato"); }} />}
      {showIns && <MInsoluto c={c} agenti={agenti} T={T} onClose={() => setShowIns(false)} onSave={ins => { onAddInsoluto(ins); setShowIns(false); }} />}
      {showAna && <MAnagrafica c={c} T={T} onClose={() => setShowAna(false)} onSave={u => { onUpdate(u); setShowAna(false); showToast("Anagrafica aggiornata"); }} />}
    </div>
  );
}

/* --- MODALI --- */
function MContatto({ c, T, onClose, onSave }) {
  const [data, setData] = useState(oggi());
  const [ora, setOra] = useState(oraOra());
  const [tipo, setTipo] = useState("Mio contatto");
  const [canale, setCanale] = useState("Telefono");
  const [testo, setTesto] = useState("");
  const [stato, setStato] = useState(c.statoOp);
  const [priorita, setPriorita] = useState(c.priorita);
  const [richiamo, setRichiamo] = useState(c.dataRichiamo || "");

  function save(e) {
    e.preventDefault();
    const v = { data, ora, tipo, canale: tipo !== "Nota personale" ? canale : null, testo, stato, ts: new Date().toISOString() };
    onSave(v, { statoOp: stato, esito: testo, ultimoContatto: tipo !== "Nota personale" ? data : c.ultimoContatto, dataRichiamo: richiamo || null, priorita });
  }
  const dc = T.diario[tipo] || T.diario["Nota personale"];

  return (
    <div className="mbg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="mo">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Registra contatto</div>
          <button className="btn-g" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={save}>
          <div className="fg">
            <label>TIPO</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {TIPI_DIARIO.map(t => {
                const tc = T.diario[t];
                return <button key={t} type="button" onClick={() => setTipo(t)} style={{ padding: "7px 14px", fontSize: 13, borderRadius: 20, background: tipo === t ? tc.bg : "transparent", color: tipo === t ? tc.label : T.textSub, border: "1px solid " + (tipo === t ? tc.border : T.borderInput), fontWeight: tipo === t ? 500 : 400 }}>{t}</button>;
              })}
            </div>
          </div>
          <div className="r2">
            <div className="fg"><label>DATA</label><input type="date" value={data} onChange={e => setData(e.target.value)} /></div>
            <div className="fg"><label>ORA</label><input type="time" value={ora} onChange={e => setOra(e.target.value)} /></div>
          </div>
          {tipo !== "Nota personale" && <div className="fg"><label>CANALE</label><select value={canale} onChange={e => setCanale(e.target.value)}>{TIPI_CONTATTO.map(t => <option key={t}>{t}</option>)}</select></div>}
          <div className="fg"><label>NOTE / ESITO</label><textarea value={testo} onChange={e => setTesto(e.target.value)} placeholder="Scrivi liberamente..." /></div>
          <div className="r2">
            <div className="fg"><label>STATO</label><select value={stato} onChange={e => setStato(e.target.value)}>{STATI_OP.map(s => <option key={s}>{s}</option>)}</select></div>
            <div className="fg"><label>PRIORITA</label><select value={priorita} onChange={e => setPriorita(e.target.value)}>{PRIORITA.map(p => <option key={p}>{p}</option>)}</select></div>
          </div>
          <div className="fg"><label>DATA RICHIAMO — vuoto = nessuno</label><input type="date" value={richiamo} onChange={e => setRichiamo(e.target.value)} /></div>
          <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }}>Salva</button>
        </form>
      </div>
    </div>
  );
}

function MIncasso({ c, T, onClose, onSave }) {
  const [tipo, setTipo] = useState("Assegno a vista");
  const [importo, setImporto] = useState("");
  const [spese, setSpese] = useState("0");
  const [note, setNote] = useState("");
  const [fatture, setFatture] = useState("");
  return (
    <div className="mbg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="mo">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Incasso — {c.ragione}</div>
          <button className="btn-g" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={e => { e.preventDefault(); onSave({ clienteId: c.id, ragione: c.ragione, agente: c.agente, tipo, importo: parseFloat(importo), spese: parseFloat(spese || 0), incassoNetto: parseFloat(importo) - parseFloat(spese || 0), note, fatture }); onClose(); }}>
          <div className="fg"><label>TIPO</label><select value={tipo} onChange={e => setTipo(e.target.value)}>{TIPI_TITOLO.map(t => <option key={t}>{t}</option>)}</select></div>
          {tipo === "Bonifico" && <div className="ib">Il bonifico viene registrato direttamente come pagato</div>}
          <div className="r2">
            <div className="fg"><label>IMPORTO (EUR)</label><input type="number" step="0.01" value={importo} onChange={e => setImporto(e.target.value)} required /></div>
            <div className="fg"><label>SPESE (EUR)</label><input type="number" step="0.01" value={spese} onChange={e => setSpese(e.target.value)} /></div>
          </div>
          {importo && <div style={{ background: T.successBg, border: "1px solid " + T.successBorder, borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#4ecb8d" }}>Netto: {fmtEur(parseFloat(importo || 0) - parseFloat(spese || 0))}</div>}
          <div className="fg"><label>FATTURE</label><input value={fatture} onChange={e => setFatture(e.target.value)} placeholder="es. Ft.001 + Ft.002" /></div>
          <div className="fg"><label>NOTE</label><textarea value={note} onChange={e => setNote(e.target.value)} style={{ minHeight: 60 }} /></div>
          <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }}>Registra</button>
        </form>
      </div>
    </div>
  );
}

function MAccorpa({ c, T, onClose, onSave }) {
  const sc = (c.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  const [sel, setSel] = useState([]);
  const [tipoOp, setTipoOp] = useState("pagata");
  const [ip, setIp] = useState("");
  const [fa, setFa] = useState("");
  function toggle(tag) { setSel(s => s.includes(tag) ? s.filter(t => t !== tag) : [...s, tag]); }
  return (
    <div className="mbg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="mo">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Gestisci pagamento</div>
          <button className="btn-g" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={e => {
          e.preventDefault();
          const ns = c.scadenze.map(s => {
            if (!sel.includes(s.tag)) return s;
            if (tipoOp === "pagata") return { ...s, stato: "PAGATA", residuo: 0 };
            if (tipoOp === "parziale") return { ...s, residuo: Math.max(0, s.importo - parseFloat(ip || 0)) };
            if (tipoOp === "accorpa") return { ...s, accorpata: true, fatturaAccorpante: fa };
            return s;
          });
          const nt = ns.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
          onSave({ scadenze: ns, totaleScaduto: nt });
        }}>
          <div className="fg"><label>OPERAZIONE</label>
            <select value={tipoOp} onChange={e => setTipoOp(e.target.value)}>
              <option value="pagata">Pagata interamente</option>
              <option value="parziale">Pagamento parziale</option>
              <option value="accorpa">Accorpa sotto altra fattura</option>
            </select>
          </div>
          {tipoOp === "parziale" && <div className="fg"><label>IMPORTO PAGATO (EUR)</label><input type="number" step="0.01" value={ip} onChange={e => setIp(e.target.value)} /></div>}
          {tipoOp === "accorpa" && <div className="fg"><label>FATTURA ACCORPANTE</label><input value={fa} onChange={e => setFa(e.target.value)} placeholder="es. 12345" /></div>}
          <div className="fg">
            <label>SCADENZE ({sel.length} sel.)</label>
            {sc.map((s, i) => (
              <div key={i} onClick={() => toggle(s.tag)} style={{ display: "flex", gap: 10, alignItems: "center", padding: "10px 12px", background: sel.includes(s.tag) ? T.successBg : T.bgAlt, border: "1px solid " + (sel.includes(s.tag) ? T.successBorder : T.border), borderRadius: 10, marginBottom: 6, cursor: "pointer" }}>
                <div style={{ width: 20, height: 20, borderRadius: 4, border: "2px solid " + (sel.includes(s.tag) ? "#4ecb8d" : T.borderInput), display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  {sel.includes(s.tag) && <span style={{ color: "#4ecb8d", fontSize: 12 }}>✓</span>}
                </div>
                <div><div style={{ fontSize: 13, fontWeight: 500, color: T.text }}>Ft. {s.doc} — {fmtEur(s.residuo || s.importo)}</div><div style={{ fontSize: 11, color: T.textSub }}>Scad: {fmtData(s.scadenza)}</div></div>
              </div>
            ))}
          </div>
          <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }} disabled={!sel.length}>Applica</button>
        </form>
      </div>
    </div>
  );
}

function MInsoluto({ c, agenti, T, onClose, onSave }) {
  const [importo, setImporto] = useState("");
  const [rif, setRif] = useState("");
  const [dataE, setDataE] = useState("");
  const [motivo, setMotivo] = useState("Mancanza fondi");
  const ag = agenti.find(a => a.nome === c.agente);
  const ins = { riferimento: rif, importo: parseFloat(importo || 0), dataEmissione: dataE, motivo };
  return (
    <div className="mbg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="mo">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600, color: "#e24b4a" }}>⚠ Assegno insoluto</div>
          <button className="btn-g" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={e => { e.preventDefault(); onSave({ clienteId: c.id, ragione: c.ragione, agente: c.agente, ...ins }); onClose(); }}>
          <div className="fg"><label>IMPORTO (EUR)</label><input type="number" step="0.01" value={importo} onChange={e => setImporto(e.target.value)} required /></div>
          <div className="r2">
            <div className="fg"><label>N. ASSEGNO</label><input value={rif} onChange={e => setRif(e.target.value)} /></div>
            <div className="fg"><label>DATA EMISSIONE</label><input type="date" value={dataE} onChange={e => setDataE(e.target.value)} /></div>
          </div>
          <div className="fg"><label>MOTIVO</label>
            <select value={motivo} onChange={e => setMotivo(e.target.value)}>
              {["Mancanza fondi", "Impagato", "Conto chiuso", "Firma non corrispondente", "Altro"].map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
          <div className="db">Il cliente passera in stato PROBLEMA · 60gg per regolarizzare</div>
          <button type="submit" className="btn-d" style={{ width: "100%", padding: 14, marginBottom: 10 }}>Registra insoluto</button>
          <div style={{ display: "flex", gap: 8 }}>
            {generaMailInsolutoCliente(c, ins) && <a href={generaMailInsolutoCliente(c, ins)} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: T.infoBg, color: "#378ADD", border: "1px solid " + T.infoBorder, borderRadius: 10, padding: 10, fontSize: 13, textDecoration: "none" }}>✉ Cliente</a>}
            {ag && generaMailInsolutoAgente(ag, c, ins) && <a href={generaMailInsolutoAgente(ag, c, ins)} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: T.successBg, color: "#4ecb8d", border: "1px solid " + T.successBorder, borderRadius: 10, padding: 10, fontSize: 13, textDecoration: "none" }}>✉ Agente</a>}
          </div>
        </form>
      </div>
    </div>
  );
}

function MAnagrafica({ c, T, onClose, onSave }) {
  const [f, setF] = useState({ email: c.email || "", telefono: c.telefono || "", cellulare: c.cellulare || "", orari: c.orari || "", personeRif: c.personeRif || "", note: c.note || "" });
  return (
    <div className="mbg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="mo">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Modifica anagrafica</div>
          <button className="btn-g" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={e => { e.preventDefault(); onSave(f); }}>
          <div className="r2">
            <div className="fg"><label>EMAIL</label><input type="email" value={f.email} onChange={e => setF(x => ({ ...x, email: e.target.value }))} /></div>
            <div className="fg"><label>TELEFONO</label><input value={f.telefono} onChange={e => setF(x => ({ ...x, telefono: e.target.value }))} /></div>
          </div>
          <div className="fg"><label>CELLULARE</label><input value={f.cellulare} onChange={e => setF(x => ({ ...x, cellulare: e.target.value }))} /></div>
          <div className="fg"><label>ORARI REPERIBILITA</label><input value={f.orari} onChange={e => setF(x => ({ ...x, orari: e.target.value }))} placeholder="es. Lun-Ven 9-12" /></div>
          <div className="fg"><label>PERSONA DI RIFERIMENTO</label><input value={f.personeRif} onChange={e => setF(x => ({ ...x, personeRif: e.target.value }))} /></div>
          <div className="fg"><label>NOTE FISSE</label><textarea value={f.note} onChange={e => setF(x => ({ ...x, note: e.target.value }))} style={{ minHeight: 80 }} /></div>
          <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }}>Salva</button>
        </form>
      </div>
    </div>
  );
}

/* --- AGENTI --- */
function Agenti({ agenti, clienti, incassi, T, onAdd, onSel }) {
  const [showF, setShowF] = useState(false);
  const [f, setF] = useState({ nome: "", email: "", telefono: "", zona: "" });
  const nomi = useMemo(() => [...new Set([...agenti.map(a => a.nome), ...clienti.map(c => c.agente).filter(Boolean)])].sort(), [agenti, clienti]);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div><div className="st">Agenti</div><div style={{ color: T.textSub, fontSize: 13 }}>{nomi.length} agenti</div></div>
        <button className="btn-p" onClick={() => setShowF(s => !s)}>+ Aggiungi</button>
      </div>
      {showF && (
        <div className="card" style={{ marginBottom: 16 }}>
          <form onSubmit={e => { e.preventDefault(); onAdd(f); setF({ nome: "", email: "", telefono: "", zona: "" }); setShowF(false); }}>
            <div className="r2">
              <div className="fg"><label>NOME COMPLETO</label><input value={f.nome} onChange={e => setF(x => ({ ...x, nome: e.target.value }))} required /></div>
              <div className="fg"><label>ZONA</label><input value={f.zona} onChange={e => setF(x => ({ ...x, zona: e.target.value }))} /></div>
            </div>
            <div className="r2">
              <div className="fg"><label>EMAIL</label><input type="email" value={f.email} onChange={e => setF(x => ({ ...x, email: e.target.value }))} /></div>
              <div className="fg"><label>TELEFONO</label><input value={f.telefono} onChange={e => setF(x => ({ ...x, telefono: e.target.value }))} /></div>
            </div>
            <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }}>Salva</button>
          </form>
        </div>
      )}
      {nomi.map(nome => {
        const ag = agenti.find(a => a.nome === nome);
        const cl = clienti.filter(c => c.agente === nome && c.statoOp !== "PAGATO");
        const inc = incassi.filter(i => i.agente === nome && i.stato === "IN MANO AGENTE");
        const tot = cl.reduce((s, c) => s + (c.totaleScaduto || 0), 0);
        const tm = inc.reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
        const alert = inc.some(i => giorniDa(i.dataRicezione) > 15);
        return (
          <div key={nome} className="cc" onClick={() => onSel(ag?.id || nome)} style={{ borderColor: alert ? T.warnBorder : T.border }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 15 }}>{nome}</div>
                <div style={{ fontSize: 11, color: T.textSub }}>{ag?.zona || "—"}{!ag && <span style={{ color: "#EF9F27", marginLeft: 8 }}>⚠ anagrafica mancante</span>}</div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(tot)}</div>
                <div style={{ fontSize: 11, color: T.textSub }}>{cl.length} clienti</div>
              </div>
            </div>
            {tm > 0 && <div style={{ fontSize: 12, color: alert ? "#EF9F27" : "#378ADD" }}>💼 {fmtEur(tm)} in mano</div>}
          </div>
        );
      })}
    </div>
  );
}

/* --- SCHEDA AGENTE --- */
function SchedaAgente({ agente, clienti, incassi, T, showToast, onUpdate, onAddDiario, onUpdCliente, onUpdIncasso, onDelIncasso, onSelectCliente }) {
  const [tab, setTab] = useState("clienti");
  const [sort, setSort] = useState("AGENTE");
  const [tracking, setTracking] = useState({});
  const [showEdit, setShowEdit] = useState(false);
  const [pending, setPending] = useState({});

  // Freeze order during call - sort only on mount
  const clientiOrd = useMemo(() => [...clienti].sort((a, b) => {
    if (sort !== "TUTTI") { if (a.statoOp === sort && b.statoOp !== sort) return -1; if (b.statoOp === sort && a.statoOp !== sort) return 1; }
    return (b.totaleScaduto || 0) - (a.totaleScaduto || 0);
  }), [clienti, sort]);

  const totS = clienti.reduce((s, c) => s + (c.totaleScaduto || 0), 0);
  const totM = incassi.reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);

  function annota(cid, testo, stato, richiamo) {
    const v = { data: oggi(), ora: oraOra(), tipo: "Via agente", testo, stato: stato || null, ts: new Date().toISOString() };
    onAddDiario(cid, v);
    const u = { ultimoContatto: oggi() };
    if (stato) u.statoOp = stato;
    if (richiamo) u.dataRichiamo = richiamo;
    setPending(p => ({ ...p, [cid]: { ...(p[cid] || {}), ...u } }));
    onUpdCliente(cid, u);
    showToast("Annotato");
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 700, marginBottom: 4 }}>{agente.nome}</div>
          <div style={{ fontSize: 12, color: T.textSub }}>{agente.zona || "—"} · {agente.email || "—"}</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div className="mono" style={{ fontSize: 20, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(totS)}</div>
          {totM > 0 && <div style={{ fontSize: 12, color: "#EF9F27", marginTop: 2 }}>💼 {fmtEur(totM)} in mano</div>}
        </div>
      </div>

      <div className="tabs">
        <button className={"tab" + (tab === "clienti" ? " on" : "")} onClick={() => setTab("clienti")}>Clienti ({clienti.length})</button>
        <button className={"tab" + (tab === "titoli" ? " on" : "")} onClick={() => setTab("titoli")}>Titoli ({incassi.length})</button>
        <button className={"tab" + (tab === "info" ? " on" : "")} onClick={() => setTab("info")}>Anagrafica</button>
      </div>

      {tab === "clienti" && (
        <div>
          <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto" }}>
            {["TUTTI", "AGENTE", "PROBLEMA", "CONTROLLA", "L/L", "AGENZIA", "RICHIAMA"].map(s => (
              <button key={s} onClick={() => setSort(s)} style={{ padding: "6px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: sort === s ? "#4a7fd4" : "transparent", color: sort === s ? "#fff" : T.textSub, border: sort === s ? "none" : "1px solid " + T.borderInput }}>{s}</button>
            ))}
          </div>
          {clientiOrd.map(c => (
            <CACard key={c.id} c={{ ...c, ...(pending[c.id] || {}) }} T={T} onApri={() => onSelectCliente(c.id)} onAnnota={(t, st, rc) => annota(c.id, t, st, rc)} />
          ))}
          {!clienti.length && <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun cliente</div>}
        </div>
      )}

      {tab === "titoli" && (
        <div>
          {!incassi.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun titolo in mano</div> : incassi.map(i => {
            const gg = giorniDa(i.dataRicezione);
            return (
              <div key={i.id} className="card" style={{ marginBottom: 10, borderColor: gg > 30 ? "#e24b4a44" : gg > 15 ? "#EF9F2744" : T.border }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <div>
                    <div style={{ fontWeight: 600, cursor: "pointer", color: "#4a7fd4" }} onClick={() => onSelectCliente(i.clienteId)}>{i.ragione}</div>
                    <div style={{ fontSize: 11, color: T.textSub }}>{i.tipo} · {fmtData(i.dataRicezione)} ({gg}gg)</div>
                  </div>
                  <div className="mono" style={{ fontWeight: 600 }}>{fmtEur(i.incassoNetto || i.importo)}</div>
                </div>
                {i.note && <div style={{ fontSize: 12, color: T.textSub, marginBottom: 8 }}>{i.note}</div>}
                {gg > 30 && <div style={{ fontSize: 12, color: "#e24b4a", marginBottom: 8, fontWeight: 500 }}>⚠ {gg}gg — sollecita!</div>}
                {gg > 15 && gg <= 30 && <div style={{ fontSize: 12, color: "#EF9F27", marginBottom: 8 }}>⏳ {gg}gg in mano</div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <input value={tracking[i.id] || ""} onChange={e => setTracking(t => ({ ...t, [i.id]: e.target.value }))} placeholder="Tracking..." style={{ flex: 1, fontSize: 13, padding: "8px 10px" }} />
                  <button className="btn-ok" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => { onUpdIncasso(i.id, { stato: "SPEDITO", dataSpedizione: oggi(), tracking: tracking[i.id] || "" }); showToast("Spedito"); }}>Spedito</button>
                  <button className="btn-d" style={{ fontSize: 13, padding: "8px 10px" }} onClick={() => { onDelIncasso(i.id); showToast("Eliminato"); }}>✕</button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {tab === "info" && (
        <div>
          <div className="card" style={{ marginBottom: 12 }}>
            {[["Nome", agente.nome], ["Email", agente.email || "—"], ["Telefono", agente.telefono || "—"], ["Zona", agente.zona || "—"]].map(([k, v]) => (
              <div key={k} className="kv"><span style={{ color: T.textSub }}>{k}</span><span style={{ fontWeight: 500 }}>{v}</span></div>
            ))}
          </div>
          <button className="btn-g" style={{ width: "100%" }} onClick={() => setShowEdit(true)}><i className="ti ti-edit" /> Modifica</button>
        </div>
      )}

      {showEdit && (
        <div className="mbg" onClick={e => e.target === e.currentTarget && setShowEdit(false)}>
          <div className="mo">
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>Modifica agente</div>
              <button className="btn-g" style={{ padding: "6px 12px" }} onClick={() => setShowEdit(false)}>✕</button>
            </div>
            <MEditAgente ag={agente} T={T} onSave={u => { onUpdate(u); setShowEdit(false); showToast("Salvato"); }} />
          </div>
        </div>
      )}
    </div>
  );
}

function MEditAgente({ ag, T, onSave }) {
  const [f, setF] = useState({ nome: ag.nome || "", email: ag.email || "", telefono: ag.telefono || "", zona: ag.zona || "" });
  return (
    <form onSubmit={e => { e.preventDefault(); onSave(f); }}>
      <div className="fg"><label>NOME</label><input value={f.nome} onChange={e => setF(x => ({ ...x, nome: e.target.value }))} /></div>
      <div className="r2">
        <div className="fg"><label>EMAIL</label><input type="email" value={f.email} onChange={e => setF(x => ({ ...x, email: e.target.value }))} /></div>
        <div className="fg"><label>TELEFONO</label><input value={f.telefono} onChange={e => setF(x => ({ ...x, telefono: e.target.value }))} /></div>
      </div>
      <div className="fg"><label>ZONA</label><input value={f.zona} onChange={e => setF(x => ({ ...x, zona: e.target.value }))} /></div>
      <button type="submit" className="btn-p" style={{ width: "100%", padding: 14 }}>Salva</button>
    </form>
  );
}

function CACard({ c, T, onApri, onAnnota }) {
  const [show, setShow] = useState(false);
  const [testo, setTesto] = useState("");
  const [stato, setStato] = useState("");
  const [richiamo, setRichiamo] = useState("");
  const ultima = (c.diario || []).find(v => v.tipo === "Via agente");
  const dc = T.diario["Via agente"];
  return (
    <div style={{ background: T.bgCard, border: "1px solid " + (c.statoOp === "AGENTE" ? T.warnBorder : T.border), borderRadius: 14, padding: 14, marginBottom: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
        <div>
          <div style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4a7fd4" }} onClick={onApri}>{c.ragione}</div>
          <div style={{ fontSize: 11, color: T.textSub }}>{c.codice}</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
          <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888" }}>{c.statoOp}</span>
        </div>
      </div>
      {ultima && (
        <div style={{ fontSize: 12, color: dc.label, borderLeft: "2px solid " + dc.border, background: dc.bg, padding: "4px 8px", borderRadius: "0 6px 6px 0", marginBottom: 8 }}>
          {fmtData(ultima.data)}: {(ultima.testo || "").slice(0, 70)}{(ultima.testo || "").length > 70 ? "…" : ""}
        </div>
      )}
      {!show ? (
        <button className="btn-g" style={{ width: "100%", fontSize: 13, padding: "8px" }} onClick={() => setShow(true)}>+ Annota</button>
      ) : (
        <div>
          <textarea value={testo} onChange={e => setTesto(e.target.value)} placeholder="Cosa ti ha detto l'agente..." style={{ width: "100%", background: dc.bg, border: "1px solid " + dc.border, borderRadius: 10, color: T.text, padding: "10px 14px", fontSize: 14, marginBottom: 8, resize: "vertical", minHeight: 60, outline: "none" }} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
            <select value={stato} onChange={e => setStato(e.target.value)} style={{ fontSize: 13, padding: "8px 10px" }}>
              <option value="">Stato invariato</option>
              {STATI_OP.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            <input type="date" value={richiamo} onChange={e => setRichiamo(e.target.value)} style={{ fontSize: 13, padding: "8px 10px" }} />
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="btn-ok" style={{ flex: 1, fontSize: 13, padding: "8px" }} onClick={() => { if (testo) { onAnnota(testo, stato || null, richiamo || null); setTesto(""); setStato(""); setRichiamo(""); setShow(false); } }}>Salva</button>
            <button className="btn-g" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => setShow(false)}>✕</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* --- L/L --- */
function PraticheLl({ pratiche, clienti, T, onUpd, onSelCl, showToast }) {
  const [filtro, setFiltro] = useState("ATTIVE");
  const ff = pratiche.filter(p => filtro === "TUTTE" ? true : filtro === "ATTIVE" ? p.stato !== "AGENZIA" && p.stato !== "CHIUSA" : p.stato === filtro);

  function avanza(p) {
    const n = p.lettera || 0;
    if (n >= 3) { onUpd(p.id, { stato: "AGENZIA" }); showToast("Passato ad Agenzia"); return; }
    onUpd(p.id, { lettera: n + 1, stato: "LETTERA " + (n + 1), dataUltimaLettera: oggi(), esito: "" });
    showToast("Lettera " + (n + 1) + " registrata");
  }

  return (
    <div>
      <div className="st">Pratiche L/L</div>
      <div style={{ color: T.textSub, fontSize: 13, marginBottom: 16 }}>{pratiche.filter(p => p.stato !== "AGENZIA" && p.stato !== "CHIUSA").length} aperte</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto" }}>
        {["ATTIVE", "IN CODA", "LETTERA 1", "LETTERA 2", "LETTERA 3", "AGENZIA", "TUTTE"].map(f => (
          <button key={f} onClick={() => setFiltro(f)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: filtro === f ? "#4a7fd4" : "transparent", color: filtro === f ? "#fff" : T.textSub, border: filtro === f ? "none" : "1px solid " + T.borderInput }}>{f}</button>
        ))}
      </div>
      {!ff.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessuna pratica</div> : ff.map(p => {
        const gg = p.dataUltimaLettera ? giorniDa(p.dataUltimaLettera) : null;
        const pronta = gg === null || gg >= 20;
        return (
          <div key={p.id} className="card" style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <div>
                <div style={{ fontWeight: 600, cursor: "pointer", color: "#4a7fd4", marginBottom: 2 }} onClick={() => { const cl = clienti.find(c => c.id === p.clienteId); if (cl) onSelCl(cl.id); }}>{p.ragione}</div>
                <div style={{ fontSize: 11, color: T.textSub }}>{p.agente} · {fmtData(p.dataDecisione)}</div>
              </div>
              <span style={{ background: "#e24b4a22", color: "#e24b4a", fontSize: 12, padding: "4px 10px", borderRadius: 8 }}>{p.stato}</span>
            </div>
            <div style={{ display: "flex", gap: 4, marginBottom: 10 }}>
              {["C", "L1", "L2", "L3", "AG"].map((s, i) => <div key={s} style={{ flex: 1, height: 4, borderRadius: 2, background: i <= (p.lettera || 0) ? "#e24b4a" : T.border }} />)}
            </div>
            {gg !== null && <div style={{ fontSize: 12, color: gg >= 20 ? "#4ecb8d" : "#EF9F27", marginBottom: 8 }}>{gg >= 20 ? "✓ Pronta (" + gg + "gg)" : "⏳ " + (20 - gg) + "gg rimanenti"}</div>}
            {p.esito && <div style={{ fontSize: 12, color: T.textSub, marginBottom: 10 }}>Esito: {p.esito}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              {p.stato !== "AGENZIA" && <select onChange={e => e.target.value && onUpd(p.id, { esito: e.target.value, dataEsito: oggi() })} defaultValue="" style={{ flex: 1, fontSize: 13, padding: "8px 10px" }}>
                <option value="">Esito raccomandata...</option>
                {ESITI_RACC.map(e => <option key={e} value={e}>{e}</option>)}
              </select>}
              {p.stato !== "AGENZIA" && pronta && <button className="btn-d" style={{ fontSize: 13, padding: "8px 14px", whiteSpace: "nowrap" }} onClick={() => avanza(p)}>{p.lettera >= 3 ? "Agenzia" : "L" + ((p.lettera || 0) + 1)}</button>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* --- INCASSI --- */
function Incassi({ incassi, clienti, T, onAdd, onUpd, onDel, onSelCl, showToast }) {
  const [tab, setTab] = useState("mano");
  const inMano = incassi.filter(i => i.stato === "IN MANO AGENTE");
  const spediti = incassi.filter(i => i.stato === "SPEDITO");
  const perAg = useMemo(() => {
    const map = {};
    inMano.forEach(i => { if (!map[i.agente]) map[i.agente] = { ag: i.agente, tot: 0, items: [] }; map[i.agente].tot += i.incassoNetto || i.importo || 0; map[i.agente].items.push(i); });
    return Object.values(map).sort((a, b) => b.tot - a.tot);
  }, [inMano]);

  const IC = ({ inc }) => {
    const [tr, setTr] = useState("");
    const gg = giorniDa(inc.dataRicezione);
    return (
      <div className="card" style={{ marginBottom: 10, borderColor: gg > 30 ? "#e24b4a44" : gg > 15 ? "#EF9F2744" : T.border }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
          <div>
            <div style={{ fontWeight: 600, cursor: "pointer", color: "#4a7fd4" }} onClick={() => { const cl = clienti.find(c => c.id === inc.clienteId); if (cl) onSelCl(cl.id); }}>{inc.ragione}</div>
            <div style={{ fontSize: 11, color: T.textSub }}>{inc.agente} · {inc.tipo} · {fmtData(inc.dataRicezione)} ({gg}gg)</div>
          </div>
          <div className="mono" style={{ fontWeight: 600 }}>{fmtEur(inc.incassoNetto || inc.importo)}</div>
        </div>
        {inc.note && <div style={{ fontSize: 12, color: T.textSub, marginBottom: 8 }}>{inc.note}</div>}
        {gg > 30 && <div style={{ fontSize: 12, color: "#e24b4a", marginBottom: 8, fontWeight: 500 }}>⚠ Sollecita spedizione!</div>}
        {gg > 15 && gg <= 30 && <div style={{ fontSize: 12, color: "#EF9F27", marginBottom: 8 }}>⏳ {gg}gg in mano</div>}
        {inc.stato === "IN MANO AGENTE" && <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <input value={tr} onChange={e => setTr(e.target.value)} placeholder="Tracking..." style={{ flex: 1, fontSize: 13, padding: "8px 10px" }} />
          <button className="btn-ok" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => { onUpd(inc.id, { stato: "SPEDITO", dataSpedizione: oggi(), tracking: tr }); showToast("Spedito"); }}>Spedito</button>
        </div>}
        {inc.stato === "SPEDITO" && <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontSize: 12, color: "#4ecb8d" }}>✓ Spedito {fmtData(inc.dataSpedizione)}{inc.tracking ? " · " + inc.tracking : ""}</div>
          <button className="btn-g" style={{ fontSize: 12, padding: "6px 12px" }} onClick={() => { onUpd(inc.id, { stato: "REGISTRATO", dataRegistrazione: oggi() }); showToast("Registrato"); }}>Registrato</button>
        </div>}
        {inc.stato === "REGISTRATO" && <div style={{ fontSize: 12, color: "#1D9E75", marginBottom: 8 }}>✓ Registrato {fmtData(inc.dataRegistrazione)}</div>}
        <button onClick={() => { onDel(inc.id); }} style={{ width: "100%", background: "transparent", color: "#e24b4a", border: "1px solid " + T.dangerBorder, borderRadius: 8, padding: 5, fontSize: 12, cursor: "pointer" }}>✕ Elimina</button>
      </div>
    );
  };

  return (
    <div>
      <div className="st">Incassi agenti</div>
      <div style={{ color: T.textSub, fontSize: 13, marginBottom: 16 }}>{inMano.length} in mano · {spediti.length} spediti</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto" }}>
        {[["mano", "In mano (" + inMano.length + ")"], ["agente", "Per agente"], ["spediti", "Spediti (" + spediti.length + ")"], ["tutti", "Tutti"]].map(([k, l]) => (
          <button key={k} onClick={() => setTab(k)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: tab === k ? "#4a7fd4" : "transparent", color: tab === k ? "#fff" : T.textSub, border: tab === k ? "none" : "1px solid " + T.borderInput }}>{l}</button>
        ))}
      </div>
      {tab === "mano" && (!inMano.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun titolo</div> : inMano.map(i => <IC key={i.id} inc={i} />))}
      {tab === "agente" && (!perAg.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun titolo</div> : perAg.map(ag => (
        <div key={ag.ag} className="card" style={{ marginBottom: 12 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}><div style={{ fontWeight: 600 }}>{ag.ag}</div><div className="mono" style={{ fontWeight: 600, color: "#EF9F27" }}>{fmtEur(ag.tot)}</div></div>
          {ag.items.map(i => <div key={i.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: "1px solid " + T.border, fontSize: 13 }}><span style={{ color: T.text }}>{i.ragione} · {i.tipo}</span><span className="mono">{fmtEur(i.incassoNetto || i.importo)}</span></div>)}
        </div>
      )))}
      {tab === "spediti" && (!spediti.length ? <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Nessun titolo</div> : spediti.map(i => <IC key={i.id} inc={i} />))}
      {tab === "tutti" && incassi.map(i => <IC key={i.id} inc={i} />)}
    </div>
  );
}

/* --- DASHBOARD --- */
function Dashboard({ kpi, clienti, agenti, insoluti, T, onSelAg }) {
  const byS = useMemo(() => { const m = {}; STATI_OP.forEach(s => { m[s] = clienti.filter(c => c.statoOp === s).length; }); return m; }, [clienti]);
  const agKpi = useMemo(() => [...new Set(clienti.map(c => c.agente).filter(Boolean))].map(nome => {
    const ag = agenti.find(a => a.nome === nome);
    const cl = clienti.filter(c => c.agente === nome && c.statoOp !== "PAGATO");
    return { nome, id: ag?.id || nome, tot: cl.reduce((s, c) => s + (c.totaleScaduto || 0), 0), n: cl.length };
  }).sort((a, b) => b.tot - a.tot), [clienti, agenti]);
  const insA = insoluti.filter(i => i.stato === "APERTO");

  return (
    <div>
      <div className="st">Dashboard</div>
      <div style={{ color: T.textSub, fontSize: 13, marginBottom: 20 }}>{new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long" })}{kpi.di && <span style={{ marginLeft: 12, fontSize: 11 }}>· import {fmtData(kpi.di)}</span>}</div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
        <div className="card" style={{ gridColumn: "1 / -1" }}>
          <div style={{ fontSize: 11, color: T.textSub, marginBottom: 6 }}>SCADUTO TOTALE</div>
          <div className="mono" style={{ fontSize: 28, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(kpi.totS)}</div>
          {kpi.di && <div style={{ fontSize: 12, color: T.textSub, marginTop: 4 }}>Ultimo import: {fmtEur(kpi.si)}</div>}
        </div>
        {[["L/L", kpi.ll, "#e24b4a"], ["Titoli mano", kpi.tm, "#EF9F27"], ["Insoluti", kpi.ins, "#FF4A8D"], ["Richiami oggi", kpi.rc, "#378ADD"]].map(([l, v, c]) => (
          <div key={l} className="card">
            <div style={{ fontSize: 11, color: T.textSub, marginBottom: 6 }}>{l.toUpperCase()}</div>
            <div className="mono" style={{ fontSize: 22, fontWeight: 700, color: c }}>{v}</div>
          </div>
        ))}
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: T.textSub }}>Incassato</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
          <div style={{ background: T.successBg, border: "1px solid " + T.successBorder, borderRadius: 12, padding: 12 }}>
            <div style={{ fontSize: 10, color: "#4ecb8d", marginBottom: 4 }}>CONFERMATO</div>
            <div className="mono" style={{ fontSize: 15, fontWeight: 700, color: "#4ecb8d" }}>{fmtEur(kpi.confT)}</div>
            <div style={{ fontSize: 11, color: "#4ecb8d88", marginTop: 2 }}>{kpi.pctC.toFixed(1)}%</div>
          </div>
          <div style={{ background: T.infoBg, border: "1px solid " + T.infoBorder, borderRadius: 12, padding: 12 }}>
            <div style={{ fontSize: 10, color: "#378ADD", marginBottom: 4 }}>IN CORSO</div>
            <div className="mono" style={{ fontSize: 15, fontWeight: 700, color: "#378ADD" }}>{fmtEur(kpi.attT)}</div>
            <div style={{ fontSize: 11, color: "#378ADD88", marginTop: 2 }}>mano+spedito</div>
          </div>
          <div style={{ background: T.bgAlt, border: "1px solid " + T.border, borderRadius: 12, padding: 12 }}>
            <div style={{ fontSize: 10, color: T.textSub, marginBottom: 4 }}>TOTALE</div>
            <div className="mono" style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{fmtEur(kpi.tot)}</div>
          </div>
        </div>
        <div style={{ borderTop: "1px solid " + T.border, paddingTop: 12 }}>
          <div style={{ fontSize: 11, color: T.textSub, marginBottom: 10 }}>QUESTO MESE</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
            {[["Confermato", fmtEur(kpi.confM), "#4ecb8d"], ["In corso", fmtEur(kpi.attM), "#378ADD"], ["% su import", kpi.pctM.toFixed(1) + "%", "#EF9F27"]].map(([l, v, c]) => (
              <div key={l}><div style={{ fontSize: 10, color: T.textSub, marginBottom: 4 }}>{l.toUpperCase()}</div><div className="mono" style={{ fontSize: 14, fontWeight: 600, color: c }}>{v}</div></div>
            ))}
          </div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: T.textSub }}>Distribuzione stati</div>
        {STATI_OP.filter(s => byS[s] > 0).map(s => (
          <div key={s} style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 8, height: 8, borderRadius: "50%", background: STATO_COLORS[s], display: "inline-block" }} />{s}</span>
              <span className="mono" style={{ color: T.textSub }}>{byS[s]}</span>
            </div>
            <div style={{ background: T.bgAlt, borderRadius: 4, height: 4, overflow: "hidden" }}>
              <div style={{ width: (byS[s] / clienti.length * 100) + "%", height: "100%", background: STATO_COLORS[s], borderRadius: 4 }} />
            </div>
          </div>
        ))}
      </div>

      <div className="card">
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: T.textSub }}>Scaduto per agente</div>
        {agKpi.map((ag, i) => (
          <div key={ag.nome} onClick={() => onSelAg(ag.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < agKpi.length - 1 ? "1px solid " + T.border : "none", cursor: "pointer" }}>
            <div><div style={{ fontSize: 13, fontWeight: 500 }}>{ag.nome}</div><div style={{ fontSize: 11, color: T.textSub }}>{ag.n} clienti</div></div>
            <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(ag.tot)}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* --- ANAGRAFICA CLIENTI --- */
function AnagraficaClienti({ clienti, T, onUpd, showToast }) {
  const [search, setSearch] = useState("");
  const [editT, setEditT] = useState(null);
  const ff = useMemo(() => clienti.filter(c => !search || c.ragione?.toLowerCase().includes(search.toLowerCase()) || c.codice?.includes(search)).sort((a, b) => (a.ragione || "").localeCompare(b.ragione || "")), [clienti, search]);
  return (
    <div>
      <div className="st">Anagrafica Clienti</div>
      <div style={{ color: T.textSub, fontSize: 13, marginBottom: 20 }}>{clienti.length} clienti</div>
      <div className="sw"><i className="ti ti-search sic" /><input className="si" placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} /></div>
      {ff.map(c => (
        <div key={c.id} className="card" style={{ marginBottom: 10 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <div><div style={{ fontWeight: 600 }}>{c.ragione}</div><div style={{ fontSize: 11, color: T.textSub }}>{c.codice} · {c.agente}</div></div>
            <button className="btn-g" style={{ padding: "6px 12px", fontSize: 12 }} onClick={() => setEditT(c)}><i className="ti ti-edit" /></button>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12, color: T.textSub }}>
            <div>✉ {c.email || "—"}</div>
            <div>📞 {c.telefono || c.cellulare || "—"}</div>
            <div>🕐 {c.orari || "—"}</div>
            <div>👤 {c.personeRif || "—"}</div>
          </div>
        </div>
      ))}
      {editT && (
        <div className="mbg" onClick={e => e.target === e.currentTarget && setEditT(null)}>
          <div className="mo">
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>{editT.ragione}</div>
              <button className="btn-g" style={{ padding: "6px 12px" }} onClick={() => setEditT(null)}>✕</button>
            </div>
            <MAnagrafica c={editT} T={T} onClose={() => setEditT(null)} onSave={u => { onUpd(editT.id, u); setEditT(null); showToast("Salvato"); }} />
          </div>
        </div>
      )}
    </div>
  );
}

/* --- ANAGRAFICA AGENTI --- */
function AnagraficaAgenti({ agenti, T, onUpd, onAdd, showToast }) {
  const [editT, setEditT] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div><div className="st">Anagrafica Agenti</div><div style={{ color: T.textSub, fontSize: 13 }}>{agenti.length} agenti</div></div>
        <button className="btn-p" onClick={() => setShowAdd(true)}>+ Aggiungi</button>
      </div>
      {!agenti.length && <div className="card" style={{ textAlign: "center", padding: 40, color: T.textMuted }}>Importa le spaccature per creare gli agenti automaticamente</div>}
      {agenti.map(ag => (
        <div key={ag.id} className="card" style={{ marginBottom: 10 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <div><div style={{ fontWeight: 600 }}>{ag.nome}</div><div style={{ fontSize: 11, color: T.textSub }}>{ag.zona || "—"}</div></div>
            <button className="btn-g" style={{ padding: "6px 12px", fontSize: 12 }} onClick={() => setEditT(ag)}><i className="ti ti-edit" /></button>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12, color: T.textSub }}>
            <div>✉ {ag.email || "—"}</div>
            <div>📞 {ag.telefono || "—"}</div>
          </div>
        </div>
      ))}
      {editT && (
        <div className="mbg" onClick={e => e.target === e.currentTarget && setEditT(null)}>
          <div className="mo">
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>Modifica {editT.nome}</div>
              <button className="btn-g" style={{ padding: "6px 12px" }} onClick={() => setEditT(null)}>✕</button>
            </div>
            <MEditAgente ag={editT} T={T} onSave={u => { onUpd(editT.id, u); setEditT(null); showToast("Salvato"); }} />
          </div>
        </div>
      )}
      {showAdd && (
        <div className="mbg" onClick={e => e.target === e.currentTarget && setShowAdd(false)}>
          <div className="mo">
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>Nuovo agente</div>
              <button className="btn-g" style={{ padding: "6px 12px" }} onClick={() => setShowAdd(false)}>✕</button>
            </div>
            <MEditAgente ag={{ nome: "", email: "", telefono: "", zona: "" }} T={T} onSave={u => { onAdd(u); setShowAdd(false); showToast("Agente aggiunto"); }} />
          </div>
        </div>
      )}
    </div>
  );
}

import { useState, useEffect, useMemo } from "react";

const SCRIPT_URL = "https://script.google.com/macros/s/PLACEHOLDER_URL/exec";

const STATI_OP = ["CONTROLLA", "AGENTE", "AGENZIA", "L/L", "RICHIAMA", "PAGATO"];
const PRIORITA = ["ALTA", "MEDIA", "BASSA"];
const STATI_CL = ["ATTIVO", "PAGATO", "AGENZIA"];
const TIPI_CONTATTO = ["Telefono", "Email", "WhatsApp", "Altro"];
const STATI_LL = ["IN CODA", "LETTERA 1", "LETTERA 2", "LETTERA 3", "AGENZIA"];
const ESITI_RACC = ["In attesa", "Consegnata", "Non consegnata", "Rifiutata", "Indirizzo errato"];
const TIPI_TITOLO = ["Assegno a vista", "Assegno postdatato", "Cambiale", "Contanti", "Bonifico"];
const STORAGE_KEY = "crm_solleciti_v1";

const STATO_COLORS = {
  "CONTROLLA": "#378ADD", "AGENTE": "#EF9F27", "AGENZIA": "#7F77DD",
  "L/L": "#e24b4a", "RICHIAMA": "#4ecb8d", "PAGATO": "#1D9E75"
};
const PRIORITA_COLORS = { "ALTA": "#e24b4a", "MEDIA": "#EF9F27", "BASSA": "#4ecb8d" };

const IBAN = "IT79G0306909496100000011059";
const AZIENDA = "Saratoga Int. Sforza SPA";
const BANCA = "Banca Intesa San Paolo – Via Lorenteggio 70";

function loadLocal() {
  try { const r = localStorage.getItem(STORAGE_KEY); return r ? JSON.parse(r) : { clienti: [], praticheLl: [], incassi: [] }; }
  catch { return { clienti: [], praticheLl: [], incassi: [] }; }
}
function saveLocal(d) { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); }
async function loadFromSheets() { const r = await fetch(`${SCRIPT_URL}?action=load`); return r.json(); }
async function saveToSheets(d) { await fetch(`${SCRIPT_URL}?action=save&data=${encodeURIComponent(JSON.stringify(d))}`); }

function oggi() { return new Date().toISOString().slice(0, 10); }
function fmtEur(n) {
  if (n == null || isNaN(n)) return "—";
  return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(n);
}
function fmtData(d) {
  if (!d) return "—";
  if (typeof d === "number") {
    const dt = new Date((d - 25569) * 86400000);
    return dt.toLocaleDateString("it-IT");
  }
  if (d.includes("-")) {
    const [y, m, dd] = d.split("-");
    return `${dd}/${m}/${y}`;
  }
  return d;
}
function parseExcelDate(v) {
  if (!v) return null;
  if (typeof v === "number") {
    const dt = new Date((v - 25569) * 86400000);
    return dt.toISOString().slice(0, 10);
  }
  return String(v);
}
function giorniDa(dataStr) {
  if (!dataStr) return 9999;
  const d = new Date(dataStr);
  return Math.floor((new Date() - d) / 86400000);
}

// Import Excel spaccature
async function parseSpaccatureExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.find(s => s.toUpperCase().includes("SCADUT") || s.toUpperCase().includes("SPACCATURE")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(rows);
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function elaboraSpaccature(rows) {
  const clientiMap = {};
  rows.forEach(row => {
    const cod = String(row["COD."] || row["CODICE"] || row["CODICE_STD"] || row["CODICE_ORIG"] || "").trim().padStart(5, "0");
    if (!cod || cod === "00000") return;
    const ragione = String(row["RAGIONE SOCIALE"] || row["RAGIONE_SOCIALE"] || "").trim();
    const agente = String(row["AGENTE"] || "").trim();
    const localita = String(row["LOCALITA'"] || row["LOCALITA"] || "").trim();
    const prov = String(row["PROV."] || row["PROV"] || "").trim();
    const doc = String(row["DOC."] || row["DOC"] || "").trim();
    const scadenzaRaw = row["SCADENZE"] || row["SCADENZA"] || row[Object.keys(row).find(k => /\d{5}/.test(String(row[k])) && !isNaN(row[k])) || ""];
    const scadenza = parseExcelDate(typeof scadenzaRaw === "number" && scadenzaRaw > 40000 ? scadenzaRaw : null) || String(scadenzaRaw);
    const importo = parseFloat(row["IMPORTO"] || 0);
    const giorni = parseInt(row["GIORNI"] || 0);
    const tag = `${cod}|${doc}|${scadenza}|${importo.toFixed(2)}`;

    if (!clientiMap[cod]) {
      clientiMap[cod] = {
        id: cod, codice: cod, ragione, agente, localita, prov,
        email: "", telefono: "", cellulare: "",
        statoOp: "CONTROLLA", statoCl: "ATTIVO", priorita: "ALTA",
        esito: "", ultimoContatto: null, dataRichiamo: null, ultimoGiro: null,
        note: "", diario: [], scadenze: [], totaleScaduto: 0, giorniMaxRitardo: 0,
      };
    }
    const cl = clientiMap[cod];
    if (ragione && !cl.ragione) cl.ragione = ragione;
    if (agente && !cl.agente) cl.agente = agente;

    const scadenzaEsistente = cl.scadenze.find(s => s.tag === tag);
    if (!scadenzaEsistente) {
      cl.scadenze.push({ tag, doc, scadenza, importo, giorni, residuo: importo, stato: "APERTA", accorpata: false, fatturaAccorpante: null });
    }
    cl.totaleScaduto += importo;
    if (giorni > cl.giorniMaxRitardo) cl.giorniMaxRitardo = giorni;
  });
  return Object.values(clientiMap);
}

function mergeClienti(esistenti, nuovi) {
  const result = [...esistenti];
  nuovi.forEach(nc => {
    const esistente = result.find(c => c.codice === nc.codice);
    if (esistente) {
      // Aggiorna dati contabili mantenendo dati operativi
      esistente.ragione = nc.ragione || esistente.ragione;
      esistente.agente = nc.agente || esistente.agente;
      esistente.localita = nc.localita || esistente.localita;
      // Merge scadenze: aggiungi solo quelle non esistenti per tag
      nc.scadenze.forEach(ns => {
        const esiste = esistente.scadenze.find(s => s.tag === ns.tag);
        if (!esiste) esistente.scadenze.push(ns);
        else {
          // Aggiorna importo se cambiato (pagamento parziale registrato dal gestionale)
          if (ns.importo !== esiste.importo) esiste.importoOriginale = esiste.importo;
        }
      });
      // Ricalcola totale
      esistente.totaleScaduto = esistente.scadenze.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
    } else {
      result.push(nc);
    }
  });
  return result;
}

function generaMail(cliente) {
  const scadenzeAperte = (cliente.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  if (scadenzeAperte.length === 0) return null;
  let corpo = `Gentile Cliente,\n\na seguito di controlli contabili risultano ancora in sospeso i pagamenti relativi alle seguenti fatture:\n\n`;
  scadenzeAperte.forEach(s => {
    corpo += `• Fattura n. ${s.doc} scaduta il ${fmtData(s.scadenza)} – € ${(s.residuo || s.importo).toFixed(2)}\n\n`;
  });
  corpo += `Al fine di agevolare la procedura di regolarizzazione, riportiamo di seguito le coordinate per il pagamento.\n\n`;
  corpo += `Bonifico bancario:\n${IBAN}\n${AZIENDA}\n${BANCA}\n\n`;
  corpo += `Qualora abbiate già provveduto al pagamento ritenete nullo il presente sollecito.\n\nRestiamo a Sua completa disposizione per ulteriori chiarimenti.\n\nRingraziando per l'attenzione, porgiamo cordiali saluti.`;
  const subject = encodeURIComponent("Verifica stato pagamento fatture");
  const body = encodeURIComponent(corpo);
  return `mailto:${cliente.email}?subject=${subject}&body=${body}`;
}

export default function App() {
  const [page, setPage] = useState("oggi");
  const [data, setData] = useState(loadLocal);
  const [toast, setToast] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [loadingCloud, setLoadingCloud] = useState(true);
  const [clienteSelezionato, setClienteSelezionato] = useState(null);

  useEffect(() => {
    loadFromSheets()
      .then(remote => {
        if (remote && !remote.error && remote.clienti?.length > 0) {
          setData(remote); saveLocal(remote);
        }
      })
      .catch(() => {})
      .finally(() => setLoadingCloud(false));
  }, []);

  useEffect(() => { saveLocal(data); }, [data]);

  function showToast(msg, type = "ok") {
    setToast({ msg, type }); setTimeout(() => setToast(null), 3000);
  }

  async function syncToCloud(nd) {
    setSyncing(true);
    try { await saveToSheets(nd); } catch { showToast("Errore sync", "warn"); }
    setSyncing(false);
  }

  function updateCliente(id, updates) {
    setData(d => {
      const nd = { ...d, clienti: d.clienti.map(c => c.id === id ? { ...c, ...updates } : c) };
      syncToCloud(nd); return nd;
    });
  }

  function addDiario(id, voce) {
    setData(d => {
      const nd = { ...d, clienti: d.clienti.map(c => c.id === id ? { ...c, diario: [voce, ...(c.diario || [])] } : c) };
      syncToCloud(nd); return nd;
    });
  }

  async function importaSpaccature(file) {
    try {
      showToast("Importazione in corso...", "info");
      const rows = await parseSpaccatureExcel(file);
      const nuovi = elaboraSpaccature(rows);
      setData(d => {
        const clientiMerged = mergeClienti(d.clienti || [], nuovi);
        const nd = { ...d, clienti: clientiMerged };
        syncToCloud(nd);
        showToast(`Importate ${nuovi.length} posizioni, ${clientiMerged.length} clienti totali`);
        return nd;
      });
    } catch (e) { showToast("Errore importazione: " + e.message, "warn"); }
  }

  function addPraticaLL(clienteId) {
    const cliente = data.clienti.find(c => c.id === clienteId);
    if (!cliente) return;
    setData(d => {
      const nuova = { id: Date.now(), clienteId, ragione: cliente.ragione, agente: cliente.agente, stato: "IN CODA", lettera: 0, dataDecisione: oggi(), dataUltimaLettera: null, esito: "", note: "", storia: [] };
      const nd = { ...d, praticheLl: [...(d.praticheLl || []), nuova] };
      showToast(`Pratica L/L aperta per ${cliente.ragione}`);
      syncToCloud(nd); return nd;
    });
  }

  function updateLL(id, updates) {
    setData(d => {
      const nd = { ...d, praticheLl: d.praticheLl.map(p => p.id === id ? { ...p, ...updates } : p) };
      syncToCloud(nd); return nd;
    });
  }

  function addIncasso(inc) {
    setData(d => {
      const nd = { ...d, incassi: [{ ...inc, id: Date.now(), dataRicezione: oggi(), stato: "IN MANO AGENTE" }, ...(d.incassi || [])] };
      showToast("Incasso registrato");
      syncToCloud(nd); return nd;
    });
  }

  function updateIncasso(id, updates) {
    setData(d => {
      const nd = { ...d, incassi: d.incassi.map(i => i.id === id ? { ...i, ...updates } : i) };
      syncToCloud(nd); return nd;
    });
  }

  // Lista Oggi — ordinata per: richiamo oggi → mai chiamati → non chiamati da più tempo (esclusi PAGATO)
  const listaOggi = useMemo(() => {
    const attivi = (data.clienti || []).filter(c => c.statoCl !== "PAGATO" && c.statoOp !== "PAGATO");
    return [...attivi].sort((a, b) => {
      const aOggi = a.dataRichiamo === oggi() ? 0 : 1;
      const bOggi = b.dataRichiamo === oggi() ? 0 : 1;
      if (aOggi !== bOggi) return aOggi - bOggi;
      const aGiorni = giorniDa(a.ultimoContatto);
      const bGiorni = giorniDa(b.ultimoContatto);
      if (aGiorni !== bGiorni) return bGiorni - aGiorni;
      if (a.agente < b.agente) return -1;
      if (a.agente > b.agente) return 1;
      return 0;
    });
  }, [data.clienti]);

  const kpi = useMemo(() => {
    const clienti = data.clienti || [];
    const totScaduto = clienti.filter(c => c.statoCl !== "PAGATO").reduce((s, c) => s + (c.totaleScaduto || 0), 0);
    const llAperte = (data.praticheLl || []).filter(p => p.stato !== "AGENZIA" && p.stato !== "CHIUSA").length;
    const titoliMano = (data.incassi || []).filter(i => i.stato === "IN MANO AGENTE").length;
    const daRichiamare = listaOggi.filter(c => c.dataRichiamo === oggi()).length;
    return { totScaduto, llAperte, titoliMano, daRichiamare };
  }, [data.clienti, data.praticheLl, data.incassi, listaOggi]);

  const NAV = [
    { id: "oggi", label: "Oggi", icon: "ti-calendar-event" },
    { id: "clienti", label: "Clienti", icon: "ti-users" },
    { id: "ll", label: "L/L", icon: "ti-file-text" },
    { id: "incassi", label: "Incassi", icon: "ti-coin" },
    { id: "dashboard", label: "Dashboard", icon: "ti-layout-dashboard" },
  ];

  if (clienteSelezionato) {
    const cl = data.clienti.find(c => c.id === clienteSelezionato);
    if (cl) return (
      <SchedaCliente
        cliente={cl}
        onBack={() => setClienteSelezionato(null)}
        onUpdate={(u) => updateCliente(cl.id, u)}
        onAddDiario={(v) => addDiario(cl.id, v)}
        onApriLL={() => { addPraticaLL(cl.id); setClienteSelezionato(null); setPage("ll"); }}
        onAddIncasso={addIncasso}
        showToast={showToast}
      />
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", color: "#e8e6df", fontFamily: "'DM Sans', 'Segoe UI', sans-serif", paddingBottom: 72 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        @import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        input, select, textarea { background: #13131a; border: 1px solid #2a2a38; border-radius: 10px; color: #e8e6df; padding: 10px 14px; font-family: inherit; font-size: 15px; width: 100%; outline: none; -webkit-appearance: none; }
        input:focus, select:focus, textarea:focus { border-color: #4a7fd4; }
        select option { background: #13131a; }
        textarea { resize: vertical; min-height: 80px; }
        button { cursor: pointer; font-family: inherit; font-size: 14px; border: none; border-radius: 10px; padding: 10px 18px; transition: all .15s; -webkit-tap-highlight-color: transparent; }
        .btn-primary { background: #4a7fd4; color: #fff; font-weight: 500; }
        .btn-primary:active { transform: scale(0.98); }
        .btn-ghost { background: transparent; color: #9a98a0; border: 1px solid #2a2a38; }
        .btn-danger { background: transparent; color: #e24b4a; border: 1px solid #3a2020; }
        .btn-green { background: #0f2a1a; color: #4ecb8d; border: 1px solid #1a4a2a; }
        .card { background: #13131a; border: 1px solid #1e1e2a; border-radius: 16px; padding: 16px; }
        .tag { display: inline-block; font-size: 11px; font-weight: 500; padding: 3px 9px; border-radius: 20px; }
        label { font-size: 12px; color: #7a7888; margin-bottom: 5px; display: block; letter-spacing: .04em; }
        .modal-bg { position: fixed; inset: 0; background: rgba(0,0,0,.85); display: flex; align-items: flex-end; justify-content: center; z-index: 200; }
        .modal { background: #13131a; border: 1px solid #2a2a38; border-radius: 20px 20px 0 0; padding: 24px 20px 40px; width: 100%; max-width: 500px; max-height: 92vh; overflow-y: auto; }
        .mono { font-family: 'DM Mono', monospace; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .spin { animation: spin 1s linear infinite; display: inline-block; }
        .bottom-nav { position: fixed; bottom: 0; left: 0; right: 0; background: #0d0d14; border-top: 1px solid #1a1a25; display: flex; z-index: 100; padding-bottom: env(safe-area-inset-bottom); }
        .nav-btn { flex: 1; display: flex; flex-direction: column; align-items: center; gap: 2px; padding: 8px 4px; background: none; border: none; border-radius: 0; font-size: 10px; letter-spacing: .04em; }
        .nav-btn.active { color: #4a7fd4; }
        .nav-btn:not(.active) { color: #4a4858; }
        .nav-icon { font-size: 22px; }
        .section-title { font-size: 22px; font-weight: 600; letter-spacing: -.02em; margin-bottom: 4px; }
        .section-sub { color: #5a5868; font-size: 13px; margin-bottom: 20px; }
        .form-group { margin-bottom: 14px; }
        .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .cl-card { background: #13131a; border: 1px solid #1e1e2a; border-radius: 14px; padding: 14px; margin-bottom: 10px; cursor: pointer; transition: border-color .15s; }
        .cl-card:active { border-color: #4a7fd4; }
        .divider { border: none; border-top: 1px solid #1e1e2a; margin: 16px 0; }
        .oggi-fascia { margin-bottom: 20px; }
        .fascia-label { font-size: 11px; color: #5a5868; letter-spacing: .08em; text-transform: uppercase; margin-bottom: 10px; font-weight: 500; }
        .priorita-bar { width: 4px; border-radius: 2px; align-self: stretch; flex-shrink: 0; }
        .badge-alert { background: #1f1808; border: 1px solid #4a3010; color: #EF9F27; border-radius: 8px; padding: 2px 8px; font-size: 11px; }
        .info-box { background: #0a1f2a; border: 1px solid #1a3a4a; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #378ADD; margin-bottom: 14px; }
        .warn-box { background: #1f1808; border: 1px solid #3a2e08; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #EF9F27; margin-bottom: 14px; }
        .success-box { background: #0f2a1a; border: 1px solid #1a4a2a; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #4ecb8d; margin-bottom: 14px; }
        .diario-item { border-left: 3px solid #378ADD; padding: 8px 12px; background: #0d0d18; border-radius: 0 8px 8px 0; margin-bottom: 8px; }
        .ll-step { display: flex; align-items: center; gap: 8px; padding: 10px 12px; background: #0d0d18; border-radius: 10px; margin-bottom: 8px; }
        .ll-step-dot { width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 600; flex-shrink: 0; }
        .search-input { background: #13131a; border: 1px solid #2a2a38; border-radius: 12px; color: #e8e6df; padding: 10px 14px 10px 38px; font-size: 15px; width: 100%; outline: none; }
        .search-wrap { position: relative; margin-bottom: 16px; }
        .search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #5a5868; font-size: 18px; }
      `}</style>

      {/* HEADER */}
      <div style={{ padding: "52px 20px 0", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 13, color: "#4a4858", letterSpacing: ".06em" }}>CRM SOLLECITI</div>
          <div style={{ fontSize: 11, marginTop: 2, color: syncing ? "#378ADD" : "#3a3848" }}>
            {syncing ? <><span className="spin">↻</span> Sync...</> : "● " + new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "short" })}
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {kpi.daRichiamare > 0 && <span className="badge-alert">📅 {kpi.daRichiamare} richiami</span>}
          {kpi.titoliMano > 0 && <span style={{ background: "#0a1f2a", border: "1px solid #1a3a4a", color: "#378ADD", borderRadius: 8, padding: "2px 8px", fontSize: 11 }}>💼 {kpi.titoliMano} titoli</span>}
        </div>
      </div>

      <div style={{ padding: "0 20px" }}>
        {loadingCloud ? (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "50vh", gap: 16, color: "#4a4858" }}>
            <span className="spin" style={{ fontSize: 36 }}>↻</span>
            <div style={{ fontSize: 14 }}>Carico dal cloud...</div>
          </div>
        ) : (
          <>
            {page === "oggi" && <Oggi lista={listaOggi} onSelectCliente={setClienteSelezionato} />}
            {page === "clienti" && <Clienti clienti={data.clienti || []} onSelectCliente={setClienteSelezionato} onImporta={importaSpaccature} showToast={showToast} />}
            {page === "ll" && <PraticheLl pratiche={data.praticheLl || []} clienti={data.clienti || []} onUpdate={updateLL} onSelectCliente={setClienteSelezionato} showToast={showToast} />}
            {page === "incassi" && <Incassi incassi={data.incassi || []} clienti={data.clienti || []} onAdd={addIncasso} onUpdate={updateIncasso} onSelectCliente={setClienteSelezionato} showToast={showToast} />}
            {page === "dashboard" && <Dashboard kpi={kpi} clienti={data.clienti || []} praticheLl={data.praticheLl || []} incassi={data.incassi || []} />}
          </>
        )}
      </div>

      {/* BOTTOM NAV */}
      <nav className="bottom-nav">
        {NAV.map(n => (
          <button key={n.id} className={`nav-btn ${page === n.id ? "active" : ""}`} onClick={() => setPage(n.id)}>
            <i className={`ti ${n.icon} nav-icon`} aria-hidden="true" />
            {n.label}
          </button>
        ))}
      </nav>

      {toast && (
        <div style={{
          position: "fixed", bottom: 88, left: 20, right: 20,
          background: toast.type === "warn" ? "#2a1f0a" : toast.type === "info" ? "#0a1f2a" : "#0f2a1a",
          border: `1px solid ${toast.type === "warn" ? "#4a3010" : toast.type === "info" ? "#1a3a4a" : "#1a4a2a"}`,
          color: toast.type === "warn" ? "#EF9F27" : toast.type === "info" ? "#378ADD" : "#4ecb8d",
          padding: "12px 18px", borderRadius: 12, fontSize: 14, zIndex: 300, textAlign: "center",
        }}>{toast.msg}</div>
      )}
    </div>
  );
}

/* ─── OGGI ─── */
function Oggi({ lista, onSelectCliente }) {
  const oggi_ = oggi();
  const richiami = lista.filter(c => c.dataRichiamo === oggi_);
  const maiChiamati = lista.filter(c => !c.ultimoContatto && c.dataRichiamo !== oggi_);
  const vecchi = lista.filter(c => c.ultimoContatto && c.dataRichiamo !== oggi_).sort((a, b) => giorniDa(a.ultimoContatto) - giorniDa(b.ultimoContatto) > 0 ? -1 : 1);

  const Gruppo = ({ label, items, colore }) => items.length === 0 ? null : (
    <div className="oggi-fascia">
      <div className="fascia-label" style={{ color: colore }}>{label} ({items.length})</div>
      {items.map(c => <CardOggi key={c.id} cliente={c} onClick={() => onSelectCliente(c.id)} />)}
    </div>
  );

  if (lista.length === 0) return (
    <div>
      <div className="section-title">Oggi</div>
      <div className="section-sub">Lista chiamate prioritizzata</div>
      <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>
        <i className="ti ti-calendar-off" style={{ fontSize: 40, marginBottom: 12, display: "block" }} />
        <div style={{ fontSize: 14 }}>Nessun cliente ancora</div>
        <div style={{ fontSize: 13, marginTop: 8 }}>Importa le spaccature dalla sezione Clienti</div>
      </div>
    </div>
  );

  return (
    <div>
      <div className="section-title">Oggi</div>
      <div className="section-sub">{lista.length} clienti attivi</div>
      <Gruppo label="📅 Richiami programmati" items={richiami} colore="#EF9F27" />
      <Gruppo label="🆕 Mai contattati" items={maiChiamati} colore="#e24b4a" />
      <Gruppo label="📞 Da richiamare" items={vecchi} colore="#9a98a0" />
    </div>
  );
}

function CardOggi({ cliente, onClick }) {
  const c = cliente;
  const giorni = giorniDa(c.ultimoContatto);
  return (
    <div className="cl-card" onClick={onClick} style={{ display: "flex", gap: 12, alignItems: "stretch" }}>
      <div className="priorita-bar" style={{ background: PRIORITA_COLORS[c.priorita] || "#888" }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2 }}>{c.ragione}</div>
            <div style={{ fontSize: 11, color: "#5a5868" }}>{c.agente} · {c.codice}</div>
          </div>
          <div style={{ textAlign: "right", flexShrink: 0 }}>
            <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#e24b4a" }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(c.totaleScaduto || 0)}</div>
            <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888", marginTop: 4, display: "inline-block" }}>{c.statoOp}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {c.ultimoContatto ? <span style={{ fontSize: 11, color: "#5a5868" }}>Ultimo: {fmtData(c.ultimoContatto)} ({giorni < 9999 ? `${giorni}gg fa` : "—"})</span> : <span style={{ fontSize: 11, color: "#e24b4a" }}>Mai contattato</span>}
          {c.dataRichiamo === oggi() && <span className="badge-alert">⏰ oggi</span>}
          {c.esito && <span style={{ fontSize: 11, color: "#7a7888", fontStyle: "italic" }}>"{c.esito}"</span>}
        </div>
      </div>
    </div>
  );
}

/* ─── CLIENTI ─── */
function Clienti({ clienti, onSelectCliente, onImporta, showToast }) {
  const [search, setSearch] = useState("");
  const [filtroStato, setFiltroStato] = useState("TUTTI");
  const [filtroAgente, setFiltroAgente] = useState("TUTTI");

  const agenti = useMemo(() => ["TUTTI", ...new Set(clienti.map(c => c.agente).filter(Boolean))].sort(), [clienti]);

  const filtrati = useMemo(() => {
    return clienti.filter(c => {
      if (filtroStato !== "TUTTI" && c.statoOp !== filtroStato) return false;
      if (filtroAgente !== "TUTTI" && c.agente !== filtroAgente) return false;
      if (search && !c.ragione?.toLowerCase().includes(search.toLowerCase()) && !c.codice?.includes(search)) return false;
      return true;
    }).sort((a, b) => {
      if (a.agente < b.agente) return -1;
      if (a.agente > b.agente) return 1;
      return (b.totaleScaduto || 0) - (a.totaleScaduto || 0);
    });
  }, [clienti, search, filtroStato, filtroAgente]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div>
          <div className="section-title">Clienti</div>
          <div className="section-sub">{clienti.length} clienti · {filtrati.length} mostrati</div>
        </div>
        <label style={{ background: "#4a7fd4", color: "#fff", borderRadius: 10, padding: "10px 14px", fontSize: 14, fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>
          <i className="ti ti-upload" /> Importa
          <input type="file" accept=".xlsx,.xls,.xlsm" style={{ display: "none" }} onChange={e => e.target.files[0] && onImporta(e.target.files[0])} />
        </label>
      </div>

      <div className="search-wrap">
        <i className="ti ti-search search-icon" />
        <input className="search-input" placeholder="Cerca per nome o codice..." value={search} onChange={e => setSearch(e.target.value)} />
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto", paddingBottom: 4 }}>
        {["TUTTI", ...STATI_OP].map(s => (
          <button key={s} onClick={() => setFiltroStato(s)} style={{
            padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap",
            background: filtroStato === s ? "#4a7fd4" : "transparent",
            color: filtroStato === s ? "#fff" : "#5a5868",
            border: filtroStato === s ? "none" : "1px solid #2a2a38",
          }}>{s}</button>
        ))}
      </div>

      <div style={{ marginBottom: 14 }}>
        <select value={filtroAgente} onChange={e => setFiltroAgente(e.target.value)}>
          {agenti.map(a => <option key={a}>{a}</option>)}
        </select>
      </div>

      {filtrati.length === 0 ? (
        <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>
          <div style={{ fontSize: 13 }}>Nessun cliente trovato</div>
        </div>
      ) : filtrati.map(c => <CardOggi key={c.id} cliente={c} onClick={() => onSelectCliente(c.id)} />)}
    </div>
  );
}

/* ─── SCHEDA CLIENTE ─── */
function SchedaCliente({ cliente, onBack, onUpdate, onAddDiario, onApriLL, onAddIncasso, showToast }) {
  const c = cliente;
  const [tab, setTab] = useState("info");
  const [showContatto, setShowContatto] = useState(false);
  const [showIncasso, setShowIncasso] = useState(false);
  const [showAccorpa, setShowAccorpa] = useState(false);

  // Form contatto
  const [cData, setCData] = useState(oggi());
  const [cOra, setCOra] = useState(new Date().toTimeString().slice(0, 5));
  const [cCanale, setCCanale] = useState("Telefono");
  const [cEsito, setCEsito] = useState("");
  const [cStato, setCStato] = useState(c.statoOp);
  const [cRichiamo, setCRichiamo] = useState(c.dataRichiamo || "");
  const [cPriorita, setCPriorita] = useState(c.priorita);

  function registraContatto(e) {
    e.preventDefault();
    const voce = { data: cData, ora: cOra, canale: cCanale, esito: cEsito, stato: cStato, ts: new Date().toISOString() };
    onAddDiario(voce);
    onUpdate({ statoOp: cStato, esito: cEsito, ultimoContatto: cData, dataRichiamo: cRichiamo || null, priorita: cPriorita });
    setShowContatto(false);
    showToast("Contatto registrato");
  }

  const mailLink = generaMail(c);
  const scadenzeAperte = (c.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", color: "#e8e6df", fontFamily: "'DM Sans', 'Segoe UI', sans-serif", paddingBottom: 40 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        @import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        input, select, textarea { background: #13131a; border: 1px solid #2a2a38; border-radius: 10px; color: #e8e6df; padding: 10px 14px; font-family: inherit; font-size: 15px; width: 100%; outline: none; -webkit-appearance: none; }
        input:focus, select:focus { border-color: #4a7fd4; }
        select option { background: #13131a; }
        button { cursor: pointer; font-family: inherit; font-size: 14px; border: none; border-radius: 10px; padding: 10px 18px; transition: all .15s; }
        .btn-primary { background: #4a7fd4; color: #fff; font-weight: 500; }
        .btn-ghost { background: transparent; color: #9a98a0; border: 1px solid #2a2a38; }
        .btn-danger { background: transparent; color: #e24b4a; border: 1px solid #3a2020; }
        .btn-green { background: #0f2a1a; color: #4ecb8d; border: 1px solid #1a4a2a; }
        .card { background: #13131a; border: 1px solid #1e1e2a; border-radius: 16px; padding: 16px; }
        .tag { display: inline-block; font-size: 11px; font-weight: 500; padding: 3px 9px; border-radius: 20px; }
        label { font-size: 12px; color: #7a7888; margin-bottom: 5px; display: block; letter-spacing: .04em; }
        .modal-bg { position: fixed; inset: 0; background: rgba(0,0,0,.85); display: flex; align-items: flex-end; justify-content: center; z-index: 200; }
        .modal { background: #13131a; border: 1px solid #2a2a38; border-radius: 20px 20px 0 0; padding: 24px 20px 40px; width: 100%; max-width: 500px; max-height: 92vh; overflow-y: auto; }
        .mono { font-family: 'DM Mono', monospace; }
        .form-group { margin-bottom: 14px; }
        .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .diario-item { border-left: 3px solid #378ADD; padding: 8px 12px; background: #0d0d18; border-radius: 0 8px 8px 0; margin-bottom: 8px; }
        .tab-bar { display: flex; gap: 8px; margin-bottom: 16px; overflow-x: auto; }
        .tab { padding: 8px 16px; border-radius: 20px; font-size: 13px; border: 1px solid #2a2a38; background: transparent; color: #5a5868; white-space: nowrap; }
        .tab.active { background: #4a7fd4; color: #fff; border-color: transparent; }
        .scad-row { background: #0d0d14; border-radius: 10px; padding: 10px 12px; margin-bottom: 8px; }
        .warn-box { background: #1f1808; border: 1px solid #3a2e08; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #EF9F27; margin-bottom: 14px; }
        .info-box { background: #0a1f2a; border: 1px solid #1a3a4a; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #378ADD; margin-bottom: 14px; }
        .kv { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #1a1a25; font-size: 13px; }
        .kv:last-child { border-bottom: none; }
      `}</style>

      {/* HEADER SCHEDA */}
      <div style={{ padding: "52px 20px 0", marginBottom: 20 }}>
        <button className="btn-ghost" style={{ marginBottom: 16, padding: "8px 14px", fontSize: 13 }} onClick={onBack}>
          <i className="ti ti-arrow-left" /> Torna
        </button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>{c.ragione}</div>
            <div style={{ fontSize: 12, color: "#5a5868" }}>{c.codice} · {c.agente}</div>
            <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
              <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888" }}>{c.statoOp}</span>
              <span className="tag" style={{ background: (PRIORITA_COLORS[c.priorita] || "#888") + "22", color: PRIORITA_COLORS[c.priorita] || "#888" }}>{c.priorita}</span>
              <span className="tag" style={{ background: "#1a1a25", color: "#5a5868" }}>{c.localita}</span>
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div className="mono" style={{ fontSize: 22, fontWeight: 700, color: "#e24b4a" }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(c.totaleScaduto || 0)}</div>
            <div style={{ fontSize: 11, color: "#5a5868", marginTop: 2 }}>scaduto totale</div>
          </div>
        </div>
      </div>

      <div style={{ padding: "0 20px" }}>
        {/* AZIONI RAPIDE */}
        <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
          <button className="btn-primary" style={{ flex: 1, minWidth: 120 }} onClick={() => setShowContatto(true)}>
            <i className="ti ti-phone" /> Registra contatto
          </button>
          {mailLink && (
            <a href={mailLink} style={{ flex: 1, minWidth: 120, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#0a1f2a", color: "#378ADD", border: "1px solid #1a3a4a", borderRadius: 10, padding: "10px 18px", fontSize: 14, textDecoration: "none" }}>
              <i className="ti ti-mail" /> Genera mail
            </a>
          )}
          <button className="btn-ghost" style={{ flex: 1, minWidth: 120 }} onClick={() => setShowIncasso(true)}>
            <i className="ti ti-coin" /> Incasso agente
          </button>
        </div>

        {/* ULTIMO ESITO */}
        {c.esito && (
          <div className="warn-box" style={{ background: "#13131a", border: "1px solid #2a2a38", color: "#9a98a0" }}>
            <span style={{ fontSize: 11, color: "#5a5868" }}>ULTIMO ESITO</span><br />
            {c.esito}
            {c.ultimoContatto && <span style={{ fontSize: 11, color: "#5a5868", marginLeft: 8 }}>· {fmtData(c.ultimoContatto)}</span>}
          </div>
        )}
        {c.dataRichiamo && (
          <div className="warn-box">
            ⏰ Richiamo programmato: {fmtData(c.dataRichiamo)}
          </div>
        )}

        {/* TAB */}
        <div className="tab-bar">
          {[["info", "Info"], ["scadenze", `Scadenze (${scadenzeAperte.length})`], ["diario", `Diario (${(c.diario || []).length})`], ["azioni", "Azioni"]].map(([k, l]) => (
            <button key={k} className={`tab ${tab === k ? "active" : ""}`} onClick={() => setTab(k)}>{l}</button>
          ))}
        </div>

        {/* TAB INFO */}
        {tab === "info" && (
          <div className="card">
            {[
              ["Email", c.email || "—"],
              ["Telefono", c.telefono || "—"],
              ["Cellulare", c.cellulare || "—"],
              ["Agente", c.agente || "—"],
              ["Località", c.localita || "—"],
              ["Scadenza più vecchia", c.giorniMaxRitardo ? `${c.giorniMaxRitardo} giorni` : "—"],
            ].map(([k, v]) => (
              <div key={k} className="kv"><span style={{ color: "#5a5868" }}>{k}</span><span style={{ fontWeight: 500 }}>{v}</span></div>
            ))}
            <div style={{ marginTop: 16, display: "flex", gap: 8 }}>
              <button className="btn-danger" style={{ fontSize: 13, padding: "8px 14px" }} onClick={onApriLL}>
                <i className="ti ti-file-text" /> Apri pratica L/L
              </button>
            </div>
          </div>
        )}

        {/* TAB SCADENZE */}
        {tab === "scadenze" && (
          <div>
            <button className="btn-ghost" style={{ width: "100%", marginBottom: 12, fontSize: 13 }} onClick={() => setShowAccorpa(true)}>
              <i className="ti ti-arrows-join" /> Registra pagamento / accorpa fatture
            </button>
            {scadenzeAperte.length === 0 ? (
              <div className="card" style={{ textAlign: "center", padding: 30, color: "#4a4858" }}>Nessuna scadenza aperta</div>
            ) : scadenzeAperte.map((s, i) => (
              <div key={i} className="scad-row">
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ fontWeight: 500, fontSize: 14 }}>Ft. {s.doc}</span>
                  <span className="mono" style={{ fontWeight: 600 }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(s.residuo || s.importo)}</span>
                </div>
                <div style={{ fontSize: 12, color: "#5a5868" }}>
                  Scad: {fmtData(s.scadenza)} · {s.giorni || 0}gg ritardo
                  {s.residuo < s.importo && <span style={{ color: "#EF9F27", marginLeft: 8 }}>Parziale (orig. {new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(s.importo)})</span>}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* TAB DIARIO */}
        {tab === "diario" && (
          <div>
            {(c.diario || []).length === 0 ? (
              <div className="card" style={{ textAlign: "center", padding: 30, color: "#4a4858" }}>Nessun contatto registrato</div>
            ) : (c.diario || []).map((v, i) => (
              <div key={i} className="diario-item">
                <div style={{ display: "flex", gap: 8, marginBottom: 4, alignItems: "center' " }}>
                  <span style={{ fontSize: 11, color: "#5a5868" }}>{fmtData(v.data)} {v.ora}</span>
                  <span className="tag" style={{ background: "#1a1a25", color: "#9a98a0" }}>{v.canale}</span>
                  {v.stato && <span className="tag" style={{ background: (STATO_COLORS[v.stato] || "#888") + "22", color: STATO_COLORS[v.stato] || "#888" }}>{v.stato}</span>}
                </div>
                <div style={{ fontSize: 13 }}>{v.esito}</div>
              </div>
            ))}
          </div>
        )}

        {/* TAB AZIONI */}
        {tab === "azioni" && (
          <div style={{ display: "grid", gap: 10 }}>
            {[
              { label: "Modifica contatti", icon: "ti-edit", action: () => {} },
              { label: "Apri pratica L/L", icon: "ti-file-text", action: onApriLL, color: "#e24b4a" },
              { label: "Registra incasso agente", icon: "ti-coin", action: () => setShowIncasso(true) },
            ].map(({ label, icon, action, color }) => (
              <button key={label} className="btn-ghost" style={{ width: "100%", textAlign: "left", padding: "14px 16px", display: "flex", alignItems: "center", gap: 10, color: color || "#9a98a0" }} onClick={action}>
                <i className={`ti ${icon}`} style={{ fontSize: 18 }} /> {label}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* MODAL REGISTRA CONTATTO */}
      {showContatto && (
        <div className="modal-bg" onClick={e => e.target === e.currentTarget && setShowContatto(false)}>
          <div className="modal">
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>Registra contatto</div>
              <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={() => setShowContatto(false)}>✕</button>
            </div>
            <form onSubmit={registraContatto}>
              <div className="row2">
                <div className="form-group"><label>DATA</label><input type="date" value={cData} onChange={e => setCData(e.target.value)} /></div>
                <div className="form-group"><label>ORA</label><input type="time" value={cOra} onChange={e => setCOra(e.target.value)} /></div>
              </div>
              <div className="form-group"><label>CANALE</label>
                <select value={cCanale} onChange={e => setCCanale(e.target.value)}>
                  {TIPI_CONTATTO.map(t => <option key={t}>{t}</option>)}
                </select>
              </div>
              <div className="form-group"><label>ESITO / NOTE</label>
                <textarea value={cEsito} onChange={e => setCEsito(e.target.value)} placeholder="es. non risponde, paga a rate, passato per L/L..." />
              </div>
              <div className="row2">
                <div className="form-group"><label>NUOVO STATO</label>
                  <select value={cStato} onChange={e => setCStato(e.target.value)}>
                    {STATI_OP.map(s => <option key={s}>{s}</option>)}
                  </select>
                </div>
                <div className="form-group"><label>PRIORITÀ</label>
                  <select value={cPriorita} onChange={e => setCPriorita(e.target.value)}>
                    {PRIORITA.map(p => <option key={p}>{p}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-group"><label>DATA RICHIAMO — lascia vuoto se non necessario</label>
                <input type="date" value={cRichiamo} onChange={e => setCRichiamo(e.target.value)} />
              </div>
              <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }}>Salva contatto</button>
            </form>
          </div>
        </div>
      )}

      {/* MODAL INCASSO AGENTE */}
      {showIncasso && (
        <ModalIncasso cliente={c} onClose={() => setShowIncasso(false)} onSave={inc => { onAddIncasso(inc); setShowIncasso(false); showToast("Incasso registrato"); }} />
      )}

      {/* MODAL ACCORPAMENTO */}
      {showAccorpa && (
        <ModalAccorpa cliente={c} onClose={() => setShowAccorpa(false)} onSave={(updates) => { onUpdate(updates); setShowAccorpa(false); showToast("Scadenze aggiornate"); }} />
      )}
    </div>
  );
}

/* ─── MODAL INCASSO ─── */
function ModalIncasso({ cliente, onClose, onSave }) {
  const [tipo, setTipo] = useState("Assegno a vista");
  const [importo, setImporto] = useState("");
  const [spese, setSpese] = useState("0");
  const [note, setNote] = useState("");
  const [fatture, setFatture] = useState("");

  function handleSave(e) {
    e.preventDefault();
    onSave({ clienteId: cliente.id, ragione: cliente.ragione, agente: cliente.agente, tipo, importo: parseFloat(importo), spese: parseFloat(spese || 0), incassoNetto: parseFloat(importo) - parseFloat(spese || 0), note, fatture, stato: "IN MANO AGENTE" });
  }

  return (
    <div className="modal-bg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <style>{`* { box-sizing: border-box; } input, select, textarea { background: #13131a; border: 1px solid #2a2a38; border-radius: 10px; color: #e8e6df; padding: 10px 14px; font-family: inherit; font-size: 15px; width: 100%; outline: none; } input:focus, select:focus { border-color: #4a7fd4; } select option { background: #13131a; } textarea { resize: vertical; min-height: 60px; } label { font-size: 12px; color: #7a7888; margin-bottom: 5px; display: block; } button { cursor: pointer; font-family: inherit; font-size: 14px; border: none; border-radius: 10px; padding: 10px 18px; } .btn-primary { background: #4a7fd4; color: #fff; } .btn-ghost { background: transparent; color: #9a98a0; border: 1px solid #2a2a38; } .form-group { margin-bottom: 14px; } .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }`}</style>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Incasso agente — {cliente.ragione}</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group"><label>TIPO TITOLO</label>
            <select value={tipo} onChange={e => setTipo(e.target.value)}>
              {TIPI_TITOLO.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div className="row2">
            <div className="form-group"><label>IMPORTO TOTALE (€)</label><input type="number" step="0.01" value={importo} onChange={e => setImporto(e.target.value)} placeholder="0.00" required /></div>
            <div className="form-group"><label>DI CUI SPESE (€)</label><input type="number" step="0.01" value={spese} onChange={e => setSpese(e.target.value)} placeholder="0.00" /></div>
          </div>
          {importo && <div style={{ background: "#0f2a1a", border: "1px solid #1a4a2a", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#4ecb8d" }}>
            Incasso netto: € {(parseFloat(importo || 0) - parseFloat(spese || 0)).toFixed(2)}
          </div>}
          <div className="form-group"><label>FATTURE DI RIFERIMENTO</label><input value={fatture} onChange={e => setFatture(e.target.value)} placeholder="es. Ft. 001 + Ft. 002 o saldo tutte" /></div>
          <div className="form-group"><label>NOTE</label><textarea value={note} onChange={e => setNote(e.target.value)} placeholder="es. saldo ft 01 +5€ spese" /></div>
          <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }}>Registra incasso</button>
        </form>
      </div>
    </div>
  );
}

/* ─── MODAL ACCORPA ─── */
function ModalAccorpa({ cliente, onClose, onSave }) {
  const scadenzeAperte = (cliente.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  const [selezionate, setSelezionate] = useState([]);
  const [tipoOp, setTipoOp] = useState("parziale");
  const [importoPagato, setImportoPagato] = useState("");
  const [fatturaAccorpante, setFatturaAccorpante] = useState("");

  function toggleSel(tag) {
    setSelezionate(s => s.includes(tag) ? s.filter(t => t !== tag) : [...s, tag]);
  }

  function handleSave(e) {
    e.preventDefault();
    const nuoveScadenze = cliente.scadenze.map(s => {
      if (!selezionate.includes(s.tag)) return s;
      if (tipoOp === "pagata") return { ...s, stato: "PAGATA", residuo: 0 };
      if (tipoOp === "parziale") return { ...s, residuo: s.importo - parseFloat(importoPagato || 0) };
      if (tipoOp === "accorpa") return { ...s, accorpata: true, fatturaAccorpante };
      return s;
    });
    const nuovoTotale = nuoveScadenze.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
    onSave({ scadenze: nuoveScadenze, totaleScaduto: nuovoTotale });
  }

  return (
    <div className="modal-bg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <style>{`* { box-sizing: border-box; } input, select { background: #13131a; border: 1px solid #2a2a38; border-radius: 10px; color: #e8e6df; padding: 10px 14px; font-family: inherit; font-size: 15px; width: 100%; outline: none; } select option { background: #13131a; } label { font-size: 12px; color: #7a7888; margin-bottom: 5px; display: block; } button { cursor: pointer; font-family: inherit; font-size: 14px; border: none; border-radius: 10px; padding: 10px 18px; } .btn-primary { background: #4a7fd4; color: #fff; } .btn-ghost { background: transparent; color: #9a98a0; border: 1px solid #2a2a38; } .form-group { margin-bottom: 14px; }`}</style>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Gestisci pagamento</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group"><label>TIPO OPERAZIONE</label>
            <select value={tipoOp} onChange={e => setTipoOp(e.target.value)}>
              <option value="pagata">Pagata interamente</option>
              <option value="parziale">Pagamento parziale</option>
              <option value="accorpa">Accorpa sotto altra fattura</option>
            </select>
          </div>

          {tipoOp === "parziale" && (
            <div className="form-group"><label>IMPORTO PAGATO (€)</label><input type="number" step="0.01" value={importoPagato} onChange={e => setImportoPagato(e.target.value)} placeholder="0.00" /></div>
          )}
          {tipoOp === "accorpa" && (
            <div className="form-group"><label>FATTURA ACCORPANTE (numero doc principale)</label><input value={fatturaAccorpante} onChange={e => setFatturaAccorpante(e.target.value)} placeholder="es. 12345" /></div>
          )}

          <div className="form-group">
            <label>SELEZIONA SCADENZE ({selezionate.length} selezionate)</label>
            {scadenzeAperte.map((s, i) => (
              <div key={i} onClick={() => toggleSel(s.tag)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: selezionate.includes(s.tag) ? "#0f2a1a" : "#0d0d14", border: `1px solid ${selezionate.includes(s.tag) ? "#1a4a2a" : "#1a1a25"}`, borderRadius: 10, marginBottom: 6, cursor: "pointer" }}>
                <div style={{ width: 20, height: 20, borderRadius: 4, border: `2px solid ${selezionate.includes(s.tag) ? "#4ecb8d" : "#2a2a38"}`, background: selezionate.includes(s.tag) ? "#4ecb8d22" : "transparent", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {selezionate.includes(s.tag) && <span style={{ color: "#4ecb8d", fontSize: 12 }}>✓</span>}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 500 }}>Ft. {s.doc} — € {(s.residuo || s.importo).toFixed(2)}</div>
                  <div style={{ fontSize: 11, color: "#5a5868" }}>Scad: {fmtData(s.scadenza)}</div>
                </div>
              </div>
            ))}
          </div>

          <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }} disabled={selezionate.length === 0}>
            Applica operazione
          </button>
        </form>
      </div>
    </div>
  );
}

/* ─── PRATICHE L/L ─── */
function PraticheLl({ pratiche, clienti, onUpdate, onSelectCliente, showToast }) {
  const [filtro, setFiltro] = useState("ATTIVE");
  const praticheFiltrate = pratiche.filter(p => filtro === "TUTTE" ? true : filtro === "ATTIVE" ? p.stato !== "AGENZIA" && p.stato !== "CHIUSA" : p.stato === filtro);

  function avanzaLettera(p) {
    const numLettera = p.lettera || 0;
    if (numLettera >= 3) { onUpdate(p.id, { stato: "AGENZIA" }); showToast("Pratica passata ad Agenzia"); return; }
    const nuovoNum = numLettera + 1;
    onUpdate(p.id, { lettera: nuovoNum, stato: `LETTERA ${nuovoNum}`, dataUltimaLettera: oggi(), esito: "" });
    showToast(`Lettera ${nuovoNum} registrata`);
  }

  function registraEsito(id, esito) {
    onUpdate(id, { esito, dataEsito: oggi() });
    showToast("Esito aggiornato");
  }

  const giorniDaLettera = (p) => p.dataUltimaLettera ? giorniDa(p.dataUltimaLettera) : null;
  const pronta = (p) => { const g = giorniDaLettera(p); return g === null || g >= 20; };

  return (
    <div>
      <div className="section-title">Pratiche L/L</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 16 }}>{pratiche.filter(p => p.stato !== "AGENZIA" && p.stato !== "CHIUSA").length} aperte</div>

      <div style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto" }}>
        {["ATTIVE", "IN CODA", "LETTERA 1", "LETTERA 2", "LETTERA 3", "AGENZIA", "TUTTE"].map(f => (
          <button key={f} onClick={() => setFiltro(f)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: filtro === f ? "#4a7fd4" : "transparent", color: filtro === f ? "#fff" : "#5a5868", border: filtro === f ? "none" : "1px solid #2a2a38" }}>{f}</button>
        ))}
      </div>

      {praticheFiltrate.length === 0 ? (
        <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>
          <i className="ti ti-file-off" style={{ fontSize: 36, marginBottom: 12, display: "block" }} />
          <div style={{ fontSize: 13 }}>Nessuna pratica L/L</div>
        </div>
      ) : praticheFiltrate.map(p => {
        const gg = giorniDaLettera(p);
        const isPronta = pronta(p);
        return (
          <div key={p.id} className="card" style={{ marginBottom: 12, borderColor: !isPronta ? "#1e1e2a" : p.lettera > 0 ? "#3a2020" : "#1e1e2a" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 2, cursor: "pointer", color: "#4a7fd4" }} onClick={() => { const cl = clienti.find(c => c.id === p.clienteId); if (cl) onSelectCliente(cl.id); }}>
                  {p.ragione}
                </div>
                <div style={{ fontSize: 11, color: "#5a5868" }}>{p.agente} · aperta il {fmtData(p.dataDecisione)}</div>
              </div>
              <span style={{ background: p.stato === "AGENZIA" ? "#7F77DD22" : p.stato === "IN CODA" ? "#EF9F2722" : "#e24b4a22", color: p.stato === "AGENZIA" ? "#7F77DD" : p.stato === "IN CODA" ? "#EF9F27" : "#e24b4a", fontSize: 12, fontWeight: 500, padding: "4px 10px", borderRadius: 8 }}>
                {p.stato}
              </span>
            </div>

            {/* STEP VISUALE */}
            <div style={{ display: "flex", gap: 4, marginBottom: 12 }}>
              {["CODA", "L1", "L2", "L3", "AG"].map((step, i) => (
                <div key={step} style={{ flex: 1, height: 4, borderRadius: 2, background: i <= (p.lettera || 0) ? "#e24b4a" : "#1e1e2a" }} />
              ))}
            </div>

            {p.dataUltimaLettera && (
              <div style={{ fontSize: 12, color: gg >= 20 ? "#4ecb8d" : "#EF9F27", marginBottom: 8 }}>
                {gg >= 20 ? `✓ Pronta per prossima lettera (${gg}gg)` : `⏳ Attendi ancora ${20 - gg} giorni`}
              </div>
            )}

            {p.esito && <div style={{ fontSize: 12, color: "#5a5868", marginBottom: 10, fontStyle: "italic" }}>Esito: {p.esito}</div>}

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {p.stato !== "AGENZIA" && (
                <select onChange={e => e.target.value && registraEsito(p.id, e.target.value)} defaultValue="" style={{ flex: 1, fontSize: 13, padding: "8px 10px" }}>
                  <option value="">Esito raccomandata...</option>
                  {ESITI_RACC.map(e => <option key={e} value={e}>{e}</option>)}
                </select>
              )}
              {p.stato !== "AGENZIA" && isPronta && (
                <button className="btn-danger" style={{ fontSize: 13, padding: "8px 14px", whiteSpace: "nowrap" }} onClick={() => avanzaLettera(p)}>
                  {p.lettera >= 3 ? "→ Agenzia" : `→ Lettera ${(p.lettera || 0) + 1}`}
                </button>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ─── INCASSI AGENTI ─── */
function Incassi({ incassi, clienti, onAdd, onUpdate, onSelectCliente, showToast }) {
  const [tab, setTab] = useState("mano");
  const [showForm, setShowForm] = useState(false);

  const inMano = incassi.filter(i => i.stato === "IN MANO AGENTE");
  const spediti = incassi.filter(i => i.stato === "SPEDITO");
  const registrati = incassi.filter(i => i.stato === "REGISTRATO");

  const totalePerAgente = useMemo(() => {
    const map = {};
    inMano.forEach(i => {
      if (!map[i.agente]) map[i.agente] = { agente: i.agente, totale: 0, count: 0, titoli: [] };
      map[i.agente].totale += i.incassoNetto || i.importo || 0;
      map[i.agente].count++;
      map[i.agente].titoli.push(i);
    });
    return Object.values(map).sort((a, b) => b.totale - a.totale);
  }, [inMano]);

  function marcaSpedito(id, tracking) {
    onUpdate(id, { stato: "SPEDITO", dataSpedizione: oggi(), tracking });
    showToast("Spedizione registrata");
  }

  function marcaRegistrato(id) {
    onUpdate(id, { stato: "REGISTRATO", dataRegistrazione: oggi() });
    showToast("Titolo registrato contabilmente");
  }

  const IncassoCard = ({ inc }) => {
    const [tracking, setTracking] = useState("");
    const gg = giorniDa(inc.dataRicezione);
    const isAlert = gg > 30;
    const isWarn = gg > 15 && !isAlert;

    return (
      <div className="card" style={{ marginBottom: 10, borderColor: isAlert ? "#3a2020" : isWarn ? "#3a2e08" : "#1e1e2a" }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4a7fd4" }} onClick={() => { const cl = clienti.find(c => c.id === inc.clienteId); if (cl) onSelectCliente(cl.id); }}>
              {inc.ragione}
            </div>
            <div style={{ fontSize: 11, color: "#5a5868" }}>{inc.agente} · {fmtData(inc.dataRicezione)} ({gg}gg fa)</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div className="mono" style={{ fontWeight: 600 }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(inc.incassoNetto || inc.importo)}</div>
            <div style={{ fontSize: 11, color: "#5a5868" }}>{inc.tipo}</div>
          </div>
        </div>
        {inc.spese > 0 && <div style={{ fontSize: 12, color: "#5a5868", marginBottom: 6 }}>Spese: €{inc.spese.toFixed(2)} · Totale: €{(inc.importo || 0).toFixed(2)}</div>}
        {inc.note && <div style={{ fontSize: 12, color: "#7a7888", marginBottom: 8, fontStyle: "italic" }}>{inc.note}</div>}
        {isAlert && <div style={{ fontSize: 12, color: "#e24b4a", marginBottom: 8, fontWeight: 500 }}>⚠ In mano da {gg} giorni — sollecita spedizione!</div>}
        {isWarn && <div style={{ fontSize: 12, color: "#EF9F27", marginBottom: 8 }}>⏳ In mano da {gg} giorni</div>}

        {inc.stato === "IN MANO AGENTE" && (
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            <input value={tracking} onChange={e => setTracking(e.target.value)} placeholder="Numero tracking..." style={{ flex: 1, fontSize: 13, padding: "8px 10px", background: "#0d0d14", border: "1px solid #2a2a38", borderRadius: 8, color: "#e8e6df" }} />
            <button className="btn-green" style={{ fontSize: 13, padding: "8px 14px", whiteSpace: "nowrap" }} onClick={() => marcaSpedito(inc.id, tracking)}>
              Spedito
            </button>
          </div>
        )}
        {inc.stato === "SPEDITO" && (
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
            <div style={{ fontSize: 12, color: "#4ecb8d" }}>✓ Spedito il {fmtData(inc.dataSpedizione)}{inc.tracking ? ` · ${inc.tracking}` : ""}</div>
            <button className="btn-ghost" style={{ fontSize: 12, padding: "6px 12px" }} onClick={() => marcaRegistrato(inc.id)}>Registrato</button>
          </div>
        )}
        {inc.stato === "REGISTRATO" && <div style={{ fontSize: 12, color: "#1D9E75", marginTop: 8 }}>✓ Registrato contabilmente il {fmtData(inc.dataRegistrazione)}</div>}
      </div>
    );
  };

  return (
    <div>
      <div className="section-title">Incassi agenti</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 16 }}>{inMano.length} titoli in mano · {spediti.length} spediti</div>

      <div className="tab-bar" style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto" }}>
        {[["mano", `In mano (${inMano.length})`], ["agente", "Per agente"], ["spediti", `Spediti (${spediti.length})`], ["tutti", "Tutti"]].map(([k, l]) => (
          <button key={k} onClick={() => setTab(k)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: tab === k ? "#4a7fd4" : "transparent", color: tab === k ? "#fff" : "#5a5868", border: tab === k ? "none" : "1px solid #2a2a38" }}>{l}</button>
        ))}
      </div>

      {tab === "mano" && (
        <div>
          {inMano.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo in mano agli agenti</div> : inMano.map(i => <IncassoCard key={i.id} inc={i} />)}
        </div>
      )}

      {tab === "agente" && (
        <div>
          {totalePerAgente.map(ag => (
            <div key={ag.agente} className="card" style={{ marginBottom: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ fontWeight: 600 }}>{ag.agente}</div>
                <div className="mono" style={{ fontWeight: 600, color: "#EF9F27" }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(ag.totale)}</div>
              </div>
              <div style={{ fontSize: 12, color: "#5a5868", marginBottom: 10 }}>{ag.count} titoli · media {new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(ag.totale / ag.count)}</div>
              {ag.titoli.map(i => (
                <div key={i.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: "1px solid #1a1a25", fontSize: 13 }}>
                  <span>{i.ragione} · {i.tipo}</span>
                  <span className="mono">{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(i.incassoNetto || i.importo)}</span>
                </div>
              ))}
            </div>
          ))}
          {totalePerAgente.length === 0 && <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo in mano</div>}
        </div>
      )}

      {tab === "spediti" && (
        <div>{spediti.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo spedito</div> : spediti.map(i => <IncassoCard key={i.id} inc={i} />)}</div>
      )}

      {tab === "tutti" && (
        <div>{incassi.map(i => <IncassoCard key={i.id} inc={i} />)}</div>
      )}
    </div>
  );
}

/* ─── DASHBOARD ─── */
function Dashboard({ kpi, clienti, praticheLl, incassi }) {
  const byStato = useMemo(() => {
    const map = {};
    STATI_OP.forEach(s => { map[s] = clienti.filter(c => c.statoOp === s).length; });
    return map;
  }, [clienti]);

  const topScaduti = [...clienti].filter(c => c.statoCl !== "PAGATO").sort((a, b) => (b.totaleScaduto || 0) - (a.totaleScaduto || 0)).slice(0, 5);

  return (
    <div>
      <div className="section-title">Dashboard</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 20 }}>{new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long" })}</div>

      {/* KPI PRINCIPALI */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
        <div className="card" style={{ gridColumn: "1 / -1" }}>
          <div style={{ fontSize: 11, color: "#5a5868", marginBottom: 6 }}>SCADUTO TOTALE</div>
          <div className="mono" style={{ fontSize: 26, fontWeight: 700, color: "#e24b4a" }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(kpi.totScaduto)}</div>
        </div>
        {[
          { label: "Pratiche L/L", val: kpi.llAperte, color: "#e24b4a" },
          { label: "Titoli in mano", val: kpi.titoliMano, color: "#EF9F27" },
          { label: "Richiami oggi", val: kpi.daRichiamare, color: "#378ADD" },
          { label: "Clienti attivi", val: clienti.filter(c => c.statoCl !== "PAGATO").length, color: "#4ecb8d" },
        ].map(k => (
          <div key={k.label} className="card">
            <div style={{ fontSize: 11, color: "#5a5868", marginBottom: 6 }}>{k.label.toUpperCase()}</div>
            <div className="mono" style={{ fontSize: 24, fontWeight: 700, color: k.color }}>{k.val}</div>
          </div>
        ))}
      </div>

      {/* DISTRIBUZIONE STATI */}
      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: "#9a98a0" }}>Distribuzione per stato</div>
        {STATI_OP.filter(s => byStato[s] > 0).map(s => (
          <div key={s} style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: STATO_COLORS[s], display: "inline-block" }} />
                {s}
              </span>
              <span className="mono" style={{ color: "#9a98a0" }}>{byStato[s]}</span>
            </div>
            <div style={{ background: "#1a1a25", borderRadius: 4, height: 4, overflow: "hidden" }}>
              <div style={{ width: `${(byStato[s] / clienti.length) * 100}%`, height: "100%", background: STATO_COLORS[s], borderRadius: 4 }} />
            </div>
          </div>
        ))}
      </div>

      {/* TOP SCADUTI */}
      <div className="card">
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: "#9a98a0" }}>Top 5 scaduti</div>
        {topScaduti.map((c, i) => (
          <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: i < 4 ? "1px solid #1a1a25" : "none" }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 500 }}>{c.ragione}</div>
              <div style={{ fontSize: 11, color: "#5a5868" }}>{c.agente}</div>
            </div>
            <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(c.totaleScaduto || 0)}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

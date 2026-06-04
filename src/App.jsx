import { useState, useEffect, useMemo } from "react";

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxM-HnsnxAehTsVTp_YxmVOKb4xuVAMSc33wzRH0mVhRclFkRrxG7-vxIDb7SQiLgg/exec";

// ─── TEMA ───
const THEMES = {
  dark: {
    bg: "#0a0a0f", bgCard: "#13131a", bgAlt: "#0d0d14", bgInput: "#13131a",
    border: "#1e1e2a", borderInput: "#2a2a38",
    text: "#e8e6df", textSub: "#5a5868", textMuted: "#4a4858",
    navBg: "#0d0d14", navBorder: "#1a1a25",
    diario: {
      "Mio contatto": { border: "#378ADD", bg: "#0a1828", label: "#378ADD" },
      "Via agente": { border: "#4ecb8d", bg: "#0a2018", label: "#4ecb8d" },
      "Nota personale": { border: "#5a5868", bg: "#0d0d14", label: "#7a7888" },
      "Cliente": { border: "#9a98a0", bg: "#121218", label: "#9a98a0" },
    }
  },
  light: {
    bg: "#f4f5f7", bgCard: "#ffffff", bgAlt: "#f0f1f3", bgInput: "#ffffff",
    border: "#e2e4e9", borderInput: "#d0d3db",
    text: "#1a1a2e", textSub: "#6b6f7d", textMuted: "#9a9fad",
    navBg: "#ffffff", navBorder: "#e2e4e9",
    diario: {
      "Mio contatto": { border: "#378ADD", bg: "#eaf3fd", label: "#2265b0" },
      "Via agente": { border: "#1D9E75", bg: "#e6f7f2", label: "#0f6e56" },
      "Nota personale": { border: "#9a9fad", bg: "#f4f5f7", label: "#6b6f7d" },
      "Cliente": { border: "#c0c3cb", bg: "#f9f9fb", label: "#9a9fad" },
    }
  }
};

const ASSET_COLORS = {};
const STATI_OP = ["CONTROLLA", "AGENTE", "AGENZIA", "L/L", "RICHIAMA", "PAGATO", "PROBLEMA"];
const PRIORITA = ["ALTA", "MEDIA", "BASSA"];
const TIPI_CONTATTO = ["Telefono", "Email", "WhatsApp", "Altro"];
const STATI_LL = ["IN CODA", "LETTERA 1", "LETTERA 2", "LETTERA 3", "AGENZIA"];
const ESITI_RACC = ["In attesa", "Consegnata", "Non consegnata", "Rifiutata", "Indirizzo errato"];
const TIPI_TITOLO = ["Assegno a vista", "Assegno postdatato", "Cambiale", "Contanti", "Bonifico"];
const TIPI_DIARIO = ["Mio contatto", "Via agente", "Nota personale", "Cliente"];
const STORAGE_KEY = "crm_solleciti_v2";

const STATO_COLORS = {
  "CONTROLLA": "#378ADD", "AGENTE": "#EF9F27", "AGENZIA": "#7F77DD",
  "L/L": "#e24b4a", "RICHIAMA": "#4ecb8d", "PAGATO": "#1D9E75", "PROBLEMA": "#FF4A8D"
};
const PRIORITA_COLORS = { "ALTA": "#e24b4a", "MEDIA": "#EF9F27", "BASSA": "#4ecb8d" };

// DIARIO_COLORS is now dynamic via theme

const IBAN = "IT79G0306909496100000011059";
const AZIENDA = "Saratoga Int. Sforza SPA";
const BANCA = "Banca Intesa San Paolo – Via Lorenteggio 70";

function loadLocal() {
  try {
    const r = localStorage.getItem(STORAGE_KEY);
    if (r) return JSON.parse(r);
    // Migra da v1 se esiste
    const v1 = localStorage.getItem("crm_solleciti_v1");
    if (v1) return { ...JSON.parse(v1), agenti: [], insoluti: [] };
    return { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [] };
  } catch { return { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [] }; }
}
function saveLocal(d) { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); }
async function loadFromSheets() { const r = await fetch(`${SCRIPT_URL}?action=load`); return r.json(); }
async function saveToSheets(d) { await fetch(`${SCRIPT_URL}?action=save&data=${encodeURIComponent(JSON.stringify(d))}`); }

function oggi() { return new Date().toISOString().slice(0, 10); }
function oraOra() { return new Date().toTimeString().slice(0, 5); }
function fmtEur(n) {
  if (n == null || isNaN(n)) return "—";
  return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(n);
}
function fmtData(d) {
  if (!d) return "—";
  if (typeof d === "number") return new Date((d - 25569) * 86400000).toLocaleDateString("it-IT");
  if (String(d).includes("-")) { const [y, m, dd] = String(d).split("-"); return `${dd}/${m}/${y}`; }
  return d;
}
function parseExcelDate(v) {
  if (!v) return null;
  if (typeof v === "number" && v > 40000) return new Date((v - 25569) * 86400000).toISOString().slice(0, 10);
  return String(v);
}
function giorniDa(dataStr) {
  if (!dataStr) return 9999;
  return Math.floor((new Date() - new Date(dataStr)) / 86400000);
}
function addGiorni(dataStr, n) {
  if (!dataStr) return "";
  const d = new Date(dataStr);
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

async function parseSpaccatureExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        // Prende il primo foglio utile (SCADUTO o simile)
        const sheetName = wb.SheetNames.find(s => s.toUpperCase().includes("SCADUT") || s.toUpperCase().includes("SPACCATURE")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        // Leggi le righe come array per gestire l'intestazione data
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (raw.length < 2) { resolve([]); return; }
        const headers = raw[0];
        // Trova l'indice della colonna scadenza (quella con una data nell'intestazione)
        const scadenzaIdx = headers.findIndex(h => h instanceof Date || (typeof h === "string" && /\d{2}\/\d{2}\/\d{4}/.test(h)));
        const rows = raw.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[String(h)] = row[i]; });
          if (scadenzaIdx >= 0) obj["__SCADENZA__"] = row[scadenzaIdx];
          return obj;
        });
        resolve(rows);
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function elaboraSpaccature(rows) {
  const clientiMap = {};
  rows.forEach(row => {
    const cod = String(row["COD."] || row["CODICE"] || "").trim().padStart(5, "0");
    if (!cod || cod === "00000") return;
    const ragione = String(row["RAGIONE SOCIALE"] || row["RAGIONE_SOCIALE"] || "").trim();
    const agente = String(row["AGENTE"] || "").trim();
    const localita = String(row["LOCALITA'"] || row["LOCALITA"] || "").trim();
    const prov = String(row["PROV."] || row["PROV"] || "").trim();
    const doc = String(row["DOC."] || row["DOC"] || "").trim();
    // Scadenza: usa la colonna trovata dinamicamente o la colonna __SCADENZA__
    const scadenzaRaw = row["__SCADENZA__"] || row["SCADENZE"] || row["SCADENZA"] || "";
    let scadenza = "";
    if (scadenzaRaw instanceof Date) {
      scadenza = scadenzaRaw.toISOString().slice(0, 10);
    } else if (typeof scadenzaRaw === "number" && scadenzaRaw > 40000) {
      scadenza = new Date((scadenzaRaw - 25569) * 86400000).toISOString().slice(0, 10);
    } else {
      scadenza = String(scadenzaRaw);
    }
    const importo = parseFloat(row["IMPORTO"] || 0);
    const giorni = parseInt(row["GIORNI"] || 0);
    const tag = `${cod}|${doc}|${scadenza}|${importo.toFixed(2)}`;
    if (!clientiMap[cod]) {
      clientiMap[cod] = {
        id: cod, codice: cod, ragione, agente, localita, prov,
        email: "", telefono: "", cellulare: "",
        orari: "", personeRif: "",
        statoOp: "CONTROLLA", statoCl: "ATTIVO", priorita: "ALTA",
        esito: "", ultimoContatto: null, dataRichiamo: null, ultimoGiro: null,
        note: "", diario: [], scadenze: [], totaleScaduto: 0, giorniMaxRitardo: 0,
      };
    }
    const cl = clientiMap[cod];
    if (ragione && !cl.ragione) cl.ragione = ragione;
    if (agente && !cl.agente) cl.agente = agente;
    const esiste = cl.scadenze.find(s => s.tag === tag);
    if (!esiste) cl.scadenze.push({ tag, doc, scadenza, importo, giorni, residuo: importo, stato: "APERTA", accorpata: false, fatturaAccorpante: null });
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
      esistente.ragione = nc.ragione || esistente.ragione;
      esistente.agente = nc.agente || esistente.agente;
      esistente.localita = nc.localita || esistente.localita;
      nc.scadenze.forEach(ns => {
        if (!esistente.scadenze.find(s => s.tag === ns.tag)) esistente.scadenze.push(ns);
      });
      esistente.totaleScaduto = esistente.scadenze.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
    } else {
      result.push(nc);
    }
  });
  return result;
}

function generaMailCliente(cliente) {
  const scadenzeAperte = (cliente.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  if (scadenzeAperte.length === 0 || !cliente.email) return null;
  let corpo = `Gentile Cliente,\n\na seguito di controlli contabili risultano ancora in sospeso i pagamenti relativi alle seguenti fatture:\n\n`;
  scadenzeAperte.forEach(s => { corpo += `• Fattura n. ${s.doc} scaduta il ${fmtData(s.scadenza)} – € ${(s.residuo || s.importo).toFixed(2)}\n\n`; });
  corpo += `Al fine di agevolare la procedura di regolarizzazione, riportiamo di seguito le coordinate per il pagamento.\n\nBonifico bancario:\n${IBAN}\n${AZIENDA}\n${BANCA}\n\nQualora abbiate già provveduto al pagamento ritenete nullo il presente sollecito.\n\nRestiamo a Sua completa disposizione per ulteriori chiarimenti.\n\nRingraziando per l'attenzione, porgiamo cordiali saluti.`;
  return `mailto:${cliente.email}?subject=${encodeURIComponent("Verifica stato pagamento fatture")}&body=${encodeURIComponent(corpo)}`;
}

function generaMailInsolutoCliente(cliente, insoluto) {
  if (!cliente.email) return null;
  const corpo = `Gentile Cliente,\n\nLa informiamo che l'assegno n. ${insoluto.riferimento || "—"} dell'importo di € ${(insoluto.importo || 0).toFixed(2)} emesso in data ${fmtData(insoluto.dataEmissione)} è tornato insoluto per ${insoluto.motivo || "mancanza fondi"}.\n\nLa preghiamo di provvedere alla regolarizzazione entro 60 giorni dalla presente comunicazione.\n\nBonifico bancario:\n${IBAN}\n${AZIENDA}\n${BANCA}\n\nPer ulteriori informazioni restiamo a disposizione.\n\nCordiali saluti.`;
  return `mailto:${cliente.email}?subject=${encodeURIComponent("Assegno insoluto — regolarizzazione entro 60 giorni")}&body=${encodeURIComponent(corpo)}`;
}

function generaMailInsolutoAgente(agente, cliente, insoluto) {
  if (!agente?.email) return null;
  const corpo = `Gentile ${agente.nome},\n\nTi informo che l'assegno del cliente ${cliente.ragione} (cod. ${cliente.codice}) dell'importo di € ${(insoluto.importo || 0).toFixed(2)} del ${fmtData(insoluto.dataEmissione)} è tornato insoluto per ${insoluto.motivo || "mancanza fondi"}.\n\nTi chiedo di non accettare ulteriori titoli da questo cliente fino a completa regolarizzazione.\n\nGrazie per la collaborazione.\n\nCordiali saluti.`;
  return `mailto:${agente.email}?subject=${encodeURIComponent(`Segnalazione insoluto cliente ${cliente.ragione} — ${cliente.codice}`)}&body=${encodeURIComponent(corpo)}`;
}

export default function App() {
  const [page, setPage] = useState("oggi");
  const [data, setData] = useState(loadLocal);
  const [toast, setToast] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [loadingCloud, setLoadingCloud] = useState(true);
  const [clienteSelezionato, setClienteSelezionato] = useState(null);
  const [agenteSelezionato, setAgenteSelezionato] = useState(null);
  const [themeMode, setThemeMode] = useState(() => localStorage.getItem("crm_theme") || "dark");
  const T = THEMES[themeMode];
  function toggleTheme() {
    const next = themeMode === "dark" ? "light" : "dark";
    setThemeMode(next);
    localStorage.setItem("crm_theme", next);
  }

  useEffect(() => {
    loadFromSheets()
      .then(remote => {
        if (remote && !remote.error && (remote.clienti?.length > 0 || remote.agenti?.length > 0)) {
          const merged = { clienti: [], praticheLl: [], incassi: [], agenti: [], insoluti: [], ...remote };
          setData(merged); saveLocal(merged);
        }
      })
      .catch(() => {})
      .finally(() => setLoadingCloud(false));
  }, []);

  useEffect(() => { saveLocal(data); }, [data]);

  function showToast(msg, type = "ok") { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); }

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
      const nd = { ...d, clienti: d.clienti.map(c => c.id === id ? { ...c, diario: [voce, ...(c.diario || [])], ultimoContatto: voce.tipo !== "Nota personale" ? voce.data : c.ultimoContatto } : c) };
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
        // Auto-crea agenti mancanti
        const agentiEsistenti = d.agenti || [];
        const nomiAgentiNuovi = [...new Set(nuovi.map(c => c.agente).filter(Boolean))];
        const agentiDaAggiungere = nomiAgentiNuovi
          .filter(nome => !agentiEsistenti.find(a => a.nome === nome))
          .map(nome => ({ id: Date.now() + Math.random(), nome, email: "", telefono: "", zona: "" }));
        const agentiMerged = [...agentiEsistenti, ...agentiDaAggiungere];
        const totaleScadutoImport = clientiMerged.filter(c => c.statoCl !== "PAGATO").reduce((s, c) => s + (c.totaleScaduto || 0), 0);
        const dataImport = oggi();
        const nd = { ...d, clienti: clientiMerged, agenti: agentiMerged, ultimoImport: { data: dataImport, scaduto: totaleScadutoImport } };
        syncToCloud(nd);
        showToast(`${clientiMerged.length} clienti · ${agentiDaAggiungere.length} nuovi agenti creati`);
        return nd;
      });
    } catch (e) { showToast("Errore: " + e.message, "warn"); }
  }

  function addPraticaLL(clienteId) {
    const cliente = data.clienti.find(c => c.id === clienteId);
    if (!cliente) return;
    setData(d => {
      const nd = { ...d, praticheLl: [...(d.praticheLl || []), { id: Date.now(), clienteId, ragione: cliente.ragione, agente: cliente.agente, stato: "IN CODA", lettera: 0, dataDecisione: oggi(), dataUltimaLettera: null, esito: "", note: "", storia: [] }] };
      showToast(`Pratica L/L aperta per ${cliente.ragione}`); syncToCloud(nd); return nd;
    });
  }

  function updateLL(id, updates) {
    setData(d => { const nd = { ...d, praticheLl: d.praticheLl.map(p => p.id === id ? { ...p, ...updates } : p) }; syncToCloud(nd); return nd; });
  }

  function addIncasso(inc) {
    const statoBonifico = inc.tipo === "Bonifico" ? "REGISTRATO" : "IN MANO AGENTE";
    setData(d => {
      const nd = { ...d, incassi: [{ ...inc, id: Date.now(), dataRicezione: oggi(), dataRegistrazione: inc.tipo === "Bonifico" ? oggi() : null, stato: statoBonifico }, ...(d.incassi || [])] };
      showToast(inc.tipo === "Bonifico" ? "Bonifico registrato come pagato" : "Incasso registrato");
      syncToCloud(nd); return nd;
    });
  }

  function updateIncasso(id, updates) {
    setData(d => { const nd = { ...d, incassi: d.incassi.map(i => i.id === id ? { ...i, ...updates } : i) }; syncToCloud(nd); return nd; });
  }

  function deleteIncasso(id) {
    setData(d => { const nd = { ...d, incassi: d.incassi.filter(i => i.id !== id) }; showToast("Incasso eliminato"); syncToCloud(nd); return nd; });
  }

  function addAgente(agente) {
    setData(d => { const nd = { ...d, agenti: [...(d.agenti || []), { ...agente, id: Date.now() }] }; showToast("Agente aggiunto"); syncToCloud(nd); return nd; });
  }

  function updateAgente(id, updates) {
    setData(d => { const nd = { ...d, agenti: d.agenti.map(a => a.id === id ? { ...a, ...updates } : a) }; syncToCloud(nd); return nd; });
  }

  function addInsoluto(ins) {
    const cliente = data.clienti.find(c => c.id === ins.clienteId);
    setData(d => {
      const nuovoIns = { ...ins, id: Date.now(), data: oggi(), stato: "APERTO", scadenza60gg: addGiorni(oggi(), 60) };
      let clienti = d.clienti;
      if (cliente) {
        clienti = d.clienti.map(c => c.id === ins.clienteId ? { ...c, statoOp: "PROBLEMA" } : c);
      }
      const nd = { ...d, clienti, insoluti: [nuovoIns, ...(d.insoluti || [])] };
      showToast("Insoluto registrato — cliente → PROBLEMA"); syncToCloud(nd); return nd;
    });
  }

  function updateInsoluto(id, updates) {
    setData(d => { const nd = { ...d, insoluti: d.insoluti.map(i => i.id === id ? { ...i, ...updates } : i) }; syncToCloud(nd); return nd; });
  }

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
    const incassi = data.incassi || [];
    const meseCorrente = oggi().slice(0, 7); // "YYYY-MM"

    const totScaduto = clienti.filter(c => c.statoCl !== "PAGATO").reduce((s, c) => s + (c.totaleScaduto || 0), 0);
    const scadutoMese = data.ultimoImport?.scaduto || totScaduto;

    // Incassato confermato (REGISTRATO)
    const incassatoConfermatoTot = incassi.filter(i => i.stato === "REGISTRATO").reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const incassatoConfermatoMese = incassi.filter(i => i.stato === "REGISTRATO" && (i.dataRegistrazione || "").startsWith(meseCorrente)).reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);

    // Incassato atteso (IN MANO AGENTE + SPEDITO)
    const incassatoAtteso = incassi.filter(i => i.stato === "IN MANO AGENTE" || i.stato === "SPEDITO").reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
    const incassatoAttesoMese = incassi.filter(i => (i.stato === "IN MANO AGENTE" || i.stato === "SPEDITO") && (i.dataRicezione || "").startsWith(meseCorrente)).reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);

    // Totale incassato
    const incassatoTotale = incassatoConfermatoTot + incassatoAtteso;
    const incassatoTotaleMese = incassatoConfermatoMese + incassatoAttesoMese;

    // Percentuali
    const pctConfermatoTot = totScaduto > 0 ? (incassatoConfermatoTot / totScaduto) * 100 : 0;
    const pctTotaleMese = scadutoMese > 0 ? (incassatoTotaleMese / scadutoMese) * 100 : 0;

    const llAperte = (data.praticheLl || []).filter(p => p.stato !== "AGENZIA" && p.stato !== "CHIUSA").length;
    const titoliMano = incassi.filter(i => i.stato === "IN MANO AGENTE").length;
    const daRichiamare = listaOggi.filter(c => c.dataRichiamo === oggi()).length;
    const insoluti = (data.insoluti || []).filter(i => i.stato === "APERTO").length;

    return {
      totScaduto, scadutoMese,
      incassatoConfermatoTot, incassatoConfermatoMese,
      incassatoAtteso, incassatoAttesoMese,
      incassatoTotale, incassatoTotaleMese,
      pctConfermatoTot, pctTotaleMese,
      llAperte, titoliMano, daRichiamare, insoluti,
      dataUltimoImport: data.ultimoImport?.data || null
    };
  }, [data.clienti, data.praticheLl, data.incassi, data.insoluti, data.ultimoImport, listaOggi]);

  const NAV = [
    { id: "oggi", label: "Oggi", icon: "ti-calendar-event" },
    { id: "clienti", label: "Clienti", icon: "ti-users" },
    { id: "agenti", label: "Agenti", icon: "ti-briefcase" },
    { id: "ll", label: "L/L", icon: "ti-file-text" },
    { id: "incassi", label: "Incassi", icon: "ti-coin" },
    { id: "dashboard", label: "KPI", icon: "ti-chart-bar" },
  ];

  if (clienteSelezionato) {
    const cl = data.clienti.find(c => c.id === clienteSelezionato);
    if (cl) return <SchedaCliente cliente={cl} agenti={data.agenti || []} themeMode={themeMode} onBack={() => setClienteSelezionato(null)} onUpdate={u => updateCliente(cl.id, u)} onAddDiario={v => addDiario(cl.id, v)} onApriLL={() => { addPraticaLL(cl.id); setClienteSelezionato(null); setPage("ll"); }} onAddIncasso={addIncasso} onAddInsoluto={addInsoluto} showToast={showToast} />;
  }

  if (agenteSelezionato) {
    const ag = data.agenti.find(a => a.id === agenteSelezionato) || data.agenti.find(a => a.nome === agenteSelezionato) || { id: null, nome: agenteSelezionato, email: "", telefono: "", zona: "" };
    const clientiAgente = (data.clienti || []).filter(c => c.agente === ag?.nome);
    const incassiAgente = (data.incassi || []).filter(i => i.agente === ag?.nome && i.stato === "IN MANO AGENTE");
    if (ag) return <SchedaAgente agente={ag} clienti={clientiAgente} incassi={incassiAgente} themeMode={themeMode} onBack={() => setAgenteSelezionato(null)} onUpdate={u => updateAgente(ag.id, u)} onAddDiario={(clienteId, voce) => addDiario(clienteId, voce)} onUpdateCliente={updateCliente} onUpdateIncasso={updateIncasso} onDeleteIncasso={deleteIncasso} onSelectCliente={id => { setAgenteSelezionato(null); setClienteSelezionato(id); }} showToast={showToast} />;
  }

  return (
    <div style={{ minHeight: "100vh", background: T.bg, color: T.text, fontFamily: "'DM Sans','Segoe UI',sans-serif", paddingBottom: 72 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        @import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');
        :root {
          --bg: ${T.bg};
          --bg-card: ${T.bgCard};
          --bg-input: ${T.bgInput};
          --border: ${T.border};
          --border-input: ${T.borderInput};
          --text: ${T.text};
          --text-sub: ${T.textSub};
          --text-muted: ${T.textMuted};
        }
        *{box-sizing:border-box;margin:0;padding:0}
        input,select,textarea{background:var(--bg-input);border:1px solid var(--border-input);border-radius:10px;color:var(--text);padding:10px 14px;font-family:inherit;font-size:15px;width:100%;outline:none;-webkit-appearance:none}
        input:focus,select:focus,textarea:focus{border-color:#4a7fd4}
        select option{background:var(--bg-input)}
        textarea{resize:vertical;min-height:80px}
        button{cursor:pointer;font-family:inherit;font-size:14px;border:none;border-radius:10px;padding:10px 18px;transition:all .15s;-webkit-tap-highlight-color:transparent}
        .btn-primary{background:#4a7fd4;color:#fff;font-weight:500}
        .btn-ghost{background:transparent;color:var(--text-sub);border:1px solid var(--border-input)}
        .btn-danger{background:transparent;color:#e24b4a;border:1px solid #3a2020}
        .btn-green{background:#0f2a1a;color:#4ecb8d;border:1px solid #1a4a2a}
        .card{background:var(--bg-card);border:1px solid var(--border);border-radius:16px;padding:16px}
        .tag{display:inline-block;font-size:11px;font-weight:500;padding:3px 9px;border-radius:20px}
        label{font-size:12px;color:var(--text-sub);margin-bottom:5px;display:block;letter-spacing:.04em}
        .modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:flex-end;justify-content:center;z-index:200}
        .modal{background:var(--bg-card);border:1px solid var(--border-input);border-radius:20px 20px 0 0;padding:24px 20px 40px;width:100%;max-width:500px;max-height:92vh;overflow-y:auto}
        .mono{font-family:'DM Mono',monospace}
        @keyframes spin{to{transform:rotate(360deg)}}
        .spin{animation:spin 1s linear infinite;display:inline-block}
        .bottom-nav{position:fixed;bottom:0;left:0;right:0;display:flex;z-index:100;padding-bottom:env(safe-area-inset-bottom)}
        .nav-btn{flex:1;display:flex;flex-direction:column;align-items:center;gap:2px;padding:8px 4px;background:none;border:none;border-radius:0;font-size:9px;letter-spacing:.04em}
        .nav-btn.active{color:#4a7fd4}
        .nav-btn:not(.active){color:var(--text-muted)}
        .nav-icon{font-size:20px}
        .section-title{font-size:22px;font-weight:600;letter-spacing:-.02em;margin-bottom:4px;color:var(--text)}
        .form-group{margin-bottom:14px}
        .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
        .cl-card{background:var(--bg-card);border:1px solid var(--border);border-radius:14px;padding:14px;margin-bottom:10px;cursor:pointer;transition:border-color .15s}
        .cl-card:active{border-color:#4a7fd4}
        .priorita-bar{width:4px;border-radius:2px;align-self:stretch;flex-shrink:0}
        .tab-bar{display:flex;gap:8px;margin-bottom:16px;overflow-x:auto;padding-bottom:2px}
        .tab{padding:8px 16px;border-radius:20px;font-size:13px;border:1px solid #2a2a38;background:transparent;color:#5a5868;white-space:nowrap}
        .tab.active{background:#4a7fd4;color:#fff;border-color:transparent}
        .search-wrap{position:relative;margin-bottom:16px}
        .search-input{background:#13131a;border:1px solid #2a2a38;border-radius:12px;color:#e8e6df;padding:10px 14px 10px 38px;font-size:15px;width:100%;outline:none}
        .search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#5a5868;font-size:18px}
        .kv{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #1a1a25;font-size:13px}
        .kv:last-child{border-bottom:none}
        .warn-box{background:#1f1808;border:1px solid #3a2e08;border-radius:10px;padding:10px 14px;font-size:13px;color:#EF9F27;margin-bottom:14px}
        .info-box{background:#0a1f2a;border:1px solid #1a3a4a;border-radius:10px;padding:10px 14px;font-size:13px;color:#378ADD;margin-bottom:14px}
        .danger-box{background:#2a0a0a;border:1px solid #4a1a1a;border-radius:10px;padding:10px 14px;font-size:13px;color:#e24b4a;margin-bottom:14px}
      `}</style>

      <div style={{ padding: "52px 20px 0", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 13, color: T.textMuted, letterSpacing: ".06em" }}>CRM SOLLECITI</div>
          <div style={{ fontSize: 11, marginTop: 2, color: syncing ? "#378ADD" : T.textMuted }}>
            {syncing ? <><span className="spin">↻</span> Sync...</> : "● " + new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "short" })}
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end", alignItems: "center" }}>
          <button onClick={toggleTheme} style={{ background: "transparent", border: `1px solid ${T.borderInput}`, borderRadius: 20, padding: "5px 12px", fontSize: 16, cursor: "pointer", color: T.textSub }}>{themeMode === "dark" ? "☀️" : "🌙"}</button>
          {kpi.daRichiamare > 0 && <span style={{ background: "#1f1808", border: "1px solid #4a3010", color: "#EF9F27", borderRadius: 8, padding: "2px 8px", fontSize: 11 }}>📅 {kpi.daRichiamare}</span>}
          {kpi.insoluti > 0 && <span style={{ background: "#2a0a0a", border: "1px solid #4a1a1a", color: "#e24b4a", borderRadius: 8, padding: "2px 8px", fontSize: 11 }}>⚠ {kpi.insoluti} insoluti</span>}
          {kpi.titoliMano > 0 && <span style={{ background: "#0a1f2a", border: "1px solid #1a3a4a", color: "#378ADD", borderRadius: 8, padding: "2px 8px", fontSize: 11 }}>💼 {kpi.titoliMano}</span>}
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
            {page === "oggi" && <Oggi lista={listaOggi} insoluti={data.insoluti || []} onSelectCliente={setClienteSelezionato} />}
            {page === "clienti" && <Clienti clienti={data.clienti || []} onSelectCliente={setClienteSelezionato} onImporta={importaSpaccature} showToast={showToast} />}
            {page === "agenti" && <Agenti agenti={data.agenti || []} clienti={data.clienti || []} incassi={data.incassi || []} onAddAgente={addAgente} onSelectAgente={setAgenteSelezionato} />}
            {page === "ll" && <PraticheLl pratiche={data.praticheLl || []} clienti={data.clienti || []} onUpdate={updateLL} onSelectCliente={setClienteSelezionato} showToast={showToast} />}
            {page === "incassi" && <Incassi incassi={data.incassi || []} clienti={data.clienti || []} onAdd={addIncasso} onUpdate={updateIncasso} onDelete={deleteIncasso} onSelectCliente={setClienteSelezionato} showToast={showToast} />}
            {page === "dashboard" && <Dashboard kpi={kpi} clienti={data.clienti || []} praticheLl={data.praticheLl || []} incassi={data.incassi || []} agenti={data.agenti || []} insoluti={data.insoluti || []} onSelectAgente={setAgenteSelezionato} />}
          </>
        )}
      </div>

      <nav className="bottom-nav" style={{ background: T.navBg, borderTop: `1px solid ${T.navBorder}` }}>
        {NAV.map(n => (
          <button key={n.id} className={`nav-btn ${page === n.id ? "active" : ""}`} onClick={() => setPage(n.id)}>
            <i className={`ti ${n.icon} nav-icon`} />
            {n.label}
          </button>
        ))}
      </nav>

      {toast && (
        <div style={{ position: "fixed", bottom: 88, left: 20, right: 20, background: toast.type === "warn" ? "#2a1f0a" : toast.type === "info" ? "#0a1f2a" : "#0f2a1a", border: `1px solid ${toast.type === "warn" ? "#4a3010" : toast.type === "info" ? "#1a3a4a" : "#1a4a2a"}`, color: toast.type === "warn" ? "#EF9F27" : toast.type === "info" ? "#378ADD" : "#4ecb8d", padding: "12px 18px", borderRadius: 12, fontSize: 14, zIndex: 300, textAlign: "center" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

/* ─── OGGI ─── */
function Oggi({ lista, insoluti, onSelectCliente }) {
  const oggi_ = oggi();
  const insolutiAperti = insoluti.filter(i => i.stato === "APERTO");
  const richiami = lista.filter(c => c.dataRichiamo === oggi_);
  const problemi = lista.filter(c => c.statoOp === "PROBLEMA" && c.dataRichiamo !== oggi_);
  const maiChiamati = lista.filter(c => !c.ultimoContatto && c.dataRichiamo !== oggi_ && c.statoOp !== "PROBLEMA");
  const vecchi = lista.filter(c => c.ultimoContatto && c.dataRichiamo !== oggi_ && c.statoOp !== "PROBLEMA");

  const Gruppo = ({ label, items, colore }) => items.length === 0 ? null : (
    <div style={{ marginBottom: 20 }}>
      <div style={{ fontSize: 11, color: colore, letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 10, fontWeight: 500 }}>{label} ({items.length})</div>
      {items.map(c => <CardCliente key={c.id} cliente={c} onClick={() => onSelectCliente(c.id)} />)}
    </div>
  );

  return (
    <div>
      <div className="section-title">Oggi</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 20 }}>{lista.length} clienti attivi</div>
      {insolutiAperti.length > 0 && (
        <div className="danger-box" style={{ marginBottom: 16 }}>
          ⚠ {insolutiAperti.length} assegni insoluti aperti — verifica scadenze 60gg
        </div>
      )}
      <Gruppo label="📅 Richiami programmati" items={richiami} colore="#EF9F27" />
      <Gruppo label="⚠ Problema / Insoluto" items={problemi} colore="#e24b4a" />
      <Gruppo label="🆕 Mai contattati" items={maiChiamati} colore="#e24b4a" />
      <Gruppo label="📞 Da richiamare" items={vecchi} colore="#9a98a0" />
      {lista.length === 0 && (
        <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>
          <i className="ti ti-calendar-off" style={{ fontSize: 40, marginBottom: 12, display: "block" }} />
          <div style={{ fontSize: 14 }}>Nessun cliente ancora</div>
          <div style={{ fontSize: 13, marginTop: 8 }}>Importa le spaccature dalla sezione Clienti</div>
        </div>
      )}
    </div>
  );
}

function CardCliente({ cliente, onClick }) {
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
            <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
            <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888", marginTop: 4, display: "inline-block" }}>{c.statoOp}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {c.ultimoContatto ? <span style={{ fontSize: 11, color: "#5a5868" }}>Ultimo: {fmtData(c.ultimoContatto)} ({giorni < 9999 ? `${giorni}gg fa` : "—"})</span> : <span style={{ fontSize: 11, color: "#e24b4a" }}>Mai contattato</span>}
          {c.dataRichiamo === oggi() && <span style={{ background: "#1f1808", border: "1px solid #4a3010", color: "#EF9F27", borderRadius: 8, padding: "2px 8px", fontSize: 11 }}>⏰ oggi</span>}
          {c.esito && <span style={{ fontSize: 11, color: "#7a7888", fontStyle: "italic" }}>"{c.esito.slice(0, 40)}{c.esito.length > 40 ? "…" : ""}"</span>}
        </div>
      </div>
    </div>
  );
}

/* ─── CLIENTI ─── */
function Clienti({ clienti, onSelectCliente, onImporta }) {
  const [search, setSearch] = useState("");
  const [filtroStato, setFiltroStato] = useState("TUTTI");
  const [filtroAgente, setFiltroAgente] = useState("TUTTI");
  const agenti = useMemo(() => ["TUTTI", ...new Set(clienti.map(c => c.agente).filter(Boolean))].sort(), [clienti]);
  const filtrati = useMemo(() => clienti.filter(c => {
    if (filtroStato !== "TUTTI" && c.statoOp !== filtroStato) return false;
    if (filtroAgente !== "TUTTI" && c.agente !== filtroAgente) return false;
    if (search && !c.ragione?.toLowerCase().includes(search.toLowerCase()) && !c.codice?.includes(search)) return false;
    return true;
  }).sort((a, b) => { if (a.agente < b.agente) return -1; if (a.agente > b.agente) return 1; return (b.totaleScaduto || 0) - (a.totaleScaduto || 0); }), [clienti, search, filtroStato, filtroAgente]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div><div className="section-title">Clienti</div><div style={{ color: "#5a5868", fontSize: 13 }}>{clienti.length} totali · {filtrati.length} mostrati</div></div>
        <label style={{ background: "#4a7fd4", color: "#fff", borderRadius: 10, padding: "10px 14px", fontSize: 14, fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>
          <i className="ti ti-upload" /> Importa
          <input type="file" accept=".xlsx,.xls,.xlsm" style={{ display: "none" }} onChange={e => e.target.files[0] && onImporta(e.target.files[0])} />
        </label>
      </div>
      <div className="search-wrap"><i className="ti ti-search search-icon" /><input className="search-input" placeholder="Cerca nome o codice..." value={search} onChange={e => setSearch(e.target.value)} /></div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12, overflowX: "auto", paddingBottom: 4 }}>
        {["TUTTI", ...STATI_OP].map(s => <button key={s} onClick={() => setFiltroStato(s)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: filtroStato === s ? "#4a7fd4" : "transparent", color: filtroStato === s ? "#fff" : "#5a5868", border: filtroStato === s ? "none" : "1px solid #2a2a38" }}>{s}</button>)}
      </div>
      <div className="form-group"><select value={filtroAgente} onChange={e => setFiltroAgente(e.target.value)}>{agenti.map(a => <option key={a}>{a}</option>)}</select></div>
      {filtrati.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}><div style={{ fontSize: 13 }}>Nessun cliente trovato</div></div> : filtrati.map(c => <CardCliente key={c.id} cliente={c} onClick={() => onSelectCliente(c.id)} />)}
    </div>
  );
}


/* ─── SHARED THEME STYLES ─── */
function getSharedStyles(T) {
  return `
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
    @import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');
    *{box-sizing:border-box;margin:0;padding:0}
    body{background:${T.bg};color:${T.text}}
    input,select,textarea{background:${T.bgInput};border:1px solid ${T.borderInput};border-radius:10px;color:${T.text};padding:10px 14px;font-family:inherit;font-size:15px;width:100%;outline:none;-webkit-appearance:none}
    input:focus,select:focus,textarea:focus{border-color:#4a7fd4}
    select option{background:${T.bgInput}}
    textarea{resize:vertical;min-height:80px}
    button{cursor:pointer;font-family:inherit;font-size:14px;border:none;border-radius:10px;padding:10px 18px;transition:all .15s}
    .btn-primary{background:#4a7fd4;color:#fff;font-weight:500}
    .btn-ghost{background:transparent;color:${T.textSub};border:1px solid ${T.borderInput}}
    .btn-danger{background:transparent;color:#e24b4a;border:1px solid #3a2020}
    .btn-green{background:#0f2a1a;color:#4ecb8d;border:1px solid #1a4a2a}
    .card{background:${T.bgCard};border:1px solid ${T.border};border-radius:16px;padding:16px}
    .tag{display:inline-block;font-size:11px;font-weight:500;padding:3px 9px;border-radius:20px}
    label{font-size:12px;color:${T.textSub};margin-bottom:5px;display:block;letter-spacing:.04em}
    .modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:flex-end;justify-content:center;z-index:200}
    .modal{background:${T.bgCard};border:1px solid ${T.borderInput};border-radius:20px 20px 0 0;padding:24px 20px 40px;width:100%;max-width:500px;max-height:92vh;overflow-y:auto}
    .mono{font-family:'DM Mono',monospace}
    .form-group{margin-bottom:14px}
    .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .tab-bar{display:flex;gap:8px;margin-bottom:16px;overflow-x:auto}
    .tab{padding:8px 16px;border-radius:20px;font-size:13px;border:1px solid ${T.borderInput};background:transparent;color:${T.textSub};white-space:nowrap}
    .tab.active{background:#4a7fd4;color:#fff;border-color:transparent}
    .kv{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid ${T.border};font-size:13px}
    .kv:last-child{border-bottom:none}
    .warn-box{background:#1f1808;border:1px solid #3a2e08;border-radius:10px;padding:10px 14px;font-size:13px;color:#EF9F27;margin-bottom:14px}
    .info-box{background:#0a1f2a;border:1px solid #1a3a4a;border-radius:10px;padding:10px 14px;font-size:13px;color:#378ADD;margin-bottom:14px}
    .scad-row{background:${T.bgAlt};border-radius:10px;padding:10px 12px;margin-bottom:8px}
  `;
}

/* ─── SCHEDA CLIENTE ─── */
function SchedaCliente({ cliente, agenti, themeMode, onBack, onUpdate, onAddDiario, onApriLL, onAddIncasso, onAddInsoluto, showToast }) {
  const T = THEMES[themeMode] || THEMES.dark;
  const c = cliente;
  const [tab, setTab] = useState("info");
  const [showContatto, setShowContatto] = useState(false);
  const [showIncasso, setShowIncasso] = useState(false);
  const [showAccorpa, setShowAccorpa] = useState(false);
  const [showInsoluto, setShowInsoluto] = useState(false);
  const [filtroTipo, setFiltroTipo] = useState("TUTTI");

  const scadenzeAperte = (c.scadenze || []).filter(s => s.stato !== "PAGATA" && !s.accorpata);
  const mailLink = generaMailCliente(c);

  const diariFiltrati = useMemo(() => {
    const d = c.diario || [];
    if (filtroTipo === "TUTTI") return d;
    return d.filter(v => v.tipo === filtroTipo);
  }, [c.diario, filtroTipo]);

  return (
    <div style={{ minHeight: "100vh", background: T.bg, color: T.text, fontFamily: "'DM Sans',sans-serif", paddingBottom: 40 }}>
      <style>{getSharedStyles(T)}</style>

      <div style={{ padding: "52px 20px 0", marginBottom: 16 }}>
        <button className="btn-ghost" style={{ marginBottom: 16, padding: "8px 14px", fontSize: 13 }} onClick={onBack}><i className="ti ti-arrow-left" /> Torna</button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>{c.ragione}</div>
            <div style={{ fontSize: 12, color: "#5a5868" }}>{c.codice} · {c.agente}</div>
            <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
              <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888" }}>{c.statoOp}</span>
              <span className="tag" style={{ background: (PRIORITA_COLORS[c.priorita] || "#888") + "22", color: PRIORITA_COLORS[c.priorita] || "#888" }}>{c.priorita}</span>
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div className="mono" style={{ fontSize: 22, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
            <div style={{ fontSize: 11, color: "#5a5868", marginTop: 2 }}>scaduto</div>
          </div>
        </div>
      </div>

      <div style={{ padding: "0 20px" }}>
        {/* AZIONI RAPIDE */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 16 }}>
          <button className="btn-primary" style={{ padding: "12px 10px" }} onClick={() => setShowContatto(true)}><i className="ti ti-phone" /> Contatto</button>
          {mailLink ? <a href={mailLink} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#0a1f2a", color: "#378ADD", border: "1px solid #1a3a4a", borderRadius: 10, padding: "12px 10px", fontSize: 14, textDecoration: "none" }}><i className="ti ti-mail" /> Mail</a> : <button className="btn-ghost" disabled style={{ opacity: 0.4 }}><i className="ti ti-mail" /> Mail</button>}
          <button className="btn-ghost" onClick={() => setShowIncasso(true)}><i className="ti ti-coin" /> Incasso</button>
          <button className="btn-danger" style={{ fontSize: 13 }} onClick={() => setShowInsoluto(true)}><i className="ti ti-alert-triangle" /> Insoluto</button>
        </div>

        {c.dataRichiamo && <div className="warn-box">⏰ Richiamo: {fmtData(c.dataRichiamo)}</div>}
        {c.esito && <div style={{ background: "#13131a", border: "1px solid #2a2a38", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#9a98a0" }}><span style={{ fontSize: 11, color: "#5a5868" }}>ULTIMO ESITO · </span>{c.esito}{c.ultimoContatto && <span style={{ fontSize: 11, color: "#5a5868" }}> · {fmtData(c.ultimoContatto)}</span>}</div>}

        <div className="tab-bar">
          {[["info", "Info"], ["scadenze", `Scad. (${scadenzeAperte.length})`], ["diario", `Diario (${(c.diario || []).length})`]].map(([k, l]) => (
            <button key={k} className={`tab ${tab === k ? "active" : ""}`} onClick={() => setTab(k)}>{l}</button>
          ))}
        </div>

        {tab === "info" && (
          <div>
            <div className="card" style={{ marginBottom: 12 }}>
              {[["Email", c.email || "—"], ["Telefono", c.telefono || "—"], ["Cellulare", c.cellulare || "—"], ["Orari", c.orari || "—"], ["Riferimento", c.personeRif || "—"], ["Agente", c.agente || "—"], ["Località", `${c.localita || "—"} (${c.prov || "—"})`], ["Max ritardo", c.giorniMaxRitardo ? `${c.giorniMaxRitardo}gg` : "—"]].map(([k, v]) => (
                <div key={k} className="kv"><span style={{ color: "#5a5868" }}>{k}</span><span style={{ fontWeight: 500 }}>{v}</span></div>
              ))}
            </div>
            <button className="btn-danger" style={{ width: "100%", fontSize: 13 }} onClick={onApriLL}><i className="ti ti-file-text" /> Apri pratica L/L</button>
          </div>
        )}

        {tab === "scadenze" && (
          <div>
            <button className="btn-ghost" style={{ width: "100%", marginBottom: 12, fontSize: 13 }} onClick={() => setShowAccorpa(true)}><i className="ti ti-arrows-join" /> Gestisci pagamento / accorpa</button>
            {scadenzeAperte.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 30, color: "#4a4858" }}>Nessuna scadenza aperta</div> : scadenzeAperte.map((s, i) => (
              <div key={i} className="scad-row">
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ fontWeight: 500, fontSize: 14 }}>Ft. {s.doc}</span>
                  <span className="mono" style={{ fontWeight: 600 }}>{fmtEur(s.residuo || s.importo)}</span>
                </div>
                <div style={{ fontSize: 12, color: "#5a5868" }}>Scad: {fmtData(s.scadenza)} · {s.giorni || 0}gg ritardo
                  {s.residuo < s.importo && <span style={{ color: "#EF9F27", marginLeft: 8 }}>Parziale (orig. {fmtEur(s.importo)})</span>}
                </div>
              </div>
            ))}
          </div>
        )}

        {tab === "diario" && (
          <div>
            <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto", paddingBottom: 2 }}>
              {["TUTTI", ...TIPI_DIARIO].map(t => (
                <button key={t} onClick={() => setFiltroTipo(t)} style={{ padding: "6px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: filtroTipo === t ? "#4a7fd4" : "transparent", color: filtroTipo === t ? "#fff" : "#5a5868", border: filtroTipo === t ? "none" : "1px solid #2a2a38" }}>{t}</button>
              ))}
            </div>
            {diariFiltrati.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 30, color: "#4a4858" }}>Nessuna voce</div> : diariFiltrati.map((v, i) => {
              const DIARIO_COLORS_LOCAL = THEMES[localStorage.getItem("crm_theme") || "dark"]?.diario || THEMES.dark.diario;
          const dc = DIARIO_COLORS_LOCAL[v.tipo] || DIARIO_COLORS_LOCAL["Nota personale"];
              return (
                <div key={i} style={{ borderLeft: `3px solid ${dc.border}`, padding: "10px 12px", background: dc.bg, borderRadius: "0 10px 10px 0", marginBottom: 8 }}>
                  <div style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ fontSize: 11, color: "#5a5868" }}>{fmtData(v.data)} {v.ora}</span>
                    <span style={{ fontSize: 11, fontWeight: 500, color: dc.label, background: dc.bg, border: `1px solid ${dc.border}`, padding: "1px 8px", borderRadius: 20 }}>{v.tipo}</span>
                    {v.canale && <span className="tag" style={{ background: "#1a1a25", color: "#9a98a0" }}>{v.canale}</span>}
                    {v.stato && <span className="tag" style={{ background: (STATO_COLORS[v.stato] || "#888") + "22", color: STATO_COLORS[v.stato] || "#888" }}>{v.stato}</span>}
                  </div>
                  <div style={{ fontSize: 13, lineHeight: 1.5 }}>{v.testo || v.esito}</div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {showContatto && <ModalContatto cliente={c} onClose={() => setShowContatto(false)} onSave={(voce, updates) => { onAddDiario(voce); onUpdate(updates); setShowContatto(false); showToast("Contatto registrato"); }} />}
      {showIncasso && <ModalIncasso cliente={c} onClose={() => setShowIncasso(false)} onSave={inc => { onAddIncasso(inc); setShowIncasso(false); showToast("Incasso registrato"); }} />}
      {showAccorpa && <ModalAccorpa cliente={c} onClose={() => setShowAccorpa(false)} onSave={updates => { onUpdate(updates); setShowAccorpa(false); showToast("Scadenze aggiornate"); }} />}
      {showInsoluto && <ModalInsoluto cliente={c} agenti={agenti} onClose={() => setShowInsoluto(false)} onSave={ins => { onAddInsoluto(ins); setShowInsoluto(false); }} />}
    </div>
  );
}

/* ─── MODAL CONTATTO ─── */
function ModalContatto({ cliente, onClose, onSave }) {
  const [data, setData] = useState(oggi());
  const [ora, setOra] = useState(oraOra());
  const [tipo, setTipo] = useState("Mio contatto");
  const [canale, setCanale] = useState("Telefono");
  const [testo, setTesto] = useState("");
  const [stato, setStato] = useState(cliente.statoOp);
  const [priorita, setPriorita] = useState(cliente.priorita);
  const [richiamo, setRichiamo] = useState(cliente.dataRichiamo || "");

  function handleSave(e) {
    e.preventDefault();
    const voce = { data, ora, tipo, canale: tipo !== "Nota personale" ? canale : null, testo, stato, ts: new Date().toISOString() };
    const updates = { statoOp: stato, esito: testo, ultimoContatto: tipo !== "Nota personale" ? data : cliente.ultimoContatto, dataRichiamo: richiamo || null, priorita };
    onSave(voce, updates);
  }

  const DIARIO_COLORS_L = THEMES[localStorage.getItem("crm_theme") || "dark"]?.diario || THEMES.dark.diario;
  const dc = DIARIO_COLORS_L[tipo] || DIARIO_COLORS_L["Nota personale"];

  return (
    <div className="modal-bg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Registra contatto</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group">
            <label>TIPO</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {TIPI_DIARIO.map(t => {
                const tc = DIARIO_COLORS[t];
                return <button key={t} type="button" onClick={() => setTipo(t)} style={{ padding: "7px 14px", fontSize: 13, borderRadius: 20, background: tipo === t ? tc.bg : "transparent", color: tipo === t ? tc.label : "#5a5868", border: `1px solid ${tipo === t ? tc.border : "#2a2a38"}`, fontWeight: tipo === t ? 500 : 400 }}>{t}</button>;
              })}
            </div>
          </div>
          <div className="row2">
            <div className="form-group"><label>DATA</label><input type="date" value={data} onChange={e => setData(e.target.value)} /></div>
            <div className="form-group"><label>ORA</label><input type="time" value={ora} onChange={e => setOra(e.target.value)} /></div>
          </div>
          {tipo !== "Nota personale" && (
            <div className="form-group"><label>CANALE</label><select value={canale} onChange={e => setCanale(e.target.value)}>{TIPI_CONTATTO.map(t => <option key={t}>{t}</option>)}</select></div>
          )}
          <div className="form-group"><label>NOTE / ESITO</label><textarea value={testo} onChange={e => setTesto(e.target.value)} placeholder="Scrivi liberamente..." /></div>
          <div className="row2">
            <div className="form-group"><label>STATO</label><select value={stato} onChange={e => setStato(e.target.value)}>{STATI_OP.map(s => <option key={s}>{s}</option>)}</select></div>
            <div className="form-group"><label>PRIORITÀ</label><select value={priorita} onChange={e => setPriorita(e.target.value)}>{PRIORITA.map(p => <option key={p}>{p}</option>)}</select></div>
          </div>
          <div className="form-group"><label>DATA RICHIAMO — vuoto = nessuno</label><input type="date" value={richiamo} onChange={e => setRichiamo(e.target.value)} /></div>
          <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }}>Salva</button>
        </form>
      </div>
    </div>
  );
}

/* ─── MODAL INSOLUTO ─── */
function ModalInsoluto({ cliente, agenti, onClose, onSave }) {
  const [importo, setImporto] = useState("");
  const [riferimento, setRiferimento] = useState("");
  const [dataEmissione, setDataEmissione] = useState("");
  const [motivo, setMotivo] = useState("Mancanza fondi");
  const agenteCliente = agenti.find(a => a.nome === cliente.agente);
  const mailCliente = cliente.email ? generaMailInsolutoCliente(cliente, { riferimento, importo: parseFloat(importo || 0), dataEmissione, motivo }) : null;
  const mailAgente = agenteCliente ? generaMailInsolutoAgente(agenteCliente, cliente, { riferimento, importo: parseFloat(importo || 0), dataEmissione, motivo }) : null;

  function handleSave(e) {
    e.preventDefault();
    onSave({ clienteId: cliente.id, ragione: cliente.ragione, agente: cliente.agente, importo: parseFloat(importo), riferimento, dataEmissione, motivo });
  }

  return (
    <div className="modal-bg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600, color: "#e24b4a" }}>⚠ Assegno insoluto</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group"><label>IMPORTO (€)</label><input type="number" step="0.01" value={importo} onChange={e => setImporto(e.target.value)} placeholder="0.00" required /></div>
          <div className="row2">
            <div className="form-group"><label>N. ASSEGNO / RIF.</label><input value={riferimento} onChange={e => setRiferimento(e.target.value)} placeholder="es. ASS001" /></div>
            <div className="form-group"><label>DATA EMISSIONE</label><input type="date" value={dataEmissione} onChange={e => setDataEmissione(e.target.value)} /></div>
          </div>
          <div className="form-group"><label>MOTIVO</label>
            <select value={motivo} onChange={e => setMotivo(e.target.value)}>
              {["Mancanza fondi", "Impagato", "Conto chiuso", "Firma non corrispondente", "Altro"].map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
          <div style={{ background: "#2a0a0a", border: "1px solid #4a1a1a", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#e24b4a" }}>
            Il cliente passerà automaticamente in stato <strong>PROBLEMA</strong> e avrai 60 giorni per la regolarizzazione.
          </div>
          <button type="submit" className="btn-danger" style={{ width: "100%", padding: 14, marginBottom: 10 }}>Registra insoluto</button>
          <div style={{ display: "flex", gap: 8 }}>
            {mailCliente && <a href={mailCliente} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#0a1f2a", color: "#378ADD", border: "1px solid #1a3a4a", borderRadius: 10, padding: "10px", fontSize: 13, textDecoration: "none" }}>✉ Mail cliente</a>}
            {mailAgente && <a href={mailAgente} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#0f2a1a", color: "#4ecb8d", border: "1px solid #1a4a2a", borderRadius: 10, padding: "10px", fontSize: 13, textDecoration: "none" }}>✉ Mail agente</a>}
          </div>
        </form>
      </div>
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
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Incasso — {cliente.ragione}</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group"><label>TIPO</label><select value={tipo} onChange={e => setTipo(e.target.value)}>{TIPI_TITOLO.map(t => <option key={t}>{t}</option>)}</select></div>
          <div className="row2">
            <div className="form-group"><label>IMPORTO (€)</label><input type="number" step="0.01" value={importo} onChange={e => setImporto(e.target.value)} required /></div>
            <div className="form-group"><label>SPESE (€)</label><input type="number" step="0.01" value={spese} onChange={e => setSpese(e.target.value)} /></div>
          </div>
          {importo && <div style={{ background: "#0f2a1a", border: "1px solid #1a4a2a", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 13, color: "#4ecb8d" }}>Netto: {fmtEur(parseFloat(importo || 0) - parseFloat(spese || 0))}</div>}
          <div className="form-group"><label>FATTURE</label><input value={fatture} onChange={e => setFatture(e.target.value)} placeholder="es. Ft.001 + Ft.002" /></div>
          <div className="form-group"><label>NOTE</label><textarea value={note} onChange={e => setNote(e.target.value)} style={{ minHeight: 60 }} /></div>
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
  const [tipoOp, setTipoOp] = useState("pagata");
  const [importoPagato, setImportoPagato] = useState("");
  const [fatturaAccorpante, setFatturaAccorpante] = useState("");

  function toggleSel(tag) { setSelezionate(s => s.includes(tag) ? s.filter(t => t !== tag) : [...s, tag]); }

  function handleSave(e) {
    e.preventDefault();
    const nuoveScadenze = cliente.scadenze.map(s => {
      if (!selezionate.includes(s.tag)) return s;
      if (tipoOp === "pagata") return { ...s, stato: "PAGATA", residuo: 0 };
      if (tipoOp === "parziale") return { ...s, residuo: Math.max(0, s.importo - parseFloat(importoPagato || 0)) };
      if (tipoOp === "accorpa") return { ...s, accorpata: true, fatturaAccorpante };
      return s;
    });
    const nuovoTotale = nuoveScadenze.filter(s => s.stato !== "PAGATA" && !s.accorpata).reduce((sum, s) => sum + (s.residuo || s.importo), 0);
    onSave({ scadenze: nuoveScadenze, totaleScaduto: nuovoTotale });
  }

  return (
    <div className="modal-bg" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>Gestisci pagamento</div>
          <button className="btn-ghost" style={{ padding: "6px 12px" }} onClick={onClose}>✕</button>
        </div>
        <form onSubmit={handleSave}>
          <div className="form-group"><label>OPERAZIONE</label>
            <select value={tipoOp} onChange={e => setTipoOp(e.target.value)}>
              <option value="pagata">Pagata interamente</option>
              <option value="parziale">Pagamento parziale</option>
              <option value="accorpa">Accorpa sotto altra fattura</option>
            </select>
          </div>
          {tipoOp === "parziale" && <div className="form-group"><label>IMPORTO PAGATO (€)</label><input type="number" step="0.01" value={importoPagato} onChange={e => setImportoPagato(e.target.value)} /></div>}
          {tipoOp === "accorpa" && <div className="form-group"><label>FATTURA ACCORPANTE</label><input value={fatturaAccorpante} onChange={e => setFatturaAccorpante(e.target.value)} placeholder="es. 12345" /></div>}
          <div className="form-group"><label>SCADENZE ({selezionate.length} selezionate)</label>
            {scadenzeAperte.map((s, i) => (
              <div key={i} onClick={() => toggleSel(s.tag)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: selezionate.includes(s.tag) ? "#0f2a1a" : "#0d0d14", border: `1px solid ${selezionate.includes(s.tag) ? "#1a4a2a" : "#1a1a25"}`, borderRadius: 10, marginBottom: 6, cursor: "pointer" }}>
                <div style={{ width: 20, height: 20, borderRadius: 4, border: `2px solid ${selezionate.includes(s.tag) ? "#4ecb8d" : "#2a2a38"}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {selezionate.includes(s.tag) && <span style={{ color: "#4ecb8d", fontSize: 12 }}>✓</span>}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 500 }}>Ft. {s.doc} — {fmtEur(s.residuo || s.importo)}</div>
                  <div style={{ fontSize: 11, color: "#5a5868" }}>Scad: {fmtData(s.scadenza)}</div>
                </div>
              </div>
            ))}
          </div>
          <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }} disabled={selezionate.length === 0}>Applica</button>
        </form>
      </div>
    </div>
  );
}

/* ─── AGENTI ─── */
function Agenti({ agenti, clienti, incassi, onAddAgente, onSelectAgente }) {
  const [showForm, setShowForm] = useState(false);
  const [nome, setNome] = useState("");
  const [cognome, setCognome] = useState("");
  const [email, setEmail] = useState("");
  const [telefono, setTelefono] = useState("");
  const [zona, setZona] = useState("");

  function handleAdd(e) {
    e.preventDefault();
    onAddAgente({ nome: `${nome} ${cognome}`.trim(), cognome, email, telefono, zona });
    setNome(""); setCognome(""); setEmail(""); setTelefono(""); setZona("");
    setShowForm(false);
  }

  // Agenti estratti dai clienti (anche se non ancora in anagrafica agenti)
  const agentiNomi = useMemo(() => {
    const fromClienti = [...new Set(clienti.map(c => c.agente).filter(Boolean))];
    const fromAgenti = agenti.map(a => a.nome);
    return [...new Set([...fromAgenti, ...fromClienti])].sort();
  }, [agenti, clienti]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div><div className="section-title">Agenti</div><div style={{ color: "#5a5868", fontSize: 13 }}>{agentiNomi.length} agenti</div></div>
        <button className="btn-primary" onClick={() => setShowForm(s => !s)}>{showForm ? "Chiudi" : "+ Aggiungi"}</button>
      </div>

      {showForm && (
        <div className="card" style={{ marginBottom: 16 }}>
          <form onSubmit={handleAdd}>
            <div className="row2">
              <div className="form-group"><label>NOME</label><input value={nome} onChange={e => setNome(e.target.value)} placeholder="Mario" required /></div>
              <div className="form-group"><label>COGNOME</label><input value={cognome} onChange={e => setCognome(e.target.value)} placeholder="Rossi" /></div>
            </div>
            <div className="form-group"><label>EMAIL</label><input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="agente@azienda.it" /></div>
            <div className="row2">
              <div className="form-group"><label>TELEFONO</label><input value={telefono} onChange={e => setTelefono(e.target.value)} /></div>
              <div className="form-group"><label>ZONA</label><input value={zona} onChange={e => setZona(e.target.value)} placeholder="es. Milano Nord" /></div>
            </div>
            <button type="submit" className="btn-primary" style={{ width: "100%", padding: 14 }}>Salva agente</button>
          </form>
        </div>
      )}

      {agentiNomi.map(nomeAg => {
        const ag = agenti.find(a => a.nome === nomeAg);
        const clientiAg = clienti.filter(c => c.agente === nomeAg && c.statoOp !== "PAGATO");
        const incassiAg = incassi.filter(i => i.agente === nomeAg && i.stato === "IN MANO AGENTE");
        const totScaduto = clientiAg.reduce((s, c) => s + (c.totaleScaduto || 0), 0);
        const totMano = incassiAg.reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);
        const hasAlert = incassiAg.some(i => giorniDa(i.dataRicezione) > 15);

        return (
          <div key={nomeAg} className="cl-card" onClick={() => { if (ag) onSelectAgente(ag.id); else onSelectAgente(nomeAg); }} style={{ borderColor: hasAlert ? "#3a2e08" : "#1e1e2a" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 15 }}>{nomeAg}</div>
                {ag?.zona && <div style={{ fontSize: 11, color: "#5a5868" }}>{ag.zona}</div>}
                {!ag && <div style={{ fontSize: 11, color: "#4a3010" }}>⚠ Anagrafica mancante</div>}
              </div>
              <div style={{ textAlign: "right" }}>
                <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(totScaduto)}</div>
                <div style={{ fontSize: 11, color: "#5a5868" }}>scaduto</div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, fontSize: 12 }}>
              <span style={{ color: "#5a5868" }}>{clientiAg.length} clienti attivi</span>
              {incassiAg.length > 0 && <span style={{ color: hasAlert ? "#EF9F27" : "#378ADD" }}>💼 {fmtEur(totMano)} in mano</span>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ─── SCHEDA AGENTE ─── */
function SchedaAgente({ agente, clienti, incassi, themeMode, onBack, onUpdate, onAddDiario, onUpdateCliente, onUpdateIncasso, onDeleteIncasso, onSelectCliente, showToast }) {
  const T = THEMES[themeMode] || THEMES.dark;
  const [tab, setTab] = useState("clienti");
  const [tracking, setTracking] = useState({});

  const [sortStato, setSortStato] = useState("AGENTE");
  // Clienti ordinati: stato selezionato in cima, poi per scaduto
  const clientiOrdinati = useMemo(() => [...clienti].sort((a, b) => {
    if (sortStato !== "TUTTI") {
      if (a.statoOp === sortStato && b.statoOp !== sortStato) return -1;
      if (b.statoOp === sortStato && a.statoOp !== sortStato) return 1;
    }
    return (b.totaleScaduto || 0) - (a.totaleScaduto || 0);
  }), [clienti, sortStato]);

  const totScaduto = clienti.reduce((s, c) => s + (c.totaleScaduto || 0), 0);
  const totMano = incassi.reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0);

  function registraContattoSuCliente(clienteId, testo, richiamo) {
    const voce = { data: oggi(), ora: oraOra(), tipo: "Via agente", testo, ts: new Date().toISOString() };
    onAddDiario(clienteId, voce);
    if (richiamo) onUpdateCliente(clienteId, { dataRichiamo: richiamo, ultimoContatto: oggi() });
    showToast("Annotazione registrata sul cliente");
  }

  return (
        <div style={{ minHeight: "100vh", background: T.bg, color: T.text, fontFamily: "'DM Sans',sans-serif", paddingBottom: 40 }}>
      <style>{getSharedStyles(T)}</style>

      <div style={{ padding: "52px 20px 0", marginBottom: 16 }}>
        <button className="btn-ghost" style={{ marginBottom: 16, padding: "8px 14px", fontSize: 13 }} onClick={onBack}><i className="ti ti-arrow-left" /> Torna</button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>{agente.nome}</div>
            <div style={{ fontSize: 12, color: "#5a5868" }}>{agente.zona || "—"} · {agente.email || "—"}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div className="mono" style={{ fontSize: 20, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(totScaduto)}</div>
            {totMano > 0 && <div style={{ fontSize: 12, color: "#EF9F27", marginTop: 2 }}>💼 {fmtEur(totMano)} in mano</div>}
          </div>
        </div>
      </div>

      <div style={{ padding: "0 20px" }}>
        <div className="tab-bar">
          <button className={`tab ${tab === "clienti" ? "active" : ""}`} onClick={() => setTab("clienti")}>Clienti ({clienti.length})</button>
          <button className={`tab ${tab === "titoli" ? "active" : ""}`} onClick={() => setTab("titoli")}>Titoli in mano ({incassi.length})</button>
          <button className={`tab ${tab === "info" ? "active" : ""}`} onClick={() => setTab("info")}>Anagrafica</button>
        </div>

        {tab === "clienti" && (
          <div>
            <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto", paddingBottom: 2 }}>
              {["TUTTI", "AGENTE", "PROBLEMA", "CONTROLLA", "L/L", "AGENZIA", "RICHIAMA"].map(s => (
                <button key={s} onClick={() => setSortStato(s)} style={{ padding: "6px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: sortStato === s ? "#4a7fd4" : "transparent", color: sortStato === s ? "#fff" : "#5a5868", border: sortStato === s ? "none" : "1px solid #2a2a38" }}>{s}</button>
              ))}
            </div>
            {clientiOrdinati.map(c => (
              <ClienteAgente key={c.id} cliente={c} onApriScheda={() => onSelectCliente(c.id)} onAnnota={(testo, richiamo) => registraContattoSuCliente(c.id, testo, richiamo)} showToast={showToast} />
            ))}
            {clienti.length === 0 && <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun cliente assegnato</div>}
          </div>
        )}

        {tab === "titoli" && (
          <div>
            {incassi.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo in mano</div> : incassi.map(i => {
              const gg = giorniDa(i.dataRicezione);
              const isAlert = gg > 30;
              const isWarn = gg > 15 && !isAlert;
              return (
                <div key={i.id} className="card" style={{ marginBottom: 10, borderColor: isAlert ? "#3a2020" : isWarn ? "#3a2e08" : "#1e1e2a" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                    <div>
                      <div style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4a7fd4" }} onClick={() => onSelectCliente(i.clienteId)}>{i.ragione}</div>
                      <div style={{ fontSize: 11, color: "#5a5868" }}>{i.tipo} · {fmtData(i.dataRicezione)} ({gg}gg fa)</div>
                    </div>
                    <div className="mono" style={{ fontWeight: 600 }}>{fmtEur(i.incassoNetto || i.importo)}</div>
                  </div>
                  {i.note && <div style={{ fontSize: 12, color: "#7a7888", marginBottom: 8 }}>{i.note}</div>}
                  {isAlert && <div style={{ fontSize: 12, color: "#e24b4a", marginBottom: 8, fontWeight: 500 }}>⚠ In mano da {gg}gg — sollecita!</div>}
                  {isWarn && <div style={{ fontSize: 12, color: "#EF9F27", marginBottom: 8 }}>⏳ In mano da {gg}gg</div>}
                  <div style={{ display: "flex", gap: 8 }}>
                    <input value={tracking[i.id] || ""} onChange={e => setTracking(t => ({ ...t, [i.id]: e.target.value }))} placeholder="Tracking spedizione..." style={{ flex: 1, fontSize: 13, padding: "8px 10px" }} />
                    <button className="btn-green" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => { onUpdateIncasso(i.id, { stato: "SPEDITO", dataSpedizione: oggi(), tracking: tracking[i.id] || "" }); showToast("Spedizione registrata"); }}>Spedito</button>
                  </div>
                  <button onClick={() => { onDeleteIncasso && onDeleteIncasso(i.id); showToast("Incasso eliminato"); }} style={{ marginTop: 8, background: "transparent", color: "#e24b4a", border: "1px solid #3a2020", borderRadius: 8, padding: "4px 12px", fontSize: 12, cursor: "pointer", width: "100%" }}>✕ Elimina</button>
                </div>
              );
            })}
          </div>
        )}

        {tab === "info" && (
          <div className="card">
            {[["Nome", agente.nome], ["Email", agente.email || "—"], ["Telefono", agente.telefono || "—"], ["Zona", agente.zona || "—"]].map(([k, v]) => (
              <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: "1px solid #1a1a25", fontSize: 13 }}>
                <span style={{ color: "#5a5868" }}>{k}</span><span style={{ fontWeight: 500 }}>{v}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ─── CLIENTE NELLA SCHEDA AGENTE ─── */
function ClienteAgente({ cliente, onApriScheda, onAnnota }) {
  const [showAnnota, setShowAnnota] = useState(false);
  const [testo, setTesto] = useState("");
  const [richiamo, setRichiamo] = useState("");
  const c = cliente;

  // Ultimi contatti via agente
  const ultimeViaAgente = (c.diario || []).filter(v => v.tipo === "Via agente").slice(0, 2);

  return (
    <div style={{ background: "#13131a", border: `1px solid ${c.statoOp === "AGENTE" ? "#3a2e08" : "#1e1e2a"}`, borderRadius: 14, padding: 14, marginBottom: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
        <div>
          <div style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4a7fd4" }} onClick={onApriScheda}>{c.ragione}</div>
          <div style={{ fontSize: 11, color: "#5a5868" }}>{c.codice}</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(c.totaleScaduto || 0)}</div>
          <span className="tag" style={{ background: (STATO_COLORS[c.statoOp] || "#888") + "22", color: STATO_COLORS[c.statoOp] || "#888" }}>{c.statoOp}</span>
        </div>
      </div>

      {/* Storico via agente */}
      {ultimeViaAgente.length > 0 && (
        <div style={{ marginBottom: 8 }}>
          {ultimeViaAgente.map((v, i) => (
            <div key={i} style={{ fontSize: 12, color: "#4ecb8d", borderLeft: "2px solid #4ecb8d", paddingLeft: 8, marginBottom: 4 }}>
              {fmtData(v.data)}: {v.testo?.slice(0, 60)}{v.testo?.length > 60 ? "…" : ""}
            </div>
          ))}
        </div>
      )}

      {!showAnnota ? (
        <button className="btn-ghost" style={{ width: "100%", fontSize: 13, padding: "8px" }} onClick={() => setShowAnnota(true)}>
          + Annota risposta agente
        </button>
      ) : (
        <div style={{ marginTop: 8 }}>
          <textarea value={testo} onChange={e => setTesto(e.target.value)} placeholder="Cosa ti ha detto l'agente su questo cliente..." style={{ marginBottom: 8, fontSize: 14, minHeight: 60, background: "#0a2018", border: "1px solid #4ecb8d33", borderRadius: 10, color: "#e8e6df", padding: "10px 14px", width: "100%", outline: "none" }} />
          <div style={{ display: "flex", gap: 8 }}>
            <input type="date" value={richiamo} onChange={e => setRichiamo(e.target.value)} placeholder="Richiamo (opc.)" style={{ flex: 1, fontSize: 13, padding: "8px 10px" }} />
            <button className="btn-green" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => { if (testo) { onAnnota(testo, richiamo ? addGiorni(richiamo, 1) : ""); setTesto(""); setRichiamo(""); setShowAnnota(false); } }}>Salva</button>
            <button className="btn-ghost" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => setShowAnnota(false)}>✕</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── PRATICHE L/L ─── */
function PraticheLl({ pratiche, clienti, onUpdate, onSelectCliente, showToast }) {
  const [filtro, setFiltro] = useState("ATTIVE");
  const praticheFiltrate = pratiche.filter(p => filtro === "TUTTE" ? true : filtro === "ATTIVE" ? p.stato !== "AGENZIA" && p.stato !== "CHIUSA" : p.stato === filtro);

  function avanzaLettera(p) {
    const n = p.lettera || 0;
    if (n >= 3) { onUpdate(p.id, { stato: "AGENZIA" }); showToast("→ Agenzia"); return; }
    onUpdate(p.id, { lettera: n + 1, stato: `LETTERA ${n + 1}`, dataUltimaLettera: oggi(), esito: "" });
    showToast(`Lettera ${n + 1} registrata`);
  }

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
        <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}><div style={{ fontSize: 13 }}>Nessuna pratica</div></div>
      ) : praticheFiltrate.map(p => {
        const gg = p.dataUltimaLettera ? giorniDa(p.dataUltimaLettera) : null;
        const pronta = gg === null || gg >= 20;
        return (
          <div key={p.id} className="card" style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 15, cursor: "pointer", color: "#4a7fd4", marginBottom: 2 }} onClick={() => { const cl = clienti.find(c => c.id === p.clienteId); if (cl) onSelectCliente(cl.id); }}>{p.ragione}</div>
                <div style={{ fontSize: 11, color: "#5a5868" }}>{p.agente} · {fmtData(p.dataDecisione)}</div>
              </div>
              <span style={{ background: p.stato === "AGENZIA" ? "#7F77DD22" : p.stato === "IN CODA" ? "#EF9F2722" : "#e24b4a22", color: p.stato === "AGENZIA" ? "#7F77DD" : p.stato === "IN CODA" ? "#EF9F27" : "#e24b4a", fontSize: 12, padding: "4px 10px", borderRadius: 8 }}>{p.stato}</span>
            </div>
            <div style={{ display: "flex", gap: 4, marginBottom: 10 }}>
              {["C", "L1", "L2", "L3", "AG"].map((s, i) => <div key={s} style={{ flex: 1, height: 4, borderRadius: 2, background: i <= (p.lettera || 0) ? "#e24b4a" : "#1e1e2a" }} />)}
            </div>
            {gg !== null && <div style={{ fontSize: 12, color: gg >= 20 ? "#4ecb8d" : "#EF9F27", marginBottom: 8 }}>{gg >= 20 ? `✓ Pronta (${gg}gg)` : `⏳ Ancora ${20 - gg}gg`}</div>}
            {p.esito && <div style={{ fontSize: 12, color: "#5a5868", marginBottom: 10 }}>Esito: {p.esito}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              {p.stato !== "AGENZIA" && <select onChange={e => e.target.value && onUpdate(p.id, { esito: e.target.value, dataEsito: oggi() })} defaultValue="" style={{ flex: 1, fontSize: 13, padding: "8px 10px" }}>
                <option value="">Esito raccomandata...</option>
                {ESITI_RACC.map(e => <option key={e} value={e}>{e}</option>)}
              </select>}
              {p.stato !== "AGENZIA" && pronta && <button className="btn-danger" style={{ fontSize: 13, padding: "8px 14px", whiteSpace: "nowrap" }} onClick={() => avanzaLettera(p)}>{p.lettera >= 3 ? "→ Agenzia" : `→ L${(p.lettera || 0) + 1}`}</button>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ─── INCASSI ─── */
function Incassi({ incassi, clienti, onAdd, onUpdate, onDelete, onSelectCliente, showToast }) {
  const [tab, setTab] = useState("mano");
  const inMano = incassi.filter(i => i.stato === "IN MANO AGENTE");
  const spediti = incassi.filter(i => i.stato === "SPEDITO");

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

  const IncassoCard = ({ inc }) => {
    const [tracking, setTracking] = useState("");
    const gg = giorniDa(inc.dataRicezione);
    return (
      <div className="card" style={{ marginBottom: 10, borderColor: gg > 30 ? "#3a2020" : gg > 15 ? "#3a2e08" : "#1e1e2a" }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4a7fd4" }} onClick={() => { const cl = clienti.find(c => c.id === inc.clienteId); if (cl) onSelectCliente(cl.id); }}>{inc.ragione}</div>
            <div style={{ fontSize: 11, color: "#5a5868" }}>{inc.agente} · {inc.tipo} · {fmtData(inc.dataRicezione)} ({gg}gg)</div>
          </div>
          <div className="mono" style={{ fontWeight: 600 }}>{fmtEur(inc.incassoNetto || inc.importo)}</div>
        </div>
        {inc.note && <div style={{ fontSize: 12, color: "#7a7888", marginBottom: 8 }}>{inc.note}</div>}
        {gg > 30 && <div style={{ fontSize: 12, color: "#e24b4a", marginBottom: 8, fontWeight: 500 }}>⚠ {gg}gg — sollecita spedizione!</div>}
        {gg > 15 && gg <= 30 && <div style={{ fontSize: 12, color: "#EF9F27", marginBottom: 8 }}>⏳ {gg}gg in mano</div>}
        {inc.stato === "IN MANO AGENTE" && (
          <div style={{ display: "flex", gap: 8 }}>
            <input value={tracking} onChange={e => setTracking(e.target.value)} placeholder="Tracking..." style={{ flex: 1, fontSize: 13, padding: "8px 10px" }} />
            <button className="btn-green" style={{ fontSize: 13, padding: "8px 14px" }} onClick={() => { onUpdate(inc.id, { stato: "SPEDITO", dataSpedizione: oggi(), tracking }); showToast("Spedito"); }}>Spedito</button>
          </div>
        )}
        {inc.stato === "SPEDITO" && (
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#4ecb8d" }}>✓ Spedito {fmtData(inc.dataSpedizione)}{inc.tracking ? ` · ${inc.tracking}` : ""}</div>
            <button className="btn-ghost" style={{ fontSize: 12, padding: "6px 12px" }} onClick={() => { onUpdate(inc.id, { stato: "REGISTRATO", dataRegistrazione: oggi() }); showToast("Registrato"); }}>Registrato</button>
          </div>
        )}
        {inc.stato === "REGISTRATO" && <div style={{ fontSize: 12, color: "#1D9E75" }}>✓ Registrato {fmtData(inc.dataRegistrazione)}</div>}
        <button onClick={() => onDelete(inc.id)} style={{ marginTop: 8, background: "transparent", color: "#e24b4a", border: "1px solid #3a2020", borderRadius: 8, padding: "4px 12px", fontSize: 12, cursor: "pointer", width: "100%" }}>✕ Elimina incasso</button>
      </div>
    );
  };

  return (
    <div>
      <div className="section-title">Incassi agenti</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 16 }}>{inMano.length} in mano · {spediti.length} spediti</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto" }}>
        {[["mano", `In mano (${inMano.length})`], ["agente", "Per agente"], ["spediti", `Spediti (${spediti.length})`], ["tutti", "Tutti"]].map(([k, l]) => (
          <button key={k} onClick={() => setTab(k)} style={{ padding: "7px 14px", fontSize: 12, borderRadius: 20, whiteSpace: "nowrap", background: tab === k ? "#4a7fd4" : "transparent", color: tab === k ? "#fff" : "#5a5868", border: tab === k ? "none" : "1px solid #2a2a38" }}>{l}</button>
        ))}
      </div>
      {tab === "mano" && (inMano.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo in mano</div> : inMano.map(i => <IncassoCard key={i.id} inc={i} />))}
      {tab === "agente" && (totalePerAgente.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo in mano</div> : totalePerAgente.map(ag => (
        <div key={ag.agente} className="card" style={{ marginBottom: 12 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
            <div style={{ fontWeight: 600 }}>{ag.agente}</div>
            <div className="mono" style={{ fontWeight: 600, color: "#EF9F27" }}>{fmtEur(ag.totale)}</div>
          </div>
          <div style={{ fontSize: 12, color: "#5a5868", marginBottom: 10 }}>{ag.count} titoli</div>
          {ag.titoli.map(i => <div key={i.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: "1px solid #1a1a25", fontSize: 13 }}><span>{i.ragione} · {i.tipo}</span><span className="mono">{fmtEur(i.incassoNetto || i.importo)}</span></div>)}
        </div>
      )))}
      {tab === "spediti" && (spediti.length === 0 ? <div className="card" style={{ textAlign: "center", padding: 40, color: "#4a4858" }}>Nessun titolo spedito</div> : spediti.map(i => <IncassoCard key={i.id} inc={i} />))}
      {tab === "tutti" && incassi.map(i => <IncassoCard key={i.id} inc={i} />)}
    </div>
  );
}

/* ─── DASHBOARD ─── */
function Dashboard({ kpi, clienti, praticheLl, incassi, agenti, insoluti, onSelectAgente }) {
  const byStato = useMemo(() => {
    const map = {};
    STATI_OP.forEach(s => { map[s] = clienti.filter(c => c.statoOp === s).length; });
    return map;
  }, [clienti]);

  const agentiKpi = useMemo(() => {
    const nomi = [...new Set(clienti.map(c => c.agente).filter(Boolean))];
    return nomi.map(nome => {
      const ag = agenti.find(a => a.nome === nome);
      const cl = clienti.filter(c => c.agente === nome && c.statoOp !== "PAGATO");
      const inc = incassi.filter(i => i.agente === nome && i.stato === "IN MANO AGENTE");
      return { nome, id: ag?.id, totScaduto: cl.reduce((s, c) => s + (c.totaleScaduto || 0), 0), nClienti: cl.length, totMano: inc.reduce((s, i) => s + (i.incassoNetto || i.importo || 0), 0), hasAlert: inc.some(i => giorniDa(i.dataRicezione) > 15) };
    }).sort((a, b) => b.totScaduto - a.totScaduto);
  }, [clienti, incassi, agenti]);

  const insolutiAperti = insoluti.filter(i => i.stato === "APERTO");

  return (
    <div>
      <div className="section-title">Dashboard</div>
      <div style={{ color: "#5a5868", fontSize: 13, marginBottom: 20 }}>
        {new Date().toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long" })}
        {kpi.dataUltimoImport && <span style={{ marginLeft: 12, fontSize: 11, color: "#4a4858" }}>· Ultimo import: {fmtData(kpi.dataUltimoImport)}</span>}
      </div>

      {/* SCADUTO */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
        <div className="card" style={{ gridColumn: "1 / -1" }}>
          <div style={{ fontSize: 11, color: "#5a5868", marginBottom: 6 }}>SCADUTO TOTALE</div>
          <div className="mono" style={{ fontSize: 26, fontWeight: 700, color: "#e24b4a" }}>{fmtEur(kpi.totScaduto)}</div>
          {kpi.dataUltimoImport && <div style={{ fontSize: 12, color: "#5a5868", marginTop: 4 }}>Ultimo import: {fmtEur(kpi.scadutoMese)}</div>}
        </div>
        {[
          { label: "Pratiche L/L", val: kpi.llAperte, color: "#e24b4a" },
          { label: "Titoli in mano", val: kpi.titoliMano, color: "#EF9F27" },
          { label: "Insoluti aperti", val: kpi.insoluti, color: "#FF4A8D" },
          { label: "Richiami oggi", val: kpi.daRichiamare, color: "#378ADD" },
        ].map(k => (
          <div key={k.label} className="card">
            <div style={{ fontSize: 11, color: "#5a5868", marginBottom: 6 }}>{k.label.toUpperCase()}</div>
            <div className="mono" style={{ fontSize: 22, fontWeight: 700, color: k.color }}>{k.val}</div>
          </div>
        ))}
      </div>

      {/* INCASSATO */}
      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: "#9a98a0" }}>Incassato</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
          <div style={{ background: "#0f2a1a", border: "1px solid #1a4a2a", borderRadius: 12, padding: "12px" }}>
            <div style={{ fontSize: 10, color: "#4ecb8d", marginBottom: 4 }}>CONFERMATO TOTALE</div>
            <div className="mono" style={{ fontSize: 16, fontWeight: 700, color: "#4ecb8d" }}>{fmtEur(kpi.incassatoConfermatoTot)}</div>
            <div style={{ fontSize: 11, color: "#4ecb8d88", marginTop: 2 }}>{kpi.pctConfermatoTot.toFixed(1)}% scaduto</div>
          </div>
          <div style={{ background: "#0a1f2a", border: "1px solid #1a3a4a", borderRadius: 12, padding: "12px" }}>
            <div style={{ fontSize: 10, color: "#378ADD", marginBottom: 4 }}>IN CORSO TOTALE</div>
            <div className="mono" style={{ fontSize: 16, fontWeight: 700, color: "#378ADD" }}>{fmtEur(kpi.incassatoAtteso)}</div>
            <div style={{ fontSize: 11, color: "#378ADD88", marginTop: 2 }}>in mano/spedito</div>
          </div>
          <div style={{ background: "#1a1a25", border: "1px solid #2a2a38", borderRadius: 12, padding: "12px" }}>
            <div style={{ fontSize: 10, color: "#9a98a0", marginBottom: 4 }}>TOTALE</div>
            <div className="mono" style={{ fontSize: 16, fontWeight: 700, color: "#e8e6df" }}>{fmtEur(kpi.incassatoTotale)}</div>
          </div>
        </div>
        <div style={{ borderTop: "1px solid #1e1e2a", paddingTop: 12 }}>
          <div style={{ fontSize: 11, color: "#5a5868", marginBottom: 10 }}>QUESTO MESE</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
            <div>
              <div style={{ fontSize: 10, color: "#5a5868", marginBottom: 4 }}>Confermato</div>
              <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#4ecb8d" }}>{fmtEur(kpi.incassatoConfermatoMese)}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, color: "#5a5868", marginBottom: 4 }}>In corso</div>
              <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#378ADD" }}>{fmtEur(kpi.incassatoAttesoMese)}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, color: "#5a5868", marginBottom: 4 }}>% su scaduto</div>
              <div className="mono" style={{ fontSize: 14, fontWeight: 600, color: "#EF9F27" }}>{kpi.pctTotaleMese.toFixed(1)}%</div>
            </div>
          </div>
        </div>
      </div>

      {insolutiAperti.length > 0 && (
        <div style={{ background: "#2a0a0a", border: "1px solid #4a1a1a", borderRadius: 12, padding: "12px 14px", marginBottom: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 500, color: "#e24b4a", marginBottom: 8 }}>⚠ Insoluti con scadenza 60gg</div>
          {insolutiAperti.map(i => {
            const gg = giorniDa(i.data);
            const rimanenti = 60 - gg;
            return <div key={i.id} style={{ fontSize: 12, color: rimanenti < 10 ? "#e24b4a" : "#EF9F27", marginBottom: 4 }}>{i.ragione} · {fmtEur(i.importo)} · {rimanenti > 0 ? `${rimanenti}gg rimasti` : "SCADUTO"}</div>;
          })}
        </div>
      )}

      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: "#9a98a0" }}>Distribuzione stati</div>
        {STATI_OP.filter(s => byStato[s] > 0).map(s => (
          <div key={s} style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 8, height: 8, borderRadius: "50%", background: STATO_COLORS[s], display: "inline-block" }} />{s}</span>
              <span className="mono" style={{ color: "#9a98a0" }}>{byStato[s]}</span>
            </div>
            <div style={{ background: "#1a1a25", borderRadius: 4, height: 4, overflow: "hidden" }}>
              <div style={{ width: `${(byStato[s] / clienti.length) * 100}%`, height: "100%", background: STATO_COLORS[s], borderRadius: 4 }} />
            </div>
          </div>
        ))}
      </div>

      <div className="card">
        <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 14, color: "#9a98a0" }}>Scaduto per agente</div>
        {agentiKpi.map((ag, i) => (
          <div key={ag.nome} onClick={() => ag.id && onSelectAgente(ag.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < agentiKpi.length - 1 ? "1px solid #1a1a25" : "none", cursor: ag.id ? "pointer" : "default" }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 500 }}>{ag.nome}</div>
              <div style={{ fontSize: 11, color: "#5a5868" }}>{ag.nClienti} clienti{ag.totMano > 0 ? ` · 💼 ${fmtEur(ag.totMano)} in mano` : ""}{ag.hasAlert ? " ⚠" : ""}</div>
            </div>
            <div className="mono" style={{ fontWeight: 600, color: "#e24b4a" }}>{fmtEur(ag.totScaduto)}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

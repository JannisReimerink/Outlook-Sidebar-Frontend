/* FMC Fuel Release Agent -- Taskpane (Phase 1: nur lesen & vorschlagen) */

Office.onReady(() => {
  wireUpUI();
  tryAutofillFromMail();
});

function wireUpUI() {
  document.getElementById("lookupBtn").addEventListener("click", onLookupClicked);
}

// --- Schritt A: Versuch, A/C-REG und Airport grob aus der geoeffneten Mail zu ziehen ---
// Bewusst simpel gehalten (Regex-Heuristik). Die eigentliche, verlaessliche
// Extraktion aus Freitext ist ein spaeterer Ausbauschritt (LLM-Extraktion,
// siehe Konzeptpapier Abschnitt 5) -- hier nur ein Startwert zum manuellen Pruefen.
function tryAutofillFromMail() {
  const item = Office.context.mailbox && Office.context.mailbox.item;
  const hint = document.getElementById("autofill-hint");
  if (!item) {
    hint.textContent = "Keine Mail geoeffnet -- Felder bitte manuell eintragen.";
    return;
  }

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    const subject = item.subject || "";
    const body = result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
    const haystack = `${subject}\n${body}`;

    const acRegMatch = haystack.match(/\b([A-Z]{1,2}-[A-Z0-9]{3,5})\b/);
    const icaoMatch = haystack.match(/\b([A-Z]{4})\b/);

    let found = [];
    if (acRegMatch) {
      document.getElementById("acReg").value = acRegMatch[1];
      found.push("A/C-REG");
    }
    if (icaoMatch) {
      document.getElementById("airport").value = icaoMatch[1];
      found.push("Airport");
    }

    hint.textContent = found.length
      ? `Automatisch erkannt: ${found.join(", ")} -- bitte pruefen.`
      : "Konnte nichts automatisch erkennen -- bitte manuell eintragen.";
  });
}

// --- Schritt B: Lookup ausloesen ---
async function onLookupClicked() {
  const acReg = document.getElementById("acReg").value.trim();
  const airport = document.getElementById("airport").value.trim();

  hide("result-section");
  hide("flagBanner");

  if (!acReg || !airport) {
    showStatus("Bitte A/C-REG und Airport eintragen.", true);
    return;
  }

  setLoading(true);
  showStatus(`Suche Vertragsdaten fuer ${acReg} @ ${airport} in der PAP ...`);

  try {
    const data = await fetchPapLookup(acReg, airport);
    hide("status-section");
    renderResult(data);
  } catch (err) {
    showStatus(`Fehler beim Lookup: ${err.message}`, true);
  } finally {
    setLoading(false);
  }
}

// --- Schritt C: Backend-Aufruf (Excel-Automatisierung der PAP) ---
// Solange der Dienst noch nicht existiert, liefert diese Funktion ein Mock-Ergebnis,
// damit die Oberflaeche end-to-end testbar ist. USE_MOCK in config.js umschalten,
// sobald der echte Endpoint (naechster Baustein) steht.
async function fetchPapLookup(acReg, airport) {
  if (window.FMC_CONFIG.USE_MOCK) {
    return mockPapLookup(acReg, airport);
  }

  const resp = await fetch(`${window.FMC_CONFIG.API_BASE_URL}/pap-lookup`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ acReg, airport }),
  });
  if (!resp.ok) {
    throw new Error(`Dienst antwortete mit Status ${resp.status}`);
  }
  return resp.json();
}

// Simulierte Antwort -- Feldnamen entsprechen 1:1 der Bauplan-Tabelle (Abschnitt 4).
function mockPapLookup(acReg, airport) {
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve({
        inScope: true,
        kunde: "Beispielkunde GmbH",
        acType: "Citation CJ3",
        iata: airport.length === 3 ? airport : "???",
        icao: airport.length === 4 ? airport : "????",
        row6: {
          supplier: "AEG Fuels",
          netto: "2.14 USD/USG",
          total: "2.61 USD/USG",
          totalEur: "0.71 EUR/ltr",
          payment: "...Fuel Release from AEG...",
          intoplane: "SASCA",
          basis: "CIF MED mean",
          addon: "0.47 USD/USG",
          remarks: "send schedule to request@fuel-more.com",
          validity: "01-JAN-2026 to 31-DEC-2026",
          treffer: "1",
        },
        options: [
          // Beispiel Mehr-Optionen-Fall; leer lassen, wenn nur eine Option existiert.
        ],
      });
    }, 600);
  });
}

// --- Rendering ---
function renderResult(data) {
  if (data.inScope === false) {
    showFlag("Operator ist laut D5 nicht im Scope -- Sonderbehandlung noetig.", "blocked");
  }

  const r = data.row6 || {};
  const hasContract = r.supplier && r.supplier !== "-" && r.payment;

  if (!hasContract) {
    showFlag("Kein Vertrag an diesem Airport -- bitte in den Recherche-Korb (guenstigste-Option-Prozess).", "blocked");
  } else {
    const optionCount = (data.options || []).filter(o => o.supplier && o.supplier !== "-").length + 1;
    if (optionCount > 1) {
      showFlag(`${optionCount} vertragliche Optionen an diesem Airport -- Entscheidung noetig.`, "warning");
    } else if (r.payment && r.payment.toLowerCase().includes("contract card")) {
      showFlag("Laut Vertrag ist eine Kartenzahlung moeglich -- Rueckfrage statt Release pruefen.", "warning");
    }
  }

  setText("r-kunde", data.kunde);
  setText("r-actype", data.acType);
  setText("r-supplier", r.supplier);
  setText("r-netto", r.netto);
  setText("r-total", r.total);
  setText("r-totaleur", r.totalEur);
  setText("r-payment", r.payment);
  setText("r-intoplane", r.intoplane);
  setText("r-basis", r.basis);
  setText("r-addon", r.addon);
  setText("r-remarks", r.remarks);
  setText("r-validity", r.validity);
  setText("r-treffer", r.treffer);

  const optionsBlock = document.getElementById("optionsBlock");
  const tbody = document.getElementById("optionsTableBody");
  tbody.innerHTML = "";
  if (data.options && data.options.length > 0) {
    data.options.forEach((o) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escapeHtml(o.supplier || "-")}</td>
                      <td>${escapeHtml(o.total || "-")}</td>
                      <td>${escapeHtml(o.payment || "-")}</td>
                      <td>${escapeHtml(o.remarks || "-")}</td>`;
      tbody.appendChild(tr);
    });
    show("optionsBlock");
  } else {
    optionsBlock.classList.add("hidden");
  }

  show("result-section");
}

// --- kleine Helfer ---
function setText(id, value) {
  document.getElementById(id).textContent = value || "-";
}
function show(id) { document.getElementById(id).classList.remove("hidden"); }
function hide(id) { document.getElementById(id).classList.add("hidden"); }
function showStatus(text, isError) {
  const section = document.getElementById("status-section");
  const el = document.getElementById("statusText");
  el.textContent = text;
  el.style.color = isError ? "#b3261e" : "#1a1a1a";
  section.classList.remove("hidden");
}
function showFlag(text, kind) {
  const el = document.getElementById("flagBanner");
  el.textContent = text;
  el.className = kind === "blocked" ? "flag-blocked" : "flag-warning";
  el.classList.remove("hidden");
}
function setLoading(isLoading) {
  document.getElementById("lookupBtn").disabled = isLoading;
}
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

// Zentrale Konfiguration der Taskpane.
// API_BASE_URL zeigt auf den Dienst, der die Excel-Automatisierung der PAP kapselt
// (naechster Baustein: "PAP-Lookup-Skript"). Bis der Dienst steht, liefert
// taskpane.js einen Mock zurueck, wenn USE_MOCK = true ist.
window.FMC_CONFIG = {
  API_BASE_URL: "https://REPLACE_WITH_HOSTING_DOMAIN/api",
  USE_MOCK: true
};

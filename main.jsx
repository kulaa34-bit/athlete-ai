import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';

// ── Supabase init ──────────────────────────────────────────────
// Set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY in Netlify env vars
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (SUPABASE_URL && SUPABASE_KEY) {
  import("@supabase/supabase-js").then(({ createClient }) => {
    window.__supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
    console.log("✅ Supabase conectado");
  }).catch(e => console.warn("Supabase não carregado:", e));
} else {
  console.warn("⚡ Modo demo: Supabase não configurado. Dados salvos localmente.");
}

// ── Anthropic API proxy ─────────────────────────────────────────
// Em produção, as chamadas vão para /.netlify/functions/ai
// Em dev local, vão direto para api.anthropic.com
// A lógica já está em IS_PROD/AI_URL no App.jsx

createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

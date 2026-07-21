import React, { useState, useEffect, useRef } from "react";
import { AreaChart, Area, LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

/* ── GLOBAL STYLES ───────────────────────────────────────────────── */
(function(){
  const s = document.createElement("style");
  s.textContent = `
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
    @keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    @keyframes slideUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:translateY(0)}}
    @keyframes popIn{0%{transform:scale(.7);opacity:0}80%{transform:scale(1.07)}100%{transform:scale(1);opacity:1}}
    @keyframes pulse{0%,100%{opacity:1}50%{opacity:.35}}
    @keyframes spin{from{stroke-dashoffset:276}to{stroke-dashoffset:0}}
    @keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-7px)}}
    .fade{animation:fadeIn .28s ease both}
    .slide{animation:slideUp .32s ease both}
    .pop{animation:popIn .38s cubic-bezier(.34,1.56,.64,1) both}
    input:focus,textarea:focus{outline:none}
    button:active{transform:scale(.97)}
    ::-webkit-scrollbar{width:3px;height:3px}
    ::-webkit-scrollbar-thumb{background:#334155;border-radius:3px}
  `;
  document.head.appendChild(s);
})();

/* ── DESIGN TOKENS ───────────────────────────────────────────────── */
const B = {
  blue:"#2563EB", green:"#22C55E", orange:"#F97316", red:"#EF4444",
  emerald:"#10B981", purple:"#7C3AED", indigo:"#6366F1", cyan:"#06B6D4",
  yellow:"#FACC15",
};
const SPORT_COLOR = { strength:B.purple, running:B.blue, bike:B.green, swim:B.cyan, rest:"#475569" };
const D = {
  bg:"#020617", bg2:"#0F172A", card:"#1E293B", bar:"#111827",
  text:"#F8FAFC", sub:"#CBD5E1", muted:"#64748B", border:"#334155",
};

/* ── EXERCISE DATABASE ───────────────────────────────────────────── */
const EXERCISES = {
  chest_shoulder:[
    {id:"cs1",name:"Supino Reto com Barra",muscles:"Peitoral · Tríceps · Deltóide Ant.",rest:90,tip:"Descida 3s. Escápulas retraídas."},
    {id:"cs2",name:"Supino Inclinado Halteres",muscles:"Peitoral Superior",rest:75,tip:"Banco 30-45°. Amplitude total."},
    {id:"cs3",name:"Crucifixo com Halteres",muscles:"Peitoral (estiramento)",rest:60,tip:"Leve flexão no cotovelo."},
    {id:"cs4",name:"Press Militar",muscles:"Deltóide · Tríceps",rest:90,tip:"Core travado. Barra ao queixo."},
    {id:"cs5",name:"Elevação Lateral",muscles:"Deltóide Lateral",rest:60,tip:"Sem balanço. 1s no topo."},
    {id:"cs6",name:"Elevação Frontal",muscles:"Deltóide Anterior",rest:60,tip:"Movimento isolado."},
  ],
  legs:[
    {id:"lg1",name:"Agachamento Livre",muscles:"Quad · Glúteos · Isquios",rest:120,tip:"Profundidade paralela ou abaixo."},
    {id:"lg2",name:"Leg Press 45°",muscles:"Quad · Glúteos",rest:90,tip:"Não trave os joelhos."},
    {id:"lg3",name:"Cadeira Extensora",muscles:"Quadríceps",rest:60,tip:"1s no topo. Descida em 3s."},
    {id:"lg4",name:"Afundo com Halteres",muscles:"Glúteos · Quad",rest:60,tip:"Tronco ereto."},
    {id:"lg5",name:"Panturrilha em Pé",muscles:"Gastrocnêmio · Sóleo",rest:60,tip:"Amplitude máxima."},
  ],
  back_biceps:[
    {id:"bb1",name:"Barra Fixa / Puxada Alta",muscles:"Dorsal · Bíceps",rest:90,tip:"Cotovelos ao chão. Peitoral ao pull."},
    {id:"bb2",name:"Remada Curvada",muscles:"Costas Média · Trapézio",rest:90,tip:"Tronco a 45°. Puxe para o umbigo."},
    {id:"bb3",name:"Remada Unilateral",muscles:"Dorsal · Rombóides",rest:75,tip:"Cotovelo junto ao tronco."},
    {id:"bb4",name:"Rosca Direta",muscles:"Bíceps Braquial",rest:75,tip:"Descida em 3s."},
    {id:"bb5",name:"Rosca Alternada",muscles:"Bíceps · Braquial",rest:60,tip:"Supinação no topo."},
  ],
  posterior:[
    {id:"pt1",name:"Terra Romeno (Stiff)",muscles:"Isquios · Glúteos · Lombar",rest:90,tip:"Quadril empurra para trás."},
    {id:"pt2",name:"Mesa Flexora",muscles:"Isquiotibiais",rest:75,tip:"Descida em 4s."},
    {id:"pt3",name:"Cadeira Flexora",muscles:"Isquios (distal)",rest:60,tip:"3s excêntrica."},
    {id:"pt4",name:"Glúteo no Cabo",muscles:"Glúteo Máximo",rest:60,tip:"Squeeze 1s no topo."},
    {id:"pt5",name:"Abdução de Quadril",muscles:"Glúteo Médio",rest:60,tip:"Controlado nos dois sentidos."},
  ],
  chest_triceps:[
    {id:"ct1",name:"Supino Reto",muscles:"Peitoral · Tríceps",rest:90,tip:"Controle a descida."},
    {id:"ct2",name:"Tríceps Pulley",muscles:"Tríceps (todos os feixes)",rest:60,tip:"Cotovelo fixo."},
    {id:"ct3",name:"Tríceps Francês",muscles:"Tríceps Longo",rest:60,tip:"Cotovelo para cima."},
    {id:"ct4",name:"Mergulho (Paralela)",muscles:"Peitoral Inferior · Tríceps",rest:75,tip:"Incline levemente o tronco."},
    {id:"ct5",name:"Peck Deck",muscles:"Peitoral",rest:60,tip:"Máxima contração."},
  ],
  back_biceps2:[
    {id:"bk1",name:"Pulldown Pegada Neutra",muscles:"Dorsal",rest:90,tip:"Deprima a escápula."},
    {id:"bk2",name:"Remada Cavalinho",muscles:"Costas Média",rest:75,tip:"Cotovelo alto."},
    {id:"bk3",name:"Rosca Scott",muscles:"Bíceps",rest:60,tip:"Amplitude total."},
    {id:"bk4",name:"Rosca Martelo",muscles:"Braquial",rest:60,tip:"Pegada neutra."},
  ],
  full_body:[
    {id:"fb1",name:"Agachamento",muscles:"Quad · Glúteos",rest:90,tip:"Profundidade paralela."},
    {id:"fb2",name:"Supino",muscles:"Peitoral",rest:90,tip:"Escápulas retraídas."},
    {id:"fb3",name:"Puxada Alta",muscles:"Dorsal · Bíceps",rest:90,tip:"Cotovelos ao chão."},
    {id:"fb4",name:"Press Ombros",muscles:"Deltóide",rest:75,tip:"Core ativado."},
    {id:"fb5",name:"Rosca Direta",muscles:"Bíceps",rest:60,tip:"Sem balanço."},
  ],
};

/* ── SPORTS META ─────────────────────────────────────────────────── */
const SPORTS = [
  { id:"strength", label:"Musculação", emoji:"💪", color:B.purple, desc:"Hipertrofia, força, recomposição corporal" },
  { id:"running",  label:"Corrida",    emoji:"🏃", color:B.blue,   desc:"Provas de rua, condicionamento, pace" },
  { id:"bike",     label:"Ciclismo",   emoji:"🚴", color:B.green,  desc:"Estrada, MTB, spinning, triathlon" },
  { id:"swim",     label:"Natação",    emoji:"🏊", color:B.cyan,   desc:"Crawl, costas, resistência, técnica" },
];

/* ── UTILS ───────────────────────────────────────────────────────── */
async function compressImage(dataURL, maxDim=480) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const r = Math.min(maxDim/img.width, maxDim/img.height, 1);
      const c = document.createElement("canvas");
      c.width = Math.round(img.width*r); c.height = Math.round(img.height*r);
      c.getContext("2d").drawImage(img,0,0,c.width,c.height);
      resolve(c.toDataURL("image/jpeg",0.72));
    };
    img.src = dataURL;
  });
}
function readFile(file) {
  return new Promise(r => { const fr=new FileReader(); fr.onload=e=>r(e.target.result); fr.readAsDataURL(file); });
}
const dayName = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];

/* ══ DATABASE ABSTRACTION ════════════════════════════════════════
   Prioridade de storage:
   1. Supabase (produção — window.__supabase definido pelo main.jsx)
   2. window.storage (Claude artifact)
   3. localStorage (dev local / PWA sem Supabase)
════════════════════════════════════════════════════════════════ */
const DB = {
  _sb: () => window.__supabase || null,
  _user: null,

  /* ── Storage abstraction ── */
  async _get(k) {
    if (window.storage) { try { const r = await window.storage.get(k); return r?.value ?? null; } catch { return null; } }
    try { return localStorage.getItem(k); } catch { return null; }
  },
  async _set(k, v) {
    if (window.storage) { try { await window.storage.set(k, v); return; } catch {} return; }
    try { localStorage.setItem(k, v); } catch {}
  },
  async _del(k) {
    if (window.storage) { try { await window.storage.delete(k); return; } catch {} return; }
    try { localStorage.removeItem(k); } catch {}
  },
  async _keys(prefix) {
    if (window.storage) { try { const r = await window.storage.list(prefix||""); return r?.keys||[]; } catch { return []; } }
    try { return Object.keys(localStorage).filter(k=>!prefix||k.startsWith(prefix)); } catch { return []; }
  },

  /* ── AUTH ── */
  async signIn(email, password) {
    const sb = DB._sb();
    if (sb) {
      try {
        const r = await sb.auth.signInWithPassword({ email, password });
        if (!r.error) DB._user = r.data.user;
        return r;
      } catch(e) {
        return { error:{ message:"Supabase inacessível. Modo demo ativado automaticamente." }, _fallback:true };
      }
    }
    const raw = await DB._get("_au");
    const users = JSON.parse(raw||"{}");
    const acc = users[email];
    // Demo mode: auto-cria conta se não existir (permite testar sem cadastro prévio)
    if (!acc) {
      const u = { id:"demo_"+btoa(unescape(encodeURIComponent(email))).replace(/[^a-z0-9]/gi,"").slice(0,12)+"_"+Date.now().toString(36), email };
      users[email] = { pwd:password, u };
      await DB._set("_au", JSON.stringify(users));
      await DB._set("_as", JSON.stringify(u));
      DB._user = u;
      return { data:{ user:u }, error:null };
    }
    if (acc.pwd !== password) return { error:{ message:"Senha incorreta." } };
    await DB._set("_as", JSON.stringify(acc.u));
    DB._user = acc.u;
    return { data:{ user:acc.u }, error:null };
  },

  async signUp(email, password) {
    const sb = DB._sb();
    if (sb) {
      try {
        const r = await sb.auth.signUp({ email, password });
        if (!r.error) DB._user = r.data.user;
        return r;
      } catch(e) {
        // Supabase unreachable — fall through to demo mode
      }
    }
    const raw = await DB._get("_au");
    const users = JSON.parse(raw||"{}");
    if (users[email]) return { error:{ message:"Email já cadastrado. Use Entrar." } };
    const u = { id:"demo_"+btoa(unescape(encodeURIComponent(email))).replace(/[^a-z0-9]/gi,"").slice(0,12)+"_"+Date.now().toString(36), email };
    users[email] = { pwd:password, u };
    await DB._set("_au", JSON.stringify(users));
    await DB._set("_as", JSON.stringify(u));
    DB._user = u;
    return { data:{ user:u }, error:null };
  },

  async getSession() {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.auth.getSession();
      const u = data?.session?.user || null;
      DB._user = u;
      return u;
    }
    const raw = await DB._get("_as");
    const u = raw ? JSON.parse(raw) : null;
    DB._user = u;
    return u;
  },

  async signOut() {
    const sb = DB._sb();
    DB._user = null;
    if (sb) return sb.auth.signOut();
    await DB._del("_as");
  },

  async resetPassword(email) {
    const sb = DB._sb();
    if (sb) return sb.auth.resetPasswordForEmail(email, { redirectTo:window.location.origin });
    return { error:{ message:"Disponível apenas com Supabase configurado." } };
  },

  /* ── PROFILE ── */
  async getProfile(uid) {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.from("profiles").select("*").eq("id",uid).single();
      if (!data) return null;
      return { basic:data.basic, selectedSports:data.selected_sports, sportData:data.sport_data, schedule:data.schedule, availability:data.availability, health:data.health, plan:data.plan, createdAt:data.created_at };
    }
    const raw = await DB._get(`_p_${uid}`);
    return raw ? JSON.parse(raw) : null;
  },

  async saveProfile(uid, profile) {
    const sb = DB._sb();
    if (sb) return sb.from("profiles").upsert({ id:uid, basic:profile.basic, selected_sports:profile.selectedSports, sport_data:profile.sportData, schedule:profile.schedule, availability:profile.availability, health:profile.health, plan:profile.plan, updated_at:new Date().toISOString() });
    await DB._set(`_p_${uid}`, JSON.stringify(profile));
  },

  async deleteProfile(uid) {
    const sb = DB._sb();
    if (sb) return sb.from("profiles").delete().eq("id",uid);
    await DB._del(`_p_${uid}`);
  },

  /* ── WORKOUTS ── */
  async getWorkouts(uid) {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.from("workout_sessions").select("*").eq("user_id",uid).order("created_at",{ascending:true});
      return (data||[]).map(s=>({ type:s.type, plan:s.plan, log:s.log, date:s.date }));
    }
    const raw = await DB._get(`_w_${uid}`);
    return JSON.parse(raw||"[]");
  },

  async addWorkout(uid, workout) {
    const sb = DB._sb();
    if (sb) return sb.from("workout_sessions").insert({ user_id:uid, type:workout.type, plan:workout.plan, log:workout.log, date:workout.date });
    const list = JSON.parse((await DB._get(`_w_${uid}`))||"[]");
    list.push(workout);
    await DB._set(`_w_${uid}`, JSON.stringify(list));
  },

  async removeWorkout(uid, workout) {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.from("workout_sessions").select("id,plan").eq("user_id",uid).eq("date",workout.date).eq("type",workout.type);
      const match = (data||[]).find(r=>r.plan?.focus===workout.plan?.focus);
      if (match) await sb.from("workout_sessions").delete().eq("id",match.id);
      return;
    }
    const list = JSON.parse((await DB._get(`_w_${uid}`))||"[]");
    await DB._set(`_w_${uid}`, JSON.stringify(list.filter(w=>!(w.date===workout.date&&w.type===workout.type&&w.plan?.focus===workout.plan?.focus))));
  },

  async clearWorkouts(uid) {
    const sb = DB._sb();
    if (sb) return sb.from("workout_sessions").delete().eq("user_id",uid);
    await DB._del(`_w_${uid}`);
  },

  /* ── BODY WEIGHTS ── */
  async getWeights(uid) {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.from("body_weights").select("*").eq("user_id",uid).order("iso_date",{ascending:true});
      return (data||[]).map(w=>({ iso:w.iso_date, date:w.date, v:w.weight }));
    }
    const keys = await DB._keys(`_bw_${uid}_`);
    const items = await Promise.all(keys.map(async k=>{ const r=await DB._get(k); return r?JSON.parse(r):null; }));
    return items.filter(Boolean).sort((a,b)=>a.iso.localeCompare(b.iso));
  },

  async saveWeight(uid, entry) {
    const sb = DB._sb();
    if (sb) return sb.from("body_weights").upsert({ user_id:uid, iso_date:entry.iso, date:entry.date, weight:entry.v },{ onConflict:"user_id,iso_date" });
    await DB._set(`_bw_${uid}_${entry.iso}`, JSON.stringify(entry));
  },

  /* ── PHOTO ASSESSMENTS ── */
  async getAssessments(uid) {
    const sb = DB._sb();
    if (sb) {
      const { data } = await sb.from("photo_assessments").select("*").eq("user_id",uid).order("created_at",{ascending:false});
      return (data||[]).map(a=>({ id:a.id, date:a.date, isoDate:a.iso_date, photos:a.photos }));
    }
    const keys = await DB._keys(`_ph_${uid}_`);
    const items = await Promise.all(keys.map(async k=>{ const r=await DB._get(k); return r?JSON.parse(r):null; }));
    return items.filter(Boolean).sort((a,b)=>(b.isoDate||"").localeCompare(a.isoDate||""));
  },

  async saveAssessment(uid, assessment) {
    const sb = DB._sb();
    if (sb) return sb.from("photo_assessments").insert({ user_id:uid, date:assessment.date, iso_date:assessment.isoDate, photos:assessment.photos });
    await DB._set(`_ph_${uid}_${assessment.isoDate}`, JSON.stringify(assessment));
  },
};

/* ── SHARED COMPONENTS ───────────────────────────────────────────── */
function Card({ children, style={}, accent }) {
  return <div style={{ background:D.card, borderRadius:16, padding:20, border:`1px solid ${D.border}`, ...(accent?{borderTop:`3px solid ${accent}`}:{}), ...style }}>{children}</div>;
}
function Lbl({ children }) {
  return <div style={{ fontSize:10, fontWeight:700, color:D.muted, textTransform:"uppercase", letterSpacing:".08em", marginBottom:8 }}>{children}</div>;
}
function SportBadge({ type, sm=false }) {
  const c = SPORT_COLOR[type]||"#475569";
  const s = SPORTS.find(s=>s.id===type)||{label:"Descanso",emoji:"😴"};
  return <span style={{ display:"inline-flex", alignItems:"center", gap:5, background:`${c}20`, color:c, padding:sm?"3px 9px":"5px 12px", borderRadius:20, fontSize:sm?10:11, fontWeight:700 }}>{s.emoji} {s.label}</span>;
}
function TipBox({ active, payload, label, unit="" }) {
  if(!active||!payload?.length) return null;
  return <div style={{ background:D.bg2, border:`1px solid ${D.border}`, borderRadius:10, padding:"8px 12px", fontSize:12 }}><div style={{ color:D.muted, marginBottom:3 }}>{label}</div>{payload.map((p,i)=><div key={i} style={{ color:p.color||B.blue, fontWeight:700 }}>{p.value}{unit}</div>)}</div>;
}
function AdjBtn({ onPress, label, accent=false }) {
  return <button onClick={onPress} style={{ width:36, height:36, borderRadius:10, border:`1.5px solid ${accent?B.purple:D.border}`, background:accent?`${B.purple}20`:D.bg2, color:accent?B.purple:D.muted, fontWeight:800, fontSize:18, cursor:"pointer", fontFamily:"'Inter',sans-serif", display:"flex", alignItems:"center", justifyContent:"center" }}>{label}</button>;
}
function FInp({ lbl, val, ch, ph, tp="text", min, max, step }) {
  const [foc,setFoc] = useState(false);
  return (
    <div>
      <Lbl>{lbl}</Lbl>
      <input type={tp} value={val} onChange={e=>ch(e.target.value)} placeholder={ph}
        min={min} max={max} step={step}
        onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)}
        style={{ width:"100%", background:D.bg2, border:`1.5px solid ${foc?B.blue:D.border}`, borderRadius:12, color:D.text, padding:"12px 14px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box", transition:"border-color .2s" }}/>
    </div>
  );
}
function SelBtn({ opts, val, onSel, color=B.blue }) {
  return (
    <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
      {opts.map(o=>(
        <button key={o.v} onClick={()=>onSel(o.v)} style={{ padding:"9px 14px", borderRadius:10, cursor:"pointer", border:`1.5px solid ${val===o.v?color:D.border}`, background:val===o.v?`${color}20`:D.card, color:val===o.v?color:D.muted, fontWeight:700, fontSize:12, fontFamily:"'Inter',sans-serif" }}>{o.l}</button>
      ))}
    </div>
  );
}
function CBox({ title, color, children }) {
  return <Card style={{ borderTop:`3px solid ${color}` }}><Lbl>{title}</Lbl>{children}</Card>;
}

/* ══ AUTH SCREEN ════════════════════════════════════════════════ */
function AuthScreen({ onAuth }) {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [resetSent, setResetSent] = useState(false);

  const isDemo = !window.__supabase;

  async function handleSubmit() {
    if (!email.trim() || !password) { setError("Preencha email e senha."); return; }
    if (password.length < 6) { setError("Senha mínima de 6 caracteres."); return; }
    setLoading(true); setError("");
    try {
      const r = mode === "register"
        ? await DB.signUp(email.trim(), password)
        : await DB.signIn(email.trim(), password);
      setLoading(false);
      if (r?.error) { setError(r.error.message); return; }
      onAuth(r.data.user);
    } catch(err) {
      setLoading(false);
      const msg = String(err?.message||err||"");
      if (msg.toLowerCase().includes("fetch") || msg.toLowerCase().includes("network") || msg.toLowerCase().includes("failed")) {
        setError("Sem conexão com o Supabase. Verifique VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY nas variáveis de ambiente do Netlify.");
      } else if (msg.toLowerCase().includes("storage") || msg.toLowerCase().includes("access") || msg.toLowerCase().includes("insecure")) {
        setError("Erro de storage. Recarregue a página e tente novamente.");
      } else {
        setError(msg || "Erro inesperado. Tente novamente.");
      }
    }
  }

  async function handleReset() {
    if (!email.trim()) { setError("Digite seu email acima primeiro."); return; }
    setLoading(true);
    const r = await DB.resetPassword(email.trim());
    setLoading(false);
    if (r.error) { setError(r.error.message); return; }
    setResetSent(true);
  }

  return (
    <div style={{ minHeight:"100vh", background:`radial-gradient(ellipse at 30% 20%, #1e3a8a22 0%, ${D.bg} 65%)`, fontFamily:"'Inter',sans-serif", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"24px 20px" }}>
      {/* Logo */}
      <div className="pop" style={{ width:72, height:72, borderRadius:20, background:`linear-gradient(135deg,${B.blue},${B.indigo})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:34, marginBottom:16, boxShadow:`0 12px 40px ${B.blue}55` }}>🏆</div>
      <div style={{ fontSize:28, fontWeight:900, color:D.text, letterSpacing:"-1px", marginBottom:4 }}>Athlete AI</div>
      <div style={{ fontSize:13, color:D.muted, marginBottom:36 }}>Coach esportivo com IA — honesto e sem enrolação</div>

      <div style={{ width:"100%", maxWidth:380 }}>
        {/* Mode toggle */}
        <div style={{ display:"flex", background:D.card, borderRadius:12, padding:4, marginBottom:22 }}>
          {[["login","Entrar"],["register","Criar conta"]].map(([m,l])=>(
            <button key={m} onClick={()=>{ setMode(m); setError(""); setResetSent(false); }} style={{ flex:1, padding:"10px 0", borderRadius:9, background:mode===m?D.bg2:"transparent", color:mode===m?D.text:D.muted, fontWeight:mode===m?700:400, fontSize:13, cursor:"pointer", border:"none", fontFamily:"'Inter',sans-serif", transition:"all .15s" }}>{l}</button>
          ))}
        </div>

        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <div>
            <Lbl>Email</Lbl>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="seu@email.com"
              style={{ width:"100%", background:D.card, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"13px 16px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/>
          </div>
          <div>
            <Lbl>Senha</Lbl>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Mínimo 6 caracteres"
              onKeyDown={e=>e.key==="Enter"&&handleSubmit()}
              style={{ width:"100%", background:D.card, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"13px 16px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/>
          </div>

          {error && <div style={{ background:`${B.red}15`, color:B.red, borderRadius:10, padding:"10px 14px", fontSize:13, border:`1px solid ${B.red}30` }}>{error}</div>}
          {resetSent && <div style={{ background:`${B.emerald}15`, color:B.emerald, borderRadius:10, padding:"10px 14px", fontSize:13 }}>✓ Email de redefinição enviado.</div>}

          <button onClick={handleSubmit} disabled={loading} style={{ padding:"15px 0", borderRadius:12, border:"none", background:loading?"#334155":B.blue, color:"#fff", fontWeight:800, fontSize:15, cursor:loading?"not-allowed":"pointer", fontFamily:"'Inter',sans-serif", boxShadow:loading?"none":`0 4px 20px ${B.blue}44`, transition:"all .15s" }}>
            {loading ? "⏳ Aguarde..." : mode==="login" ? "Entrar →" : "Criar conta →"}
          </button>

          {mode==="login" && (
            <button onClick={handleReset} disabled={loading} style={{ background:"transparent", border:"none", color:D.muted, fontSize:12, cursor:"pointer", fontFamily:"'Inter',sans-serif", padding:"4px 0" }}>
              Esqueci minha senha
            </button>
          )}
        </div>

        {isDemo && (
          <div style={{ marginTop:20, fontSize:11, color:D.muted, textAlign:"center", lineHeight:1.7, background:D.card, borderRadius:10, padding:"10px 14px", border:`1px solid ${D.border}` }}>
            ⚡ <strong style={{color:D.sub}}>Modo demo</strong> — Supabase não detectado.<br/>
            Dados salvos localmente neste dispositivo.<br/>
            <span style={{color:B.emerald}}>Qualquer email/senha funciona — conta criada automaticamente.</span>
          </div>
        )}
      </div>
    </div>
  );
}

/* ══ SPLASH ══════════════════════════════════════════════════════ */
function Splash() {
  return (
    <div style={{ minHeight:"100vh", background:`radial-gradient(ellipse at 30% 20%, #1e3a8a 0%, #020617 65%)`, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:"'Inter',sans-serif", gap:28 }}>
      <div className="pop" style={{ width:96, height:96, borderRadius:28, background:`linear-gradient(135deg,${B.blue},${B.indigo})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:46, boxShadow:`0 0 0 1px ${B.blue}30, 0 20px 60px ${B.blue}55` }}>🏆</div>
      <div className="fade" style={{ textAlign:"center" }}>
        <div style={{ fontSize:36, fontWeight:900, color:D.text, letterSpacing:"-1.5px" }}>Athlete AI</div>
        <div style={{ fontSize:12, color:B.blue, fontWeight:600, marginTop:6, letterSpacing:"0.16em", textTransform:"uppercase" }}>Seu Coach Esportivo com IA</div>
      </div>
      <div style={{ display:"flex", gap:8 }}>
        {[0,1,2].map(i=><div key={i} style={{ width:8, height:8, borderRadius:"50%", background:i===0?B.blue:D.border, animation:`pulse 1.4s ease ${i*.22}s infinite` }}/>)}
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════
   ONBOARDING — totalmente dinâmico por esporte
══════════════════════════════════════════════════════════════════ */
function Onboarding({ onComplete }) {
  const [step, setStep] = useState(0);
  const [selectedSports, setSelectedSports] = useState([]);
  const [sportStep, setSportStep] = useState(0); // which sport question we're on
  const [aiGenerating, setAiGenerating] = useState(false);
  const [genMsg, setGenMsg] = useState("");

  // Basic profile
  const [basic, setBasic] = useState({ name:"", age:"", sex:"M", height:"", weight:"", fat:"" });
  const upBasic = (k,v) => setBasic(p=>({...p,[k]:v}));

  // Sport-specific data (filled per sport)
  const [sportData, setSportData] = useState({
    strength:{ goal:"hypertrophy", experience:"intermediate", daysPerWeek:4, hasGym:true, equipment:"completo", currentBench:"", currentSquat:"" },
    running:{ currentPace5k:"6:00", longestRun:"5", goalDistance:"10", goalTime:"", daysPerWeek:2, weeklyKm:"15" },
    bike:{ currentFTP:"", bikeType:"speed", daysPerWeek:2, weeklyKm:"", goal:"condicionamento" },
    swim:{ canSwim:true, style:"crawl", currentPace100:"2:20", weeklyVolume:"", daysPerWeek:2, goal:"resistencia" },
  });
  const upSport = (sport,k,v) => setSportData(p=>({...p,[sport]:{...p[sport],[k]:v}}));

  // Schedule
  const [schedule, setSchedule] = useState({ sessionDuration:60 });

  // Availability — days available + preferred per sport
  const ALL_DAYS = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  const [availability, setAvailability] = useState({
    availableDays: ["Seg","Ter","Qua","Qui","Sex","Sáb"], // all except Sunday by default
    preferredDays: {},  // { running: ["Qua","Dom"], strength: ["Seg","Ter","Qui","Sex"], ... }
  });

  function toggleAvailDay(day) {
    setAvailability(prev => {
      const has = prev.availableDays.includes(day);
      const newDays = has ? prev.availableDays.filter(d=>d!==day) : [...prev.availableDays, day];
      // Remove preferred days that are no longer available
      const newPref = {};
      for (const sp of Object.keys(prev.preferredDays)) {
        newPref[sp] = (prev.preferredDays[sp]||[]).filter(d=>newDays.includes(d));
      }
      return { ...prev, availableDays: newDays, preferredDays: newPref };
    });
  }

  function togglePrefDay(sport, day) {
    setAvailability(prev => {
      const curr = prev.preferredDays[sport]||[];
      const has = curr.includes(day);
      return { ...prev, preferredDays: { ...prev.preferredDays, [sport]: has ? curr.filter(d=>d!==day) : [...curr, day] } };
    });
  }

  // Health
  const [health, setHealth] = useState({ injuries:"", medications:"", limitações:"" });

  function toggleSport(id) {
    setSelectedSports(prev => prev.includes(id) ? prev.filter(s=>s!==id) : [...prev, id]);
  }

  // Calculate total steps dynamically
  // Step 0: Welcome
  // Step 1: Basic info
  // Step 2: Sport selection
  // Step 3..N: One question step per selected sport
  // Step N+1: Schedule
  // Step N+2: Health
  // Step N+3: AI generates
  const sportSteps = selectedSports; // each sport gets 1 step
  const totalSteps = 3 + sportSteps.length + 2; // welcome + basic + sportSelect + sports + schedule + health
  const currentSportInStep = step >= 3 && step < 3+sportSteps.length ? sportSteps[step-3] : null;

  function nextStep() {
    if (step < totalSteps - 2) { setStep(s=>s+1); }
    else { generatePlan(); }
  }

  async function generatePlan() {
    setStep(totalSteps - 1);
    setAiGenerating(true);
    const msgs = ["Analisando seu perfil...","Calculando periodização...","Ajustando volumes...","Sincronizando esportes...","Finalizando plano..."];
    for (const m of msgs) { setGenMsg(m); await new Promise(r=>setTimeout(r,700)); }

    const sportsDesc = selectedSports.map(s => {
      const d = sportData[s];
      if (s==="strength") return `MUSCULAÇÃO: Objetivo=${d.goal}, Experiência=${d.experience}, ${d.daysPerWeek}x/sem, Academia=${d.hasGym?"sim":"não"}, Supino=${d.currentBench||"?"}kg, Agachamento=${d.currentSquat||"?"}kg`;
      if (s==="running") return `CORRIDA: Pace 5km=${d.currentPace5k}/km, Maior distância=${d.longestRun}km, Objetivo=${d.goalDistance}km, Km/sem=${d.weeklyKm}, ${d.daysPerWeek}x/sem`;
      if (s==="bike") return `CICLISMO: FTP=${d.currentFTP||"não testado"}, Tipo=${d.bikeType}, ${d.daysPerWeek}x/sem, Km/sem=${d.weeklyKm||"?"}, Objetivo=${d.goal}`;
      if (s==="swim") return `NATAÇÃO: Pace/100m=${d.currentPace100}, Estilo=${d.style}, Volume/sem=${d.weeklyVolume||"?"}m, ${d.daysPerWeek}x/sem, Objetivo=${d.goal}`;
      return "";
    }).join("\n");

    const availDesc = `Dias DISPONÍVEIS para treinar: ${availability.availableDays.join(", ") || "Nenhum selecionado"}
Dias NÃO disponíveis: ${ALL_DAYS.filter(d=>!availability.availableDays.includes(d)).join(", ") || "Nenhum"}
Preferências por esporte:
${selectedSports.map(s=>{
  const pref = availability.preferredDays[s];
  const sp = SPORTS.find(x=>x.id===s);
  return `  ${sp?.label}: ${pref?.length>0 ? pref.join(", ") : "sem preferência específica (qualquer dia disponível)"}`;
}).join("\n")}
Duração de cada sessão: ${schedule.sessionDuration} min`;

    const prompt = `Você é um coach esportivo especializado em periodização mista. Crie um plano semanal personalizado em JSON puro (sem markdown).

ATLETA: ${basic.name||"Atleta"}, ${basic.age} anos, ${basic.sex==="M"?"Masculino":"Feminino"}, ${basic.weight}kg, ${basic.height}cm, ${basic.fat||"?"}% gordura
ESPORTES:\n${sportsDesc}
ROTINA:\n${availDesc}
LESÕES/LIMITAÇÕES: ${health.injuries||"Nenhuma"}

REGRA CRÍTICA: Respeite OBRIGATORIAMENTE os dias disponíveis e as preferências de dias por esporte. Se o atleta prefere correr na Quarta e Domingo, coloque corrida nesses dias. Se não pode treinar na Segunda, deixe Segunda como rest.

Responda SOMENTE com este JSON:
{
  "planName": "Nome do plano personalizado",
  "description": "Descrição de 2 frases explicando a lógica do plano",
  "weeklySchedule": [
    {"day":"Dom","type":"rest"},
    {"day":"Seg","type":"strength|running|bike|swim|rest","focus":"Nome do treino","distance":0,"paceTarget":"","hrZone":"","notes":""},
    {"day":"Ter","type":"...","focus":"..."},
    {"day":"Qua","type":"...","focus":"..."},
    {"day":"Qui","type":"...","focus":"..."},
    {"day":"Sex","type":"...","focus":"..."},
    {"day":"Sáb","type":"...","focus":"..."}
  ],
  "aiCoachNote": "Mensagem motivacional personalizada de 2-3 frases"
}

FOCOS VÁLIDOS — Musculação: 'Peito & Ombros', 'Pernas', 'Costas & Bíceps', 'Posterior', 'Corpo Todo'. Corrida: 'Corrida Fácil', 'Tempo Run', 'Longão', 'Regenerativo', 'Intervalado'. Bike: 'Endurance', 'Intervalado FTP', 'Longão Bike'. Natação: 'Técnica', 'Intervalado', 'Resistência'.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1500, messages:[{ role:"user", content:prompt }] })
      });
      const data = await res.json();
      const raw = data.content?.[0]?.text || "";
      const jsonMatch = raw.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const plan = JSON.parse(jsonMatch[0]);
        onComplete({ basic, selectedSports, sportData, schedule, availability, health, plan, createdAt:new Date().toISOString() });
        return;
      }
    } catch {}

    // Fallback plan if API fails — respects availability
    const fallback = buildFallbackPlan(selectedSports, sportData, basic, availability);
    onComplete({ basic, selectedSports, sportData, schedule, availability, health, plan:fallback, createdAt:new Date().toISOString() });
  }

  function buildFallbackPlan(sports, data, basicInfo, avail) {
    const allDays2 = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
    const available = avail?.availableDays || ["Seg","Ter","Qua","Qui","Sex"];
    const preferred = avail?.preferredDays || {};

    // Build initial rest schedule
    const sched = allDays2.map(d => ({ day:d, type:"rest" }));

    const focuses = {
      strength: ["Peito & Ombros","Pernas","Costas & Bíceps","Posterior","Corpo Todo"],
      running: [{focus:"Corrida Fácil",distance:6,paceTarget:"6:00",hrZone:"Z2"},{focus:"Tempo Run",distance:8,paceTarget:"5:20",hrZone:"Z3-Z4"},{focus:"Longão",distance:14,paceTarget:"6:20",hrZone:"Z2"}],
      bike: [{focus:"Endurance",distance:40,hrZone:"Z2"},{focus:"Intervalado FTP",distance:30,hrZone:"Z4-Z5"}],
      swim: [{focus:"Técnica",distance:1500},{focus:"Resistência",distance:2000},{focus:"Intervalado",distance:1800}],
    };

    let si = { strength:0, running:0, bike:0, swim:0 };

    for (const sp of sports) {
      const daysForSport = data[sp]?.daysPerWeek || 2;
      // Use preferred days if set, otherwise use available days
      const pref = preferred[sp]?.length > 0 ? preferred[sp] : available;
      const candidates = pref.filter(d => available.includes(d));
      let assigned = 0;

      for (const d of candidates) {
        if (assigned >= daysForSport) break;
        const dayIdx = allDays2.indexOf(d);
        if (dayIdx === -1 || sched[dayIdx].type !== "rest") continue;
        const f = focuses[sp];
        if (sp === "strength") {
          sched[dayIdx] = { day:d, type:"strength", focus:f[si[sp]%f.length] };
        } else {
          const sess = f[si[sp]%f.length];
          sched[dayIdx] = { day:d, type:sp, ...sess };
        }
        si[sp]++; assigned++;
      }

      // If preferred days weren't enough, fill from remaining available days
      if (assigned < daysForSport) {
        for (const d of available) {
          if (assigned >= daysForSport) break;
          const dayIdx = allDays2.indexOf(d);
          if (dayIdx === -1 || sched[dayIdx].type !== "rest") continue;
          const f = focuses[sp];
          if (sp === "strength") {
            sched[dayIdx] = { day:d, type:"strength", focus:f[si[sp]%f.length] };
          } else {
            const sess = f[si[sp]%f.length];
            sched[dayIdx] = { day:d, type:sp, ...sess };
          }
          si[sp]++; assigned++;
        }
      }
    }

    return {
      planName:`Plano Personalizado — ${sports.map(s=>SPORTS.find(x=>x.id===s)?.label).join(" + ")}`,
      description:`Plano distribuído nos seus dias disponíveis (${available.join(", ")}). A IA ajustará progressivamente conforme você evolui.`,
      weeklySchedule: sched,
      aiCoachNote:`Bem-vindo, ${basicInfo.name||"atleta"}! Seu plano respeita sua rotina. Mantenha a consistência — é o que mais importa.`,
    };
  }

  const canNext = () => {
    if (step===1) return basic.name.trim().length>0 && basic.age && basic.weight && basic.height;
    if (step===2) return selectedSports.length>0;
    if (step === 3+sportSteps.length) return availability.availableDays.length>0;
    return true;
  };

  // ── STEP CONTENT ──
  function renderStep() {
    // Step 0: Welcome
    if (step===0) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", gap:18, alignItems:"center", textAlign:"center" }}>
        <div className="pop" style={{ width:80, height:80, borderRadius:24, background:`linear-gradient(135deg,${B.blue},${B.indigo})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:38, boxShadow:`0 12px 40px ${B.blue}55` }}>🏆</div>
        <div>
          <div style={{ fontSize:26, fontWeight:900, color:D.text, letterSpacing:"-.5px" }}>Bem-vindo ao Athlete AI</div>
          <div style={{ color:D.muted, fontSize:14, marginTop:8, lineHeight:1.7 }}>Vou criar um plano de treinamento <strong style={{color:D.text}}>100% personalizado</strong> para você — com base nos seus esportes, nível atual e objetivos.</div>
        </div>
        {[["🎯","Você escolhe os esportes que pratica"],["📊","Informa seu nível atual e objetivo"],["🤖","A IA cria o plano ideal para você"],["💬","Coach IA disponível para ajustes a qualquer momento"]].map(([e,t])=>(
          <div key={t} style={{ display:"flex", gap:14, alignItems:"center", textAlign:"left", background:D.card, borderRadius:14, padding:"13px 16px", width:"100%", borderLeft:`3px solid ${B.blue}` }}>
            <span style={{ fontSize:22, flexShrink:0 }}>{e}</span><span style={{ fontSize:14, color:D.sub, fontWeight:500 }}>{t}</span>
          </div>
        ))}
      </div>
    );

    // Step 1: Basic info
    if (step===1) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", gap:16 }}>
        <div><div style={{ fontSize:32, marginBottom:10 }}>👤</div><div style={{ fontSize:20, fontWeight:800, color:D.text }}>Seus dados básicos</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Usados para personalizar cargas e metas</div></div>
        <FInp lbl="Seu nome" val={basic.name} ch={v=>upBasic("name",v)} ph="Como devo te chamar?" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <FInp lbl="Idade" val={basic.age} ch={v=>upBasic("age",v)} ph="Anos" tp="number" />
          <div>
            <Lbl>Sexo biológico</Lbl>
            <div style={{ display:"flex", gap:8 }}>
              {[["M","Masculino"],["F","Feminino"]].map(([k,l])=>(
                <button key={k} onClick={()=>upBasic("sex",k)} style={{ flex:1, padding:"12px 0", borderRadius:12, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontWeight:700, fontSize:13, border:`1.5px solid ${basic.sex===k?B.blue:D.border}`, background:basic.sex===k?`${B.blue}20`:D.bg2, color:basic.sex===k?B.blue:D.muted }}>{l}</button>
              ))}
            </div>
          </div>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10 }}>
          <FInp lbl="Altura (cm)" val={basic.height} ch={v=>upBasic("height",v)} ph="175" tp="number" />
          <FInp lbl="Peso (kg)" val={basic.weight} ch={v=>upBasic("weight",v)} ph="80" tp="number" />
          <FInp lbl="Gordura %" val={basic.fat} ch={v=>upBasic("fat",v)} ph="20?" tp="number" />
        </div>
      </div>
    );

    // Step 2: Sport selection
    if (step===2) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", gap:16 }}>
        <div><div style={{ fontSize:32, marginBottom:10 }}>🏅</div><div style={{ fontSize:20, fontWeight:800, color:D.text }}>Quais esportes você pratica?</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Selecione um ou mais. Você pode ter combinações como musculação + corrida ou natação + bike.</div></div>
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {SPORTS.map(sp => {
            const sel = selectedSports.includes(sp.id);
            return (
              <button key={sp.id} onClick={()=>toggleSport(sp.id)} style={{ display:"flex", gap:14, alignItems:"center", padding:"16px", borderRadius:14, cursor:"pointer", textAlign:"left", border:`2px solid ${sel?sp.color:D.border}`, background:sel?`${sp.color}15`:D.card, fontFamily:"'Inter',sans-serif", transition:"all .15s" }}>
                <div style={{ width:48, height:48, borderRadius:14, background:sel?`${sp.color}30`:D.bg2, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26, flexShrink:0 }}>{sp.emoji}</div>
                <div style={{ flex:1 }}>
                  <div style={{ fontWeight:700, color:sel?sp.color:D.text, fontSize:16 }}>{sp.label}</div>
                  <div style={{ fontSize:12, color:D.muted, marginTop:2 }}>{sp.desc}</div>
                </div>
                <div style={{ width:24, height:24, borderRadius:"50%", border:`2px solid ${sel?sp.color:D.border}`, background:sel?sp.color:"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                  {sel&&<div style={{ fontSize:12, color:"#fff", fontWeight:800 }}>✓</div>}
                </div>
              </button>
            );
          })}
        </div>
        {selectedSports.length>0&&<div style={{ fontSize:12, color:B.emerald, fontWeight:600, textAlign:"center" }}>✓ {selectedSports.length} esporte{selectedSports.length>1?"s":""} selecionado{selectedSports.length>1?"s":""} — próximo: perguntas específicas para cada um</div>}
      </div>
    );

    // Steps 3..N: Sport-specific questions
    if (currentSportInStep) return <SportQuestions sport={currentSportInStep} data={sportData[currentSportInStep]} onUpdate={(k,v)=>upSport(currentSportInStep,k,v)}/>;

    // Availability step — days available + preferred days per sport + session duration
    if (step === 3+sportSteps.length) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", gap:22 }}>
        <div>
          <div style={{ fontSize:32, marginBottom:10 }}>📅</div>
          <div style={{ fontSize:20, fontWeight:800, color:D.text }}>Sua Rotina de Treinos</div>
          <div style={{ fontSize:13, color:D.muted, marginTop:4 }}>A IA vai montar o plano respeitando exatamente sua disponibilidade.</div>
        </div>

        {/* Available days */}
        <div>
          <Lbl>Quais dias você PODE treinar?</Lbl>
          <div style={{ display:"flex", gap:6 }}>
            {ALL_DAYS.map(d => {
              const avail = availability.availableDays.includes(d);
              return (
                <button key={d} onClick={()=>toggleAvailDay(d)} style={{
                  flex:1, padding:"11px 0", borderRadius:11, cursor:"pointer",
                  border:`1.5px solid ${avail?B.blue:D.border}`,
                  background:avail?`${B.blue}25`:D.card,
                  color:avail?B.blue:D.muted,
                  fontWeight:avail?700:400, fontSize:11,
                  fontFamily:"'Inter',sans-serif", transition:"all .15s",
                }}>
                  <div style={{ fontSize:9, marginBottom:3, color:avail?B.blue:D.muted }}>
                    {d==="Dom"||d==="Sáb"?"fim":""}
                  </div>
                  {d}
                </button>
              );
            })}
          </div>
          <div style={{ marginTop:8, fontSize:12, color:B.emerald, fontWeight:600 }}>
            {availability.availableDays.length === 0
              ? "⚠️ Selecione pelo menos 1 dia"
              : `✓ ${availability.availableDays.length} dias disponíveis por semana`}
          </div>
        </div>

        {/* Preferred days per sport */}
        {selectedSports.length > 0 && availability.availableDays.length > 0 && (
          <div>
            <Lbl>Em quais dias prefere cada esporte?</Lbl>
            <div style={{ fontSize:12, color:D.muted, marginBottom:14 }}>
              Opcional — deixe em branco e a IA distribui automaticamente.
            </div>
            {selectedSports.map(spId => {
              const sp = SPORTS.find(s=>s.id===spId);
              const prefDays = availability.preferredDays[spId]||[];
              return (
                <div key={spId} style={{ marginBottom:18 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10 }}>
                    <span style={{ fontSize:20 }}>{sp.emoji}</span>
                    <span style={{ fontSize:14, fontWeight:700, color:sp.color }}>{sp.label}</span>
                    {prefDays.length>0 && (
                      <span style={{ fontSize:11, color:sp.color, background:`${sp.color}20`, padding:"2px 8px", borderRadius:12, fontWeight:600 }}>
                        {prefDays.join(", ")}
                      </span>
                    )}
                  </div>
                  <div style={{ display:"flex", gap:6 }}>
                    {availability.availableDays.map(d => {
                      const sel = prefDays.includes(d);
                      return (
                        <button key={d} onClick={()=>togglePrefDay(spId,d)} style={{
                          padding:"9px 0", flex:1, borderRadius:9, cursor:"pointer",
                          border:`1.5px solid ${sel?sp.color:D.border}`,
                          background:sel?`${sp.color}25`:D.card,
                          color:sel?sp.color:D.muted,
                          fontWeight:sel?700:400, fontSize:11,
                          fontFamily:"'Inter',sans-serif", transition:"all .15s",
                        }}>{d}</button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Session duration */}
        <div>
          <Lbl>Duração média de cada sessão · <span style={{color:B.blue}}>{schedule.sessionDuration} min</span></Lbl>
          <input type="range" min={30} max={120} step={15} value={schedule.sessionDuration}
            onChange={e=>setSchedule(p=>({...p,sessionDuration:+e.target.value}))}
            style={{ width:"100%", accentColor:B.blue, cursor:"pointer" }}/>
          <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:D.muted, marginTop:4 }}>
            <span>30 min</span><span>2 horas</span>
          </div>
        </div>

        {/* Preview */}
        {availability.availableDays.length > 0 && (
          <Card style={{ background:`${B.indigo}12`, border:`1px solid ${B.indigo}25` }}>
            <Lbl>📋 Resumo da sua rotina</Lbl>
            <div style={{ display:"flex", gap:5 }}>
              {ALL_DAYS.map(d => {
                const avail = availability.availableDays.includes(d);
                const sportPrefs = selectedSports.filter(sp => (availability.preferredDays[sp]||[]).includes(d));
                const anyPref = sportPrefs.length > 0;
                const sp = anyPref ? SPORTS.find(s=>s.id===sportPrefs[0]) : null;
                return (
                  <div key={d} style={{ flex:1, textAlign:"center" }}>
                    <div style={{ fontSize:9, color:D.muted, marginBottom:4 }}>{d}</div>
                    <div style={{
                      width:"100%", aspectRatio:"1", borderRadius:8,
                      background: !avail ? D.bg2 : anyPref ? `${sp.color}30` : `${B.blue}15`,
                      border:`1px solid ${!avail?D.border:anyPref?sp.color:B.blue}40`,
                      display:"flex", alignItems:"center", justifyContent:"center", fontSize:12,
                    }}>
                      {!avail ? "🚫" : anyPref ? sp.emoji : "✓"}
                    </div>
                  </div>
                );
              })}
            </div>
            <div style={{ fontSize:11, color:D.muted, marginTop:10, lineHeight:1.6 }}>
              🚫 = dia bloqueado · ✓ = disponível · emoji = dia preferido para aquele esporte
            </div>
          </Card>
        )}
      </div>
    );

    // Health step
    if (step === 4+sportSteps.length) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", gap:16 }}>
        <div><div style={{ fontSize:32, marginBottom:10 }}>🏥</div><div style={{ fontSize:20, fontWeight:800, color:D.text }}>Saúde & Limitações</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Informações importantes para evitar lesões</div></div>
        <div>
          <Lbl>Lesões ou dores atuais</Lbl>
          <textarea value={health.injuries} onChange={e=>setHealth(p=>({...p,injuries:e.target.value}))} placeholder="Ex: dor no joelho, hérnia lombar, tendinite no ombro, cirurgia recente..." style={{ width:"100%", minHeight:80, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"12px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/>
        </div>
        <div>
          <Lbl>Outras limitações ou observações</Lbl>
          <textarea value={health.limitações} onChange={e=>setHealth(p=>({...p,limitações:e.target.value}))} placeholder="Preferências, restrições médicas, objetivos de prazo..." style={{ width:"100%", minHeight:70, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"12px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/>
        </div>
      </div>
    );

    // AI generating step
    if (step === totalSteps - 1) return (
      <div className="fade" style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:24, textAlign:"center", paddingTop:20 }}>
        <div style={{ position:"relative" }}>
          <svg width={96} height={96} style={{ transform:"rotate(-90deg)" }}>
            <circle cx={48} cy={48} r={40} fill="none" stroke={D.card} strokeWidth={8}/>
            <circle cx={48} cy={48} r={40} fill="none" stroke={B.blue} strokeWidth={8} strokeLinecap="round" strokeDasharray="251" style={{ animation:"spin 1.6s linear infinite" }}/>
          </svg>
          <div style={{ position:"absolute", inset:0, display:"flex", alignItems:"center", justifyContent:"center", fontSize:30 }}>🤖</div>
        </div>
        <div>
          <div style={{ fontSize:22, fontWeight:800, color:D.text, letterSpacing:"-.5px" }}>Criando seu plano, {basic.name||"atleta"}!</div>
          <div style={{ color:B.blue, fontSize:13, marginTop:8, fontWeight:600 }}>{genMsg}</div>
        </div>
        <Card style={{ width:"100%", textAlign:"left" }}>
          <Lbl>Seu plano vai incluir</Lbl>
          {selectedSports.map(s=>{
            const sp = SPORTS.find(x=>x.id===s);
            const d = sportData[s];
            const pref = availability.preferredDays[s];
            return <div key={s} style={{ display:"flex", gap:10, marginBottom:10 }}>
              <span style={{fontSize:18,flexShrink:0}}>{sp?.emoji}</span>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:sp?.color}}>{sp?.label} — {d.daysPerWeek||2}× por semana</div>
                <div style={{fontSize:12,color:D.muted}}>
                  {s==="strength"&&`Objetivo: ${d.goal==="hypertrophy"?"Hipertrofia":d.goal==="strength"?"Força":d.goal==="weight_loss"?"Emagrecimento":"Recomposição"}`}
                  {s==="running"&&`Pace atual: ${d.currentPace5k}/km → Meta: ${d.goalDistance}km`}
                  {s==="bike"&&`${d.weeklyKm||"?"} km/sem → Objetivo: ${d.goal}`}
                  {s==="swim"&&`Pace: ${d.currentPace100}/100m → ${d.goal}`}
                </div>
                {pref?.length>0 && <div style={{fontSize:11,color:sp?.color,marginTop:2}}>📅 Dias preferidos: {pref.join(", ")}</div>}
              </div>
            </div>;
          })}
        </Card>
      </div>
    );

    return null;
  }

  const isLastDataStep = step === 4+sportSteps.length;
  const isAIStep = step === totalSteps - 1;
  const progressPct = Math.min(step / (totalSteps-2), 1);

  return (
    <div style={{ minHeight:"100vh", background:D.bg, fontFamily:"'Inter',sans-serif", paddingBottom:90 }}>
      <div style={{ background:D.bar, padding:"16px 20px", borderBottom:`1px solid ${D.card}`, position:"sticky", top:0, zIndex:20 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
          <div style={{ fontSize:14, fontWeight:800, color:D.text }}>Athlete AI</div>
          <div style={{ fontSize:11, color:D.muted }}>Passo {Math.min(step+1,totalSteps-1)} de {totalSteps-1}</div>
        </div>
        <div style={{ background:D.card, borderRadius:6, height:5 }}>
          <div style={{ width:`${progressPct*100}%`, background:`linear-gradient(90deg,${B.blue},${B.indigo})`, height:"100%", borderRadius:6, transition:"width .4s ease" }}/>
        </div>
      </div>
      <div style={{ padding:"24px 20px", maxWidth:480, margin:"0 auto" }}>{renderStep()}</div>
      {!isAIStep&&(
        <div style={{ position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)", width:"100%", maxWidth:480, padding:"14px 20px 24px", background:D.bg, borderTop:`1px solid ${D.card}`, display:"flex", gap:10, boxSizing:"border-box" }}>
          {step>0&&<button onClick={()=>setStep(s=>s-1)} style={{ flex:"0 0 80px", padding:"13px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:14 }}>← Voltar</button>}
          <button disabled={!canNext()} onClick={nextStep} style={{ flex:1, padding:"14px 0", borderRadius:12, border:"none", background:canNext()?B.blue:"#334155", color:canNext()?"#fff":"#475569", fontWeight:800, fontSize:15, cursor:canNext()?"pointer":"not-allowed", fontFamily:"'Inter',sans-serif", boxShadow:canNext()?`0 4px 20px ${B.blue}44`:"none" }}>
            {isLastDataStep?"🚀 Gerar Meu Plano":"Continuar →"}
          </button>
        </div>
      )}
    </div>
  );
}

/* ── SPORT-SPECIFIC QUESTIONS ────────────────────────────────────── */
function SportQuestions({ sport, data, onUpdate }) {
  const sp = SPORTS.find(s=>s.id===sport);
  if (sport==="strength") return (
    <div className="fade" style={{ display:"flex", flexDirection:"column", gap:18 }}>
      <div><div style={{ fontSize:32 }}>{sp.emoji}</div><div style={{ fontSize:20, fontWeight:800, color:sp.color, marginTop:8 }}>Musculação</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Vou adaptar as cargas ao seu nível</div></div>
      <div>
        <Lbl>Objetivo principal</Lbl>
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {[{v:"hypertrophy",l:"💪 Hipertrofia",d:"Ganhar massa muscular"},{v:"strength",l:"🏋️ Força",d:"Aumentar cargas máximas"},{v:"weight_loss",l:"🔥 Emagrecimento",d:"Perder gordura mantendo músculo"},{v:"body_recomp",l:"⚖️ Recomposição",d:"Perder gordura e ganhar músculo"}].map(o=>(
            <button key={o.v} onClick={()=>onUpdate("goal",o.v)} style={{ display:"flex", gap:12, alignItems:"center", padding:"12px 14px", borderRadius:12, cursor:"pointer", textAlign:"left", border:`1.5px solid ${data.goal===o.v?sp.color:D.border}`, background:data.goal===o.v?`${sp.color}15`:D.card, fontFamily:"'Inter',sans-serif" }}>
              <div style={{ flex:1 }}><div style={{ fontWeight:700, color:data.goal===o.v?sp.color:D.text, fontSize:14 }}>{o.l}</div><div style={{ fontSize:12, color:D.muted }}>{o.d}</div></div>
              {data.goal===o.v&&<span style={{ color:sp.color, fontWeight:800 }}>✓</span>}
            </button>
          ))}
        </div>
      </div>
      <div>
        <Lbl>Experiência em musculação</Lbl>
        <SelBtn opts={[{v:"beginner",l:"Iniciante (<1 ano)"},{v:"intermediate",l:"Intermediário (1-3 anos)"},{v:"advanced",l:"Avançado (3+ anos)"}]} val={data.experience} onSel={v=>onUpdate("experience",v)} color={sp.color}/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <FInp lbl="Supino aprox. (kg)" val={data.currentBench} ch={v=>onUpdate("currentBench",v)} ph="Ex: 60" tp="number"/>
        <FInp lbl="Agachamento aprox. (kg)" val={data.currentSquat} ch={v=>onUpdate("currentSquat",v)} ph="Ex: 80" tp="number"/>
      </div>
      <div>
        <Lbl>Dias por semana</Lbl>
        <SelBtn opts={[{v:2,l:"2×"},{v:3,l:"3×"},{v:4,l:"4×"},{v:5,l:"5×"},{v:6,l:"6×"}]} val={data.daysPerWeek} onSel={v=>onUpdate("daysPerWeek",v)} color={sp.color}/>
      </div>
      <div style={{ display:"flex", gap:10 }}>
        <button onClick={()=>onUpdate("hasGym",true)} style={{ flex:1, padding:"12px", borderRadius:12, cursor:"pointer", border:`1.5px solid ${data.hasGym?sp.color:D.border}`, background:data.hasGym?`${sp.color}20`:D.card, color:data.hasGym?sp.color:D.muted, fontWeight:700, fontSize:13, fontFamily:"'Inter',sans-serif" }}>🏋️ Academia</button>
        <button onClick={()=>onUpdate("hasGym",false)} style={{ flex:1, padding:"12px", borderRadius:12, cursor:"pointer", border:`1.5px solid ${!data.hasGym?sp.color:D.border}`, background:!data.hasGym?`${sp.color}20`:D.card, color:!data.hasGym?sp.color:D.muted, fontWeight:700, fontSize:13, fontFamily:"'Inter',sans-serif" }}>🏠 Casa/Outro</button>
      </div>
    </div>
  );

  if (sport==="running") return (
    <div className="fade" style={{ display:"flex", flexDirection:"column", gap:18 }}>
      <div><div style={{ fontSize:32 }}>{sp.emoji}</div><div style={{ fontSize:20, fontWeight:800, color:sp.color, marginTop:8 }}>Corrida</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Seu histórico atual para gerar o plano ideal</div></div>
      <div>
        <Lbl>Pace atual nos 5km (min:seg/km)</Lbl>
        <SelBtn opts={[{v:"7:00",l:">7:00"},{v:"6:30",l:"6:30"},{v:"6:00",l:"6:00"},{v:"5:30",l:"5:30"},{v:"5:00",l:"5:00"},{v:"4:30",l:"<4:30"}]} val={data.currentPace5k} onSel={v=>onUpdate("currentPace5k",v)} color={sp.color}/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <FInp lbl="Maior distância já corrida (km)" val={data.longestRun} ch={v=>onUpdate("longestRun",v)} ph="Ex: 10" tp="number"/>
        <FInp lbl="Km por semana atualmente" val={data.weeklyKm} ch={v=>onUpdate("weeklyKm",v)} ph="Ex: 20" tp="number"/>
      </div>
      <div>
        <Lbl>Objetivo de distância</Lbl>
        <SelBtn opts={[{v:"5",l:"5km"},{v:"10",l:"10km"},{v:"21",l:"21km (meia)"},{v:"42",l:"42km (maratona)"},{v:"condicionamento",l:"Condicionamento"}]} val={data.goalDistance} onSel={v=>onUpdate("goalDistance",v)} color={sp.color}/>
      </div>
      <div>
        <Lbl>Dias de corrida por semana</Lbl>
        <SelBtn opts={[{v:2,l:"2×"},{v:3,l:"3×"},{v:4,l:"4×"},{v:5,l:"5×"}]} val={data.daysPerWeek} onSel={v=>onUpdate("daysPerWeek",v)} color={sp.color}/>
      </div>
      {data.goalDistance!=="condicionamento"&&(
        <FInp lbl="Prazo para a prova (opcional)" val={data.goalTime} ch={v=>onUpdate("goalTime",v)} ph="Ex: Dezembro 2025"/>
      )}
    </div>
  );

  if (sport==="bike") return (
    <div className="fade" style={{ display:"flex", flexDirection:"column", gap:18 }}>
      <div><div style={{ fontSize:32 }}>{sp.emoji}</div><div style={{ fontSize:20, fontWeight:800, color:sp.color, marginTop:8 }}>Ciclismo</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Sua situação atual no pedal</div></div>
      <div>
        <Lbl>Tipo de bicicleta</Lbl>
        <SelBtn opts={[{v:"speed",l:"Speed/Road"},{v:"mtb",l:"MTB"},{v:"gravel",l:"Gravel"},{v:"spinning",l:"Spinning"},{v:"qualquer",l:"Qualquer"}]} val={data.bikeType} onSel={v=>onUpdate("bikeType",v)} color={sp.color}/>
      </div>
      <FInp lbl="FTP atual (Watts) — deixe vazio se não sabe" val={data.currentFTP} ch={v=>onUpdate("currentFTP",v)} ph="Ex: 220" tp="number"/>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <FInp lbl="Km por semana atualmente" val={data.weeklyKm} ch={v=>onUpdate("weeklyKm",v)} ph="Ex: 80" tp="number"/>
        <div>
          <Lbl>Dias por semana</Lbl>
          <SelBtn opts={[{v:2,l:"2×"},{v:3,l:"3×"},{v:4,l:"4×"}]} val={data.daysPerWeek} onSel={v=>onUpdate("daysPerWeek",v)} color={sp.color}/>
        </div>
      </div>
      <div>
        <Lbl>Objetivo</Lbl>
        <SelBtn opts={[{v:"condicionamento",l:"Condicionamento"},{v:"performance",l:"Performance/FTP"},{v:"evento",l:"Evento/Prova"},{v:"triathlon",l:"Triathlon"}]} val={data.goal} onSel={v=>onUpdate("goal",v)} color={sp.color}/>
      </div>
    </div>
  );

  if (sport==="swim") return (
    <div className="fade" style={{ display:"flex", flexDirection:"column", gap:18 }}>
      <div><div style={{ fontSize:32 }}>{sp.emoji}</div><div style={{ fontSize:20, fontWeight:800, color:sp.color, marginTop:8 }}>Natação</div><div style={{ fontSize:13, color:D.muted, marginTop:4 }}>Vou adaptar o treino ao seu nível na água</div></div>
      <div>
        <Lbl>Você já nada?</Lbl>
        <div style={{ display:"flex", gap:10 }}>
          <button onClick={()=>onUpdate("canSwim",true)} style={{ flex:1, padding:"12px", borderRadius:12, cursor:"pointer", border:`1.5px solid ${data.canSwim?sp.color:D.border}`, background:data.canSwim?`${sp.color}20`:D.card, color:data.canSwim?sp.color:D.muted, fontWeight:700, fontSize:13, fontFamily:"'Inter',sans-serif" }}>Sim, já sei nadar</button>
          <button onClick={()=>onUpdate("canSwim",false)} style={{ flex:1, padding:"12px", borderRadius:12, cursor:"pointer", border:`1.5px solid ${!data.canSwim?sp.color:D.border}`, background:!data.canSwim?`${sp.color}20`:D.card, color:!data.canSwim?sp.color:D.muted, fontWeight:700, fontSize:13, fontFamily:"'Inter',sans-serif" }}>Estou aprendendo</button>
        </div>
      </div>
      {data.canSwim&&(
        <>
          <div>
            <Lbl>Estilo principal</Lbl>
            <SelBtn opts={[{v:"crawl",l:"Crawl"},{v:"costas",l:"Costas"},{v:"peito",l:"Peito"},{v:"borboleta",l:"Borboleta"},{v:"variado",l:"Variado"}]} val={data.style} onSel={v=>onUpdate("style",v)} color={sp.color}/>
          </div>
          <div>
            <Lbl>Pace atual por 100m (min:seg)</Lbl>
            <SelBtn opts={[{v:"2:40",l:">2:40"},{v:"2:20",l:"2:20"},{v:"2:00",l:"2:00"},{v:"1:45",l:"1:45"},{v:"1:30",l:"<1:30"}]} val={data.currentPace100} onSel={v=>onUpdate("currentPace100",v)} color={sp.color}/>
          </div>
          <FInp lbl="Volume semanal atual (metros)" val={data.weeklyVolume} ch={v=>onUpdate("weeklyVolume",v)} ph="Ex: 2000" tp="number"/>
        </>
      )}
      <div>
        <Lbl>Objetivo</Lbl>
        <SelBtn opts={[{v:"tecnica",l:"Melhorar técnica"},{v:"resistencia",l:"Resistência"},{v:"velocidade",l:"Velocidade"},{v:"triathlon",l:"Triathlon"},{v:"condicionamento",l:"Condicionamento geral"}]} val={data.goal} onSel={v=>onUpdate("goal",v)} color={sp.color}/>
      </div>
      <div>
        <Lbl>Dias por semana</Lbl>
        <SelBtn opts={[{v:2,l:"2×"},{v:3,l:"3×"},{v:4,l:"4×"}]} val={data.daysPerWeek} onSel={v=>onUpdate("daysPerWeek",v)} color={sp.color}/>
      </div>
    </div>
  );
  return null;
}

/* ══ STATS UTILITIES ════════════════════════════════════════════ */
function parseLogDate(str) {
  if (!str) return new Date(0);
  const parts = str.split("/");
  if (parts.length === 3) return new Date(+parts[2], +parts[1]-1, +parts[0]);
  return new Date(str);
}

function calcAccumulated(history, type) {
  const now = new Date();
  const startWeek = new Date(now); startWeek.setDate(now.getDate() - now.getDay()); startWeek.setHours(0,0,0,0);
  const startMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const todayStr = now.toLocaleDateString("pt-BR");

  const rel = history.filter(w => w.type === type);
  const week = rel.filter(w => parseLogDate(w.date) >= startWeek);
  const month = rel.filter(w => parseLogDate(w.date) >= startMonth);
  const today = rel.filter(w => w.date === todayStr);

  if (type === "running") {
    const km = arr => arr.reduce((s,w) => s + (parseFloat(w.log?.distance)||0), 0);
    return { today: km(today), week: km(week), month: km(month), total: km(rel), sessions: rel.length, weekSessions: week.length };
  }
  if (type === "bike") {
    const km = arr => arr.reduce((s,w) => s + (parseFloat(w.log?.distance)||0), 0);
    return { today: km(today), week: km(week), month: km(month), total: km(rel), sessions: rel.length };
  }
  if (type === "swim") {
    const m = arr => arr.reduce((s,w) => s + (parseFloat(w.log?.volume)||0), 0);
    return { today: m(today), week: m(week), month: m(month), total: m(rel), sessions: rel.length };
  }
  if (type === "strength") {
    return { week: week.length, month: month.length, total: rel.length, weekSessions: week.length };
  }
  return {};
}

function parsePaceSec(pace) {
  if (!pace) return 999;
  const [m,s] = String(pace).split(":").map(Number);
  return (m||0)*60 + (s||0);
}

function secToPace(sec) {
  if (!sec || sec >= 999) return "—";
  const m = Math.floor(sec/60); const s = sec%60;
  return `${m}:${String(s).padStart(2,"0")}`;
}

// Check if recent sessions indicate accumulated fatigue
function checkFatigue(history) {
  const recent = history.slice(-6).filter(w => w.log?.logs || w.log?.feel);
  if (recent.length < 3) return null;
  const allRPEs = recent.flatMap(w => Object.values(w.log?.logs||{}).flat().map(s=>s.rpe).filter(Boolean));
  const feels = recent.map(w=>w.log?.feel).filter(Boolean);
  if (allRPEs.length >= 3) {
    const avg = allRPEs.reduce((a,b)=>a+b,0)/allRPEs.length;
    if (avg > 8.7) return { type:"rpe", avg: avg.toFixed(1) };
  }
  if (feels.length >= 3) {
    const avgFeel = feels.reduce((a,b)=>a+b,0)/feels.length;
    if (avgFeel < 5) return { type:"feel", avg: avgFeel.toFixed(1) };
  }
  return null;
}

// Build pace trend data for running
function buildRunPaceData(history) {
  return history
    .filter(w => w.type==="running" && w.log?.pace)
    .map(w => ({
      d: w.date?.split("/").slice(0,2).join("/") || w.date,
      v: parsePaceSec(w.log.pace),
      dist: parseFloat(w.log.distance)||0,
    }))
    .filter(d => d.v > 0 && d.v < 900);
}

// Build strength load data per exercise
function buildLoadData(history, exId) {
  return history
    .filter(w => w.type==="strength" && w.log?.logs?.[exId]?.length>0)
    .map(w => ({
      d: w.date?.split("/").slice(0,2).join("/") || w.date,
      v: Math.max(...(w.log.logs[exId]||[]).map(s=>s.load||0)),
      reps: Math.max(...(w.log.logs[exId]||[]).map(s=>s.reps||0)),
    }))
    .filter(d => d.v > 0);
}

// Get all exercise IDs that appear in history
function getTrackedExercises(history) {
  const counts = {};
  history.filter(w=>w.type==="strength").forEach(w => {
    Object.keys(w.log?.logs||{}).forEach(id => {
      counts[id] = (counts[id]||0) + 1;
    });
  });
  return Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([id])=>id);
}

const IS_PROD = typeof window !== "undefined" && !window.location.hostname.includes("localhost") && !window.location.hostname.includes("claude");
const AI_URL = IS_PROD ? "/.netlify/functions/ai" : "https://api.anthropic.com/v1/messages";

async function callAI(body) {
  const res = await fetch(AI_URL, {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body: JSON.stringify(IS_PROD ? body : body),
  });
  return res.json();
}

function localFeedback(type, planned, log) {
  const feel = parseInt(log.feel) || 7;

  if (type === "running") {
    const pDist = parseFloat(planned.distance) || 0;
    const aDist = parseFloat(log.distance) || 0;
    const pPaceSec = parsePaceSec(planned.paceTarget);
    const aPaceSec = parsePaceSec(log.pace);
    const paceDiff = aPaceSec - pPaceSec; // positivo = mais lento
    const distPct = pDist > 0 ? aDist / pDist : 1;
    const parts = [];

    // Volume
    if (distPct < 0.75) {
      parts.push(`Você completou apenas ${Math.round(distPct*100)}% do volume planejado — faltaram ${(pDist-aDist).toFixed(1)}km. Isso não é ajuste fino, é um treino diferente do que foi prescrito. Se não houve lesão ou sintoma físico real, o problema é de comprometimento com o processo.`);
    } else if (distPct < 0.92) {
      parts.push(`${(pDist-aDist).toFixed(1)}km abaixo do planejado. Se foi por sinal real do corpo, ok. Se foi falta de disposição, cuidado — esse padrão de encurtar treinos acumula déficit de volume sem você perceber.`);
    } else if (distPct > 1.12) {
      parts.push(`Rodou ${(aDist-pDist).toFixed(1)}km a mais que o prescrito. Parece dedicação, mas não é — sair do plano por excesso é tão problemático quanto por falta. A periodização tem motivo. Volume extra hoje = recuperação comprometida amanhã.`);
    }

    // Pace
    if (pPaceSec < 900 && aPaceSec < 900) {
      if (paceDiff > 45) {
        parts.push(`Pace ${paceDiff}s/km acima do alvo. Isso não é uma variação — é um treino em zona completamente diferente do prescrito. Avalie honestamente: foi o terreno, o calor, ou você simplesmente não foi com intensidade suficiente?`);
      } else if (paceDiff > 20) {
        parts.push(`Pace ${paceDiff}s/km acima do alvo. Margem relevante. Em treinos de qualidade (tempo, intervalo), isso significa que o estímulo fisiológico foi menor do que o planejado.`);
      } else if (paceDiff < -25) {
        parts.push(`Pace ${Math.abs(paceDiff)}s/km mais rápido que o planejado. Em treinos de base (Z1/Z2), isso é erro — você saiu da zona aeróbia e transformou um treino de recuperação em um treino de intensidade não programada. O custo aparece nos próximos dias.`);
      }
    }

    // Sensação vs desempenho
    if (feel <= 4) {
      parts.push(`Sensação ${feel}/10. Se isso é frequente, você está com recuperação insuficiente — reveja sono (mínimo 7h), alimentação pré-treino e volume semanal total.`);
    } else if (feel >= 9 && distPct < 0.9) {
      parts.push(`Sensação ${feel}/10 mas volume abaixo do plano. Disposição boa que não foi convertida em treino completo é oportunidade desperdiçada.`);
    }

    if (parts.length === 0) {
      if (Math.abs(paceDiff) <= 10 && distPct >= 0.97) {
        parts.push(`Execução dentro dos parâmetros. Distância e pace dentro da margem aceitável. Isso é o que se espera — não menos. A consistência de executar o plano como prescrito é o que separa atletas que evoluem dos que ficam no lugar.`);
      } else {
        parts.push(`Treino registrado. Avalie se os dados batem com o esforço percebido — números honestos são mais úteis do que números bonitos.`);
      }
    }

    return parts.join(" ");
  }

  if (type === "strength") {
    const logs = log.logs || {};
    const totalSets = Object.values(logs).reduce((s,a) => s+a.length, 0);
    const allRPEs = Object.values(logs).flat().map(s => s.rpe).filter(Boolean);
    const avgRPE = allRPEs.length > 0 ? (allRPEs.reduce((a,b)=>a+b,0)/allRPEs.length).toFixed(1) : null;
    const parts = [];

    if (totalSets === 0) {
      return "Treino registrado sem dados de séries. Sem números, não há análise possível e nenhum ajuste pode ser feito no seu plano. Registre cargas e reps em pelo menos nos exercícios principais.";
    }

    if (avgRPE !== null) {
      const rpe = parseFloat(avgRPE);
      if (rpe < 6.5) {
        parts.push(`RPE médio de ${avgRPE}/10. Isso é treinamento turístico — você está se movimentando, não gerando estímulo de hipertrofia. Para ganho de massa, RPE mínimo de 7-8 nos exercícios compostos é obrigatório. Se as cargas estão fáceis, elas estão erradas.`);
      } else if (rpe < 7.5) {
        parts.push(`RPE médio ${avgRPE}/10. Aceitável, mas no limite inferior. Aumente a carga nos próximos treinos — se você sente que poderia fazer 4+ reps a mais, a carga está subestimada.`);
      } else if (rpe > 9.3) {
        parts.push(`RPE médio ${avgRPE}/10. Intensidade muito alta de forma consistente gera fadiga acumulada e aumenta risco de lesão. RPE 8-9 é o sweet spot — não confunda sofrimento com estímulo.`);
      } else {
        parts.push(`RPE médio ${avgRPE}/10. Zona de intensidade adequada para hipertrofia.`);
      }
    }

    if (feel <= 5) {
      parts.push(`Sensação ${feel}/10 após musculação. Fadiga excessiva pode indicar volume alto demais para a sua recuperação atual, ou déficit de sono/alimentação.`);
    }

    parts.push(`Pergunta crítica para o próximo treino de ${planned?.focus||"musculação"}: houve progressão de carga em pelo menos 1 exercício? Sem sobrecarga progressiva, o corpo não tem motivo para adaptar.`);

    return parts.join(" ");
  }

  if (type === "bike") {
    const pDist = parseFloat(planned.distance) || 0;
    const aDist = parseFloat(log.distance) || 0;
    const power = parseFloat(log.avgPower) || 0;
    const cadence = parseFloat(log.cadence) || 0;
    const distPct = pDist > 0 ? aDist/pDist : 1;
    const parts = [];

    if (distPct < 0.8) {
      parts.push(`${Math.round((1-distPct)*100)}% do volume cortado. Se não foi mecânico ou condição climática extrema, o treino foi incompleto.`);
    } else if (distPct > 1.15) {
      parts.push(`Volume acima do prescrito. Ciclismo tem curva de fadiga traiçoeira — você não sente no momento, mas paga nas próximas 48h.`);
    }

    if (cadence > 0 && cadence < 80) {
      parts.push(`Cadência de ${cadence}rpm está baixa. Pedalar pesado e lento força isometricamente o quadríceps e prejudica eficiência. Trabalhe para manter 85-95rpm mesmo que precise baixar a carga.`);
    }

    if (power > 0) {
      parts.push(`Potência média de ${power}W registrada. Compare com suas sessões anteriores — potência é o único dado objetivo que mostra se você está evoluindo ou apenas pedalando.`);
    }

    if (parts.length === 0) {
      parts.push(`Pedal registrado. Potência e cadência são os dados que realmente importam no ciclismo — se não está medindo, está treinando no escuro.`);
    }

    return parts.join(" ");
  }

  if (type === "swim") {
    const pVol = parseFloat(planned.distance) || 0;
    const aVol = parseFloat(log.volume) || 0;
    const strokes = parseInt(log.strokes) || 0;
    const pace = log.pace100 || "";
    const volPct = pVol > 0 ? aVol/pVol : 1;
    const parts = [];

    if (volPct < 0.8) {
      parts.push(`Volume ${Math.round(volPct*100)}% do planejado. Na natação, consistência de volume é ainda mais crítica que em outros esportes porque a adaptação técnica depende de repetição acumulada.`);
    }

    if (strokes > 0) {
      if (strokes > 22) {
        parts.push(`${strokes} braçadas por comprimento indica stroke curto — você está "remando" mais do que deslizando. Foque no alcance frontal e na rotação de ombro antes de aumentar velocidade.`);
      } else if (strokes < 14) {
        parts.push(`${strokes} braçadas por comprimento é baixo — verifique se não está cobrando de forma passiva (deslize longo mas força insuficiente).`);
      } else {
        parts.push(`${strokes} braçadas/comp. — contagem dentro da faixa adequada para crawl eficiente.`);
      }
    }

    if (pace) {
      parts.push(`Pace ${pace}/100m registrado. Acompanhe essa métrica a cada treino — é o único indicador objetivo de evolução técnica e aeróbia na natação.`);
    }

    if (parts.length === 0) {
      parts.push(`Natação registrada. Braçadas por comprimento e pace/100m são os dados que separam quem evolui de quem só se cansa na água. Registre sempre.`);
    }

    return parts.join(" ");
  }

  return "Atividade registrada. Dados insuficientes para análise. Preencha as métricas do treino para receber feedback útil.";
}

function CompareRow({ icon, label, planned, executed, good }) {
  const statusColor = good === true ? B.emerald : good === false ? B.orange : D.muted;
  const statusIcon = good === true ? "✅" : good === false ? "⚠️" : "—";
  return (
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:0, padding:"10px 0", borderBottom:`1px solid ${D.bg2}` }}>
      <div style={{ display:"flex", gap:8, alignItems:"center" }}>
        <span style={{ fontSize:14 }}>{icon}</span>
        <div>
          <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>{label}</div>
          <div style={{ fontSize:13, fontWeight:600, color:D.sub }}>{planned||"—"}</div>
        </div>
      </div>
      <div style={{ display:"flex", alignItems:"center", gap:8, justifyContent:"flex-end" }}>
        <div style={{ textAlign:"right" }}>
          <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>Executado</div>
          <div style={{ fontSize:14, fontWeight:800, color:statusColor }}>{executed||"—"}</div>
        </div>
        <span style={{ fontSize:16 }}>{statusIcon}</span>
      </div>
    </div>
  );
}

function StatBar({ label, value, max, color, unit="km" }) {
  const pct = max > 0 ? Math.min(value/max*100, 100) : 0;
  return (
    <div style={{ marginBottom:14 }}>
      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
        <span style={{ fontSize:12, color:D.sub, fontWeight:500 }}>{label}</span>
        <span style={{ fontSize:12, fontWeight:700, color }}>
          {value.toFixed(1)}{unit} <span style={{ color:D.muted, fontWeight:400 }}>/ {max}{unit}</span>
        </span>
      </div>
      <div style={{ background:D.bg2, borderRadius:6, height:8, overflow:"hidden" }}>
        <div style={{
          width:`${pct}%`, background:color, borderRadius:6, height:"100%",
          transition:"width 1.2s ease", boxShadow:`0 0 8px ${color}66`,
        }}/>
      </div>
      <div style={{ fontSize:10, color:D.muted, marginTop:4 }}>{Math.round(pct)}% da meta</div>
    </div>
  );
}

function WorkoutSummaryCard({ entry, history, profile, onDelete }) {
  const [confirmDel, setConfirmDel] = useState(false);
  const { type, plan, log, date } = entry;
  const sp = SPORTS.find(s=>s.id===type);
  const sc = SPORT_COLOR[type]||B.blue;
  const acc = calcAccumulated(history, type);
  const feedback = localFeedback(type, plan||{}, log||{});

  // Sport-specific goal from profile
  const runGoalKmWeek = parseFloat(profile?.sportData?.running?.weeklyKm)||25;
  const runGoalKmMonth = runGoalKmWeek * 4;

  return (
    <div className="pop" style={{ margin:"14px 16px 0" }}>
      <Card accent={sc}>
        {/* Header */}
        <div style={{ display:"flex", gap:12, alignItems:"center", marginBottom:16 }}>
          <div style={{
            width:44, height:44, borderRadius:12,
            background:`linear-gradient(135deg,${sc},${sc}99)`,
            display:"flex", alignItems:"center", justifyContent:"center", fontSize:22,
            boxShadow:`0 4px 16px ${sc}55`, flexShrink:0,
          }}>🏆</div>
          <div>
            <div style={{ fontSize:12, color:sc, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em" }}>
              ✓ Treino Concluído
            </div>
            <div style={{ fontSize:18, fontWeight:900, color:D.text, letterSpacing:"-.3px" }}>
              {plan?.focus || sp?.label || "Atividade"}
            </div>
            <div style={{ fontSize:11, color:D.muted }}>{date}</div>
          </div>
        </div>

        {/* ── PLANNED vs EXECUTED ── */}
        <div style={{ background:D.bg2, borderRadius:12, padding:"12px 14px", marginBottom:16 }}>
          <div style={{ fontSize:10, color:D.muted, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:4 }}>
            📋 Planejado vs Executado
          </div>

          {type==="running" && <>
            <CompareRow icon="📍" label="Distância" planned={`${plan?.distance||"?"}km`} executed={`${log?.distance||"?"}km`}
              good={parseFloat(log?.distance)>=(parseFloat(plan?.distance)||0)*0.95}/>
            <CompareRow icon="⚡" label="Pace" planned={plan?.paceTarget?`${plan.paceTarget}/km`:"—"} executed={log?.pace?`${log.pace}/km`:"—"}
              good={log?.pace&&plan?.paceTarget ? parsePaceSec(log.pace)-parsePaceSec(plan.paceTarget)<=20 : null}/>
            <CompareRow icon="❤️" label="FC" planned={plan?.hrZone||"—"} executed={log?.hr?`${log.hr}bpm`:"—"} good={null}/>
            <CompareRow icon="😊" label="Sensação" planned="—" executed={log?.feel?`${log.feel}/10`:"—"}
              good={parseInt(log?.feel)>=7}/>
          </>}

          {type==="strength" && <>
            <CompareRow icon="💪" label="Sessão" planned={plan?.focus||"—"} executed="Concluída" good={true}/>
            <CompareRow icon="📊" label="Volume" planned="Treino completo" executed={log?.logs?`${Object.values(log.logs).reduce((s,a)=>s+a.length,0)} séries`:"Registrado"} good={true}/>
            <CompareRow icon="😊" label="Sensação" planned="—" executed={log?.feel?`${log.feel}/10`:"—"} good={parseInt(log?.feel)>=7}/>
          </>}

          {type==="bike" && <>
            <CompareRow icon="📍" label="Distância" planned={`${plan?.distance||"?"}km`} executed={`${log?.distance||"?"}km`}
              good={parseFloat(log?.distance)>=(parseFloat(plan?.distance)||0)*0.9}/>
            <CompareRow icon="⚡" label="Potência" planned="—" executed={log?.avgPower?`${log.avgPower}W`:"—"} good={null}/>
            <CompareRow icon="🔄" label="Cadência" planned="—" executed={log?.cadence?`${log.cadence}rpm`:"—"} good={null}/>
            <CompareRow icon="😊" label="Sensação" planned="—" executed={log?.feel?`${log.feel}/10`:"—"} good={parseInt(log?.feel)>=7}/>
          </>}

          {type==="swim" && <>
            <CompareRow icon="📍" label="Volume" planned={`${plan?.distance||"?"}m`} executed={`${log?.volume||"?"}m`}
              good={parseFloat(log?.volume)>=(parseFloat(plan?.distance)||0)*0.9}/>
            <CompareRow icon="⚡" label="Pace/100m" planned="—" executed={log?.pace100||"—"} good={null}/>
            <CompareRow icon="🔄" label="Braçadas" planned="—" executed={log?.strokes?`${log.strokes}/comp`:"—"} good={null}/>
            <CompareRow icon="😊" label="Sensação" planned="—" executed={log?.feel?`${log.feel}/10`:"—"} good={parseInt(log?.feel)>=7}/>
          </>}
        </div>

        {/* ── ACCUMULATED STATS ── */}
        {(type==="running"||type==="bike"||type==="swim") && (
          <div style={{ background:D.bg2, borderRadius:12, padding:"14px 16px", marginBottom:16 }}>
            <div style={{ fontSize:10, color:D.muted, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:14 }}>
              📊 {type==="running"?"Corridas acumuladas":type==="bike"?"Pedais acumulados":"Natação acumulada"}
            </div>

            {/* Quick numbers row */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8, marginBottom:16 }}>
              {[
                {l:"Hoje",v:`${(type==="swim"?acc.today:acc.today).toFixed(1)}${type==="swim"?"m":"km"}`,c:sc},
                {l:"Esta semana",v:`${(acc.week||0).toFixed(1)}${type==="swim"?"m":"km"}`,c:sc},
                {l:"Este mês",v:`${(acc.month||0).toFixed(1)}${type==="swim"?"m":"km"}`,c:sc},
              ].map(m=>(
                <div key={m.l} style={{ textAlign:"center", background:D.card, borderRadius:10, padding:"10px 6px" }}>
                  <div style={{ fontSize:15, fontWeight:800, color:m.c }}>{m.v}</div>
                  <div style={{ fontSize:10, color:D.muted, marginTop:3 }}>{m.l}</div>
                </div>
              ))}
            </div>

            {/* Progress bars (running only - has goals) */}
            {type==="running" && <>
              <StatBar label="Meta semanal de km" value={acc.week||0} max={runGoalKmWeek} color={sc}/>
              <StatBar label="Meta mensal de km" value={acc.month||0} max={runGoalKmMonth} color={sc}/>
            </>}

            {type==="bike" && <>
              <StatBar label="Km esta semana" value={acc.week||0} max={parseFloat(profile?.sportData?.bike?.weeklyKm)||80} color={sc}/>
            </>}

            {type==="swim" && <>
              <StatBar label="Volume esta semana (m)" value={acc.week||0} max={parseFloat(profile?.sportData?.swim?.weeklyVolume)||4000} color={sc} unit="m"/>
            </>}

            {/* Total line */}
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", paddingTop:10, borderTop:`1px solid ${D.card}` }}>
              <span style={{ fontSize:12, color:D.muted }}>{type==="swim"?"Volume total":"Total acumulado"}</span>
              <span style={{ fontSize:15, fontWeight:800, color:sc }}>
                {(acc.total||0).toFixed(1)}{type==="swim"?"m":"km"} · {acc.sessions||0} sessões
              </span>
            </div>
          </div>
        )}

        {/* Strength stats */}
        {type==="strength" && (
          <div style={{ background:D.bg2, borderRadius:12, padding:"14px 16px", marginBottom:16 }}>
            <div style={{ fontSize:10, color:D.muted, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:12 }}>
              📊 Musculação acumulada
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8 }}>
              {[
                {l:"Esta semana",v:`${acc.weekSessions||0}`,u:"treinos",c:sc},
                {l:"Este mês",v:`${acc.month||0}`,u:"treinos",c:sc},
                {l:"Total",v:`${acc.total||0}`,u:"sessões",c:sc},
              ].map(m=>(
                <div key={m.l} style={{ textAlign:"center", background:D.card, borderRadius:10, padding:"10px 6px" }}>
                  <div style={{ fontSize:18, fontWeight:800, color:m.c }}>{m.v}</div>
                  <div style={{ fontSize:10, color:D.muted, marginTop:2 }}>{m.u}</div>
                  <div style={{ fontSize:9, color:D.muted }}>{m.l}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── AI FEEDBACK ── */}
        <div style={{ background:`${B.indigo}18`, borderRadius:12, padding:"12px 14px", marginBottom:14 }}>
          <div style={{ fontSize:10, color:B.indigo, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:6 }}>🤖 Análise do Coach IA</div>
          <div style={{ fontSize:13, color:D.sub, lineHeight:1.7 }}>{feedback}</div>
        </div>

        {/* ── SHARE ── */}
        <button onClick={()=>{
          const metric = type==="running"&&log?.distance?`${log.distance}km @ ${log.pace||"?"}/km`:type==="bike"&&log?.distance?`${log.distance}km`:type==="swim"&&log?.volume?`${log.volume}m`:plan?.focus||"Treino";
          const text = `✅ ${plan?.focus||sp?.label} concluído!\n${metric} · Sensação ${log?.feel||"?"}/10\n\nCoach IA: "${feedback.slice(0,120)}..."\n\n#AthleteAI`;
          try { navigator.share({ title:"Athlete AI", text }); } catch { navigator.clipboard?.writeText(text); }
        }} style={{ width:"100%", padding:"10px 0", borderRadius:10, border:`1px solid ${D.border}`, background:"transparent", color:D.muted, fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif", marginBottom:8 }}>
          📤 Compartilhar treino
        </button>

        {/* ── DELETE / UNDO ── */}
        {!confirmDel ? (
          <button onClick={()=>setConfirmDel(true)} style={{
            width:"100%", padding:"10px 0", borderRadius:10, border:`1px solid ${D.border}`,
            background:"transparent", color:D.muted, fontSize:12, fontWeight:600,
            cursor:"pointer", fontFamily:"'Inter',sans-serif",
          }}>
            ↩️ Desfazer — remover este registro
          </button>
        ) : (
          <div style={{ background:`${B.red}12`, borderRadius:12, padding:"14px 16px", border:`1px solid ${B.red}30` }}>
            <div style={{ fontSize:13, fontWeight:700, color:D.text, marginBottom:12 }}>
              Remover este registro do histórico?
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={()=>setConfirmDel(false)} style={{
                flex:1, padding:"11px 0", borderRadius:10, border:`1.5px solid ${D.border}`,
                background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif",
              }}>Cancelar</button>
              <button onClick={onDelete} style={{
                flex:1, padding:"11px 0", borderRadius:10, border:"none",
                background:B.red, color:"#fff", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif",
              }}>🗑️ Remover</button>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ══ HOME TAB ════════════════════════════════════════════════════ */
function HomeTab({ profile, onStartWorkout, todayCompleted, workoutHistory, onDelete }) {
  const [showCI, setShowCI] = useState(false);
  const [checkin, setCheckin] = useState(null);
  const [ci, setCi] = useState({ sleep:7, energy:7, soreness:4 });

  const todayIdx = new Date().getDay();
  const todayName = dayName[todayIdx];
  const plan = profile?.plan?.weeklySchedule || [];
  const todayPlan = plan.find(p=>p.day===todayName) || { type:"rest", day:todayName };
  const h = new Date().getHours();
  const greet = h<12?"Bom dia":h<18?"Boa tarde":"Boa noite";
  const sc = SPORT_COLOR[todayPlan.type]||"#475569";
  const sp = SPORTS.find(s=>s.id===todayPlan.type);

  return (
    <div>
      <div style={{ background:`linear-gradient(160deg,${D.bg2} 0%,${D.bg} 100%)`, padding:"24px 20px 20px", borderBottom:`1px solid ${D.card}` }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
          <div>
            <div style={{ fontSize:12, color:D.muted }}>{greet} · {todayName}-feira</div>
            <div style={{ fontSize:24, fontWeight:900, color:D.text, letterSpacing:"-.5px", marginTop:4 }}>{profile?.basic?.name||"Atleta"}</div>
            <div style={{ fontSize:12, color:D.muted, marginTop:4 }}>{profile?.plan?.planName||"Plano personalizado"}</div>
          </div>
          <div style={{ width:42, height:42, borderRadius:12, background:`linear-gradient(135deg,${B.blue},${B.indigo})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, boxShadow:`0 4px 16px ${B.blue}44` }}>🏆</div>
        </div>
        <div style={{ display:"flex", gap:8, marginTop:14, flexWrap:"wrap" }}>
          {(profile?.selectedSports||[]).map(s=>{
            const sp2 = SPORTS.find(x=>x.id===s);
            return <SportBadge key={s} type={s} sm/>;
          })}
        </div>
      </div>

      <div style={{ padding:"16px", display:"flex", flexDirection:"column", gap:14 }}>
        {/* Fatigue alert */}
        {(()=>{
          const f = checkFatigue(workoutHistory||[]);
          if (!f) return null;
          return (
            <div style={{ background:`${B.red}15`, border:`1px solid ${B.red}30`, borderRadius:13, padding:"12px 16px" }}>
              <div style={{ fontSize:11, color:B.red, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:4 }}>⚠️ Alerta de Fadiga</div>
              <div style={{ fontSize:13, color:D.sub, lineHeight:1.6 }}>
                {f.type==="rpe"
                  ? `RPE médio das últimas sessões: ${f.avg}/10. Acima de 8.7 indica carga acumulada alta. Considere um treino regenerativo ou deload esta semana.`
                  : `Sensação média das últimas sessões: ${f.avg}/10. Sinais de recuperação insuficiente. Reveja sono, alimentação e volume semanal.`}
              </div>
            </div>
          );
        })()}

        {/* Check-in */}
        {!checkin ? (
          <div style={{ background:D.card, borderRadius:14, padding:"14px 18px", border:`1px solid ${B.indigo}30`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div><div style={{ fontSize:13, fontWeight:700, color:D.text, marginBottom:3 }}>📋 Check-in Diário</div><div style={{ fontSize:12, color:D.muted }}>Como você está se sentindo?</div></div>
            <button onClick={()=>setShowCI(true)} style={{ background:B.indigo, color:"#fff", border:"none", borderRadius:10, padding:"9px 16px", fontWeight:700, cursor:"pointer", fontSize:12, fontFamily:"'Inter',sans-serif" }}>Iniciar</button>
          </div>
        ) : (
          <div style={{ background:D.card, borderRadius:14, padding:"12px 18px", borderTop:`3px solid ${B.emerald}` }}>
            <div style={{ fontSize:10, color:B.emerald, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em" }}>✓ Check-in realizado</div>
            <div style={{ display:"flex", gap:18, marginTop:8 }}>
              {[["🌙",`${checkin.sleep}h sono`],["⚡",`Energia ${checkin.energy}/10`],["💪",`Dor ${checkin.soreness}/10`]].map(([ic,lb])=><span key={lb} style={{ fontSize:12, color:D.sub }}>{ic} {lb}</span>)}
            </div>
          </div>
        )}

        {/* Today's workout card OR summary after completion */}
        {todayCompleted ? (
          /* ── COMPLETED STATE — summary replaces the "Iniciar" card ── */
          <WorkoutSummaryCard entry={todayCompleted} history={workoutHistory||[]} profile={profile} onDelete={onDelete}/>
        ) : todayPlan.type!=="rest" ? (
          /* ── PENDING — show workout card ── */
          <Card accent={sc}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14 }}>
              <div>
                <SportBadge type={todayPlan.type} sm/>
                <div style={{ fontSize:22, fontWeight:900, color:D.text, marginTop:8, letterSpacing:"-.5px" }}>{todayPlan.focus||sp?.label||"Treino"}</div>
                <div style={{ fontSize:12, color:D.muted, marginTop:4 }}>
                  {todayPlan.type==="running"&&todayPlan.distance&&`${todayPlan.distance}km · ${todayPlan.paceTarget||""}/km · ${todayPlan.hrZone||""}`}
                  {todayPlan.type==="bike"&&todayPlan.distance&&`${todayPlan.distance}km · ${todayPlan.hrZone||""}`}
                  {todayPlan.type==="swim"&&todayPlan.distance&&`${todayPlan.distance}m · ${todayPlan.hrZone||""}`}
                  {todayPlan.type==="strength"&&"Treino de musculação"}
                </div>
                {todayPlan.notes&&<div style={{ fontSize:12, color:sc, marginTop:4, fontWeight:500 }}>{todayPlan.notes}</div>}
              </div>
              <div style={{ fontSize:34 }}>{sp?.emoji||"💪"}</div>
            </div>
            {profile?.plan?.aiCoachNote&&(
              <div style={{ background:`${B.indigo}15`, borderRadius:10, padding:"10px 12px", marginBottom:14 }}>
                <div style={{ fontSize:10, color:B.indigo, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:4 }}>🤖 Coach IA</div>
                <div style={{ fontSize:12, color:D.sub, lineHeight:1.5 }}>{profile.plan.aiCoachNote}</div>
              </div>
            )}
            <button onClick={()=>onStartWorkout(todayPlan)} style={{ width:"100%", padding:"13px 0", borderRadius:12, border:"none", background:sc, color:sc===B.green||sc===B.cyan?"#0A1628":"#fff", fontWeight:800, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 6px 24px ${sc}44` }}>
              ▶ {todayPlan.type==="strength"?"Iniciar Treino":todayPlan.type==="running"?"Iniciar Corrida":todayPlan.type==="bike"?"Iniciar Pedal":"Iniciar Natação"}
            </button>
          </Card>
        ) : (
          /* ── REST DAY ── */
          <Card style={{ textAlign:"center", padding:"28px 20px" }}>
            <div style={{ fontSize:44 }}>😴</div>
            <div style={{ fontSize:18, fontWeight:700, color:D.text, marginTop:12 }}>Dia de Descanso</div>
            <div style={{ fontSize:13, color:D.muted, marginTop:8, lineHeight:1.7 }}>Recuperação é parte do treino. Aproveite para alongar, caminhar levemente ou simplesmente descansar.</div>
          </Card>
        )}

        {/* Weekly schedule visual */}
        <Card>
          <Lbl>Semana Atual</Lbl>
          <div style={{ display:"flex", gap:5 }}>
            {plan.map((p,i) => {
              const isT = p.day===todayName;
              const c = SPORT_COLOR[p.type]||"#475569";
              const sp3 = SPORTS.find(s=>s.id===p.type);
              return (
                <div key={i} style={{ flex:1, textAlign:"center" }}>
                  <div style={{ fontSize:9, color:isT?D.text:D.muted, fontWeight:isT?700:400, marginBottom:5 }}>{p.day.slice(0,3)}</div>
                  <div style={{ width:"100%", aspectRatio:"1", borderRadius:9, background:isT?c:p.type==="rest"?D.bg2:`${c}25`, border:`${isT?"2px":"1px"} solid ${isT?c:p.type==="rest"?D.border:`${c}40`}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13 }}>
                    {p.type==="rest"?"·":sp3?.emoji||"💪"}
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{ display:"flex", gap:12, marginTop:10, justifyContent:"center", flexWrap:"wrap" }}>
            {[...new Set(plan.filter(p=>p.type!=="rest").map(p=>p.type))].map(t=>{
              const sp4=SPORTS.find(s=>s.id===t);
              return <div key={t} style={{ display:"flex", gap:5, alignItems:"center", fontSize:10, color:D.muted }}><div style={{ width:7, height:7, borderRadius:2, background:SPORT_COLOR[t] }}/>{sp4?.emoji} {sp4?.label}</div>;
            })}
          </div>
        </Card>

        {/* AI plan description */}
        {profile?.plan?.description&&(
          <Card style={{ background:`${B.indigo}10`, border:`1px solid ${B.indigo}25` }}>
            <Lbl>🤖 Sobre o seu plano</Lbl>
            <div style={{ fontSize:13, color:D.sub, lineHeight:1.7 }}>{profile.plan.description}</div>
          </Card>
        )}
      </div>

      {/* Check-in modal */}
      {showCI&&(
        <div style={{ position:"fixed", inset:0, background:"rgba(2,6,23,.85)", display:"flex", alignItems:"flex-end", zIndex:200 }}>
          <div className="slide" style={{ background:D.card, borderRadius:"20px 20px 0 0", padding:"24px 20px 40px", width:"100%", maxWidth:480, margin:"0 auto", borderTop:`1px solid ${D.border}` }}>
            <div style={{ width:36, height:4, borderRadius:2, background:D.border, margin:"0 auto 20px" }}/>
            <div style={{ fontSize:18, fontWeight:800, color:D.text, marginBottom:22 }}>📋 Como você está hoje?</div>
            {[{k:"sleep",l:"Horas dormidas",ic:"🌙",min:3,max:12,u:"h",c:B.indigo},{k:"energy",l:"Energia",ic:"⚡",min:1,max:10,u:"/10",c:B.blue},{k:"soreness",l:"Dor muscular",ic:"💪",min:1,max:10,u:"/10",c:B.purple}].map(f=>(
              <div key={f.k} style={{ marginBottom:20 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:9 }}>
                  <span style={{ fontSize:13, color:D.sub, fontWeight:500 }}>{f.ic} {f.l}</span>
                  <span style={{ fontSize:15, fontWeight:800, color:f.c }}>{ci[f.k]}{f.u}</span>
                </div>
                <input type="range" min={f.min} max={f.max} value={ci[f.k]} onChange={e=>setCi(p=>({...p,[f.k]:+e.target.value}))} style={{ width:"100%", accentColor:f.c, cursor:"pointer" }}/>
              </div>
            ))}
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={()=>setShowCI(false)} style={{ flex:1, padding:"12px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Cancelar</button>
              <button onClick={()=>{setCheckin(ci);setShowCI(false);}} style={{ flex:2, padding:"12px 0", borderRadius:12, border:"none", background:B.blue, color:"#fff", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 16px ${B.blue}44` }}>Confirmar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ══ WORKOUT TAB — routes to the right component ════════════════ */
function WorkoutTab({ todayPlan, profile, onLogWorkout }) {
  if (!todayPlan || todayPlan.type==="rest") return (
    <div style={{ padding:20, textAlign:"center", paddingTop:60 }}>
      <div style={{ fontSize:52 }}>😴</div>
      <div style={{ fontSize:20, fontWeight:700, color:D.text, marginTop:16 }}>Descanso hoje</div>
      <div style={{ color:D.muted, marginTop:8 }}>Próximo treino amanhã. Recupere bem!</div>
    </div>
  );
  if (todayPlan.type==="strength") return <StrengthWorkout plan={todayPlan} profile={profile} onFinish={onLogWorkout}/>;
  if (todayPlan.type==="running") return <RunWorkout plan={todayPlan} profile={profile} onFinish={onLogWorkout}/>;
  if (todayPlan.type==="bike") return <BikeWorkout plan={todayPlan} onFinish={onLogWorkout}/>;
  if (todayPlan.type==="swim") return <SwimWorkout plan={todayPlan} onFinish={onLogWorkout}/>;
  return null;
}

/* ── Strength Workout ─── */
function StrengthWorkout({ plan, profile, onFinish }) {
  const focusMap = {
    "Peito & Ombros":"chest_shoulder","Pernas":"legs","Costas & Bíceps":"back_biceps",
    "Posterior":"posterior","Peito & Tríceps":"chest_triceps","Corpo Todo":"full_body",
  };
  const focusKey = Object.keys(focusMap).find(k=>plan.focus?.includes(k.split(" ")[0])) || "full_body";
  const exList = EXERCISES[focusMap[focusKey]||"full_body"];
  const exp = profile?.sportData?.strength?.experience||"intermediate";
  const goal = profile?.sportData?.strength?.goal||"hypertrophy";
  const sets = goal==="strength"?5:goal==="hypertrophy"?4:3;
  const reps = goal==="strength"?"3-5":goal==="hypertrophy"?"8-12":"12-20";

  const [active, setActive] = useState(0);
  const [setIdx, setSetIdx] = useState(0);
  const [logs, setLogs] = useState({});
  const [rest, setRest] = useState(null);
  const [done, setDone] = useState(false);
  const timer = useRef(null);

  function logSet(exId, data) {
    setLogs(p=>({...p,[exId]:[...(p[exId]||[]),data]}));
    const restTime = goal==="strength"?180:90;
    setRest(restTime);
    timer.current = setInterval(()=>setRest(p=>{if(p<=1){clearInterval(timer.current);return null;}return p-1;}),1000);
  }
  useEffect(()=>()=>clearInterval(timer.current),[]);

  function advance() {
    const ex = exList[active];
    const done_sets = (logs[ex.id]||[]).length;
    if (done_sets < sets-1) { setSetIdx(done_sets+1); }
    else if (active < exList.length-1) { setActive(a=>a+1); setSetIdx(0); }
    else { setDone(true); }
  }

  if (done) return (
    <div className="pop" style={{ padding:20, display:"flex", flexDirection:"column", alignItems:"center", textAlign:"center", gap:20, paddingTop:40 }}>
      <div style={{ fontSize:72 }}>🏆</div>
      <div style={{ fontSize:26, fontWeight:900, color:D.text, letterSpacing:"-.5px" }}>Treino Concluído!</div>
      <Card style={{ background:`${B.indigo}15`, border:`1px solid ${B.indigo}30`, width:"100%", textAlign:"left" }}>
        <div style={{ fontSize:11, color:B.indigo, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:8 }}>🤖 Coach IA</div>
        <div style={{ fontSize:13, color:D.sub, lineHeight:1.7 }}>Excelente! Complete os dados de cada série para que a IA possa ajustar as cargas na próxima sessão de {plan.focus}.</div>
      </Card>
      <button onClick={()=>onFinish({type:"strength",plan,log:{logs,feel:8},date:new Date().toLocaleDateString("pt-BR")})} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:B.purple, color:"#fff", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>✓ Salvar e Concluir</button>
    </div>
  );

  const ex = exList[active];
  const exLogs = logs[ex.id]||[];
  const total = exList.length*sets;
  const doneSets = Object.values(logs).reduce((s,a)=>s+a.length,0);

  return (
    <div style={{ padding:20 }}>
      <SportBadge type="strength"/>
      <div style={{ fontSize:18, fontWeight:800, color:D.text, marginTop:8, marginBottom:16 }}>{plan.focus}</div>
      <div style={{ marginBottom:18 }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:D.muted, marginBottom:7 }}>
          <span>Progresso</span><span style={{ color:B.purple, fontWeight:700 }}>{doneSets}/{total} séries</span>
        </div>
        <div style={{ background:D.bg2, borderRadius:8, height:7, overflow:"hidden" }}>
          <div style={{ width:`${total>0?doneSets/total*100:0}%`, background:`linear-gradient(90deg,${B.purple},${B.indigo})`, height:"100%", borderRadius:8, transition:"width .4s" }}/>
        </div>
      </div>
      <Card style={{ marginBottom:14, background:`linear-gradient(135deg,#1e2545,${D.card})` }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:10 }}>
          <div>
            <div style={{ fontSize:10, color:D.muted, fontWeight:600, textTransform:"uppercase", letterSpacing:".08em" }}>Exercício {active+1}/{exList.length}</div>
            <div style={{ fontSize:19, fontWeight:800, color:D.text, marginTop:4 }}>{ex.name}</div>
            <div style={{ fontSize:12, color:D.muted, marginTop:2 }}>{ex.muscles}</div>
          </div>
          <div style={{ background:`${B.purple}25`, borderRadius:10, padding:"7px 11px", textAlign:"center" }}>
            <div style={{ fontSize:18, fontWeight:800, color:B.purple }}>{setIdx+1}/{sets}</div>
            <div style={{ fontSize:10, color:D.muted }}>série</div>
          </div>
        </div>
        <div style={{ background:`${B.indigo}15`, borderRadius:10, padding:"9px 12px", fontSize:12, color:D.sub }}>💡 {ex.tip}</div>
        <div style={{ fontSize:12, color:D.muted, marginTop:8 }}>Meta: {sets}×{reps} reps · Descanso: {ex.rest}s</div>
      </Card>
      {exLogs.length>0&&(
        <Card style={{ marginBottom:14, padding:14 }}>
          <Lbl>Séries anteriores</Lbl>
          <div style={{ display:"flex", gap:8 }}>
            {exLogs.map((l,i)=>(
              <div key={i} style={{ flex:1, background:D.bg2, borderRadius:10, padding:"8px 6px", textAlign:"center" }}>
                <div style={{ fontSize:10, color:D.muted }}>S{i+1}</div>
                <div style={{ fontSize:13, fontWeight:800, color:B.purple }}>{l.load}kg</div>
                <div style={{ fontSize:11, color:D.sub }}>{l.reps}×</div>
              </div>
            ))}
          </div>
        </Card>
      )}
      {!rest ? (
        <QuickSetLogger goal={goal} onLog={(data)=>{logSet(ex.id,data);}} onAdvance={advance}/>
      ) : (
        <Card style={{ textAlign:"center", padding:"24px 20px" }}>
          <Lbl>Descanso</Lbl>
          <div style={{ fontSize:60, fontWeight:900, color:B.purple, letterSpacing:"-2px" }}>{rest}s</div>
          <button onClick={()=>{clearInterval(timer.current);setRest(null);advance();}} style={{ marginTop:16, background:"transparent", border:`1.5px solid ${D.border}`, color:D.muted, borderRadius:10, padding:"9px 20px", cursor:"pointer", fontFamily:"'Inter',sans-serif", fontWeight:600, fontSize:13 }}>Pular descanso →</button>
        </Card>
      )}
      <div style={{ display:"flex", gap:5, marginTop:14, overflowX:"auto" }}>
        {exList.map((e,i)=>{
          const d2=(logs[e.id]||[]).length>=sets;
          return <div key={e.id} style={{ flexShrink:0, width:30, height:30, borderRadius:8, background:i===active?B.purple:d2?`${B.purple}30`:D.card, color:i===active?"#fff":d2?B.purple:D.muted, border:`1px solid ${i===active?B.purple:d2?`${B.purple}50`:D.border}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700 }}>{i+1}</div>;
        })}
      </div>
    </div>
  );
}

function QuickSetLogger({ goal, onLog, onAdvance }) {
  const defaultLoad = 20;
  const defaultReps = goal==="strength"?5:goal==="hypertrophy"?10:15;
  const [reps,setReps]=useState(defaultReps);
  const [load,setLoad]=useState(defaultLoad);
  const [rpe,setRpe]=useState(7);
  const rpeColor = rpe<=7?B.emerald:rpe===8?B.orange:B.red;
  const RPE={6:"Fácil",7:"Moderado",8:"Pesado",9:"Muito pesado",10:"Máximo"};
  return (
    <Card>
      <Lbl>Registrar série</Lbl>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:18 }}>
        <div>
          <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".08em", marginBottom:9 }}>Repetições</div>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <AdjBtn onPress={()=>setReps(p=>Math.max(1,p-1))} label="−"/>
            <span style={{ fontSize:30, fontWeight:900, color:D.text, minWidth:34, textAlign:"center" }}>{reps}</span>
            <AdjBtn onPress={()=>setReps(p=>p+1)} label="+" accent/>
          </div>
        </div>
        <div>
          <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".08em", marginBottom:9 }}>Carga (kg)</div>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <AdjBtn onPress={()=>setLoad(p=>Math.max(0,+(p-2.5).toFixed(1)))} label="−"/>
            <span style={{ fontSize:30, fontWeight:900, color:D.text, minWidth:34, textAlign:"center" }}>{load}</span>
            <AdjBtn onPress={()=>setLoad(p=>+(p+2.5).toFixed(1))} label="+" accent/>
          </div>
        </div>
      </div>
      <div style={{ marginBottom:18 }}>
        <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
          <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".08em" }}>RPE</div>
          <div style={{ fontSize:12, fontWeight:700, color:rpeColor }}>{rpe} · {RPE[rpe]}</div>
        </div>
        <div style={{ display:"flex", gap:5 }}>
          {[6,7,8,9,10].map(v=><button key={v} onClick={()=>setRpe(v)} style={{ flex:1, padding:"9px 0", borderRadius:8, border:"none", background:v===rpe?rpeColor:`${rpeColor}18`, color:v===rpe?"#fff":v<=rpe?rpeColor:D.muted, fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>{v}</button>)}
        </div>
      </div>
      <button onClick={()=>{onLog({reps,load,rpe});}} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:B.purple, color:"#fff", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 18px ${B.purple}44` }}>✓ Concluir Série</button>
    </Card>
  );
}

/* ── Run/Bike/Swim simple logger ─── */
/* ══ LIVE SESSION — cronômetro com timer real ════════════════════ */
function LiveSession({ plan, type, color, onComplete, onBack }) {
  const [elapsed, setElapsed] = useState(0);
  const [paused, setPaused] = useState(false);
  const [showLog, setShowLog] = useState(false);
  const [ld, setLd] = useState({
    distance: parseFloat(plan?.distance)||0,
    pace: plan?.paceTarget||"",
    hr: "", feel: 8, notes: "", time: "",
    avgPower: "", cadence: "90",
    volume: parseFloat(plan?.distance)||0,
    pace100: "", strokes: "",
  });

  useEffect(() => {
    if (paused || showLog) return;
    const id = setInterval(() => setElapsed(e => e+1), 1000);
    return () => clearInterval(id);
  }, [paused, showLog]);

  function fmt(s) {
    const h = Math.floor(s/3600);
    const m = Math.floor((s%3600)/60);
    const sec = s%60;
    if (h>0) return `${h}:${String(m).padStart(2,"0")}:${String(sec).padStart(2,"0")}`;
    return `${String(m).padStart(2,"0")}:${String(sec).padStart(2,"0")}`;
  }

  function handleFinalize() {
    setPaused(true);
    setLd(p => ({ ...p, time: fmt(elapsed) }));
    setShowLog(true);
  }

  const FIELDS = {
    running:[{l:"Distância percorrida (km)",k:"distance",tp:"number"},{l:"Pace médio (ex: 5:45)",k:"pace",tp:"text"},{l:"FC média (bpm)",k:"hr",tp:"number"}],
    bike:[{l:"Distância (km)",k:"distance",tp:"number"},{l:"Tempo",k:"time",tp:"text"},{l:"Potência média (W)",k:"avgPower",tp:"number"},{l:"Cadência (rpm)",k:"cadence",tp:"number"},{l:"FC média (bpm)",k:"hr",tp:"number"}],
    swim:[{l:"Volume nadado (metros)",k:"volume",tp:"number"},{l:"Tempo total",k:"time",tp:"text"},{l:"Pace/100m (ex: 2:10)",k:"pace100",tp:"text"},{l:"Braçadas/comp.",k:"strokes",tp:"number"},{l:"FC média (bpm)",k:"hr",tp:"number"}],
  };

  if (showLog) return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={()=>{ setShowLog(false); setPaused(false); }} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <div style={{ fontSize:17, fontWeight:800, color:D.text }}>📝 {plan?.focus||"Registrar"}</div>
      </div>
      <div style={{ background:`${color}18`, borderRadius:12, padding:"10px 16px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        <div><div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".08em" }}>Tempo registrado</div><div style={{ fontSize:22, fontWeight:800, color, marginTop:2 }}>⏱ {fmt(elapsed)}</div></div>
        <div style={{ fontSize:32 }}>{SPORTS.find(s=>s.id===type)?.emoji||"💪"}</div>
      </div>
      <Card>
        {(FIELDS[type]||[]).map(f=>(
          <div key={f.k} style={{ marginBottom:14 }}>
            <Lbl>{f.l}</Lbl>
            <input type={f.tp} value={ld[f.k]||""} onChange={e=>setLd(p=>({...p,[f.k]:f.tp==="number"?+e.target.value:e.target.value}))}
              style={{ width:"100%", background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/>
          </div>
        ))}
        <div style={{ marginBottom:18 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
            <Lbl>Sensação geral</Lbl><span style={{ fontSize:14, fontWeight:800, color }}>{ld.feel}/10</span>
          </div>
          <input type="range" min={1} max={10} value={ld.feel} onChange={e=>setLd(p=>({...p,feel:+e.target.value}))} style={{ width:"100%", accentColor:color }}/>
        </div>
        <div style={{ marginBottom:18 }}>
          <Lbl>Notas</Lbl>
          <textarea value={ld.notes||""} onChange={e=>setLd(p=>({...p,notes:e.target.value}))} placeholder="Como foi? Condições, sensações..."
            style={{ width:"100%", minHeight:60, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/>
        </div>
        <button onClick={()=>onComplete(ld)} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:color, color:color===B.green||color===B.cyan?"#0A1628":"#fff", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 18px ${color}44` }}>
          ✓ Salvar Atividade
        </button>
      </Card>
    </div>
  );

  return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={onBack} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <SportBadge type={type}/>
        {paused && <span style={{ fontSize:11, color:B.orange, fontWeight:700, background:`${B.orange}20`, padding:"3px 10px", borderRadius:20 }}>⏸ PAUSADO</span>}
      </div>

      {/* Timer */}
      <Card style={{ textAlign:"center", padding:"36px 20px", background:`linear-gradient(160deg,${D.bg2},${D.card})` }}>
        <div style={{ fontSize:11, color:paused?B.orange:D.muted, fontWeight:700, textTransform:"uppercase", letterSpacing:".14em", marginBottom:16 }}>
          {paused ? "⏸ Pausado" : "🔴 Em andamento"}
        </div>
        <div style={{ fontSize:68, fontWeight:900, color, letterSpacing:"-3px", textShadow:`0 0 40px ${color}44`, fontVariantNumeric:"tabular-nums" }}>
          {fmt(elapsed)}
        </div>
        <div style={{ fontSize:14, color:D.muted, marginTop:10 }}>{plan?.focus||"Sessão ativa"}</div>

        {/* Targets row */}
        <div style={{ display:"flex", justifyContent:"center", gap:20, marginTop:22, flexWrap:"wrap" }}>
          {plan?.distance&&<div style={{ textAlign:"center" }}>
            <div style={{ fontSize:20, fontWeight:800, color }}>{plan.distance}{type==="swim"?"m":"km"}</div>
            <div style={{ fontSize:10, color:D.muted }}>meta</div>
          </div>}
          {plan?.paceTarget&&<div style={{ textAlign:"center" }}>
            <div style={{ fontSize:20, fontWeight:800, color }}>{plan.paceTarget}/km</div>
            <div style={{ fontSize:10, color:D.muted }}>pace alvo</div>
          </div>}
          {plan?.hrZone&&<div style={{ textAlign:"center" }}>
            <div style={{ fontSize:20, fontWeight:800, color:B.red }}>{plan.hrZone}</div>
            <div style={{ fontSize:10, color:D.muted }}>zona FC</div>
          </div>}
          {plan?.cadenceTarget&&<div style={{ textAlign:"center" }}>
            <div style={{ fontSize:20, fontWeight:800, color }}>{plan.cadenceTarget}</div>
            <div style={{ fontSize:10, color:D.muted }}>cadência</div>
          </div>}
        </div>
      </Card>

      {/* Controls */}
      <div style={{ display:"flex", gap:12 }}>
        <button onClick={()=>setPaused(p=>!p)} style={{
          flex:1, padding:"15px 0", borderRadius:13,
          border:`2px solid ${D.border}`, background:"transparent",
          color:D.sub, fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif",
        }}>
          {paused ? "▶ Retomar" : "⏸ Pausar"}
        </button>
        <button onClick={handleFinalize} style={{
          flex:2, padding:"15px 0", borderRadius:13, border:"none",
          background:color, color:color===B.green||color===B.cyan?"#0A1628":"#fff",
          fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif",
          boxShadow:`0 6px 28px ${color}55`,
        }}>
          🏁 Finalizar e Registrar
        </button>
      </div>
      <div style={{ textAlign:"center", fontSize:11, color:D.muted }}>
        Você poderá editar os dados antes de salvar.
      </div>
    </div>
  );
}

function RunWorkout({ plan, onFinish }) {
  const [mode, setMode] = useState("idle"); // idle | live | log
  const [ld, setLd] = useState({ distance:plan?.distance||5, pace:plan?.paceTarget||"6:00", hr:"", feel:8, notes:"" });
  const finish = (data) => onFinish({ type:"running", plan, log:data, date:new Date().toLocaleDateString("pt-BR") });

  if (mode==="live") return <LiveSession plan={plan} type="running" color={B.blue} onComplete={finish} onBack={()=>setMode("idle")}/>;
  if (mode==="log") return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={()=>setMode("idle")} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <div style={{ fontSize:17, fontWeight:800, color:D.text }}>📝 {plan?.focus||"Registrar Corrida"}</div>
      </div>
      <Card>
        {[{l:"Distância (km)",k:"distance",tp:"number"},{l:"Pace médio (ex: 5:45)",k:"pace",tp:"text"},{l:"FC média (bpm)",k:"hr",tp:"number"}].map(f=>(
          <div key={f.k} style={{ marginBottom:14 }}>
            <Lbl>{f.l}</Lbl>
            <input type={f.tp} value={ld[f.k]||""} onChange={e=>setLd(p=>({...p,[f.k]:f.tp==="number"?+e.target.value:e.target.value}))}
              style={{ width:"100%", background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/>
          </div>
        ))}
        <div style={{ marginBottom:18 }}><div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}><Lbl>Sensação</Lbl><span style={{ fontSize:14, fontWeight:800, color:B.blue }}>{ld.feel}/10</span></div><input type="range" min={1} max={10} value={ld.feel} onChange={e=>setLd(p=>({...p,feel:+e.target.value}))} style={{ width:"100%", accentColor:B.blue }}/></div>
        <div style={{ marginBottom:18 }}><Lbl>Notas</Lbl><textarea value={ld.notes||""} onChange={e=>setLd(p=>({...p,notes:e.target.value}))} placeholder="Como foi?" style={{ width:"100%", minHeight:60, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/></div>
        <button onClick={()=>finish(ld)} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:B.blue, color:"#fff", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>✓ Salvar</button>
      </Card>
    </div>
  );

  return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <SportBadge type="running"/>
      <Card accent={B.blue}>
        <div style={{ fontSize:22, fontWeight:900, color:D.text, marginBottom:8, letterSpacing:"-.5px" }}>{plan?.focus||"Corrida"}</div>
        {plan?.notes&&<div style={{ fontSize:13, color:D.muted, marginBottom:14, lineHeight:1.6 }}>{plan.notes}</div>}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8, marginBottom:16 }}>
          {[{l:"Distância",v:plan?.distance?`${plan.distance}km`:"—",c:B.blue},{l:"Pace alvo",v:plan?.paceTarget?`${plan.paceTarget}/km`:"—",c:B.blue},{l:"Zona FC",v:plan?.hrZone||"—",c:B.red}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"9px 8px", textAlign:"center" }}>
              <div style={{ fontSize:13, fontWeight:800, color:m.c }}>{m.v}</div><div style={{ fontSize:10, color:D.muted, marginTop:3 }}>{m.l}</div>
            </div>
          ))}
        </div>
        <div style={{ display:"flex", gap:10 }}>
          <button onClick={()=>setMode("log")} style={{ flex:1, padding:"12px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.sub, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>📝 Só registrar</button>
          <button onClick={()=>setMode("live")} style={{ flex:2, padding:"12px 0", borderRadius:12, border:"none", background:B.blue, color:"#fff", fontWeight:800, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 18px ${B.blue}44` }}>▶ Iniciar Corrida</button>
        </div>
      </Card>
    </div>
  );
}

function BikeWorkout({ plan, onFinish }) {
  const [mode, setMode] = useState("idle");
  const [ld, setLd] = useState({ distance:plan?.distance||40, time:"", avgPower:"", cadence:"90", hr:"", feel:8, notes:"" });
  const finish = (data) => onFinish({ type:"bike", plan, log:data, date:new Date().toLocaleDateString("pt-BR") });

  if (mode==="live") return <LiveSession plan={plan} type="bike" color={B.green} onComplete={finish} onBack={()=>setMode("idle")}/>;
  if (mode==="log") return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={()=>setMode("idle")} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <div style={{ fontSize:17, fontWeight:800, color:D.text }}>📝 {plan?.focus||"Registrar Pedal"}</div>
      </div>
      <Card>
        {[{l:"Distância (km)",k:"distance",tp:"number"},{l:"Tempo (ex: 1:20)",k:"time",tp:"text"},{l:"Potência média (W)",k:"avgPower",tp:"number"},{l:"Cadência (rpm)",k:"cadence",tp:"number"},{l:"FC média (bpm)",k:"hr",tp:"number"}].map(f=>(
          <div key={f.k} style={{ marginBottom:14 }}><Lbl>{f.l}</Lbl><input type={f.tp} value={ld[f.k]||""} onChange={e=>setLd(p=>({...p,[f.k]:f.tp==="number"?+e.target.value:e.target.value}))} style={{ width:"100%", background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/></div>
        ))}
        <div style={{ marginBottom:18 }}><div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}><Lbl>Sensação</Lbl><span style={{ fontSize:14, fontWeight:800, color:B.green }}>{ld.feel}/10</span></div><input type="range" min={1} max={10} value={ld.feel} onChange={e=>setLd(p=>({...p,feel:+e.target.value}))} style={{ width:"100%", accentColor:B.green }}/></div>
        <div style={{ marginBottom:18 }}><Lbl>Notas</Lbl><textarea value={ld.notes||""} onChange={e=>setLd(p=>({...p,notes:e.target.value}))} placeholder="Como foi?" style={{ width:"100%", minHeight:60, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/></div>
        <button onClick={()=>finish(ld)} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:B.green, color:"#0A1628", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>✓ Salvar</button>
      </Card>
    </div>
  );

  return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <SportBadge type="bike"/>
      <Card accent={B.green}>
        <div style={{ fontSize:22, fontWeight:900, color:D.text, marginBottom:8, letterSpacing:"-.5px" }}>{plan?.focus||"Ciclismo"}</div>
        {plan?.notes&&<div style={{ fontSize:13, color:D.muted, marginBottom:14, lineHeight:1.6 }}>{plan.notes}</div>}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:16 }}>
          {[{l:"Distância",v:plan?.distance?`${plan.distance}km`:"—",c:B.green},{l:"Zona",v:plan?.hrZone||"—",c:B.orange}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"9px 8px", textAlign:"center" }}><div style={{ fontSize:13, fontWeight:800, color:m.c }}>{m.v}</div><div style={{ fontSize:10, color:D.muted, marginTop:3 }}>{m.l}</div></div>
          ))}
        </div>
        <div style={{ display:"flex", gap:10 }}>
          <button onClick={()=>setMode("log")} style={{ flex:1, padding:"12px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.sub, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>📝 Só registrar</button>
          <button onClick={()=>setMode("live")} style={{ flex:2, padding:"12px 0", borderRadius:12, border:"none", background:B.green, color:"#0A1628", fontWeight:800, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 18px ${B.green}44` }}>▶ Iniciar Pedal</button>
        </div>
      </Card>
    </div>
  );
}

function SwimWorkout({ plan, onFinish }) {
  const [mode, setMode] = useState("idle");
  const [ld, setLd] = useState({ volume:plan?.distance||1500, time:"", pace100:"", strokes:"", hr:"", feel:8, notes:"" });
  const finish = (data) => onFinish({ type:"swim", plan, log:data, date:new Date().toLocaleDateString("pt-BR") });

  if (mode==="live") return <LiveSession plan={plan} type="swim" color={B.cyan} onComplete={finish} onBack={()=>setMode("idle")}/>;
  if (mode==="log") return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={()=>setMode("idle")} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <div style={{ fontSize:17, fontWeight:800, color:D.text }}>📝 {plan?.focus||"Registrar Natação"}</div>
      </div>
      <Card>
        {[{l:"Volume (metros)",k:"volume",tp:"number"},{l:"Tempo total (ex: 40:00)",k:"time",tp:"text"},{l:"Pace/100m (ex: 2:10)",k:"pace100",tp:"text"},{l:"Braçadas/comp.",k:"strokes",tp:"number"},{l:"FC média (bpm)",k:"hr",tp:"number"}].map(f=>(
          <div key={f.k} style={{ marginBottom:14 }}><Lbl>{f.l}</Lbl><input type={f.tp} value={ld[f.k]||""} onChange={e=>setLd(p=>({...p,[f.k]:f.tp==="number"?+e.target.value:e.target.value}))} style={{ width:"100%", background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:15, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/></div>
        ))}
        <div style={{ marginBottom:18 }}><div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}><Lbl>Sensação</Lbl><span style={{ fontSize:14, fontWeight:800, color:B.cyan }}>{ld.feel}/10</span></div><input type="range" min={1} max={10} value={ld.feel} onChange={e=>setLd(p=>({...p,feel:+e.target.value}))} style={{ width:"100%", accentColor:B.cyan }}/></div>
        <div style={{ marginBottom:18 }}><Lbl>Notas</Lbl><textarea value={ld.notes||""} onChange={e=>setLd(p=>({...p,notes:e.target.value}))} placeholder="Como foi?" style={{ width:"100%", minHeight:60, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:13, fontFamily:"'Inter',sans-serif", resize:"none", boxSizing:"border-box" }}/></div>
        <button onClick={()=>finish(ld)} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:B.cyan, color:"#0A1628", fontWeight:800, fontSize:15, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>✓ Salvar</button>
      </Card>
    </div>
  );

  return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:14 }}>
      <SportBadge type="swim"/>
      <Card accent={B.cyan}>
        <div style={{ fontSize:22, fontWeight:900, color:D.text, marginBottom:8, letterSpacing:"-.5px" }}>{plan?.focus||"Natação"}</div>
        {plan?.notes&&<div style={{ fontSize:13, color:D.muted, marginBottom:14, lineHeight:1.6 }}>{plan.notes}</div>}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:16 }}>
          {[{l:"Volume",v:plan?.distance?`${plan.distance}m`:"—",c:B.cyan},{l:"Zona",v:plan?.hrZone||"—",c:B.blue}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"9px 8px", textAlign:"center" }}><div style={{ fontSize:13, fontWeight:800, color:m.c }}>{m.v}</div><div style={{ fontSize:10, color:D.muted, marginTop:3 }}>{m.l}</div></div>
          ))}
        </div>
        <div style={{ display:"flex", gap:10 }}>
          <button onClick={()=>setMode("log")} style={{ flex:1, padding:"12px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.sub, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>📝 Só registrar</button>
          <button onClick={()=>setMode("live")} style={{ flex:2, padding:"12px 0", borderRadius:12, border:"none", background:B.cyan, color:"#0A1628", fontWeight:800, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif", boxShadow:`0 4px 18px ${B.cyan}44` }}>▶ Iniciar Natação</button>
        </div>
      </Card>
    </div>
  );
}


/* ══ COACH TAB ══════════════════════════════════════════════════ */
const QUICK = ["Estou cansado, devo treinar?","Rodei menos que o planejado hoje","Meu pace não melhora há semanas","Quero pular o treino de amanhã","Estou treinando há 2 meses e não vejo resultado","Aumentei carga mas sinto que não evoluí"];

function CoachTab({ messages, input, loading, chatRef, onInput, onSend, profile }) {
  return (
    <div style={{ display:"flex", flexDirection:"column", height:"calc(100vh - 130px)" }}>
      <div style={{ padding:"14px 20px", background:D.bar, borderBottom:`1px solid ${D.card}` }}>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <div style={{ width:42, height:42, borderRadius:12, background:`linear-gradient(135deg,${B.indigo},${B.blue})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, boxShadow:`0 4px 16px ${B.indigo}55` }}>🤖</div>
          <div>
            <div style={{ fontSize:15, fontWeight:700, color:D.text }}>Coach IA</div>
            <div style={{ fontSize:11, color:B.indigo }}>● {(profile?.selectedSports||[]).map(s=>SPORTS.find(x=>x.id===s)?.label).join(" · ")}</div>
          </div>
        </div>
      </div>
      <div ref={chatRef} style={{ flex:1, overflowY:"auto", padding:"16px", display:"flex", flexDirection:"column", gap:12 }}>
        {messages.length===0&&(
          <div className="fade" style={{ textAlign:"center", padding:"16px 8px" }}>
            <div style={{ width:56, height:56, borderRadius:16, background:`linear-gradient(135deg,${B.indigo},${B.blue})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26, margin:"0 auto 14px", boxShadow:`0 8px 24px ${B.indigo}44` }}>🤖</div>
            <div style={{ fontSize:16, fontWeight:700, color:D.text, marginBottom:6 }}>{profile?.basic?.name||"Atleta"}, aqui não tem elogio fácil.</div>
            <div style={{ fontSize:13, color:D.muted, marginBottom:18, lineHeight:1.7 }}>Sou seu coach — direto, baseado em dados, sem paciência para desculpa. Se quer "parabéns por ter saído do sofá", use outro app. Se quer evoluir de verdade, pergunte.</div>
            <div style={{ display:"flex", flexWrap:"wrap", gap:7, justifyContent:"center" }}>
              {QUICK.map(m=><button key={m} onClick={()=>onInput(m)} style={{ background:D.card, border:`1px solid ${D.border}`, borderRadius:20, color:D.sub, padding:"7px 13px", fontSize:12, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>{m}</button>)}
            </div>
          </div>
        )}
        {messages.map((m,i)=>(
          <div key={i} style={{ display:"flex", flexDirection:m.role==="user"?"row-reverse":"row", gap:8, alignItems:"flex-end" }}>
            {m.role==="assistant"&&<div style={{ width:28, height:28, borderRadius:8, flexShrink:0, background:`linear-gradient(135deg,${B.indigo},${B.blue})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13 }}>🤖</div>}
            <div style={{ maxWidth:"80%", padding:"11px 14px", fontSize:14, lineHeight:1.65, borderRadius:m.role==="user"?"16px 4px 16px 16px":"4px 16px 16px 16px", background:m.role==="user"?B.blue:D.card, color:"#fff" }}>
              {m.role==="assistant"&&<div style={{ fontSize:10, color:B.indigo, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:5 }}>Coach IA</div>}
              {m.content}
            </div>
          </div>
        ))}
        {loading&&(
          <div style={{ display:"flex", gap:8, alignItems:"flex-end" }}>
            <div style={{ width:28, height:28, borderRadius:8, background:`linear-gradient(135deg,${B.indigo},${B.blue})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13 }}>🤖</div>
            <div style={{ background:D.card, padding:"12px 16px", borderRadius:"4px 16px 16px 16px" }}>
              <div style={{ display:"flex", gap:5 }}>{[0,1,2].map(i=><div key={i} style={{ width:7, height:7, borderRadius:"50%", background:B.indigo, animation:`bounce 1s ease ${i*.15}s infinite` }}/>)}</div>
            </div>
          </div>
        )}
      </div>
      <div style={{ padding:"10px 14px 14px", background:D.bar, borderTop:`1px solid ${D.card}` }}>
        <div style={{ display:"flex", gap:8 }}>
          <input value={input} onChange={e=>onInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&onSend()} placeholder="Mensagem para o Coach IA..." style={{ flex:1, background:D.card, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"12px 14px", fontSize:14, fontFamily:"'Inter',sans-serif" }}/>
          <button onClick={onSend} disabled={!input.trim()||loading} style={{ width:46, height:46, borderRadius:12, border:"none", background:input.trim()&&!loading?B.blue:D.card, color:input.trim()&&!loading?"#fff":D.border, fontWeight:800, fontSize:18, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}>↑</button>
        </div>
      </div>
    </div>
  );
}

/* ══ PROGRESS TAB ════════════════════════════════════════════════ */
const ANGLES=[{key:"front",label:"Frente",icon:"🧍"},{key:"back",label:"Costas",icon:"🔄"},{key:"left",label:"Lat. Esq.",icon:"◀"},{key:"right",label:"Lat. Dir.",icon:"▶"}];

/* ══ WORKOUT DETAIL MODAL ════════════════════════════════════════ */
function WorkoutDetail({ entry, onClose, onDelete }) {
  const [confirmDel, setConfirmDel] = useState(false);
  const { type, plan, log, date } = entry;
  const sp = SPORTS.find(s=>s.id===type);
  const sc = SPORT_COLOR[type]||B.blue;
  const feedback = localFeedback(type, plan||{}, log||{});

  const metrics = {
    running:[
      {l:"Planejado",v:plan?.distance?`${plan.distance}km`:"—"},
      {l:"Executado",v:log?.distance?`${log.distance}km`:"—",bold:true,c:sc},
      {l:"Pace planejado",v:plan?.paceTarget?`${plan.paceTarget}/km`:"—"},
      {l:"Pace executado",v:log?.pace?`${log.pace}/km`:"—",bold:true,c:sc},
      {l:"Zona FC",v:plan?.hrZone||"—"},
      {l:"FC média",v:log?.hr?`${log.hr}bpm`:"—"},
      {l:"Sensação",v:log?.feel?`${log.feel}/10`:"—",bold:true},
    ],
    strength:[
      {l:"Sessão",v:plan?.focus||"—"},
      {l:"Séries totais",v:log?.logs?`${Object.values(log.logs).reduce((s,a)=>s+a.length,0)}`:"—",bold:true},
      {l:"Sensação",v:log?.feel?`${log.feel}/10`:"—",bold:true},
    ],
    bike:[
      {l:"Planejado",v:plan?.distance?`${plan.distance}km`:"—"},
      {l:"Executado",v:log?.distance?`${log.distance}km`:"—",bold:true,c:sc},
      {l:"Tempo",v:log?.time||"—"},
      {l:"Potência média",v:log?.avgPower?`${log.avgPower}W`:"—",bold:true},
      {l:"Cadência",v:log?.cadence?`${log.cadence}rpm`:"—"},
      {l:"FC média",v:log?.hr?`${log.hr}bpm`:"—"},
      {l:"Sensação",v:log?.feel?`${log.feel}/10`:"—",bold:true},
    ],
    swim:[
      {l:"Planejado",v:plan?.distance?`${plan.distance}m`:"—"},
      {l:"Executado",v:log?.volume?`${log.volume}m`:"—",bold:true,c:sc},
      {l:"Tempo",v:log?.time||"—"},
      {l:"Pace/100m",v:log?.pace100||"—",bold:true},
      {l:"Braçadas/comp.",v:log?.strokes?`${log.strokes}`:"—"},
      {l:"Sensação",v:log?.feel?`${log.feel}/10`:"—",bold:true},
    ],
  };

  return (
    <div style={{ position:"fixed", inset:0, background:D.bg, zIndex:300, overflowY:"auto", fontFamily:"'Inter',sans-serif" }}>
      {/* Header */}
      <div style={{ background:D.bg2, padding:"16px 20px", borderBottom:`1px solid ${D.card}`, display:"flex", alignItems:"center", gap:12, position:"sticky", top:0, zIndex:10 }}>
        <button onClick={onClose} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:22, padding:0, lineHeight:1 }}>←</button>
        <div style={{ flex:1 }}>
          <div style={{ fontSize:16, fontWeight:800, color:D.text }}>{plan?.focus||sp?.label||"Treino"}</div>
          <div style={{ fontSize:11, color:D.muted }}>{sp?.emoji} {sp?.label} · {date}</div>
        </div>
        <div style={{ width:36, height:36, borderRadius:10, background:`${sc}25`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>{sp?.emoji||"💪"}</div>
      </div>

      <div style={{ padding:"16px 20px", display:"flex", flexDirection:"column", gap:14 }}>
        {/* Metrics */}
        <Card>
          <Lbl>Métricas da sessão</Lbl>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
            {(metrics[type]||[]).map((m,i)=>(
              <div key={i} style={{ background:D.bg2, borderRadius:10, padding:"10px 12px" }}>
                <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em", marginBottom:4 }}>{m.l}</div>
                <div style={{ fontSize:m.bold?16:14, fontWeight:m.bold?800:500, color:m.c||D.sub }}>{m.v}</div>
              </div>
            ))}
          </div>
        </Card>

        {/* Strength sets detail */}
        {type==="strength" && log?.logs && Object.keys(log.logs).length>0 && (
          <Card>
            <Lbl>Séries registradas</Lbl>
            {Object.entries(log.logs).map(([exId, sets])=>{
              const ex = Object.values(EXERCISES).flat().find(e=>e.id===exId);
              return (
                <div key={exId} style={{ marginBottom:14 }}>
                  <div style={{ fontSize:13, fontWeight:700, color:D.text, marginBottom:8 }}>{ex?.name||exId}</div>
                  <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                    {sets.map((s,i)=>(
                      <div key={i} style={{ background:D.bg2, borderRadius:8, padding:"7px 12px", textAlign:"center", minWidth:64 }}>
                        <div style={{ fontSize:10, color:D.muted }}>Série {i+1}</div>
                        <div style={{ fontSize:13, fontWeight:800, color:B.purple }}>{s.load}kg</div>
                        <div style={{ fontSize:11, color:D.sub }}>{s.reps}× · RPE {s.rpe}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </Card>
        )}

        {/* Notes */}
        {log?.notes && (
          <Card>
            <Lbl>📝 Suas anotações</Lbl>
            <div style={{ fontSize:13, color:D.sub, lineHeight:1.7, fontStyle:"italic" }}>"{log.notes}"</div>
          </Card>
        )}

        {/* Coach feedback */}
        <Card style={{ background:`${B.indigo}12`, border:`1px solid ${B.indigo}25` }}>
          <Lbl>🤖 Análise do Coach IA</Lbl>
          <div style={{ fontSize:13, color:D.sub, lineHeight:1.7 }}>{feedback}</div>
        </Card>

        {/* Delete */}
        {!confirmDel ? (
          <button onClick={()=>setConfirmDel(true)} style={{ padding:"11px 0", borderRadius:10, border:`1px solid ${D.border}`, background:"transparent", color:D.muted, fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>
            🗑️ Remover este registro
          </button>
        ) : (
          <div style={{ background:`${B.red}12`, borderRadius:12, padding:"14px 16px", border:`1px solid ${B.red}30` }}>
            <div style={{ fontSize:13, fontWeight:700, color:D.text, marginBottom:12 }}>Remover permanentemente?</div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={()=>setConfirmDel(false)} style={{ flex:1, padding:"11px 0", borderRadius:10, border:`1.5px solid ${D.border}`, background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Cancelar</button>
              <button onClick={()=>{onDelete(entry);onClose();}} style={{ flex:1, padding:"11px 0", borderRadius:10, border:"none", background:B.red, color:"#fff", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Remover</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ══ WORKOUT HISTORY LIST ════════════════════════════════════════ */
function WorkoutHistoryList({ history, onSelect, onDelete }) {
  const [filter, setFilter] = useState("all");
  const [month, setMonth] = useState(null); // null = all months

  const allMonths = [...new Set(history.map(w=>{
    const p = w.date?.split("/"); return p?.length===3 ? `${p[2]}-${p[1]}` : null;
  }).filter(Boolean))].sort().reverse();

  const filtered = history.filter(w=>{
    if (filter!=="all" && w.type!==filter) return false;
    if (month) { const p=w.date?.split("/"); return p?.length===3 && `${p[2]}-${p[1]}`===month; }
    return true;
  }).slice().reverse();

  const FILTERS = [
    {id:"all",l:"Todos"},
    ...["strength","running","bike","swim"].map(id=>({id,l:SPORTS.find(s=>s.id===id)?.emoji||id})),
  ];

  return (
    <div>
      {/* Sport filter */}
      <div style={{ display:"flex", gap:6, marginBottom:12, overflowX:"auto", paddingBottom:4 }}>
        {FILTERS.map(f=>(
          <button key={f.id} onClick={()=>setFilter(f.id)} style={{
            flexShrink:0, padding:"7px 14px", borderRadius:20, cursor:"pointer",
            border:`1.5px solid ${filter===f.id?B.blue:D.border}`,
            background:filter===f.id?`${B.blue}20`:D.card,
            color:filter===f.id?B.blue:D.muted,
            fontWeight:filter===f.id?700:400, fontSize:12, fontFamily:"'Inter',sans-serif",
          }}>{f.l}</button>
        ))}
      </div>

      {/* Month filter */}
      {allMonths.length>1 && (
        <div style={{ display:"flex", gap:6, marginBottom:14, overflowX:"auto", paddingBottom:4 }}>
          <button onClick={()=>setMonth(null)} style={{ flexShrink:0, padding:"5px 12px", borderRadius:16, cursor:"pointer", border:`1px solid ${!month?B.indigo:D.border}`, background:!month?`${B.indigo}20`:D.card, color:!month?B.indigo:D.muted, fontSize:11, fontFamily:"'Inter',sans-serif", fontWeight:!month?700:400 }}>Todos</button>
          {allMonths.map(m=>{
            const [y,mo] = m.split("-");
            const label = new Date(+y,+mo-1,1).toLocaleDateString("pt-BR",{month:"short",year:"numeric"});
            return <button key={m} onClick={()=>setMonth(m)} style={{ flexShrink:0, padding:"5px 12px", borderRadius:16, cursor:"pointer", border:`1px solid ${month===m?B.indigo:D.border}`, background:month===m?`${B.indigo}20`:D.card, color:month===m?B.indigo:D.muted, fontSize:11, fontFamily:"'Inter',sans-serif", fontWeight:month===m?700:400 }}>{label}</button>;
          })}
        </div>
      )}

      {/* Count */}
      <div style={{ fontSize:12, color:D.muted, marginBottom:10 }}>
        {filtered.length} {filtered.length===1?"treino":"treinos"} encontrado{filtered.length===1?"":"s"}
      </div>

      {/* List */}
      {filtered.length===0 ? (
        <div style={{ textAlign:"center", padding:"40px 20px", color:D.muted }}>
          <div style={{ fontSize:36, marginBottom:12 }}>📭</div>
          <div style={{ fontSize:14 }}>Nenhum treino neste filtro ainda.</div>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {filtered.map((w,i)=>{
            const sp = SPORTS.find(s=>s.id===w.type);
            const sc = SPORT_COLOR[w.type]||B.blue;
            const metric = w.type==="running"&&w.log?.distance?`${w.log.distance}km · ${w.log.pace||""}`:
                           w.type==="bike"&&w.log?.distance?`${w.log.distance}km`:
                           w.type==="swim"&&w.log?.volume?`${w.log.volume}m · ${w.log.pace100||""}`:
                           w.type==="strength"&&w.log?.logs?`${Object.values(w.log.logs||{}).reduce((s,a)=>s+a.length,0)} séries`:"Registrado";
            const feel = w.log?.feel;
            const feelColor = !feel?"":feel>=8?B.emerald:feel>=6?B.orange:B.red;
            return (
              <button key={i} onClick={()=>onSelect(w)} style={{
                display:"flex", gap:12, alignItems:"center",
                background:D.card, borderRadius:13, padding:"12px 14px",
                border:`1px solid ${D.border}`, cursor:"pointer",
                textAlign:"left", fontFamily:"'Inter',sans-serif", width:"100%",
              }}>
                <div style={{ width:42, height:42, borderRadius:11, background:`${sc}25`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, flexShrink:0 }}>{sp?.emoji||"💪"}</div>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:14, fontWeight:700, color:D.text, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{w.plan?.focus||sp?.label||"Treino"}</div>
                  <div style={{ fontSize:11, color:sc, fontWeight:600, marginTop:2 }}>{metric}</div>
                  <div style={{ fontSize:10, color:D.muted, marginTop:2 }}>{w.date}</div>
                </div>
                {feel && (
                  <div style={{ textAlign:"center", flexShrink:0 }}>
                    <div style={{ fontSize:15, fontWeight:800, color:feelColor }}>{feel}</div>
                    <div style={{ fontSize:9, color:D.muted }}>sensação</div>
                  </div>
                )}
                <div style={{ color:D.muted, fontSize:16, flexShrink:0 }}>›</div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ══ PROGRESS TAB ════════════════════════════════════════════════ */
function ProgressTab({ profile, workoutHistory, onDeleteEntry, userId }) {
  const [view,setView]=useState("overview");
  const [selectedEntry, setSelectedEntry] = useState(null);
  const sports = profile?.selectedSports||[];
  const VIEWS=[
    {id:"overview",l:"Geral"},
    {id:"history",l:"📋 Histórico"},
    ...sports.map(s=>({id:s,l:SPORTS.find(x=>x.id===s)?.emoji+" "+SPORTS.find(x=>x.id===s)?.label})),
    {id:"photos",l:"📸 Fotos"},
  ];

  if (selectedEntry) return (
    <WorkoutDetail
      entry={selectedEntry}
      onClose={()=>setSelectedEntry(null)}
      onDelete={(e)=>{ onDeleteEntry&&onDeleteEntry(e); setSelectedEntry(null); }}
    />
  );

  return (
    <div style={{ paddingBottom:20 }}>
      <div style={{ background:D.bg2, borderBottom:`1px solid ${D.card}` }}>
        <div style={{ padding:"16px 20px 0" }}>
          <div style={{ fontSize:22, fontWeight:900, color:D.text, letterSpacing:"-.5px", marginBottom:14 }}>Progresso</div>
        </div>
        <div style={{ display:"flex", gap:0, padding:"0 16px", overflowX:"auto" }}>
          {VIEWS.map(v=>(
            <button key={v.id} onClick={()=>setView(v.id)} style={{ flexShrink:0, padding:"10px 14px", background:"transparent", color:view===v.id?D.text:D.muted, border:"none", borderBottom:`2px solid ${view===v.id?B.blue:"transparent"}`, fontWeight:view===v.id?700:500, fontSize:12, cursor:"pointer", fontFamily:"'Inter',sans-serif", whiteSpace:"nowrap" }}>{v.l}</button>
          ))}
        </div>
      </div>

      <div style={{ padding:"16px", display:"flex", flexDirection:"column", gap:14 }}>

        {/* ── OVERVIEW ── */}
        {view==="overview"&&(
          <>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
              {[
                {ic:"🏋️",lb:"Total registrado",v:workoutHistory.length,u:"treinos",c:B.purple},
                {ic:"📅",lb:"Esta semana",v:workoutHistory.filter(w=>new Date()-parseLogDate(w.date)<7*864e5).length,u:"treinos",c:B.blue},
                {ic:"🔥",lb:"Semanas consecutivas",v:Math.max(1,Math.ceil(workoutHistory.length/4)),u:"sem.",c:B.orange},
                {ic:"⏱️",lb:"Horas estimadas",v:(workoutHistory.length*0.85).toFixed(0),u:"h",c:B.emerald},
              ].map(m=>(
                <Card key={m.lb} accent={m.c} style={{ padding:14 }}>
                  <Lbl>{m.ic} {m.lb}</Lbl>
                  <div style={{ display:"flex", alignItems:"baseline", gap:4, marginTop:4 }}>
                    <span style={{ fontSize:24, fontWeight:800, color:D.text }}>{m.v}</span>
                    <span style={{ fontSize:11, color:D.muted }}>{m.u}</span>
                  </div>
                </Card>
              ))}
            </div>

            {/* Recent 4 */}
            <Card>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                <Lbl>Últimas atividades</Lbl>
                {workoutHistory.length>4&&(
                  <button onClick={()=>setView("history")} style={{ fontSize:12, color:B.blue, background:"transparent", border:"none", cursor:"pointer", fontFamily:"'Inter',sans-serif", fontWeight:600 }}>Ver tudo →</button>
                )}
              </div>
              {workoutHistory.length===0 ? (
                <div style={{ textAlign:"center", padding:"24px 0", color:D.muted }}>
                  <div style={{ fontSize:32 }}>📊</div>
                  <div style={{ marginTop:10, fontSize:13 }}>Seus treinos aparecerão aqui conforme você for registrando.</div>
                </div>
              ) : (
                workoutHistory.slice(-4).reverse().map((w,i)=>{
                  const sp2=SPORTS.find(s=>s.id===w.type);
                  const c=SPORT_COLOR[w.type]||B.blue;
                  const metric=w.type==="running"&&w.log?.distance?`${w.log.distance}km`:w.type==="swim"&&w.log?.volume?`${w.log.volume}m`:w.type==="bike"&&w.log?.distance?`${w.log.distance}km`:"Feito";
                  return (
                    <button key={i} onClick={()=>setSelectedEntry(w)} style={{ display:"flex", gap:12, alignItems:"center", padding:"10px 0", borderBottom:i<3?`1px solid ${D.bg2}`:"none", background:"transparent", border:"none", borderBottom:i<Math.min(workoutHistory.length,4)-1?`1px solid ${D.bg2}`:"none", cursor:"pointer", width:"100%", textAlign:"left", fontFamily:"'Inter',sans-serif" }}>
                      <div style={{ width:36, height:36, borderRadius:10, background:`${c}20`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, flexShrink:0 }}>{sp2?.emoji||"💪"}</div>
                      <div style={{ flex:1 }}>
                        <div style={{ fontSize:13, fontWeight:600, color:D.text }}>{w.plan?.focus||sp2?.label||"Treino"}</div>
                        <div style={{ fontSize:11, color:D.muted }}>{w.date}</div>
                      </div>
                      <div style={{ fontSize:11, color:c, fontWeight:700 }}>{metric}</div>
                      <span style={{ color:D.muted, fontSize:14 }}>›</span>
                    </button>
                  );
                })
              )}
            </Card>

            {/* Per-sport quick stats */}
            {sports.length>0&&(
              <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                {sports.map(spId=>{
                  const sp2=SPORTS.find(s=>s.id===spId);
                  const sc=SPORT_COLOR[spId];
                  const acc=calcAccumulated(workoutHistory,spId);
                  return (
                    <button key={spId} onClick={()=>setView(spId)} style={{ display:"flex", gap:12, alignItems:"center", background:D.card, borderRadius:13, padding:"13px 16px", border:`1px solid ${D.border}`, cursor:"pointer", textAlign:"left", fontFamily:"'Inter',sans-serif", width:"100%" }}>
                      <div style={{ width:40, height:40, borderRadius:10, background:`${sc}25`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>{sp2?.emoji}</div>
                      <div style={{ flex:1 }}>
                        <div style={{ fontSize:14, fontWeight:700, color:sp2?.color }}>{sp2?.label}</div>
                        <div style={{ fontSize:11, color:D.muted, marginTop:2 }}>
                          {spId==="strength"?`${acc.weekSessions||0} treinos esta semana · ${acc.total||0} total`:
                           spId==="running"?`${(acc.week||0).toFixed(1)}km esta semana · ${(acc.total||0).toFixed(1)}km total`:
                           spId==="bike"?`${(acc.week||0).toFixed(1)}km esta semana · ${(acc.total||0).toFixed(1)}km total`:
                           spId==="swim"?`${(acc.week||0).toFixed(0)}m esta semana · ${(acc.total||0).toFixed(0)}m total`:""}
                        </div>
                      </div>
                      <span style={{ color:D.muted, fontSize:14 }}>›</span>
                    </button>
                  );
                })}
              </div>
            )}
          </>
        )}

        {/* ── HISTORY ── */}
        {view==="history"&&(
          <WorkoutHistoryList
            history={workoutHistory}
            onSelect={setSelectedEntry}
            onDelete={onDeleteEntry}
          />
        )}

        {/* ── PER-SPORT ── */}
        {sports.includes(view)&&(
          <SportProgress type={view} history={workoutHistory.filter(w=>w.type===view)} profile={profile} onSelect={setSelectedEntry}/>
        )}

        {view==="photos"&&<PhotoSection userId={userId}/>}
      </div>
    </div>
  );
}

function SportProgress({ type, history, profile, onSelect }) {
  const sp2 = SPORTS.find(s=>s.id===type);
  const c = SPORT_COLOR[type];
  const data = profile?.sportData?.[type];
  const [selEx, setSelEx] = useState(null);

  const trackedExercises = type==="strength" ? getTrackedExercises(history) : [];
  const exId = selEx || trackedExercises[0];
  const loadData = type==="strength" && exId ? buildLoadData(history, exId) : [];
  const paceData = type==="running" ? buildRunPaceData(history) : [];
  const exName = exId ? (Object.values(EXERCISES).flat().find(e=>e.id===exId)?.name||exId) : "";

  const paceTrend = paceData.length>=2 ? paceData[paceData.length-1].v - paceData[0].v : 0;
  const loadTrend = loadData.length>=2 ? loadData[loadData.length-1].v - loadData[0].v : 0;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      {/* Profile card */}
      <Card accent={c}>
        <div style={{ fontSize:16, fontWeight:700, color:D.text, marginBottom:12 }}>{sp2?.emoji} {sp2?.label} — Seu Perfil</div>
        {type==="strength"&&<div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          {[{l:"Objetivo",v:data?.goal==="hypertrophy"?"Hipertrofia":data?.goal==="strength"?"Força":data?.goal==="weight_loss"?"Emagrecimento":"Recomposição"},{l:"Experiência",v:data?.experience==="beginner"?"Iniciante":data?.experience==="intermediate"?"Intermediário":"Avançado"},{l:"Supino inicial",v:`${data?.currentBench||"?"}kg`},{l:"Agach. inicial",v:`${data?.currentSquat||"?"}kg`}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"10px 12px" }}><div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>{m.l}</div><div style={{ fontSize:14, fontWeight:700, color:c, marginTop:4 }}>{m.v}</div></div>
          ))}
        </div>}
        {type==="running"&&<div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          {[{l:"Pace inicial 5km",v:`${data?.currentPace5k||"?"}/km`},{l:"Objetivo",v:data?.goalDistance==="condicionamento"?"Condicionamento":`${data?.goalDistance||"?"}km`},{l:"Maior distância",v:`${data?.longestRun||"?"}km`},{l:"Km/sem inicial",v:`${data?.weeklyKm||"?"}km`}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"10px 12px" }}><div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>{m.l}</div><div style={{ fontSize:14, fontWeight:700, color:c, marginTop:4 }}>{m.v}</div></div>
          ))}
        </div>}
        {type==="bike"&&<div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          {[{l:"FTP inicial",v:data?.currentFTP?`${data.currentFTP}W`:"Não testado"},{l:"Objetivo",v:data?.goal||"—"},{l:"Tipo de bike",v:data?.bikeType||"—"},{l:"Km/sem inicial",v:`${data?.weeklyKm||"?"}km`}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"10px 12px" }}><div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>{m.l}</div><div style={{ fontSize:14, fontWeight:700, color:c, marginTop:4 }}>{m.v}</div></div>
          ))}
        </div>}
        {type==="swim"&&<div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          {[{l:"Pace/100m inicial",v:`${data?.currentPace100||"?"}`},{l:"Estilo",v:data?.style||"—"},{l:"Objetivo",v:data?.goal||"—"},{l:"Volume inicial",v:`${data?.weeklyVolume||"?"}m`}].map(m=>(
            <div key={m.l} style={{ background:D.bg2, borderRadius:10, padding:"10px 12px" }}><div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", letterSpacing:".06em" }}>{m.l}</div><div style={{ fontSize:14, fontWeight:700, color:c, marginTop:4 }}>{m.v}</div></div>
          ))}
        </div>}
      </Card>

      {/* Pace trend chart — running */}
      {type==="running" && paceData.length>=2 && (
        <Card>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
            <Lbl>📈 Evolução de Pace</Lbl>
            <span style={{ fontSize:12, fontWeight:700, color:paceTrend<0?B.emerald:B.red }}>
              {paceTrend<0?"▼":"▲"} {Math.abs(paceTrend)}s/km
            </span>
          </div>
          <div style={{ fontSize:11, color:D.muted, marginBottom:10 }}>
            {paceTrend<0?`Pace melhorou ${Math.abs(paceTrend)}s/km desde o início — evolução real.`:`Pace está ${Math.abs(paceTrend)}s/km mais lento que o início. Avalie o volume e intensidade.`}
          </div>
          <ResponsiveContainer width="100%" height={180}>
            <LineChart data={paceData} margin={{top:4,right:4,left:0,bottom:4}}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="d" tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false}/>
              <YAxis domain={["auto","auto"]} tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false} tickFormatter={secToPace} width={38}/>
              <Tooltip content={({active,payload,label})=>{
                if(!active||!payload?.length)return null;
                return <div style={{background:D.bg2,border:`1px solid ${D.border}`,borderRadius:8,padding:"8px 12px",fontSize:12}}>
                  <div style={{color:D.muted,marginBottom:3}}>{label}</div>
                  <div style={{color:c,fontWeight:700}}>{secToPace(payload[0]?.value)}/km</div>
                  <div style={{color:D.muted,fontSize:10}}>{payload[0]?.payload?.dist}km</div>
                </div>;
              }}/>
              <Line type="monotone" dataKey="v" stroke={c} strokeWidth={2.5} dot={{r:3,fill:c}} activeDot={{r:5}}/>
            </LineChart>
          </ResponsiveContainer>
        </Card>
      )}

      {/* Load progression chart — strength */}
      {type==="strength" && trackedExercises.length>0 && (
        <Card>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
            <Lbl>📈 Progressão de Carga</Lbl>
            {loadTrend!==0&&<span style={{ fontSize:12, fontWeight:700, color:loadTrend>0?B.emerald:B.red }}>{loadTrend>0?`▲ +${loadTrend}kg`:`▼ ${loadTrend}kg`}</span>}
          </div>
          {/* Exercise selector */}
          <div style={{ display:"flex", gap:5, flexWrap:"wrap", marginBottom:12 }}>
            {trackedExercises.map(id=>{
              const name = Object.values(EXERCISES).flat().find(e=>e.id===id)?.name||id;
              const sel = (selEx||trackedExercises[0])===id;
              return <button key={id} onClick={()=>setSelEx(id)} style={{ padding:"5px 10px", borderRadius:16, cursor:"pointer", border:`1.5px solid ${sel?c:D.border}`, background:sel?`${c}20`:D.card, color:sel?c:D.muted, fontSize:10, fontWeight:sel?700:400, fontFamily:"'Inter',sans-serif" }}>{name.split(" ").slice(0,2).join(" ")}</button>;
            })}
          </div>
          {loadData.length<2 ? (
            <div style={{textAlign:"center",padding:"20px 0",color:D.muted,fontSize:13}}>Registre mais sessões de {exName.split(" ").slice(0,2).join(" ")} para ver a evolução.</div>
          ) : (
            <ResponsiveContainer width="100%" height={160}>
              <LineChart data={loadData} margin={{top:4,right:4,left:0,bottom:4}}>
                <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
                <XAxis dataKey="d" tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false} unit="kg"/>
                <Tooltip content={({active,payload,label})=>{
                  if(!active||!payload?.length)return null;
                  return <div style={{background:D.bg2,border:`1px solid ${D.border}`,borderRadius:8,padding:"8px 12px",fontSize:12}}>
                    <div style={{color:D.muted,marginBottom:3}}>{label}</div>
                    <div style={{color:c,fontWeight:700}}>{payload[0]?.value}kg × {payload[0]?.payload?.reps} reps</div>
                  </div>;
                }}/>
                <Line type="monotone" dataKey="v" stroke={c} strokeWidth={2.5} dot={{r:3,fill:c}} activeDot={{r:5}}/>
              </LineChart>
            </ResponsiveContainer>
          )}
        </Card>
      )}

      {/* Session history */}
      <Card>
        <Lbl>Histórico de sessões ({history.length} registros)</Lbl>
        {history.length===0?(
          <div style={{ textAlign:"center", padding:"20px 0", color:D.muted, fontSize:13 }}>Registre seus treinos para ver o histórico aqui.</div>
        ):(
          history.slice(-10).reverse().map((w,i)=>(
            <button key={i} onClick={()=>onSelect&&onSelect(w)} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"9px 0", borderBottom:i<Math.min(history.length,10)-1?`1px solid ${D.bg2}`:"none", background:"transparent", border:"none", cursor:"pointer", width:"100%", textAlign:"left", fontFamily:"'Inter',sans-serif" }}>
              <div>
                <div style={{ fontSize:13, fontWeight:600, color:D.text }}>{w.plan?.focus||"Treino"}</div>
                <div style={{ fontSize:11, color:D.muted }}>{w.date}</div>
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <div style={{ fontSize:12, color:c, fontWeight:700 }}>
                  {type==="running"&&w.log?.pace?`${w.log.pace}/km`:
                   type==="strength"&&w.log?.logs?`${Object.values(w.log.logs).reduce((s,a)=>s+a.length,0)} séries`:
                   type==="bike"&&w.log?.distance?`${w.log.distance}km`:
                   type==="swim"&&w.log?.volume?`${w.log.volume}m`:
                   w.log?.feel?`${w.log.feel}/10`:""}
                </div>
                <span style={{ color:D.muted, fontSize:13 }}>›</span>
              </div>
            </button>
          ))
        )}
      </Card>
    </div>
  );
}

/* ══ PHOTO SECTION (unchanged logic) ════════════════════════════ */
function PhotoSection({ userId }) {
  const [assessments,setAssessments]=useState([]);
  const [newPhotos,setNewPhotos]=useState({});
  const [addMode,setAddMode]=useState(false);
  const [comparing,setComparing]=useState(null);
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);

  useEffect(()=>{ if(userId) loadAssessments(); },[userId]);

  async function loadAssessments() {
    setLoading(true);
    try { setAssessments(await DB.getAssessments(userId)); } catch {}
    setLoading(false);
  }
  async function handleUpload(angle,file){const raw=await readFile(file);const compressed=await compressImage(raw,480);setNewPhotos(p=>({...p,[angle]:compressed}));}
  async function saveAssessment(){
    if(Object.keys(newPhotos).length===0||!userId)return;setSaving(true);
    const obj={date:new Date().toLocaleDateString("pt-BR"),isoDate:new Date().toISOString().split("T")[0],photos:newPhotos};
    try{ await DB.saveAssessment(userId, obj); setAssessments(p=>[{...obj,id:`local_${Date.now()}`},...p]); setNewPhotos({}); setAddMode(false); }catch{}
    setSaving(false);
  }

  if(comparing){
    const cur=assessments[0]; const comp=assessments.find(a=>a.id===comparing);
    return (
      <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <button onClick={()=>setComparing(null)} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
          <div style={{ fontSize:16, fontWeight:800, color:D.text }}>Comparação</div>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:4 }}>
          <div style={{ textAlign:"center", fontSize:12, fontWeight:700, color:B.blue }}>Atual · {cur?.date}</div>
          <div style={{ textAlign:"center", fontSize:12, fontWeight:700, color:B.purple }}>Anterior · {comp?.date}</div>
        </div>
        {ANGLES.map(a=>{
          const cu=cur?.photos?.[a.key]; const ol=comp?.photos?.[a.key];
          if(!cu&&!ol)return null;
          return (
            <div key={a.key}>
              <div style={{ fontSize:11, color:D.muted, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:8 }}>{a.icon} {a.label}</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                {[cu,ol].map((ph,i)=>(
                  <div key={i} style={{ borderRadius:14, overflow:"hidden", background:D.card, border:`1.5px solid ${i===0?B.blue:B.purple}40`, aspectRatio:"3/4" }}>
                    {ph?<img src={ph} alt={a.label} style={{ width:"100%", height:"100%", objectFit:"cover", objectPosition:"top" }}/>:<div style={{ width:"100%", height:"100%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:24, color:D.muted }}>—</div>}
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  if(addMode) return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={()=>{setAddMode(false);setNewPhotos({});}} style={{ background:"transparent", border:"none", color:D.muted, cursor:"pointer", fontSize:20, padding:0 }}>←</button>
        <div style={{ fontSize:16, fontWeight:800, color:D.text }}>Nova Avaliação Física</div>
      </div>
      <div style={{ fontSize:13, color:D.muted, background:D.card, borderRadius:12, padding:"12px 16px", lineHeight:1.6 }}>📌 Mesma luz, mesmo horário, mesma postura para comparações precisas.</div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        {ANGLES.map(a=>(
          <label key={a.key} htmlFor={`ph_${a.key}`} style={{ cursor:"pointer" }}>
            <div style={{ borderRadius:14, overflow:"hidden", background:D.card, border:`2px dashed ${newPhotos[a.key]?B.blue:D.border}`, aspectRatio:"3/4", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", position:"relative" }}>
              {newPhotos[a.key]?(
                <><img src={newPhotos[a.key]} alt={a.label} style={{ width:"100%", height:"100%", objectFit:"cover", objectPosition:"top", position:"absolute", inset:0 }}/><div style={{ position:"absolute", bottom:8, left:"50%", transform:"translateX(-50%)", background:`${B.blue}dd`, color:"#fff", fontSize:10, fontWeight:700, padding:"4px 10px", borderRadius:20, whiteSpace:"nowrap" }}>✓ {a.label}</div></>
              ):(
                <><div style={{ fontSize:28, marginBottom:8 }}>{a.icon}</div><div style={{ fontSize:12, color:D.sub, fontWeight:600 }}>{a.label}</div><div style={{ fontSize:10, color:D.muted, marginTop:3 }}>Toque para foto</div></>
              )}
            </div>
            <input id={`ph_${a.key}`} type="file" accept="image/*" capture="camera" onChange={e=>e.target.files[0]&&handleUpload(a.key,e.target.files[0])} style={{ display:"none" }}/>
          </label>
        ))}
      </div>
      {Object.keys(newPhotos).length>0&&<button onClick={saveAssessment} disabled={saving} style={{ width:"100%", padding:"14px 0", borderRadius:12, border:"none", background:saving?"#334155":B.blue, color:"#fff", fontWeight:800, fontSize:15, cursor:saving?"wait":"pointer", fontFamily:"'Inter',sans-serif" }}>{saving?"⏳ Salvando...":` Salvar (${Object.keys(newPhotos).length}/4 fotos)`}</button>}
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        <div style={{ fontSize:16, fontWeight:800, color:D.text }}>📸 Avaliações Físicas</div>
        <button onClick={()=>setAddMode(true)} style={{ background:B.blue, color:"#fff", border:"none", borderRadius:10, padding:"9px 16px", fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>+ Nova</button>
      </div>
      {loading?<div style={{ textAlign:"center", padding:"30px 0", color:D.muted }}>Carregando...</div>:
      assessments.length===0?(
        <div style={{ textAlign:"center", padding:"40px 20px" }}>
          <div style={{ fontSize:52, marginBottom:16 }}>📸</div>
          <div style={{ fontSize:16, fontWeight:700, color:D.text, marginBottom:8 }}>Nenhuma avaliação ainda</div>
          <div style={{ fontSize:13, color:D.muted, lineHeight:1.7, marginBottom:20 }}>Fotos de frente, costas e laterais para acompanhar sua transformação.</div>
          <button onClick={()=>setAddMode(true)} style={{ background:B.blue, color:"#fff", border:"none", borderRadius:12, padding:"13px 24px", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:14 }}>📸 Fazer Primeira Avaliação</button>
        </div>
      ):(
        <>
          {assessments.length>1&&(
            <Card style={{ background:`${B.indigo}15`, border:`1px solid ${B.indigo}30` }}>
              <div style={{ fontSize:11, color:B.indigo, fontWeight:700, textTransform:"uppercase", letterSpacing:".08em", marginBottom:8 }}>Comparação Rápida</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
                {ANGLES.slice(0,2).map(a=>{
                  const cu=assessments[0]?.photos?.[a.key]; const ol=assessments[1]?.photos?.[a.key];
                  return (
                    <div key={a.key}>
                      <div style={{ fontSize:10, color:D.muted, textTransform:"uppercase", marginBottom:5 }}>{a.icon} {a.label}</div>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4 }}>
                        {[{ph:cu,c:B.blue},{ph:ol,c:B.purple}].map((x,i)=>(
                          <div key={i} style={{ borderRadius:10, overflow:"hidden", background:D.card, border:`1.5px solid ${x.c}40`, aspectRatio:"3/4" }}>
                            {x.ph?<img src={x.ph} alt={a.label} style={{ width:"100%", height:"100%", objectFit:"cover", objectPosition:"top" }}/>:<div style={{ width:"100%", height:"100%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16, color:D.muted }}>—</div>}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
              <button onClick={()=>setComparing(assessments[1]?.id)} style={{ width:"100%", padding:"10px 0", borderRadius:10, border:"none", background:B.indigo, color:"#fff", fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>Ver comparação completa →</button>
            </Card>
          )}
          {assessments.map((a,ai)=>(
            <Card key={a.id}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                <div><div style={{ fontSize:14, fontWeight:700, color:D.text }}>{a.date}</div><div style={{ fontSize:11, color:D.muted }}>{Object.keys(a.photos||{}).length} fotos{ai===0?" · Mais recente":""}</div></div>
                {ai>0&&<button onClick={()=>setComparing(a.id)} style={{ background:`${B.indigo}20`, color:B.indigo, border:`1px solid ${B.indigo}40`, borderRadius:8, padding:"6px 12px", fontWeight:700, fontSize:11, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Comparar</button>}
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:7 }}>
                {ANGLES.map(ang=>{const ph=a.photos?.[ang.key]; return (
                  <div key={ang.key} style={{ display:"flex", flexDirection:"column", gap:4 }}>
                    <div style={{ borderRadius:9, overflow:"hidden", background:D.bg2, aspectRatio:"3/4" }}>
                      {ph?<img src={ph} alt={ang.label} style={{ width:"100%", height:"100%", objectFit:"cover", objectPosition:"top" }}/>:<div style={{ width:"100%", height:"100%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, color:D.muted }}>{ang.icon}</div>}
                    </div>
                    <div style={{ fontSize:9, color:D.muted, textAlign:"center", fontWeight:600 }}>{ang.label}</div>
                  </div>
                );})}
              </div>
            </Card>
          ))}
        </>
      )}
    </div>
  );
}

/* ══ BOTTOM NAV ══════════════════════════════════════════════════ */
function BottomNav({ active, onChange }) {
  const tabs=[{id:"home",icon:"🏠",l:"Início"},{id:"treino",icon:"🏋️",l:"Treino"},{id:"coach",icon:"🤖",l:"Coach"},{id:"progresso",icon:"📊",l:"Progresso"},{id:"perfil",icon:"👤",l:"Perfil"}];
  return (
    <div style={{ position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)", width:"100%", maxWidth:480, background:D.bar, borderTop:`1px solid ${D.card}`, display:"flex", zIndex:50, boxSizing:"border-box", paddingBottom:"env(safe-area-inset-bottom,0px)" }}>
      {tabs.map(t=>(
        <button key={t.id} onClick={()=>onChange(t.id)} style={{ flex:1, padding:"10px 0 8px", display:"flex", flexDirection:"column", alignItems:"center", gap:3, background:"transparent", border:"none", cursor:"pointer", fontFamily:"'Inter',sans-serif", position:"relative" }}>
          {active===t.id&&<div style={{ position:"absolute", top:0, left:"50%", transform:"translateX(-50%)", width:28, height:3, borderRadius:"0 0 3px 3px", background:`linear-gradient(90deg,${B.blue},${B.indigo})` }}/>}
          <span style={{ fontSize:20, filter:active===t.id?"none":"grayscale(.6) opacity(.55)", transition:"filter .2s" }}>{t.icon}</span>
          <span style={{ fontSize:9, fontWeight:active===t.id?700:400, color:active===t.id?B.blue:D.muted, letterSpacing:".04em" }}>{t.l}</span>
        </button>
      ))}
    </div>
  );
}

/* ══ PROFILE TAB ════════════════════════════════════════════════ */
/* ══ BODY WEIGHT SECTION ════════════════════════════════════════ */
function BodyWeightSection({ userId }) {
  const [weights, setWeights] = useState([]);
  const [todayVal, setTodayVal] = useState("");
  const [saving, setSaving] = useState(false);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => { if(userId) loadWeights(); }, [userId]);

  async function loadWeights() {
    try {
      const items = await DB.getWeights(userId);
      setWeights(items);
      const today = new Date().toISOString().split("T")[0];
      const existing = items.find(w=>w.iso===today);
      if (existing) setTodayVal(String(existing.v));
    } catch {}
    setLoaded(true);
  }

  async function saveWeight() {
    const v = parseFloat(todayVal);
    if (!v || v < 20 || v > 300 || !userId) return;
    setSaving(true);
    const iso = new Date().toISOString().split("T")[0];
    const entry = { iso, date: new Date().toLocaleDateString("pt-BR"), v };
    try {
      await DB.saveWeight(userId, entry);
      setWeights(prev => [...prev.filter(w=>w.iso!==iso), entry].sort((a,b)=>a.iso.localeCompare(b.iso)));
    } catch {}
    setSaving(false);
  }

  const chartData = weights.slice(-30).map(w=>({ d: w.date?.split("/").slice(0,2).join("/"), v: w.v }));
  const first = weights[0]?.v; const last = weights[weights.length-1]?.v;
  const diff = first&&last ? (last-first).toFixed(1) : null;
  const trend = diff!==null ? parseFloat(diff) : null;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card>
        <Lbl>⚖️ Registrar Peso Hoje</Lbl>
        <div style={{ display:"flex", gap:10, alignItems:"center" }}>
          <input type="number" step="0.1" value={todayVal} onChange={e=>setTodayVal(e.target.value)}
            placeholder="Ex: 80.5" style={{ flex:1, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:12, color:D.text, padding:"11px 14px", fontSize:16, fontFamily:"'Inter',sans-serif", boxSizing:"border-box" }}/>
          <span style={{ fontSize:14, color:D.muted, fontWeight:600 }}>kg</span>
          <button onClick={saveWeight} disabled={saving||!todayVal} style={{ padding:"11px 18px", borderRadius:12, border:"none", background:todayVal?B.blue:"#334155", color:todayVal?"#fff":"#475569", fontWeight:700, cursor:todayVal?"pointer":"not-allowed", fontFamily:"'Inter',sans-serif", fontSize:14 }}>
            {saving?"...":"Salvar"}
          </button>
        </div>
      </Card>

      {chartData.length>=2 && (
        <Card>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
            <Lbl>Evolução do Peso</Lbl>
            {trend!==null && (
              <span style={{ fontSize:12, fontWeight:700, color:trend<0?B.emerald:trend>0?B.orange:D.muted }}>
                {trend<0?`▼ ${Math.abs(trend)}kg`:trend>0?`▲ +${trend}kg`:"Estável"}
              </span>
            )}
          </div>
          <ResponsiveContainer width="100%" height={160}>
            <AreaChart data={chartData} margin={{top:4,right:4,left:0,bottom:4}}>
              <defs><linearGradient id="gw" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.blue} stopOpacity={.25}/><stop offset="95%" stopColor={B.blue} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="d" tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fontSize:10,fill:D.muted}} axisLine={false} tickLine={false} domain={["auto","auto"]} unit="kg" width={42}/>
              <Tooltip content={<TipBox unit="kg"/>}/>
              <Area type="monotone" dataKey="v" stroke={B.blue} fill="url(#gw)" strokeWidth={2.5} dot={{r:2,fill:B.blue}}/>
            </AreaChart>
          </ResponsiveContainer>
        </Card>
      )}

      {loaded && weights.length===0 && (
        <div style={{ textAlign:"center", padding:"30px 20px", color:D.muted }}>
          <div style={{ fontSize:32, marginBottom:10 }}>⚖️</div>
          <div style={{ fontSize:13 }}>Registre seu peso diariamente para acompanhar a evolução com precisão.</div>
        </div>
      )}
    </div>
  );
}

/* ══ PROFILE TAB ════════════════════════════════════════════════ */
function ProfileTab({ profile, onReset, onRegenerate, workoutHistory, userId, onSignOut }) {
  const [confirm, setConfirm] = useState(false);
  const [regenState, setRegenState] = useState("idle"); // idle | confirming | loading | done
  const [notifPerm, setNotifPerm] = useState(typeof Notification!=="undefined" ? Notification.permission : "default");
  const [notifTime, setNotifTime] = useState("07:00");
  const [activeSection, setActiveSection] = useState("profile"); // profile | weight | export

  async function requestNotifications() {
    if (typeof Notification === "undefined") return;
    const perm = await Notification.requestPermission();
    setNotifPerm(perm);
    if (perm === "granted") {
      // Schedule a local notification check
      try { await window.storage.set("notif_time", JSON.stringify({ time: notifTime })); } catch {}
      new Notification("Athlete AI 🏆", {
        body: `Notificações ativadas para ${notifTime}. Até o próximo treino!`,
        icon: "/icon-192.png",
      });
    }
  }

  function exportXLSX() {
    if (!workoutHistory?.length) return;
    const rows = workoutHistory.map(w => ({
      Data: w.date||"",
      Esporte: SPORTS.find(s=>s.id===w.type)?.label||w.type||"",
      Sessão: w.plan?.focus||"",
      "Distância/Volume": w.type==="running"?`${w.log?.distance||""}km`:w.type==="swim"?`${w.log?.volume||""}m`:w.type==="bike"?`${w.log?.distance||""}km`:"",
      Pace: w.log?.pace||w.log?.pace100||"",
      "FC média": w.log?.hr||"",
      Potência: w.log?.avgPower||"",
      Cadência: w.log?.cadence||"",
      Sensação: w.log?.feel||"",
      Notas: w.log?.notes||"",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Histórico");
    XLSX.writeFile(wb, `AthleteAI_Historico_${new Date().toLocaleDateString("pt-BR").replace(/\//g,"-")}.xlsx`);
  }

  const SECTIONS = [{id:"profile",l:"👤 Perfil"},{id:"weight",l:"⚖️ Peso"},{id:"export",l:"📤 Dados"}];

  return (
    <div style={{ padding:20, display:"flex", flexDirection:"column", gap:0, paddingBottom:20 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div style={{ fontSize:22, fontWeight:900, color:D.text, letterSpacing:"-.5px" }}>Perfil</div>
        <button onClick={onSignOut} style={{ background:"transparent", border:`1px solid ${D.border}`, color:D.muted, borderRadius:10, padding:"7px 14px", fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>
          Sair →
        </button>
      </div>

      {/* Section tabs */}
      <div style={{ display:"flex", gap:0, marginBottom:18, background:D.card, borderRadius:12, padding:4 }}>
        {SECTIONS.map(s=>(
          <button key={s.id} onClick={()=>setActiveSection(s.id)} style={{ flex:1, padding:"9px 0", borderRadius:9, background:activeSection===s.id?D.bg2:"transparent", color:activeSection===s.id?D.text:D.muted, fontWeight:activeSection===s.id?700:400, fontSize:12, cursor:"pointer", border:"none", fontFamily:"'Inter',sans-serif", transition:"all .15s" }}>{s.l}</button>
        ))}
      </div>

      {/* PROFILE SECTION */}
      {activeSection==="profile"&&(
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <Card>
            <div style={{ display:"flex", gap:14, alignItems:"center", marginBottom:16 }}>
              <div style={{ width:56, height:56, borderRadius:16, background:`linear-gradient(135deg,${B.blue},${B.indigo})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26 }}>🏆</div>
              <div>
                <div style={{ fontSize:20, fontWeight:800, color:D.text }}>{profile?.basic?.name||"Atleta"}</div>
                <div style={{ fontSize:13, color:D.muted }}>{profile?.basic?.age} anos · {profile?.basic?.weight}kg · {profile?.basic?.height}cm</div>
              </div>
            </div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>{(profile?.selectedSports||[]).map(s=><SportBadge key={s} type={s} sm/>)}</div>
          </Card>

          {(profile?.selectedSports||[]).map(s=>{
            const sp2=SPORTS.find(x=>x.id===s); const d=profile?.sportData?.[s];
            return (
              <Card key={s} accent={SPORT_COLOR[s]}>
                <div style={{ fontWeight:700, color:SPORT_COLOR[s], marginBottom:8 }}>{sp2?.emoji} {sp2?.label}</div>
                <div style={{ fontSize:13, color:D.sub, lineHeight:1.8 }}>
                  {s==="strength"&&`Objetivo: ${d?.goal==="hypertrophy"?"Hipertrofia":d?.goal==="strength"?"Força":d?.goal==="weight_loss"?"Emagrecimento":"Recomposição"} · ${d?.daysPerWeek||"?"}×/sem · ${d?.experience==="beginner"?"Iniciante":d?.experience==="intermediate"?"Intermediário":"Avançado"}`}
                  {s==="running"&&`Pace 5km: ${d?.currentPace5k||"?"}/km · Meta: ${d?.goalDistance||"?"}km · ${d?.weeklyKm||"?"}km/sem`}
                  {s==="bike"&&`FTP: ${d?.currentFTP||"Não testado"} · ${d?.bikeType||"?"} · ${d?.weeklyKm||"?"}km/sem`}
                  {s==="swim"&&`Pace: ${d?.currentPace100||"?"}/100m · ${d?.style||"?"} · ${d?.weeklyVolume||"?"}m/sem`}
                </div>
              </Card>
            );
          })}

          {/* Availability summary */}
          <Card style={{ background:`${B.indigo}10`, border:`1px solid ${B.indigo}25` }}>
            <Lbl>📅 Rotina configurada</Lbl>
            <div style={{ fontSize:13, color:D.sub, marginBottom:8 }}>Disponível: {profile?.availability?.availableDays?.join(", ")||"—"}</div>
            {(profile?.selectedSports||[]).filter(s=>(profile?.availability?.preferredDays?.[s]||[]).length>0).map(s=>{
              const sp2=SPORTS.find(x=>x.id===s);
              return <div key={s} style={{ fontSize:12, color:D.muted, marginBottom:3 }}>{sp2?.emoji} {sp2?.label}: {profile.availability.preferredDays[s].join(", ")}</div>;
            })}
          </Card>

          {/* AI Plan */}
          <Card style={{ background:`${B.indigo}10`, border:`1px solid ${B.indigo}25` }}>
            <Lbl>🤖 Plano atual</Lbl>
            <div style={{ fontSize:14, fontWeight:700, color:D.text, marginBottom:4 }}>{profile?.plan?.planName||"—"}</div>
            <div style={{ fontSize:13, color:D.muted, lineHeight:1.6, marginBottom:14 }}>{profile?.plan?.description||"—"}</div>
            {regenState==="idle"&&<button onClick={()=>setRegenState("confirming")} style={{ width:"100%", padding:"11px 0", borderRadius:10, border:`1.5px solid ${B.indigo}50`, background:`${B.indigo}15`, color:B.indigo, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>🔄 Regenerar plano com IA</button>}
            {regenState==="confirming"&&(
              <div>
                <div style={{ fontSize:13, color:D.sub, marginBottom:12 }}>A IA vai reconstruir o plano usando seu histórico de {workoutHistory?.length||0} treinos. Isso substitui o plano atual.</div>
                <div style={{ display:"flex", gap:10 }}>
                  <button onClick={()=>setRegenState("idle")} style={{ flex:1, padding:"10px 0", borderRadius:10, border:`1.5px solid ${D.border}`, background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Cancelar</button>
                  <button onClick={async()=>{ setRegenState("loading"); await onRegenerate(); setRegenState("done"); setTimeout(()=>setRegenState("idle"),3000); }} style={{ flex:2, padding:"10px 0", borderRadius:10, border:"none", background:B.indigo, color:"#fff", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Gerar novo plano</button>
                </div>
              </div>
            )}
            {regenState==="loading"&&<div style={{ textAlign:"center", padding:"10px 0", color:B.indigo, fontSize:13, fontWeight:600 }}>🤖 Gerando novo plano...</div>}
            {regenState==="done"&&<div style={{ textAlign:"center", padding:"10px 0", color:B.emerald, fontSize:13, fontWeight:600 }}>✓ Plano regenerado!</div>}
          </Card>

          {/* Notifications */}
          <Card>
            <Lbl>🔔 Lembretes de Treino</Lbl>
            {notifPerm==="granted" ? (
              <div>
                <div style={{ fontSize:13, color:B.emerald, fontWeight:600, marginBottom:10 }}>✓ Notificações ativas</div>
                <div style={{ display:"flex", gap:10, alignItems:"center" }}>
                  <input type="time" value={notifTime} onChange={e=>setNotifTime(e.target.value)} style={{ flex:1, background:D.bg2, border:`1.5px solid ${D.border}`, borderRadius:10, color:D.text, padding:"9px 12px", fontSize:14, fontFamily:"'Inter',sans-serif" }}/>
                  <button onClick={requestNotifications} style={{ padding:"9px 14px", borderRadius:10, border:"none", background:B.blue, color:"#fff", fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>Salvar</button>
                </div>
              </div>
            ) : notifPerm==="denied" ? (
              <div style={{ fontSize:13, color:B.red }}>Notificações bloqueadas. Ative nas configurações do navegador.</div>
            ) : (
              <div>
                <div style={{ fontSize:13, color:D.muted, marginBottom:12 }}>Receba um lembrete no horário do seu treino.</div>
                <button onClick={requestNotifications} style={{ width:"100%", padding:"11px 0", borderRadius:10, border:`1.5px solid ${D.border}`, background:"transparent", color:D.sub, fontWeight:600, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:13 }}>Ativar lembretes</button>
              </div>
            )}
          </Card>

          {/* Reset */}
          {!confirm ? (
            <button onClick={()=>setConfirm(true)} style={{ padding:"12px 0", borderRadius:12, border:`1.5px solid ${B.red}40`, background:`${B.red}10`, color:B.red, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif", fontSize:14 }}>🔄 Reconfigurar perfil</button>
          ) : (
            <div style={{ background:`${B.red}15`, borderRadius:12, padding:16 }}>
              <div style={{ fontSize:14, fontWeight:700, color:D.text, marginBottom:8 }}>Tem certeza? O histórico de treinos será mantido.</div>
              <div style={{ display:"flex", gap:10 }}>
                <button onClick={()=>setConfirm(false)} style={{ flex:1, padding:"11px 0", borderRadius:10, border:`1.5px solid ${D.border}`, background:"transparent", color:D.muted, fontWeight:700, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Cancelar</button>
                <button onClick={onReset} style={{ flex:1, padding:"11px 0", borderRadius:10, border:"none", background:B.red, color:"#fff", fontWeight:800, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>Confirmar</button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* WEIGHT SECTION */}
      {activeSection==="weight"&&<BodyWeightSection userId={userId}/>}

      {/* EXPORT / DATA SECTION */}
      {activeSection==="export"&&(
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <Card>
            <Lbl>📊 Exportar histórico</Lbl>
            <div style={{ fontSize:13, color:D.muted, marginBottom:14, lineHeight:1.7 }}>
              Gera uma planilha Excel com todos os seus {workoutHistory?.length||0} treinos registrados — data, esporte, sessão, métricas e notas.
            </div>
            <button onClick={exportXLSX} disabled={!workoutHistory?.length} style={{
              width:"100%", padding:"13px 0", borderRadius:12, border:"none",
              background:workoutHistory?.length?B.emerald:"#334155",
              color:workoutHistory?.length?"#0A1628":"#475569",
              fontWeight:800, fontSize:14, cursor:workoutHistory?.length?"pointer":"not-allowed", fontFamily:"'Inter',sans-serif",
            }}>⬇ Baixar AthleteAI_Historico.xlsx</button>
          </Card>

          <Card>
            <Lbl>📱 Compartilhar app</Lbl>
            <div style={{ fontSize:13, color:D.muted, marginBottom:14 }}>Indique o Athlete AI para outros atletas que também querem feedback real.</div>
            <button onClick={()=>{ try { navigator.share({ title:"Athlete AI", text:"Coach esportivo com IA — honesto e baseado em dados.", url:window.location.href }); } catch { navigator.clipboard?.writeText(window.location.href); } }}
              style={{ width:"100%", padding:"13px 0", borderRadius:12, border:`1.5px solid ${D.border}`, background:"transparent", color:D.sub, fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>
              📤 Compartilhar link do app
            </button>
          </Card>

          <Card style={{ background:`${D.bg2}`, border:`1px solid ${D.border}` }}>
            <Lbl>ℹ️ Sobre</Lbl>
            <div style={{ fontSize:12, color:D.muted, lineHeight:1.8 }}>
              Athlete AI v1.0<br/>
              {workoutHistory?.length||0} treinos registrados<br/>
              Plano: {profile?.plan?.planName||"—"}<br/>
              Esportes: {(profile?.selectedSports||[]).map(s=>SPORTS.find(x=>x.id===s)?.label).join(", ")||"—"}
            </div>
          </Card>
        </div>
      )}
    </div>
  );
}

/* ══ ROOT APP ════════════════════════════════════════════════════ */
export default function AthleteAI() {
  const [screen, setScreen] = useState("splash"); // splash | auth | onboarding | main
  const [tab, setTab] = useState("home");
  const [userId, setUserId] = useState(null);
  const [profile, setProfile] = useState(null);
  const [todayPlan, setTodayPlan] = useState(null);
  const [workoutHistory, setWorkoutHistory] = useState([]);
  const [todayCompleted, setTodayCompleted] = useState(null);
  const [chat, setChat] = useState([]);
  const [chatIn, setChatIn] = useState("");
  const [aiLoad, setAiLoad] = useState(false);
  const chatRef = useRef(null);
  const TODAY_STR = new Date().toLocaleDateString("pt-BR");

  // Check auth on mount
  useEffect(() => {
    async function init() {
      try {
        const user = await DB.getSession();
        if (user) {
          setUserId(user.id);
          await loadUserData(user.id);
        } else {
          setTimeout(() => setScreen("auth"), 2200);
        }
      } catch {
        setTimeout(() => setScreen("auth"), 2200);
      }
    }
    init();
  }, []);

  async function loadUserData(uid) {
    try {
      const [p, workouts] = await Promise.all([DB.getProfile(uid), DB.getWorkouts(uid)]);
      setWorkoutHistory(workouts || []);
      // Check if today already completed
      const tc = (workouts||[]).filter(w=>w.date===TODAY_STR).slice(-1)[0];
      if (tc) setTodayCompleted(tc);
      if (p) {
        setProfile(p);
        setTodayPlanFromProfile(p);
        setScreen("main");
      } else {
        setScreen("onboarding");
      }
    } catch {
      setScreen("onboarding");
    }
  }

  function setTodayPlanFromProfile(p) {
    const todayN = dayName[new Date().getDay()];
    const tp = (p?.plan?.weeklySchedule||[]).find(x=>x.day===todayN) || { type:"rest", day:todayN };
    setTodayPlan(tp);
  }

  async function handleAuth(user) {
    setUserId(user.id);
    setScreen("splash"); // brief loading
    await loadUserData(user.id);
  }

  async function saveProfile(p) {
    setProfile(p);
    setTodayPlanFromProfile(p);
    try { await DB.saveProfile(userId, p); } catch {}
    setScreen("main");
  }

  async function logWorkout(entry) {
    const fullEntry = { ...entry, date: entry.date || TODAY_STR };
    const updated = [...workoutHistory, fullEntry];
    setWorkoutHistory(updated);
    setTodayCompleted(fullEntry);
    try { await DB.addWorkout(userId, fullEntry); } catch {}
    setTab("home");
  }

  async function deleteWorkout(entry) {
    const updated = workoutHistory.filter(w =>
      !(w.date===entry.date && w.type===entry.type && w.plan?.focus===entry.plan?.focus)
    );
    setWorkoutHistory(updated);
    setTodayCompleted(prev => (prev?.date===entry.date&&prev?.type===entry.type) ? null : prev);
    try { await DB.removeWorkout(userId, entry); } catch {}
  }

  async function regeneratePlan() {
    if (!profile || !userId) return;
    const sportsDesc = (profile.selectedSports||[]).map(s => {
      const d = profile.sportData?.[s];
      if (s==="strength") return `MUSCULAÇÃO: ${d?.goal}, ${d?.experience}, ${d?.daysPerWeek}x/sem`;
      if (s==="running") return `CORRIDA: pace ${d?.currentPace5k}/km, meta ${d?.goalDistance}km, ${d?.weeklyKm}km/sem`;
      if (s==="bike") return `CICLISMO: FTP ${d?.currentFTP||"?"}, ${d?.bikeType}, ${d?.daysPerWeek}x/sem`;
      if (s==="swim") return `NATAÇÃO: ${d?.currentPace100}/100m, ${d?.style}, ${d?.weeklyVolume||"?"}m/sem`;
      return "";
    }).join("\n");
    const histSummary = workoutHistory.slice(-20).map(w=>`${w.date} ${w.type} ${w.plan?.focus||""} sensação:${w.log?.feel||"?"}`).join("; ");
    const avail = profile.availability;
    try {
      const res = await fetch(AI_URL, { method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1500, messages:[{ role:"user",
          content:`Regenere o plano semanal com base no histórico. Responda SOMENTE com JSON.
ATLETA: ${profile.basic?.name}, ${profile.basic?.age} anos, ${profile.basic?.weight}kg
ESPORTES:\n${sportsDesc}
DIAS DISPONÍVEIS: ${avail?.availableDays?.join(", ")||"todos"}
PREFERÊNCIAS: ${(profile.selectedSports||[]).map(s=>`${s}:${avail?.preferredDays?.[s]?.join(",")||"qualquer"}`).join("; ")}
HISTÓRICO (últimas 20): ${histSummary}
DURAÇÃO: ${profile.schedule?.sessionDuration||60}min
JSON: { "planName":"...","description":"...","weeklySchedule":[{"day":"Dom","type":"rest"},{"day":"Seg","type":"strength","focus":"...","distance":0,"paceTarget":"","hrZone":"","notes":""},...7 dias],"aiCoachNote":"..." }` }] })
      });
      const data = await res.json();
      const match = (data.content?.[0]?.text||"").match(/\{[\s\S]*\}/);
      if (match) {
        const newPlan = JSON.parse(match[0]);
        const updated = { ...profile, plan: newPlan };
        setProfile(updated); setTodayPlanFromProfile(updated);
        await DB.saveProfile(userId, updated);
      }
    } catch(e) { console.error(e); }
  }

  async function resetProfile() {
    try { await DB.deleteProfile(userId); } catch {}
    setProfile(null); setTodayPlan(null); setChat([]); setTodayCompleted(null);
    setScreen("onboarding");
  }

  async function signOut() {
    await DB.signOut();
    setUserId(null); setProfile(null); setWorkoutHistory([]);
    setTodayCompleted(null); setChat([]); setTodayPlan(null);
    setScreen("auth");
  }

  async function sendMsg(msg) {
    if (!msg.trim() || aiLoad) return;
    const uMsg = { role:"user", content:msg };
    const newMsgs = [...chat, uMsg];
    setChat(newMsgs); setChatIn(""); setAiLoad(true);
    const sportsContext = (profile?.selectedSports||[]).map(s => {
      const d = profile?.sportData?.[s];
      if (s==="strength") return `Musculação: ${d?.goal==="hypertrophy"?"hipertrofia":d?.goal}, ${d?.experience}, supino ${d?.currentBench||"?"}kg`;
      if (s==="running") return `Corrida: pace 5km ${d?.currentPace5k}/km, meta ${d?.goalDistance}km, ${d?.weeklyKm}km/sem`;
      if (s==="bike") return `Ciclismo: FTP ${d?.currentFTP||"desconhecido"}, ${d?.bikeType}, meta ${d?.goal}`;
      if (s==="swim") return `Natação: ${d?.currentPace100}/100m, ${d?.style}, meta ${d?.goal}`;
      return "";
    }).join(" | ");
    try {
      const res = await fetch(AI_URL, { method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000,
          system:`Você é o Coach IA do Athlete AI — treinador esportivo experiente, direto e cético. Seu diferencial é a honestidade que outros apps não têm coragem de ter.
FILOSOFIA: Resultados vêm de consistência e esforço honesto. Você não recompensa mediocridade — "completar o treino" é o mínimo esperado. Usa dados concretos. NUNCA usa: "Ótimo trabalho!", "Continue assim!", "Cada passo conta!", "O importante é participar!". Se o atleta busca validação fácil, oriente-o a usar outro app.
PERFIL: ${profile?.basic?.name||"Atleta"} | ${profile?.basic?.age} anos | ${profile?.basic?.weight}kg | ${profile?.basic?.height}cm
Esportes: ${sportsContext}
Plano: ${profile?.plan?.planName||"Em andamento"} | Hoje: ${todayPlan?.focus||todayPlan?.type||"Descanso"} | ${workoutHistory.length} sessões registradas
POSTURA: Direto. Máx 3 parágrafos. Base científica quando relevante. Nunca substitua médico. Português do Brasil.`,
          messages: newMsgs,
        })
      });
      const data = await res.json();
      setChat(prev => [...prev, { role:"assistant", content:data.content?.[0]?.text||"Erro ao processar." }]);
    } catch {
      setChat(prev => [...prev, { role:"assistant", content:"⚠️ Erro de conexão." }]);
    }
    setAiLoad(false);
  }

  useEffect(() => { if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight; }, [chat, aiLoad]);

  if (screen==="splash") return <Splash/>;
  if (screen==="auth") return <AuthScreen onAuth={handleAuth}/>;
  if (screen==="onboarding") return <Onboarding onComplete={saveProfile}/>;

  return (
    <div style={{ background:D.bg, minHeight:"100vh", fontFamily:"'Inter',sans-serif", color:D.text, maxWidth:480, margin:"0 auto", position:"relative", paddingBottom:80 }}>
      {tab==="home"&&<HomeTab profile={profile} onStartWorkout={plan=>{ setTodayPlan(plan); setTab("treino"); }} todayCompleted={todayCompleted} workoutHistory={workoutHistory} onDelete={()=>deleteWorkout(todayCompleted)}/>}
      {tab==="treino"&&<WorkoutTab todayPlan={todayPlan} profile={profile} onLogWorkout={logWorkout}/>}
      {tab==="coach"&&<CoachTab messages={chat} input={chatIn} loading={aiLoad} chatRef={chatRef} onInput={setChatIn} onSend={()=>sendMsg(chatIn)} profile={profile}/>}
      {tab==="progresso"&&<ProgressTab profile={profile} workoutHistory={workoutHistory} onDeleteEntry={deleteWorkout} userId={userId}/>}
      {tab==="perfil"&&<ProfileTab profile={profile} onReset={resetProfile} onRegenerate={regeneratePlan} workoutHistory={workoutHistory} userId={userId} onSignOut={signOut}/>}
      <BottomNav active={tab} onChange={setTab}/>
    </div>
  );
}

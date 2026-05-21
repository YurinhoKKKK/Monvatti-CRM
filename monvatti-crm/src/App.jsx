/**
 * MONVATTI CRM V4
 * Novidades: dark mode · responsivo (mobile) · portal dropdowns (fix status/user)
 * exportação XLSX + PDF · upload de foto · column DnD · filtro por data
 * duplicate/move copia campos equivalentes + atualizações
 */
import {
  useState, useEffect, useCallback, useRef,
  createContext, useContext, useMemo
} from "react";
import { createPortal } from "https://esm.sh/react-dom@18";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// EmailJS — envio automático de email do browser (sem backend)
// Docs: https://www.emailjs.com
// Para ativar: crie conta em emailjs.com → pegue Service ID, Template ID e Public Key
const EMAILJS_SERVICE  = "service_lz2vrni";   // ← substitua pelo seu Service ID
const EMAILJS_TEMPLATE = "template_fzh6cfq";     // ← substitua pelo seu Template ID
const EMAILJS_KEY      = "jtcaA8HV7AeVZy6_2";    // ← substitua pela sua Public Key
const sendEmailJS = async (params) => {
  try {
    const { default: emailjs } = await import("https://esm.sh/@emailjs/browser@4");
    emailjs.init({publicKey: EMAILJS_KEY});
    await emailjs.send(EMAILJS_SERVICE, EMAILJS_TEMPLATE, params);
    return {ok:true};
  } catch(err) {
    console.error("EmailJS error:", err);
    return {ok:false, error:err};
  }
};
// ─── CALLMEBOT — notificação WhatsApp para múltiplos números (sem backend) ───
// Para ativar cada número:
//   1. Salve +34 694 26 48 06 na agenda como "CallMeBot"
//   2. Envie "I allow callmebot to send me messages" via WhatsApp para esse número
//   3. Você receberá sua apikey em segundos — preencha abaixo
const WA_DESTINATARIOS = [
  { nome: "Yuri",     phone: "554899706309", apikey: "6328106"          },
  { nome: "Mauricio", phone: "554898652493", apikey: "3460088"          },
  { nome: "Alice",    phone: "",             apikey: ""                 }, // ← preencher
];
const sendWhatsApp = async (text) => {
  const ativos = WA_DESTINATARIOS.filter(d => d.phone && d.apikey);
  if (!ativos.length) return { ok: false, sent: 0 };
  const resultados = await Promise.allSettled(
    ativos.map(d =>
      fetch(`https://api.callmebot.com/whatsapp.php?phone=${d.phone}&text=${encodeURIComponent(text)}&apikey=${d.apikey}`)
        .then(r => ({ nome: d.nome, ok: r.ok }))
        .catch(() => ({ nome: d.nome, ok: false }))
    )
  );
  const enviados = resultados.filter(r => r.status === "fulfilled" && r.value.ok).map(r => r.value.nome);
  return { ok: enviados.length > 0, sent: enviados.length, total: ativos.length, enviados };
};

import * as XLSX from "https://esm.sh/xlsx@0.18.5";
import { jsPDF } from "https://esm.sh/jspdf@2.5.1";
import autoTable from "https://esm.sh/jspdf-autotable@3.8.3";

// ─── SUPABASE ────────────────────────────────────────────────────────────────
const SUPA_URL = "https://hyhealogmqylciuzdmkz.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh5aGVhbG9nbXF5bGNpdXpkbWt6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYxNzQxNDgsImV4cCI6MjA5MTc1MDE0OH0.vDr-lQp6lBcT0f2bZGnTk1Jh6jjle9-oL7m2iAgw3nA";
const db = createClient(SUPA_URL, SUPA_KEY);

// ─── THEME CSS (injected once) ───────────────────────────────────────────────
const THEME_CSS = `
:root{--bg:#DDE0EA;--surface:#FFFFFF;--surface2:#EFF1F7;--surface3:#E3E6EF;
--text:#161E2D;--text2:#3D4A5C;--text3:#6B7A90;--border:#B8BECE;--borderStrong:#9AA3B5;
--sidebar:#2B333B;--sidebarText:rgba(255,255,255,0.75);--sidebarActive:rgba(255,255,255,0.18);
--blue:#001AD8;--alt:#3145FF;--shadow:rgba(0,0,0,0.08);--shadowMd:rgba(0,0,0,0.14);--row-sel:#dbeafe;}
[data-dark=true]{--bg:#0D0F14;--surface:#1A1D27;--surface2:#1E2233;--surface3:#252A3E;
--text:#F0F2F5;--text2:#9CA3AF;--text3:#6B7280;--border:#2D3748;--borderStrong:#3D4A5C;
--sidebar:#111318;--sidebarText:rgba(255,255,255,0.7);--sidebarActive:rgba(255,255,255,0.14);
--blue:#3D5AFE;--alt:#536DFE;--shadow:rgba(0,0,0,0.3);--shadowMd:rgba(0,0,0,0.45);--row-sel:#1e3a5f;}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body,#root{height:100%;width:100%;overflow:hidden;font-family:system-ui,-apple-system,sans-serif;}
body{background:var(--bg);color:var(--text);}
input,select,textarea,button{font-family:inherit;}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
`;

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const BRAND = { blue:"var(--blue)", alt:"var(--alt)" };

const SC = {
  "Novo":"#3145FF","Em Contato":"#d97706","Sem Resposta":"#64748b",
  "Qualificado":"#059669","Não Qualificado":"#dc2626","Descartado":"#6b7280",
  "Proposta na Rua":"#3145FF","Proposta Enviada":"#7c3aed","Em Análise":"#d97706",
  "Fechado/Ganho":"#059669","Recusado":"#dc2626","Negócio Futuro":"#ea580c",
  "Não compareceu":"#94a3b8","Inativo":"#94a3b8","Retomar Contato":"#d97706",
  "Prospecção Fria":"#3145FF","Tráfego":"#7c3aed","Indicação":"#059669","Outro":"#6b7280",
};

// Normaliza options do config: aceita tanto string[] (legado) quanto objeto[] (novo)
// Formato novo: [{id, label, color, ordem}]
const resolveOpts = (raw=[]) => {
  if(!raw?.length) return [];
  if(typeof raw[0] === "string")
    return raw.map((s,i)=>({id:s, label:s, color:SC[s]||"#94a3b8", ordem:i}));
  return [...raw].sort((a,b)=>a.ordem-b.ordem);
};

// Gera id único para opção
const optId = () => Math.random().toString(36).slice(2,9);

// Paleta de cores para o seletor de opções
const OPT_COLORS = [
  "#3145FF","#001AD8","#7c3aed","#d97706","#059669","#dc2626",
  "#ea580c","#0891b2","#64748b","#94a3b8","#e11d48","#0F7B6C",
  "#B5451B","#7B2D8B","#1565C0","#065f46","#92400e","#1e3a5f",
];
const GC = ["#3145FF","#001AD8","#0F7B6C","#B5451B","#7B2D8B","#1565C0","#059669","#d97706"];
const BICONS = ["📋","🎯","📣","🤝","📁","💼","📊","🚀","⭐","🔥","💡","📈","🗂️","📌"];

// ─── ROLES & PERMISSIONS ─────────────────────────────────────────────────────
// Cada cargo define o que o usuário pode fazer no sistema.
// "all" = super-admin override.  Para SDR/Closer, acesso é controlado por group_access.
const ROLE_MATRIX = {
  administrador:     {all:true,sendToVendas:true},
  ceo:               {viewAll:true,editAny:true,manageBoards:true,manageCols:true,manageUsers:true,manageRoles:true,export:true,deleteAny:true,createGroups:true,sendToVendas:true},
  gerente_comercial: {viewAll:true,editAny:true,manageBoards:true,manageCols:true,manageUsers:true,manageRoles:false,export:true,deleteAny:true,createGroups:true,sendToVendas:true},
  financeiro:        {viewAll:true,editAny:false,manageBoards:false,manageCols:false,manageUsers:false,manageRoles:false,export:true,deleteAny:false,createGroups:false,sendToVendas:true},
  closer:            {viewAll:false,editAny:false,manageBoards:false,manageCols:false,manageUsers:false,manageRoles:false,export:false,deleteAny:false,createGroups:false,sendToVendas:false},
  sdr:               {viewAll:false,editAny:false,manageBoards:false,manageCols:false,manageUsers:false,manageRoles:false,export:false,deleteAny:false,createGroups:false,sendToVendas:false},
};
const DEFAULT_ROLES = [
  {nome:"Administrador",   slug:"administrador",   cor:"#dc2626"},
  {nome:"CEO",             slug:"ceo",             cor:"#001AD8"},
  {nome:"Gerente Comercial",slug:"gerente_comercial",cor:"#7c3aed"},
  {nome:"Financeiro",      slug:"financeiro",      cor:"#059669"},
  {nome:"Closer",          slug:"closer",          cor:"#d97706"},
  {nome:"SDR",             slug:"sdr",             cor:"#0891b2"},
];

// Context de permissões — derivado do perfil carregado
const PermCtx = createContext({perms:{},roleSlug:"",isAdmin:false});
const usePerms = ()=>useContext(PermCtx);

function getPerms(profile) {
  const slug=(profile?.role_slug||"sdr").toLowerCase();
  // Prioridade: permissões salvas no banco (roles.permissions) → fallback ROLE_MATRIX
  const dbPerms = profile?.roles?.permissions||{};
  const raw = Object.keys(dbPerms).length>0 ? dbPerms : (ROLE_MATRIX[slug]||ROLE_MATRIX.sdr);
  const isAdmin=!!(raw.all || slug==="administrador");
  return {
    all:          isAdmin,
    viewAll:      isAdmin||!!raw.viewAll,
    editAny:      isAdmin||!!raw.editAny,
    manageBoards: isAdmin||!!raw.manageBoards,
    manageCols:   isAdmin||!!raw.manageCols,
    manageUsers:  isAdmin||!!raw.manageUsers,
    manageRoles:  isAdmin||!!raw.manageRoles,
    export:       isAdmin||!!raw.export,
    deleteAny:    isAdmin||!!raw.deleteAny,
    createGroups: isAdmin||!!raw.createGroups,
    sendToVendas: isAdmin||!!raw.sendToVendas,
    isAdmin,
    isFull: isAdmin||!!raw.viewAll,
    slug,
  };
}
const CTYPES = [
  {v:"text",l:"Texto"},{v:"status",l:"Status (múltipla escolha)"},
  {v:"date",l:"Data"},{v:"number",l:"Número"},
  {v:"currency",l:"Moeda (R$)"},{v:"email",l:"E-mail"},
  {v:"phone",l:"Telefone"},{v:"link",l:"Link"},
  {v:"user",l:"Responsável"},{v:"calculated",l:"Calculado (Valor÷Parcelas)"},
];

// ─── UTILS ───────────────────────────────────────────────────────────────────
const deep   = x => JSON.parse(JSON.stringify(x));
const fmtBRL = n => `R$ ${Number(n).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtDate = d => {
  if(!d) return "";
  // Datas no formato YYYY-MM-DD sem horário são tratadas como UTC midnight
  // o que causa -1 dia em fusos negativos. Forçamos noon local para evitar isso.
  const s = String(d);
  const date = s.length===10 ? new Date(s+"T12:00:00") : new Date(s);
  return date.toLocaleDateString("pt-BR");
};
const uColor  = u => {
  const p=["#3145FF","#0F7B6C","#B5451B","#7B2D8B","#1565C0","#d97706","#059669","#e11d48","#0891b2"];
  return p[((u?.nome||u?.email||"x").charCodeAt(0))%p.length];
};
const uInit   = u => {
  if(!u?.nome) return "?";
  return u.nome.split(" ").slice(0,2).map(w=>w[0].toUpperCase()).join("");
};
const parseMD = t => {
  if(!t) return "";
  return t
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
    .replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>")
    .replace(/\*(.+?)\*/g,"<em>$1</em>")
    .replace(/^- (.+)$/gm,"<li>$1</li>")
    .replace(/((<li>.*?<\/li>\n?)+)/gs,m=>`<ul style="margin:4px 0;padding-left:18px">${m}</ul>`)
    .replace(/\n/g,"<br>");
};
const logAct = async (uid,wsId,type,id,action,meta={}) => {
  try{ await db.from("activity_logs").insert({user_id:uid,workspace_id:wsId,entity_type:type,entity_id:id,action,metadata:meta}); }catch(_){}
};

// ─── STYLE TOKENS ────────────────────────────────────────────────────────────
const T = {
  btn:  {display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,borderRadius:8,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer",border:"none"},
  inp:  {border:"1.5px solid var(--border)",borderRadius:8,padding:"10px 13px",fontSize:13,outline:"none",width:"100%",background:"var(--surface)",color:"var(--text)"},
  lbl:  {display:"block",fontSize:11,fontWeight:700,color:"var(--text3)",marginBottom:6,textTransform:"uppercase",letterSpacing:.5},
  iBtn: {background:"none",border:"1px solid var(--border)",borderRadius:6,padding:"4px 9px",cursor:"pointer",fontSize:12,color:"var(--text2)"},
};

// ─── BREAKPOINT HOOK ─────────────────────────────────────────────────────────
function useBreakpoint() {
  const [w, setW] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(()=>{
    const h = ()=>setW(window.innerWidth);
    window.addEventListener("resize",h);
    return()=>window.removeEventListener("resize",h);
  },[]);
  return { isMobile: w < 768, isTablet: w < 1024, width: w };
}

// ─────────────────────────────────────────────────────────────────────────────
// CONTEXTS
// ─────────────────────────────────────────────────────────────────────────────

// AUTH
const AuthCtx = createContext(null);
function AuthProvider({children}) {
  const [session,setSession] = useState(null);
  const [profile,setProfile] = useState(null);
  const [authLoading,setAuthLoading] = useState(true);
  const fetchProfile = useCallback(async uid=>{
    // CRÍTICO: incluir permissions no join para getPerms ler do banco
    const {data}=await db.from("profiles")
      .select("*, roles(id,slug,nome,cor,permissions)")
      .eq("id",uid).single();
    if(data) setProfile({
      ...data,
      role_slug: data.roles?.slug||"sdr",
      role_nome: data.roles?.nome||"Sem cargo",
    });
    else setProfile(null);
  },[]);
  useEffect(()=>{
    db.auth.getSession().then(({data:{session:s}})=>{
      setSession(s);
      if(s) fetchProfile(s.user.id).finally(()=>setAuthLoading(false));
      else setAuthLoading(false);
    });
    const {data:{subscription}}=db.auth.onAuthStateChange((_,s)=>{
      setSession(s); if(s) fetchProfile(s.user.id); else{setProfile(null);setAuthLoading(false);}
    });
    return()=>subscription.unsubscribe();
  },[fetchProfile]);
  const signIn  = (e,p)=>db.auth.signInWithPassword({email:e,password:p});
  const signUp  = async(e,p,nome)=>{
    const {data,error}=await db.auth.signUp({email:e,password:p,options:{data:{nome}}});
    if(error) return{data,error};
    // Supabase pode retornar user null quando email confirmation está ativo mas rate limit foi atingido
    if(data?.user){
      // Garante perfil mesmo antes do trigger disparar
      await db.from("profiles")
        .upsert({id:data.user.id, email:e, nome, role_id:null},{onConflict:"id"});
      await fetchProfile(data.user.id);
    }
    return{data,error};
  };
  const signOut = async()=>{await db.auth.signOut();setSession(null);setProfile(null);};
  const updateProfile = async upd=>{
    if(!session) return{error:"no session"};
    const {data,error}=await db.from("profiles").update({...upd,updated_at:new Date().toISOString()}).eq("id",session.user.id).select().single();
    if(!error) await fetchProfile(session.user.id); // re-fetch with roles join to preserve cargo
    return{data,error};
  };
  const updatePassword = pw=>db.auth.updateUser({password:pw});
  return <AuthCtx.Provider value={{session,profile,authLoading,signIn,signUp,signOut,updateProfile,updatePassword,fetchProfile}}>{children}</AuthCtx.Provider>;
}
const useAuth = ()=>useContext(AuthCtx);

// THEME
const ThemeCtx = createContext(null);
function ThemeProvider({children}) {
  const [dark,setDark] = useState(()=>localStorage.getItem("crm-theme")==="dark");
  useEffect(()=>{
    document.documentElement.setAttribute("data-dark",String(dark));
    localStorage.setItem("crm-theme",dark?"dark":"light");
  },[dark]);
  return (
    <ThemeCtx.Provider value={{dark,toggle:()=>setDark(p=>!p)}}>
      {/* Injeta variáveis CSS globalmente, incluindo nas páginas de auth */}
      <style>{THEME_CSS}</style>
      {children}
    </ThemeCtx.Provider>
  );
}
const useTheme = ()=>useContext(ThemeCtx);

// TOAST
const ToastCtx = createContext(null);
function ToastProvider({children}) {
  const [toasts,setToasts] = useState([]);
  const add = useCallback((msg,type="success")=>{
    const id=Date.now()+Math.random();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3800);
  },[]);
  const col={success:"#059669",error:"#dc2626",info:"#3145FF",warning:"#d97706"};
  return (
    <ToastCtx.Provider value={add}>
      {children}
      {createPortal(
        <div style={{position:"fixed",bottom:28,right:24,zIndex:4000,display:"flex",flexDirection:"column-reverse",gap:10}}>
          <style>{`@keyframes ti{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}`}</style>
          {toasts.map(t=>(
            <div key={t.id} style={{background:col[t.type]||col.success,color:"#fff",padding:"13px 20px",borderRadius:12,fontSize:13,fontWeight:600,boxShadow:"0 6px 28px rgba(0,0,0,.24)",maxWidth:340,animation:"ti .2s ease",display:"flex",alignItems:"center",gap:10}}>
              {t.type==="error"?"✕":t.type==="warning"?"⚠":"✓"} {t.msg}
            </div>
          ))}
        </div>,
        document.body
      )}
    </ToastCtx.Provider>
  );
}
const useToast = ()=>useContext(ToastCtx);

// ─────────────────────────────────────────────────────────────────────────────
// PORTAL DROPDOWN — FIX para células status/user não clicáveis em tabelas
// ─────────────────────────────────────────────────────────────────────────────
function PortalDrop({trigger, children, minWidth=220, maxHeight=320}) {
  const [open,setOpen] = useState(false);
  const [pos,setPos]   = useState({top:0,left:0,showAbove:false});
  const trigRef = useRef();
  const dropRef = useRef();

  const openDrop = ()=>{
    if(trigRef.current){
      const r = trigRef.current.getBoundingClientRect();
      const spaceBelow = window.innerHeight - r.bottom - 8;
      const spaceAbove = r.top - 8;
      const showAbove  = spaceBelow < Math.min(maxHeight, 200) && spaceAbove > spaceBelow;
      // Garante que não sai horizontalmente
      const left = Math.max(8, Math.min(r.left, window.innerWidth - minWidth - 8));
      setPos({top: r.bottom + 4, bottom: window.innerHeight - r.top + 4, left, showAbove});
    }
    setOpen(p=>!p);
  };

  useEffect(()=>{
    if(!open) return;
    const h = e=>{
      if(dropRef.current&&!dropRef.current.contains(e.target)&&trigRef.current&&!trigRef.current.contains(e.target))
        setOpen(false);
    };
    const t=setTimeout(()=>document.addEventListener("mousedown",h),0);
    return()=>{clearTimeout(t);document.removeEventListener("mousedown",h);};
  },[open]);

  // Após render do dropdown, corrige posição se ainda sair da tela
  useEffect(()=>{
    if(!open||!dropRef.current) return;
    const dr = dropRef.current.getBoundingClientRect();
    if(dr.right > window.innerWidth - 8){
      setPos(p=>({...p,left:Math.max(8,window.innerWidth-dr.width-8)}));
    }
    if(!pos.showAbove && dr.bottom > window.innerHeight - 8){
      setPos(p=>({...p,showAbove:true}));
    }
  },[open]);

  return (
    <>
      <div ref={trigRef} onClick={openDrop} style={{cursor:"pointer"}}>
        {trigger}
      </div>
      {open && createPortal(
        <div ref={dropRef} style={{
          position:"fixed",
          ...(pos.showAbove ? {bottom:pos.bottom,top:"auto"} : {top:pos.top}),
          left:pos.left,
          zIndex:9999,
          background:"var(--surface)",
          border:"1px solid var(--border)",
          borderRadius:12,
          boxShadow:"0 12px 48px var(--shadowMd)",
          minWidth,
          maxHeight,
          overflowY:"auto",
          padding:8,
        }}>
          {children(()=>setOpen(false))}
        </div>,
        document.body
      )}
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// UI PRIMITIVES
// ─────────────────────────────────────────────────────────────────────────────
function Spinner({size=36}) {
  return (
    <>
      <style>{`@keyframes sp{to{transform:rotate(360deg)}}`}</style>
      <div style={{width:size,height:size,border:"3px solid var(--border)",borderTop:"3px solid var(--blue)",borderRadius:"50%",animation:"sp .8s linear infinite",flexShrink:0}}/>
    </>
  );
}

function Avatar({user,size=30}) {
  if(!user) return null;
  const bg=uColor(user), label=uInit(user);
  return (
    <div title={user.nome} style={{width:size,height:size,borderRadius:"50%",background:bg,color:"#fff",
      display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.max(9,size*.36),fontWeight:700,
      flexShrink:0,border:"2.5px solid var(--surface)",overflow:"hidden",marginRight:-5}}>
      {user.foto_url?<img src={user.foto_url} style={{width:"100%",height:"100%",objectFit:"cover"}} alt=""/>:label}
    </div>
  );
}

// color prop opcional — quando vem de opção rica usa a cor definida, senão cai no SC map
function Badge({value,color=null,size="normal"}) {
  const col=color||SC[value]||"#94a3b8";
  const p=size==="sm"?"2px 8px":"3px 11px";
  const fs=size==="sm"?11:12;
  return (
    <span style={{display:"inline-flex",alignItems:"center",background:col+"22",color:col,
      border:`1px solid ${col}45`,borderRadius:6,padding:p,fontSize:fs,fontWeight:600,whiteSpace:"nowrap",
      maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}}>
      {value}
    </span>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MODAL
// ─────────────────────────────────────────────────────────────────────────────
function Modal({title,onClose,children,width=490,footer=null}) {
  useEffect(()=>{
    const h=e=>{if(e.key==="Escape")onClose();};
    document.addEventListener("keydown",h);
    return()=>document.removeEventListener("keydown",h);
  },[onClose]);
  return createPortal(
    <div onClick={e=>{if(e.target===e.currentTarget)onClose();}}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,.52)",display:"flex",alignItems:"center",
        justifyContent:"center",zIndex:2000,padding:16}}>
      <div style={{background:"var(--surface)",borderRadius:16,width,maxWidth:"100%",maxHeight:"92vh",
        display:"flex",flexDirection:"column",boxShadow:"0 24px 72px var(--shadowMd)"}}>
        <div style={{padding:"18px 24px",borderBottom:"1px solid var(--border)",display:"flex",
          alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <span style={{fontWeight:800,fontSize:16,color:"var(--text)"}}>{title}</span>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:26,color:"var(--text3)",lineHeight:1}}>×</button>
        </div>
        <div style={{padding:"22px 24px",overflowY:"auto",flex:1}}>{children}</div>
        {footer&&<div style={{padding:"14px 24px",borderTop:"1px solid var(--border)",display:"flex",gap:10,justifyContent:"flex-end",flexShrink:0}}>{footer}</div>}
      </div>
    </div>,
    document.body
  );
}

function ConfirmModal({title,message,danger=false,onConfirm,onCancel}) {
  return (
    <Modal title={title} onClose={onCancel} width={420}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={onConfirm} style={{...T.btn,background:danger?"#dc2626":"var(--blue)",color:"#fff"}}>{danger?"Excluir":"Confirmar"}</button>
      </>}>
      <p style={{margin:0,fontSize:14,color:"var(--text2)",lineHeight:1.7}}>{message}</p>
    </Modal>
  );
}

function GroupSelectorModal({title,groups,onSelect,onCancel}) {
  return (
    <Modal title={title} onClose={onCancel} width={380}>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {groups.map(g=>(
          <button key={g.id} onClick={()=>onSelect(g)}
            style={{...T.btn,background:"var(--surface2)",color:"var(--text)",border:`1.5px solid ${g.color}50`,
              justifyContent:"flex-start",padding:"12px 16px",width:"100%"}}>
            <div style={{width:11,height:11,borderRadius:"50%",background:g.color,flexShrink:0}}/>
            {g.nome}
          </button>
        ))}
      </div>
    </Modal>
  );
}

// ─── MODAL ENVIAR PARA NEGOCIAÇÕES — 2 etapas: Closer → Sub-grupo ────────────
function SendToNegModal({groups, onSelect, onCancel}) {
  const [step,setStep]=useState(1);
  const [selParent,setSelParent]=useState(null);
  const parents=groups.filter(g=>g.is_parent);
  const childrenOf=pid=>groups.filter(g=>g.parent_group_id===pid);
  const closerLabel=nome=>(nome||"").replace(/Closer\s*[-\u2013]\s*/i,"").trim();
  const getInitials=nome=>closerLabel(nome).split(" ").slice(0,2).map(p=>p[0]||"").join("").toUpperCase();

  if(step===1) return (
    <Modal title="🤝 Enviar para Negociações" onClose={onCancel} width={430}>
      <div style={{marginBottom:14,fontSize:13,color:"var(--text3)"}}>
        Selecione o <strong style={{color:"var(--text)"}}>Closer</strong> que receberá o lead:
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {parents.map(p=>{
          const subs=childrenOf(p.id);
          return (
            <button key={p.id} onClick={()=>{setSelParent(p);setStep(2);}}
              style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",
                background:"var(--surface2)",border:`2px solid ${p.color}35`,borderRadius:12,
                cursor:"pointer",width:"100%",textAlign:"left",transition:"all .15s"}}
              onMouseEnter={e=>{e.currentTarget.style.borderColor=p.color;e.currentTarget.style.background="var(--surface3)";}}
              onMouseLeave={e=>{e.currentTarget.style.borderColor=`${p.color}35`;e.currentTarget.style.background="var(--surface2)";}}>
              <div style={{width:44,height:44,borderRadius:"50%",background:p.color,flexShrink:0,
                display:"flex",alignItems:"center",justifyContent:"center",
                fontWeight:800,fontSize:17,color:"#fff",letterSpacing:.5}}>
                {getInitials(p.nome)}
              </div>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:15,color:"var(--text)",marginBottom:3}}>{closerLabel(p.nome)}</div>
                <div style={{fontSize:12,color:"var(--text3)"}}>{subs.map(s=>s.nome).join(" · ")}</div>
              </div>
              <div style={{fontSize:20,color:p.color,opacity:.6,fontWeight:300}}>›</div>
            </button>
          );
        })}
      </div>
    </Modal>
  );

  const subs=childrenOf(selParent.id);
  return (
    <Modal title="🤝 Enviar para Negociações" onClose={onCancel} width={430}>
      <button onClick={()=>setStep(1)}
        style={{display:"flex",alignItems:"center",gap:10,width:"100%",
          padding:"10px 14px",background:`${selParent.color}15`,border:`1.5px solid ${selParent.color}40`,
          borderRadius:10,cursor:"pointer",marginBottom:16,textAlign:"left"}}>
        <span style={{fontSize:16,color:selParent.color}}>←</span>
        <div style={{width:32,height:32,borderRadius:"50%",background:selParent.color,flexShrink:0,
          display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,color:"#fff"}}>
          {getInitials(selParent.nome)}
        </div>
        <div>
          <div style={{fontWeight:700,fontSize:13,color:"var(--text)"}}>{closerLabel(selParent.nome)}</div>
          <div style={{fontSize:11,color:"var(--text3)"}}>Escolha o sub-grupo de destino</div>
        </div>
      </button>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {subs.map(sg=>(
          <button key={sg.id} onClick={()=>onSelect(sg)}
            style={{display:"flex",alignItems:"center",gap:12,padding:"13px 16px",
              background:"var(--surface2)",border:`1.5px solid ${sg.color}45`,borderRadius:10,
              cursor:"pointer",width:"100%",textAlign:"left",transition:"all .15s"}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor=sg.color;e.currentTarget.style.background="var(--surface3)";}}
            onMouseLeave={e=>{e.currentTarget.style.borderColor=`${sg.color}45`;e.currentTarget.style.background="var(--surface2)";}}>
            <div style={{width:10,height:10,borderRadius:"50%",background:sg.color,flexShrink:0}}/>
            <span style={{fontWeight:600,fontSize:14,color:"var(--text)",flex:1}}>{sg.nome}</span>
          </button>
        ))}
      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// CELLS — todas usando PortalDrop para dropdowns
// ─────────────────────────────────────────────────────────────────────────────
function EditableCell({value,onChange,type="text",placeholder=""}) {
  const [ed,setEd]=useState(false);
  const [v,setV]=useState(value??"");
  const ref=useRef();
  useEffect(()=>setV(value??""),[value]);
  useEffect(()=>{if(ed)ref.current?.focus();},[ed]);
  const commit=()=>{setEd(false);if(v!==(value??""))onChange(v||null);};
  if(!ed) return (
    <div onClick={()=>setEd(true)}
      style={{padding:"5px 8px",minHeight:30,minWidth:80,cursor:"text",fontSize:13,
        color:v?"var(--text)":"var(--text3)",borderRadius:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
      {type==="date"&&v?fmtDate(v):v||placeholder}
    </div>
  );
  return (
    <input ref={ref} type={type==="date"?"date":type==="number"?"number":"text"} value={v}
      onChange={e=>setV(e.target.value)} onBlur={commit}
      onKeyDown={e=>{if(e.key==="Enter")commit();if(e.key==="Escape"){setEd(false);setV(value??"");}}}
      style={{border:"2px solid var(--blue)",borderRadius:5,padding:"4px 8px",fontSize:13,outline:"none",
        width:"100%",boxSizing:"border-box",background:"var(--surface)",color:"var(--text)"}}
    />
  );
}

function CurrencyCell({value,onChange}) {
  const [ed,setEd]=useState(false);
  const [v,setV]=useState(value??"");
  useEffect(()=>setV(value??""),[value]);
  if(!ed) return (
    <div onClick={()=>setEd(true)}
      style={{padding:"5px 8px",minWidth:110,cursor:"text",fontSize:13,color:value!=null?"var(--text)":"var(--text3)"}}>
      {value!=null?fmtBRL(value):"R$ —"}
    </div>
  );
  return (
    <input autoFocus type="number" value={v} onChange={e=>setV(e.target.value)}
      onBlur={()=>{setEd(false);onChange(v!==""?parseFloat(v):null);}}
      onKeyDown={e=>{if(e.key==="Enter"){setEd(false);onChange(v!==""?parseFloat(v):null);}}}
      style={{border:"2px solid var(--blue)",borderRadius:5,padding:"4px 8px",fontSize:13,outline:"none",
        width:"100%",boxSizing:"border-box",background:"var(--surface)",color:"var(--text)"}}
    />
  );
}

// FIX: PortalDrop garante que o menu apareça mesmo dentro de td com overflow:hidden
// options aceita tanto string[] (legado) quanto objeto[] {id,label,color,ordem} (novo)
function StatusCell({value,options=[],onChange}) {
  const resolved = useMemo(()=>resolveOpts(options),[options]);
  const current  = resolved.find(o=>o.label===value)||null;
  const useGrid  = resolved.length >= 5;
  const minW     = useGrid ? Math.min(480, Math.max(280, resolved.length * 48)) : 230;

  const trigger = (
    <div style={{padding:"4px 4px",userSelect:"none",minWidth:120}}>
      {current
        ? <div style={{
            display:"inline-flex",alignItems:"center",gap:6,
            padding:"4px 10px 4px 8px",borderRadius:6,cursor:"pointer",
            background:`${current.color}18`,border:`1.5px solid ${current.color}50`,
            transition:"all .15s",maxWidth:"100%",overflow:"hidden"}}>
            <div style={{width:7,height:7,borderRadius:"50%",background:current.color,flexShrink:0,
              boxShadow:`0 0 5px ${current.color}80`}}/>
            <span style={{fontSize:11,fontWeight:700,color:current.color,
              whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",letterSpacing:.2}}>
              {current.label}
            </span>
          </div>
        : <span style={{fontSize:11,color:"var(--text3)",padding:"4px 10px",
            border:"1.5px dashed var(--border)",borderRadius:6,
            display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",whiteSpace:"nowrap"}}>
            <span style={{opacity:.5}}>◯</span> Selecionar
          </span>
      }
    </div>
  );
  if(!resolved.length) return trigger;
  return (
    <PortalDrop trigger={trigger} minWidth={minW} maxHeight={400}>
      {close=>(
        <>
          {useGrid
            ? <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:5,padding:4}}>
                {resolved.map(opt=>(
                  <div key={opt.id} onClick={()=>{onChange(opt.label);close();}}
                    style={{padding:"8px 10px",borderRadius:8,cursor:"pointer",display:"flex",alignItems:"center",
                      gap:7,background:value===opt.label?"var(--surface3)":"transparent",
                      border:value===opt.label?`1.5px solid ${opt.color}40`:"1.5px solid transparent",transition:"all .1s"}}
                    onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"}
                    onMouseLeave={e=>e.currentTarget.style.background=value===opt.label?"var(--surface3)":"transparent"}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:opt.color,flexShrink:0}}/>
                    <span style={{fontSize:11,fontWeight:600,color:"var(--text)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{opt.label}</span>
                    {value===opt.label&&<span style={{marginLeft:"auto",color:opt.color,fontSize:13,flexShrink:0}}>✓</span>}
                  </div>
                ))}
              </div>
            : <div>
                {resolved.map(opt=>(
                  <div key={opt.id} onClick={()=>{onChange(opt.label);close();}}
                    style={{padding:"8px 12px",borderRadius:7,cursor:"pointer",display:"flex",alignItems:"center",
                      gap:9,marginBottom:2,background:value===opt.label?"var(--surface3)":"transparent"}}
                    onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"}
                    onMouseLeave={e=>e.currentTarget.style.background=value===opt.label?"var(--surface3)":"transparent"}>
                    <Badge value={opt.label} color={opt.color}/>
                    {value===opt.label&&<span style={{marginLeft:"auto",color:"var(--blue)",fontSize:15}}>✓</span>}
                  </div>
                ))}
              </div>
          }
          {value&&<>
            <div style={{height:1,background:"var(--border)",margin:"5px 0"}}/>
            <div onClick={()=>{onChange(null);close();}}
              style={{padding:"8px 12px",borderRadius:7,cursor:"pointer",color:"var(--text3)",fontSize:12}}
              onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              ✕ Limpar seleção
            </div>
          </>}
        </>
      )}
    </PortalDrop>
  );
}

function UserCell({value=[],allUsers=[],onChange}) {
  const cur=Array.isArray(value)?value:[];
  const sel=cur.map(id=>allUsers.find(u=>u.id===id)).filter(Boolean);
  const toggle=uid=>onChange(cur.includes(uid)?cur.filter(x=>x!==uid):[...cur,uid]);
  const trigger=(
    <div style={{display:"flex",alignItems:"center",padding:"5px 5px",minHeight:32,minWidth:56}}>
      {!sel.length
        ? <div style={{width:28,height:28,borderRadius:"50%",border:"1.5px dashed var(--border)",
            display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:"var(--text3)"}}>+</div>
        : <div style={{display:"flex",alignItems:"center",maxWidth:68,overflow:"hidden",flexShrink:0}}>{sel.slice(0,3).map((u,idx)=><div key={u.id} style={{marginLeft:idx?-7:0,flexShrink:0,zIndex:10-idx,borderRadius:"50%",border:"2px solid var(--surface)"}}><Avatar user={u} size={24}/></div>)}{sel.length>3&&<div style={{marginLeft:-7,width:24,height:24,borderRadius:"50%",background:"var(--surface3)",border:"2px solid var(--surface)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:"var(--text2)",flexShrink:0}}>+{sel.length-3}</div>}</div>
      }
    </div>
  );
  return (
    <PortalDrop trigger={trigger} minWidth={260} maxHeight={380}>
      {close=>(
        <>
          {!allUsers.length&&<div style={{padding:"12px 14px",fontSize:13,color:"var(--text3)"}}>Nenhum usuário</div>}
          {allUsers.map(u=>(
            <div key={u.id} onClick={()=>toggle(u.id)}
              style={{display:"flex",alignItems:"center",gap:11,padding:"9px 13px",borderRadius:8,cursor:"pointer",
                background:cur.includes(u.id)?"var(--surface3)":"transparent",
                border:cur.includes(u.id)?"1px solid var(--blue)30":"1px solid transparent"}}
              onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"}
              onMouseLeave={e=>e.currentTarget.style.background=cur.includes(u.id)?"var(--surface3)":"transparent"}>
              <Avatar user={u} size={32}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:600,color:"var(--text)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{u.nome}</div>
                <div style={{fontSize:11,color:"var(--text3)"}}>{u.funcao||u.email}</div>
              </div>
              {cur.includes(u.id)&&<span style={{color:"var(--blue)",fontWeight:700,flexShrink:0}}>✓</span>}
            </div>
          ))}
        </>
      )}
    </PortalDrop>
  );
}

// FIX: window.open com noopener
function LinkCell({value,onChange}) {
  const [ed,setEd]=useState(false);
  const [v,setV]=useState(value||"");
  useEffect(()=>setV(value||""),[value]);
  const href=v?(v.startsWith("http")?v:`https://${v}`):null;
  if(!ed) return (
    <div style={{display:"flex",alignItems:"center",gap:6,padding:"5px 8px",minHeight:30,minWidth:100}}>
      <div onClick={()=>setEd(true)} style={{flex:1,fontSize:13,color:v?"#3b82f6":"var(--text3)",cursor:"text",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
        {v||"https://"}
      </div>
      {href&&(
        <button onClick={e=>{e.stopPropagation();window.open(href,"_blank","noopener,noreferrer");}}
          style={{...T.iBtn,width:24,height:24,padding:0,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>↗</button>
      )}
    </div>
  );
  return (
    <input autoFocus value={v} onChange={e=>setV(e.target.value)}
      onBlur={()=>{setEd(false);onChange(v||null);}}
      onKeyDown={e=>{if(e.key==="Enter"){setEd(false);onChange(v||null);}if(e.key==="Escape"){setEd(false);setV(value||"");}}}
      style={{border:"2px solid var(--blue)",borderRadius:5,padding:"4px 8px",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",background:"var(--surface)",color:"var(--text)"}}
    />
  );
}

// Torna links clicáveis no HTML das atualizações
function linkifyHTML(html) {
  if(!html) return "";
  // 1. Garante target=_blank em <a> já existentes
  let out = html.replace(/<a(\s)/gi,'<a target="_blank" rel="noopener noreferrer"$1');
  // Wrapper para criar link
  const mkLink = (url, display, href) =>
    `<a href="${href||url}" target="_blank" rel="noopener noreferrer" style="color:var(--blue);text-decoration:underline;word-break:break-all">${display||url}</a>`;
  // 2. https:// ou http:// (já dentro de href="" ignora — lookbehind de ="  ou href)
  out = out.replace(/(?<![="'\/])(https?:\/\/[^\s<>"']+)/g, url=>mkLink(url));
  // 3. www. sem protocolo
  out = out.replace(/(?<![="'\/\w])(www\.[a-zA-Z0-9][a-zA-Z0-9.-]+\.[a-zA-Z]{2,}[^\s<>"']*)/g,
    url=>mkLink(url, url, "https://"+url));
  // 4. Domínios comuns sem www nem protocolo (ex: youtube.com, google.com.br)
  out = out.replace(/(?<![="'\/\w.])((?:[a-zA-Z0-9-]+\.)+(?:com\.br|com|net|org|io|app|dev|co|me|tv|info|biz|store|shop|edu|gov|br)(?:\/[^\s<>"']*)?)/g,
    url=>mkLink(url, url, "https://"+url));
  return out;
}

// FIX: detecta colunas currency/number dinamicamente
function CalcCell({values,allColumns}) {
  const curCol=useMemo(()=>allColumns.find(c=>c.tipo==="currency"),[allColumns]);
  const parCol=useMemo(()=>allColumns.find(c=>c.tipo==="number"&&c.nome.toLowerCase().includes("parcela")),[allColumns]);
  if(!curCol||!parCol) return <div style={{padding:"5px 8px",color:"var(--text3)",fontSize:13}}>—</div>;
  const v=parseFloat(values?.[curCol.id])||0;
  const p=parseFloat(values?.[parCol.id])||0;
  const r=p>0?v/p:null;
  return <div style={{padding:"5px 8px",fontSize:13,color:r!=null?"#059669":"var(--text3)",fontWeight:r!=null?600:400}}>{r!=null?fmtBRL(r):"—"}</div>;
}

function Cell({col,values,allColumns,responsibles,allUsers,onChange,onRespChange}) {
  const v=values?.[col.id];
  const opts=col.config?.options||[];
  switch(col.tipo){
    case "calculated": return <CalcCell values={values} allColumns={allColumns}/>;
    case "status":     return <StatusCell value={v} options={opts} onChange={onChange}/>;
    case "user":       return <UserCell value={responsibles} allUsers={allUsers} onChange={onRespChange}/>;
    case "currency":   return <CurrencyCell value={v} onChange={onChange}/>;
    case "link":       return <LinkCell value={v} onChange={onChange}/>;
    case "date":       return <EditableCell value={v} onChange={onChange} type="date"/>;
    case "number":     return <EditableCell value={v} onChange={onChange} type="number" placeholder="0"/>;
    case "phone":      return <EditableCell value={v} onChange={onChange} placeholder="(11) 9…"/>;
    case "email":      return <EditableCell value={v} onChange={onChange} placeholder="email@"/>;
    default:           return <EditableCell value={v} onChange={onChange} placeholder={col.nome}/>;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// RICH EDITOR — contentEditable proper (sem execCommand obsoleto)
// ─────────────────────────────────────────────────────────────────────────────
function RichEditor({onSubmit,disabled=false}) {
  const editorRef = useRef(null);
  const [isEmpty, setIsEmpty] = useState(true);
  const [fontSize, setFontSize] = useState(13);
  const [align, setAlign] = useState("left");

  // Aplica comando de formatação via Selection API
  const fmt = (tag, attrs={}) => {
    editorRef.current?.focus();
    const sel = window.getSelection();
    if(!sel||!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const el = document.createElement(tag);
    Object.entries(attrs).forEach(([k,v])=>el.style[k]=v);
    try {
      range.surroundContents(el);
    } catch {
      // If range crosses node boundaries, wrap extracted contents
      el.appendChild(range.extractContents());
      range.insertNode(el);
    }
    sel.removeAllRanges();
    const newRange = document.createRange();
    newRange.selectNodeContents(el);
    newRange.collapse(false);
    sel.addRange(newRange);
    checkEmpty();
  };

  const insertList = (ordered) => {
    editorRef.current?.focus();
    const sel = window.getSelection();
    if(!sel||!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const list = document.createElement(ordered?"ol":"ul");
    list.style.paddingLeft="20px";list.style.margin="4px 0";
    const li = document.createElement("li");
    li.appendChild(range.extractContents()||document.createTextNode("​"));
    list.appendChild(li);
    range.insertNode(list);
    const r2=document.createRange();
    r2.setStartAfter(li);r2.collapse(true);
    sel.removeAllRanges();sel.addRange(r2);
    checkEmpty();
  };

  const insertChecklist = () => {
    editorRef.current?.focus();
    const div = document.createElement("div");
    div.style.display="flex";div.style.alignItems="center";div.style.gap="6px";div.style.margin="3px 0";
    const cb = document.createElement("input");
    cb.type="checkbox";cb.style.accentColor="var(--blue)";cb.style.width="14px";cb.style.height="14px";cb.style.flexShrink="0";
    const span = document.createElement("span");
    span.textContent="​";
    cb.onchange=()=>{span.style.textDecoration=cb.checked?"line-through":"none";span.style.opacity=cb.checked?"0.5":"1";};
    div.appendChild(cb);div.appendChild(span);
    const sel=window.getSelection();
    if(sel&&sel.rangeCount){const r=sel.getRangeAt(0);r.insertNode(div);r.setStartAfter(div);r.collapse(true);sel.removeAllRanges();sel.addRange(r);}
    span.focus?.();
    checkEmpty();
  };

  const setAlignStyle = (a) => {
    setAlign(a);
    editorRef.current?.focus();
    const sel=window.getSelection();
    if(!sel||!sel.rangeCount) return;
    const node=sel.getRangeAt(0).commonAncestorContainer;
    const block=node.nodeType===3?node.parentElement:node;
    if(block&&editorRef.current?.contains(block)) block.style.textAlign=a;
  };

  const applyFontSize = (size) => {
    setFontSize(size);
    editorRef.current?.style && (editorRef.current.style.fontSize=size+"px");
    editorRef.current?.focus();
  };

  const checkEmpty = () => {
    const t = editorRef.current?.textContent||"";
    setIsEmpty(t.trim()==="");
  };

  const submit = () => {
    const html = editorRef.current?.innerHTML||"";
    if(!html||isEmpty||disabled) return;
    onSubmit(html);
    editorRef.current.innerHTML="";
    setIsEmpty(true);
  };

  const ToolBtn=({onClick,active=false,children,title=""})=>(
    <button onMouseDown={e=>{e.preventDefault();onClick();}} title={title}
      style={{background:active?"var(--blue)":"none",color:active?"#fff":"var(--text2)",
        border:"1px solid var(--border)",borderRadius:5,padding:"3px 8px",cursor:"pointer",
        fontSize:12,fontFamily:"inherit",minWidth:28,display:"flex",alignItems:"center",justifyContent:"center",
        transition:"all .12s"}}>
      {children}
    </button>
  );

  const Divider=()=><div style={{width:1,height:16,background:"var(--border)",flexShrink:0}}/>;

  return (
    <div style={{border:"1.5px solid var(--border)",borderRadius:10,overflow:"hidden",background:"var(--surface)"}}>
      {/* Toolbar */}
      <div style={{display:"flex",gap:4,padding:"8px 10px",borderBottom:"1px solid var(--border)",
        background:"var(--surface2)",flexWrap:"wrap",alignItems:"center"}}>
        {/* Font size */}
        <select value={fontSize} onChange={e=>applyFontSize(Number(e.target.value))}
          style={{border:"1px solid var(--border)",borderRadius:5,padding:"3px 6px",fontSize:11,
            background:"var(--surface)",color:"var(--text)",cursor:"pointer",outline:"none"}}>
          {[10,11,12,13,14,16,18,20,24].map(s=><option key={s} value={s}>{s}px</option>)}
        </select>
        <Divider/>
        {/* Bold, Italic, Underline, Strikethrough */}
        <ToolBtn onClick={()=>fmt("strong")} title="Negrito (⌘B)"><strong>N</strong></ToolBtn>
        <ToolBtn onClick={()=>fmt("em")} title="Itálico (⌘I)"><em>I</em></ToolBtn>
        <ToolBtn onClick={()=>fmt("u")} title="Sublinhado"><span style={{textDecoration:"underline"}}>S</span></ToolBtn>
        <ToolBtn onClick={()=>fmt("s")} title="Tachado"><span style={{textDecoration:"line-through"}}>T</span></ToolBtn>
        <Divider/>
        {/* Alignment */}
        <ToolBtn onClick={()=>setAlignStyle("left")} active={align==="left"} title="Alinhar à esquerda">⬅</ToolBtn>
        <ToolBtn onClick={()=>setAlignStyle("center")} active={align==="center"} title="Centralizar">⬛</ToolBtn>
        <ToolBtn onClick={()=>setAlignStyle("right")} active={align==="right"} title="Alinhar à direita">➡</ToolBtn>
        <Divider/>
        {/* Lists */}
        <ToolBtn onClick={()=>insertList(false)} title="Lista com marcadores">• ≡</ToolBtn>
        <ToolBtn onClick={()=>insertList(true)} title="Lista numerada">1. ≡</ToolBtn>
        <ToolBtn onClick={insertChecklist} title="Checklist">☑</ToolBtn>
        <Divider/>
        {/* Color */}
        <div style={{display:"flex",alignItems:"center",gap:4}}>
          <span style={{fontSize:11,color:"var(--text3)"}}>Cor:</span>
          {["#2B333B","#dc2626","#d97706","#059669","#3145FF","#7c3aed"].map(c=>(
            <div key={c} onClick={()=>fmt("span",{color:c})}
              style={{width:16,height:16,borderRadius:3,background:c,cursor:"pointer",
                border:"2px solid transparent",transition:"transform .1s"}}
              onMouseEnter={e=>e.currentTarget.style.transform="scale(1.25)"}
              onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}/>
          ))}
        </div>
      </div>

      {/* Editor area */}
      <div
        ref={editorRef}
        contentEditable={!disabled}
        suppressContentEditableWarning
        onInput={checkEmpty}
        onKeyDown={e=>{
          if((e.metaKey||e.ctrlKey)&&e.key==="Enter"){e.preventDefault();submit();}
          if((e.metaKey||e.ctrlKey)&&e.key==="b"){e.preventDefault();fmt("strong");}
          if((e.metaKey||e.ctrlKey)&&e.key==="i"){e.preventDefault();fmt("em");}
          if((e.metaKey||e.ctrlKey)&&e.key==="u"){e.preventDefault();fmt("u");}
        }}
        data-placeholder="Escreva uma atualização… ⌘↵ para publicar"
        style={{minHeight:100,maxHeight:300,overflowY:"auto",
          padding:"12px 14px",outline:"none",fontSize:fontSize,
          lineHeight:1.7,color:"var(--text)",background:"var(--surface)",
          textAlign:align,
          // Placeholder via CSS
        }}
      />
      <style>{`
        [data-placeholder]:empty:before{content:attr(data-placeholder);color:var(--text3);pointer-events:none;}
      `}</style>

      {/* Footer */}
      <div style={{padding:"8px 12px",borderTop:"1px solid var(--border)",display:"flex",
        justifyContent:"flex-end",alignItems:"center",gap:12,background:"var(--surface2)"}}>
        <span style={{fontSize:11,color:"var(--text3)"}}>⌘↵ publicar · ⌘B negrito · ⌘I itálico</span>
        <button onClick={submit} disabled={disabled||isEmpty}
          style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 20px",
            opacity:disabled||isEmpty?0.5:1,fontSize:13}}>
          Publicar
        </button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// OPTION EDITOR — editor visual de opções de status (nome, cor, posição)
// ─────────────────────────────────────────────────────────────────────────────
function OptionEditor({options,onChange}) {
  const [dragIdx,setDragIdx]=useState(null);
  const [dragOverIdx,setDragOverIdx]=useState(null);
  const [editingId,setEditingId]=useState(null); // id da opção com color picker aberto

  const add=()=>{
    const newOpt={id:optId(),label:"Nova opção",color:OPT_COLORS[options.length%OPT_COLORS.length],ordem:options.length};
    onChange([...options,newOpt]);
    // Foca no novo campo após render
    setTimeout(()=>{
      const el=document.getElementById(`opt-label-${newOpt.id}`);
      if(el){el.focus();el.select();}
    },50);
  };

  const update=(id,field,val)=>onChange(options.map(o=>o.id===id?{...o,[field]:val}:o));

  const remove=id=>onChange(options.filter(o=>o.id!==id).map((o,i)=>({...o,ordem:i})));

  const drop=()=>{
    if(dragIdx===null||dragOverIdx===null||dragIdx===dragOverIdx){setDragIdx(null);setDragOverIdx(null);return;}
    const next=[...options];
    const [moved]=next.splice(dragIdx,1);
    next.splice(dragOverIdx,0,moved);
    onChange(next.map((o,i)=>({...o,ordem:i})));
    setDragIdx(null);setDragOverIdx(null);
  };

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <label style={T.lbl}>Opções</label>
        <button onClick={add}
          style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"5px 14px",fontSize:12,borderRadius:7}}>
          + Adicionar opção
        </button>
      </div>

      {!options.length&&(
        <div style={{textAlign:"center",padding:"20px 0",color:"var(--text3)",fontSize:13,
          border:"1.5px dashed var(--border)",borderRadius:9}}>
          Nenhuma opção. Clique em "+ Adicionar opção" para começar.
        </div>
      )}

      {options.map((opt,i)=>(
        <div key={opt.id} draggable
          onDragStart={()=>setDragIdx(i)}
          onDragOver={e=>{e.preventDefault();setDragOverIdx(i);}}
          onDrop={drop}
          onDragEnd={()=>{setDragIdx(null);setDragOverIdx(null);}}
          style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:9,marginBottom:7,
            background:dragOverIdx===i?"var(--surface3)":"var(--surface2)",
            border:`1.5px solid ${dragOverIdx===i?"var(--blue)":"var(--border)"}`,
            transition:"border .15s,background .15s",cursor:"grab"}}>

          {/* Handle de arrastar */}
          <span style={{color:"var(--text3)",fontSize:15,flexShrink:0,cursor:"grab"}}>⋮⋮</span>

          {/* Seletor de cor — clique abre paleta inline */}
          <div style={{position:"relative",flexShrink:0}}>
            <div onClick={()=>setEditingId(editingId===opt.id?null:opt.id)}
              title="Escolher cor"
              style={{width:26,height:26,borderRadius:7,background:opt.color,cursor:"pointer",
                border:"2px solid rgba(0,0,0,.12)",flexShrink:0,transition:"transform .1s"}}
              onMouseEnter={e=>e.currentTarget.style.transform="scale(1.15)"}
              onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}
            />
            {editingId===opt.id&&(
              <div style={{position:"absolute",top:"100%",left:0,zIndex:100,background:"var(--surface)",
                border:"1px solid var(--border)",borderRadius:12,padding:12,marginTop:6,
                boxShadow:"0 8px 32px var(--shadowMd)",width:220}}>
                <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",marginBottom:10,textTransform:"uppercase",letterSpacing:.5}}>Escolher cor</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
                  {OPT_COLORS.map(c=>(
                    <div key={c} onClick={()=>{update(opt.id,"color",c);}}
                      title={c}
                      style={{width:28,height:28,borderRadius:7,background:c,cursor:"pointer",
                        border:opt.color===c?"3px solid var(--text)":"2px solid transparent",
                        transition:"transform .1s,border .1s"}}
                      onMouseEnter={e=>e.currentTarget.style.transform="scale(1.15)"}
                      onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}
                    />
                  ))}
                </div>
                {/* Input hex manual */}
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:20,height:20,borderRadius:4,background:opt.color,flexShrink:0,border:"1px solid var(--border)"}}/>
                  <input value={opt.color} onChange={e=>update(opt.id,"color",e.target.value)}
                    placeholder="#3145FF"
                    style={{...T.inp,padding:"5px 9px",fontSize:12,flex:1}}/>
                  <button onClick={()=>setEditingId(null)}
                    style={{...T.iBtn,padding:"5px 9px",fontSize:12,flexShrink:0}}>OK</button>
                </div>
              </div>
            )}
          </div>

          {/* Preview do badge */}
          <div style={{flexShrink:0}}>
            <Badge value={opt.label||"…"} color={opt.color}/>
          </div>

          {/* Input do nome */}
          <input id={`opt-label-${opt.id}`}
            value={opt.label} onChange={e=>update(opt.id,"label",e.target.value)}
            onKeyDown={e=>{if(e.key==="Enter"){e.preventDefault();add();}}}
            placeholder="Nome da opção"
            style={{...T.inp,flex:1,padding:"6px 10px",fontSize:13,minWidth:0}}/>

          {/* Botão excluir */}
          <button onClick={()=>remove(opt.id)} title="Remover opção"
            style={{...T.iBtn,color:"#dc2626",borderColor:"#FEE2E2",padding:"5px 8px",flexShrink:0,lineHeight:1}}>✕</button>
        </div>
      ))}

      {options.length>0&&(
        <p style={{fontSize:11,color:"var(--text3)",marginTop:8}}>
          Arraste ⋮⋮ para reordenar · Clique na cor para alterar · Enter para adicionar próxima
        </p>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN MANAGER — CRUD + status options + drag reorder
// ─────────────────────────────────────────────────────────────────────────────
function ColumnFormModal({initial={},boardId,onSave,onCancel}) {
  const [nome,setNome]=useState(initial.nome||"");
  const [tipo,setTipo]=useState(initial.tipo||"text");
  // Normaliza opções existentes para o formato rico
  const [opts,setOpts]=useState(()=>resolveOpts(initial.config?.options||[]));

  const save=()=>{
    if(!nome.trim()) return;
    const config=tipo==="status"
      ? {options:opts.map((o,i)=>({...o,ordem:i}))}  // salva objetos ricos
      : {};
    onSave({nome:nome.trim(),tipo,config});
  };

  return (
    <Modal title={initial.id?"Editar coluna":"Nova coluna"} onClose={onCancel} width={520}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={save} disabled={!nome.trim()} style={{...T.btn,background:"var(--blue)",color:"#fff",opacity:nome.trim()?1:.5}}>Salvar</button>
      </>}>
      <label style={T.lbl}>Nome da coluna</label>
      <input value={nome} onChange={e=>setNome(e.target.value)} autoFocus
        onKeyDown={e=>{if(e.key==="Enter"&&tipo!=="status")save();}}
        style={{...T.inp,marginBottom:18}} placeholder="Ex: Status, Data Reunião, Valor…"/>
      <label style={T.lbl}>Tipo</label>
      <select value={tipo} onChange={e=>{setTipo(e.target.value);}} style={{...T.inp,marginBottom:24}}>
        {CTYPES.map(t=><option key={t.v} value={t.v}>{t.l}</option>)}
      </select>
      {tipo==="status"&&(
        <OptionEditor options={opts} onChange={setOpts}/>
      )}
    </Modal>
  );
}

function ColumnManagerModal({board,onClose,onRefresh,toast}) {
  const [cols,setCols]=useState(deep(board.columns||[]));
  const [formM,setFormM]=useState(null);
  const [confirmM,setConfirmM]=useState(null);
  const [saving,setSaving]=useState(false);
  const [dragIdx,setDragIdx]=useState(null);
  const [dragOverIdx,setDragOverIdx]=useState(null);

  const addCol=async({nome,tipo,config})=>{
    setSaving(true);
    const {data,error}=await db.from("columns").insert({board_id:board.id,nome,tipo,config,ordem:cols.length+1}).select().single();
    if(!error&&data) setCols(p=>[...p,{...data,config:data.config||{}}]);
    else toast(error?.message||"Erro","error");
    setSaving(false);setFormM(null);onRefresh();
  };
  const editCol=async({nome,tipo,config})=>{
    const id=formM.column.id; setSaving(true);
    const {error}=await db.from("columns").update({nome,tipo,config}).eq("id",id);
    if(!error) setCols(p=>p.map(c=>c.id===id?{...c,nome,tipo,config}:c));
    else toast(error.message,"error");
    setSaving(false);setFormM(null);onRefresh();
  };
  const delCol=async id=>{
    setSaving(true);
    await db.from("columns").delete().eq("id",id);
    setCols(p=>p.filter(c=>c.id!==id));
    setConfirmM(null);setSaving(false);onRefresh();
  };
  const moveCol=async(id,dir)=>{
    const idx=cols.findIndex(c=>c.id===id); if(idx<0) return;
    const swap=idx+dir; if(swap<0||swap>=cols.length) return;
    const next=[...cols];[next[idx],next[swap]]=[next[swap],next[idx]];
    next.forEach((c,i)=>c.ordem=i+1);
    setCols(next);
    await Promise.all(next.map(c=>db.from("columns").update({ordem:c.ordem}).eq("id",c.id)));
    onRefresh();
  };
  const dropCol=async()=>{
    if(dragIdx===null||dragOverIdx===null||dragIdx===dragOverIdx){setDragIdx(null);setDragOverIdx(null);return;}
    const next=[...cols];const [moved]=next.splice(dragIdx,1);next.splice(dragOverIdx,0,moved);
    next.forEach((c,i)=>c.ordem=i+1);
    setCols(next);setDragIdx(null);setDragOverIdx(null);
    await Promise.all(next.map(c=>db.from("columns").update({ordem:c.ordem}).eq("id",c.id)));
    onRefresh();
  };

  return (
    <>
      <Modal title={`Colunas — ${board.nome}`} onClose={onClose} width={580}
        footer={<>
          <button onClick={()=>setFormM({})} style={{...T.btn,background:"var(--blue)",color:"#fff"}}>+ Nova coluna</button>
          <button onClick={onClose} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Fechar</button>
        </>}>
        <p style={{fontSize:12,color:"var(--text3)",marginBottom:16}}>Arraste para reordenar. Duplo-clique no ícone ⋮⋮ para arrastar.</p>
        {!cols.length&&<div style={{textAlign:"center",color:"var(--text3)",padding:"24px 0",fontSize:14}}>Nenhuma coluna.</div>}
        {cols.map((col,i)=>(
          <div key={col.id} draggable
            onDragStart={()=>setDragIdx(i)}
            onDragOver={e=>{e.preventDefault();setDragOverIdx(i);}}
            onDrop={dropCol}
            style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:9,
              background:dragOverIdx===i?"var(--surface3)":"var(--surface2)",marginBottom:8,
              border:`1.5px solid ${dragOverIdx===i?"var(--blue)":"var(--border)"}`,transition:"border .15s",cursor:"grab"}}>
            <span style={{color:"var(--text3)",fontSize:16,flexShrink:0}}>⋮⋮</span>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:600,fontSize:14,color:"var(--text)"}}>{col.nome}</div>
              <div style={{fontSize:11,color:"var(--text3)",marginTop:2}}>{CTYPES.find(t=>t.v===col.tipo)?.l||col.tipo}</div>
            </div>
            <div style={{display:"flex",gap:6,flexShrink:0}}>
              <button onClick={()=>moveCol(col.id,-1)} disabled={i===0} style={{...T.iBtn,opacity:i===0?0.3:1}}>↑</button>
              <button onClick={()=>moveCol(col.id,1)} disabled={i===cols.length-1} style={{...T.iBtn,opacity:i===cols.length-1?0.3:1}}>↓</button>
              <button onClick={()=>setFormM({column:col})} style={T.iBtn}>✏️</button>
              <button onClick={()=>setConfirmM({id:col.id,nome:col.nome})} style={{...T.iBtn,color:"#dc2626",borderColor:"#FEE2E2"}}>🗑️</button>
            </div>
          </div>
        ))}
      </Modal>
      {formM!==null&&(formM.column
        ? <ColumnFormModal initial={formM.column} boardId={board.id} onSave={editCol} onCancel={()=>setFormM(null)}/>
        : <ColumnFormModal initial={{}} boardId={board.id} onSave={addCol} onCancel={()=>setFormM(null)}/>
      )}
      {confirmM&&<ConfirmModal title="Excluir coluna" danger
        message={`Excluir "${confirmM.nome}"? Todos os dados desta coluna serão perdidos.`}
        onConfirm={()=>delCol(confirmM.id)} onCancel={()=>setConfirmM(null)}/>}
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// EXPORT MODAL — XLSX + PDF com cores da marca
// ─────────────────────────────────────────────────────────────────────────────
function ExportModal({board,onClose}) {
  const [busy,setBusy]=useState(false);
  const toast=useToast();

  const getRows=()=>{
    const rows=[];
    for(const g of board.groups){
      rows.push({_group:g.nome,_color:g.color});
      for(const item of g.items){
        const row={_groupName:g.nome};
        for(const col of board.columns){
          if(col.tipo==="user") row[col.nome]="—";
          else if(col.tipo==="calculated"){
            const cC=board.columns.find(c=>c.tipo==="currency");
            const pC=board.columns.find(c=>c.tipo==="number"&&c.nome.toLowerCase().includes("parcela"));
            if(cC&&pC){const v=parseFloat(item.values?.[cC.id])||0,p=parseFloat(item.values?.[pC.id])||0;row[col.nome]=p>0?v/p:null;}
            else row[col.nome]=null;
          }else{
            const v=item.values?.[col.id];
            row[col.nome]=v!=null?String(v):"";
          }
        }
        rows.push(row);
      }
    }
    return rows;
  };

  const exportXLSX=()=>{
    setBusy(true);
    try{
      const wb=XLSX.utils.book_new();
      const colNames=board.columns.map(c=>c.nome);
      const wsData=[["Grupo",...colNames]];
      const styles=[];
      for(const g of board.groups){
        styles.push({row:wsData.length,type:"group",color:g.color});
        wsData.push([g.nome,...Array(colNames.length).fill("")]);
        for(const item of g.items){
          const row=[g.nome];
          for(const col of board.columns){
            if(col.tipo==="user"){ row.push(""); continue; }
            if(col.tipo==="calculated"){
              const cC=board.columns.find(c=>c.tipo==="currency");
              const pC=board.columns.find(c=>c.tipo==="number"&&c.nome.toLowerCase().includes("parcela"));
              if(cC&&pC){const v=parseFloat(item.values?.[cC.id])||0,p=parseFloat(item.values?.[pC.id])||0;row.push(p>0?(v/p):0);}
              else row.push(0);
              continue;
            }
            const v=item.values?.[col.id];
            if(col.tipo==="currency"||col.tipo==="number") row.push(v!=null?Number(v):null);
            else row.push(v!=null?String(v):"");
          }
          wsData.push(row);
        }
      }
      const ws=XLSX.utils.aoa_to_sheet(wsData);
      // Largura de colunas
      ws["!cols"]=[{wch:18},...colNames.map(()=>({wch:22}))];
      XLSX.utils.book_append_sheet(wb,ws,board.nome.substring(0,30));
      XLSX.writeFile(wb,`${board.nome.replace(/\s+/g,"_")}_${new Date().toLocaleDateString("pt-BR").replace(/\//g,"-")}.xlsx`);
      toast("Excel exportado!");
    }catch(e){toast("Erro ao exportar","error");console.error(e);}
    setBusy(false);
  };

  const exportPDF=()=>{
    setBusy(true);
    try{
      const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});
      const colNames=board.columns.filter(c=>c.tipo!=="user"&&c.tipo!=="calculated").map(c=>c.nome);
      const colIds=board.columns.filter(c=>c.tipo!=="user"&&c.tipo!=="calculated").map(c=>c.id);

      doc.setFont("helvetica","bold");
      doc.setFontSize(18);
      doc.setTextColor(0,26,216);
      doc.text(`${board.icon||""} ${board.nome}`, 14, 18);
      doc.setFontSize(10);
      doc.setTextColor(107,114,128);
      doc.text(`Gerado em ${new Date().toLocaleString("pt-BR")}`, 14, 26);

      const body=[];
      for(const g of board.groups){
        if(!g.items.length) continue;
        body.push([{content:`📁 ${g.nome}`,colSpan:colNames.length,styles:{fontStyle:"bold",fillColor:[245,247,250],textColor:[43,51,59]}}]);
        for(const item of g.items){
          const row=[];
          for(const col of board.columns.filter(c=>c.tipo!=="user"&&c.tipo!=="calculated")){
            const v=item.values?.[col.id];
            if(col.tipo==="currency") row.push(v!=null?fmtBRL(v):"");
            else if(col.tipo==="date") row.push(v?fmtDate(v):"");
            else row.push(v!=null?String(v):"");
          }
          body.push(row);
        }
      }

      autoTable(doc,{
        head:[colNames],
        body,
        startY:32,
        headStyles:{fillColor:[0,26,216],textColor:[255,255,255],fontStyle:"bold",fontSize:9},
        bodyStyles:{fontSize:8,textColor:[43,51,59]},
        alternateRowStyles:{fillColor:[248,249,252]},
        tableLineColor:[229,231,235],
        tableLineWidth:.3,
        margin:{left:14,right:14},
      });

      doc.save(`${board.nome.replace(/\s+/g,"_")}_${new Date().toLocaleDateString("pt-BR").replace(/\//g,"-")}.pdf`);
      toast("PDF exportado!");
    }catch(e){toast("Erro ao exportar PDF","error");console.error(e);}
    setBusy(false);
  };

  return (
    <Modal title="Exportar relatório" onClose={onClose} width={440}>
      <p style={{fontSize:13,color:"var(--text2)",marginBottom:24,lineHeight:1.6}}>
        Os relatórios incluem todas as colunas visíveis do quadro <strong>{board.nome}</strong>, organizados por grupo.
      </p>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        <button onClick={exportXLSX} disabled={busy}
          style={{...T.btn,background:"#059669",color:"#fff",padding:"16px 20px",justifyContent:"flex-start",gap:14,width:"100%"}}>
          <span style={{fontSize:28}}>📊</span>
          <div style={{textAlign:"left"}}>
            <div style={{fontWeight:700}}>Exportar Excel (.xlsx)</div>
            <div style={{fontSize:12,opacity:.85,fontWeight:400}}>Planilha editável com dados por grupo</div>
          </div>
        </button>
        <button onClick={exportPDF} disabled={busy}
          style={{...T.btn,background:"#dc2626",color:"#fff",padding:"16px 20px",justifyContent:"flex-start",gap:14,width:"100%"}}>
          <span style={{fontSize:28}}>📄</span>
          <div style={{textAlign:"left"}}>
            <div style={{fontWeight:700}}>Exportar PDF</div>
            <div style={{fontSize:12,opacity:.85,fontWeight:400}}>Relatório formatado para impressão</div>
          </div>
        </button>
      </div>
      {busy&&<div style={{display:"flex",justifyContent:"center",marginTop:20}}><Spinner size={28}/></div>}
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ITEM ROW — FIX hover sem deslocamento (checkbox sempre visível)
// ─────────────────────────────────────────────────────────────────────────────
function ItemMenuOpt({icon,label,onClick,danger=false}) {
  return (
    <div onClick={onClick}
      style={{display:"flex",alignItems:"center",gap:10,padding:"9px 13px",borderRadius:7,cursor:"pointer",
        fontSize:13,color:danger?"#dc2626":"var(--text)"}}
      onMouseEnter={e=>e.currentTarget.style.background=danger?"#fef2f2":"var(--surface3)"}
      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
      {icon} {label}
    </div>
  );
}

function ItemRow({item,columns,gc,allUsers,selected,onToggle,onOpen,onDelete,onMoveInativa,onDupNeg,onSendToNeg,onSendToVendas,
  onDragStart,onDragOver,onDrop,onUpdateValue,onRespChange,sentToNegIds=new Set()}) {
  const [hov,setHov]=useState(false);
  const [menu,setMenu]=useState(false);
  const [menuPos,setMenuPos]=useState({top:0,right:0});
  const btnRef=useRef();
  const menuRef=useRef();
  useEffect(()=>{
    if(!menu) return;
    const h=e=>{
      if(!btnRef.current?.contains(e.target)&&!menuRef.current?.contains(e.target))
        setMenu(false);
    };
    const t=setTimeout(()=>document.addEventListener("mousedown",h),0);
    return()=>{clearTimeout(t);document.removeEventListener("mousedown",h);};
  },[menu]);
  const openMenu=()=>{
    const r=btnRef.current?.getBoundingClientRect();
    if(r) setMenuPos({top:r.bottom+4,right:window.innerWidth-r.right});
    setMenu(p=>!p);
  };
  const alreadySent=sentToNegIds.has(item.id);
  const rowBg=selected?"var(--row-sel)":hov?"var(--surface2)":"var(--surface)";
  return (
    <tr draggable onDragStart={onDragStart} onDragOver={onDragOver} onDrop={onDrop}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}
      style={{background:rowBg,transition:"background .1s"}}>
      {/* Célula de controles — layout original preservado, badge via borda + dot */}
      <td style={{width:52,textAlign:"center",
        borderLeft:`3px solid ${alreadySent?"#059669":gc}`,
        verticalAlign:"middle",padding:"0 6px",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:5,position:"relative"}}>
          <button onClick={onOpen} title={alreadySent?"Atualizações (já enviado para Negociações)":"Atualizações"}
            style={{...T.iBtn,opacity:hov||selected?1:0.25,transition:"opacity .15s",fontSize:13,padding:"2px 3px",flexShrink:0,position:"relative"}}>
            📝
            {alreadySent&&<span title="Lead enviado para Negociações" style={{position:"absolute",top:-4,right:-4,width:10,height:10,
              borderRadius:"50%",background:"#00C46A",border:"2px solid var(--surface)",
              display:"block",boxShadow:"0 0 6px #00C46A90"}}/>}
          </button>
          <input type="checkbox" checked={selected} onChange={onToggle}
            style={{cursor:"pointer",accentColor:"var(--blue)",width:13,height:13,flexShrink:0}}/>
        </div>
      </td>
      {columns.map(col=>(
        <td key={col.id} style={{padding:"2px 3px",borderRight:"1px solid var(--border)",verticalAlign:"middle",maxWidth:220}}>
          <Cell col={col} values={item.values} allColumns={columns}
            responsibles={item.responsibles} allUsers={allUsers}
            onChange={v=>onUpdateValue(col.id,v)} onRespChange={onRespChange}/>
        </td>
      ))}
      <td style={{width:72,padding:"0 6px",textAlign:"right",verticalAlign:"middle",flexShrink:0}}>
        <div style={{display:"flex",gap:4,justifyContent:"flex-end",opacity:hov||selected||menu?1:0,transition:"opacity .1s"}}>
          <button ref={btnRef} onClick={openMenu} title="Ações" style={T.iBtn}>⋯</button>
        </div>
        {menu&&createPortal(
          <div ref={menuRef} style={{position:"fixed",top:menuPos.top,right:menuPos.right,zIndex:5000,
            background:"var(--surface)",border:"1px solid var(--border)",borderRadius:12,
            boxShadow:"0 12px 48px var(--shadowMd)",minWidth:240,padding:7}}>
            {onDupNeg&&<ItemMenuOpt icon="📋" label="Duplicar → Negociação" onClick={()=>{setMenu(false);onDupNeg();}}/>}
            {onMoveInativa&&<ItemMenuOpt icon="📁" label="Mover para Base Inativa" onClick={()=>{setMenu(false);onMoveInativa();}}/>}
            {onSendToNeg&&(
              alreadySent
                ? <ItemMenuOpt icon="🔁" label="Reenviar para Negociações" onClick={()=>{setMenu(false);onSendToNeg();}}/>
                : <ItemMenuOpt icon="🤝" label="Enviar para Negociações" onClick={()=>{setMenu(false);onSendToNeg();}}/>
            )}
            {onSendToVendas&&<ItemMenuOpt icon="🏆" label="Enviar para Vendas" onClick={()=>{setMenu(false);onSendToVendas();}}/>}
            <div style={{height:1,background:"var(--border)",margin:"5px 0"}}/>
            <ItemMenuOpt icon="🗑️" label="Excluir item" danger onClick={()=>{setMenu(false);onDelete();}}/>
          </div>,
          document.body
        )}
      </td>
    </tr>
  );
}

// Mobile card view para itens
function ItemCard({item,columns,gc,allUsers,selected,onToggle,onOpen,onDelete}) {
  const firstCol=columns[0];
  const title=item.values?.[firstCol?.id]||"Sem título";
  const statusCol=columns.find(c=>c.tipo==="status");
  const statusVal=statusCol?item.values?.[statusCol.id]:null;
  return (
    <div style={{background:"var(--surface)",borderRadius:10,padding:"14px 16px",marginBottom:8,
      borderLeft:`4px solid ${gc}`,border:`1px solid var(--border)`,borderLeftWidth:4,borderLeftColor:gc,
      boxShadow:`0 2px 8px var(--shadow)`}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
        <input type="checkbox" checked={selected} onChange={onToggle} style={{accentColor:"var(--blue)",width:14,height:14}}/>
        <div style={{flex:1,fontWeight:600,fontSize:14,color:"var(--text)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{title}</div>
        {statusVal&&<Badge value={statusVal} size="sm"/>}
        <button onClick={onOpen} style={T.iBtn}>✏️</button>
        <button onClick={onDelete} style={{...T.iBtn,color:"#dc2626",borderColor:"#FEE2E2"}}>🗑️</button>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
        {columns.slice(1,5).map(col=>{
          const v=item.values?.[col.id];
          if(!v) return null;
          return (
            <div key={col.id} style={{fontSize:11,color:"var(--text3)"}}>
              <span style={{fontWeight:600}}>{col.nome}: </span>
              <span>{col.tipo==="currency"?fmtBRL(v):col.tipo==="date"?fmtDate(v):String(v)}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Linha de totais do grupo ───────────────────────────────────────────────────
function GroupTotals({columns,items,gc}) {
  const currencyCols = columns.filter(c=>c.tipo==="currency");
  const curCol = columns.find(c=>c.tipo==="currency");
  const parCol = columns.find(c=>c.tipo==="number"&&c.nome.toLowerCase().includes("parcela"));

  if(!currencyCols.length||!items.length) return null;

  const parse = raw=>{
    if(raw===null||raw===undefined||raw==="") return 0;
    const v=typeof raw==="object"&&raw!==null&&"value" in raw?raw.value:raw;
    return parseFloat(v)||0;
  };

  // Soma somente itens onde a coluna específica tem valor > 0 (evita itens "fantasma")
  const sums = currencyCols.map(col=>{
    const seen=new Set(); let total=0,count=0;
    for(const it of items){
      if(seen.has(it.id))continue; seen.add(it.id);
      const v=parse(it.values?.[col.id]);
      if(v>0){total+=v;count++;}
    }
    return{col,total,count};
  }).filter(x=>x.total>0);

  let mrrTotal=0,mrrCount=0;
  if(curCol&&parCol){
    const seen=new Set();
    for(const it of items){
      if(seen.has(it.id))continue; seen.add(it.id);
      const v=parse(it.values?.[curCol.id]);
      const p=parse(it.values?.[parCol.id]);
      if(v>0&&p>0){mrrTotal+=v/p;mrrCount++;}
    }
  }

  if(!sums.length&&mrrTotal===0) return null;

  return (
    <div style={{borderLeft:`3px solid ${gc}`,background:`${gc}08`,borderTop:`1.5px dashed ${gc}30`,
      padding:"9px 16px",display:"flex",flexWrap:"wrap",alignItems:"center",gap:6}}>
      <span style={{fontSize:10,fontWeight:800,color:gc,textTransform:"uppercase",letterSpacing:.7,
        paddingRight:10,marginRight:4,borderRight:`1px solid ${gc}30`,flexShrink:0}}>Σ Totais</span>
      {sums.map(({col,total,count})=>(
        <span key={col.id} style={{fontSize:13,fontWeight:700,color:"var(--text)",display:"inline-flex",
          alignItems:"center",gap:5,background:"var(--surface3)",borderRadius:8,padding:"3px 10px",border:`1px solid ${gc}20`}}>
          <span style={{fontSize:10,color:"var(--text3)",fontWeight:500}}>{col.nome}: </span>
          {fmtBRL(total)}
          <span style={{fontSize:10,color:"var(--text3)",background:"var(--surface)",borderRadius:5,padding:"1px 5px"}}>{count}×</span>
        </span>
      ))}
      {mrrTotal>0&&(
        <span style={{fontSize:13,fontWeight:700,color:"#059669",display:"inline-flex",alignItems:"center",gap:5,
          background:"#05996912",borderRadius:8,padding:"3px 10px",border:"1px solid #05996930"}}>
          <span style={{fontSize:10,color:"#059669",fontWeight:500,opacity:.8}}>Valor Mensal: </span>
          {fmtBRL(mrrTotal)}
          <span style={{fontSize:10,color:"#059669",background:"#05996920",borderRadius:5,padding:"1px 5px",opacity:.8}}>{mrrCount}×</span>
        </span>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// GROUP
// ─────────────────────────────────────────────────────────────────────────────
const SB={background:"none",border:"1px solid",borderRadius:6,padding:"4px 11px",cursor:"pointer",fontSize:12};

function Group({group,columns,items,isDraggingOver,allUsers,selectedItems,isMobile,
  perms,currentUser,groupAccess,
  onToggleItem,onSelectAll,onAddItem,onDelGroup,onRenameGroup,onToggle,
  onOpenItem,onUpdateValue,onRespChange,onDelItem,onMoveInativa,onDupNeg,onSendToNeg,onSendToVendas,
  onDragStart,onDragOver,onDrop,onItemDragOver,onItemDrop,onGroupSettings,sentToNegIds=new Set(),
  sortCfg={colId:null,dir:1},setSortCfg=()=>{}}) {
  const [renaming,setRenaming]=useState(false);
  const [gname,setGname]=useState(group.nome);
  const nref=useRef();
  const stickyRef=useRef();   // header scroll container (overflow:hidden)
  const bodyRef=useRef();     // body scroll container (overflow:auto)
  useEffect(()=>{ if(renaming)nref.current?.focus(); },[renaming]);
  const allSel=items.length>0&&items.every(i=>selectedItems.has(i.id));
  const canEdit=perms.all||perms.editAny||(perms.slug==="sdr"&&group.owner_id===currentUser?.id)||((perms.slug==="closer")&&(groupAccess||[]).includes(currentUser?.id));
  const canRename=perms.all||["administrador","ceo","financeiro","gerente_comercial"].includes(perms.slug);

  // Sincroniza scroll horizontal: arrastar o body move o header junto
  const syncHeader=()=>{
    if(stickyRef.current&&bodyRef.current)
      stickyRef.current.scrollLeft=bodyRef.current.scrollLeft;
  };

  // Colgroup compartilhado — garante alinhamento exato entre header e body
  const colWidths=columns.map(c=>
    c.tipo==="date"?105:c.tipo==="currency"||c.tipo==="number"||c.tipo==="calculated"?110:
    c.tipo==="user"?75:c.tipo==="status"?130:140
  );
  const totalW=52+colWidths.reduce((s,w)=>s+w,0)+72;
  const cg=(
    <colgroup>
      <col style={{width:52}}/>
      {columns.map((col,i)=><col key={col.id} style={{width:colWidths[i]}}/>)}
      <col style={{width:72}}/>
    </colgroup>
  );

  return (
    <div style={{marginTop:16,borderRadius:11,
      border:isDraggingOver?`2px solid ${group.color}`:"2px solid transparent",transition:"border .15s"}}
      onDragOver={onDragOver} onDrop={onDrop}>

      {/* ── CABEÇALHO STICKY (desktop): overflow:hidden → synced via JS ── */}
      {!isMobile&&(
        <div ref={stickyRef} style={{
          position:"sticky",top:0,zIndex:5,
          overflow:"hidden",          // sem scrollbar, mas scrollLeft é setável via JS
          background:"var(--surface)",
          borderRadius:group.collapsed?"11px":"11px 11px 0 0",
          boxShadow:"0 2px 10px var(--shadow)"}}>

          {/* Linha do grupo */}
          <div style={{display:"flex",alignItems:"center",gap:9,padding:"10px 14px",
            background:"var(--surface)",borderLeft:`4px solid ${group.color}`,minWidth:totalW}}>
            <button onClick={onToggle}
              style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:group.color,
                padding:0,transform:group.collapsed?"rotate(-90deg)":"none",transition:"transform .2s",lineHeight:1,flexShrink:0}}>▼</button>
            {renaming&&canRename
              ? <input ref={nref} value={gname} onChange={e=>setGname(e.target.value)}
                  onBlur={()=>{setRenaming(false);if(gname.trim())onRenameGroup(gname.trim());else setGname(group.nome);}}
                  onKeyDown={e=>{if(e.key==="Enter"){setRenaming(false);if(gname.trim())onRenameGroup(gname.trim());}}}
                  style={{fontWeight:700,fontSize:14,color:group.color,border:"none",borderBottom:`2px solid ${group.color}`,
                    background:"transparent",outline:"none",minWidth:160}}/>
              : <span onDoubleClick={()=>canRename&&setRenaming(true)} title={canRename?"Clique duplo para renomear":""}
                  style={{fontWeight:700,fontSize:14,color:group.color,cursor:canRename?"pointer":"default"}}>{group.nome}</span>
            }
            <span style={{fontSize:12,color:"var(--text3)",background:"var(--surface3)",borderRadius:20,padding:"1px 8px",fontWeight:600}}>{items.length}</span>
            {group.owner_id&&(()=>{const owner=allUsers.find(u=>u.id===group.owner_id);return owner?<Avatar user={owner} size={22}/>:null;})()}
            <div style={{marginLeft:"auto",display:"flex",gap:7}}>
              {canRename&&<button onClick={()=>setRenaming(true)} title="Renomear" style={{...SB,color:"var(--text2)",borderColor:"var(--border)"}}>✎</button>}
              {canEdit&&<button onClick={onAddItem} style={{...SB,color:group.color,borderColor:group.color+"55"}}>+ Item</button>}
              {onGroupSettings&&<button onClick={onGroupSettings} style={{...SB,color:"var(--text2)",borderColor:"var(--border)"}}>⚙</button>}
              {(perms.all||perms.deleteAny)&&<button onClick={onDelGroup} style={{...SB,color:"#dc2626",borderColor:"#FEE2E2"}}>🗑️</button>}
            </div>
          </div>

          {/* Linha das colunas — mesma estrutura de colgroup do body */}
          {!group.collapsed&&(
            <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
              {cg}
              <tbody>
                <tr style={{background:"var(--surface2)",borderBottom:"2px solid var(--border)"}}>
                  <th style={{width:52,padding:"7px 6px",textAlign:"center",
                    borderLeft:`4px solid ${group.color}`,borderRight:"1px solid var(--border)"}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                      <span style={{fontSize:11,opacity:0,userSelect:"none"}}>📝</span>
                      <input type="checkbox" checked={allSel} onChange={()=>onSelectAll(items,!allSel)}
                        style={{cursor:"pointer",accentColor:"var(--blue)",width:13,height:13}}/>
                    </div>
                  </th>
                  {columns.map(col=>{
                    const isActive=sortCfg.colId===col.id;
                    const isAsc=isActive&&sortCfg.dir===1;
                    const isDesc=isActive&&sortCfg.dir===-1;
                    const cycleSort=()=>{
                      if(!isActive) setSortCfg({colId:col.id,dir:1});
                      else if(isAsc) setSortCfg({colId:col.id,dir:-1});
                      else setSortCfg({colId:null,dir:1});
                    };
                    return (
                      <th key={col.id} onClick={cycleSort}
                        style={{padding:"7px 11px",textAlign:"left",fontWeight:700,fontSize:11,
                          color:isActive?"var(--blue)":"var(--text3)",whiteSpace:"nowrap",
                          borderRight:"1px solid var(--border)",textTransform:"uppercase",letterSpacing:.4,
                          overflow:"hidden",textOverflow:"ellipsis",background:"var(--surface2)",
                          cursor:"pointer",userSelect:"none",transition:"color .15s"}}
                        onMouseEnter={e=>{if(!isActive)e.currentTarget.style.color="var(--text2)";}}
                        onMouseLeave={e=>{if(!isActive)e.currentTarget.style.color="var(--text3)";}}>
                        <div style={{display:"flex",alignItems:"center",gap:4}}>
                          <span style={{overflow:"hidden",textOverflow:"ellipsis"}}>{col.nome}</span>
                          <div style={{display:"flex",flexDirection:"column",gap:0,flexShrink:0,lineHeight:.8}}>
                            <span style={{fontSize:8,opacity:isAsc?1:.25,color:isActive?"var(--blue)":"var(--text3)"}}>▲</span>
                            <span style={{fontSize:8,opacity:isDesc?1:.25,color:isActive?"var(--blue)":"var(--text3)"}}>▼</span>
                          </div>
                        </div>
                      </th>
                    );
                  })}
                  <th style={{width:72,background:"var(--surface2)"}}/>
                </tr>
              </tbody>
            </table>
          )}
        </div>
      )}

      {/* ── Mobile: header simples ── */}
      {isMobile&&(
        <div style={{display:"flex",alignItems:"center",gap:9,padding:"10px 14px",
          background:"var(--surface)",borderLeft:`4px solid ${group.color}`,borderRadius:"11px 11px 0 0"}}>
          <button onClick={onToggle}
            style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:group.color,
              padding:0,transform:group.collapsed?"rotate(-90deg)":"none",transition:"transform .2s",lineHeight:1}}>▼</button>
          <span style={{fontWeight:700,fontSize:14,color:group.color}}>{group.nome}</span>
          <span style={{fontSize:12,color:"var(--text3)",background:"var(--surface3)",borderRadius:20,padding:"1px 8px",fontWeight:600}}>{items.length}</span>
          {group.owner_id&&(()=>{const owner=allUsers.find(u=>u.id===group.owner_id);return owner?<Avatar user={owner} size={22}/>:null;})()}
          <div style={{marginLeft:"auto",display:"flex",gap:7}}>
            {canEdit&&<button onClick={onAddItem} style={{...SB,color:group.color,borderColor:group.color+"55"}}>+ Item</button>}
            {(perms.all||perms.deleteAny)&&<button onClick={onDelGroup} style={{...SB,color:"#dc2626",borderColor:"#FEE2E2"}}>🗑️</button>}
          </div>
        </div>
      )}

      {!group.collapsed&&(
        isMobile
          ? <div style={{padding:"8px 0",background:"var(--bg)"}}>
              {items.map(item=>(
                <ItemCard key={item.id} item={item} columns={columns} gc={group.color} allUsers={allUsers}
                  selected={selectedItems.has(item.id)} onToggle={()=>onToggleItem(item.id)}
                  onOpen={()=>onOpenItem(item)} onDelete={()=>onDelItem(item.id)}/>
              ))}
              {!items.length&&<div style={{padding:"12px 16px",color:"var(--text3)",fontSize:13}}>Nenhum item</div>}
              <GroupTotals columns={columns} items={items} gc={group.color}/>
              <button onClick={onAddItem} style={{...T.btn,padding:"10px 16px",color:"var(--text3)",background:"none",width:"100%",justifyContent:"flex-start"}}>
                + Adicionar item
              </button>
            </div>

          // ── Desktop: body com scroll horizontal sincronizado com o header ──
          : <>
              <div ref={bodyRef} style={{overflowX:"auto",background:"var(--surface)"}} onScroll={syncHeader}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,tableLayout:"fixed"}}>
                  {cg}
                  <tbody>
                    {items.map((item)=>(
                      <ItemRow key={item.id} item={item} columns={columns} gc={group.color} allUsers={allUsers}
                        selected={selectedItems.has(item.id)}
                        onToggle={()=>onToggleItem(item.id)}
                        onOpen={()=>onOpenItem(item)}
                        onDelete={()=>onDelItem(item.id)}
                        onMoveInativa={onMoveInativa?()=>onMoveInativa(item):null}
                        onDupNeg={onDupNeg?()=>onDupNeg(item):null}
                        onSendToNeg={onSendToNeg?()=>onSendToNeg(item):null}
                        onSendToVendas={onSendToVendas?()=>onSendToVendas(item):null}
                        onDragStart={e=>onDragStart(e,item,group.id)}
                        onDragOver={e=>onItemDragOver(e,item.id,group.id)}
                        onDrop={e=>onItemDrop(e,item.id,group.id)}
                        onUpdateValue={(cid,v)=>onUpdateValue(item.id,cid,v)}
                        onRespChange={ids=>onRespChange(item.id,ids)}
                        sentToNegIds={sentToNegIds}
                      />
                    ))}
                  </tbody>
                </table>
                {!items.length&&<div style={{padding:"14px 56px",color:"var(--text3)",fontSize:13,borderLeft:`3px solid ${group.color}`}}>Nenhum item</div>}
              </div>
              <GroupTotals columns={columns} items={items} gc={group.color}/>
              <div style={{borderLeft:`3px solid ${group.color}`,background:"var(--surface)"}}>
                <button onClick={onAddItem} style={{display:"flex",alignItems:"center",gap:7,padding:"10px 18px",
                  background:"none",border:"none",cursor:"pointer",color:"var(--text3)",fontSize:13}}>
                  + Adicionar item
                </button>
              </div>
            </>
      )}
    </div>
  );
}
// ─────────────────────────────────────────────────────────────────────────────
function ItemPanel({item,board,allUsers,currentUser,onClose,onUpdateValue,onRespChange,onAddUpdate,onDelUpdate}) {
  const [submitting,setSubmitting]=useState(false);
  const {isMobile}=useBreakpoint();
  const firstCol=board?.columns?.[0];
  const title=item.values?.[firstCol?.id]||"Sem título";
  const handleUpdate=async html=>{setSubmitting(true);await onAddUpdate(html);setSubmitting(false);};
  const panelW=isMobile?"100vw":"500px";

  return createPortal(
    <div style={{position:"fixed",top:0,right:0,bottom:0,width:panelW,background:"var(--surface)",
      borderLeft:"1.5px solid var(--border)",display:"flex",flexDirection:"column",
      zIndex:1100,boxShadow:"-8px 0 40px var(--shadowMd)"}}>
      <div style={{padding:"17px 22px",borderBottom:"1px solid var(--border)",display:"flex",
        alignItems:"center",gap:13,background:"var(--surface2)",flexShrink:0}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:16,fontWeight:800,color:"var(--text)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{title}</div>
          <div style={{fontSize:12,color:"var(--text3)",marginTop:2}}>{board?.icon} {board?.nome}</div>
        </div>
        <button onClick={onClose}
          style={{background:"none",border:"1px solid var(--border)",borderRadius:7,fontSize:22,cursor:"pointer",
            color:"var(--text3)",width:35,height:35,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>×</button>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"18px 22px"}}>
        <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",marginBottom:16,textTransform:"uppercase",letterSpacing:.5}}>Campos</div>
        {board?.columns?.map(col=>(
          <div key={col.id} style={{display:"flex",alignItems:"flex-start",marginBottom:9,gap:12,paddingBottom:9,borderBottom:"1px solid var(--border)"}}>
            <div style={{width:130,fontSize:12,color:"var(--text3)",flexShrink:0,paddingTop:7,fontWeight:600}}>{col.nome}</div>
            <div style={{flex:1,minWidth:0}}>
              <Cell col={col} values={item.values} allColumns={board.columns}
                responsibles={item.responsibles} allUsers={allUsers}
                onChange={v=>onUpdateValue(col.id,v)} onRespChange={onRespChange}/>
            </div>
          </div>
        ))}
        <div style={{marginTop:28,paddingTop:24,borderTop:"1.5px solid var(--border)"}}>
          <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",marginBottom:16,textTransform:"uppercase",letterSpacing:.5}}>Atualizações</div>
          <RichEditor onSubmit={handleUpdate} disabled={submitting}/>
          <div style={{marginTop:20}}>
            {(item.updates||[]).map(u=>(
              <div key={u.id} style={{padding:"13px 15px",background:"var(--surface2)",borderRadius:9,marginBottom:10,borderLeft:"3px solid var(--alt)"}}>
                <div style={{fontSize:11,color:"var(--text3)",marginBottom:7,display:"flex",alignItems:"center",gap:8,justifyContent:"space-between"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <Avatar user={allUsers.find(x=>x.id===u.created_by)||currentUser} size={20}/>
                    {new Date(u.created_at).toLocaleString("pt-BR")}
                  </div>
                  <button onClick={()=>onDelUpdate(u.id)} style={{...T.iBtn,color:"#dc2626",borderColor:"#FEE2E2",fontSize:11}}>✕</button>
                </div>
                <div style={{fontSize:13,color:"var(--text)",lineHeight:1.7}} dangerouslySetInnerHTML={{__html:linkifyHTML(u.content)}}/>
              </div>
            ))}
            {!(item.updates||[]).length&&<div style={{fontSize:13,color:"var(--text3)",textAlign:"center",padding:"24px 0"}}>Nenhuma atualização</div>}
          </div>
        </div>
      </div>
    </div>,
    document.body
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// FILTER PANEL — status + responsável + intervalo de datas
// ─────────────────────────────────────────────────────────────────────────────
function FilterPanel({board,allUsers,filters,setFilters,onClose}) {
  const statusCols=board?.columns?.filter(c=>c.tipo==="status")||[];
  const [local,setLocal]=useState(()=>({...filters}));
  const allStatuses=useMemo(()=>{
    const seen=new Map();
    for(const col of statusCols){
      for(const opt of resolveOpts(col.config?.options||[])){
        if(!seen.has(opt.label)) seen.set(opt.label,opt);
      }
    }
    return [...seen.values()];
  },[statusCols]);
  const toggleF=(key,val)=>setLocal(f=>({...f,[key]:(f[key]||[]).includes(val)?(f[key]||[]).filter(x=>x!==val):[...(f[key]||[]),val]}));
  const hasAny=Object.values(local).some(v=>Array.isArray(v)?v.length>0:!!v);
  const {isMobile}=useBreakpoint();
  const apply=()=>{setFilters(local);onClose();};
  const clear=()=>setLocal({});

  return createPortal(
    <div style={{position:"fixed",top:0,right:0,bottom:0,width:isMobile?"100vw":"320px",
      background:"var(--surface)",borderLeft:"1.5px solid var(--border)",zIndex:1050,
      boxShadow:"-8px 0 40px var(--shadowMd)",display:"flex",flexDirection:"column",fontFamily:"system-ui"}}>
      <div style={{padding:"17px 22px",borderBottom:"1px solid var(--border)",display:"flex",
        alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <span style={{fontWeight:800,fontSize:16,color:"var(--text)"}}>Filtros</span>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          {hasAny&&<button onClick={clear}
            style={{fontSize:12,color:"#dc2626",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>Limpar tudo</button>}
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:24,color:"var(--text3)",lineHeight:1}}>×</button>
        </div>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"20px 22px"}}>
        {/* Intervalo de datas */}
        <div style={{...T.lbl,marginBottom:12}}>Intervalo de datas</div>
        <div style={{display:"flex",gap:10,marginBottom:24}}>
          <div style={{flex:1}}>
            <label style={{fontSize:11,color:"var(--text3)",display:"block",marginBottom:4}}>De</label>
            <input type="date" value={local.dateFrom||""} onChange={e=>setLocal(f=>({...f,dateFrom:e.target.value||null}))}
              style={{...T.inp,padding:"8px 10px",fontSize:12}}/>
          </div>
          <div style={{flex:1}}>
            <label style={{fontSize:11,color:"var(--text3)",display:"block",marginBottom:4}}>Até</label>
            <input type="date" value={local.dateTo||""} onChange={e=>setLocal(f=>({...f,dateTo:e.target.value||null}))}
              style={{...T.inp,padding:"8px 10px",fontSize:12}}/>
          </div>
        </div>

        {/* Filtro por mês */}
        <div style={{...T.lbl,marginBottom:12}}>Mês específico</div>
        <input type="month" value={local.month||""} onChange={e=>setLocal(f=>({...f,month:e.target.value||null}))}
          style={{...T.inp,marginBottom:24}}/>

        {/* Status */}
        {allStatuses.length>0&&<>
          <div style={{...T.lbl,marginBottom:12}}>Status</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:7,marginBottom:24}}>
            {allStatuses.map(opt=>{
              const active=(local.status||[]).includes(opt.label);
              return (
                <div key={opt.id||opt.label} onClick={()=>toggleF("status",opt.label)}
                  style={{cursor:"pointer",opacity:active?1:.45,transform:active?"scale(1.04)":"none",transition:"all .15s"}}>
                  <Badge value={opt.label} color={opt.color}/>
                </div>
              );
            })}
          </div>
        </>}

        {/* Responsável */}
        {allUsers.length>0&&<>
          <div style={{...T.lbl,marginBottom:12}}>Responsável</div>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {allUsers.map(u=>{
              const active=(local.resp||[]).includes(u.id);
              return (
                <div key={u.id} onClick={()=>toggleF("resp",u.id)}
                  style={{display:"flex",alignItems:"center",gap:11,padding:"10px 14px",borderRadius:9,cursor:"pointer",
                    border:`1.5px solid ${active?"var(--blue)":"var(--border)"}`,background:active?"var(--surface3)":"var(--surface2)",transition:"all .15s"}}>
                  <Avatar user={u} size={32}/>
                  <div>
                    <div style={{fontSize:13,fontWeight:active?700:400,color:"var(--text)"}}>{u.nome}</div>
                    <div style={{fontSize:11,color:"var(--text3)"}}>{u.funcao||""}</div>
                  </div>
                  {active&&<span style={{marginLeft:"auto",color:"var(--blue)",fontWeight:700}}>✓</span>}
                </div>
              );
            })}
          </div>
        </>}
      </div>
      {/* Botão Aplicar Filtros */}
      <div style={{padding:"14px 22px",borderTop:"1px solid var(--border)",display:"flex",gap:10,flexShrink:0}}>
        <button onClick={apply}
          style={{...T.btn,flex:1,background:"var(--blue)",color:"#fff",padding:"11px 0",fontSize:14,fontWeight:700,justifyContent:"center"}}>
          ✓ Aplicar Filtros
        </button>
        <button onClick={onClose}
          style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"11px 16px",fontSize:13,border:"1px solid var(--border)"}}>
          Cancelar
        </button>
      </div>
    </div>,
    document.body
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PARENT GROUP CONTAINER — "grupo mãe" do board Negociações
// ─────────────────────────────────────────────────────────────────────────────
function ParentGroupContainer({parentGroup,subGroups,columns,allUsers,selectedItems,isMobile,
  perms,currentUser,groupAccess,canManageParent,
  onToggleItem,onSelectAll,onAddItem,onDelItem,onOpenItem,
  onUpdateValue,onRespChange,onMoveInativa,onDupNeg,onSendToNeg,onSendToVendas,sentToNegIds=new Set(),
  sortCfg={colId:null,dir:1},setSortCfg=()=>{},
  onDragStart,onDragOver,onDrop,onItemDragOver,onItemDrop,onGroupSettings,
  onRenameSubGroup,onDelSubGroup,onEditParent,onDelParent}) {

  const [collapsed,setCollapsed]=useState(false);
  // Estado local de colapso por sub-grupo (id → bool)
  const [sgCollapsed,setSgCollapsed]=useState({});
  const toggleSg=id=>setSgCollapsed(p=>({...p,[id]:!p[id]}));
  const [renaming,setRenaming]=useState(false);
  const [pgName,setPgName]=useState(parentGroup.nome);
  const nref=useRef();
  const canRenameParent=["administrador","ceo","financeiro","gerente_comercial"].includes(perms?.slug)||perms?.all;
  useEffect(()=>{if(renaming)nref.current?.focus();},[renaming]);

  const totalItems=subGroups.reduce((s,g)=>g.items?.length?s+g.items.length:s,0);
  const owner=allUsers.find(u=>u.id===parentGroup.owner_id);
  const color=parentGroup.color||"#3145FF";

  return (
    <div style={{marginTop:20,borderRadius:13,overflow:"hidden",
      border:`2px solid ${color}30`,boxShadow:`0 2px 16px ${color}12`}}>

      {/* Header do grupo mãe */}
      <div style={{display:"flex",alignItems:"center",gap:10,padding:"12px 16px",
        background:`${color}12`,borderLeft:`5px solid ${color}`,cursor:"pointer"}}
        onClick={()=>!renaming&&setCollapsed(p=>!p)}>
        <button onClick={e=>{e.stopPropagation();setCollapsed(p=>!p);}}
          style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color,padding:0,
            transform:collapsed?"rotate(-90deg)":"none",transition:"transform .2s",lineHeight:1,flexShrink:0}}>▼</button>
        {renaming&&canRenameParent
          ? <input ref={nref} value={pgName} onChange={e=>setPgName(e.target.value)}
              onClick={e=>e.stopPropagation()}
              onBlur={()=>{setRenaming(false);if(pgName.trim()&&pgName!==parentGroup.nome)onEditParent?.({...parentGroup,nome:pgName.trim()});else setPgName(parentGroup.nome);}}
              onKeyDown={e=>{if(e.key==="Enter"){setRenaming(false);if(pgName.trim())onEditParent?.({...parentGroup,nome:pgName.trim()});}if(e.key==="Escape"){setRenaming(false);setPgName(parentGroup.nome);}}}
              style={{fontWeight:800,fontSize:15,color,border:"none",borderBottom:`2px solid ${color}`,
                background:"transparent",outline:"none",minWidth:180}} autoFocus/>
          : <span onDoubleClick={e=>{e.stopPropagation();canRenameParent&&setRenaming(true);}}
              style={{fontWeight:800,fontSize:15,color,flex:1}}>{parentGroup.nome}</span>
        }
        {owner&&<Avatar user={owner} size={24}/>}
        <span style={{fontSize:12,color:"var(--text3)",background:"var(--surface)",borderRadius:20,padding:"2px 10px",fontWeight:700,flexShrink:0}}>
          {totalItems} lead{totalItems!==1?"s":""}
        </span>
        {canManageParent&&<div style={{display:"flex",gap:6,marginLeft:"auto"}} onClick={e=>e.stopPropagation()}>
          {canRenameParent&&<button onClick={()=>setRenaming(true)} style={{...SB,color:"var(--text2)",borderColor:"var(--border)",fontSize:11}}>✎</button>}
          {canManageParent&&<button onClick={()=>onEditParent&&onEditParent(parentGroup)}
            style={{...SB,color:color,borderColor:color+"55",fontSize:11}}>⚙ Editar</button>}
          <button onClick={()=>onDelParent&&onDelParent(parentGroup.id)}
            style={{...SB,color:"#dc2626",borderColor:"#FEE2E2",fontSize:11}}>🗑️</button>
        </div>}
      </div>

      {/* Sub-grupos */}
      {!collapsed&&(
        <div style={{padding:"8px 12px 12px",background:"var(--bg)"}}>
          {subGroups.map(sg=>(
            <Group key={sg.id} group={{...sg,collapsed:!!sgCollapsed[sg.id]}} columns={columns}
              items={sg.items||[]} isDraggingOver={false}
              allUsers={allUsers} selectedItems={selectedItems} isMobile={isMobile}
              perms={perms} currentUser={currentUser}
              groupAccess={groupAccess[sg.id]||[]}
              onToggleItem={onToggleItem} onSelectAll={onSelectAll}
              onAddItem={()=>onAddItem(sg.id)}
              onDelGroup={()=>onDelSubGroup(sg.id)}
              onRenameGroup={n=>onRenameSubGroup(sg.id,n)}
              onToggle={()=>toggleSg(sg.id)}
              onOpenItem={item=>onOpenItem(item,sg.id)}
              onUpdateValue={(iid,cid,v)=>onUpdateValue(iid,sg.id,cid,v)}
              onRespChange={(iid,ids)=>onRespChange(iid,sg.id,ids)}
              onDelItem={iid=>onDelItem(sg.id,iid)}
              onMoveInativa={onMoveInativa?item=>onMoveInativa(sg.id,item):null}
              onDupNeg={onDupNeg?item=>onDupNeg(sg.id,item):null}
              onSendToNeg={onSendToNeg?item=>onSendToNeg(sg.id,item):null}
              onSendToVendas={onSendToVendas?item=>onSendToVendas(sg.id,item):null}
              sentToNegIds={sentToNegIds}
              sortCfg={sortCfg} setSortCfg={setSortCfg}
              onDragStart={onDragStart} onDragOver={e=>onDragOver(e,sg.id)} onDrop={e=>onDrop(e,sg.id)}
              onItemDragOver={onItemDragOver} onItemDrop={onItemDrop}
              onGroupSettings={null}/>
          ))}
          {!subGroups.length&&<div style={{padding:"12px 16px",color:"var(--text3)",fontSize:13}}>Nenhum sub-grupo</div>}
        </div>
      )}
    </div>
  );
}

// Modal para criar/editar grupo mãe de Negociações
function ParentGroupModal({initial,allUsers,onSave,onCancel}) {
  const CLOSER_ROLES=["closer"];
  const closers=allUsers.filter(u=>CLOSER_ROLES.includes(u.funcao?.toLowerCase())||CLOSER_ROLES.includes(u.role?.toLowerCase()));
  const [nome,setNome]=useState(initial?.nome||"");
  const [color,setColor]=useState(initial?.color||"#3145FF");
  const [ownerId,setOwnerId]=useState(initial?.owner_id||"");
  const isEdit=!!initial?.id;

  const PALETTE=["#3145FF","#059669","#d97706","#7c3aed","#0891b2","#dc2626","#e11d48","#0f766e","#1d4ed8","#b45309"];

  return (
    <Modal title={isEdit?"Editar Grupo Mãe":"+ Criar Grupo Mãe (Closer)"}
      onClose={onCancel}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"9px 20px",border:"1px solid var(--border)"}}>Cancelar</button>
        <button onClick={()=>nome.trim()&&onSave({...(initial?.id?{id:initial.id}:{}),nome:nome.trim(),color,owner_id:ownerId||null})}
          style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"9px 24px"}} disabled={!nome.trim()}>
          {isEdit?"Salvar":"Criar Grupo"}
        </button>
      </>}>
      <div style={{display:"flex",flexDirection:"column",gap:18}}>

        <div>
          <label style={T.lbl}>Nome do Grupo (ex: Closer - João)</label>
          <input autoFocus value={nome} onChange={e=>setNome(e.target.value)}
            style={{...T.inp}} placeholder="Ex: João Silva — Closer"/>
        </div>

        <div>
          <label style={T.lbl}>Closer Responsável</label>
          <select value={ownerId} onChange={e=>setOwnerId(e.target.value)} style={{...T.inp}}>
            <option value="">Selecionar closer…</option>
            {(allUsers).map(u=>(
              <option key={u.id} value={u.id}>{u.nome} {u.funcao?`(${u.funcao})`:""}</option>
            ))}
          </select>
          <div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>O closer selecionado terá acesso exclusivo a este grupo</div>
        </div>

        <div>
          <label style={T.lbl}>Cor do Grupo</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:6}}>
            {PALETTE.map(c=>(
              <div key={c} onClick={()=>setColor(c)}
                style={{width:30,height:30,borderRadius:8,background:c,cursor:"pointer",
                  outline:color===c?`3px solid ${c}`:"3px solid transparent",
                  outlineOffset:2,transition:"outline .15s"}}/>
            ))}
            <input type="color" value={color} onChange={e=>setColor(e.target.value)}
              style={{width:30,height:30,borderRadius:8,border:"1.5px solid var(--border)",cursor:"pointer",padding:2}}/>
          </div>
        </div>

        {!isEdit&&<div style={{background:"var(--surface2)",borderRadius:10,padding:"12px 14px",fontSize:12,color:"var(--text3)",borderLeft:"3px solid var(--blue)"}}>
          <strong style={{color:"var(--text)"}}>Ao criar, serão gerados automaticamente 4 sub-grupos:</strong>
          <ul style={{margin:"6px 0 0",paddingLeft:18,lineHeight:1.8}}>
            <li>Oportunidades Closer</li><li>Negociando</li>
            <li>Fechado/Ganho</li><li>Recusado/Negócio Futuro</li>
          </ul>
        </div>}

      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// BOARD VIEW
// ─────────────────────────────────────────────────────────────────────────────
function BoardView({boardId,boards,allUsers,currentUser,wsId,perms,onBoardCountChange}) {
  const toast=useToast();
  const {isMobile}=useBreakpoint();
  const [board,setBoard]=useState(null);
  const [loading,setLoading]=useState(true);
  const [selItem,setSelItem]=useState(null);
  const [search,setSearch]=useState("");
  const [filters,setFilters]=useState({});
  const [showFilters,setShowFilters]=useState(false);
  const [sortCfg,setSortCfg]=useState({colId:null,dir:1}); // dir: 1=asc, -1=desc
  const [showColMgr,setShowColMgr]=useState(false);
  const [showExport,setShowExport]=useState(false);
  const [selected,setSelected]=useState(new Set());
  const [dragState,setDragState]=useState(null);
  const [dragOverGroup,setDragOverGroup]=useState(null);
  const [confirmM,setConfirmM]=useState(null);
  const [groupSelM,setGroupSelM]=useState(null);
  const [sendToNegM,setSendToNegM]=useState(null);
  const [sentToNegIds,setSentToNegIds]=useState(new Set());
  // Permissões de grupo: map groupId → [userId, ...]
  const [groupAccess,setGroupAccess]=useState({});
  // Modal de criação/configurações de grupo
  const [groupCreateM,setGroupCreateM]=useState(false);
  const [groupSettingsM,setGroupSettingsM]=useState(null); // group obj
  const [moveItemsM,setMoveItemsM]=useState(false); // mover selecionados
  const [parentGroupM,setParentGroupM]=useState(null); // null=fechado, {}=criar, {id,...}=editar

  const loadBoard=useCallback(async bid=>{
    if(!bid){setBoard(null);setLoading(false);return;}
    setLoading(true);setSelItem(null);setSelected(new Set());
    const [{data:bData},{data:cols},{data:grps},{data:itens}]=await Promise.all([
      db.from("boards").select("*").eq("id",bid).single(),
      db.from("columns").select("*").eq("board_id",bid).order("ordem"),
      db.from("groups").select("*").eq("board_id",bid).order("ordem"),
      db.from("items").select("id,group_id,ordem,created_at").eq("board_id",bid).order("ordem"),
    ]);
    const ids=(itens||[]).map(i=>i.id);
    let vMap={},rMap={};
    if(ids.length){
      const [{data:vals},{data:resps}]=await Promise.all([
        db.from("item_values").select("*").in("item_id",ids),
        db.from("item_responsables").select("*").in("item_id",ids),
      ]);
      for(const v of vals||[]){if(!vMap[v.item_id])vMap[v.item_id]={};vMap[v.item_id][v.column_id]=v.value;}
      for(const r of resps||[]){if(!rMap[r.item_id])rMap[r.item_id]=[];rMap[r.item_id].push(r.user_id);}
    }
    // Carrega group_access para todos os grupos do board
    const gids=(grps||[]).map(g=>g.id);
    if(gids.length){
      const {data:accesses}=await db.from("group_access").select("*").in("group_id",gids);
      const aMap={};
      for(const a of accesses||[]){
        if(!aMap[a.group_id])aMap[a.group_id]=[];
        aMap[a.group_id].push(a.user_id);
      }
      setGroupAccess(aMap);
    }
    const columns=(cols||[]).map(c=>({...c,config:typeof c.config==="string"?JSON.parse(c.config):(c.config||{})}));
    const groups=(grps||[]).map(g=>({
      ...g,
      items:(itens||[]).filter(i=>i.group_id===g.id).map(i=>({...i,values:vMap[i.id]||{},responsibles:rMap[i.id]||[],updates:[]}))
    }));
    setBoard({...bData,columns,groups});

    // Detecta quais itens do Pré-Vendas já foram enviados para Negociações
    // via origin_item_id — vínculo estrutural permanente, imune a edições de campos
    if(bData?.nome==="Pré - Vendas"&&ids.length){
      const {data:linked}=await db.from("items")
        .select("origin_item_id")
        .in("origin_item_id",ids)
        .not("origin_item_id","is",null);
      setSentToNegIds(new Set((linked||[]).map(r=>r.origin_item_id)));
    }else{
      setSentToNegIds(new Set());
    }

    setLoading(false);
  },[]);

  useEffect(()=>{loadBoard(boardId);},[boardId,loadBoard]);

  const upd=fn=>setBoard(prev=>{if(!prev)return prev;const n=deep(prev);fn(n);return n;});
  const bump=(bid,d)=>onBoardCountChange(bid,d);

  // ── CRUD itens
  const addItem=async gid=>{
    const {data,error}=await db.from("items").insert({board_id:boardId,group_id:gid,ordem:9999}).select().single();
    if(error){toast("Erro ao criar item","error");return;}
    // Auto-preenche responsável com o owner do grupo
    const group=board?.groups.find(g=>g.id===gid);
    const ownerIds=group?.owner_id?[group.owner_id]:[];
    if(ownerIds.length){
      await db.from("item_responsables").insert(ownerIds.map(uid=>({item_id:data.id,user_id:uid})));
    }
    upd(b=>{b.groups.find(g=>g.id===gid)?.items.push({...data,values:{},responsibles:ownerIds,updates:[]});});
    bump(boardId,1);
  };

  const updateValue=async(iid,gid,cid,val)=>{
    upd(b=>{const item=b.groups.find(g=>g.id===gid)?.items.find(i=>i.id===iid);if(item)item.values[cid]=val;});
    if(selItem?.id===iid)setSelItem(p=>({...p,values:{...p.values,[cid]:val}}));
    await db.from("item_values").upsert({item_id:iid,column_id:cid,value:val},{onConflict:"item_id,column_id"});
  };

  const updateResp=async(iid,gid,ids)=>{
    upd(b=>{const item=b.groups.find(g=>g.id===gid)?.items.find(i=>i.id===iid);if(item)item.responsibles=ids;});
    if(selItem?.id===iid)setSelItem(p=>({...p,responsibles:ids}));
    await db.from("item_responsables").delete().eq("item_id",iid);
    if(ids.length)await db.from("item_responsables").insert(ids.map(uid=>({item_id:iid,user_id:uid})));
  };

  const delItem=async(gid,iid)=>{
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.items=g.items.filter(i=>i.id!==iid);});
    if(selItem?.id===iid)setSelItem(null);
    setSelected(s=>{const n=new Set(s);n.delete(iid);return n;});
    bump(boardId,-1);
    await db.from("items").delete().eq("id",iid);
    toast("Item excluído");
  };

  const bulkDelete=()=>{
    const ids=[...selected];
    setConfirmM({title:"Excluir itens",danger:true,message:`Excluir ${ids.length} item(s)?`,
      onConfirm:async()=>{
        upd(b=>{b.groups.forEach(g=>{g.items=g.items.filter(i=>!ids.includes(i.id));});});
        setSelected(new Set());bump(boardId,-ids.length);
        if(ids.includes(selItem?.id))setSelItem(null);
        await db.from("items").delete().in("id",ids);
        toast(`${ids.length} item(s) excluído(s)`);setConfirmM(null);
      }
    });
  };

  // Verifica se item tem todas as colunas preenchidas (para sendToVendas #5)
  const itemIsComplete=(item)=>{
    return board.columns.every(col=>{
      if(col.tipo==="calculated") return true; // calculado não precisa
      if(col.tipo==="user") return (item.responsibles||[]).length>0;
      const v=item.values?.[col.id];
      return v!=null&&v!=="";
    });
  };

  const getSelectedItems=()=>{
    const ids=[...selected];
    return board.groups.flatMap(g=>g.items.filter(i=>ids.includes(i.id)));
  };

  // ── CRUD grupos
  const addGroup=async({nome,color,ownerId})=>{
    const {data,error}=await db.from("groups")
      .insert({nome:nome.trim(),board_id:boardId,color,ordem:board?.groups.length||0,owner_id:ownerId||null})
      .select().single();
    if(error){toast("Erro ao criar grupo","error");return;}
    // Registra acesso do owner em group_access
    if(ownerId){
      await db.from("group_access").upsert({group_id:data.id,user_id:ownerId},{onConflict:"group_id,user_id",ignoreDuplicates:true});
    }
    const newAccess=ownerId?[ownerId]:[];
    setGroupAccess(p=>({...p,[data.id]:newAccess}));
    upd(b=>b.groups.push({...data,items:[]}));
    setGroupCreateM(false);
    toast("Grupo criado!");
  };
  const delGroup=gid=>{
    const qty=board?.groups.find(g=>g.id===gid)?.items.length||0;
    setConfirmM({title:"Excluir grupo",danger:true,message:`Excluir grupo e seus ${qty} item(s)?`,
      onConfirm:async()=>{
        upd(b=>{b.groups=b.groups.filter(g=>g.id!==gid);});bump(boardId,-qty);
        await db.from("groups").delete().eq("id",gid);
        toast("Grupo excluído");setConfirmM(null);
      }
    });
  };

  // Atualiza lista de usuários com acesso a um grupo
  const updateGroupAccess=async(gid,userIds)=>{
    // Delete all existing, then re-insert fresh — avoids constraint issues
    const {error:delErr}=await db.from("group_access").delete().eq("group_id",gid);
    if(delErr){toast("Erro ao limpar acessos: "+delErr.message,"error");return;}
    if(userIds.length){
      const rows=userIds.map(uid=>({group_id:gid,user_id:uid}));
      const {error:insErr}=await db.from("group_access").insert(rows);
      if(insErr){toast("Erro ao salvar acessos: "+insErr.message,"error");return;}
    }
    setGroupAccess(p=>({...p,[gid]:userIds}));
    toast("Acessos atualizados!");
  };

  // Atualiza owner de um grupo
  const updateGroupOwner=async(gid,ownerId)=>{
    await db.from("groups").update({owner_id:ownerId||null}).eq("id",gid);
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.owner_id=ownerId;});
    // Garante que o owner também está em group_access
    if(ownerId){
      const cur=groupAccess[gid]||[];
      if(!cur.includes(ownerId)) await updateGroupAccess(gid,[...cur,ownerId]);
    }
  };
  const renameGroup=async(gid,nome)=>{upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.nome=nome;});await db.from("groups").update({nome}).eq("id",gid);};
  const toggleGroup=gid=>upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.collapsed=!g.collapsed;});

  // ── CRUD grupos mãe (Negociações hierárquico)
  const SUB_GRUPOS_PADRAO=[
    {nome:"Oportunidades Closer",color:"#3B82F6"},
    {nome:"Negociando",color:"#F59E0B"},
    {nome:"Fechado/Ganho",color:"#10B981"},
    {nome:"Recusado/Negócio Futuro",color:"#EF4444"},
  ];

  const addParentGroup=async({nome,color,owner_id})=>{
    const ordem=(board?.groups||[]).filter(g=>g.is_parent).length;
    const{data:pg,error}=await db.from("groups")
      .insert({nome:nome.trim(),board_id:boardId,color,ordem,owner_id:owner_id||null,is_parent:true})
      .select().single();
    if(error){toast("Erro ao criar grupo mãe: "+error.message,"error");return;}

    // Registra acesso do owner no grupo mãe
    if(owner_id) await db.from("group_access").upsert({group_id:pg.id,user_id:owner_id},{onConflict:"group_id,user_id",ignoreDuplicates:true});

    // Cria os 4 sub-grupos padrão
    const sgRows=SUB_GRUPOS_PADRAO.map((sg,i)=>({
      nome:sg.nome,board_id:boardId,color:sg.color,
      ordem:i,owner_id:owner_id||null,parent_group_id:pg.id,is_parent:false
    }));
    const{data:sgs}=await db.from("groups").insert(sgRows).select();

    // Registra acesso do owner nos sub-grupos
    if(owner_id&&sgs?.length){
      const accRows=sgs.map(sg=>({group_id:sg.id,user_id:owner_id}));
      await db.from("group_access").insert(accRows).then(()=>{});
    }

    upd(b=>{
      b.groups.push({...pg,items:[],is_parent:true});
      (sgs||[]).forEach(sg=>b.groups.push({...sg,items:[]}));
    });
    setGroupAccess(prev=>{
      const n={...prev,[pg.id]:owner_id?[owner_id]:[]};
      (sgs||[]).forEach(sg=>{n[sg.id]=owner_id?[owner_id]:[];});
      return n;
    });
    setParentGroupM(null);
    toast("Grupo mãe criado com 4 sub-grupos!");
  };

  const editParentGroup=async(updated)=>{
    const{id,nome,color,owner_id}=updated;
    await db.from("groups").update({nome,color,owner_id:owner_id||null}).eq("id",id);
    // Atualiza owner nos sub-grupos também
    const subIds=(board?.groups||[]).filter(g=>g.parent_group_id===id).map(g=>g.id);
    if(subIds.length) await db.from("groups").update({owner_id:owner_id||null}).in("id",subIds);
    // Atualiza acessos
    if(owner_id){
      const allIds=[id,...subIds];
      await Promise.all(allIds.map(gid=>db.from("group_access").upsert({group_id:gid,user_id:owner_id},{onConflict:"group_id,user_id",ignoreDuplicates:true})));
    }
    upd(b=>{
      const pg=b.groups.find(g=>g.id===id);if(pg){pg.nome=nome;pg.color=color;pg.owner_id=owner_id||null;}
      b.groups.filter(g=>g.parent_group_id===id).forEach(sg=>{sg.owner_id=owner_id||null;});
    });
    setGroupAccess(prev=>{
      const n={...prev};
      [id,...subIds].forEach(gid=>{if(owner_id&&!n[gid]?.includes(owner_id))n[gid]=[...(n[gid]||[]),owner_id];});
      return n;
    });
    setParentGroupM(null);
    toast("Grupo mãe atualizado!");
  };

  const delParentGroup=pgid=>{
    const pg=board?.groups.find(g=>g.id===pgid);
    const subs=board?.groups.filter(g=>g.parent_group_id===pgid)||[];
    const totalItems=subs.reduce((s,g)=>s+g.items.length,0);
    setConfirmM({title:"Excluir grupo mãe",danger:true,
      message:`Excluir "${pg?.nome}" e todos os ${subs.length} sub-grupos com ${totalItems} lead(s)? Esta ação é irreversível.`,
      onConfirm:async()=>{
        upd(b=>{b.groups=b.groups.filter(g=>g.id!==pgid&&g.parent_group_id!==pgid);});
        bump(boardId,-totalItems);
        await db.from("groups").delete().eq("id",pgid); // CASCADE remove sub-grupos
        toast("Grupo mãe excluído");setConfirmM(null);
      }
    });
  };

  // ── Updates
  const loadUpdates=async(item,gid)=>{
    const {data}=await db.from("item_updates").select("*").eq("item_id",item.id).order("created_at",{ascending:false});
    setSelItem({...item,updates:data||[],_gid:gid});
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g){const idx=g.items.findIndex(i=>i.id===item.id);if(idx>=0)g.items[idx]={...g.items[idx],updates:data||[]};}});
  };
  const addUpdate=async(iid,gid,content)=>{
    const {data}=await db.from("item_updates").insert({item_id:iid,content,created_by:currentUser?.id}).select().single();
    if(!data)return;
    upd(b=>{const item=b.groups.find(g=>g.id===gid)?.items.find(i=>i.id===iid);if(item)item.updates.unshift(data);});
    setSelItem(p=>p?({...p,updates:[data,...(p.updates||[])]}):p);
    toast("Publicado");
  };
  const delUpdate=async(uid)=>{
    await db.from("item_updates").delete().eq("id",uid);
    setSelItem(p=>p?({...p,updates:(p.updates||[]).filter(u=>u.id!==uid)}):p);
    toast("Atualização removida");
  };

  // ── FIX: duplicate com campos equivalentes + updates
  const copyMatchingValues=async(srcItemId,srcCols,tgtItemId,tgtCols)=>{
    const {data:vals}=await db.from("item_values").select("*").eq("item_id",srcItemId);
    if(!vals?.length) return;
    // Mapa: nome_coluna → column_id no destino
    const tgtByName=Object.fromEntries(tgtCols.map(c=>[c.nome.toLowerCase().trim(),c.id]));
    const toInsert=[];
    for(const v of vals){
      const srcCol=srcCols.find(c=>c.id===v.column_id);
      if(!srcCol) continue;
      const tgtId=tgtByName[srcCol.nome.toLowerCase().trim()];
      if(tgtId) toInsert.push({item_id:tgtItemId,column_id:tgtId,value:v.value});
    }
    if(toInsert.length) await db.from("item_values").insert(toInsert);
  };

  const copyUpdates=async(srcItemId,tgtItemId)=>{
    const {data:upds}=await db.from("item_updates").select("*").eq("item_id",srcItemId).order("created_at");
    if(!upds?.length) return;
    await db.from("item_updates").insert(upds.map(u=>({item_id:tgtItemId,content:u.content,created_by:u.created_by})));
  };

  const moveInativa=async(gid,item)=>{
    const iBoard=boards.find(b=>b.nome==="Base Inativa");
    if(!iBoard){toast("Board 'Base Inativa' não encontrado","error");return;}
    const {data:tgs}=await db.from("groups").select("*").eq("board_id",iBoard.id).order("ordem");
    const isTraf=board?.nome==="Tráfego";
    const tg=isTraf?tgs?.find(g=>g.nome==="Tráfego"):tgs?.find(g=>g.nome==="Prospecção Fria")||tgs?.[0];
    if(!tg){toast("Grupo destino não encontrado","error");return;}
    // Cria novo item no destino
    const {data:newItem}=await db.from("items").insert({board_id:iBoard.id,group_id:tg.id,ordem:9999}).select().single();
    if(!newItem){toast("Erro ao mover","error");return;}
    // Copia campos equivalentes
    const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",iBoard.id);
    await copyMatchingValues(item.id,board.columns,newItem.id,tgtCols||[]);
    // Copia responsáveis
    if(item.responsibles?.length) await db.from("item_responsables").insert(item.responsibles.map(uid=>({item_id:newItem.id,user_id:uid})));
    // Copia atualizações
    await copyUpdates(item.id,newItem.id);
    // Exclui original
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.items=g.items.filter(i=>i.id!==item.id);});
    if(selItem?.id===item.id)setSelItem(null);
    bump(boardId,-1);bump(iBoard.id,1);
    await db.from("items").delete().eq("id",item.id);
    toast("Movido para Base Inativa (com dados e atualizações)");
    await logAct(currentUser?.id,wsId,"item",item.id,"moved",{to:"Base Inativa"});
  };

  const dupNeg=async(gid,item)=>{
    const nBoard=boards.find(b=>b.nome==="Negociações");
    if(!nBoard){toast("Board 'Negociações' não encontrado","error");return;}
    const {data:ngs}=await db.from("groups").select("*").eq("board_id",nBoard.id).order("ordem");
    if(!ngs?.length) return;
    setGroupSelM({title:"Grupo de destino em Negociações",groups:ngs,
      onSelect:async tg=>{
        const {data:ni}=await db.from("items").insert({board_id:nBoard.id,group_id:tg.id,ordem:9999}).select().single();
        if(!ni){toast("Erro ao duplicar","error");setGroupSelM(null);return;}
        // Campos equivalentes
        const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",nBoard.id);
        await copyMatchingValues(item.id,board.columns,ni.id,tgtCols||[]);
        // Responsáveis
        if(item.responsibles?.length) await db.from("item_responsables").insert(item.responsibles.map(uid=>({item_id:ni.id,user_id:uid})));
        // Atualizações
        await copyUpdates(item.id,ni.id);
        bump(nBoard.id,1);
        toast("Duplicado em Negociações (com dados e atualizações)");
        setGroupSelM(null);
        await logAct(currentUser?.id,wsId,"item",ni.id,"created",{via:"duplicate",from:item.id});
      }
    });
  };

  // ── Enviar para Negociações (Pré-Vendas → Negociações) — modal 2 etapas
  const sendToNeg=async(gid,item)=>{
    const negBoard=boards.find(b=>b.nome==="Negociações"||b.nome==="🤝 Negociações");
    if(!negBoard){toast("Quadro 'Negociações' não encontrado. Verifique o nome exato.","error");return;}
    const {data:ngs}=await db.from("groups").select("*").eq("board_id",negBoard.id).order("ordem");
    if(!ngs?.length){toast("Nenhum grupo em Negociações","error");return;}
    setSendToNegM({
      groups:ngs,
      onSelect:async tg=>{
        setSendToNegM(null);
        const {data:ni}=await db.from("items").insert({board_id:negBoard.id,group_id:tg.id,ordem:9999,origin_item_id:item.id}).select().single();
        if(!ni){toast("Erro ao criar item","error");return;}
        const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",negBoard.id);
        await copyMatchingValues(item.id,board.columns,ni.id,tgtCols||[]);
        if(item.responsibles?.length) await db.from("item_responsables").insert(item.responsibles.map(uid=>({item_id:ni.id,user_id:uid})));
        await copyUpdates(item.id,ni.id);
        bump(negBoard.id,1);
        await logAct(currentUser?.id,wsId,"item",ni.id,"created",{via:"sendToNeg",from:item.id});
        // Atualiza estado local imediatamente — na próxima carga o cruzamento confirma automaticamente
        setSentToNegIds(prev=>new Set([...prev,item.id]));
        toast("✅ Lead enviado para Negociações!");
      }
    });
  };

  // ── Enviar para Vendas (Negociações → Vendas, apenas cargos autorizados + campos completos)
  const sendToVendas=async(gid,item)=>{
    if(!perms.sendToVendas){toast("Sem permissão para enviar para Vendas","error");return;}
    if(!itemIsComplete(item)){
      toast("Preencha TODOS os campos do lead antes de enviar para Vendas","warning");return;
    }
    // Busca pelo nome exato do quadro Vendas
    // Normaliza nome: remove emojis/espaços/hifens e compara exato "vendas"
    const vendasBoard=boards.find(b=>{
      const n=b.nome.toLowerCase()
        .replace(/[^a-z0-9]/g,"")  // remove emojis, espaços, hifens, acentos ascii
        .normalize?.("NFD").replace(/[\u0300-\u036f]/g,"")||
        b.nome.toLowerCase().replace(/[^a-z0-9]/g,"");
      // Deve ser exatamente "vendas", NÃO "prvendas" ou "prevendas"
      return n==="vendas";
    });
    if(!vendasBoard){toast("Quadro 'Vendas' não encontrado. Crie um quadro com 'Vendas' no nome.","error");return;}
    const {data:vgs}=await db.from("groups").select("*").eq("board_id",vendasBoard.id).order("ordem");
    if(!vgs?.length){toast("Nenhum grupo em Vendas","error");return;}
    setGroupSelM({title:"📌 Enviar para Vendas — escolha o grupo",groups:vgs,
      onSelect:async tg=>{
        const {data:ni}=await db.from("items").insert({board_id:vendasBoard.id,group_id:tg.id,ordem:9999}).select().single();
        if(!ni){toast("Erro ao criar item em Vendas","error");setGroupSelM(null);return;}
        const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",vendasBoard.id);
        await copyMatchingValues(item.id,board.columns,ni.id,tgtCols||[]);
        if(item.responsibles?.length) await db.from("item_responsables").insert(item.responsibles.map(uid=>({item_id:ni.id,user_id:uid})));
        await copyUpdates(item.id,ni.id);
        bump(vendasBoard.id,1);
        toast("🏆 Lead enviado para Vendas!");
        setGroupSelM(null);
        await logAct(currentUser?.id,wsId,"item",ni.id,"created",{via:"sendToVendas",from:item.id});
        // ── Notificação por email
        await sendVendasEmail(item,tg);
      }
    });
  };

  // ── Email automático de notificação de venda via EmailJS
  const sendVendasEmail=async(item,grupo)=>{
    try{
      const campos=board.columns
        .filter(c=>c.tipo!=="calculated")
        .map(c=>{
          const v=c.tipo==="user"
            ? (item.responsibles||[]).map(id=>allUsers.find(u=>u.id===id)?.nome||id).join(", ")
            : item.values?.[c.id];
          return v?`${c.nome}: ${v}`:null;
        })
        .filter(Boolean)
        .join("\n");

      const vendedor=allUsers.find(u=>u.id===currentUser?.id);
      const dataHora=new Date().toLocaleString("pt-BR");
      const leadNome=item.values?.[board.columns[0]?.id]||"Lead";

      // Parâmetros para o template EmailJS
      // Configure em emailjs.com os campos: to_email, subject, vendedor, lead_nome, data_hora, grupo, campos, mensagem
      const emailParams = {
        to_email:   "yurifrancomonvatti@gmail.com",
        reply_to:   "yurifrancomonvatti@gmail.com",
        subject:    `✅ Nova Venda Registrada — ${leadNome}`,
        vendedor:   vendedor?.nome||"Sistema",
        lead_nome:  leadNome,
        data_hora:  dataHora,
        grupo:      grupo.nome,
        campos:     campos,
        mensagem:   [
          "Olá Yuri,",
          "",
          "Temos ótimas notícias! Uma nova venda acaba de ser fechada e registrada no Monvatti CRM.",
          "",
          "━━━━━━━━━━━━━━━━━━━━━━━━━━",
          "🏆 NOVA VENDA CONFIRMADA",
          "━━━━━━━━━━━━━━━━━━━━━━━━━━",
          "",
          `Data e hora: ${dataHora}`,
          `Registrado por: ${vendedor?.nome||"Sistema"}`,
          `Grupo de destino: ${grupo.nome}`,
          "",
          "📋 INFORMAÇÕES DO LEAD:",
          campos,
          "",
          "━━━━━━━━━━━━━━━━━━━━━━━━━━",
          "Esta mensagem foi gerada automaticamente pelo Monvatti CRM.",
        ].join("\n"),
      };

      // Mensagem WhatsApp — compacta e legível
      const waMensagem = [
        "🏆 *NOVA VENDA — Monvatti CRM*",
        "",
        `📋 *Lead:* ${leadNome}`,
        `👤 *Registrado por:* ${vendedor?.nome||"Sistema"}`,
        `📅 *Data/hora:* ${dataHora}`,
        `📂 *Grupo:* ${grupo.nome}`,
        "",
        "📄 *Detalhes:*",
        campos,
        "",
        "─────────────────────────",
        "Monvatti CRM",
      ].join("\n");

      // Dispara email e WhatsApp em paralelo
      const [emailResult, waResult] = await Promise.all([
        EMAILJS_KEY!=="YOUR_PUBLIC_KEY" ? sendEmailJS(emailParams) : Promise.resolve({ok:false}),
        sendWhatsApp(waMensagem),
      ]);

      const notices = [];
      if(emailResult.ok) notices.push("📧 E-mail");
      if(waResult.ok)    notices.push(`💬 WhatsApp (${waResult.enviados?.join(", ")})`);
      if(notices.length){ toast(notices.join(" + ") + " enviado(s)!","success"); return; }

      // Fallback mailto se nenhuma integração estiver ativa
      const subject=encodeURIComponent(emailParams.subject);
      const body=encodeURIComponent(emailParams.mensagem);
      window.open("mailto:yurifrancomonvatti@gmail.com?subject="+subject+"&body="+body,"_blank");
      toast("Configure EMAILJS_KEY ou WA_DESTINATARIOS para notificações automáticas.","warning");
    }catch(err){
      console.error("Erro ao enviar email:",err);
      toast("Erro ao enviar e-mail de notificação","error");
    }
  };

  // ── Notificação em lote — agrupa múltiplos leads em 1 única mensagem
  const sendVendasEmailBulk=async(items,grupo)=>{
    if(!items?.length) return;
    if(items.length===1){ await sendVendasEmail(items[0],grupo); return; }
    try{
      const nl="\n";
      const sep="\n\n\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n\n";
      const vendedor=allUsers.find(u=>u.id===currentUser?.id);
      const dataHora=new Date().toLocaleString("pt-BR");
      const resumos=items.map((item,idx)=>{
        const leadNome=item.values?.[board.columns[0]?.id]||("Lead "+(idx+1));
        const campos=board.columns
          .filter(c=>c.tipo!=="calculated"&&c.tipo!=="user")
          .map(c=>{ const v=item.values?.[c.id]; return v?("  "+c.nome+": "+v):null; })
          .filter(Boolean).join(nl);
        return ("*"+(idx+1)+". "+leadNome+"*"+nl+campos);
      }).join(sep);
      const waMensagem=[
        "\uD83C\uDFC6 *"+items.length+" NOVAS VENDAS \u2014 Monvatti CRM*",
        "",
        "\uD83D\uDC64 *Registrado por:* "+(vendedor?.nome||"Sistema"),
        "\uD83D\uDCC5 *Data/hora:* "+dataHora,
        "\uD83D\uDCC2 *Grupo:* "+grupo.nome,
        "",
        "\uD83D\uDCC4 *Leads:*",
        "",
        resumos,
        "",
        "\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500",
        "Monvatti CRM",
      ].join(nl);
      const emailMensagem=[
        "Ol\u00E1 Yuri,","",
        items.length+" novas vendas foram registradas no Monvatti CRM.","",
        "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501",
        "\uD83C\uDFC6 "+items.length+" VENDAS CONFIRMADAS",
        "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501","",
        "Data e hora: "+dataHora,
        "Registrado por: "+(vendedor?.nome||"Sistema"),
        "Grupo: "+grupo.nome,"",
        resumos,"",
        "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501",
        "Monvatti CRM",
      ].join(nl);
      const emailParams={
        to_email:"yurifrancomonvatti@gmail.com",
        reply_to:"yurifrancomonvatti@gmail.com",
        subject:"\u2705 "+items.length+" Novas Vendas \u2014 "+dataHora,
        vendedor:vendedor?.nome||"Sistema",
        lead_nome:items.length+" leads",
        data_hora:dataHora,
        grupo:grupo.nome,
        campos:resumos,
        mensagem:emailMensagem,
      };
      const [emailResult,waResult]=await Promise.all([
        EMAILJS_KEY!=="YOUR_PUBLIC_KEY"?sendEmailJS(emailParams):Promise.resolve({ok:false}),
        sendWhatsApp(waMensagem),
      ]);
      const notices=[];
      if(emailResult.ok) notices.push("\uD83D\uDCE7 E-mail");
      if(waResult.ok)    notices.push("\uD83D\uDCAC WhatsApp ("+(waResult.enviados?.join(", "))+")");
      if(notices.length) toast(notices.join(" + ")+" enviado(s)!","success");
    }catch(err){
      console.error("Erro ao enviar notificações em lote:",err);
    }
  };

    // ── Drag & drop com persistência de ordem
  const handleDragStart=(e,item,gid)=>{setDragState({itemId:item.id,srcGroupId:gid,item});e.dataTransfer.effectAllowed="move";};
  const handleGroupDragOver=(e,gid)=>{e.preventDefault();setDragOverGroup(gid);};
  const handleGroupDrop=async(e,tgtGid)=>{
    e.preventDefault();
    if(!dragState||dragState.srcGroupId===tgtGid){setDragState(null);setDragOverGroup(null);return;}
    const {srcGroupId,item}=dragState;
    upd(b=>{
      const src=b.groups.find(g=>g.id===srcGroupId);
      const tgt=b.groups.find(g=>g.id===tgtGid);
      if(src&&tgt){src.items=src.items.filter(i=>i.id!==item.id);tgt.items.push({...item,group_id:tgtGid});}
    });
    setDragState(null);setDragOverGroup(null);
    await db.from("items").update({group_id:tgtGid,ordem:9999}).eq("id",item.id);
  };
  const handleItemDragOver=(e,overItemId,gid)=>{e.preventDefault();setDragState(p=>p?({...p,targetItemId:overItemId,targetGroupId:gid}):p);};
  const handleItemDrop=async(e,overItemId,gid)=>{
    e.preventDefault();
    if(!dragState) return;
    if(dragState.srcGroupId===gid&&overItemId!==dragState.itemId){
      let sorted=null;
      upd(b=>{
        const g=b.groups.find(g=>g.id===gid); if(!g) return;
        const from=g.items.findIndex(i=>i.id===dragState.itemId);
        const to=g.items.findIndex(i=>i.id===overItemId);
        if(from<0||to<0) return;
        const arr=[...g.items];arr.splice(to,0,arr.splice(from,1)[0]);
        g.items=arr.map((it,idx)=>({...it,ordem:idx}));
        sorted=g.items;
      });
      if(sorted) await Promise.all(sorted.map(it=>db.from("items").update({ordem:it.ordem}).eq("id",it.id)));
    }
    setDragState(null);setDragOverGroup(null);
  };

  // ── Filtros combinados
  const sortItems=(items)=>{
    const {colId,dir}=sortCfg;
    if(!colId) return items;
    return [...items].sort((a,b)=>{
      // Ordenação por responsável
      if(colId==="__resp__"){
        const nameA=(allUsers.find(u=>a.responsibles?.[0]===u.id)?.nome||"").toLowerCase();
        const nameB=(allUsers.find(u=>b.responsibles?.[0]===u.id)?.nome||"").toLowerCase();
        return dir*(nameA<nameB?-1:nameA>nameB?1:0);
      }
      const vA=a.values?.[colId]; const vB=b.values?.[colId];
      // Datas
      if(vA&&vB&&String(vA).match(/^\d{4}-\d{2}-\d{2}/))
        return dir*(new Date(vA)-new Date(vB));
      // Números e moeda
      const nA=parseFloat(vA); const nB=parseFloat(vB);
      if(!isNaN(nA)&&!isNaN(nB)) return dir*(nA-nB);
      // Texto e status
      const sA=String(vA||"").toLowerCase(); const sB=String(vB||"").toLowerCase();
      return dir*(sA<sB?-1:sA>sB?1:0);
    });
  };
  const applyFilters=items=>sortItems(applyFiltersRaw(items));
  const applyFiltersRaw=items=>{
    let r=items;
    if(search.trim()){
      const q=search.toLowerCase();
      r=r.filter(i=>Object.values(i.values||{}).some(v=>v&&String(v).toLowerCase().includes(q))||allUsers.filter(u=>i.responsibles?.includes(u.id)).some(u=>u.nome.toLowerCase().includes(q)));
    }
    if(filters.status?.length){
      const sCols=board?.columns?.filter(c=>c.tipo==="status")||[];
      // Compara com o label armazenado (string) — compatível com formato antigo e novo
      r=r.filter(i=>filters.status.some(label=>sCols.some(c=>i.values?.[c.id]===label)));
    }
    if(filters.resp?.length) r=r.filter(i=>filters.resp.some(uid=>i.responsibles?.includes(uid)));
    // Filtro por intervalo de datas
    if(filters.dateFrom||filters.dateTo||filters.month){
      const dateCols=board?.columns?.filter(c=>c.tipo==="date")||[];
      r=r.filter(i=>{
        for(const col of dateCols){
          const v=i.values?.[col.id]; if(!v) continue;
          const d=new Date(v);
          if(filters.dateFrom&&d<new Date(filters.dateFrom)) return false;
          if(filters.dateTo&&d>new Date(filters.dateTo)) return false;
          if(filters.month){
            const [y,m]=filters.month.split("-").map(Number);
            if(d.getFullYear()!==y||d.getMonth()+1!==m) return false;
          }
        }
        return true;
      });
    }
    return r;
  };

  const toggleItem=id=>setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const selectAll=(items,on)=>setSelected(s=>{const n=new Set(s);items.forEach(i=>on?n.add(i.id):n.delete(i.id));return n;});
  // Botões contextuais por quadro
  const isPreVendas      = board?.nome==="Pré - Vendas";     // botão → Negociações (todos)
  const isNegociacoes    = board?.nome==="Negociações";       // botão → Vendas (sendToVendas)
  const canActions       = isPreVendas;                       // mover inativa habilitado no Pré-Vendas

  // ── Hierarquia de grupos
  const canManageParent=perms.all||perms.manageBoards||["administrador","ceo","gerente_comercial"].includes(perms.slug);

  // Filtra grupos conforme permissão do usuário
  const visibleGroups=useMemo(()=>{
    if(!board) return [];

    if(!isNegociacoes){
      let gs=board.groups;
      if(perms.slug==="sdr") gs=gs.filter(g=>g.owner_id===currentUser?.id);
      else if(perms.slug==="closer") gs=gs.filter(g=>(groupAccess[g.id]||[]).includes(currentUser?.id));
      // Aplica filtro/busca nos itens — grupos com 0 resultados ficam com items=[]
      return gs.map(g=>({...g,items:applyFilters(g.items||[])}));
    }

    // Para Negociações: retorna grupos mãe com seus sub-grupos embutidos
    const allGroups=board.groups;
    const parentGroups=allGroups.filter(g=>g.is_parent);
    const subGroups=allGroups.filter(g=>!!g.parent_group_id);
    const legacyGroups=allGroups.filter(g=>!g.is_parent&&!g.parent_group_id); // grupos antigos sem pai

    // Enriquece grupos mãe com sub-grupos
    let enriched=parentGroups.map(pg=>({
      ...pg,
      _isParent:true,
      subGroups:subGroups.filter(sg=>sg.parent_group_id===pg.id)
        .sort((a,b)=>a.ordem-b.ordem)
        .map(sg=>({...sg,items:applyFilters(sg.items||[])})),
    }));

    // Closer: vê apenas o seu grupo mãe (pelo owner_id)
    if(perms.slug==="closer"){
      enriched=enriched.filter(pg=>pg.owner_id===currentUser?.id||(groupAccess[pg.id]||[]).includes(currentUser?.id));
    }

    // Grupos legados (antes da migração) ficam visíveis normalmente
    const legacyFiltered=legacyGroups.filter(g=>{
      if(perms.all||perms.viewAll) return true;
      if(perms.slug==="sdr") return g.owner_id===currentUser?.id;
      if(perms.slug==="closer") return (groupAccess[g.id]||[]).includes(currentUser?.id);
      return true;
    });

    return [...enriched,...legacyFiltered.map(g=>({...g,items:applyFilters(g.items||[])}))];
  },[board,perms,currentUser,groupAccess,isNegociacoes,filters,search,sortCfg]);
  const filterCount=(filters.status?.length||0)+(filters.resp?.length||0)+(filters.dateFrom?1:0)+(filters.dateTo?1:0)+(filters.month?1:0)+(search.trim()?1:0);

  if(loading) return <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}><Spinner size={40}/><span style={{color:"var(--text3)",fontSize:14}}>Carregando…</span></div>;
  if(!board) return <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",color:"var(--text3)",fontSize:14}}>Selecione um quadro</div>;

  return (
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
      {/* Top bar */}
      <div style={{background:"var(--surface)",borderBottom:"1.5px solid var(--border)",
        padding:isMobile?"10px 12px":"0 16px",
        display:"flex",alignItems:"center",gap:8,minHeight:58,flexShrink:0,
        flexWrap:isMobile?"wrap":"nowrap"}}>
        <div style={{fontSize:isMobile?15:17,fontWeight:900,color:"var(--text)",letterSpacing:-.4,
          flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
          {board.icon||"📋"} {board.nome}
        </div>
        {/* Em mobile: linha 1 tem título + botões ícone compactos */}
        {isMobile&&(
          <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}>
            <button onClick={()=>setShowFilters(p=>!p)}
              style={{...T.btn,background:filterCount>0?"var(--blue)":"var(--surface3)",color:filterCount>0?"#fff":"var(--text2)",padding:"7px 10px",fontSize:13,minWidth:0}}>
              {filterCount>0?`⚙ (${filterCount})`:"⚙"}
            </button>
            {perms.export&&<button onClick={()=>setShowExport(true)} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"7px 10px",fontSize:13}}>📥</button>}
            {perms.createGroups&&<button onClick={()=>setGroupCreateM(true)} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"7px 12px",fontSize:12}}>+ Grupo</button>}
          </div>
        )}
        {/* Em desktop: botões normais */}
        {!isMobile&&<>
          <input placeholder="🔍 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}
            style={{border:`1.5px solid ${search?"var(--blue)":"var(--border)"}`,borderRadius:8,padding:"7px 12px",
              fontSize:13,width:190,outline:"none",background:"var(--surface2)",color:"var(--text)"}}/>
          <button onClick={()=>setShowFilters(p=>!p)}
            style={{...T.btn,background:filterCount>0?"var(--blue)":"var(--surface3)",color:filterCount>0?"#fff":"var(--text2)",padding:"8px 13px",fontSize:12}}>
            ⚙ {filterCount>0?`(${filterCount})`:"Filtros"}
          </button>
                    {perms.manageCols&&<button onClick={()=>setShowColMgr(true)} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"8px 13px",fontSize:12}}>⊞ Colunas</button>}
          {perms.export&&<button onClick={()=>setShowExport(true)} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"8px 13px",fontSize:12}}>📥 Exportar</button>}
          {perms.createGroups&&<button onClick={()=>setGroupCreateM(true)} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 16px",fontSize:12}}>+ Grupo</button>}
        </>}
        {/* Em mobile: linha 2 com busca full-width */}
        {isMobile&&<div style={{width:"100%",paddingTop:4}}>
          <input placeholder="🔍 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}
            style={{...T.inp,padding:"8px 12px",fontSize:13}}/>
        </div>}
      </div>

      {/* Content */}
      <div style={{flex:1,overflowY:"auto",padding:isMobile?"8px 12px 120px":"8px 22px 120px"}}>
        {visibleGroups
          .map(group=>{
            // Quando há busca ou filtro ativos, remove sub-grupos vazios e grupos mãe completamente vazios
            if(filterCount>0&&group._isParent){
              const nonEmptySubs=(group.subGroups||[]).filter(sg=>(sg.items||[]).length>0);
              if(!nonEmptySubs.length) return null;
              return {...group,subGroups:nonEmptySubs};
            }
            if(filterCount>0&&!group._isParent&&(group.items||[]).length===0) return null;
            return group;
          })
          .filter(Boolean)
          .map(group=>(
          group._isParent
            ? <ParentGroupContainer key={group.id}
                parentGroup={group}
                subGroups={group.subGroups||[]}
                columns={board.columns}
                allUsers={allUsers}
                selectedItems={selected}
                isMobile={isMobile}
                perms={perms}
                currentUser={currentUser}
                groupAccess={groupAccess}
                canManageParent={canManageParent}
                onToggleItem={toggleItem}
                onSelectAll={selectAll}
                onAddItem={gid=>addItem(gid)}
                onDelItem={(gid,iid)=>setConfirmM({title:"Excluir item",danger:true,message:"Excluir permanentemente?",onConfirm:()=>{delItem(gid,iid);setConfirmM(null);}})}
                onOpenItem={(item,gid)=>loadUpdates(item,gid)}
                onUpdateValue={(iid,gid,cid,v)=>updateValue(iid,gid,cid,v)}
                onRespChange={(iid,gid,ids)=>updateResp(iid,gid,ids)}
                onMoveInativa={canActions?(gid,item)=>moveInativa(gid,item):null}
                onDupNeg={null}
                onSendToNeg={isPreVendas?(gid,item)=>sendToNeg(gid,item):null}
                onSendToVendas={isNegociacoes&&perms.sendToVendas?(gid,item)=>sendToVendas(gid,item):null}
                sentToNegIds={sentToNegIds}
                sortCfg={sortCfg} setSortCfg={setSortCfg}
                onDragStart={handleDragStart}
                onDragOver={(e,gid)=>handleGroupDragOver(e,gid)}
                onDrop={(e,gid)=>handleGroupDrop(e,gid)}
                onItemDragOver={handleItemDragOver}
                onItemDrop={handleItemDrop}
                onRenameSubGroup={(gid,n)=>renameGroup(gid,n)}
                onDelSubGroup={gid=>delGroup(gid)}
                onEditParent={pg=>setParentGroupM(pg)}
                onDelParent={pgid=>delParentGroup(pgid)}/>
            : <Group key={group.id} group={group} columns={board.columns}
                items={group.items}
                isDraggingOver={dragOverGroup===group.id}
                allUsers={allUsers} selectedItems={selected} isMobile={isMobile}
                perms={perms} currentUser={currentUser}
                groupAccess={groupAccess[group.id]||[]}
                onToggleItem={toggleItem} onSelectAll={selectAll}
                onAddItem={()=>addItem(group.id)} onDelGroup={()=>delGroup(group.id)}
                onRenameGroup={n=>renameGroup(group.id,n)} onToggle={()=>toggleGroup(group.id)}
                onOpenItem={item=>loadUpdates(item,group.id)}
                onUpdateValue={(iid,cid,v)=>updateValue(iid,group.id,cid,v)}
                onRespChange={(iid,ids)=>updateResp(iid,group.id,ids)}
                onDelItem={iid=>setConfirmM({title:"Excluir item",danger:true,message:"Excluir permanentemente?",onConfirm:()=>{delItem(group.id,iid);setConfirmM(null);}})}
                onMoveInativa={canActions?item=>moveInativa(group.id,item):null}
                onDupNeg={null}
                onSendToNeg={isPreVendas?item=>sendToNeg(group.id,item):null}
                onSendToVendas={isNegociacoes&&perms.sendToVendas?item=>sendToVendas(group.id,item):null}
                sentToNegIds={sentToNegIds}
                sortCfg={sortCfg} setSortCfg={setSortCfg}
                onDragStart={handleDragStart} onDragOver={e=>handleGroupDragOver(e,group.id)} onDrop={e=>handleGroupDrop(e,group.id)}
                onItemDragOver={handleItemDragOver} onItemDrop={handleItemDrop}
                onGroupSettings={perms.isFull||perms.all?()=>setGroupSettingsM(group):null}
              />
        ))}
        {perms.createGroups&&(
          <div style={{marginTop:18,display:"flex",gap:10,flexWrap:"wrap"}}>
            {isNegociacoes&&canManageParent&&(
              <button onClick={()=>setParentGroupM({})}
                style={{display:"flex",alignItems:"center",gap:9,background:"none",
                  border:"1.5px dashed var(--blue)",borderRadius:9,padding:"11px 24px",cursor:"pointer",
                  color:"var(--blue)",fontSize:13,fontWeight:700}}>
                + Grupo Mãe (Closer)
              </button>
            )}
            {!isNegociacoes&&<button onClick={()=>setGroupCreateM(true)}
              style={{display:"flex",alignItems:"center",gap:9,background:"none",
                border:"1.5px dashed var(--border)",borderRadius:9,padding:"11px 24px",cursor:"pointer",
                color:"var(--text3)",fontSize:13}}>
              + Adicionar grupo
            </button>}
          </div>
        )}
      </div>

      {selItem&&board&&(
        <ItemPanel item={selItem} board={board} allUsers={allUsers} currentUser={currentUser}
          onClose={()=>setSelItem(null)}
          onUpdateValue={(cid,v)=>updateValue(selItem.id,selItem._gid,cid,v)}
          onRespChange={ids=>updateResp(selItem.id,selItem._gid,ids)}
          onAddUpdate={html=>addUpdate(selItem.id,selItem._gid,html)}
          onDelUpdate={uid=>delUpdate(uid)}
        />
      )}
      {showFilters&&<FilterPanel board={board} allUsers={allUsers} filters={filters} setFilters={setFilters} onClose={()=>setShowFilters(false)}/>}
      {showColMgr&&<ColumnManagerModal board={board} toast={toast} onClose={()=>setShowColMgr(false)} onRefresh={()=>loadBoard(boardId)}/>}
      {showExport&&<ExportModal board={{...board,groups:board.groups.map(g=>({...g,items:applyFilters(g.items)}))}} onClose={()=>setShowExport(false)}/>}
      {parentGroupM!==null&&(
        <ParentGroupModal
          initial={parentGroupM?.id?parentGroupM:null}
          allUsers={allUsers}
          onSave={parentGroupM?.id?editParentGroup:addParentGroup}
          onCancel={()=>setParentGroupM(null)}/>
      )}

      {selected.size>0&&createPortal(
        <div style={{position:"fixed",bottom:28,left:"50%",transform:"translateX(-50%)",background:"var(--sidebar)",
          color:"#fff",borderRadius:14,padding:"13px 20px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",
          boxShadow:"0 12px 48px rgba(0,0,0,.36)",zIndex:1500,whiteSpace:"nowrap",maxWidth:"96vw"}}>
          <span style={{fontSize:14,fontWeight:700}}>{selected.size} item(s)</span>
          <button onClick={()=>setMoveItemsM(true)}
            style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 14px",fontSize:12}}>
            📦 Mover para quadro
          </button>
          <button onClick={bulkDelete}
            style={{...T.btn,background:"#dc2626",color:"#fff",padding:"8px 14px",fontSize:12}}>
            🗑️ Excluir
          </button>
          <button onClick={()=>setSelected(new Set())}
            style={{...T.btn,background:"rgba(255,255,255,.15)",color:"#fff",border:"1px solid rgba(255,255,255,.3)",padding:"8px 12px",fontSize:12}}>
            ✕
          </button>
        </div>,
        document.body
      )}

      {confirmM&&<ConfirmModal title={confirmM.title} message={confirmM.message} danger={confirmM.danger} onConfirm={confirmM.onConfirm} onCancel={()=>setConfirmM(null)}/>}
      {groupSelM&&<GroupSelectorModal title={groupSelM.title} groups={groupSelM.groups} onSelect={groupSelM.onSelect} onCancel={()=>setGroupSelM(null)}/>}
      {sendToNegM&&<SendToNegModal groups={sendToNegM.groups} onSelect={sendToNegM.onSelect} onCancel={()=>setSendToNegM(null)}/>}
      {groupCreateM&&<GroupCreateModal allUsers={allUsers} onSave={addGroup} onCancel={()=>setGroupCreateM(false)}/>}
      {groupSettingsM&&<GroupSettingsModal group={groupSettingsM} allUsers={allUsers}
        currentAccess={groupAccess[groupSettingsM.id]||[]}
        onSaveAccess={ids=>updateGroupAccess(groupSettingsM.id,ids)}
        onSaveOwner={uid=>updateGroupOwner(groupSettingsM.id,uid)}
        onClose={()=>setGroupSettingsM(null)}/>}
      {moveItemsM&&board&&<MoveItemsModal
        items={getSelectedItems()}
        srcBoard={board}
        allBoards={boards}
        onMove={async(destBid,movedItemIds,destGroup)=>{
          const cnt=[...selected].length;
          setMoveItemsM(false);
          // Se destino é Negociações, atualiza badges localmente
          if(movedItemIds?.length){
            setSentToNegIds(prev=>new Set([...prev,...movedItemIds]));
          }
          // Se destino é Vendas, dispara notificação agrupada (1 msg para N leads)
          if(destGroup){
            const movedItems=getSelectedItems();
            await sendVendasEmailBulk(movedItems,destGroup);
          }
          setSelected(new Set());
          bump(destBid,cnt);
          toast(`${cnt} item(s) copiado(s) com sucesso!`);
        }}
        onCancel={()=>setMoveItemsM(false)}/>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// GROUP CREATE MODAL — nome + cor + responsável
// ─────────────────────────────────────────────────────────────────────────────
function GroupCreateModal({allUsers,onSave,onCancel}) {
  const [nome,setNome]=useState("");
  const [color,setColor]=useState(GC[0]);
  const [ownerId,setOwnerId]=useState("");

  const save=()=>{
    if(!nome.trim()) return;
    onSave({nome:nome.trim(),color,ownerId:ownerId||null});
  };

  return (
    <Modal title="Novo grupo" onClose={onCancel} width={460}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={save} disabled={!nome.trim()} style={{...T.btn,background:"var(--blue)",color:"#fff",opacity:nome.trim()?1:.5}}>Criar grupo</button>
      </>}>
      <label style={T.lbl}>Nome do grupo</label>
      <input value={nome} onChange={e=>setNome(e.target.value)} autoFocus
        onKeyDown={e=>{if(e.key==="Enter")save();}}
        style={{...T.inp,marginBottom:22}} placeholder="Ex: SDR - João, Time Brasil…"/>

      <label style={T.lbl}>Cor do grupo</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:9,marginBottom:22}}>
        {[...GC,"#059669","#d97706","#e11d48","#0891b2","#7c3aed","#94a3b8"].map(c=>(
          <div key={c} onClick={()=>setColor(c)}
            style={{width:32,height:32,borderRadius:8,background:c,cursor:"pointer",
              border:color===c?"3px solid var(--text)":"2.5px solid transparent",
              transition:"transform .1s,border .1s"}}
            onMouseEnter={e=>e.currentTarget.style.transform="scale(1.15)"}
            onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}/>
        ))}
      </div>
      {/* Preview */}
      <div style={{padding:"8px 14px",borderRadius:8,background:color+"18",borderLeft:`4px solid ${color}`,marginBottom:22,fontWeight:700,fontSize:14,color}}>
        {nome||"Nome do grupo"}
      </div>

      <label style={T.lbl}>Responsável pelo grupo</label>
      <select value={ownerId} onChange={e=>setOwnerId(e.target.value)} style={{...T.inp}}>
        <option value="">— Sem responsável —</option>
        {allUsers.map(u=><option key={u.id} value={u.id}>{u.nome} {u.role_nome?`(${u.role_nome})`:""}</option>)}
      </select>
      <p style={{fontSize:11,color:"var(--text3)",marginTop:8}}>O responsável terá acesso automático e será preenchido nos novos itens.</p>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// GROUP SETTINGS MODAL — gerenciar acesso + trocar responsável
// ─────────────────────────────────────────────────────────────────────────────
function GroupSettingsModal({group,allUsers,currentAccess,onSaveAccess,onSaveOwner,onClose}) {
  const toast=useToast();
  const [access,setAccess]=useState([...currentAccess]);
  const [ownerId,setOwnerId]=useState(group.owner_id||"");
  const [saving,setSaving]=useState(false);
  const toggle=uid=>setAccess(p=>p.includes(uid)?p.filter(x=>x!==uid):[...p,uid]);

  const save=async()=>{
    setSaving(true);
    await onSaveAccess(access);
    if(ownerId!==group.owner_id) await onSaveOwner(ownerId||null);
    setSaving(false);
    onClose();
  };

  return (
    <Modal title={`Configurações — ${group.nome}`} onClose={onClose} width={480}
      footer={<>
        <button onClick={onClose} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={save} disabled={saving} style={{...T.btn,background:"var(--blue)",color:"#fff"}}>
          {saving?"Salvando…":"Salvar configurações"}
        </button>
      </>}>
      {/* Responsável (owner) */}
      <label style={T.lbl}>Responsável pelo grupo</label>
      <select value={ownerId} onChange={e=>setOwnerId(e.target.value)} style={{...T.inp,marginBottom:24}}>
        <option value="">— Sem responsável —</option>
        {allUsers.map(u=><option key={u.id} value={u.id}>{u.nome} {u.role_nome?`(${u.role_nome})`:""}</option>)}
      </select>

      {/* Lista de acesso */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
        <label style={T.lbl}>Quem pode acessar este grupo</label>
        <span style={{fontSize:11,color:"var(--text3)"}}>{access.length} selecionado(s)</span>
      </div>
      <p style={{fontSize:12,color:"var(--text3)",marginBottom:16,lineHeight:1.6}}>
        Usuários com "todas as permissões" (CEO, Gerente, etc.) sempre veem todos os grupos independentemente desta lista.
        Esta configuração afeta apenas SDR e Closer.
      </p>
      <div style={{display:"flex",flexDirection:"column",gap:8,maxHeight:320,overflowY:"auto"}}>
        {allUsers.map(u=>{
          const active=access.includes(u.id);
          const isOwner=u.id===ownerId;
          return (
            <div key={u.id} onClick={()=>toggle(u.id)}
              style={{display:"flex",alignItems:"center",gap:12,padding:"11px 14px",borderRadius:10,cursor:"pointer",
                border:`1.5px solid ${active?"var(--blue)":"var(--border)"}`,
                background:active?"var(--surface3)":"var(--surface2)",transition:"all .15s"}}>
              <Avatar user={u} size={36}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:600,color:"var(--text)",display:"flex",alignItems:"center",gap:8}}>
                  {u.nome}
                  {isOwner&&<span style={{fontSize:10,background:"var(--blue)",color:"#fff",borderRadius:4,padding:"1px 6px",fontWeight:700}}>OWNER</span>}
                </div>
                <div style={{fontSize:11,color:"var(--text3)"}}>{u.role_nome||u.funcao||u.email}</div>
              </div>
              <div style={{width:22,height:22,borderRadius:6,border:`2px solid ${active?"var(--blue)":"var(--border)"}`,
                background:active?"var(--blue)":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"all .15s"}}>
                {active&&<span style={{color:"#fff",fontSize:14,lineHeight:1}}>✓</span>}
              </div>
            </div>
          );
        })}
      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN PANEL — gerenciar cargos e usuários (somente admin/ceo/gerente)
// ─────────────────────────────────────────────────────────────────────────────
function AdminPanel({onBack}) {
  const toast=useToast();
  const {profile}=useAuth();
  const perms=getPerms(profile);
  const [tab,setTab]=useState("users"); // "users" | "roles"
  const [users,setUsers]=useState([]);
  const [roles,setRoles]=useState([]);
  const [loading,setLoading]=useState(true);
  const [editingRole,setEditingRole]=useState(null);
  const [newRoleForm,setNewRoleForm]=useState(false);

  useEffect(()=>{
    (async()=>{
      const [{data:u},{data:r}]=await Promise.all([
        db.from("profiles").select("*, roles(id,nome,slug,cor)").order("nome"),
        db.from("roles").select("*").order("nome"),
      ]);
      setUsers(u||[]);setRoles(r||[]);setLoading(false);
    })();
  },[]);

  const assignRole=async(userId,roleId)=>{
    const {error}=await db.from("profiles").update({role_id:roleId||null}).eq("id",userId);
    if(error){toast(error.message,"error");return;}
    const newRole=roles.find(r=>r.id===roleId)||null;
    setUsers(p=>p.map(u=>u.id===userId?{...u,role_id:roleId||null,roles:newRole}:u));
    toast("Cargo atribuído com sucesso!");
  };

  const saveRole=async(roleData)=>{
    const slug=roleData.slug||(roleData.nome||"").toLowerCase().replace(/[^a-z0-9]+/g,"_");
    // Salva permissões no campo JSON do banco
    const permissions=roleData.permsLocal||{};
    const payload={nome:roleData.nome, cor:roleData.cor, slug, permissions};
    if(roleData.id){
      const {error}=await db.from("roles").update(payload).eq("id",roleData.id);
      if(error){toast("Erro ao salvar cargo: "+error.message,"error");return;}
      setRoles(p=>p.map(r=>r.id===roleData.id?{...r,...payload}:r));
      toast("Cargo e permissões salvos!");
    } else {
      const {data,error}=await db.from("roles").insert(payload).select().single();
      if(error){toast("Erro ao criar cargo: "+error.message,"error");return;}
      if(data) setRoles(p=>[...p,data]);
      toast("Cargo criado!");
    }
    setEditingRole(null);setNewRoleForm(false);
  };

  const pad=14;

  return (
    <div style={{flex:1,overflowY:"auto",background:"var(--bg)",padding:"24px"}}>
      <div style={{maxWidth:720,margin:"0 auto"}}>
        <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:24}}>
          <button onClick={onBack} style={{...T.btn,background:"none",border:"1.5px solid var(--border)",color:"var(--text)",padding:"8px 16px",fontSize:13}}>← Voltar</button>
          <div style={{fontWeight:900,fontSize:20,color:"var(--text)"}}>Administração</div>
        </div>

        {/* Tabs */}
        <div style={{display:"flex",gap:4,marginBottom:20,background:"var(--surface2)",borderRadius:10,padding:4}}>
          {[["users","👥 Usuários & Cargos"],["roles","🏷️ Gerenciar Cargos"]].map(([id,lbl])=>(
            <button key={id} onClick={()=>setTab(id)}
              style={{...T.btn,flex:1,padding:"10px",fontSize:13,borderRadius:8,
                background:tab===id?"var(--surface)":"transparent",
                color:tab===id?"var(--text)":"var(--text3)",
                boxShadow:tab===id?"0 2px 8px var(--shadow)":"none"}}>
              {lbl}
            </button>
          ))}
        </div>

        {loading&&<div style={{display:"flex",justifyContent:"center",padding:40}}><Spinner size={36}/></div>}

        {/* USERS TAB */}
        {!loading&&tab==="users"&&(
          <div style={{background:"var(--surface)",borderRadius:14,overflow:"hidden",boxShadow:"0 2px 16px var(--shadow)"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid var(--border)",fontWeight:700,fontSize:14,color:"var(--text)"}}>
              {users.length} usuário(s) cadastrado(s)
            </div>
            {users.map((u,i)=>(
              <div key={u.id} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",
                borderBottom:i<users.length-1?"1px solid var(--border)":"none",flexWrap:"wrap",gap:12}}>
                <Avatar user={u} size={42}/>
                <div style={{flex:1,minWidth:160}}>
                  <div style={{fontWeight:700,fontSize:14,color:"var(--text)"}}>{u.nome}</div>
                  <div style={{fontSize:12,color:"var(--text3)"}}>{u.email}</div>
                </div>
                {/* Badge do cargo atual */}
                {u.roles&&<span style={{fontSize:11,background:(u.roles.cor||"#94a3b8")+"22",color:u.roles.cor||"#94a3b8",border:`1px solid ${u.roles.cor||"#94a3b8"}40`,borderRadius:6,padding:"3px 10px",fontWeight:600,whiteSpace:"nowrap"}}>{u.roles.nome}</span>}
                {/* Select de cargo */}
                {perms.manageUsers&&<select value={u.role_id||""} onChange={e=>assignRole(u.id,e.target.value||null)}
                  style={{...T.inp,width:"auto",padding:"7px 10px",fontSize:12,maxWidth:200}}>
                  <option value="">— Sem cargo —</option>
                  {roles.map(r=><option key={r.id} value={r.id}>{r.nome}</option>)}
                </select>}
              </div>
            ))}
          </div>
        )}

        {/* ROLES TAB */}
        {!loading&&tab==="roles"&&(
          <div>
            <div style={{display:"flex",justifyContent:"flex-end",marginBottom:14}}>
              {perms.manageRoles&&<button onClick={()=>setNewRoleForm(true)} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"9px 18px"}}>+ Novo cargo</button>}
            </div>
            <div style={{background:"var(--surface)",borderRadius:14,overflow:"hidden",boxShadow:"0 2px 16px var(--shadow)"}}>
              {roles.map((r,i)=>(
                <div key={r.id} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",borderBottom:i<roles.length-1?"1px solid var(--border)":"none"}}>
                  <div style={{width:14,height:14,borderRadius:4,background:r.cor||"#94a3b8",flexShrink:0}}/>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:14,color:"var(--text)"}}>{r.nome}</div>
                    <div style={{fontSize:11,color:"var(--text3)",marginTop:2}}>{r.slug}</div>
                  </div>
                  {perms.manageRoles&&<button onClick={()=>setEditingRole(r)} style={T.iBtn}>✏️ Editar</button>}
                </div>
              ))}
              {!roles.length&&<div style={{padding:"24px",textAlign:"center",color:"var(--text3)",fontSize:14}}>Nenhum cargo cadastrado.</div>}
            </div>
          </div>
        )}
      </div>

      {/* Role form modal */}
      {(editingRole||newRoleForm)&&(
        <RoleFormModal
          initial={editingRole||{}}
          onSave={saveRole}
          onCancel={()=>{setEditingRole(null);setNewRoleForm(false);}}
        />
      )}
    </div>
  );
}

function RoleFormModal({initial={},onSave,onCancel}) {
  const [nome,setNome]=useState(initial.nome||"");
  const [cor,setCor]=useState(initial.cor||"#3145FF");
  // Permissões editáveis visualmente (para cargos customizados e edição dos padrão)
  const PERM_LABELS=[
    ["viewAll","Ver todos os grupos"],
    ["editAny","Editar qualquer item"],
    ["manageBoards","Gerenciar quadros"],
    ["manageCols","Gerenciar colunas"],
    ["manageUsers","Gerenciar usuários"],
    ["manageRoles","Gerenciar cargos"],
    ["export","Exportar relatórios"],
    ["deleteAny","Excluir qualquer item"],
    ["createGroups","Criar grupos"],
  ];
  const existingSlug=initial.slug||(initial.nome||"").toLowerCase().replace(/\s+/g,"_");
  const [permsLocal,setPermsLocal]=useState(()=>{
    // PRIORIDADE: initial.permissions (banco) → ROLE_MATRIX (fallback hardcoded)
    const dbPerms = initial.permissions && Object.keys(initial.permissions).length>0
      ? initial.permissions : null;
    const matrixPerms = ROLE_MATRIX[existingSlug]||{};
    const src = dbPerms || matrixPerms;
    const isAdmin = !!(src.all || existingSlug==="administrador");
    const p={};
    PERM_LABELS.forEach(([k])=>{ p[k]=isAdmin||!!src[k]; });
    return p;
  });
  const togglePerm=k=>setPermsLocal(p=>({...p,[k]:!p[k]}));
  const isBuiltIn=!!ROLE_MATRIX[existingSlug];

  const save=()=>{
    if(!nome.trim()) return;
    onSave({...initial, nome:nome.trim(), cor, permsLocal});
  };

  return (
    <Modal title={initial.id?"Editar cargo":"Novo cargo"} onClose={onCancel} width={480}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={save} disabled={!nome.trim()} style={{...T.btn,background:"var(--blue)",color:"#fff",opacity:nome.trim()?1:.5}}>Salvar</button>
      </>}>
      <label style={T.lbl}>Nome do cargo</label>
      <input value={nome} onChange={e=>setNome(e.target.value)} autoFocus
        onKeyDown={e=>{if(e.key==="Enter")save();}}
        style={{...T.inp,marginBottom:20}} placeholder="Ex: SDR Senior, Coordenador…"/>

      <label style={T.lbl}>Cor de identificação</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:9,marginBottom:10}}>
        {OPT_COLORS.map(c=>(
          <div key={c} onClick={()=>setCor(c)}
            style={{width:30,height:30,borderRadius:7,background:c,cursor:"pointer",
              border:cor===c?"3px solid var(--text)":"2px solid transparent",transition:"transform .1s"}}
            onMouseEnter={e=>e.currentTarget.style.transform="scale(1.15)"}
            onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}/>
        ))}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:24}}>
        <div style={{width:22,height:22,borderRadius:5,background:cor,border:"1px solid var(--border)",flexShrink:0}}/>
        <input value={cor} onChange={e=>setCor(e.target.value)} placeholder="#3145FF"
          style={{...T.inp,padding:"6px 10px",fontSize:12,flex:1}}/>
        {nome&&<div style={{padding:"5px 14px",borderRadius:8,background:cor+"22",border:`1px solid ${cor}40`,whiteSpace:"nowrap"}}>
          <span style={{fontSize:12,fontWeight:700,color:cor}}>{nome}</span>
        </div>}
      </div>

      <div style={{borderTop:"1px solid var(--border)",paddingTop:20,marginTop:4}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
          <label style={T.lbl}>Permissões</label>
          {isBuiltIn&&<span style={{fontSize:11,color:"var(--text3)"}}>Cargo padrão — permissões são visuais</span>}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {PERM_LABELS.map(([key,label])=>(
            <div key={key} onClick={()=>togglePerm(key)}
              style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:9,cursor:"pointer",
                border:`1.5px solid ${permsLocal[key]?"var(--blue)":"var(--border)"}`,
                background:permsLocal[key]?"var(--row-sel)":"var(--surface2)",transition:"all .15s"}}>
              <div style={{width:20,height:20,borderRadius:5,border:`2px solid ${permsLocal[key]?"var(--blue)":"var(--border)"}`,
                background:permsLocal[key]?"var(--blue)":"transparent",display:"flex",alignItems:"center",
                justifyContent:"center",flexShrink:0,transition:"all .15s"}}>
                {permsLocal[key]&&<span style={{color:"#fff",fontSize:13,lineHeight:1}}>✓</span>}
              </div>
              <span style={{fontSize:13,fontWeight:permsLocal[key]?600:400,color:"var(--text)"}}>{label}</span>
            </div>
          ))}
        </div>
      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MOVE ITEMS MODAL — move selected items to another board/group (#3)
// ─────────────────────────────────────────────────────────────────────────────
function MoveItemsModal({items,srcBoard,allBoards,onMove,onCancel}) {
  const [destBoardId,setDestBoardId]=useState("");
  const [destGroups,setDestGroups]=useState([]);
  const [destGroupId,setDestGroupId]=useState("");
  const [loading,setLoading]=useState(false);
  const [moving,setMoving]=useState(false);
  // Para Negociações: picker em 2 etapas
  const [negStep,setNegStep]=useState(1);   // 1=closer, 2=subgrupo
  const [negParent,setNegParent]=useState(null);

  const destBoard=allBoards.find(b=>b.id===destBoardId);
  const isNeg=destBoard?.nome==="Negociações";
  const negParents=destGroups.filter(g=>g.is_parent);
  const negChildren=pid=>destGroups.filter(g=>g.parent_group_id===pid);
  const closerLabel=nome=>(nome||"").replace(/Closer\s*[-\u2013]\s*/i,"").trim();
  const getInitials=nome=>closerLabel(nome).split(" ").slice(0,2).map(p=>p[0]||"").join("").toUpperCase();

  useEffect(()=>{
    if(!destBoardId){setDestGroups([]);setDestGroupId("");setNegStep(1);setNegParent(null);return;}
    setLoading(true);
    db.from("groups").select("*").eq("board_id",destBoardId).order("ordem")
      .then(({data})=>{
        setDestGroups(data||[]);
        setDestGroupId("");
        setNegStep(1);setNegParent(null);
        if(!allBoards.find(b=>b.id===destBoardId)?.nome?.includes("Negociações"))
          setDestGroupId(data?.[0]?.id||"");
        setLoading(false);
      });
  },[destBoardId]);

  const otherBoards=allBoards.filter(b=>b.id!==srcBoard.id);

  const doMove=async()=>{
    if(!destBoardId||!destGroupId) return;
    setMoving(true);
    const {data:destCols}=await db.from("columns").select("*").eq("board_id",destBoardId);
    const matchCols=(srcCols,tgtCols)=>{
      const map={};
      for(const sc of srcCols){
        const tc=tgtCols.find(c=>c.nome===sc.nome&&c.tipo===sc.tipo);
        if(tc) map[sc.id]=tc.id;
      }
      return map;
    };
    const colMap=matchCols(srcBoard.columns,destCols||[]);
    const isNegDest=destBoard?.nome==="Negociações";
    for(const item of items){
      const {data:ni}=await db.from("items")
        .insert({board_id:destBoardId,group_id:destGroupId,ordem:9999,
          ...(isNegDest?{origin_item_id:item.id}:{})})
        .select().single();
      if(!ni) continue;
      const {data:vals}=await db.from("item_values").select("*").eq("item_id",item.id);
      const toInsert=[];
      for(const v of vals||[]){
        const tgtId=colMap[v.column_id];
        if(tgtId) toInsert.push({item_id:ni.id,column_id:tgtId,value:v.value});
      }
      if(toInsert.length) await db.from("item_values").insert(toInsert);
      if(item.responsibles?.length)
        await db.from("item_responsables").insert(item.responsibles.map(uid=>({item_id:ni.id,user_id:uid})));
      const {data:upds}=await db.from("item_updates").select("*").eq("item_id",item.id);
      if(upds?.length)
        await db.from("item_updates").insert(upds.map(u=>({item_id:ni.id,content:u.content,created_by:u.created_by})));
    }
    setMoving(false);
    const isVendasDest=destBoard?.nome==="Vendas";
    const destGroupObj=destGroups.find(g=>g.id===destGroupId)||{nome:"Vendas"};
    // Passa IDs dos itens originais do Pré-Vendas quando destino é Negociações
    // Passa destGroupObj quando destino é Vendas para o email ser disparado no BoardPage
    onMove(destBoardId, isNegDest?items.map(i=>i.id):null, isVendasDest?destGroupObj:null);
  };

  // Sub-picker de Negociações inline dentro do modal
  const NegPicker=()=>{
    if(loading) return <div style={{padding:"10px 0",color:"var(--text3)",fontSize:13}}>Carregando grupos…</div>;
    if(negStep===1) return (
      <div style={{display:"flex",flexDirection:"column",gap:8,marginTop:4}}>
        {negParents.map(p=>{
          const subs=negChildren(p.id);
          return (
            <button key={p.id} onClick={()=>{setNegParent(p);setNegStep(2);setDestGroupId("");}}
              style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",
                background:"var(--surface2)",border:`2px solid ${p.color}35`,borderRadius:10,
                cursor:"pointer",width:"100%",textAlign:"left",transition:"all .15s"}}
              onMouseEnter={e=>{e.currentTarget.style.borderColor=p.color;e.currentTarget.style.background="var(--surface3)";}}
              onMouseLeave={e=>{e.currentTarget.style.borderColor=`${p.color}35`;e.currentTarget.style.background="var(--surface2)";}}>
              <div style={{width:38,height:38,borderRadius:"50%",background:p.color,flexShrink:0,
                display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:14,color:"#fff"}}>
                {getInitials(p.nome)}
              </div>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:14,color:"var(--text)"}}>{closerLabel(p.nome)}</div>
                <div style={{fontSize:11,color:"var(--text3)"}}>{subs.map(s=>s.nome).join(" · ")}</div>
              </div>
              <span style={{fontSize:18,color:p.color,opacity:.6}}>›</span>
            </button>
          );
        })}
      </div>
    );
    const subs=negChildren(negParent.id);
    return (
      <div>
        <button onClick={()=>{setNegStep(1);setNegParent(null);setDestGroupId("");}}
          style={{display:"flex",alignItems:"center",gap:8,width:"100%",padding:"8px 12px",
            background:`${negParent.color}15`,border:`1.5px solid ${negParent.color}40`,borderRadius:8,
            cursor:"pointer",marginBottom:10,textAlign:"left"}}>
          <span style={{fontSize:15,color:negParent.color}}>←</span>
          <div style={{width:28,height:28,borderRadius:"50%",background:negParent.color,flexShrink:0,
            display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,color:"#fff"}}>
            {getInitials(negParent.nome)}
          </div>
          <div>
            <div style={{fontWeight:700,fontSize:13,color:"var(--text)"}}>{closerLabel(negParent.nome)}</div>
            <div style={{fontSize:11,color:"var(--text3)"}}>Escolha o sub-grupo</div>
          </div>
        </button>
        <div style={{display:"flex",flexDirection:"column",gap:7}}>
          {subs.map(sg=>(
            <button key={sg.id} onClick={()=>setDestGroupId(sg.id)}
              style={{display:"flex",alignItems:"center",gap:10,padding:"11px 14px",
                background:destGroupId===sg.id?`${sg.color}20`:"var(--surface2)",
                border:`1.5px solid ${destGroupId===sg.id?sg.color:`${sg.color}45`}`,borderRadius:9,
                cursor:"pointer",width:"100%",textAlign:"left",transition:"all .15s"}}
              onMouseEnter={e=>{if(destGroupId!==sg.id){e.currentTarget.style.borderColor=sg.color;e.currentTarget.style.background="var(--surface3)";}}}
              onMouseLeave={e=>{if(destGroupId!==sg.id){e.currentTarget.style.borderColor=`${sg.color}45`;e.currentTarget.style.background="var(--surface2)";}}} >
              <div style={{width:10,height:10,borderRadius:"50%",background:sg.color,flexShrink:0}}/>
              <span style={{fontWeight:600,fontSize:13,color:"var(--text)",flex:1}}>{sg.nome}</span>
              {destGroupId===sg.id&&<span style={{color:sg.color,fontWeight:700,fontSize:13}}>✓</span>}
            </button>
          ))}
        </div>
      </div>
    );
  };

  return (
    <Modal title={`📦 Mover ${items.length} item(s) para outro quadro`} onClose={onCancel} width={490}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={doMove} disabled={!destBoardId||!destGroupId||moving}
          style={{...T.btn,background:"var(--blue)",color:"#fff",opacity:!destBoardId||!destGroupId?0.5:1}}>
          {moving?"Movendo…":"Mover itens"}
        </button>
      </>}>
      <div style={{padding:"12px 14px",background:"var(--surface2)",borderRadius:10,marginBottom:20,
        border:"1px solid var(--border)",fontSize:13,color:"var(--text2)"}}>
        <strong style={{color:"var(--text)"}}>{items.length} item(s)</strong> serão copiados para o destino.
        Colunas com o <strong>mesmo nome e tipo</strong> serão preenchidas automaticamente.
      </div>

      <label style={T.lbl}>Quadro de destino</label>
      <select value={destBoardId} onChange={e=>setDestBoardId(e.target.value)}
        style={{...T.inp,marginBottom:18}}>
        <option value="">— Selecionar quadro —</option>
        {otherBoards.map(b=><option key={b.id} value={b.id}>{b.icon||"📋"} {b.nome}</option>)}
      </select>

      {destBoardId&&(
        <>
          <label style={{...T.lbl,marginBottom:10}}>
            {isNeg?"Escolha o Closer e sub-grupo":"Grupo de destino"}
          </label>
          {isNeg
            ? <NegPicker/>
            : loading
              ? <div style={{padding:"10px 0",color:"var(--text3)",fontSize:13}}>Carregando grupos…</div>
              : <select value={destGroupId} onChange={e=>setDestGroupId(e.target.value)}
                  style={T.inp} disabled={!destBoardId}>
                  <option value="">— Selecionar grupo —</option>
                  {destGroups.map(g=><option key={g.id} value={g.id}>{g.nome}</option>)}
                </select>
          }
          {!isNeg&&<p style={{fontSize:11,color:"var(--text3)",marginTop:12,lineHeight:1.6}}>
            Os itens originais não serão removidos. Atualizações e responsáveis também serão copiados.
          </p>}
        </>
      )}
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// BOARD SETTINGS MODAL — privacidade e acesso ao quadro
// ─────────────────────────────────────────────────────────────────────────────
function BoardSettingsModal({board,allUsers,currentAccess,isPrivate,onSave,onClose}) {
  const toast=useToast();
  const [priv,setPriv]=useState(!!isPrivate);
  const [access,setAccess]=useState([...currentAccess]);
  const [saving,setSaving]=useState(false);
  const toggle=uid=>setAccess(p=>p.includes(uid)?p.filter(x=>x!==uid):[...p,uid]);
  const save=async()=>{setSaving(true);await onSave({isPrivate:priv,access});setSaving(false);onClose();};
  return (
    <Modal title={`Configurações — ${board.nome}`} onClose={onClose} width={500}
      footer={<>
        <button onClick={onClose} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={save} disabled={saving} style={{...T.btn,background:"var(--blue)",color:"#fff"}}>
          {saving?"Salvando…":"Salvar"}
        </button>
      </>}>
      {/* Privacidade */}
      <div style={{marginBottom:22}}>
        <label style={T.lbl}>Visibilidade do quadro</label>
        <div style={{display:"flex",gap:12,marginTop:8}}>
          {[[false,"🌐","Público","Todos os membros veem"],[true,"🔒","Privado","Apenas membros selecionados"]].map(([val,icon,label,desc])=>(
            <div key={String(val)} onClick={()=>setPriv(val)}
              style={{flex:1,padding:"14px 16px",borderRadius:10,cursor:"pointer",
                border:`2px solid ${priv===val?"var(--blue)":"var(--border)"}`,
                background:priv===val?"var(--row-sel)":"var(--surface2)",transition:"all .15s"}}>
              <div style={{fontSize:20,marginBottom:6}}>{icon}</div>
              <div style={{fontWeight:700,fontSize:13,color:"var(--text)",marginBottom:3}}>{label}</div>
              <div style={{fontSize:11,color:"var(--text3)"}}>{desc}</div>
            </div>
          ))}
        </div>
      </div>
      {/* Membros com acesso (só aparece se privado) */}
      {priv&&<>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
          <label style={T.lbl}>Membros com acesso</label>
          <span style={{fontSize:11,color:"var(--text3)"}}>{access.length} selecionado(s)</span>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:7,maxHeight:280,overflowY:"auto"}}>
          {allUsers.map(u=>{
            const active=access.includes(u.id);
            return (
              <div key={u.id} onClick={()=>toggle(u.id)}
                style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:9,cursor:"pointer",
                  border:`1.5px solid ${active?"var(--blue)":"var(--border)"}`,
                  background:active?"var(--row-sel)":"var(--surface2)",transition:"all .15s"}}>
                <Avatar user={u} size={34}/>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:600,color:"var(--text)"}}>{u.nome}</div>
                  <div style={{fontSize:11,color:"var(--text3)"}}>{u.role_nome||u.email}</div>
                </div>
                <div style={{width:22,height:22,borderRadius:6,border:`2px solid ${active?"var(--blue)":"var(--border)"}`,
                  background:active?"var(--blue)":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                  {active&&<span style={{color:"#fff",fontSize:13}}>✓</span>}
                </div>
              </div>
            );
          })}
        </div>
        <p style={{fontSize:11,color:"var(--text3)",marginTop:10}}>
          Usuários com cargo de Administrador, CEO e Gerente sempre têm acesso independentemente desta lista.
        </p>
      </>}
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// BOARD FORM MODAL
// ─────────────────────────────────────────────────────────────────────────────
function BoardFormModal({initial={},onSave,onCancel}) {
  const [nome,setNome]=useState(initial.nome||"");
  const [icon,setIcon]=useState(initial.icon||"📋");
  return (
    <Modal title={initial.id?"Editar quadro":"Novo quadro"} onClose={onCancel} width={420}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={()=>nome.trim()&&onSave({nome:nome.trim(),icon})} disabled={!nome.trim()}
          style={{...T.btn,background:"var(--blue)",color:"#fff",opacity:nome.trim()?1:.5}}>Salvar</button>
      </>}>
      <label style={T.lbl}>Nome</label>
      <input value={nome} onChange={e=>setNome(e.target.value)} autoFocus
        onKeyDown={e=>{if(e.key==="Enter"&&nome.trim())onSave({nome:nome.trim(),icon});}}
        style={{...T.inp,marginBottom:20}} placeholder="Nome do quadro"/>
      <label style={T.lbl}>Ícone</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
        {BICONS.map(ic=>(
          <button key={ic} onClick={()=>setIcon(ic)}
            style={{width:44,height:44,border:`2px solid ${ic===icon?"var(--blue)":"var(--border)"}`,borderRadius:9,
              background:ic===icon?"var(--surface3)":"var(--surface2)",cursor:"pointer",fontSize:22,transition:"all .15s"}}>
            {ic}
          </button>
        ))}
      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SIDEBAR — responsiva com drawer no mobile
// ─────────────────────────────────────────────────────────────────────────────
function Sidebar({boards,currentBoardId,currentUser,wsNome,dark,onToggleDark,perms,onSelectBoard,onAddBoard,onEditBoard,onDelBoard,onReorderBoards,onBoardSettings,onDuplicateBoard,onProfile,onAdmin,onDashboard,onLogout,open,onClose}) {
  const {isMobile}=useBreakpoint();
  const [boardMenu,setBoardMenu]=useState(null);
  const [boardMenuPos,setBoardMenuPos]=useState({top:0,left:248});
  const [dragBoardIdx,setDragBoardIdx]=useState(null);
  const [dragBoardOver,setDragBoardOver]=useState(null);

  const dropBoard=()=>{
    if(dragBoardIdx===null||dragBoardOver===null||dragBoardIdx===dragBoardOver){
      setDragBoardIdx(null);setDragBoardOver(null);return;
    }
    const next=[...boards];
    const [moved]=next.splice(dragBoardIdx,1);
    next.splice(dragBoardOver,0,moved);
    setDragBoardIdx(null);setDragBoardOver(null);
    onReorderBoards(next);
  };

  const content=(
    <div style={{display:"flex",flexDirection:"column",height:"100%",overflowY:"auto"}}>
      {/* Logo */}
      <div style={{padding:"20px 18px 16px",borderBottom:"1px solid rgba(255,255,255,.12)",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:12,justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:40,height:40,background:"var(--alt)",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:22,color:"#fff",flexShrink:0}}>M</div>
            <div>
              <div style={{fontWeight:900,fontSize:15,letterSpacing:-.5,color:"#fff"}}>Monvatti CRM</div>
              <div style={{fontSize:11,color:"rgba(255,255,255,.4)",marginTop:1}}>{wsNome||"Workspace"}</div>
            </div>
          </div>
          {isMobile&&<button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.5)",fontSize:22,lineHeight:1}}>×</button>}
        </div>
      </div>
      {/* Boards */}
      <div style={{padding:"14px 14px 6px",flex:1,overflowY:"auto"}}>
        {/* Dashboard link */}
        <div onClick={()=>{onDashboard&&onDashboard();if(isMobile)onClose();}}
          style={{display:"flex",alignItems:"center",gap:9,padding:"9px 11px",borderRadius:9,cursor:"pointer",
            color:"rgba(255,255,255,.65)",fontSize:13,marginBottom:10,
            background:"transparent"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.1)"}
          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <span style={{fontSize:17,flexShrink:0}}>📊</span>
          <span style={{flex:1}}>Dashboard</span>
        </div>
        <div style={{height:1,background:"rgba(255,255,255,.1)",marginBottom:10}}/>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,padding:"0 4px"}}>
          <div style={{fontSize:10,color:"rgba(255,255,255,.4)",textTransform:"uppercase",letterSpacing:.7,fontWeight:700}}>Quadros</div>
          <button onClick={onAddBoard} style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.5)",fontSize:22,lineHeight:1}}
            onMouseEnter={e=>e.currentTarget.style.color="#fff"} onMouseLeave={e=>e.currentTarget.style.color="rgba(255,255,255,.5)"}>+</button>
        </div>
        {boards.map((b,bi)=>(
          <div key={b.id} draggable
            onDragStart={()=>setDragBoardIdx(bi)}
            onDragOver={e=>{e.preventDefault();setDragBoardOver(bi);}}
            onDrop={dropBoard}
            onDragEnd={()=>{setDragBoardIdx(null);setDragBoardOver(null);}}
            style={{position:"relative",marginBottom:3,opacity:dragBoardIdx===bi?0.4:1,
              outline:dragBoardOver===bi&&dragBoardIdx!==bi?"2px dashed rgba(255,255,255,.5)":"2px solid transparent",
              borderRadius:10,transition:"opacity .15s,outline .1s"}}>
            <div onClick={()=>{onSelectBoard(b.id);if(isMobile)onClose();}}
              style={{display:"flex",alignItems:"center",gap:9,padding:"9px 11px",borderRadius:9,cursor:"pointer",
                background:b.id===currentBoardId?"rgba(255,255,255,.18)":"transparent",
                color:b.id===currentBoardId?"#fff":"rgba(255,255,255,.65)",fontSize:13}}
              onMouseEnter={e=>{if(b.id!==currentBoardId)e.currentTarget.style.background="rgba(255,255,255,.1)";}}
              onMouseLeave={e=>{if(b.id!==currentBoardId)e.currentTarget.style.background="transparent";}}>
              <span style={{fontSize:17,flexShrink:0}}>{b.icon||"📋"}</span>
              <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.nome}</span>
              <span style={{fontSize:11,color:"rgba(255,255,255,.3)",background:"rgba(255,255,255,.12)",borderRadius:10,padding:"1px 8px",flexShrink:0}}>{b.totalItems||0}</span>
              <button onClick={e=>{e.stopPropagation();const r=e.currentTarget.getBoundingClientRect();setBoardMenuPos({top:r.bottom+4,left:isMobile?16:r.right+8});setBoardMenu(boardMenu===b.id?null:b.id);}}
                style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.4)",fontSize:14,padding:"0 3px",lineHeight:1,flexShrink:0}}
                onMouseEnter={e=>e.currentTarget.style.color="#fff"} onMouseLeave={e=>e.currentTarget.style.color="rgba(255,255,255,.4)"}>⋯</button>
            </div>
            {boardMenu===b.id&&createPortal(
              <div style={{position:"fixed",top:boardMenuPos.top,left:boardMenuPos.left,zIndex:6000,background:"var(--surface)",border:"1px solid var(--border)",borderRadius:12,boxShadow:"0 12px 48px var(--shadowMd)",padding:7,minWidth:200}}>
                <div onClick={()=>{setBoardMenu(null);onEditBoard(b);}} style={{padding:"9px 13px",borderRadius:7,cursor:"pointer",fontSize:13,color:"var(--text)"}} onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>✏️ Renomear quadro</div>
                <div onClick={()=>{setBoardMenu(null);onBoardSettings&&onBoardSettings(b);}} style={{padding:"9px 13px",borderRadius:7,cursor:"pointer",fontSize:13,color:"var(--text)"}} onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>⚙️ Configurações</div>
                <div onClick={()=>{setBoardMenu(null);onDuplicateBoard&&onDuplicateBoard(b);}} style={{padding:"9px 13px",borderRadius:7,cursor:"pointer",fontSize:13,color:"var(--text)"}} onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>🗂️ Duplicar quadro</div>
                <div onClick={()=>{setBoardMenu(null);onDelBoard(b);}} style={{padding:"9px 13px",borderRadius:7,cursor:"pointer",fontSize:13,color:"#dc2626"}} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>🗑️ Excluir quadro</div>
              </div>,
              document.body
            )}
          </div>
        ))}
      </div>
      {/* Bottom bar */}
      <div style={{padding:"12px 16px",borderTop:"1px solid rgba(255,255,255,.1)",flexShrink:0}}>
        {currentUser?.role_nome&&(
          <div style={{marginBottom:9,padding:"4px 10px",borderRadius:7,background:"rgba(255,255,255,.08)",display:"inline-flex",alignItems:"center",gap:7}}>
            <div style={{width:7,height:7,borderRadius:"50%",background:currentUser.roles?.cor||"#94a3b8",flexShrink:0}}/>
            <span style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,.7)"}}>{currentUser.role_nome}</span>
          </div>
        )}
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <Avatar user={currentUser} size={36}/>
          <div style={{flex:1,minWidth:0,overflow:"hidden"}}>
            <div style={{fontSize:13,fontWeight:700,color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{currentUser?.nome}</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,.45)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{currentUser?.email}</div>
          </div>
          <button onClick={onToggleDark} title={dark?"Modo claro":"Modo escuro"}
            style={{background:"none",border:"none",cursor:"pointer",fontSize:17,lineHeight:1}}>{dark?"☀️":"🌙"}</button>
          {(perms?.all||perms?.manageUsers)&&(
            <button onClick={onAdmin} title="Administração"
              style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.45)",fontSize:17,lineHeight:1}}
              onMouseEnter={e=>e.currentTarget.style.color="#fff"}
              onMouseLeave={e=>e.currentTarget.style.color="rgba(255,255,255,.45)"}>⚙</button>
          )}
          <button onClick={onProfile} title="Perfil"
            style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.45)",fontSize:17,lineHeight:1}}
            onMouseEnter={e=>e.currentTarget.style.color="#fff"} onMouseLeave={e=>e.currentTarget.style.color="rgba(255,255,255,.45)"}>👤</button>
          <button onClick={onLogout} title="Sair"
            style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,.45)",fontSize:17,lineHeight:1}}
            onMouseEnter={e=>e.currentTarget.style.color="#fff"} onMouseLeave={e=>e.currentTarget.style.color="rgba(255,255,255,.45)"}>⏻</button>
        </div>
      </div>
    </div>
  );

  if(isMobile){
    return createPortal(
      <>
        {open&&<div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:1600}}/>}
        <div style={{position:"fixed",top:0,left:0,bottom:0,width:260,background:"var(--sidebar)",
          transform:open?"translateX(0)":"translateX(-100%)",transition:"transform .25s ease",zIndex:1700}}>
          {content}
        </div>
      </>,
      document.body
    );
  }

  return (
    <div style={{width:240,background:"var(--sidebar)",flexShrink:0,height:"100%"}}>
      {content}
    </div>
  );
}

function ProfilePage({onBack}) {
  const {profile,updateProfile,updatePassword,session}=useAuth();
  const toast=useToast();
  const {isMobile}=useBreakpoint();
  const [nome,setNome]=useState(profile?.nome||"");
  const [pw,setPw]=useState("");
  const [pw2,setPw2]=useState("");
  const [saving,setSaving]=useState(false);
  const [uploading,setUploading]=useState(false);
  const fileRef=useRef();

  const roleNome=profile?.role_nome||profile?.roles?.nome||profile?.funcao||"Sem cargo";
  const roleCor=profile?.roles?.cor||"#94a3b8";

  const handlePhotoUpload=async e=>{
    const file=e.target.files?.[0]; if(!file) return;
    if(!file.type.startsWith("image/")){toast("Apenas imagens são permitidas","error");return;}
    if(file.size>5*1024*1024){toast("Imagem muito grande (máx 5MB)","error");return;}
    setUploading(true);
    const ext=file.name.split(".").pop();
    const path=`${session.user.id}/avatar.${ext}`;
    const {error:upErr}=await db.storage.from("avatars").upload(path,file,{upsert:true,contentType:file.type});
    if(upErr){toast("Erro ao fazer upload: "+upErr.message,"error");setUploading(false);return;}
    const {data:{publicUrl}}=db.storage.from("avatars").getPublicUrl(path);
    await updateProfile({foto_url:publicUrl});
    toast("Foto atualizada!");setUploading(false);
  };

  const save=async()=>{
    setSaving(true);
    const {error}=await updateProfile({nome});
    if(error)toast("Erro ao salvar","error"); else toast("Perfil atualizado!");
    setSaving(false);
  };

  const changePw=async()=>{
    if(pw!==pw2){toast("Senhas não conferem","error");return;}
    if(pw.length<6){toast("Mínimo 6 caracteres","warning");return;}
    setSaving(true);
    const {error}=await updatePassword(pw);
    if(error)toast(error.message,"error"); else{toast("Senha alterada!");setPw("");setPw2("");}
    setSaving(false);
  };

  const pad=isMobile?14:32;

  return (
    <div style={{flex:1,overflowY:"auto",background:"var(--bg)",padding:isMobile?"16px":"40px 24px"}}>
      <div style={{maxWidth:580,margin:"0 auto"}}>
        <button onClick={onBack} style={{...T.btn,background:"none",border:"1.5px solid var(--border)",color:"var(--text)",marginBottom:24,padding:"8px 16px",fontSize:13}}>← Voltar</button>

        <div style={{background:"var(--surface)",borderRadius:16,padding:pad,boxShadow:"0 4px 24px var(--shadow)",marginBottom:20}}>
          {/* Avatar + info */}
          <div style={{display:"flex",alignItems:"center",gap:20,marginBottom:28}}>
            <div style={{position:"relative"}}>
              <Avatar user={profile} size={76}/>
              <button onClick={()=>fileRef.current?.click()}
                style={{position:"absolute",bottom:-4,right:-4,width:28,height:28,borderRadius:"50%",background:"var(--blue)",color:"#fff",border:"2px solid var(--surface)",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>
                {uploading?<Spinner size={14}/>:"📷"}
              </button>
              <input ref={fileRef} type="file" accept="image/*" onChange={handlePhotoUpload} style={{display:"none"}}/>
            </div>
            <div>
              <div style={{fontSize:20,fontWeight:800,color:"var(--text)"}}>{profile?.nome||"—"}</div>
              <div style={{fontSize:13,color:"var(--text3)",marginTop:3}}>{profile?.email}</div>
              {/* Cargo — badge readonly, apenas admins podem alterar via painel */}
              <div style={{marginTop:8,display:"inline-flex",alignItems:"center",gap:8,
                background:roleCor+"20",border:`1px solid ${roleCor}40`,borderRadius:8,
                padding:"4px 12px"}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:roleCor,flexShrink:0}}/>
                <span style={{fontSize:12,fontWeight:700,color:roleCor}}>{roleNome}</span>
                <span style={{fontSize:11,color:"var(--text3)"}}>· definido pelo administrador</span>
              </div>
            </div>
          </div>

          <label style={T.lbl}>Nome completo</label>
          <input value={nome} onChange={e=>setNome(e.target.value)} style={{...T.inp,marginBottom:24}}/>
          <button onClick={save} disabled={saving} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"11px 28px"}}>
            {saving?"Salvando…":"Salvar alterações"}
          </button>
        </div>

        <div style={{background:"var(--surface)",borderRadius:16,padding:pad,boxShadow:"0 4px 24px var(--shadow)"}}>
          <div style={{fontWeight:800,fontSize:16,color:"var(--text)",marginBottom:22}}>Alterar senha</div>
          <label style={T.lbl}>Nova senha</label>
          <input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Mínimo 6 caracteres" style={{...T.inp,marginBottom:16}}/>
          <label style={T.lbl}>Confirmar senha</label>
          <input type="password" value={pw2} onChange={e=>setPw2(e.target.value)} style={{...T.inp,marginBottom:24}}/>
          <button onClick={changePw} disabled={saving||!pw} style={{...T.btn,background:"var(--sidebar)",color:"#fff",padding:"11px 28px",opacity:!pw?0.5:1}}>
            {saving?"Salvando…":"Alterar senha"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// AUTH PAGES
// ─────────────────────────────────────────────────────────────────────────────
function AuthLayout({children}) {
  const {isMobile}=useBreakpoint();
  return (
    <div style={{
      minHeight:"100vh",
      background:"var(--sidebar)",
      display:"flex",alignItems:"center",justifyContent:"center",
      padding:20,
      position:"relative",
      overflow:"hidden",
    }}>
      {/* Decoração de fundo */}
      <div style={{position:"absolute",top:-120,right:-120,width:400,height:400,borderRadius:"50%",
        background:"var(--alt)",opacity:.07,pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:-80,left:-80,width:280,height:280,borderRadius:"50%",
        background:"var(--blue)",opacity:.06,pointerEvents:"none"}}/>

      <div style={{
        background:"var(--surface)",
        borderRadius:20,
        padding:isMobile?"28px 22px":"44px 44px",
        width:440,
        maxWidth:"100%",
        boxShadow:"0 32px 80px rgba(0,0,0,.45)",
        position:"relative",
        zIndex:1,
      }}>
        {/* Logo */}
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:32}}>
          <div style={{
            width:48,height:48,
            background:"linear-gradient(135deg, var(--alt) 0%, var(--blue) 100%)",
            borderRadius:14,
            display:"flex",alignItems:"center",justifyContent:"center",
            fontWeight:900,fontSize:24,color:"#fff",flexShrink:0,
            boxShadow:"0 4px 16px var(--blue)44",
          }}>M</div>
          <div>
            <div style={{fontWeight:900,fontSize:20,color:"var(--text)",letterSpacing:-.5}}>Monvatti CRM</div>
            <div style={{fontSize:12,color:"var(--text3)",marginTop:1}}>Plataforma comercial interna</div>
          </div>
        </div>

        {/* Linha divisória sutil */}
        <div style={{height:1,background:"var(--border)",marginBottom:28}}/>

        {children}
      </div>
    </div>
  );
}

function LoginPage({onSwitch}) {
  const {signIn}=useAuth();
  const toast=useToast();
  const [email,setEmail]=useState("");
  const [senha,setSenha]=useState("");
  const [busy,setBusy]=useState(false);

  const submit=async e=>{
    e.preventDefault(); setBusy(true);
    const {error}=await signIn(email.trim().toLowerCase(),senha);
    if(error) toast(
      error.message==="Invalid login credentials"?"E-mail ou senha incorretos":error.message,
      "error"
    );
    setBusy(false);
  };

  const foc=e=>e.target.style.borderColor="var(--blue)";
  const blr=e=>e.target.style.borderColor="var(--border)";

  return (
    <AuthLayout>
      <div style={{marginBottom:24}}>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:4}}>Bem-vindo de volta</div>
        <div style={{fontSize:13,color:"var(--text2)"}}>Entre com suas credenciais para acessar o CRM.</div>
      </div>
      <form onSubmit={submit}>
        <label style={T.lbl}>E-mail</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoFocus
          placeholder="nome@empresa.com"
          style={{...T.inp,marginBottom:16}} onFocus={foc} onBlur={blr}/>

        <label style={T.lbl}>Senha</label>
        <input type="password" value={senha} onChange={e=>setSenha(e.target.value)} required
          placeholder="Sua senha"
          style={{...T.inp,marginBottom:26}} onFocus={foc} onBlur={blr}/>

        <button type="submit" disabled={busy||!email||!senha}
          style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"13px",fontSize:15,
            opacity:busy||!email||!senha?0.6:1}}>
          {busy?"Entrando…":"Entrar"}
        </button>
      </form>
      <p style={{textAlign:"center",fontSize:13,color:"var(--text3)",marginTop:22,marginBottom:0}}>
        Sem conta?{" "}
        <button onClick={onSwitch}
          style={{background:"none",border:"none",color:"var(--blue)",cursor:"pointer",fontWeight:700,fontSize:13}}>
          Criar conta
        </button>
      </p>
    </AuthLayout>
  );
}

function RegisterPage({onSwitch}) {
  const {signUp}=useAuth();
  const toast=useToast();
  const {isMobile}=useBreakpoint();
  const [nome,setNome]=useState("");
  const [email,setEmail]=useState("");
  const [senha,setSenha]=useState("");
  const [conf,setConf]=useState("");
  const [busy,setBusy]=useState(false);
  const [done,setDone]=useState(false);

  const submit=async e=>{
    e.preventDefault();
    if(senha.length<6){toast("Mínimo 6 caracteres na senha","warning");return;}
    if(senha!==conf){toast("As senhas não conferem","error");return;}
    setBusy(true);
    const {error}=await signUp(email.trim().toLowerCase(),senha,nome.trim());
    if(error){
      if(error.message?.toLowerCase().includes("rate limit"))
        toast("Limite de e-mails atingido. Peça ao admin para criar sua conta via painel Supabase, ou aguarde alguns minutos.","error");
      else
        toast(error.message,"error");
    } else setDone(true);
    setBusy(false);
  };

  const foc=e=>e.target.style.borderColor="var(--blue)";
  const blr=e=>e.target.style.borderColor="var(--border)";

  if(done) return (
    <AuthLayout>
      <div style={{textAlign:"center",padding:"10px 0"}}>
        <div style={{fontSize:52,marginBottom:18}}>✅</div>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:10}}>Conta criada com sucesso!</div>
        <p style={{fontSize:13,color:"var(--text2)",lineHeight:1.7,marginBottom:26}}>
          Seu acesso foi solicitado.<br/>
          Um administrador irá atribuir seu cargo em breve.<br/>
          Após isso, faça login normalmente.
        </p>
        <button onClick={onSwitch}
          style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"13px",fontSize:15}}>
          Ir para o login
        </button>
      </div>
    </AuthLayout>
  );

  return (
    <AuthLayout>
      <div style={{marginBottom:22}}>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:4}}>Criar conta</div>
        <div style={{fontSize:13,color:"var(--text2)"}}>Preencha seus dados. O cargo será atribuído pelo administrador.</div>
      </div>
      <form onSubmit={submit}>
        <label style={T.lbl}>Nome completo</label>
        <input type="text" value={nome} onChange={e=>setNome(e.target.value)} required autoFocus
          placeholder="Seu nome completo"
          style={{...T.inp,marginBottom:16}} onFocus={foc} onBlur={blr}/>

        <label style={T.lbl}>E-mail corporativo</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required
          placeholder="nome@empresa.com"
          style={{...T.inp,marginBottom:16}} onFocus={foc} onBlur={blr}/>

        <div style={{display:"flex",gap:12,marginBottom:24}}>
          <div style={{flex:1}}>
            <label style={T.lbl}>Senha</label>
            <input type="password" value={senha} onChange={e=>setSenha(e.target.value)} required
              placeholder="Mín. 6 caracteres"
              style={{...T.inp}} onFocus={foc} onBlur={blr}/>
          </div>
          <div style={{flex:1}}>
            <label style={T.lbl}>Confirmar senha</label>
            <input type="password" value={conf} onChange={e=>setConf(e.target.value)} required
              placeholder="Repita a senha"
              style={{...T.inp,borderColor:conf&&conf!==senha?"#dc2626":"var(--border)"}} onFocus={foc} onBlur={blr}/>
          </div>
        </div>

        {/* Aviso sobre cargo */}
        <div style={{display:"flex",alignItems:"flex-start",gap:10,padding:"12px 14px",
          background:"var(--blue)12",border:"1px solid var(--blue)30",borderRadius:10,marginBottom:22}}>
          <span style={{fontSize:16,flexShrink:0}}>ℹ️</span>
          <p style={{fontSize:12,color:"var(--text2)",margin:0,lineHeight:1.6}}>
            O cargo e as permissões de acesso são atribuídos pelo administrador do sistema após a criação da conta.
          </p>
        </div>

        <button type="submit" disabled={busy||!nome.trim()||!email.trim()||!senha||!conf}
          style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"13px",fontSize:15,
            opacity:(busy||!nome.trim()||!email.trim()||!senha||!conf)?0.6:1}}>
          {busy?"Criando conta…":"Criar conta"}
        </button>
      </form>
      <p style={{textAlign:"center",fontSize:13,color:"var(--text3)",marginTop:22,marginBottom:0}}>
        Já tem conta?{" "}
        <button onClick={onSwitch}
          style={{background:"none",border:"none",color:"var(--blue)",cursor:"pointer",fontWeight:700,fontSize:13}}>
          Fazer login
        </button>
      </p>
    </AuthLayout>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD PAGE — métricas em tempo real do CRM Monvatti
// ─────────────────────────────────────────────────────────────────────────────

function BarChartSVG({data=[],color="var(--blue)",height=80}){
  if(!data.length)return null;
  const max=Math.max(...data.map(d=>d.value),1);
  const w=100/data.length;
  return(
    <svg width="100%" height={height} style={{overflow:"visible"}}>
      {data.map((d,i)=>{
        const bh=(d.value/max)*(height-20);
        const x=i*w+w*.1;
        const bw=w*.8;
        const y=height-20-bh;
        return(<g key={i}>
          <rect x={`${x}%`} y={y} width={`${bw}%`} height={Math.max(bh,2)}
            fill={d.highlight?color:"var(--border)"} rx={3}
            style={{transition:"height .4s ease,y .4s ease"}}/>
          <text x={`${x+bw/2}%`} y={height-4} textAnchor="middle"
            style={{fontSize:9,fill:"var(--text3)",fontFamily:"system-ui"}}>{d.label}</text>
        </g>);
      })}
    </svg>
  );
}

function DonutSVG({segments=[],size=120}){
  const total=segments.reduce((s,x)=>s+x.value,0)||1;
  let angle=-90;
  const r=40,cx=50,cy=50;
  const p2xy=(deg,radius)=>{
    const rad=(deg*Math.PI)/180;
    return{x:cx+radius*Math.cos(rad),y:cy+radius*Math.sin(rad)};
  };
  const paths=[];
  for(const seg of segments){
    if(!seg.value)continue;
    const sweep=(seg.value/total)*360;
    const large=sweep>180?1:0;
    const start=p2xy(angle,r);
    const end=p2xy(angle+sweep-.5,r);
    paths.push({d:`M ${cx} ${cy} L ${start.x} ${start.y} A ${r} ${r} 0 ${large} 1 ${end.x} ${end.y} Z`,color:seg.color});
    angle+=sweep;
  }
  return(
    <svg viewBox="0 0 100 100" width={size} height={size}>
      <circle cx={cx} cy={cy} r={r} fill="var(--surface2)"/>
      {paths.map((p,i)=><path key={i} d={p.d} fill={p.color} opacity={.9}/>)}
      <circle cx={cx} cy={cy} r={r*.55} fill="var(--surface)"/>
    </svg>
  );
}

function HBarSVG({items=[]}){
  if(!items.length)return null;
  const max=Math.max(...items.map(x=>x.value),1);
  return(
    <div style={{display:"flex",flexDirection:"column",gap:6,width:"100%"}}>
      {items.map((item,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8,width:"100%"}}>
          <div style={{fontSize:11,color:"var(--text2)",width:120,flexShrink:0,textAlign:"right",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.label}</div>
          <div style={{flex:1,background:"var(--surface3)",borderRadius:4,height:14,overflow:"hidden"}}>
            <div style={{height:"100%",width:`${(item.value/max)*100}%`,background:item.color||"var(--blue)",borderRadius:4,transition:"width .4s ease",minWidth:item.value>0?4:0}}/>
          </div>
          <div style={{fontSize:11,fontWeight:700,color:"var(--text)",width:30,textAlign:"right",flexShrink:0}}>{item.value}</div>
        </div>
      ))}
    </div>
  );
}

function KpiCard({title,value,sub,icon,color="#3145FF",progress=null,badge=null}){
  const{isMobile}=useBreakpoint();
  return(
    <div style={{background:"var(--surface)",borderRadius:14,padding:isMobile?"14px 16px":"18px 22px",
      border:"1px solid var(--border)",boxShadow:"0 2px 12px var(--shadow)",
      display:"flex",flexDirection:"column",gap:8,minWidth:0,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-10,right:-10,width:60,height:60,borderRadius:"50%",background:color,opacity:.07}}/>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:8}}>
        <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.6,lineHeight:1.3}}>{title}</div>
        {icon&&<div style={{fontSize:18,flexShrink:0,opacity:.7}}>{icon}</div>}
      </div>
      <div style={{fontSize:isMobile?20:24,fontWeight:900,color:"var(--text)",lineHeight:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{value}</div>
      {progress!=null&&(
        <div>
          <div style={{height:5,background:"var(--surface3)",borderRadius:3,overflow:"hidden",marginBottom:4}}>
            <div style={{height:"100%",width:`${Math.min(progress,100)}%`,background:progress>=100?"#059669":color,borderRadius:3,transition:"width .5s ease"}}/>
          </div>
          <div style={{fontSize:10,color:"var(--text3)"}}>{Math.round(progress)}% qualificados</div>
        </div>
      )}
      {sub&&<div style={{fontSize:11,color:"var(--text3)",marginTop:-2}}>{sub}</div>}
      {badge&&<div style={{display:"inline-flex",alignItems:"center",gap:5,background:(badge.bg||"#05996920"),borderRadius:6,padding:"2px 8px",alignSelf:"flex-start"}}>
        <div style={{width:6,height:6,borderRadius:"50%",background:badge.color||"#059669"}}/>
        <span style={{fontSize:10,fontWeight:700,color:badge.color||"#059669"}}>{badge.label}</span>
      </div>}
    </div>
  );
}

function ChartCard({title,subtitle,children,action}){
  const{isMobile}=useBreakpoint();
  return(
    <div style={{background:"var(--surface)",borderRadius:14,padding:isMobile?"14px":"20px 22px",
      border:"1px solid var(--border)",boxShadow:"0 2px 12px var(--shadow)",
      display:"flex",flexDirection:"column",gap:14,minWidth:0}}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12}}>
        <div>
          <div style={{fontWeight:800,fontSize:13,color:"var(--text)"}}>{title}</div>
          {subtitle&&<div style={{fontSize:11,color:"var(--text3)",marginTop:2}}>{subtitle}</div>}
        </div>
        {action}
      </div>
      {children}
    </div>
  );
}

function DashboardPage({onBack,wsId,allUsers,perms,profile}){
  const{isMobile,width}=useBreakpoint();
  const toast=useToast();
  const containerRef=useRef();

  const todayMes=()=>{const n=new Date();return`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`;};
  const[filterMode,setFilterMode]=useState("month");
  const[selMonth,setSelMonth]=useState(todayMes);
  const[rangeStart,setRangeStart]=useState("");
  const[rangeEnd,setRangeEnd]=useState("");
  const[rawItems,setRawItems]=useState([]);
  const[rawVals,setRawVals]=useState([]);
  const[boards,setBoards]=useState([]);
  const[cols,setCols]=useState([]);
  const[loading,setLoading]=useState(true);
  const[lastUpdate,setLastUpdate]=useState(null);
  const[goal,setGoal]=useState(0);
  const[editGoal,setEditGoal]=useState(false);
  const[goalInput,setGoalInput]=useState("");
  const[savingGoal,setSavingGoal]=useState(false);
  const[fullscreen,setFullscreen]=useState(false);

  const toggleFS=()=>{
    if(!fullscreen)(containerRef.current?.requestFullscreen||containerRef.current?.webkitRequestFullscreen)?.call(containerRef.current);
    else(document.exitFullscreen||document.webkitExitFullscreen)?.call(document);
  };
  useEffect(()=>{
    const h=()=>setFullscreen(!!document.fullscreenElement);
    document.addEventListener("fullscreenchange",h);
    document.addEventListener("webkitfullscreenchange",h);
    return()=>{document.removeEventListener("fullscreenchange",h);document.removeEventListener("webkitfullscreenchange",h);};
  },[]);

  const loadGoal=useCallback(async()=>{
    if(!wsId)return;
    const{data}=await db.from("dashboard_goals").select("meta_valor").eq("workspace_id",wsId).eq("mes",selMonth).maybeSingle();
    setGoal(data?.meta_valor||0);setGoalInput(String(data?.meta_valor||0));
  },[wsId,selMonth]);

  const saveGoal=async()=>{
    if(!wsId)return;setSavingGoal(true);
    const val=parseFloat(goalInput)||0;
    const{error}=await db.from("dashboard_goals").upsert({workspace_id:wsId,mes:selMonth,meta_valor:val,created_by:profile?.id},{onConflict:"workspace_id,mes"});
    if(error)toast("Erro ao salvar meta: "+error.message,"error");
    else{setGoal(val);setEditGoal(false);toast("Meta salva!");}
    setSavingGoal(false);
  };

  const loadData=useCallback(async()=>{
    setLoading(true);
    try{
      const{data:bds}=await db.from("boards").select("id,nome").in("nome",["Vendas","Negociações","Pré - Vendas"]);
      if(!bds?.length){setLoading(false);return;}
      setBoards(bds);
      const boardIds=bds.map(b=>b.id);
      const{data:allCols}=await db.from("columns").select("id,board_id,nome,tipo,config").in("board_id",boardIds);
      setCols(allCols||[]);

      // Items do período selecionado
      let itemQ=db.from("items").select("id,board_id,created_at").in("board_id",boardIds);
      if(filterMode==="month"&&selMonth){
        const[y,m]=selMonth.split("-");
        // Usa meia-noite LOCAL (igual ao inPeriod) para consistência de timezone
        const startLocal=new Date(parseInt(y),parseInt(m)-1,1);
        const endLocal=new Date(parseInt(y),parseInt(m),1);
        itemQ=itemQ.gte("created_at",startLocal.toISOString()).lt("created_at",endLocal.toISOString());
      }else if(filterMode==="range"&&rangeStart&&rangeEnd){
        itemQ=itemQ.gte("created_at",rangeStart+"T00:00:00Z").lte("created_at",rangeEnd+"T23:59:59Z");
      }
      // filterMode==="all": sem filtro, busca todos
      const{data:items}=await itemQ.order("created_at",{ascending:true});

      // Todos negociações (sem filtro) para proposta na rua
      const negB=bds.find(b=>b.nome==="Negociações");
      let allNeg=[];
      if(negB){const{data:ni}=await db.from("items").select("id,board_id,created_at").eq("board_id",negB.id);allNeg=ni||[];}

      // Últimos 8 meses vendas (para gráfico)
      const venB=bds.find(b=>b.nome==="Vendas");
      let allVen=[];
      if(venB){
        const dt8=new Date();dt8.setMonth(dt8.getMonth()-7);dt8.setDate(1);dt8.setHours(0,0,0,0);
        const{data:vi}=await db.from("items").select("id,board_id,created_at").eq("board_id",venB.id).gte("created_at",dt8.toISOString());
        allVen=vi||[];
      }

      // Todos pré-vendas (para MQL total)
      const preB=bds.find(b=>b.nome==="Pré - Vendas");
      let allPre=[];
      if(preB){const{data:pi}=await db.from("items").select("id,board_id,created_at").eq("board_id",preB.id);allPre=pi||[];}

      // ── DEDUPLICAR por id antes de salvar ────────────────────────────────
      const uniqMap=new Map();
      for(const it of [...(items||[]),...allNeg,...allVen,...allPre]) uniqMap.set(it.id,it);
      const uniqItems=[...uniqMap.values()];
      const uniqIds=uniqItems.map(x=>x.id);
      setRawItems(uniqItems);

      if(uniqIds.length){
        // Busca item_values em chunks de 100 IDs por request com limit alto explícito
        // para evitar o corte silencioso do Supabase (default 1000 linhas por query)
        const CHUNK=100;
        const chunks=[];
        for(let i=0;i<uniqIds.length;i+=CHUNK) chunks.push(uniqIds.slice(i,i+CHUNK));
        const valsArr=await Promise.all(
          chunks.map(chunk=>
            db.from("item_values")
              .select("item_id,column_id,value")
              .in("item_id",chunk)
              .limit(10000)   // explícito: nunca deixar o default 1000 cortar resultados
          )
        );
        // Deduplicar item_values por (item_id, column_id) — pega o último
        const vMap=new Map();
        for(const v of valsArr.flatMap(r=>r.data||[])) vMap.set(v.item_id+"_"+v.column_id,v);
        setRawVals([...vMap.values()]);
      }else setRawVals([]);

      setLastUpdate(new Date());
    }catch(e){console.error(e);toast("Erro ao carregar dashboard","error");}
    setLoading(false);
  },[filterMode,selMonth,rangeStart,rangeEnd]);

  useEffect(()=>{loadData();loadGoal();},[loadData,loadGoal]);

  useEffect(()=>{
    const ch=db.channel("dash-rt").on("postgres_changes",{event:"*",schema:"public",table:"items"},()=>loadData()).on("postgres_changes",{event:"*",schema:"public",table:"item_values"},()=>loadData()).subscribe();
    return()=>db.removeChannel(ch);
  },[loadData]);

  const metrics=useMemo(()=>{
    if(!boards.length||!cols.length)return{};
    const gB=nome=>boards.find(b=>b.nome===nome);
    const gC=(bNome,cNome,tipo)=>cols.find(c=>{const b=gB(bNome);return b&&c.board_id===b.id&&(cNome?c.nome===cNome:true)&&(tipo?c.tipo===tipo:true);});

    // Índice de valores: {itemId: {colId: value}} — acesso O(1), sem buscas repetidas
    const valIdx={};
    for(const row of rawVals){
      if(!valIdx[row.item_id]) valIdx[row.item_id]={};
      const v=row.value;
      valIdx[row.item_id][row.column_id]=(v!==null&&v!==undefined&&typeof v==="object"&&"value" in v)?v.value:v;
    }
    const gV=(itemId,colId)=>colId?(valIdx[itemId]?.[colId]??null):null;

    // Itens "preenchidos": têm ao menos 1 valor não-nulo no item_values
    const itemsComValor=new Set(rawVals.filter(r=>r.value!==null&&r.value!==undefined&&r.value!=="").map(r=>r.item_id));
    const isPreenchido=id=>itemsComValor.has(id);

    const venB=gB("Vendas"),negB=gB("Negociações"),preB=gB("Pré - Vendas");
    const venValC=gC("Vendas","Valor do Projeto","currency");
    const venParC=gC("Vendas","Parcelas","number");
    const negValC=gC("Negociações","Valor do Projeto","currency");
    const negEtapaC=gC("Negociações","Etapa","status");
    const negOriC=gC("Negociações","Origem","status");
    const preMqlC=gC("Pré - Vendas","MQL","status");

    // Função de filtro de período
    const inPeriod=it=>{
      if(filterMode==="all") return true;
      if(filterMode==="month"&&selMonth){
        const[y,m]=selMonth.split("-");
        const s=new Date(parseInt(y),parseInt(m)-1,1);
        const e=new Date(parseInt(y),parseInt(m),1);
        const d=new Date(it.created_at);return d>=s&&d<e;
      }
      if(filterMode==="range"&&rangeStart&&rangeEnd){
        const s=new Date(rangeStart);const e=new Date(rangeEnd+"T23:59:59Z");
        const d=new Date(it.created_at);return d>=s&&d<=e;
      }
      // filterMode definido mas sem parâmetros suficientes → não exibe nada
      return false;
    };

    // Conjuntos únicos por board (rawItems já está deduplicado)
    // Apenas itens com pelo menos 1 valor preenchido
    const venFilt  = rawItems.filter(i=>i.board_id===venB?.id && inPeriod(i) && isPreenchido(i.id));
    const negAll   = rawItems.filter(i=>i.board_id===negB?.id && isPreenchido(i.id));   // sem filtro data — proposta na rua
    const negFilt  = rawItems.filter(i=>i.board_id===negB?.id && inPeriod(i) && isPreenchido(i.id));
    const preAll   = rawItems.filter(i=>i.board_id===preB?.id && isPreenchido(i.id));   // sem filtro — MQL total
    const preFilt  = rawItems.filter(i=>i.board_id===preB?.id && inPeriod(i) && isPreenchido(i.id));
    const venAll8m = rawItems.filter(i=>i.board_id===venB?.id && isPreenchido(i.id));   // sem filtro data — gráfico 8 meses

    // ── Métricas financeiras (todas vindas do quadro Vendas) ──────────────────
    const receitaEfetiva=venFilt.reduce((s,it)=>s+(parseFloat(gV(it.id,venValC?.id))||0),0);
    const mrr=venFilt.reduce((s,it)=>{
      const v=parseFloat(gV(it.id,venValC?.id))||0;
      const p=parseFloat(gV(it.id,venParC?.id))||1;
      return s+(p>0?v/p:0);
    },0);
    // Ticket médio = receita / nº de vendas com valor > 0
    const venComValor=venFilt.filter(it=>(parseFloat(gV(it.id,venValC?.id))||0)>0);
    const ticketMedio=venComValor.length>0?receitaEfetiva/venComValor.length:0;
    // Valor médio fechado = média das vendas no board Vendas (não Negociações)
    const valMedioFechados=venComValor.length>0?receitaEfetiva/venComValor.length:0;

    // ── Proposta na rua (estado atual = Negociando, todos, sem filtro data) ─────
    const propostaNaRua=negAll.filter(it=>(gV(it.id,negEtapaC?.id)||"")==="Negociando")
      .reduce((s,it)=>s+(parseFloat(gV(it.id,negValC?.id))||0),0);

    // ── Negócios fechados em Negociações (para funil) ─────────────────────────
    const negFechadosItems=negFilt.filter(it=>gV(it.id,negEtapaC?.id)==="Negócio Fechado");

    // ── Etapas ────────────────────────────────────────────────────────────────
    const etapaOpts=negEtapaC?.config?.options||[];
    const etapasMap={};
    // Apenas itens que têm etapa definida
    negFilt.filter(it=>!!gV(it.id,negEtapaC?.id)).forEach(it=>{
      const e=gV(it.id,negEtapaC?.id);
      etapasMap[e]=(etapasMap[e]||0)+1;
    });
    const etapasData=etapaOpts.length
      ?etapaOpts.map(o=>({label:o.label,value:etapasMap[o.label]||0,color:o.color})).filter(x=>x.value>0)
      :Object.entries(etapasMap).map(([k,v])=>({label:k,value:v,color:"#3145FF"}));

    // ── Origens ───────────────────────────────────────────────────────────────
    const oriOpts=negOriC?.config?.options||[];
    const oriMap={};
    // Apenas itens com origem definida
    negFilt.filter(it=>!!gV(it.id,negOriC?.id)).forEach(it=>{
      const o=gV(it.id,negOriC?.id);
      oriMap[o]=(oriMap[o]||0)+1;
    });
    const totalOri=Object.values(oriMap).reduce((s,v)=>s+v,0)||1;
    const origemData=Object.entries(oriMap).map(([label,value])=>({
      label,value,pct:Math.round((value/totalOri)*100),
      color:oriOpts.find(o=>o.label===label)?.color||"#94a3b8"
    })).sort((a,b)=>b.value-a.value);

    // ── MQL (todos os leads de Pré-Vendas, sem filtro de data) ───────────────
    let mqlQ=0,mqlDQ=0;
    preAll.forEach(it=>{
      const m=gV(it.id,preMqlC?.id);
      if(m==="Qualificado")mqlQ++;
      else if(m==="Desqualificado"||m==="Não Qualificado")mqlDQ++;
    });

    // ── Receita por mês (últimos 8 meses, board Vendas) ───────────────────────
    const mesMap={};
    const now=new Date();
    for(let i=7;i>=0;i--){
      const d=new Date(now.getFullYear(),now.getMonth()-i,1);
      const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      mesMap[k]=0;
    }
    venAll8m.forEach(it=>{
      const d=new Date(it.created_at);
      const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      if(k in mesMap){mesMap[k]+=(parseFloat(gV(it.id,venValC?.id))||0);}
    });
    const receitaMesData=Object.entries(mesMap).map(([k,value])=>({label:k.slice(5),fullLabel:k,value,highlight:k===selMonth}));

    // ── Taxa de conversão: Vendas realizadas / Leads Pré-Vendas × 100 ─────────
    // Definição clara: de cada 100 leads que entraram no Pré-Vendas no período,
    // quantos viraram uma venda efetiva no board Vendas?
    const totalLeads=preFilt.length;
    const taxaConversao=totalLeads>0?(venFilt.length/totalLeads*100):0;

    return{
      receitaEfetiva,mrr,ticketMedio,valMedioFechados,
      negFechados:negFechadosItems.length,
      propostaNaRua,etapasData,origemData,
      mqlQ,mqlDQ,receitaMesData,
      totalLeads,taxaConversao,
      venFiltLen:venComValor.length,
      negFiltLen:negFilt.length,
      preAllLen:preAll.length
    };
  },[rawItems,rawVals,boards,cols,filterMode,selMonth,rangeStart,rangeEnd]);

  const monthOptions=useMemo(()=>{
    const opts=[];const now=new Date();
    for(let i=11;i>=0;i--){const d=new Date(now.getFullYear(),now.getMonth()-i,1);const val=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;const label=d.toLocaleDateString("pt-BR",{month:"short",year:"numeric"}).replace("de ","");opts.push({val,label});}
    return opts;
  },[]);

  const canEditGoal=perms?.all||perms?.manageBoards||perms?.editAny||perms?.isFull;
  const progresso=goal>0?(metrics.receitaEfetiva||0)/goal*100:0;
  const gridCols=width>=1400?"repeat(4,1fr)":width>=1024?"repeat(3,1fr)":width>=640?"repeat(2,1fr)":"1fr";
  const chartCols=width>=1200?"repeat(3,1fr)":width>=768?"repeat(2,1fr)":"1fr";

  return(
    <div ref={containerRef} style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",background:"var(--bg)"}}>

      {/* Header */}
      <div style={{background:"var(--surface)",borderBottom:"1px solid var(--border)",padding:isMobile?"12px 16px":"14px 24px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",flexShrink:0,zIndex:10}}>
        {!fullscreen&&<button onClick={onBack} style={{...T.btn,background:"none",border:"1.5px solid var(--border)",color:"var(--text)",padding:"7px 14px",fontSize:12,flexShrink:0}}>← Voltar</button>}
        <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
          <div style={{fontSize:isMobile?16:18,fontWeight:900,color:"var(--text)"}}>📊 Dashboard</div>
          {loading&&<Spinner size={15}/>}
          {!loading&&lastUpdate&&<div style={{fontSize:10,color:"var(--text3)",background:"var(--surface3)",borderRadius:6,padding:"2px 8px"}}>⟳ {lastUpdate.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}</div>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",flex:1}}>
          <div style={{display:"flex",borderRadius:8,overflow:"hidden",border:"1.5px solid var(--border)",flexShrink:0}}>
            {[["month","Mês"],["range","Período"],["all","Todos"]].map(([mode,lbl])=>(
              <button key={mode} onClick={()=>setFilterMode(mode)} style={{...T.btn,borderRadius:0,padding:"6px 12px",fontSize:11,background:filterMode===mode?"var(--blue)":"var(--surface)",color:filterMode===mode?"#fff":"var(--text2)"}}>{lbl}</button>
            ))}
          </div>
          {filterMode==="month"&&<select value={selMonth} onChange={e=>setSelMonth(e.target.value)} style={{...T.inp,width:"auto",padding:"7px 10px",fontSize:12,borderRadius:8}}>{monthOptions.map(o=><option key={o.val} value={o.val}>{o.label}</option>)}</select>}
          {filterMode==="range"&&<div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
            <input type="date" value={rangeStart} onChange={e=>setRangeStart(e.target.value)} style={{...T.inp,width:140,padding:"7px 10px",fontSize:12,borderRadius:8}}/>
            <span style={{color:"var(--text3)",fontSize:12}}>até</span>
            <input type="date" value={rangeEnd} onChange={e=>setRangeEnd(e.target.value)} style={{...T.inp,width:140,padding:"7px 10px",fontSize:12,borderRadius:8}}/>
          </div>}
        </div>
        <div style={{display:"flex",gap:8,flexShrink:0}}>
          <button onClick={loadData} title="Atualizar" style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"7px 11px",fontSize:14,border:"1px solid var(--border)"}}>↻</button>
          <button onClick={toggleFS} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"7px 14px",fontSize:12}}>{fullscreen?"⊠ Sair":"⛶ Tela Cheia"}</button>
        </div>
      </div>

      {/* Conteúdo */}
      <div style={{flex:1,overflowY:"auto",padding:isMobile?"12px":"20px 24px"}}>
        {loading&&!metrics.receitaEfetiva&&<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:200,gap:12,flexDirection:"column"}}><Spinner size={36}/><span style={{color:"var(--text3)",fontSize:14}}>Carregando dados…</span></div>}

        {/* Meta Mensal */}
        <div style={{background:"var(--surface)",borderRadius:14,padding:isMobile?"14px":"20px",border:"1px solid var(--border)",boxShadow:"0 2px 12px var(--shadow)",marginBottom:16}}>
          <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12,flexWrap:"wrap",marginBottom:16}}>
            <div>
              <div style={{fontWeight:900,fontSize:isMobile?15:18,color:"var(--text)"}}>🎯 Meta Mensal</div>
              <div style={{fontSize:11,color:"var(--text3)",marginTop:2}}>{filterMode==="month"?monthOptions.find(m=>m.val===selMonth)?.label||selMonth:"Período selecionado"}</div>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
              {canEditGoal&&!editGoal&&<button onClick={()=>{setEditGoal(true);setGoalInput(String(goal));}} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"7px 14px",fontSize:12,border:"1px solid var(--border)"}}>{goal>0?"✏️ Editar meta":"+ Definir meta"}</button>}
              {editGoal&&<div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <div style={{position:"relative"}}>
                  <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"var(--text3)",fontSize:13}}>R$</span>
                  <input type="number" value={goalInput} onChange={e=>setGoalInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")saveGoal();if(e.key==="Escape")setEditGoal(false);}} style={{...T.inp,width:160,paddingLeft:28}} autoFocus placeholder="0,00"/>
                </div>
                <button onClick={saveGoal} disabled={savingGoal} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 16px",fontSize:12}}>{savingGoal?"…":"Salvar"}</button>
                <button onClick={()=>setEditGoal(false)} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"8px 14px",fontSize:12,border:"1px solid var(--border)"}}>Cancelar</button>
              </div>}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:goal>0?16:0}}>
            {[
              {lbl:"Receita Efetiva",val:fmtBRL(metrics.receitaEfetiva||0),clr:"#059669"},
              {lbl:"Meta",val:goal>0?fmtBRL(goal):"—",clr:"var(--text)"},
              {lbl:"Restante",val:goal>0?fmtBRL(Math.max(goal-(metrics.receitaEfetiva||0),0)):"—",clr:goal>0&&(metrics.receitaEfetiva||0)>=goal?"#059669":"#dc2626"},
              {lbl:"Atingimento",val:goal>0?`${Math.round(progresso)}%`:"—",clr:progresso>=100?"#059669":progresso>=70?"#d97706":"#dc2626"},
            ].map((item,i)=>(
              <div key={i}>
                <div style={{fontSize:10,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>{item.lbl}</div>
                <div style={{fontSize:isMobile?18:24,fontWeight:900,color:item.clr}}>{item.val}</div>
              </div>
            ))}
          </div>
          {goal>0&&<div>
            <div style={{height:10,background:"var(--surface3)",borderRadius:5,overflow:"hidden",marginBottom:6}}>
              <div style={{height:"100%",width:`${Math.min(progresso,100)}%`,background:progresso>=100?"#059669":progresso>=70?"#d97706":"var(--blue)",borderRadius:5,transition:"width .6s ease"}}/>
            </div>
            <div style={{fontSize:11,color:"var(--text3)"}}>{progresso>=100?"🎉 Meta atingida!":progresso>=70?`Quase lá — faltam ${fmtBRL(goal-(metrics.receitaEfetiva||0))}`:`${Math.round(100-progresso)}% abaixo da meta`}</div>
          </div>}
        </div>

        {/* KPIs */}
        <div style={{display:"grid",gridTemplateColumns:gridCols,gap:12,marginBottom:16}}>
          <KpiCard title="MRR" value={fmtBRL(metrics.mrr||0)} icon="🔄" color="#059669" sub="Média recorrente mensal"/>
          <KpiCard title="Proposta na Rua" value={fmtBRL(metrics.propostaNaRua||0)} icon="📤" color="#d97706" sub="Negociando (em aberto, todos)" badge={{label:"Ativo",color:"#d97706",bg:"#d9770620"}}/>
          <KpiCard title="Ticket Médio" value={fmtBRL(metrics.ticketMedio||0)} icon="💰" color="#7c3aed" sub={`${metrics.venFiltLen||0} vendas no período`}/>
          <KpiCard title="Valor Médio Fechado" value={fmtBRL(metrics.valMedioFechados||0)} icon="🤝" color="#0891b2" sub={`${metrics.negFechados||0} negócios fechados`}/>
          <KpiCard title="Total Leads" value={metrics.totalLeads||0} icon="👥" color="#3145FF" sub="Pré-Vendas no período"/>
          <KpiCard title="Taxa de Conversão" value={`${(metrics.taxaConversao||0).toFixed(1)}%`} icon="📈" color="#059669" sub="Leads → Negócio Fechado"/>
          <KpiCard title="MQL Qualificados" value={metrics.mqlQ||0} icon="✅" color="#059669" sub={`De ${(metrics.mqlQ||0)+(metrics.mqlDQ||0)} avaliados`} progress={(metrics.mqlQ||0)+(metrics.mqlDQ||0)>0?((metrics.mqlQ||0)/((metrics.mqlQ||0)+(metrics.mqlDQ||0)))*100:null}/>
          <KpiCard title="Negócios Fechados" value={metrics.negFechados||0} icon="🏆" color="#059669" sub="Negociações no período"/>
        </div>

        {/* Gráficos linha 1 */}
        <div style={{display:"grid",gridTemplateColumns:chartCols,gap:16,marginBottom:16}}>

          <ChartCard title="📅 Receita por Mês" subtitle="Últimos 8 meses — Vendas">
            {(metrics.receitaMesData||[]).length>0?(
              <><BarChartSVG data={metrics.receitaMesData||[]} color="var(--blue)" height={90}/>
                <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:4}}>
                  {(metrics.receitaMesData||[]).filter(d=>d.value>0).slice(-4).map((d,i)=>(
                    <div key={i} style={{fontSize:10,color:"var(--text3)",background:"var(--surface3)",borderRadius:6,padding:"3px 8px"}}>
                      <strong>{d.label}:</strong> {fmtBRL(d.value)}
                    </div>
                  ))}
                </div>
              </>
            ):<div style={{color:"var(--text3)",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sem dados</div>}
          </ChartCard>

          <ChartCard title="🌐 Origens de Lead" subtitle="Negociações — distribuição">
            {(metrics.origemData||[]).length>0?(
              <div style={{display:"flex",gap:12,alignItems:"flex-start",flexWrap:"wrap"}}>
                <DonutSVG segments={metrics.origemData||[]} size={isMobile?90:110}/>
                <div style={{flex:1,display:"flex",flexDirection:"column",gap:5,minWidth:0}}>
                  {(metrics.origemData||[]).slice(0,7).map((s,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:7}}>
                      <div style={{width:9,height:9,borderRadius:2,background:s.color,flexShrink:0}}/>
                      <div style={{fontSize:11,color:"var(--text2)",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.label}</div>
                      <div style={{fontSize:11,fontWeight:700,color:"var(--text)",flexShrink:0}}>{s.pct}%</div>
                    </div>
                  ))}
                </div>
              </div>
            ):<div style={{color:"var(--text3)",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sem dados de origem</div>}
          </ChartCard>

          <ChartCard title="🔄 Etapas — Negociações" subtitle="Leads por etapa no período">
            {(metrics.etapasData||[]).length>0?<HBarSVG items={metrics.etapasData||[]}/>:<div style={{color:"var(--text3)",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sem dados de etapas</div>}
          </ChartCard>

        </div>

        {/* Gráficos linha 2 */}
        <div style={{display:"grid",gridTemplateColumns:width>=1024?"repeat(2,1fr)":"1fr",gap:16,marginBottom:16}}>

          <ChartCard title="🎯 MQL — Qualificação (Pré-Vendas)" subtitle="Todos os leads avaliados">
            <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
              <DonutSVG segments={[
                {label:"Qualificado",value:metrics.mqlQ||0,color:"#059669"},
                {label:"Desqualificado",value:metrics.mqlDQ||0,color:"#dc2626"},
                {label:"Sem avaliação",value:Math.max((metrics.preAllLen||0)-(metrics.mqlQ||0)-(metrics.mqlDQ||0),0),color:"var(--border)"}
              ]} size={100}/>
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:10,minWidth:0}}>
                {[{label:"Qualificados",value:metrics.mqlQ||0,color:"#059669"},{label:"Desqualificados",value:metrics.mqlDQ||0,color:"#dc2626"},{label:"Sem avaliação",value:Math.max((metrics.preAllLen||0)-(metrics.mqlQ||0)-(metrics.mqlDQ||0),0),color:"var(--border)"}].map((row,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:row.color,flexShrink:0}}/>
                    <div style={{flex:1,fontSize:12,color:"var(--text2)"}}>{row.label}</div>
                    <div style={{fontSize:14,fontWeight:800,color:"var(--text)"}}>{row.value}</div>
                  </div>
                ))}
              </div>
            </div>
          </ChartCard>

          <ChartCard title="🔽 Funil de Conversão" subtitle="Leads → Vendas no período">
            {(()=>{
              const preB2=boards.find(b=>b.nome==="Pré - Vendas");
              const negB2=boards.find(b=>b.nome==="Negociações");
              const venB2=boards.find(b=>b.nome==="Vendas");
              const inPeriod=item=>{
                if(filterMode==="all") return true;
                if(filterMode==="month"&&selMonth){const[y,m]=selMonth.split("-");const s=new Date(parseInt(y),parseInt(m)-1,1);const e=new Date(parseInt(y),parseInt(m),1);const d=new Date(item.created_at);return d>=s&&d<e;}
                if(filterMode==="range"&&rangeStart&&rangeEnd){const s=new Date(rangeStart);const e=new Date(rangeEnd+"T23:59:59Z");const d=new Date(item.created_at);return d>=s&&d<=e;}
                return false;
              };
              const steps=[
                {label:"Pré-Vendas (Leads)",value:rawItems.filter(i=>i.board_id===preB2?.id&&inPeriod(i)).length,color:"#3145FF"},
                {label:"Negociações",value:rawItems.filter(i=>i.board_id===negB2?.id&&inPeriod(i)).length,color:"#7c3aed"},
                {label:"Negócio Fechado",value:metrics.negFechados||0,color:"#059669"},
                {label:"Vendas Realizadas",value:rawItems.filter(i=>i.board_id===venB2?.id&&inPeriod(i)).length,color:"#0891b2"},
              ];
              const maxV=Math.max(...steps.map(s=>s.value),1);
              return(
                <div style={{display:"flex",flexDirection:"column",gap:10}}>
                  {steps.map((step,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
                      <div style={{width:20,height:20,borderRadius:6,background:step.color,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,color:"#fff",fontWeight:800}}>{i+1}</div>
                      <div style={{flex:1}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                          <span style={{fontSize:11,color:"var(--text2)"}}>{step.label}</span>
                          <span style={{fontSize:12,fontWeight:800,color:"var(--text)"}}>{step.value}</span>
                        </div>
                        <div style={{height:10,background:"var(--surface3)",borderRadius:4,overflow:"hidden"}}>
                          <div style={{height:"100%",width:`${(step.value/maxV)*100}%`,background:step.color,borderRadius:4,transition:"width .5s ease"}}/>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              );
            })()}
          </ChartCard>
        </div>

      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// APP CONTENT
// ─────────────────────────────────────────────────────────────────────────────
function AppContent() {
  const {session,profile,authLoading,signOut}=useAuth();
  const {dark,toggle:toggleDark}=useTheme();
  const toast=useToast();
  const {isMobile}=useBreakpoint();
  const [authPage,setAuthPage]=useState("login");
  const [page,setPage]=useState("board"); // "board"|"profile"|"admin"|"dashboard"
  const [boards,setBoards]=useState([]);
  const [boardId,setBoardId]=useState(null);
  const [allUsers,setAllUsers]=useState([]);
  const [wsId,setWsId]=useState(null);
  const [wsNome,setWsNome]=useState("");
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [confirmM,setConfirmM]=useState(null);
  const [boardFormM,setBoardFormM]=useState(null);
  const [boardSettingsM,setBoardSettingsM]=useState(null); // board obj
  const [boardAccess,setBoardAccess]=useState({}); // {boardId:[userId,...]}

  // Permissões derivadas do cargo do perfil logado
  const perms=useMemo(()=>getPerms(profile),[profile]);

  // Boards visíveis conforme privacidade
  const visibleBoards=useMemo(()=>{
    if(perms.all||perms.viewAll) return boards;
    return boards.filter(b=>!b.is_private||(boardAccess[b.id]||[]).includes(profile?.id));
  },[boards,perms,boardAccess,profile]);

  // Carrega usuários com join de roles
  useEffect(()=>{
    if(!session) return;
    db.from("profiles").select("*, roles(id,nome,slug,cor)").order("nome")
      .then(({data})=>setAllUsers((data||[]).map(u=>({...u,role_nome:u.roles?.nome||u.funcao||"",role_slug:u.roles?.slug||"sdr"}))));
  },[session]);

  const loadBoards=useCallback(async()=>{
    if(!session) return;
    const {data:bds}=await db.from("boards").select("*").order("ordem");
    if(!bds) return;
    if(bds.length){
      const {data:ws}=await db.from("workspaces").select("*").eq("id",bds[0].workspace_id).single();
      if(ws){setWsId(ws.id);setWsNome(ws.nome);}
    }
    const counts=await Promise.all(bds.map(async b=>{
      const{count}=await db.from("items").select("*",{count:"exact",head:true}).eq("board_id",b.id);
      return{id:b.id,total:count||0};
    }));
    const cmap=Object.fromEntries(counts.map(c=>[c.id,c.total]));
    setBoards(bds.map(b=>({...b,totalItems:cmap[b.id]||0})));
    if(!boardId&&bds.length) setBoardId(bds[0].id);
  },[session,boardId]);

  useEffect(()=>{loadBoards();},[session]);

  const bumpCount=(bid,d)=>setBoards(prev=>prev.map(b=>b.id===bid?{...b,totalItems:Math.max(0,(b.totalItems||0)+d)}:b));

  // Carrega acessos de quadros privados
  const loadBoardAccess=useCallback(async(bids)=>{
    if(!bids?.length) return;
    const {data}=await db.from("board_access").select("*").in("board_id",bids);
    const map={};
    for(const a of data||[]){if(!map[a.board_id])map[a.board_id]=[];map[a.board_id].push(a.user_id);}
    setBoardAccess(map);
  },[]);

  // Salva configurações de acesso do quadro
  const saveBoardSettings=async(board,{isPrivate,access})=>{
    const {error:updErr}=await db.from("boards").update({is_private:isPrivate}).eq("id",board.id);
    if(updErr){toast("Erro ao salvar visibilidade: "+updErr.message,"error");return;}
    const {error:delErr}=await db.from("board_access").delete().eq("board_id",board.id);
    if(delErr){toast("Erro ao limpar acessos: "+delErr.message,"error");return;}
    if(isPrivate&&access.length){
      const rows=access.map(uid=>({board_id:board.id,user_id:uid}));
      const {error:insErr}=await db.from("board_access").upsert(rows,{onConflict:"board_id,user_id",ignoreDuplicates:true});
      if(insErr){toast("Erro ao salvar acessos: "+insErr.message,"error");return;}
    }
    setBoardAccess(p=>({...p,[board.id]:isPrivate?access:[]}));
    setBoards(p=>p.map(b=>b.id===board.id?{...b,is_private:isPrivate}:b));
    setBoardSettingsM(null);
    toast("Configurações do quadro salvas!");
  };

  // Reordena boards com DnD
  const reorderBoards=async(newOrder)=>{
    setBoards(newOrder.map((b,i)=>({...b,ordem:i})));
    await Promise.all(newOrder.map((b,i)=>db.from("boards").update({ordem:i}).eq("id",b.id)));
  };

  // Duplica quadro completo (grupos + colunas, sem os items)
  const duplicateBoard=async(srcBoard)=>{
    toast("Duplicando quadro…","info");
    // Cria novo board
    const {data:nb,error}=await db.from("boards")
      .insert({nome:srcBoard.nome+" (cópia)",icon:srcBoard.icon,workspace_id:wsId,ordem:boards.length,is_private:false})
      .select().single();
    if(error||!nb){toast("Erro ao duplicar quadro","error");return;}

    // Copia colunas
    const {data:srcCols}=await db.from("columns").select("*").eq("board_id",srcBoard.id).order("ordem");
    const colIdMap={}; // srcColId → newColId
    for(const col of srcCols||[]){
      const {data:nc}=await db.from("columns")
        .insert({board_id:nb.id,nome:col.nome,tipo:col.tipo,config:col.config,ordem:col.ordem})
        .select().single();
      if(nc) colIdMap[col.id]=nc.id;
    }

    // Copia grupos (sem itens)
    const {data:srcGroups}=await db.from("groups").select("*").eq("board_id",srcBoard.id).order("ordem");
    for(const g of srcGroups||[]){
      await db.from("groups").insert({board_id:nb.id,nome:g.nome,color:g.color,ordem:g.ordem});
    }

    setBoards(p=>[...p,{...nb,totalItems:0}]);
    toast("Quadro duplicado com sucesso!");
  };

  const addBoard=async({nome,icon})=>{
    if(!wsId){toast("Workspace não encontrado","error");setBoardFormM(null);return;}
    const{data,error}=await db.from("boards").insert({nome,icon,workspace_id:wsId,ordem:boards.length}).select().single();
    if(error){toast(error.message,"error");return;}
    setBoards(p=>[...p,{...data,totalItems:0}]);setBoardId(data.id);setBoardFormM(null);toast("Quadro criado!");
  };
  const editBoard=async({nome,icon})=>{
    const b=boardFormM;
    await db.from("boards").update({nome,icon}).eq("id",b.id);
    setBoards(p=>p.map(x=>x.id===b.id?{...x,nome,icon}:x));
    setBoardFormM(null);toast("Quadro atualizado!");
  };
  const delBoard=b=>{
    setConfirmM({title:"Excluir quadro",danger:true,message:`Excluir "${b.nome}" e todos os dados?`,
      onConfirm:async()=>{
        await db.from("boards").delete().eq("id",b.id);
        const rem=boards.filter(x=>x.id!==b.id);setBoards(rem);
        if(boardId===b.id)setBoardId(rem[0]?.id||null);
        toast("Quadro excluído");setConfirmM(null);
      }
    });
  };

  if(authLoading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--bg)",flexDirection:"column",gap:16}}>
      <Spinner size={44}/><span style={{color:"var(--text3)",fontSize:14}}>Iniciando…</span>
    </div>
  );

  if(!session) return authPage==="login"
    ?<LoginPage onSwitch={()=>setAuthPage("register")}/>
    :<RegisterPage onSwitch={()=>setAuthPage("login")}/>;

  const sidebar=(
    <Sidebar boards={boards} currentBoardId={boardId} currentUser={profile}
      wsNome={wsNome} dark={dark} onToggleDark={toggleDark} perms={perms}
      onSelectBoard={id=>{setBoardId(id);setPage("board");}}
      onAddBoard={()=>perms.manageBoards&&setBoardFormM({})}
      onEditBoard={b=>perms.manageBoards&&setBoardFormM(b)}
      onDelBoard={b=>perms.manageBoards&&delBoard(b)}
      onReorderBoards={reorderBoards}
      onBoardSettings={b=>perms.manageBoards&&setBoardSettingsM(b)}
      onDuplicateBoard={b=>perms.manageBoards&&duplicateBoard(b)}
      onProfile={()=>setPage("profile")}
      onAdmin={()=>setPage("admin")}
      onDashboard={()=>setPage("dashboard")}
      onLogout={signOut}
      open={sidebarOpen} onClose={()=>setSidebarOpen(false)}
    />
  );

  return (
    <div style={{display:"flex",height:"100vh",width:"100%",overflow:"hidden",background:"var(--bg)"}}>
      {!isMobile&&sidebar}
      {isMobile&&sidebar}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
        {/* Mobile top bar */}
        {isMobile&&<div style={{background:"var(--surface)",borderBottom:"1px solid var(--border)",padding:"0 16px",height:52,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <button onClick={()=>setSidebarOpen(true)} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"var(--text)",lineHeight:1}}>☰</button>
          <div style={{fontWeight:800,fontSize:15,color:"var(--text)",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            {page==="dashboard"?"📊 Dashboard":page==="profile"?"👤 Perfil":page==="admin"?"⚙️ Admin":(boards.find(b=>b.id===boardId)?.icon||"📋")+" "+(boards.find(b=>b.id===boardId)?.nome||"CRM")}
          </div>
        </div>}
        {page==="profile"&&<ProfilePage onBack={()=>setPage("board")}/>}
        {page==="admin"&&(perms.all||perms.manageUsers)
          ?<AdminPanel onBack={()=>setPage("board")}/>
          :page==="admin"&&<div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",color:"var(--text3)",fontSize:14,flexDirection:"column",gap:12}}>
              <div style={{fontSize:32}}>🔒</div>
              <div>Sem permissão para acessar esta área.</div>
              <button onClick={()=>setPage("board")} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"9px 20px",marginTop:8}}>Voltar</button>
            </div>
        }
        {page==="dashboard"&&(
          <DashboardPage onBack={()=>setPage("board")} wsId={wsId} allUsers={allUsers} perms={perms} profile={profile}/>
        )}
        {page==="board"&&(
          <BoardView key={boardId} boardId={boardId} boards={visibleBoards} allUsers={allUsers}
            currentUser={profile} wsId={wsId} perms={perms} onBoardCountChange={bumpCount}/>
        )}
      </div>
      {confirmM&&<ConfirmModal title={confirmM.title} message={confirmM.message} danger={confirmM.danger} onConfirm={confirmM.onConfirm} onCancel={()=>setConfirmM(null)}/>}
      {boardFormM!==null&&<BoardFormModal initial={boardFormM} onCancel={()=>setBoardFormM(null)} onSave={boardFormM?.id?editBoard:addBoard}/>}
      {boardSettingsM&&<BoardSettingsModal board={boardSettingsM} allUsers={allUsers}
        currentAccess={boardAccess[boardSettingsM.id]||[]}
        isPrivate={!!boardSettingsM.is_private}
        onSave={s=>saveBoardSettings(boardSettingsM,s)}
        onClose={()=>setBoardSettingsM(null)}/>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ROOT
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  return (
    <ThemeProvider>
      <AuthProvider>
        <ToastProvider>
          <AppContent/>
        </ToastProvider>
      </AuthProvider>
    </ThemeProvider>
  );
}
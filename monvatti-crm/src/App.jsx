/**
 * MONVATTI CRM V4
 * Novidades: dark mode · responsivo (mobile) · portal dropdowns (fix status/user)
 * exportação XLSX + PDF · upload de foto · column DnD · filtro por data
 * duplicate/move copia campos equivalentes + atualizações
 */
import {
  useState, useEffect, useCallback, useRef,
  createContext, useContext, useMemo, Fragment
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
  { nome: "Isaías",    phone: "554896321559",             apikey: "1681947"                 }, // ← preencher
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

import * as XLSX_MOD from "https://esm.sh/xlsx-js-style@1.2.0";
// esm.sh pode expor o pacote como named exports OU sob .default — resolve o objeto que tem .utils
const XLSX = (XLSX_MOD && XLSX_MOD.utils) ? XLSX_MOD : (XLSX_MOD?.default || XLSX_MOD);

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
  {v:"phone",l:"Telefone"},{v:"link",l:"Link"},{v:"cnpj",l:"CNPJ"},
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
  const [isRecovery,setIsRecovery] = useState(false); // true quando vem do link de redefinição
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
    const {data:{subscription}}=db.auth.onAuthStateChange((event,s)=>{
      if(event==="PASSWORD_RECOVERY") setIsRecovery(true);
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
  const resetPassword  = email=>db.auth.resetPasswordForEmail(email,{redirectTo:"https://monvatticrm.com"});
  return <AuthCtx.Provider value={{session,profile,authLoading,isRecovery,setIsRecovery,signIn,signUp,signOut,updateProfile,updatePassword,resetPassword,fetchProfile}}>{children}</AuthCtx.Provider>;
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
      <div ref={trigRef} onClick={openDrop} style={{cursor:"pointer",overflow:"hidden",maxWidth:"100%"}}>
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

function ConfirmModal({title,message,danger=false,onConfirm,onCancel,requirePhrase=null,boardName=""}) {
  const [typed,setTyped]=useState("");
  const phraseOk=!requirePhrase||typed===requirePhrase;
  const lines=message.split("\n");
  return (
    <Modal title={title} onClose={onCancel} width={480}
      footer={<>
        <button onClick={onCancel} style={{...T.btn,background:"var(--surface3)",color:"var(--text)"}}>Cancelar</button>
        <button onClick={onConfirm} disabled={!phraseOk}
          style={{...T.btn,background:danger?"#dc2626":"var(--blue)",color:"#fff",
            opacity:phraseOk?1:0.4,cursor:phraseOk?"pointer":"not-allowed",transition:"opacity .2s"}}>
          {danger?"Excluir permanentemente":"Confirmar"}
        </button>
      </>}>
      {/* Aviso visual de perigo */}
      {danger&&requirePhrase&&(
        <div style={{background:"#fef2f2",border:"1.5px solid #fca5a5",borderRadius:10,
          padding:"12px 16px",marginBottom:18,display:"flex",gap:12,alignItems:"flex-start"}}>
          <span style={{fontSize:22,flexShrink:0}}>⚠️</span>
          <div>
            <div style={{fontWeight:800,color:"#dc2626",fontSize:14,marginBottom:4}}>Ação irreversível</div>
            <div style={{fontSize:13,color:"#7f1d1d",lineHeight:1.6}}>
              Todos os leads, campos e anotações do quadro <strong>"{boardName}"</strong> serão apagados para sempre. Esta ação não pode ser desfeita.
            </div>
          </div>
        </div>
      )}
      {lines.filter(l=>!l.includes("Para confirmar")&&!l.includes('"Eu realmente')).map((l,i)=>(
        l?<p key={i} style={{margin:"0 0 6px",fontSize:14,color:"var(--text2)",lineHeight:1.7}}>{l}</p>:null
      ))}
      {requirePhrase&&(
        <div style={{marginTop:16}}>
          <div style={{fontSize:13,color:"var(--text2)",marginBottom:8}}>
            Para confirmar, digite exatamente:
          </div>
          <div style={{background:"var(--surface2)",border:"1px solid var(--border)",borderRadius:8,
            padding:"10px 14px",fontFamily:"monospace",fontSize:13,color:"var(--text)",
            marginBottom:12,userSelect:"none",letterSpacing:.2}}>
            {requirePhrase}
          </div>
          <input
            value={typed}
            onChange={e=>setTyped(e.target.value)}
            placeholder="Digite a frase acima..."
            style={{...T.inp,fontFamily:"monospace",fontSize:13,
              borderColor:typed&&!phraseOk?"#dc2626":typed&&phraseOk?"#10b981":"var(--border)",
              boxShadow:typed&&phraseOk?"0 0 0 2px #10b98130":typed&&!phraseOk?"0 0 0 2px #dc262620":"none"}}
            autoFocus
          />
          {typed&&!phraseOk&&(
            <div style={{fontSize:12,color:"#dc2626",marginTop:5}}>A frase não corresponde exatamente.</div>
          )}
          {typed&&phraseOk&&(
            <div style={{fontSize:12,color:"#10b981",marginTop:5,fontWeight:600}}>✓ Confirmado — você pode excluir agora.</div>
          )}
        </div>
      )}
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
// ─── CNPJ ────────────────────────────────────────────────────────────────────
const soDigitos = v => String(v??"").replace(/\D/g,"").slice(0,14);
const formataCNPJ = v => {
  const d = soDigitos(v);
  if(!d) return "";
  return d
    .replace(/^(\d{2})(\d)/,"$1.$2")
    .replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3")
    .replace(/\.(\d{3})(\d)/,".$1/$2")
    .replace(/(\d{4})(\d)/,"$1-$2");
};
// Validação oficial pelos dois dígitos verificadores
const cnpjValido = v => {
  const d = soDigitos(v);
  if(d.length!==14 || /^(\d)\1{13}$/.test(d)) return false;
  const calc = (base,pesos) => {
    const soma = base.reduce((acc,n,i)=>acc+n*pesos[i],0);
    const r = soma % 11;
    return r < 2 ? 0 : 11 - r;
  };
  const n = d.split("").map(Number);
  const dv1 = calc(n.slice(0,12),[5,4,3,2,9,8,7,6,5,4,3,2]);
  const dv2 = calc(n.slice(0,13),[6,5,4,3,2,9,8,7,6,5,4,3,2]);
  return dv1===n[12] && dv2===n[13];
};

function CnpjCell({value,onChange}) {
  // Sem efeito de sincronia: em leitura o valor vem das props; o estado local
  // só existe enquanto o campo está em edição.
  const [ed,setEd]=useState(false);
  const [v,setV]=useState("");
  const ref=useRef();
  useEffect(()=>{if(ed)ref.current?.focus();},[ed]);
  const abrir=()=>{setV(formataCNPJ(value));setEd(true);};
  const digitos=soDigitos(v);
  const invalido=digitos.length>0&&!cnpjValido(digitos);
  const commit=()=>{
    setEd(false);
    const novo=digitos?formataCNPJ(digitos):null;
    if(novo!==(value??null)) onChange(novo);
  };
  const invalidoLeitura=value&&!cnpjValido(value);
  if(!ed) return (
    <div onClick={abrir}
      title={invalidoLeitura?"CNPJ com dígitos verificadores inválidos":undefined}
      style={{padding:"5px 7px",borderRadius:5,cursor:"text",minHeight:22,fontSize:12.5,
        color:value?"var(--text)":"var(--text3)",
        borderBottom:invalidoLeitura?"1.5px solid #d97706":"none"}}>
      {value?formataCNPJ(value):"00.000.000/0000-00"}
    </div>
  );
  return (
    <input ref={ref} value={v}
      onChange={e=>setV(formataCNPJ(e.target.value))}
      onBlur={commit}
      onKeyDown={e=>{if(e.key==="Enter")commit();if(e.key==="Escape")setEd(false);}}
      inputMode="numeric" placeholder="00.000.000/0000-00"
      style={{width:"100%",padding:"5px 7px",borderRadius:5,fontSize:12.5,
        border:`1px solid ${invalido?"#d97706":"var(--blue)"}`,
        background:"var(--surface)",color:"var(--text)",outline:"none"}}/>
  );
}

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
    <div style={{padding:"4px 4px",userSelect:"none",maxWidth:"100%",overflow:"hidden"}}>
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
    <div style={{display:"flex",alignItems:"center",padding:"5px 5px",minHeight:32,maxWidth:"100%",overflow:"hidden"}}>
      {!sel.length
        ? <div style={{width:28,height:28,borderRadius:"50%",border:"1.5px dashed var(--border)",
            display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:"var(--text3)"}}>+</div>
        : <div style={{display:"flex",alignItems:"center",maxWidth:68,overflow:"hidden",flexShrink:0,position:"relative",zIndex:0}}>{sel.slice(0,3).map((u,idx)=><div key={u.id} style={{marginLeft:idx?-7:0,flexShrink:0,zIndex:10-idx,borderRadius:"50%",border:"2px solid var(--surface)",position:"relative"}}><Avatar user={u} size={24}/></div>)}{sel.length>3&&<div style={{marginLeft:-7,width:24,height:24,borderRadius:"50%",background:"var(--surface3)",border:"2px solid var(--surface)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:"var(--text2)",flexShrink:0}}>+{sel.length-3}</div>}</div>
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

function Cell({col,values,allColumns,responsibles,respByCol,allUsers,onChange,onRespChange,item}) {
  const v=values?.[col.id];
  const opts=col.config?.options||[];
  switch(col.tipo){
    case "calculated": return <CalcCell values={values} allColumns={allColumns}/>;
    // Coluna nativa: fonte única em items.numero_cliente, nunca em item_values
    // Renovação: o código é herdado do lead original e NÃO pode ser editado aqui
    // (o banco também trava: gatilho copia o código do original em qualquer insert/update).
    case "numero_cliente":
      if(item?.renovacao_de) return (
        <div title="Renovação: o código é o mesmo do lead original e só pode ser alterado nele"
          style={{display:"flex",alignItems:"center",gap:6,padding:"5px 8px",fontSize:13,fontWeight:600,color:"var(--text2)"}}>
          <span>{item?.numero_cliente??"—"}</span>
          <span style={{fontSize:10,fontWeight:700,color:"#7c3aed",background:"#7c3aed18",borderRadius:10,padding:"1px 7px",letterSpacing:.3}}>RENOVAÇÃO</span>
        </div>
      );
      return <EditableCell value={item?.numero_cliente??""} onChange={onChange} type="number" placeholder="—"/>;
    case "status":     return <StatusCell value={v} options={opts} onChange={onChange}/>;
    case "user":       return <UserCell value={respByCol?.[col.id]||[]} allUsers={allUsers} onChange={ids=>onRespChange(col.id,ids)}/>;
    case "currency":   return <CurrencyCell value={v} onChange={onChange}/>;
    case "link":       return <LinkCell value={v} onChange={onChange}/>;
    case "date":       return <EditableCell value={v} onChange={onChange} type="date"/>;
    case "number":     return <EditableCell value={v} onChange={onChange} type="number" placeholder="0"/>;
    case "phone":      return <EditableCell value={v} onChange={onChange} placeholder="(11) 9…"/>;
    case "cnpj":       return <CnpjCell value={v} onChange={onChange}/>;
    case "email":      return <EditableCell value={v} onChange={onChange} placeholder="email@"/>;
    default:           return <EditableCell value={v} onChange={onChange} placeholder={col.nome}/>;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// RICH EDITOR — contentEditable proper (sem execCommand obsoleto)
// ─────────────────────────────────────────────────────────────────────────────
function RichEditor({onSubmit,disabled=false,initialHTML="",submitLabel="Publicar",onCancel=null}) {
  const editorRef = useRef(null);
  const [isEmpty, setIsEmpty] = useState(!initialHTML);
  // Modo edição: carrega o texto atual uma vez, ao montar
  useEffect(()=>{
    if(initialHTML&&editorRef.current){editorRef.current.innerHTML=initialHTML;editorRef.current.focus();}
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);
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
    if(!onCancel){editorRef.current.innerHTML="";setIsEmpty(true);}
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
        <span style={{fontSize:11,color:"var(--text3)"}}>⌘↵ {onCancel?"salvar":"publicar"} · ⌘B negrito · ⌘I itálico</span>
        {onCancel&&<button onClick={onCancel} disabled={disabled}
          style={{...T.btn,background:"none",color:"var(--text2)",border:"1px solid var(--border)",padding:"8px 16px",fontSize:13}}>
          Cancelar
        </button>}
        <button onClick={submit} disabled={disabled||isEmpty}
          style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 20px",
            opacity:disabled||isEmpty?0.5:1,fontSize:13}}>
          {submitLabel}
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
// ─────────────────────────────────────────────────────────────────────────────
// ENVIO PARA O CRM/TIMER — formulário do cliente fechado
// A chamada sai da Edge Function "crm-timer". O segredo nunca vem para cá.
// Datas trafegam como TEXTO AAAA-MM-DD: sem Date, sem toISOString, sem fuso.
// ─────────────────────────────────────────────────────────────────────────────

// CONFIRMAR O CAMINHO: usado apenas para o atalho até a empresa no CRM/Timer.
const TIMER_EMPRESA_URL = id => `https://www.timermonvatti.com.br/empresas/${id}`;

const SERVICOS_TIMER = [
  {v:"mercado_livre",       l:"Mercado Livre"},
  {v:"shopee",              l:"Shopee"},
  {v:"amazon",              l:"Amazon"},
  {v:"trafego",             l:"Tráfego"},
  {v:"gestao_site",         l:"Gestão de site"},
  {v:"desenvolvimento_site",l:"Desenvolvimento de site"},
];
const CADENCIAS_TIMER = [
  {v:"semanal",            l:"Semanal"},
  {v:"quinzenal",          l:"Quinzenal"},
  {v:"semanal_quinzenal",  l:"Semanal depois quinzenal"},
  {v:"quinzenal_semanal",  l:"Quinzenal depois semanal"},
];

// Soma meses a uma data AAAA-MM-DD usando só aritmética de texto e número.
function somaMesesISO(iso,meses){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(iso||"")) return "";
  const n=parseInt(meses,10);
  if(!Number.isFinite(n)||n<=0) return "";
  const [y,m,d]=iso.split("-").map(Number);
  const total=y*12+(m-1)+n;
  const ny=Math.floor(total/12), nm=(total%12)+1;
  const bissexto=(ny%4===0&&ny%100!==0)||ny%400===0;
  const ultimo=[31,bissexto?29:28,31,30,31,30,31,31,30,31,30,31][nm-1];
  const nd=Math.min(d,ultimo);
  return `${ny}-${String(nm).padStart(2,"0")}-${String(nd).padStart(2,"0")}`;
}

// Campo do formulário — fora do render para não remontar os inputs a cada tecla
function CampoTimer({label,children,obs}){
  return (
    <div style={{marginBottom:14}}>
      <label style={{...T.lbl,marginBottom:6}}>{label}</label>
      {children}
      {obs&&<div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>{obs}</div>}
    </div>
  );
}

// Lê o corpo JSON da resposta mesmo quando o status é de erro. O supabase-js
// trata status fora de 2xx como exceção e guarda a resposta em error.context.
async function corpoDaResposta(data,error){
  if(data&&typeof data==="object") return data;
  const ctx=error?.context;
  try{
    if(ctx&&typeof ctx.json==="function") return await ctx.json();
    if(ctx&&typeof ctx.text==="function") return JSON.parse(await ctx.text());
  }catch{/* corpo ilegível: cai no erro genérico */}
  return null;
}

function TimerIntakeModal({item,board,allUsers=[],onClose,onSent}) {
  const toast=useToast();
  const col=n=>(board.columns||[]).find(c=>c.nome===n);
  const val=n=>{const c=col(n);return c?(item.values?.[c.id]??""):"";};

  const numero=item.numero_cliente??null;
  const razaoInicial=String(val("Razão Social")||"");
  const [razao,setRazao]      =useState(razaoInicial);
  // Coluna de CNPJ do quadro (qualquer tipo); pré-preenche quando já houver valor.
  const colCnpj=(board.columns||[]).find(c=>c.tipo==="cnpj")
             ||(board.columns||[]).find(c=>/cnpj/i.test(c.nome||""));
  const cnpjInicial=formataCNPJ(colCnpj?item.values?.[colCnpj.id]:"");
  const [cnpj,setCnpj]        =useState(cnpjInicial);

  // ── Campos derivados das colunas do quadro (não editáveis aqui; a fonte é a coluna)
  // Valor do Projeto (currency) é guardado como número; Parcelas (number) como texto.
  const colValor=(board.columns||[]).find(c=>c.nome==="Valor do Projeto");
  const colParcelas=(board.columns||[]).find(c=>c.nome==="Parcelas");
  const colCloser=(board.columns||[]).find(c=>c.tipo==="user"&&c.nome==="Closer");
  // Nome fantasia: campo do formulário, obrigatório, pré-preenchido com a coluna
  // "Empresa" do item quando ela tiver valor. Editável aqui, como a razão social.
  const colEmpresa=(board.columns||[]).find(c=>c.nome==="Empresa");
  const fantasiaInicial=(()=>{
    const v=colEmpresa?item.values?.[colEmpresa.id]:null;
    return typeof v==="string"?v.trim():"";
  })();
  const [fantasia,setFantasia]=useState(fantasiaInicial);
  const valorProjeto=(()=>{
    const bruto=colValor?item.values?.[colValor.id]:null;
    if(bruto==null||bruto==="") return null;
    const n=typeof bruto==="number"?bruto:parseFloat(String(bruto).replace(/[^0-9.,-]/g,"").replace(/\./g,"").replace(",","."));
    return Number.isFinite(n)&&n>=0?n:null;
  })();
  const parcelas=(()=>{
    const bruto=colParcelas?item.values?.[colParcelas.id]:null;
    if(bruto==null||bruto==="") return null;
    const n=parseInt(String(bruto).replace(/[^0-9]/g,""),10);
    return Number.isInteger(n)&&n>0?n:null;
  })();
  // Apenas o NOME do closer, como texto. Sem id, sem e-mail, sem foto.
  const closerNome=(()=>{
    const ids=colCloser?(item.respByCol?.[colCloser.id]||[]):[];
    const nomes=ids.map(uid=>allUsers.find(u=>u.id===uid)?.nome).filter(Boolean);
    return nomes.length?nomes.join(", "):null;
  })();
  // Coluna AUSENTE é diferente de coluna vazia: ausente vira aviso visível, para
  // que uma coluna renomeada ou removida não faça o campo sumir em silêncio.
  const derivados=[
    {rotulo:"Valor",    coluna:"Valor do Projeto", ausente:!colValor,
     texto:valorProjeto!=null?fmtBRL(valorProjeto):null},
    {rotulo:"Parcelas", coluna:"Parcelas",         ausente:!colParcelas,
     texto:parcelas!=null?`${parcelas}x`:null},
    {rotulo:"Closer",   coluna:"Closer",           ausente:!colCloser,
     texto:closerNome||null},
  ];
  const [contato,setContato]  =useState(String(val("Nome")||""));
  const [inicio,setInicio]    =useState(String(val("Data Entrada")||""));
  const [meses,setMeses]      =useState(String(val("Tempo de Projeto")||""));
  const [about,setAbout]      =useState(String(val("Obs")||""));
  // Sugestão a partir do Plano: só BPO e Ema coincidem. Os demais planos
  // (Gold, Platinum, Renovação, Feira, Mentoria, Feel) não sugerem nada.
  const planoLower=String(val("Plano")||"").trim().toLowerCase();
  const [modelo,setModelo]    =useState(planoLower==="bpo"?"bpo":planoLower==="ema"?"ema":"");
  const [servicos,setServicos]=useState([]);
  const [sistema,setSistema]  =useState("");
  const [cadencia,setCadencia]=useState("");
  const [dor,setDor]          =useState("");

  const [enviando,setEnviando]=useState(false);
  const [erro,setErro]        =useState(null);
  const [colisao,setColisao]  =useState(null);
  const [parecidas,setParecidas]=useState([]);
  const [conferindo,setConferindo]=useState(true);

  const fim=somaMesesISO(inicio,meses);
  // Todos os campos são obrigatórios para enviar.
  const faltando=[
    !numero&&"Número do cliente (coluna Código)",
    !razao.trim()&&"Razão social",
    !fantasia.trim()&&"Nome fantasia",
    (!soDigitos(cnpj)&&"CNPJ")||(!cnpjValido(cnpj)&&"CNPJ válido (dígitos verificadores)"),
    !contato.trim()&&"Nome do contato",
    !modelo&&"Modelo do projeto",
    !inicio&&"Data de início",
    !String(meses).trim()&&"Tempo de projeto (meses)",
    (inicio&&String(meses).trim()&&!fim)&&"Data de término (confira início e meses)",
    servicos.length===0&&"Serviços contratados (ao menos um)",
    !cadencia&&"Cadência de contato",
    !sistema.trim()&&"Sistema utilizado",
    !dor.trim()&&"Maior dor",
    !about.trim()&&"Sobre",
    // Derivados das colunas do quadro: obrigatórios, mas corrigidos NA COLUNA,
    // não aqui, para não criar um segundo lugar de edição do mesmo dado.
    valorProjeto==null&&"Valor do Projeto (preencha a coluna no quadro)",
    parcelas==null&&"Parcelas (preencha a coluna no quadro)",
    !closerNome&&"Closer (preencha a coluna no quadro)",
  ].filter(Boolean);
  const completo=faltando.length===0;
  // Conferência de duplicado: SOMENTE por CNPJ. A verificação por semelhança de
  // nome foi removida por decisão do Mauricio, junto com a porcentagem.
  //
  // LIMITAÇÃO CONHECIDA: das 159 empresas já cadastradas no CRM/Timer, apenas 2
  // têm CNPJ hoje. Enquanto os CNPJs das antigas não forem preenchidos lá, esta
  // conferência não detecta reenvio de cliente antigo e o envio criará empresa
  // duplicada. Decisão consciente, não é bug.
  //
  // Empresa sem CNPJ cadastrado do outro lado não é erro: só não há como conferir.
  const porCnpj=parecidas;

  // Conferência consultiva de duplicado ao ABRIR
  useEffect(()=>{
    let vivo=true;
    (async()=>{
      try{
        const {data,error}=await db.functions.invoke("crm-timer",{body:{action:"check",payload:{cnpj:soDigitos(cnpjInicial)}}});
        const res=await corpoDaResposta(data,error);
        if(vivo&&res?.ok) setParecidas(res.matches||[]);
      }catch{/* consultivo: falhar aqui não impede o envio */}
      if(vivo) setConferindo(false);
    })();
    return()=>{vivo=false;};
  },[razaoInicial,cnpjInicial]);

  const enviar=async()=>{
    setErro(null);setColisao(null);
    if(!completo){
      setErro("Preencha todos os campos antes de enviar. Faltando: "+faltando.join(", ")+".");
      return;
    }
    setEnviando(true);
    try{
      const payload={
        number:numero,
        razao_social:razao.trim(),
        cnpj:soDigitos(cnpj), // 14 dígitos, sem máscara
        contato:contato.trim()||undefined,
        project_model:modelo||undefined,
        started_on:inicio||undefined,
        ends_on:fim||undefined,
        cadence:cadencia||undefined,
        system_used:sistema.trim()||undefined,
        main_pain:dor.trim()||undefined,
        about:about.trim()||undefined,
        contracted_services:servicos.length?servicos:undefined,
        project_value:valorProjeto??undefined,
        installments:parcelas??undefined,
        closer_name:closerNome??undefined,
        nome_fantasia:fantasia.trim()||undefined,
      };
      const {data,error}=await db.functions.invoke("crm-timer",{body:{action:"create",payload}});
      // Em resposta com status de erro, o cliente Supabase devolve data vazio e
      // guarda o corpo real em error.context. Sem ler dali, a mensagem específica
      // (CNPJ repetido, número em uso, etc.) se perderia num erro genérico.
      const res=await corpoDaResposta(data,error);
      if(!res){
        setErro("Não foi possível falar com a integração. Verifique a conexão e tente de novo.");
        setEnviando(false);return;
      }
      if(!res.ok){
        setErro(res.message||"O envio não foi concluído.");
        if(res.collision) setColisao(res.collision);
        // Leva a colisão para o quadro de atenção, no topo do formulário
        // Só entram no quadro de atenção os casos de empresa duplicada.
        // "número em uso" não é semelhança de empresa, fica só na mensagem.
        if(res.collision&&res.code==="cnpj_in_use"){
          const lista=Array.isArray(res.collision)?res.collision:[res.collision];
          const novos=lista.map(c=>({
            name:c?.name||c?.razao_social||c?.company||"(empresa sem nome)",
            group:c?.group||null,
          }));
          setParecidas(prev=>{
            const chave=x=>String(x.name);
            const vistos=new Set(novos.map(chave));
            return [...novos,...prev.filter(x=>!vistos.has(chave(x)))];
          });
        }
        setEnviando(false);return; // mantém tudo preenchido
      }
      // Só marca como enviado com id confirmado de volta
      // Persiste o CNPJ digitado na coluna do lead, para não se perder
      await onSent(res.id,res.name,{cnpj:formataCNPJ(cnpj),colCnpjId:colCnpj?.id||null});
      toast("Cliente enviado para o CRM/Timer!");
      onClose();
    }catch{
      setErro("Falha inesperada no envio. Nada foi criado.");
      setEnviando(false);
    }
  };

  return (
    <Modal title="🚀 Enviar para o CRM/Timer" onClose={onClose} width={600}>
      {!numero&&(
        <div style={{padding:"11px 13px",background:"rgba(220,38,38,.12)",border:"1px solid #dc2626",
          borderRadius:9,marginBottom:16,fontSize:13,color:"var(--text)"}}>
          Este lead está <strong>sem número de cliente</strong>. Preencha a coluna Código antes de enviar.
        </div>
      )}

      {conferindo&&(
        <div style={{padding:"9px 13px",background:"var(--surface2)",borderRadius:9,marginBottom:14,
          fontSize:12.5,color:"var(--text3)"}}>Conferindo se já existe empresa com este CNPJ…</div>
      )}
      {!conferindo&&porCnpj.length>0&&(
        <div style={{padding:"11px 13px",background:"rgba(220,38,38,.12)",border:"1px solid #dc2626",
          borderRadius:9,marginBottom:12,fontSize:12.5,color:"var(--text)"}}>
          <strong>Já existe empresa com este CNPJ</strong> no CRM/Timer. É o mesmo cliente:
          <ul style={{margin:"7px 0 0",paddingLeft:18}}>
            {porCnpj.slice(0,5).map((m,i)=>(
              <li key={i} style={{marginBottom:2}}>{m.name}{m.group?` · ${m.group}`:""}</li>
            ))}
          </ul>
          <div style={{marginTop:6,color:"var(--text2)"}}>O envio será recusado se insistir neste CNPJ.</div>
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:2}}>
        <CampoTimer label="Número do cliente">
          <div style={{...T.inp,display:"flex",alignItems:"center",fontWeight:700,
            color:numero?"var(--text)":"var(--text3)",background:"var(--surface2)"}}>
            {numero??"— sem número —"}
          </div>
        </CampoTimer>
        <CampoTimer label="Modelo do projeto *">
          <select value={modelo} onChange={e=>setModelo(e.target.value)} style={T.inp}>
            <option value="">— Selecionar —</option>
            <option value="bpo">BPO</option>
            <option value="consultoria">Consultoria</option>
            <option value="ema">Ema</option>
          </select>
        </CampoTimer>
      </div>

      <CampoTimer label="Razão social *">
        <input value={razao} onChange={e=>setRazao(e.target.value)} style={T.inp} maxLength={200}/>
      </CampoTimer>
      <CampoTimer label="Nome fantasia *" obs={!colEmpresa?"Coluna \"Empresa\" não encontrada neste quadro; preencha manualmente.":undefined}>
        <input value={fantasia} onChange={e=>setFantasia(e.target.value)} style={T.inp} maxLength={200}
          placeholder="Nome pelo qual a empresa é conhecida"/>
      </CampoTimer>
      <CampoTimer label="CNPJ *" obs={soDigitos(cnpj)&&!cnpjValido(cnpj)?"Dígitos verificadores não conferem.":undefined}>
        <input value={cnpj} onChange={e=>setCnpj(formataCNPJ(e.target.value))}
          inputMode="numeric" placeholder="00.000.000/0000-00"
          style={{...T.inp,borderColor:soDigitos(cnpj)&&!cnpjValido(cnpj)?"#d97706":undefined}}/>
      </CampoTimer>
      <CampoTimer label="Nome do contato *">
        <input value={contato} onChange={e=>setContato(e.target.value)} style={T.inp}/>
      </CampoTimer>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14}}>
        <CampoTimer label="Início *">
          <input type="date" value={inicio} onChange={e=>setInicio(e.target.value)} style={T.inp}/>
        </CampoTimer>
        <CampoTimer label="Tempo (meses) *">
          <input type="number" min="1" value={meses} onChange={e=>setMeses(e.target.value)} style={T.inp}/>
        </CampoTimer>
        <CampoTimer label="Término" obs="calculado">
          <div style={{...T.inp,display:"flex",alignItems:"center",background:"var(--surface2)",color:fim?"var(--text)":"var(--text3)"}}>
            {fim||"—"}
          </div>
        </CampoTimer>
      </div>

      <CampoTimer label="Do quadro (obrigatórios, editados na coluna) *"
        obs="Vem das colunas do lead. Para preencher ou corrigir, edite a coluna no quadro e reabra este formulário. Valor Mensal não é enviado: o CRM/Timer recalcula.">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
          {derivados.map(d=>(
            <div key={d.rotulo}
              title={d.ausente?`A coluna "${d.coluna}" não existe neste quadro`:(d.texto||"")}
              style={{...T.inp,background:(d.ausente||!d.texto)?"rgba(217,119,6,.10)":"var(--surface2)",
                borderColor:(d.ausente||!d.texto)?"#d97706":undefined,
                display:"flex",alignItems:"center",gap:6,
                color:(d.ausente||!d.texto)?"#d97706":"var(--text)",
                overflow:"hidden",whiteSpace:"nowrap",textOverflow:"ellipsis"}}>
              {d.ausente?"⚠ coluna não encontrada":(d.texto||`⚠ ${d.rotulo}: não preenchido`)}
            </div>
          ))}
        </div>
      </CampoTimer>
      {derivados.some(d=>d.ausente)&&(
        <div style={{padding:"10px 13px",background:"rgba(217,119,6,.12)",border:"1px solid #d97706",
          borderRadius:9,marginBottom:14,fontSize:12.5,color:"var(--text)"}}>
          <strong>Coluna não encontrada neste quadro:</strong>{" "}
          {derivados.filter(d=>d.ausente).map(d=>`"${d.coluna}"`).join(", ")}.
          Esses campos são obrigatórios no envio, então <strong>o envio fica bloqueado</strong>
          até a coluna existir e estar preenchida. Verifique se ela foi renomeada ou removida no quadro.
        </div>
      )}

      <CampoTimer label="Serviços contratados *">
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {SERVICOS_TIMER.map(s=>{
            const on=servicos.includes(s.v);
            return (
              <button key={s.v} type="button"
                onClick={()=>setServicos(p=>on?p.filter(x=>x!==s.v):[...p,s.v])}
                style={{...T.btn,padding:"7px 12px",fontSize:12,
                  background:on?"var(--blue)":"var(--surface2)",
                  color:on?"#fff":"var(--text2)",
                  border:`1px solid ${on?"var(--blue)":"var(--border)"}`}}>
                {on?"✓ ":""}{s.l}
              </button>
            );
          })}
        </div>
      </CampoTimer>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <CampoTimer label="Cadência de contato *">
          <select value={cadencia} onChange={e=>setCadencia(e.target.value)} style={T.inp}>
            <option value="">— Selecionar —</option>
            {CADENCIAS_TIMER.map(c=><option key={c.v} value={c.v}>{c.l}</option>)}
          </select>
        </CampoTimer>
        <CampoTimer label="Sistema utilizado *">
          <input value={sistema} onChange={e=>setSistema(e.target.value)} style={T.inp} placeholder="ERP, plataforma…"/>
        </CampoTimer>
      </div>

      <CampoTimer label="Maior dor *">
        <textarea value={dor} onChange={e=>setDor(e.target.value)} style={{...T.inp,minHeight:64,resize:"vertical"}}/>
      </CampoTimer>
      <CampoTimer label="Sobre *">
        <textarea value={about} onChange={e=>setAbout(e.target.value)} maxLength={5000}
          style={{...T.inp,minHeight:78,resize:"vertical"}}/>
      </CampoTimer>

      {erro&&(
        <div style={{padding:"11px 13px",background:"rgba(220,38,38,.12)",border:"1px solid #dc2626",
          borderRadius:9,marginBottom:14,fontSize:12.5,color:"var(--text)"}}>
          {erro}
          {colisao&&(
            <div style={{marginTop:6,color:"var(--text2)"}}>
              Colidiu com: <strong>{colisao.name||colisao.razao_social||JSON.stringify(colisao)}</strong>
              {colisao.number?` (nº ${colisao.number})`:""}
            </div>
          )}
        </div>
      )}

      {!completo&&(
        <div style={{padding:"10px 13px",background:"var(--surface2)",border:"1px solid var(--border)",
          borderRadius:9,marginBottom:14,fontSize:12.5,color:"var(--text2)"}}>
          <strong style={{color:"var(--text)"}}>Todos os campos são obrigatórios.</strong> Faltando: {faltando.join(", ")}.
        </div>
      )}

      <div style={{display:"flex",justifyContent:"flex-end",gap:10,marginTop:4}}>
        <button onClick={onClose} style={{...T.btn,background:"var(--surface2)",color:"var(--text2)",padding:"10px 16px"}}>
          Cancelar
        </button>
        <button onClick={enviar} disabled={enviando||!completo}
          title={completo?undefined:"Faltando: "+faltando.join(", ")}
          style={{...T.btn,background:"#059669",color:"#fff",padding:"10px 18px",
            opacity:(enviando||!completo)?.55:1}}>
          {enviando?"Enviando…":"Enviar para o CRM/Timer"}
        </button>
      </div>
    </Modal>
  );
}

function ExportModal({board,allUsers=[],onClose}) {
  const [busy,setBusy]=useState(false);
  const [selCols,setSelCols]=useState(()=>new Set(board.columns.map(c=>c.id)));
  const toast=useToast();
  const date=new Date().toLocaleDateString("pt-BR").replace(/\//g,"-");
  const dateStr=new Date().toLocaleString("pt-BR");

  const visibleCols=board.columns.filter(c=>selCols.has(c.id));

  // Achata grupos: resolve estrutura mãe/filho (Negociações) e grupos simples.
  // board.groups já chega com itens pré-filtrados (applyFilters + vendasFilter aplicados no BoardView).
  const flatGroups=(()=>{
    const result=[];
    for(const g of board.groups||[]){
      if(g._isParent){
        for(const sg of g.subGroups||[]){
          if((sg.items||[]).length>0) result.push({...sg,_parentNome:g.nome,_parentColor:g.color});
        }
      } else {
        if((g.items||[]).length>0) result.push(g);
      }
    }
    return result;
  })();

  // Resolve cor de uma opção de status buscando nas options da coluna
  const getStatusColor=(col,label)=>{
    if(!label) return null;
    const opts=resolveOpts(col.config?.options||[]);
    const opt=opts.find(o=>o.label===label);
    return opt?.color||SC[label]||null;
  };

  const getCellVal=(item,col)=>{
    if(col.tipo==="user"){
      const ids=item.respByCol?.[col.id]||[];
      return ids.map(uid=>{const u=allUsers.find(x=>x.id===uid);return u?.nome||"";}).filter(Boolean).join(", ");
    }
    if(col.tipo==="calculated"){
      const cC=board.columns.find(c=>c.tipo==="currency");
      const pC=board.columns.find(c=>c.tipo==="number"&&c.nome.toLowerCase().includes("parcela"));
      if(cC&&pC){const v=parseFloat(item.values?.[cC.id])||0,p=parseFloat(item.values?.[pC.id])||0;return p>0?v/p:0;}
      return 0;
    }
    const v=item.values?.[col.id];
    if(v==null||v==="") return "";
    if(col.tipo==="currency") return parseFloat(v)||0;
    if(col.tipo==="number") return parseFloat(v)||0;
    if(col.tipo==="date") return fmtDate(v)||"";
    return String(v);
  };

  const exportXLSX=()=>{
    setBusy(true);
    try{
      const wb=XLSX.utils.book_new();
      const HEADER_BG="1E293B";
      const HEADER_FG="FFFFFF";
      const TOTAL_BG ="0F4C35";
      const TOTAL_FG ="FFFFFF";
      const GROUP_FG ="FFFFFF";
      const ALT_BG   ="F1F5F9";
      const BORDER   ="CBD5E1";
      const TEXT_MAIN="1E293B";
      const TEXT_MUTED="64748B";
      const fmtCur=v=>v!=null&&v!==""?fmtBRL(parseFloat(v)||0):"";
      const fmtNum=v=>v!=null&&v!==""?String(parseFloat(v)||0):"";
      const hexToARGB=h=>("FF"+(h||"4F46E5").replace("#","").toUpperCase().padStart(6,"0"));
      // Clareia hex para fundo suave de badge (18% cor + 82% branco)
      const hexToLightBg=h=>{
        const r=parseInt((h||"#94a3b8").slice(1,3),16)||148;
        const g=parseInt((h||"#94a3b8").slice(3,5),16)||163;
        const b=parseInt((h||"#94a3b8").slice(5,7),16)||184;
        return [Math.round(r*0.18+255*0.82),Math.round(g*0.18+255*0.82),Math.round(b*0.18+255*0.82)]
          .map(x=>x.toString(16).padStart(2,"0").toUpperCase()).join("");
      };
      const cellStyle=(opts={})=>({
        font:{name:"Calibri",sz:opts.sz||10,...(opts.bold?{bold:true}:{}),...(opts.italic?{italic:true}:{}),...(opts.color?{color:{rgb:opts.color}}:{})},
        fill:opts.fill?{fgColor:{rgb:opts.fill},patternType:"solid"}:{patternType:"none"},
        alignment:{vertical:"center",horizontal:opts.align||"left",wrapText:opts.wrap||false},
        border:{top:{style:"thin",color:{rgb:BORDER}},bottom:{style:"thin",color:{rgb:BORDER}},
                left:{style:"thin",color:{rgb:BORDER}},right:{style:"thin",color:{rgb:BORDER}}},
      });
      const moneyCols=visibleCols.filter(c=>c.tipo==="currency"||c.tipo==="calculated");

      // ── Aba 1: Resumo ──────────────────────────────────────────────────────
      const summaryRows=[];
      summaryRows.push([{v:"MONVATTI CRM — "+board.nome,s:cellStyle({bold:true,sz:15,color:HEADER_FG,fill:HEADER_BG,align:"center"})}]);
      summaryRows.push([{v:"Exportado em: "+dateStr,s:cellStyle({sz:9,color:TEXT_MUTED,align:"center"})}]);
      summaryRows.push([{v:""}]);
      summaryRows.push([
        {v:"Grupo",s:cellStyle({bold:true,sz:11,color:HEADER_FG,fill:HEADER_BG})},
        {v:"Leads",s:cellStyle({bold:true,sz:11,color:HEADER_FG,fill:HEADER_BG,align:"center"})},
        ...moneyCols.map(c=>({v:"Total "+c.nome,s:cellStyle({bold:true,sz:11,color:HEADER_FG,fill:HEADER_BG,align:"right"})}))
      ]);
      const grandTotal=moneyCols.map(()=>0);
      flatGroups.forEach(g=>{
        const itens=g.items||[];
        const sums=moneyCols.map((c,i)=>{const v=itens.reduce((s,it)=>s+(parseFloat(getCellVal(it,c))||0),0);grandTotal[i]+=v;return v;});
        const gc=hexToARGB(g.color||"#4F46E5");
        const parentLabel=g._parentNome?" ("+g._parentNome+")":"";
        summaryRows.push([
          {v:g.nome+parentLabel,s:cellStyle({bold:true,sz:10,color:GROUP_FG,fill:gc.slice(2)})},
          {v:itens.length,s:cellStyle({align:"center",sz:10,fill:gc.slice(2),color:GROUP_FG})},
          ...sums.map(s=>({v:fmtBRL(s),s:cellStyle({align:"right",sz:10,fill:gc.slice(2),color:GROUP_FG})}))
        ]);
      });
      summaryRows.push([
        {v:"TOTAL GERAL",s:cellStyle({bold:true,sz:11,color:TOTAL_FG,fill:TOTAL_BG})},
        {v:flatGroups.reduce((s,g)=>s+(g.items||[]).length,0),s:cellStyle({bold:true,align:"center",sz:11,color:TOTAL_FG,fill:TOTAL_BG})},
        ...grandTotal.map(v=>({v:fmtBRL(v),s:cellStyle({bold:true,align:"right",sz:11,color:TOTAL_FG,fill:TOTAL_BG})}))
      ]);
      const wsSummary=XLSX.utils.aoa_to_sheet(summaryRows.map(r=>r.map(c=>typeof c==="object"&&"v" in c?c.v:c)));
      summaryRows.forEach((row,ri)=>{row.forEach((cell,ci)=>{if(cell?.s){const ref=XLSX.utils.encode_cell({r:ri,c:ci});if(!wsSummary[ref])wsSummary[ref]={};wsSummary[ref].s=cell.s;}});});
      const sumWidth=2+moneyCols.length;
      wsSummary["!merges"]=[{s:{r:0,c:0},e:{r:0,c:sumWidth-1}},{s:{r:1,c:0},e:{r:1,c:sumWidth-1}}];
      wsSummary["!cols"]=[{wch:32},{wch:10},...moneyCols.map(()=>({wch:20}))];
      wsSummary["!rows"]=[{hpt:32},{hpt:18},{hpt:8},{hpt:22},...flatGroups.map(()=>({hpt:22})),{hpt:24}];
      XLSX.utils.book_append_sheet(wb,wsSummary,"Resumo");

      // ── Aba 2: Dados detalhados ─────────────────────────────────────────────
      const detailRows=[];
      detailRows.push([{v:"MONVATTI CRM — "+board.nome,s:cellStyle({bold:true,sz:14,color:HEADER_FG,fill:HEADER_BG,align:"center"})}]);
      detailRows.push([{v:"Exportado em: "+dateStr,s:cellStyle({sz:9,color:TEXT_MUTED,align:"center"})}]);
      detailRows.push([{v:""}]);
      detailRows.push([
        {v:"Grupo",s:cellStyle({bold:true,sz:11,color:HEADER_FG,fill:HEADER_BG})},
        ...visibleCols.map(c=>({v:c.nome,s:cellStyle({bold:true,sz:11,color:HEADER_FG,fill:HEADER_BG,align:c.tipo==="currency"||c.tipo==="number"||c.tipo==="calculated"?"right":"left"})}))
      ]);
      flatGroups.forEach(g=>{
        const gc=hexToARGB(g.color||"#4F46E5");
        const parentLabel=g._parentNome?" — "+g._parentNome:"";
        detailRows.push([
          {v:g.nome+parentLabel+" ("+((g.items||[]).length)+" leads)",s:cellStyle({bold:true,sz:11,color:GROUP_FG,fill:gc.slice(2)})},
          ...visibleCols.map(()=>({v:"",s:cellStyle({fill:gc.slice(2)})}))
        ]);
        const groupSubtotals=visibleCols.map(()=>0);
        (g.items||[]).forEach((item,idx)=>{
          const alt=idx%2===1;
          const rowBg=alt?ALT_BG:undefined;
          const row=[
            {v:g.nome,s:cellStyle({sz:9,color:TEXT_MUTED,fill:rowBg})},
            ...visibleCols.map((c,ci)=>{
              const v=getCellVal(item,c);
              const isNum=c.tipo==="currency"||c.tipo==="calculated"||c.tipo==="number";
              const isMoney=c.tipo==="currency"||c.tipo==="calculated"; if(isMoney&&v) groupSubtotals[ci]=(groupSubtotals[ci]||0)+(parseFloat(v)||0);
              if(c.tipo==="status"&&v){
                const sColor=getStatusColor(c,String(v));
                if(sColor){
                  const sBg=hexToLightBg(sColor);
                  const sFg=sColor.replace("#","").toUpperCase().padStart(6,"0");
                  return {v:String(v),s:cellStyle({align:"center",fill:sBg,color:sFg,bold:true,sz:9})};
                }
              }
              const display=c.tipo==="currency"||c.tipo==="calculated"?fmtCur(v):c.tipo==="number"?fmtNum(v):String(v||"");
              return {v:display,s:cellStyle({align:isNum?"right":"left",fill:rowBg,sz:9,color:TEXT_MAIN})};
            })
          ];
          detailRows.push(row);
        });
        const hasMoney=visibleCols.some(c=>c.tipo==="currency"||c.tipo==="calculated");
        if(hasMoney){
          detailRows.push([
            {v:"Subtotal — "+g.nome,s:cellStyle({bold:true,sz:10,color:TOTAL_FG,fill:TOTAL_BG})},
            ...visibleCols.map((c,ci)=>{
              const isNum=c.tipo==="currency"||c.tipo==="calculated"||c.tipo==="number";
              const isMoney2=c.tipo==="currency"||c.tipo==="calculated"; return {v:isMoney2?fmtBRL(groupSubtotals[ci]||0):"",s:cellStyle({bold:true,sz:10,align:isMoney2?"right":"left",color:TOTAL_FG,fill:TOTAL_BG})};
            })
          ]);
        }
        detailRows.push([{v:""}]);
      });
      const wsDetail=XLSX.utils.aoa_to_sheet(detailRows.map(r=>r.map(c=>typeof c==="object"&&"v" in c?c.v:c)));
      detailRows.forEach((row,ri)=>{row.forEach((cell,ci)=>{if(cell?.s){const ref=XLSX.utils.encode_cell({r:ri,c:ci});if(!wsDetail[ref])wsDetail[ref]={};wsDetail[ref].s=cell.s;}});});
      const numCols=1+visibleCols.length;
      wsDetail["!merges"]=[{s:{r:0,c:0},e:{r:0,c:numCols-1}},{s:{r:1,c:0},e:{r:1,c:numCols-1}}];
      wsDetail["!cols"]=[{wch:24},...visibleCols.map(c=>({wch:c.tipo==="currency"||c.tipo==="calculated"?18:c.tipo==="date"?14:c.tipo==="status"?18:22}))];
      wsDetail["!rows"]=[{hpt:30},{hpt:16},{hpt:8},{hpt:22}];
      XLSX.utils.book_append_sheet(wb,wsDetail,board.nome.substring(0,30));

      XLSX.writeFile(wb,board.nome.replace(/\s+/g,"_")+"_"+date+".xlsx");
      toast("Excel exportado com sucesso!");
    }catch(e){toast("Erro ao exportar Excel","error");console.error(e);}
    setBusy(false);
  };

  return (
    <Modal title={"📤 Exportar — "+board.nome} onClose={onClose} width={520}>
      <div style={{marginBottom:20}}>
        <div style={{fontSize:12,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.4,marginBottom:10}}>
          Colunas a exportar
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {board.columns.map(col=>{
            const on=selCols.has(col.id);
            return (
              <button key={col.id} onClick={()=>setSelCols(prev=>{const n=new Set(prev);on?n.delete(col.id):n.add(col.id);return n;})}
                style={{padding:"5px 12px",borderRadius:20,border:"1.5px solid",fontSize:12,fontWeight:600,
                  cursor:"pointer",transition:"all .15s",
                  background:on?"var(--blue)":"var(--surface2)",
                  color:on?"#fff":"var(--text2)",
                  borderColor:on?"var(--blue)":"var(--border)"}}>
                {col.nome}
              </button>
            );
          })}
        </div>
        <div style={{fontSize:11,color:"var(--text3)",marginTop:8}}>
          {selCols.size} coluna(s) · {flatGroups.reduce((s,g)=>s+(g.items||[]).length,0)} lead(s) filtrado(s)
        </div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        <button onClick={exportXLSX} disabled={busy||selCols.size===0}
          style={{...T.btn,background:"#059669",color:"#fff",padding:"14px 18px",
            justifyContent:"flex-start",gap:14,width:"100%",opacity:selCols.size===0?.5:1}}>
          <div style={{fontSize:26,lineHeight:1,flexShrink:0}}>📊</div>
          <div style={{textAlign:"left"}}>
            <div style={{fontWeight:700,fontSize:14}}>Exportar Excel (.xlsx)</div>
            <div style={{fontSize:11,opacity:.85,fontWeight:400,marginTop:2}}>
              2 abas: Resumo por grupo + dados detalhados com cores de status e subtotais
            </div>
          </div>
        </button>
      </div>
      {busy&&<div style={{display:"flex",alignItems:"center",gap:10,marginTop:16,color:"var(--text3)",fontSize:13}}>
        <Spinner size={20}/> Gerando arquivo…
      </div>}
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
  onDragStart,onDragOver,onDrop,onUpdateValue,onRespChange,sentToNegIds=new Set(),stickyFirstCols=0,colWidths=null,
  siblingStages=null,onMoveStage=null,sentLabel="Negociações",onSendTimer=null,onReativar=null,onRenovar=null}) {
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
  const enviadoTimer=!!item.timer_company_id;
  // Marcador: mesmo padrão visual do pin de processo concluído usado nos outros quadros
  const alreadySent=sentToNegIds.has(item.id)||enviadoTimer;
  const rotuloEnviado=enviadoTimer?"CRM/Timer":sentLabel;
  const rowBg=selected?"var(--row-sel)":hov?"var(--surface2)":"var(--surface)";
  return (
    <tr draggable onDragStart={onDragStart} onDragOver={onDragOver} onDrop={onDrop}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}
      style={{background:rowBg,transition:"background .1s"}}>
      {/* Célula de controles — layout original preservado, badge via borda + dot */}
      <td style={{width:52,textAlign:"center",borderLeft:`3px solid ${alreadySent?"#059669":gc}`,
        verticalAlign:"middle",padding:"0 6px",flexShrink:0,
        position:"sticky",left:0,zIndex:3,background:"var(--surface)",
        boxShadow:"2px 0 0 var(--border)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:5,position:"relative"}}>
          <button onClick={onOpen} title={alreadySent?`Atualizações (já enviado para ${rotuloEnviado})`:"Atualizações"}
            style={{...T.iBtn,opacity:hov||selected?1:0.25,transition:"opacity .15s",fontSize:13,padding:"2px 3px",flexShrink:0,position:"relative"}}>
            📝
            {alreadySent&&<span title={`Lead enviado para ${rotuloEnviado}`} style={{position:"absolute",top:-4,right:-4,width:10,height:10,
              borderRadius:"50%",background:"#00C46A",border:"2px solid var(--surface)",
              display:"block",boxShadow:"0 0 6px #00C46A90"}}/>}
          </button>
          <input type="checkbox" checked={selected} onChange={onToggle}
            style={{cursor:"pointer",accentColor:"var(--blue)",width:13,height:13,flexShrink:0}}/>
        </div>
      </td>
      {columns.map(col=>(
        <td key={col.id} style={{padding:"2px 3px",borderRight:"1px solid var(--border)",verticalAlign:"middle",maxWidth:220,overflow:"hidden",
          ...(stickyFirstCols>0&&columns.indexOf(col)<stickyFirstCols?{position:"sticky",left:(52+columns.slice(0,columns.indexOf(col)).reduce((s,_,i)=>s+(colWidths?.[i]||140),0))+"px",background:rowBg,zIndex:2,
            ...(columns.indexOf(col)===stickyFirstCols-1?{boxShadow:"4px 0 8px -2px var(--shadow)"}:{})}:{})}}>
          <Cell col={col} item={item} values={item.values} allColumns={columns}
            responsibles={item.responsibles} respByCol={item.respByCol} allUsers={allUsers}
            onChange={v=>onUpdateValue(col.id,v)} onRespChange={onRespChange}/>
        </td>
      ))}
      <td style={{width:40,padding:"0 3px",textAlign:"right",verticalAlign:"middle",flexShrink:0,position:"sticky",right:0,background:"var(--surface)",zIndex:2}}>
        <div style={{display:"flex",gap:4,justifyContent:"flex-end",opacity:hov||selected||menu?1:0,transition:"opacity .1s"}}>
          <button ref={btnRef} onClick={openMenu} title="Ações" style={T.iBtn}>⋯</button>
        </div>
        {menu&&createPortal(
          <div ref={menuRef} style={{position:"fixed",top:menuPos.top,right:menuPos.right,zIndex:5000,
            background:"var(--surface)",border:"1px solid var(--border)",borderRadius:12,
            boxShadow:"0 12px 48px var(--shadowMd)",minWidth:240,padding:7}}>
            {onDupNeg&&<ItemMenuOpt icon="📋" label="Duplicar → Negociação" onClick={()=>{setMenu(false);onDupNeg();}}/>}
            {onMoveInativa&&<ItemMenuOpt icon="📁" label="Mover para Base Inativa" onClick={()=>{setMenu(false);onMoveInativa();}}/>}
            {onReativar&&<ItemMenuOpt icon="↩️" label="Reativar → Pré-Vendas" onClick={()=>{setMenu(false);onReativar();}}/>}
            {onSendToNeg&&(
              alreadySent
                ? <ItemMenuOpt icon="🔁" label="Reenviar para Negociações" onClick={()=>{setMenu(false);onSendToNeg();}}/>
                : <ItemMenuOpt icon="🤝" label="Enviar para Negociações" onClick={()=>{setMenu(false);onSendToNeg();}}/>
            )}
            {onRenovar&&<ItemMenuOpt icon="🔄" label="Renovar lead" onClick={()=>{setMenu(false);onRenovar();}}/>}
            {/* Renovação não vai ao CRM/Timer: a empresa já existe lá com este número e CNPJ */}
            {onSendTimer&&!item.renovacao_de&&(
              enviadoTimer
                ? <ItemMenuOpt icon="✅" label="Já enviado ao CRM/Timer" onClick={()=>{setMenu(false);onSendTimer(item);}}/>
                : <ItemMenuOpt icon="🚀" label="Enviar para o CRM/Timer" onClick={()=>{setMenu(false);onSendTimer(item);}}/>
            )}
            {onSendToVendas&&(
              alreadySent
                ? <ItemMenuOpt icon="🔁" label="Reenviar para Vendas" onClick={()=>{setMenu(false);onSendToVendas();}}/>
                : <ItemMenuOpt icon="🏆" label="Enviar para Vendas" onClick={()=>{setMenu(false);onSendToVendas();}}/>
            )}
            {siblingStages&&siblingStages.length>0&&<>
              <div style={{height:1,background:"var(--border)",margin:"5px 0"}}/>
              <div style={{padding:"3px 13px 4px",fontSize:10.5,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.5}}>Mover para</div>
              {siblingStages.map(sg=>(
                <ItemMenuOpt key={sg.id}
                  icon={<span style={{display:"inline-block",width:9,height:9,borderRadius:"50%",background:sg.color||"#888",flexShrink:0}}/>}
                  label={sg.nome}
                  onClick={()=>{setMenu(false);onMoveStage&&onMoveStage(item,sg.id);}}/>
              ))}
            </>}
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
  const firstCol=columns.find(c=>c.tipo==="text"||c.tipo==="email"||c.tipo==="phone")||columns[0];
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
  onOpenItem,onUpdateValue,onRespChange,onDelItem,onMoveInativa,onDupNeg,onSendToNeg,onSendToVendas,onSendTimer,
  onDragStart,onDragOver,onDrop,onItemDragOver,onItemDrop,onGroupSettings,sentToNegIds=new Set(),sentLabel="Negociações",
  sortCfg={colId:null,dir:1},setSortCfg=()=>{},stickyFirstCols=0,siblingStages=null,onMoveStage=null,onReativar=null,onRenovar=null,
  onGroupDragStart=null,onGroupDragEnd=null,canReorderGroups=false}) {
  const [renaming,setRenaming]=useState(false);
  const [gname,setGname]=useState(group.nome);
  const [localCollapsed,setLocalCollapsed]=useState(true); // inicia minimizado
  // Usa collapsed do group prop se passado via sgCollapsed, senão estado local
  // collapsed: undefined/true = fechado, false = aberto
  const isCollapsed=group.collapsed===false?false:localCollapsed;
  const toggleCollapsed=()=>{
    if(onToggle) onToggle(); // sempre propaga ao board state
    else setLocalCollapsed(p=>!p);
  };
  // Larguras de coluna redimensionáveis com persistência em localStorage
  const storageKey="colWidths_"+group.id;
  const defaultWidths=()=>columns.map(c=>
    c.tipo==="date"?105:c.tipo==="currency"||c.tipo==="number"||c.tipo==="calculated"?110:
    c.tipo==="user"?75:c.tipo==="status"?130:140
  );
  const [colWidthsState,setColWidthsState]=useState(()=>{
    try{
      const saved=localStorage.getItem(storageKey);
      if(saved){
        const parsed=JSON.parse(saved);
        if(Array.isArray(parsed)&&parsed.length===columns.length) return parsed;
      }
    }catch(_){}
    return defaultWidths();
  });
  // Persistir ao alterar
  const setColWidths=newWidths=>{
    setColWidthsState(newWidths);
    try{ localStorage.setItem(storageKey,JSON.stringify(newWidths)); }catch(_){}
  };
  const resizingRef=useRef(null); // {colIdx, startX, startW}
  const startResize=(e,colIdx)=>{
    e.preventDefault();
    resizingRef.current={colIdx,startX:e.clientX,startW:colWidthsState[colIdx]};
    const onMove=ev=>{
      const {colIdx:ci,startX,startW}=resizingRef.current;
      const newW=Math.max(50,startW+(ev.clientX-startX));
      setColWidths(colWidthsState.map((w,i)=>i===ci?newW:w));
    };
    const onUp=()=>{
      resizingRef.current=null;
      document.removeEventListener("mousemove",onMove);
      document.removeEventListener("mouseup",onUp);
    };
    document.addEventListener("mousemove",onMove);
    document.addEventListener("mouseup",onUp);
  };
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
  const colWidths=colWidthsState;
  // setColWidths já definido acima com persistência
  const totalW=52+colWidths.reduce((s,w)=>s+w,0)+72;
  const cg=(
    <colgroup>
      <col style={{width:52}}/>
      {columns.map((col,i)=><col key={col.id} style={{width:colWidths[i]}}/>)}
      <col style={{width:40}}/>
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
          borderRadius:isCollapsed?"11px":"11px 11px 0 0",
          boxShadow:"0 2px 10px var(--shadow)"}}>

          {/* Linha do grupo */}
          <div style={{display:"flex",alignItems:"center",gap:9,padding:"10px 14px",
            background:"var(--surface)",borderLeft:`4px solid ${group.color}`,minWidth:totalW}}>
            {canReorderGroups&&onGroupDragStart&&(
              <span draggable title="Arraste para reordenar os grupos"
                onDragStart={e=>onGroupDragStart(e,group.id)} onDragEnd={onGroupDragEnd||undefined}
                style={{cursor:"grab",color:"var(--text3)",fontSize:15,lineHeight:1,userSelect:"none",flexShrink:0}}>⠿</span>
            )}
            <button onClick={toggleCollapsed}
              style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:group.color,
                padding:0,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform .2s",lineHeight:1,flexShrink:0}}>▼</button>
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
          {!isCollapsed&&(
            <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
              {cg}
              <tbody>
                <tr style={{background:"var(--surface2)",borderBottom:"2px solid var(--border)"}}>
                  <th style={{width:52,padding:"7px 6px",textAlign:"center",
                    borderLeft:`4px solid ${group.color}`,
                    position:"sticky",left:0,zIndex:4,background:"var(--surface2)",
                    boxShadow:"2px 0 0 var(--border)"}}>
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
                      <th key={col.id}
                        style={{padding:"7px 11px",textAlign:"left",fontWeight:700,fontSize:11,
                          color:isActive?"var(--blue)":"var(--text3)",whiteSpace:"nowrap",
                          borderRight:"1px solid var(--border)",textTransform:"uppercase",letterSpacing:.4,
                          background:"var(--surface2)",cursor:"pointer",userSelect:"none",
                          transition:"color .15s",position:stickyFirstCols>0&&columns.indexOf(col)<stickyFirstCols?"sticky":"relative",
                          left:stickyFirstCols>0&&columns.indexOf(col)<stickyFirstCols?(52+columns.slice(0,columns.indexOf(col)).reduce((s,c,i)=>s+colWidths[i],0))+"px":undefined,
                          zIndex:stickyFirstCols>0&&columns.indexOf(col)<stickyFirstCols?3:undefined,
                          boxShadow:stickyFirstCols>0&&columns.indexOf(col)===stickyFirstCols-1?"4px 0 8px -2px var(--shadow)":undefined,
                          overflow:"visible"}}>
                        <div onClick={cycleSort} style={{display:"flex",alignItems:"center",gap:4,overflow:"hidden"}}
                          onMouseEnter={e=>{if(!isActive)e.currentTarget.parentElement.style.color="var(--text2)";}}
                          onMouseLeave={e=>{if(!isActive)e.currentTarget.parentElement.style.color="var(--text3)";}}>
                          <span style={{overflow:"hidden",textOverflow:"ellipsis"}}>{col.nome}</span>
                          <div style={{display:"flex",flexDirection:"column",gap:0,flexShrink:0,lineHeight:.8}}>
                            <span style={{fontSize:8,opacity:isAsc?1:.25,color:isActive?"var(--blue)":"var(--text3)"}}>▲</span>
                            <span style={{fontSize:8,opacity:isDesc?1:.25,color:isActive?"var(--blue)":"var(--text3)"}}>▼</span>
                          </div>
                        </div>
                        {/* Alça de resize — drag na borda direita */}
                        <div onMouseDown={e=>{e.stopPropagation();startResize(e,columns.indexOf(col));}}
                          style={{position:"absolute",top:0,right:0,width:6,height:"100%",cursor:"col-resize",
                            zIndex:10,background:"transparent"}}
                          onMouseEnter={e=>e.currentTarget.style.background="var(--blue)30"}
                          onMouseLeave={e=>e.currentTarget.style.background="transparent"}/>
                      </th>
                    );
                  })}
                  <th style={{width:40,background:"var(--surface2)",position:"sticky",right:0,zIndex:4}}/>
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
              <div ref={bodyRef} style={{overflowX:"auto",background:"var(--surface)",position:"relative"}} onScroll={syncHeader}>
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
                        onReativar={onReativar?()=>onReativar(item):null}
                        onRenovar={onRenovar?()=>onRenovar(item):null}
                        onDupNeg={onDupNeg?()=>onDupNeg(item):null}
                        onSendToNeg={onSendToNeg?()=>onSendToNeg(item):null}
                        onSendToVendas={onSendToVendas?()=>onSendToVendas(item):null}
                        onSendTimer={onSendTimer}
                        onDragStart={e=>onDragStart(e,item,group.id)}
                        onDragOver={e=>onItemDragOver(e,item.id,group.id)}
                        onDrop={e=>onItemDrop(e,item.id,group.id)}
                        onUpdateValue={(cid,v)=>onUpdateValue(item.id,cid,v)}
                        onRespChange={(colId,ids)=>onRespChange(item.id,colId,ids)}
                        sentToNegIds={sentToNegIds} sentLabel={sentLabel}
                        stickyFirstCols={stickyFirstCols}
                        colWidths={colWidths}
                        siblingStages={siblingStages}
                        onMoveStage={onMoveStage?(it,tgtGid)=>onMoveStage(it,group.id,tgtGid):null}
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
function ItemPanel({item,board,allUsers,currentUser,onClose,onUpdateValue,onRespChange,onAddUpdate,onDelUpdate,onEditUpdate,canModerar=false,verExcluidas=false}) {
  const [submitting,setSubmitting]=useState(false);
  const [editId,setEditId]=useState(null);      // anotação em edição
  const [histAberto,setHistAberto]=useState({}); // {updateId: true}
  const [hist,setHist]=useState({});            // {updateId: [versões]}
  const [excl,setExcl]=useState(null);          // anotações excluídas deste lead (gestão)
  const [exclAberto,setExclAberto]=useState(false);
  const nomeDe=id=>allUsers.find(x=>x.id===id)?.nome||"Usuário removido";
  const dataHora=d=>d?new Date(d).toLocaleString("pt-BR"):"";
  const carregaHist=async uid=>{
    const {data}=await db.from("item_updates_historico").select("*").eq("update_id",uid).eq("acao","edicao").order("alterado_em",{ascending:false});
    setHist(h=>({...h,[uid]:data||[]}));
  };
  const toggleHist=uid=>{
    const abre=!histAberto[uid];
    setHistAberto(h=>({...h,[uid]:abre}));
    if(abre) carregaHist(uid);
  };
  const salvaEdicao=async(u,html)=>{
    if(html===u.content){setEditId(null);return;}
    setSubmitting(true);
    const ok=await onEditUpdate(u.id,html);
    setSubmitting(false);
    if(ok){setEditId(null);if(histAberto[u.id]) carregaHist(u.id);}
  };
  const itemId=item.id;
  useEffect(()=>{
    if(!verExcluidas) return;
    let vivo=true;
    db.from("item_updates_historico").select("*").eq("item_id",itemId).eq("acao","exclusao").order("alterado_em",{ascending:false})
      .then(({data})=>{if(vivo) setExcl(data||[]);});
    return ()=>{vivo=false;};
  },[itemId,verExcluidas,item.updates?.length]);
  const [activeTab,setActiveTab]=useState("anotacoes"); // "anotacoes" | "campos"
  const {isMobile}=useBreakpoint();
  const firstCol=board?.columns?.find(c=>c.tipo==="text"||c.tipo==="email"||c.tipo==="phone")||board?.columns?.[0];
  const title=item.values?.[firstCol?.id]||"Sem título";
  const handleUpdate=async html=>{setSubmitting(true);await onAddUpdate(html);setSubmitting(false);};
  const panelW=isMobile?"100vw":"520px";
  const updCount=(item.updates||[]).length;

  const tabStyle=(tab)=>({
    padding:"10px 20px",fontSize:13,fontWeight:700,cursor:"pointer",
    border:"none",background:"none",
    color:activeTab===tab?"var(--blue)":"var(--text3)",
    borderBottom:activeTab===tab?"2.5px solid var(--blue)":"2.5px solid transparent",
    transition:"all .15s",whiteSpace:"nowrap",
  });

  return createPortal(
    <div style={{position:"fixed",top:0,right:0,bottom:0,width:panelW,background:"var(--surface)",
      borderLeft:"1.5px solid var(--border)",display:"flex",flexDirection:"column",
      zIndex:1100,boxShadow:"-8px 0 40px var(--shadowMd)"}}>

      {/* Header */}
      <div style={{padding:"17px 22px 14px",borderBottom:"1px solid var(--border)",
        background:"var(--surface2)",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:14}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:16,fontWeight:800,color:"var(--text)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{title}</div>
            <div style={{fontSize:12,color:"var(--text3)",marginTop:2}}>{board?.icon} {board?.nome}</div>
          </div>
          <button onClick={onClose}
            style={{background:"none",border:"1px solid var(--border)",borderRadius:7,fontSize:22,cursor:"pointer",
              color:"var(--text3)",width:35,height:35,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>×</button>
        </div>
        {/* Abas */}
        <div style={{display:"flex",borderBottom:"1px solid var(--border)",marginBottom:-14,marginLeft:-2}}>
          <button style={tabStyle("anotacoes")} onClick={()=>setActiveTab("anotacoes")}>
            💬 Anotações {updCount>0&&<span style={{marginLeft:5,background:"var(--blue)",color:"#fff",borderRadius:20,
              padding:"0 6px",fontSize:11,fontWeight:700}}>{updCount}</span>}
          </button>
          <button style={tabStyle("campos")} onClick={()=>setActiveTab("campos")}>
            📋 Campos
          </button>
        </div>
      </div>

      {/* Conteúdo da aba */}
      <div style={{flex:1,overflowY:"auto",padding:"22px 22px"}}>

        {/* ABA: ANOTAÇÕES */}
        {activeTab==="anotacoes"&&(
          <div>
            <RichEditor onSubmit={handleUpdate} disabled={submitting}/>
            <div style={{marginTop:22}}>
              {updCount===0&&(
                <div style={{textAlign:"center",padding:"32px 0",color:"var(--text3)"}}>
                  <div style={{fontSize:32,marginBottom:10}}>💬</div>
                  <div style={{fontSize:14,fontWeight:600,marginBottom:4}}>Nenhuma anotação ainda</div>
                  <div style={{fontSize:12}}>Use o editor acima para registrar informações sobre este lead.</div>
                </div>
              )}
              {(item.updates||[]).map(u=>{
                const meu=!!u.created_by&&u.created_by===currentUser?.id;
                const versoes=hist[u.id];
                return (
                <div key={u.id} style={{padding:"13px 16px",background:"var(--surface2)",borderRadius:10,
                  marginBottom:10,border:"1px solid var(--border)",borderLeft:"3px solid var(--alt)"}}>
                  <div style={{fontSize:11,color:"var(--text3)",marginBottom:8,display:"flex",alignItems:"center",gap:8,justifyContent:"space-between"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",minWidth:0}}>
                      <Avatar user={allUsers.find(x=>x.id===u.created_by)||currentUser} size={22}/>
                      <span>{dataHora(u.created_at)}</span>
                      {u.edited_at&&(
                        <button onClick={()=>toggleHist(u.id)} title="Ver versões anteriores"
                          style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:11,color:"#b45309",fontWeight:700,textDecoration:"underline"}}>
                          editada em {dataHora(u.edited_at)}
                        </button>
                      )}
                    </div>
                    <div style={{display:"flex",gap:6,flexShrink:0}}>
                      {meu&&editId!==u.id&&<button onClick={()=>setEditId(u.id)} title="Editar anotação" style={{...T.iBtn,fontSize:11}}>✎</button>}
                      {(meu||canModerar)&&<button onClick={()=>onDelUpdate(u.id)} title="Excluir anotação" style={{...T.iBtn,color:"#dc2626",borderColor:"#FEE2E280",fontSize:11}}>✕</button>}
                    </div>
                  </div>
                  {editId===u.id
                    ? <RichEditor initialHTML={u.content} submitLabel="Salvar" disabled={submitting}
                        onSubmit={html=>salvaEdicao(u,html)} onCancel={()=>setEditId(null)}/>
                    : <div style={{fontSize:13,color:"var(--text)",lineHeight:1.75}} dangerouslySetInnerHTML={{__html:linkifyHTML(u.content)}}/>}
                  {histAberto[u.id]&&(
                    <div style={{marginTop:12,paddingTop:10,borderTop:"1px dashed var(--border)"}}>
                      <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",marginBottom:8}}>Versões anteriores (não podem ser apagadas)</div>
                      {!versoes&&<div style={{fontSize:12,color:"var(--text3)"}}>Carregando…</div>}
                      {versoes&&versoes.length===0&&<div style={{fontSize:12,color:"var(--text3)"}}>Nenhuma versão registrada.</div>}
                      {(versoes||[]).map(h=>(
                        <div key={h.id} style={{marginBottom:8,padding:"9px 11px",background:"var(--surface)",border:"1px solid var(--border)",borderRadius:8}}>
                          <div style={{fontSize:11,color:"var(--text3)",marginBottom:5}}>
                            Texto até {dataHora(h.alterado_em)}, alterado por {nomeDe(h.alterado_por)}
                          </div>
                          <div style={{fontSize:12.5,color:"var(--text2)",lineHeight:1.7}} dangerouslySetInnerHTML={{__html:linkifyHTML(h.conteudo_anterior||"")}}/>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );})}
              {verExcluidas&&excl&&excl.length>0&&(
                <div style={{marginTop:18}}>
                  <button onClick={()=>setExclAberto(v=>!v)}
                    style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:12,fontWeight:700,color:"#dc2626"}}>
                    {exclAberto?"▾":"▸"} {excl.length} {excl.length===1?"anotação excluída":"anotações excluídas"}
                  </button>
                  {exclAberto&&excl.map(h=>(
                    <div key={h.id} style={{marginTop:8,padding:"10px 12px",background:"var(--surface2)",border:"1px dashed #dc262660",borderRadius:8}}>
                      <div style={{fontSize:11,color:"var(--text3)",marginBottom:5}}>
                        Escrita por {nomeDe(h.autor_original)} em {dataHora(h.criada_em)}. Excluída por {nomeDe(h.alterado_por)} em {dataHora(h.alterado_em)}.
                      </div>
                      <div style={{fontSize:12.5,color:"var(--text2)",lineHeight:1.7}} dangerouslySetInnerHTML={{__html:linkifyHTML(h.conteudo_anterior||"")}}/>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ABA: CAMPOS */}
        {activeTab==="campos"&&(
          <div>
            {board?.columns?.map(col=>(
              <div key={col.id} style={{display:"flex",alignItems:"flex-start",marginBottom:10,
                gap:12,paddingBottom:10,borderBottom:"1px solid var(--border)"}}>
                <div style={{width:130,fontSize:12,color:"var(--text3)",flexShrink:0,paddingTop:7,fontWeight:600,lineHeight:1.4}}>{col.nome}</div>
                <div style={{flex:1,minWidth:0}}>
                  <Cell col={col} item={item} values={item.values} allColumns={board.columns}
                    responsibles={item.responsibles} respByCol={item.respByCol} allUsers={allUsers}
                    onChange={v=>onUpdateValue(col.id,v)} onRespChange={onRespChange}/>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>,
    document.body
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// FILTER PANEL — status + responsável + intervalo de datas
// ─────────────────────────────────────────────────────────────────────────────
function FilterPanel({board,allUsers,filters,setFilters,onClose}) {
  // Colunas de status lidas diretamente do board — sempre atualizadas
  const statusCols=board?.columns?.filter(c=>c.tipo==="status")||[];
  // Colunas de data do quadro: o filtro usa UMA delas como âncora, escolhida aqui.
  const dateColsBoard=(board?.columns||[]).filter(c=>c.tipo==="date");
  const isVendasBoard=board?.nome==="Vendas";
  const [local,setLocal]=useState(()=>({...filters}));

  // Agrupa opções por coluna para exibição organizada
  const statusByCol=useMemo(()=>
    statusCols.map(col=>({
      col,
      opts:resolveOpts(col.config?.options||[]),
    })).filter(({opts})=>opts.length>0)
  ,[statusCols]);

  const toggleF=(key,val)=>setLocal(f=>({...f,[key]:(f[key]||[]).includes(val)?(f[key]||[]).filter(x=>x!==val):[...(f[key]||[]),val]}));
  const hasAny=Object.values(local).some(v=>Array.isArray(v)?v.length>0:!!v);
  const {isMobile}=useBreakpoint();
  const apply=()=>{setFilters(local);onClose();};
  const clear=()=>setLocal({});

  return createPortal(
    <div style={{position:"fixed",top:0,right:0,bottom:0,width:isMobile?"100vw":"340px",
      background:"var(--surface)",borderLeft:"1.5px solid var(--border)",zIndex:1050,
      boxShadow:"-8px 0 40px var(--shadowMd)",display:"flex",flexDirection:"column",fontFamily:"system-ui"}}>

      {/* Header */}
      <div style={{padding:"17px 22px",borderBottom:"1px solid var(--border)",display:"flex",
        alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <span style={{fontWeight:800,fontSize:16,color:"var(--text)"}}>Filtros</span>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          {hasAny&&<button onClick={clear}
            style={{fontSize:12,color:"#dc2626",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>Limpar tudo</button>}
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:24,color:"var(--text3)",lineHeight:1}}>×</button>
        </div>
      </div>

      <div style={{flex:1,overflowY:"auto",padding:"20px 22px",display:"flex",flexDirection:"column",gap:28}}>

        {/* Filtros de data — ocultos no Vendas pois já existe a barra de período */}
        {!isVendasBoard&&<div>
          <div style={{...T.lbl,marginBottom:12}}>Intervalo de datas</div>
          {dateColsBoard.length>0&&(
            <div style={{marginBottom:12}}>
              <label style={{fontSize:11,color:"var(--text3)",display:"block",marginBottom:4}}>
                Filtrar pela coluna
              </label>
              <select value={local.dateCol||dateColsBoard[0].id}
                onChange={e=>setLocal(f=>({...f,dateCol:e.target.value}))}
                style={{...T.inp,padding:"8px 10px",fontSize:12}}>
                {dateColsBoard.map(c=><option key={c.id} value={c.id}>{c.nome}</option>)}
              </select>
              {dateColsBoard.length>1&&(
                <div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>
                  Este quadro tem {dateColsBoard.length} colunas de data. O filtro usa apenas a escolhida acima.
                </div>
              )}
            </div>
          )}
          <div style={{display:"flex",gap:10,marginBottom:16}}>
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
          <div style={{...T.lbl,marginBottom:8}}>Mês específico</div>
          <input type="month" value={local.month||""} onChange={e=>setLocal(f=>({...f,month:e.target.value||null}))}
            style={{...T.inp}}/>
        </div>}

        {/* Status — agrupado por coluna, sempre refletindo o estado atual do board */}
        {statusByCol.length>0&&<div>
          <div style={{...T.lbl,marginBottom:14}}>Status</div>
          <div style={{display:"flex",flexDirection:"column",gap:18}}>
            {statusByCol.map(({col,opts})=>(
              <div key={col.id}>
                {/* Nome da coluna como separador visual */}
                <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",
                  letterSpacing:"0.07em",marginBottom:8,paddingBottom:5,
                  borderBottom:"1px solid var(--border)"}}>
                  {col.nome}
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                  {opts.map(opt=>{
                    const active=(local.status||[]).includes(opt.label);
                    return (
                      <div key={opt.id||opt.label} onClick={()=>toggleF("status",opt.label)}
                        style={{cursor:"pointer",transition:"all .15s",
                          opacity:active?1:.5,
                          transform:active?"scale(1.06)":"none",
                          outline:active?"2px solid var(--blue)":"2px solid transparent",
                          outlineOffset:2,borderRadius:20}}>
                        <Badge value={opt.label} color={opt.color}/>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>}

        {/* Responsável */}
        {allUsers.length>0&&<div>
          <div style={{...T.lbl,marginBottom:12}}>Responsável</div>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {allUsers.map(u=>{
              const active=(local.resp||[]).includes(u.id);
              return (
                <div key={u.id} onClick={()=>toggleF("resp",u.id)}
                  style={{display:"flex",alignItems:"center",gap:11,padding:"10px 14px",borderRadius:9,cursor:"pointer",
                    border:`1.5px solid ${active?"var(--blue)":"var(--border)"}`,
                    background:active?"var(--surface3)":"var(--surface2)",transition:"all .15s"}}>
                  <Avatar user={u} size={32}/>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:active?700:400,color:"var(--text)"}}>{u.nome}</div>
                    <div style={{fontSize:11,color:"var(--text3)"}}>{u.funcao||""}</div>
                  </div>
                  {active&&<span style={{color:"var(--blue)",fontWeight:700,fontSize:15}}>✓</span>}
                </div>
              );
            })}
          </div>
        </div>}

      </div>

      {/* Footer */}
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
  onUpdateValue,onRespChange,onMoveInativa,onDupNeg,onSendToNeg,onSendToVendas,onSendTimer,sentToNegIds=new Set(),sentLabel="Negociações",
  sortCfg={colId:null,dir:1},setSortCfg=()=>{},
  onDragStart,onDragOver,onDrop,onItemDragOver,onItemDrop,onGroupSettings,
  onRenameSubGroup,onDelSubGroup,onEditParent,onDelParent,onMoveStage,
  onGroupDragStart=null,onGroupDragEnd=null,canReorderGroups=false}) {

  const [collapsed,setCollapsed]=useState(true);  // inicia minimizado
  // Estado local de colapso por sub-grupo (id → bool) — inicia minimizado
  const [sgCollapsed,setSgCollapsed]=useState(()=>{
    const init={};
    (subGroups||[]).forEach(sg=>{init[sg.id]=true;});
    return init;
  });
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
        onClick={()=>!renaming&&setCollapsed(p=>!p)}
        onDragOver={e=>onDragOver&&onDragOver(e,parentGroup.id)} onDrop={e=>onDrop&&onDrop(e,parentGroup.id)}>
        {canReorderGroups&&onGroupDragStart&&(
          <span draggable title="Arraste para reordenar os closers"
            onClick={e=>e.stopPropagation()}
            onDragStart={e=>onGroupDragStart(e,parentGroup.id)} onDragEnd={onGroupDragEnd||undefined}
            style={{cursor:"grab",color:"var(--text3)",fontSize:16,lineHeight:1,userSelect:"none",flexShrink:0}}>⠿</span>
        )}
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
              onRespChange={(iid,colId,ids)=>onRespChange(iid,sg.id,colId,ids)}
              onDelItem={iid=>onDelItem(sg.id,iid)}
              onMoveInativa={onMoveInativa?item=>onMoveInativa(sg.id,item):null}
              onDupNeg={onDupNeg?item=>onDupNeg(sg.id,item):null}
              onSendToNeg={onSendToNeg?item=>onSendToNeg(sg.id,item):null}
              onSendToVendas={onSendToVendas?item=>onSendToVendas(sg.id,item):null}
              onSendTimer={onSendTimer}
              sentToNegIds={sentToNegIds} sentLabel={sentLabel}
              sortCfg={sortCfg} setSortCfg={setSortCfg}
              onDragStart={onDragStart} onDragOver={e=>onDragOver(e,sg.id)} onDrop={e=>onDrop(e,sg.id)}
              onItemDragOver={onItemDragOver} onItemDrop={onItemDrop}
              onGroupSettings={null}
              siblingStages={subGroups.filter(x=>x.id!==sg.id).map(x=>({id:x.id,nome:x.nome,color:x.color,ordem:x.ordem}))}
              onMoveStage={onMoveStage}/>
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
          <strong style={{color:"var(--text)"}}>Ao criar, serão gerados automaticamente 5 sub-grupos:</strong>
          <ul style={{margin:"6px 0 0",paddingLeft:18,lineHeight:1.8}}>
            <li>Oportunidades Closer</li><li>Negociando</li><li>Fechado/Ganho</li>
            <li>Não comparecido</li><li>Recusado/Negócio Futuro</li>
          </ul>
        </div>}

      </div>
    </Modal>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// VENDAS FILTER BAR — barra de filtro de período para o quadro Vendas
// ─────────────────────────────────────────────────────────────────────────────
function VendasFilterBar({vendasFilter,setVendasFilter,visibleGroups}) {
  const now=new Date();
  const mesAtual=now.toISOString().slice(0,7);
  const mesAnterior=new Date(now.getFullYear(),now.getMonth()-1,1).toISOString().slice(0,7);
  const btn=(active)=>({
    padding:"5px 14px",borderRadius:20,border:"1.5px solid",fontSize:12,fontWeight:700,
    cursor:"pointer",transition:"all .15s",
    background:active?"var(--blue)":"var(--surface)",
    color:active?"#fff":"var(--text2)",
    borderColor:active?"var(--blue)":"var(--border)",
  });
  const isMesAtual=vendasFilter.mode==="month"&&vendasFilter.month===mesAtual;
  const isMesAnt=vendasFilter.mode==="month"&&vendasFilter.month===mesAnterior;
  const isRange=vendasFilter.mode==="range";
  const isAll=vendasFilter.mode==="all";
  const totalLeads=visibleGroups.reduce((s,g)=>{
    const items=g._isParent?(g.subGroups||[]).flatMap(sg=>sg.items||[]):g.items||[];
    return s+items.length;
  },0);
  return (
    <div style={{background:"var(--surface2)",borderBottom:"1px solid var(--border)",
      padding:"8px 22px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",flexShrink:0}}>
      <span style={{fontSize:12,fontWeight:700,color:"var(--text3)",marginRight:2}}>📅 Período:</span>
      <button style={btn(isMesAtual)} onClick={()=>setVendasFilter({mode:"month",month:mesAtual,start:"",end:""})}>Mês atual</button>
      <button style={btn(isMesAnt)} onClick={()=>setVendasFilter({mode:"month",month:mesAnterior,start:"",end:""})}>Mês anterior</button>
      <button style={btn(isAll)} onClick={()=>setVendasFilter({mode:"all",month:"",start:"",end:""})}>Todos</button>
      <div style={{display:"flex",alignItems:"center",gap:6,marginLeft:4}}>
        <input type="month" value={vendasFilter.mode==="month"?vendasFilter.month:""} max={mesAtual}
          onChange={e=>setVendasFilter({mode:"month",month:e.target.value,start:"",end:""})}
          style={{border:"1.5px solid var(--border)",borderRadius:8,padding:"4px 8px",fontSize:12,
            background:"var(--surface)",color:"var(--text)",cursor:"pointer"}}/>
      </div>
      <span style={{fontSize:12,color:"var(--text3)",margin:"0 2px"}}>ou</span>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <input type="date" value={vendasFilter.start}
          onChange={e=>setVendasFilter(p=>({...p,mode:"range",start:e.target.value}))}
          style={{border:"1.5px solid "+(isRange?"var(--blue)":"var(--border)"),borderRadius:8,
            padding:"4px 8px",fontSize:12,background:"var(--surface)",color:"var(--text)"}}/>
        <span style={{fontSize:11,color:"var(--text3)"}}>até</span>
        <input type="date" value={vendasFilter.end}
          onChange={e=>setVendasFilter(p=>({...p,mode:"range",end:e.target.value}))}
          style={{border:"1.5px solid "+(isRange?"var(--blue)":"var(--border)"),borderRadius:8,
            padding:"4px 8px",fontSize:12,background:"var(--surface)",color:"var(--text)"}}/>
      </div>
      <div style={{marginLeft:"auto",fontSize:12,color:"var(--text3)",fontWeight:600}}>
        {totalLeads} lead(s) no período
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// BOARD VIEW
// ─────────────────────────────────────────────────────────────────────────────
function BoardView({boardId,boards,allBoardsRaw,allUsers,currentUser,wsId,perms,onBoardCountChange}) {
  const toast=useToast();
  const {isMobile}=useBreakpoint();
  const [board,setBoard]=useState(null);
  const [loading,setLoading]=useState(true);
  const [selItem,setSelItem]=useState(null);
  const [search,setSearch]=useState("");
  const [filters,setFilters]=useState({});
  const [showFilters,setShowFilters]=useState(false);
  const [sortCfg,setSortCfg]=useState({colId:null,dir:1}); // dir: 1=asc, -1=desc
  // Filtro de período para o quadro Vendas
  const todayStr=new Date().toISOString().slice(0,7); // YYYY-MM
  const [vendasFilter,setVendasFilter]=useState({mode:"month",month:todayStr,start:"",end:""});
  const [showColMgr,setShowColMgr]=useState(false);
  const [showExport,setShowExport]=useState(false);
  const [selected,setSelected]=useState(new Set());
  const [stageMenuOpen,setStageMenuOpen]=useState(false);
  const [timerItem,setTimerItem]=useState(null); // item aberto no envio ao CRM/Timer // menu "Mover para etapa" da barra de seleção
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
      db.from("items").select("id,group_id,ordem,created_at,numero_cliente,timer_company_id,renovacao_de").eq("board_id",bid).order("ordem"),
    ]);
    const ids=(itens||[]).map(i=>i.id);
    let vMap={},rMap={},rColMap={};
    if(ids.length){
      // ATENÇÃO: o Supabase corta a resposta em 1000 linhas por requisição e IGNORA
      // .limit() maior que isso. Um quadro com 89 itens e ~11 valores cada já passa
      // de 1000 linhas, e as linhas excedentes (as mais recentes) simplesmente não
      // chegavam: o dado existia no banco, mas sumia da tela ao recarregar.
      //
      // Por isso aqui NÃO basta dividir por item: é preciso PAGINAR com range()
      // dentro de cada bloco, até vir uma página menor que o tamanho da página.
      // Assim funciona independentemente de quantos itens ou colunas o quadro tenha.
      const PAGINA=1000;
      const CHUNK=50;
      const chunks=[];
      for(let i=0;i<ids.length;i+=CHUNK) chunks.push(ids.slice(i,i+CHUNK));
      const buscaPaginado=async(tabela,chunk)=>{
        const out=[];
        for(let de=0;;de+=PAGINA){
          const {data,error}=await db.from(tabela).select("*").in("item_id",chunk).range(de,de+PAGINA-1);
          if(error){console.error("Falha ao carregar "+tabela,error);break;}
          out.push(...(data||[]));
          if(!data||data.length<PAGINA) break;
        }
        return out;
      };
      const [valsArr,respsArr]=await Promise.all([
        Promise.all(chunks.map(chunk=>buscaPaginado("item_values",chunk))),
        Promise.all(chunks.map(chunk=>buscaPaginado("item_responsables",chunk))),
      ]);
      const vals=valsArr.flat();
      const resps=respsArr.flat();
      for(const v of vals){if(!vMap[v.item_id])vMap[v.item_id]={};vMap[v.item_id][v.column_id]=v.value;}
      // Colunas de responsável deste board (ordenadas); a primeira é a "primária" (ex.: Closer)
      const userColIds=(cols||[]).filter(c=>c.tipo==="user").sort((a,b)=>(a.ordem||0)-(b.ordem||0)).map(c=>c.id);
      const validUserCol=new Set(userColIds);
      const primaryUserColId=userColIds[0]||null;
      for(const r of resps){
        if(!rMap[r.item_id])rMap[r.item_id]=[];
        if(!rMap[r.item_id].includes(r.user_id))rMap[r.item_id].push(r.user_id);
        // Coluna efetiva: a do registro, se for uma coluna de responsável válida deste board;
        // senão (legado: column_id nulo/inválido) → coluna primária, para NÃO sumir da tela.
        const eff=(r.column_id&&validUserCol.has(r.column_id))?r.column_id:primaryUserColId;
        if(eff){
          if(!rColMap[r.item_id])rColMap[r.item_id]={};
          if(!rColMap[r.item_id][eff])rColMap[r.item_id][eff]=[];
          if(!rColMap[r.item_id][eff].includes(r.user_id))rColMap[r.item_id][eff].push(r.user_id);
        }
      }
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
      items:(itens||[]).filter(i=>i.group_id===g.id).map(i=>({...i,values:vMap[i.id]||{},responsibles:rMap[i.id]||[],respByCol:rColMap[i.id]||{},updates:[]}))
    }));
    setBoard({...bData,columns,groups});

    // Detecta quais itens já foram encaminhados para o próximo quadro do funil,
    // via origin_item_id — vínculo estrutural permanente, imune a edições de campos.
    // Pré-Vendas → Negociações   |   Negociações → Vendas
    const nextBoardName = bData?.nome==="Pré - Vendas" ? "Negociações"
                        : bData?.nome==="Negociações"  ? "Vendas" : null;
    if(nextBoardName&&ids.length){
      const {data:nb}=await db.from("boards").select("id").eq("nome",nextBoardName).maybeSingle();
      if(nb?.id){
        const {data:linked}=await db.from("items")
          .select("origin_item_id")
          .eq("board_id",nb.id)
          .in("origin_item_id",ids)
          .not("origin_item_id","is",null);
        setSentToNegIds(new Set((linked||[]).map(r=>r.origin_item_id)));
      }else setSentToNegIds(new Set());
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
    // O número é atribuído por gatilho no banco (AFTER INSERT), então não vem no
    // retorno do insert. Relê o item para exibir o código na hora, sem recarregar.
    let numeroCliente=data.numero_cliente??null;
    if(numeroCliente==null){
      const {data:fresh}=await db.from("items").select("numero_cliente").eq("id",data.id).maybeSingle();
      numeroCliente=fresh?.numero_cliente??null;
    }
    // Auto-preenche responsável (owner do grupo) na PRIMEIRA coluna de responsável do board
    const group=board?.groups.find(g=>g.id===gid);
    const ownerIds=group?.owner_id?[group.owner_id]:[];
    const primaryUserCol=(board?.columns||[]).filter(c=>c.tipo==="user").sort((a,b)=>(a.ordem||0)-(b.ordem||0))[0];
    if(ownerIds.length&&primaryUserCol){
      await db.from("item_responsables").insert(ownerIds.map(uid=>({item_id:data.id,user_id:uid,column_id:primaryUserCol.id})));
    }
    upd(b=>{b.groups.find(g=>g.id===gid)?.items.push({...data,numero_cliente:numeroCliente,values:{},responsibles:ownerIds,respByCol:primaryUserCol&&ownerIds.length?{[primaryUserCol.id]:ownerIds}:{},updates:[]});});
    bump(boardId,1);
  };

  // ── Renovação de lead (só no quadro Vendas): cria uma cópia do lead com o MESMO
  // código, e com Razão Social, Nome, Empresa, CNPJ, Closer e SDR já preenchidos.
  // O restante (Plano, Valor, Parcelas, Data Entrada etc.) é preenchido à mão.
  // `renovacao_de` aponta sempre para o lead RAIZ (renovar uma renovação aponta
  // para o original), para que o código se propague de um único lugar.
  const RENOVACAO_CAMPOS=["Razão Social","Nome","Empresa","CNPJ"];
  const RENOVACAO_RESP=["Closer","SDR"];
  const renovarLead=async(gid,item)=>{
    if(!isVendas) return;
    if(item.numero_cliente==null){toast("Este lead ainda não tem código; preencha o código antes de renovar.","error");return;}
    const raizId=item.renovacao_de||item.id;
    const {data,error}=await db.from("items").insert({
      board_id:boardId,group_id:gid,ordem:9999,
      renovacao_de:raizId,numero_cliente:item.numero_cliente,
    }).select().single();
    if(error){toast("Não foi possível criar a renovação: "+(error.message||"erro"),"error");return;}
    const cols=board?.columns||[];
    // Valores de texto/CNPJ
    const vals={};
    const inserts=[];
    cols.filter(c=>RENOVACAO_CAMPOS.includes(c.nome)).forEach(c=>{
      const v=item.values?.[c.id];
      if(v!=null&&String(v).trim()!==""){vals[c.id]=v;inserts.push({item_id:data.id,column_id:c.id,value:v});}
    });
    if(inserts.length){
      const {error:e1}=await db.from("item_values").insert(inserts);
      if(e1) toast("A renovação foi criada, mas alguns campos não foram copiados.","warning");
    }
    // Responsáveis (Closer e SDR), por coluna
    const respByCol={};
    const respIns=[];
    cols.filter(c=>c.tipo==="user"&&RENOVACAO_RESP.includes(c.nome)).forEach(c=>{
      const ids=item.respByCol?.[c.id]||[];
      if(ids.length){respByCol[c.id]=ids;ids.forEach(uid=>respIns.push({item_id:data.id,user_id:uid,column_id:c.id}));}
    });
    if(respIns.length){
      const {error:e2}=await db.from("item_responsables").insert(respIns);
      if(e2) toast("A renovação foi criada, mas Closer/SDR não foram copiados.","warning");
    }
    const responsibles=[...new Set(Object.values(respByCol).flat())];
    upd(b=>{b.groups.find(g=>g.id===gid)?.items.push({...data,values:vals,responsibles,respByCol,updates:[]});});
    bump(boardId,1);
    setSelected(new Set());
    await logAct(currentUser?.id,wsId,"item",data.id,"created",{renovacao_de:raizId,numero_cliente:item.numero_cliente});
    toast(`Renovação criada para o cliente ${item.numero_cliente}. Preencha Plano, Valor, Parcelas e Data Entrada.`,"success");
  };

  const updateValue=async(iid,gid,cid,val)=>{
    const col=(board?.columns||[]).find(c=>c.id===cid);
    // Coluna nativa: fonte única em items.numero_cliente. Nunca grava em item_values.
    if(col?.tipo==="numero_cliente"){
      const limpo=String(val??"").replace(/[^0-9]/g,"");
      const num=limpo===""?null:parseInt(limpo,10);
      // O código do original se propaga às renovações (o banco faz o mesmo por gatilho)
      upd(b=>{b.groups.forEach(g=>g.items.forEach(it=>{ if(it.id===iid||it.renovacao_de===iid) it.numero_cliente=num; }));});
      if(selItem?.id===iid)setSelItem(p=>({...p,numero_cliente:num}));
      const {error}=await db.from("items").update({numero_cliente:num}).eq("id",iid);
      if(error){
        toast(error.code==="23505"
          ? "Esse número já está em uso por outro cliente."
          : "Não foi possível salvar o número do cliente.","error");
        loadBoard(boardId);
      }
      return;
    }
    upd(b=>{const item=b.groups.find(g=>g.id===gid)?.items.find(i=>i.id===iid);if(item)item.values[cid]=val;});
    if(selItem?.id===iid)setSelItem(p=>({...p,values:{...p.values,[cid]:val}}));
    await db.from("item_values").upsert({item_id:iid,column_id:cid,value:val},{onConflict:"item_id,column_id"});
  };

  const updateResp=async(iid,gid,colId,ids)=>{
    upd(b=>{const item=b.groups.find(g=>g.id===gid)?.items.find(i=>i.id===iid);if(item){if(!item.respByCol)item.respByCol={};item.respByCol[colId]=ids;item.responsibles=Object.values(item.respByCol).flat();}});
    if(selItem?.id===iid)setSelItem(p=>{const rc={...(p.respByCol||{}),[colId]:ids};return{...p,respByCol:rc,responsibles:Object.values(rc).flat()};});
    // Escopo por coluna: apaga/insere APENAS os responsáveis desta coluna (column_id)
    await db.from("item_responsables").delete().eq("item_id",iid).eq("column_id",colId);
    // Se esta é a coluna primária (onde o legado sem column_id é exibido), limpa o legado
    // nulo também — assim remoções persistem e não geram duplicação após o refresh.
    const userCols=(board?.columns||[]).filter(c=>c.tipo==="user").sort((a,b)=>(a.ordem||0)-(b.ordem||0));
    if(userCols[0]?.id===colId) await db.from("item_responsables").delete().eq("item_id",iid).is("column_id",null);
    if(ids.length){
      const {error}=await db.from("item_responsables").insert(ids.map(uid=>({item_id:iid,user_id:uid,column_id:colId})));
      if(error){
        // Não falha em silêncio: avisa e recarrega para refletir o estado real do banco
        toast("Não foi possível salvar o responsável: "+(error.message||"erro"),"error");
        loadBoard(boardId);
      }
    }
  };

  const delItem=async(gid,iid)=>{
    // Lead original com renovações não pode ser excluído (o banco bloqueia: as
    // renovações herdam o código dele). Avisa antes de mexer na tela.
    const temRenovacao=(board?.groups||[]).some(g=>(g.items||[]).some(i=>i.renovacao_de===iid));
    if(temRenovacao){toast("Este lead tem renovações ligadas a ele. Exclua as renovações antes.","error");return;}
    const {error}=await db.from("items").delete().eq("id",iid);
    if(error){
      toast(error.code==="23503"?"Este lead tem renovações ligadas a ele. Exclua as renovações antes.":"Não foi possível excluir o item.","error");
      return;
    }
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.items=g.items.filter(i=>i.id!==iid);});
    if(selItem?.id===iid)setSelItem(null);
    setSelected(s=>{const n=new Set(s);n.delete(iid);return n;});
    bump(boardId,-1);
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
  // collapsed: undefined/true = fechado, false = aberto
  const toggleGroup=gid=>upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.collapsed=g.collapsed===false?true:false;});

  // ── CRUD grupos mãe (Negociações hierárquico)
  const SUB_GRUPOS_PADRAO=[
    {nome:"Oportunidades Closer",color:"#3B82F6"},
    {nome:"Negociando",color:"#F59E0B"},
    {nome:"Fechado/Ganho",color:"#10B981"},
    {nome:"Não comparecido",color:"#8B5CF6"},
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
  const editUpdate=async(uid,content)=>{
    const {data,error}=await db.from("item_updates").update({content}).eq("id",uid).select().single();
    if(error||!data){toast(error?.message||"Não foi possível salvar a edição","error");return false;}
    upd(b=>{for(const g of b.groups){const it=g.items.find(i=>(i.updates||[]).some(u=>u.id===uid));if(it){it.updates=it.updates.map(u=>u.id===uid?data:u);break;}}});
    setSelItem(p=>p?({...p,updates:(p.updates||[]).map(u=>u.id===uid?data:u)}):p);
    toast("Anotação editada");
    return true;
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

  // Transporta responsáveis para a coluna "gêmea" (mesmo nome e tipo 'user') do destino.
  // Espelha copyMatchingValues. Linhas legacy (column_id null) são ignoradas no transporte.
  const copyMatchingResponsables=async(srcItemId,srcCols,tgtItemId,tgtCols)=>{
    const {data:resps}=await db.from("item_responsables").select("*").eq("item_id",srcItemId);
    if(!resps?.length) return;
    const tgtByName=Object.fromEntries(tgtCols.filter(c=>c.tipo==="user").map(c=>[c.nome.toLowerCase().trim(),c.id]));
    const toInsert=[];
    for(const r of resps){
      if(!r.column_id) continue;
      const srcCol=srcCols.find(c=>c.id===r.column_id);
      if(!srcCol) continue;
      const tgtId=tgtByName[srcCol.nome.toLowerCase().trim()];
      if(tgtId) toInsert.push({item_id:tgtItemId,user_id:r.user_id,column_id:tgtId});
    }
    if(toInsert.length) await db.from("item_responsables").insert(toInsert);
  };

  const copyUpdates=async(srcItemId,tgtItemId)=>{
    const {data:upds}=await db.from("item_updates").select("*").eq("item_id",srcItemId).order("created_at");
    if(!upds?.length) return;
    await db.from("item_updates").insert(upds.map(u=>({item_id:tgtItemId,content:u.content,created_by:u.created_by})));
  };

  // Núcleo de "mover de verdade" para outro quadro: cria o item no destino com
  // valores, responsáveis e atualizações, e só então exclui o original.
  // Usado pela Base Inativa (ida) e pela reativação (volta ao Pré-Vendas).
  const _moveToBoardCore=async(gid,item,destBoard,destGroup)=>{
    const {data:newItem}=await db.from("items").insert({board_id:destBoard.id,group_id:destGroup.id,ordem:9999}).select().single();
    if(!newItem){toast("Erro ao mover","error");return false;}
    const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",destBoard.id);
    await copyMatchingValues(item.id,board.columns,newItem.id,tgtCols||[]);
    await copyMatchingResponsables(item.id,board.columns,newItem.id,tgtCols||[]);
    await copyUpdates(item.id,newItem.id);
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.items=g.items.filter(i=>i.id!==item.id);});
    if(selItem?.id===item.id)setSelItem(null);
    bump(boardId,-1);bump(destBoard.id,1);
    await db.from("items").delete().eq("id",item.id);
    await logAct(currentUser?.id,wsId,"item",item.id,"moved",{to:destBoard.nome,group:destGroup.nome});
    return true;
  };
  // ── Duplicados na Base Inativa: a chave é o TELEFONE (coluna "Número").
  // A Base Inativa não tem CNPJ preenchido e o nome da empresa é texto livre,
  // então o telefone é a única chave confiável. Normaliza para dígitos, ignora
  // o +55 e compara os 8 últimos dígitos; se os dois lados têm DDD, o DDD também
  // precisa bater (o 9º dígito do celular pode faltar em registros antigos).
  const chaveFone=raw=>{
    let d=String(raw??"").replace(/\D/g,"");
    if(d.startsWith("55")&&d.length>=12) d=d.slice(2);
    if(d.length<8) return null;
    // 11 dígitos = DDD + 9 + 8; 10 dígitos = DDD + 8 (sem o nono dígito)
    const ddd=d.length>=11?d.slice(-11,-9):d.length===10?d.slice(-10,-8):null;
    return {d8:d.slice(-8),ddd};
  };
  const mesmoFone=(a,b)=>!!a&&!!b&&a.d8===b.d8&&(!a.ddd||!b.ddd||a.ddd===b.ddd);
  const colFoneDe=cols=>(cols||[]).find(c=>c.tipo==="phone")||(cols||[]).find(c=>c.nome.toLowerCase().trim()==="número");
  // Procura na Base Inativa um lead com o mesmo telefone do item que está chegando.
  const buscaDuplicadoInativa=async(item,iBoard,tgtCols)=>{
    const cSrc=colFoneDe(board?.columns), cDst=colFoneDe(tgtCols);
    if(!cSrc||!cDst) return null;
    const chave=chaveFone(item.values?.[cSrc.id]);
    if(!chave) return null;
    const PAG=1000, linhas=[];
    for(let de=0;;de+=PAG){
      const {data,error}=await db.from("item_values").select("item_id,value").eq("column_id",cDst.id).range(de,de+PAG-1);
      if(error){console.error("Base Inativa: falha ao conferir duplicado",error);return null;}
      linhas.push(...(data||[]));
      if(!data||data.length<PAG) break;
    }
    const achado=linhas.find(r=>{
      const v=(r.value&&typeof r.value==="object"&&"value" in r.value)?r.value.value:r.value;
      return mesmoFone(chave,chaveFone(v));
    });
    if(!achado) return null;
    const {data:ex}=await db.from("items").select("id,board_id,group_id").eq("id",achado.item_id).eq("board_id",iBoard.id).maybeSingle();
    return ex||null;
  };
  // Funde o item que chega no lead que JÁ existe na Base Inativa: preenche só os
  // campos vazios do existente, leva responsáveis e atualizações (com a data
  // original), registra uma atualização automática e exclui a origem.
  const _fundeNaInativa=async(gid,item,existente,iBoard,tgtCols)=>{
    const {data:valsSrc}=await db.from("item_values").select("*").eq("item_id",item.id);
    const {data:valsDst}=await db.from("item_values").select("column_id,value").eq("item_id",existente.id);
    const vazio=v=>v==null||v===""||(typeof v==="object"&&"value" in v&&(v.value==null||v.value===""));
    const temNoDestino=new Set((valsDst||[]).filter(v=>!vazio(v.value)).map(v=>v.column_id));
    const tgtByName=Object.fromEntries(tgtCols.map(c=>[c.nome.toLowerCase().trim(),c.id]));
    const inserts=[];
    for(const v of (valsSrc||[])){
      if(vazio(v.value)) continue;
      const srcCol=(board?.columns||[]).find(c=>c.id===v.column_id);
      const tgtId=srcCol?tgtByName[srcCol.nome.toLowerCase().trim()]:null;
      if(tgtId&&!temNoDestino.has(tgtId)) inserts.push({item_id:existente.id,column_id:tgtId,value:v.value});
    }
    if(inserts.length) await db.from("item_values").upsert(inserts,{onConflict:"item_id,column_id"});
    // Responsáveis: só os que ainda não estão na linha existente (evita violar a chave única)
    const {data:respsSrc}=await db.from("item_responsables").select("*").eq("item_id",item.id);
    const {data:respsDst}=await db.from("item_responsables").select("user_id,column_id").eq("item_id",existente.id);
    const tgtUserByName=Object.fromEntries(tgtCols.filter(c=>c.tipo==="user").map(c=>[c.nome.toLowerCase().trim(),c.id]));
    const jaTem=new Set((respsDst||[]).map(r=>r.user_id+"|"+r.column_id));
    const rIns=[];
    for(const r of (respsSrc||[])){
      if(!r.column_id) continue;
      const srcCol=(board?.columns||[]).find(c=>c.id===r.column_id);
      const tgtId=srcCol?tgtUserByName[srcCol.nome.toLowerCase().trim()]:null;
      if(tgtId&&!jaTem.has(r.user_id+"|"+tgtId)){jaTem.add(r.user_id+"|"+tgtId);rIns.push({item_id:existente.id,user_id:r.user_id,column_id:tgtId});}
    }
    if(rIns.length) await db.from("item_responsables").insert(rIns);
    // Atualizações, preservando a data original
    const {data:upds}=await db.from("item_updates").select("*").eq("item_id",item.id).order("created_at");
    if(upds?.length) await db.from("item_updates").insert(upds.map(u=>({item_id:existente.id,content:u.content,created_by:u.created_by,created_at:u.created_at})));
    const hoje=new Date().toLocaleDateString("pt-BR");
    await db.from("item_updates").insert({item_id:existente.id,created_by:currentUser?.id||null,
      content:`🔁 Lead chegou de novo à Base Inativa em ${hoje}, vindo de ${board?.nome||"outro quadro"} (${currentUser?.nome||"usuário"}). Já existia aqui pelo mesmo telefone; histórico unido nesta linha.`});
    // Remove a origem
    upd(b=>{const g=b.groups.find(g=>g.id===gid);if(g)g.items=g.items.filter(i=>i.id!==item.id);});
    if(selItem?.id===item.id)setSelItem(null);
    bump(boardId,-1);
    await db.from("items").delete().eq("id",item.id);
    await logAct(currentUser?.id,wsId,"item",item.id,"moved",{to:iBoard.nome,merged_into:existente.id});
    return true;
  };

  // Núcleo do "mover para Base Inativa" — sem toast, reutilizado pelo item único e pelo lote.
  // Antes de criar linha nova, confere duplicado por telefone; se já existe, FUNDE.
  const _moveToInativaCore=async(gid,item)=>{
    const iBoard=boards.find(b=>b.nome==="Base Inativa");
    if(!iBoard){toast("Board 'Base Inativa' não encontrado","error");return false;}
    const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",iBoard.id);
    const existente=await buscaDuplicadoInativa(item,iBoard,tgtCols||[]);
    if(existente){
      const ok=await _fundeNaInativa(gid,item,existente,iBoard,tgtCols||[]);
      return ok?"merged":false; // quem chama decide a mensagem
    }
    const {data:tgs}=await db.from("groups").select("*").eq("board_id",iBoard.id).order("ordem");
    const isTraf=board?.nome==="Tráfego";
    const tg=isTraf?tgs?.find(g=>g.nome==="Tráfego"):tgs?.find(g=>g.nome==="Prospecção Fria")||tgs?.[0];
    if(!tg){toast("Grupo destino não encontrado","error");return false;}
    return _moveToBoardCore(gid,item,iBoard,tg);
  };

  // ── Reativar lead: Base Inativa → Pré-Vendas (move real, com escolha de grupo).
  // SDR só pode escolher entre os grupos dos quais é dono no Pré-Vendas; os demais
  // cargos escolhem qualquer grupo. Único destino permitido: Pré-Vendas.
  const abrirReativar=async(alvos)=>{ // alvos = [{gid,item}]
    if(!alvos.length) return;
    const pBoard=(allBoardsRaw||boards).find(b=>b.nome==="Pré - Vendas");
    if(!pBoard){toast("Quadro 'Pré - Vendas' não encontrado","error");return;}
    const {data:pgs}=await db.from("groups").select("*").eq("board_id",pBoard.id).order("ordem");
    let destinos=pgs||[];
    if(perms.slug==="sdr") destinos=destinos.filter(g=>g.owner_id===currentUser?.id);
    if(!destinos.length){
      toast(perms.slug==="sdr"?"Você ainda não tem um grupo próprio no Pré-Vendas. Peça ao gestor para criar um.":"Nenhum grupo encontrado no Pré-Vendas","error");
      return;
    }
    setGroupSelM({
      title:`Reativar ${alvos.length} lead(s) → Pré-Vendas: escolha o grupo`,
      groups:destinos,
      onSelect:async(g)=>{
        setGroupSelM(null);
        setSelected(new Set());
        let ok=0;
        for(const a of alvos){ if(await _moveToBoardCore(a.gid,a.item,pBoard,g)) ok++; }
        toast(`${ok} lead(s) reativado(s) em "${g.nome}"`+(ok<alvos.length?` · ${alvos.length-ok} com erro`:""), ok?"success":"error");
      },
    });
  };
  const reativarLead=(gid,item)=>abrirReativar([{gid,item}]);
  const bulkReativar=()=>{
    const ids=[...selected];
    const alvos=[];
    board.groups.forEach(g=>(g.items||[]).forEach(it=>{ if(ids.includes(it.id)) alvos.push({gid:g.id,item:it}); }));
    abrirReativar(alvos);
  };
  const moveInativa=async(gid,item)=>{
    const r=await _moveToInativaCore(gid,item);
    if(r==="merged") toast("Este lead já estava na Base Inativa (mesmo telefone). O histórico foi unido na linha existente.","warning");
    else if(r) toast("Movido para Base Inativa (com dados e atualizações)");
  };
  // Move em lote os leads selecionados para a Base Inativa (usado pela barra do checkbox)
  const bulkMoveInativa=async()=>{
    const ids=[...selected];
    const targets=[];
    board.groups.forEach(g=>(g.items||[]).forEach(it=>{ if(ids.includes(it.id)) targets.push({gid:g.id,item:it}); }));
    if(!targets.length) return;
    setSelected(new Set());
    let ok=0, fundidos=0;
    for(const t of targets){ const r=await _moveToInativaCore(t.gid,t.item); if(r){ok++; if(r==="merged") fundidos++;} }
    toast(`${ok} lead(s) movido(s) para Base Inativa`+(fundidos?` · ${fundidos} já existia(m) lá e foi(ram) unido(s) pelo telefone`:"")+(ok<targets.length?` · ${targets.length-ok} com erro`:""), ok?"success":"error");
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
        await copyMatchingResponsables(item.id,board.columns,ni.id,tgtCols||[]);
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
    // Busca direto no DB para não depender do boards state filtrado por permissão
    const {data:negBoardRow}=await db.from("boards").select("id,nome").eq("nome","Negociações").single();
    const negBoard=negBoardRow;
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
        await copyMatchingResponsables(item.id,board.columns,ni.id,tgtCols||[]);
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
        const {data:ni}=await db.from("items").insert({board_id:vendasBoard.id,group_id:tg.id,ordem:9999,origin_item_id:item.id}).select().single();
        if(!ni){toast("Erro ao criar item em Vendas","error");setGroupSelM(null);return;}
        const {data:tgtCols}=await db.from("columns").select("*").eq("board_id",vendasBoard.id);
        await copyMatchingValues(item.id,board.columns,ni.id,tgtCols||[]);
        await copyMatchingResponsables(item.id,board.columns,ni.id,tgtCols||[]);
        await copyUpdates(item.id,ni.id);
        bump(vendasBoard.id,1);
        setSentToNegIds(prev=>new Set([...prev,item.id]));
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
  // Move um único item para outro grupo — fonte única usada pelo arrastar E pelo menu "..."
  const moveItemToGroup=async(item,srcGroupId,tgtGroupId)=>{
    if(!item||!tgtGroupId||srcGroupId===tgtGroupId) return;
    upd(b=>{
      const src=b.groups.find(g=>g.id===srcGroupId);
      const tgt=b.groups.find(g=>g.id===tgtGroupId);
      if(src&&tgt){src.items=src.items.filter(i=>i.id!==item.id);tgt.items.push({...item,group_id:tgtGroupId});}
    });
    await db.from("items").update({group_id:tgtGroupId,ordem:9999}).eq("id",item.id);
  };
  // ── Reordenação de grupos por arrastar (alça ⠿ no cabeçalho do grupo).
  // Só entre irmãos do mesmo nível: grupos soltos entre si (Pré-Vendas, Vendas,
  // Base Inativa) ou grupos mãe entre si (Negociações). Sub-grupos (etapas) NÃO
  // são reordenáveis: a "ordem" deles identifica a etapa entre todos os closers
  // (ver stageOptions), e mexer nela quebraria o "Mover para etapa".
  const groupDragRef=useRef(null);
  const canReorderGroups=!!(perms.all||perms.manageBoards||perms.createGroups);
  const nivelGrupo=g=>g.is_parent?"mae":g.parent_group_id?"sub":"solto";
  const handleGroupHandleDragStart=(e,gid)=>{
    e.stopPropagation();
    groupDragRef.current=gid;
    e.dataTransfer.effectAllowed="move";
    try{e.dataTransfer.setData("text/plain",gid);}catch{/* alguns navegadores exigem setData para iniciar o arrasto */}
  };
  const handleGroupHandleDragEnd=()=>{groupDragRef.current=null;setDragOverGroup(null);};
  const reorderGroups=async(srcGid,tgtGid)=>{
    if(!board||srcGid===tgtGid) return;
    const src=board.groups.find(g=>g.id===srcGid), tgt=board.groups.find(g=>g.id===tgtGid);
    if(!src||!tgt) return;
    const nivel=nivelGrupo(src);
    if(nivel==="sub"||nivel!==nivelGrupo(tgt)){toast("Só é possível reordenar grupos do mesmo nível","warning");return;}
    const irmaos=board.groups.filter(g=>nivelGrupo(g)===nivel).sort((a,b)=>a.ordem-b.ordem);
    const from=irmaos.findIndex(g=>g.id===srcGid), to=irmaos.findIndex(g=>g.id===tgtGid);
    if(from<0||to<0) return;
    const nova=[...irmaos]; nova.splice(to,0,nova.splice(from,1)[0]);
    const ordemPor=Object.fromEntries(nova.map((g,idx)=>[g.id,idx]));
    // Atualiza local: aplica a nova ordem e reposiciona os irmãos nos lugares que já ocupavam
    upd(b=>{
      const slots=[]; b.groups.forEach((g,i)=>{ if(ordemPor[g.id]!=null) slots.push(i); });
      const reordenados=b.groups.filter(g=>ordemPor[g.id]!=null).sort((a,c)=>ordemPor[a.id]-ordemPor[c.id]);
      reordenados.forEach(g=>{g.ordem=ordemPor[g.id];});
      slots.forEach((pos,k)=>{ b.groups[pos]=reordenados[k]; });
    });
    const results=await Promise.all(nova.map((g,idx)=>db.from("groups").update({ordem:idx}).eq("id",g.id)));
    if(results.some(r=>r.error)) toast("A nova ordem não foi salva por completo. Recarregue o quadro.","error");
  };
  const handleGroupDrop=async(e,tgtGid)=>{
    e.preventDefault();
    // Arrasto de GRUPO (alça do cabeçalho) tem prioridade sobre arrasto de item
    if(groupDragRef.current){
      const srcGid=groupDragRef.current;
      groupDragRef.current=null;
      setDragState(null);setDragOverGroup(null);
      await reorderGroups(srcGid,tgtGid);
      return;
    }
    if(!dragState||dragState.srcGroupId===tgtGid){setDragState(null);setDragOverGroup(null);return;}
    // Item nunca cai dentro de um grupo mãe (só em sub-grupos ou grupos soltos)
    if(board?.groups.find(g=>g.id===tgtGid)?.is_parent){setDragState(null);setDragOverGroup(null);return;}
    const {srcGroupId,item}=dragState;
    setDragState(null);setDragOverGroup(null);
    await moveItemToGroup(item,srcGroupId,tgtGid);
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

  // ── Etapas (grupos filhos padronizados) para o "Mover para etapa" da barra de seleção
  const stageOptions=(()=>{
    const seen={};
    (board?.groups||[]).filter(g=>g.parent_group_id).forEach(g=>{
      if(seen[g.ordem]==null) seen[g.ordem]={ordem:g.ordem,nome:g.nome,color:g.color};
    });
    return Object.values(seen).sort((a,b)=>a.ordem-b.ordem);
  })();
  // Move os leads selecionados para a etapa (ordem) do PRÓPRIO grupo-mãe de cada um.
  // Leads em grupos sem grupo-mãe (legacy) são ignorados.
  const bulkMoveToStage=async(ordem)=>{
    const ids=[...selected];
    const plan=[]; let skipped=0;
    ids.forEach(id=>{
      const src=board.groups.find(g=>(g.items||[]).some(i=>i.id===id));
      if(!src||!src.parent_group_id){skipped++;return;}
      const tgt=board.groups.find(g=>g.parent_group_id===src.parent_group_id&&g.ordem===ordem);
      if(!tgt){skipped++;return;}
      if(tgt.id===src.id) return; // já está nessa etapa
      plan.push({id,srcId:src.id,tgtId:tgt.id});
    });
    if(plan.length){
      upd(b=>{
        plan.forEach(p=>{
          const src=b.groups.find(g=>g.id===p.srcId);
          const tgt=b.groups.find(g=>g.id===p.tgtId);
          if(src&&tgt){const it=src.items.find(i=>i.id===p.id);if(it){src.items=src.items.filter(i=>i.id!==p.id);tgt.items.push({...it,group_id:p.tgtId});}}
        });
      });
      await Promise.all(plan.map(p=>db.from("items").update({group_id:p.tgtId,ordem:9999}).eq("id",p.id)));
    }
    toast(`${plan.length} lead(s) movido(s)`+(skipped?` · ${skipped} fora de grupo-mãe ignorado(s)`:""),plan.length?"success":"warning");
    setSelected(new Set());
    setStageMenuOpen(false);
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
      // ÂNCORA ÚNICA: o filtro usa UMA coluna de data, escolhida pelo usuário.
      // Antes percorria todas as colunas de data exigindo que TODAS casassem, o que
      // fazia o filtro de um quadro com "Data Reunião" e "Follow Up" responder pela
      // coluna errada. Sem escolha explícita, usa a primeira coluna de data do quadro.
      const dateCols=(board?.columns||[]).filter(c=>c.tipo==="date");
      const ancora=dateCols.find(c=>c.id===filters.dateCol)||dateCols[0];
      if(ancora){
        r=r.filter(i=>{
          // Datas são texto AAAA-MM-DD: compara como texto, sem Date e sem fuso.
          const bruto=i.values?.[ancora.id];
          const v=typeof bruto==="string"?bruto.slice(0,10):"";
          if(!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false; // sem data na âncora sai do filtro
          if(filters.dateFrom&&v<filters.dateFrom) return false;
          if(filters.dateTo&&v>filters.dateTo) return false;
          if(filters.month&&v.slice(0,7)!==filters.month) return false;
          return true;
        });
      }
    }
    return r;
  };

  const toggleItem=id=>setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const selectAll=(items,on)=>setSelected(s=>{const n=new Set(s);items.forEach(i=>on?n.add(i.id):n.delete(i.id));return n;});
  // Botões contextuais por quadro
  const isPreVendas      = board?.nome==="Pré - Vendas";     // botão → Negociações (todos)
  const isNegociacoes    = board?.nome==="Negociações";       // botão → Vendas (sendToVendas)
  // Rótulo do pin verde: indica para qual quadro o lead já foi encaminhado
  const sentLabel        = isNegociacoes ? "Vendas" : "Negociações";
  // Envio ao CRM/Timer: só no quadro Vendas e para quem tem a permissão.
  const abrirEnvioTimer=(item)=>{
    if(item?.timer_company_id){
      toast("Este cliente já foi enviado ao CRM/Timer (empresa "+item.timer_company_id+").","warning");
      return;
    }
    if(!item?.numero_cliente){
      toast("Preencha o Código do cliente antes de enviar ao CRM/Timer.","error");
      return;
    }
    setTimerItem(item);
  };
  // Grava o identificador devolvido. Só é chamado com id confirmado.
  const marcarEnviadoTimer=async(companyId,_nome,extras)=>{
    const iid=timerItem?.id;
    if(!iid||!companyId) return;
    const cid=String(companyId);
    upd(b=>{b.groups.forEach(g=>{const it=(g.items||[]).find(i=>i.id===iid);if(it)it.timer_company_id=cid;});});
    await db.from("items").update({timer_company_id:cid}).eq("id",iid);
    // Guarda o CNPJ informado no formulário na coluna de CNPJ do quadro
    const colId=extras?.colCnpjId, valorCnpj=extras?.cnpj;
    if(colId&&valorCnpj){
      upd(b=>{b.groups.forEach(g=>{const it=(g.items||[]).find(i=>i.id===iid);if(it)it.values[colId]=valorCnpj;});});
      await db.from("item_values").upsert({item_id:iid,column_id:colId,value:valorCnpj},{onConflict:"item_id,column_id"});
    }
  };
  const isVendas         = board?.nome==="Vendas";            // filtro de período
  const isInativa        = board?.nome==="Base Inativa";      // SDR vê todos os grupos aqui
  const canRenovar       = isVendas&&!!(perms.all||perms.editAny); // botão "Renovar lead" (só Vendas)
  const canActions       = isPreVendas;                       // mover inativa habilitado no Pré-Vendas

  // ── Hierarquia de grupos
  const canManageParent=perms.all||perms.manageBoards||["administrador","ceo","gerente_comercial"].includes(perms.slug);

  // Filtra grupos conforme permissão do usuário
  const visibleGroups=useMemo(()=>{
    if(!board) return [];

    if(!isNegociacoes){
      let gs=board.groups;
      // SDR vê só os grupos dos quais é dono, EXCETO na Base Inativa: lá vê tudo
      // (o grupo único da Base Inativa não tem dono, e ficava invisível ao SDR).
      if(perms.slug==="sdr"&&!isInativa) gs=gs.filter(g=>g.owner_id===currentUser?.id);
      else if(perms.slug==="closer") gs=gs.filter(g=>(groupAccess[g.id]||[]).includes(currentUser?.id));
      // Para Vendas: aplica filtro de período adicional.
      // Âncora = valor da coluna "Data Entrada" (e NÃO o created_at, que é quando o lead
      // entrou no quadro Vendas). Fallback para created_at se a data não estiver preenchida.
      const deCol=(board.columns||[]).find(c=>c.tipo==="date"&&c.nome==="Data Entrada");
      const anchorDate=item=>{
        const raw=deCol?item.values?.[deCol.id]:null;
        if(typeof raw==="string"&&/^\d{4}-\d{2}-\d{2}/.test(raw)){
          const[y,m,dd]=raw.slice(0,10).split("-").map(Number);
          return new Date(y,m-1,dd); // parse LOCAL, alinhado ao cálculo de mês abaixo (evita erro de fuso)
        }
        return new Date(item.created_at);
      };
      const inVendasPeriod=isVendas?item=>{
        const d=anchorDate(item);
        if(vendasFilter.mode==="month"&&vendasFilter.month){
          const[y,m]=vendasFilter.month.split("-");
          const s=new Date(parseInt(y),parseInt(m)-1,1);
          const e=new Date(parseInt(y),parseInt(m),1);
          return d>=s&&d<e;
        }
        if(vendasFilter.mode==="range"&&vendasFilter.start&&vendasFilter.end){
          return d>=new Date(vendasFilter.start)&&d<=new Date(vendasFilter.end+"T23:59:59");
        }
        return true;
      }:null;
      // Aplica filtro/busca nos itens — grupos com 0 resultados ficam com items=[]
      return gs.map(g=>({...g,items:applyFilters(inVendasPeriod?(g.items||[]).filter(inVendasPeriod):(g.items||[]))}));
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

    // SDR: não tem acesso ao board de Negociações — retorna vazio
    if(perms.slug==="sdr") return [];

    // Closer: vê apenas o seu grupo mãe (pelo owner_id ou group_access)
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
  },[board,perms,currentUser,groupAccess,isNegociacoes,isVendas,isInativa,filters,search,sortCfg,vendasFilter]);
  const filterCount=(filters.status?.length||0)+(filters.resp?.length||0)+(!isVendas&&filters.dateFrom?1:0)+(!isVendas&&filters.dateTo?1:0)+(!isVendas&&filters.month?1:0)+(search.trim()?1:0);

  if(loading) return <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}><Spinner size={40}/><span style={{color:"var(--text3)",fontSize:14}}>Carregando…</span></div>;
  if(!board) return (
    <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:20,
      background:"var(--bg)",padding:32}}>
      <div style={{fontSize:64,lineHeight:1}}>🔒</div>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:20,fontWeight:800,color:"var(--text)",marginBottom:8}}>Acesso Restrito</div>
        <div style={{fontSize:14,color:"var(--text3)",maxWidth:320,lineHeight:1.6}}>
          Você não tem permissão para visualizar este quadro.<br/>
          Entre em contato com um administrador para solicitar acesso.
        </div>
      </div>
    </div>
  );

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

      {/* Vendas: barra de filtro de período */}
      {isVendas&&!isMobile&&<VendasFilterBar
        vendasFilter={vendasFilter} setVendasFilter={setVendasFilter}
        visibleGroups={visibleGroups}/>}

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
                onRespChange={(iid,gid,colId,ids)=>updateResp(iid,gid,colId,ids)}
                onMoveInativa={canActions?(gid,item)=>moveInativa(gid,item):null}
                onDupNeg={null}
                onSendToNeg={isPreVendas?(gid,item)=>sendToNeg(gid,item):null}
                onSendToVendas={isNegociacoes&&perms.sendToVendas?(gid,item)=>sendToVendas(gid,item):null}
                onSendTimer={isVendas&&perms.sendToVendas?abrirEnvioTimer:null}
                sentToNegIds={sentToNegIds} sentLabel={sentLabel}
                sortCfg={sortCfg} setSortCfg={setSortCfg}
                onDragStart={handleDragStart}
                onDragOver={(e,gid)=>handleGroupDragOver(e,gid)}
                onDrop={(e,gid)=>handleGroupDrop(e,gid)}
                onItemDragOver={handleItemDragOver}
                onItemDrop={handleItemDrop}
                onRenameSubGroup={(gid,n)=>renameGroup(gid,n)}
                onDelSubGroup={gid=>delGroup(gid)}
                onEditParent={pg=>setParentGroupM(pg)}
                onDelParent={pgid=>delParentGroup(pgid)}
                onMoveStage={(item,srcGid,tgtGid)=>moveItemToGroup(item,srcGid,tgtGid)}
                canReorderGroups={canReorderGroups}
                onGroupDragStart={handleGroupHandleDragStart} onGroupDragEnd={handleGroupHandleDragEnd}/>
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
                onRespChange={(iid,colId,ids)=>updateResp(iid,group.id,colId,ids)}
                onDelItem={iid=>setConfirmM({title:"Excluir item",danger:true,message:"Excluir permanentemente?",onConfirm:()=>{delItem(group.id,iid);setConfirmM(null);}})}
                onMoveInativa={canActions?item=>moveInativa(group.id,item):null}
                onReativar={isInativa?item=>reativarLead(group.id,item):null}
                onRenovar={canRenovar?item=>setConfirmM({title:"Renovar lead",message:`Criar uma renovação para o cliente ${item.numero_cliente??"sem código"}? A cópia terá o mesmo código e virá com Razão Social, Nome, Empresa, CNPJ, Closer e SDR preenchidos.`,onConfirm:()=>{setConfirmM(null);renovarLead(group.id,item);}}):null}
                onDupNeg={null}
                onSendToNeg={isPreVendas?item=>sendToNeg(group.id,item):null}
                onSendToVendas={isNegociacoes&&perms.sendToVendas?item=>sendToVendas(group.id,item):null}
                onSendTimer={isVendas&&perms.sendToVendas?abrirEnvioTimer:null}
                sentToNegIds={sentToNegIds} sentLabel={sentLabel}
                sortCfg={sortCfg} setSortCfg={setSortCfg}
                stickyFirstCols={isVendas?3:0}
                onDragStart={handleDragStart} onDragOver={e=>handleGroupDragOver(e,group.id)} onDrop={e=>handleGroupDrop(e,group.id)}
                onItemDragOver={handleItemDragOver} onItemDrop={handleItemDrop}
                onGroupSettings={perms.isFull||perms.all?()=>setGroupSettingsM(group):null}
                canReorderGroups={canReorderGroups}
                onGroupDragStart={handleGroupHandleDragStart} onGroupDragEnd={handleGroupHandleDragEnd}
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
          onRespChange={(colId,ids)=>updateResp(selItem.id,selItem._gid,colId,ids)}
          onAddUpdate={html=>addUpdate(selItem.id,selItem._gid,html)}
          onDelUpdate={uid=>delUpdate(uid)}
          onEditUpdate={editUpdate}
          canModerar={!!(perms.all||perms.deleteAny)}
          verExcluidas={!!(perms.all||perms.viewAll)}
        />
      )}
      {showFilters&&<FilterPanel board={board} allUsers={allUsers} filters={filters} setFilters={setFilters} onClose={()=>setShowFilters(false)}/>}
      {showColMgr&&<ColumnManagerModal board={board} toast={toast} onClose={()=>setShowColMgr(false)} onRefresh={()=>loadBoard(boardId)}/>}
      {showExport&&<ExportModal board={{...board,groups:visibleGroups}} allUsers={allUsers} onClose={()=>setShowExport(false)}/>}
      {timerItem&&<TimerIntakeModal item={timerItem} board={board} allUsers={allUsers}
        onClose={()=>setTimerItem(null)} onSent={marcarEnviadoTimer}/>}
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
          {/* Na Base Inativa, o SDR só pode reativar (→ Pré-Vendas); "Mover para quadro" fica fora para ele */}
          {!(isInativa&&perms.slug==="sdr")&&<button onClick={()=>setMoveItemsM(true)}
            style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"8px 14px",fontSize:12}}>
            📦 Mover para quadro
          </button>}
          {isInativa&&<button onClick={bulkReativar}
            style={{...T.btn,background:"#16a34a",color:"#fff",padding:"8px 14px",fontSize:12}}>
            ↩️ Reativar → Pré-Vendas
          </button>}
          {/* Renovação: um lead por vez, só em Vendas */}
          {canRenovar&&(()=>{
            const alvo=selected.size===1?(()=>{const id=[...selected][0];for(const g of board.groups){const it=(g.items||[]).find(i=>i.id===id);if(it)return {gid:g.id,item:it};}return null;})():null;
            return <button onClick={()=>{
                if(!alvo){toast("Selecione apenas um lead para renovar","warning");return;}
                setConfirmM({title:"Renovar lead",message:`Criar uma renovação para o cliente ${alvo.item.numero_cliente??"sem código"}? A cópia terá o mesmo código e virá com Razão Social, Nome, Empresa, CNPJ, Closer e SDR preenchidos.`,onConfirm:()=>{setConfirmM(null);renovarLead(alvo.gid,alvo.item);}});
              }}
              title={selected.size===1?"Criar renovação deste lead":"Selecione apenas um lead"}
              style={{...T.btn,background:selected.size===1?"#7c3aed":"var(--surface3)",color:selected.size===1?"#fff":"var(--text3)",padding:"8px 14px",fontSize:12}}>
              🔄 Renovar lead
            </button>;
          })()}
          {stageOptions.length>0&&(
            <div style={{position:"relative"}}>
              <button onClick={()=>setStageMenuOpen(p=>!p)}
                style={{...T.btn,background:"rgba(255,255,255,.12)",color:"#fff",border:"1px solid rgba(255,255,255,.25)",padding:"8px 14px",fontSize:12}}>
                ↔ Mover para etapa ▾
              </button>
              {stageMenuOpen&&(
                <div style={{position:"absolute",bottom:"calc(100% + 8px)",left:0,minWidth:230,
                  background:"var(--surface)",border:"1px solid var(--border)",borderRadius:11,
                  boxShadow:"0 12px 48px var(--shadowMd)",padding:7,zIndex:1600}}>
                  {stageOptions.map(st=>(
                    <div key={st.ordem} onClick={()=>bulkMoveToStage(st.ordem)}
                      style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:7,cursor:"pointer",fontSize:13,color:"var(--text)"}}
                      onMouseEnter={e=>e.currentTarget.style.background="var(--surface3)"}
                      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                      <span style={{display:"inline-block",width:9,height:9,borderRadius:"50%",background:st.color||"#888",flexShrink:0}}/>
                      {st.nome}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
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

      {confirmM&&<ConfirmModal title={confirmM.title} message={confirmM.message} danger={confirmM.danger} onConfirm={confirmM.onConfirm} onCancel={()=>setConfirmM(null)} requirePhrase={confirmM.requirePhrase||null} boardName={confirmM.boardName||""}/>}
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
        allBoards={perms.slug==="sdr"?(allBoardsRaw||boards).filter(b=>b.nome==="Negociações"||b.nome==="Base Inativa"):boards}
        onMove={async(destBid,movedItemIds,destGroup,inativa)=>{
          setMoveItemsM(false);
          // Destino Base Inativa → move real (remove do quadro de origem), reutilizando bulkMoveInativa
          if(inativa){ await bulkMoveInativa(); return; }
          const cnt=[...selected].length;
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
  // Ramal do PABX (VTcall): liga as ligações sincronizadas ao usuário. Único por usuário.
  const salvarRamal=async(userId,ramal)=>{
    const r=String(ramal||"").replace(/\D/g,"")||null;
    const {error}=await db.from("profiles").update({ramal:r}).eq("id",userId);
    if(error){toast(error.code==="23505"?"Esse ramal já está em outro usuário.":error.message,"error");return;}
    setUsers(p=>p.map(u=>u.id===userId?{...u,ramal:r}:u));
    toast(r?`Ramal ${r} salvo`:"Ramal removido");
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
                {/* Ramal do PABX (VTcall) */}
                {perms.manageUsers
                  ?<input defaultValue={u.ramal||""} placeholder="Ramal" title="Ramal no PABX VTcall (ex.: 4602)" maxLength={6}
                      onBlur={e=>{if((e.target.value.replace(/\D/g,"")||null)!==(u.ramal||null))salvarRamal(u.id,e.target.value);}}
                      onKeyDown={e=>{if(e.key==="Enter")e.target.blur();}}
                      style={{...T.inp,width:80,padding:"7px 10px",fontSize:12}}/>
                  :u.ramal&&<span style={{fontSize:11,color:"var(--text3)"}}>📞 {u.ramal}</span>}
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
  const isInativaDest=destBoard?.nome==="Base Inativa";
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
    // Base Inativa = inativar: move real (remove do Pré-Vendas). Delega ao BoardView.
    if(isInativaDest){ onMove(null,null,null,true); return; }
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
          origin_item_id:item.id})
        .select().single();
      if(!ni) continue;
      const {data:vals}=await db.from("item_values").select("*").eq("item_id",item.id);
      const toInsert=[];
      for(const v of vals||[]){
        const tgtId=colMap[v.column_id];
        if(tgtId) toInsert.push({item_id:ni.id,column_id:tgtId,value:v.value});
      }
      if(toInsert.length) await db.from("item_values").insert(toInsert);
      // Responsáveis → coluna gêmea (mesmo nome e tipo) via colMap; ignora legacy (column_id null)
      const {data:srcResps}=await db.from("item_responsables").select("*").eq("item_id",item.id);
      const respIns=[];
      for(const r of srcResps||[]){
        const tgtCid=r.column_id?colMap[r.column_id]:null;
        if(tgtCid) respIns.push({item_id:ni.id,user_id:r.user_id,column_id:tgtCid});
      }
      if(respIns.length) await db.from("item_responsables").insert(respIns);
      const {data:upds}=await db.from("item_updates").select("*").eq("item_id",item.id);
      if(upds?.length)
        await db.from("item_updates").insert(upds.map(u=>({item_id:ni.id,content:u.content,created_by:u.created_by})));
    }
    setMoving(false);
    const isVendasDest=destBoard?.nome==="Vendas";
    const destGroupObj=destGroups.find(g=>g.id===destGroupId)||{nome:"Vendas"};
    // Passa IDs dos itens originais do Pré-Vendas quando destino é Negociações
    // Passa destGroupObj quando destino é Vendas para o email ser disparado no BoardPage
    onMove(destBoardId, (isNegDest||isVendasDest)?items.map(i=>i.id):null, isVendasDest?destGroupObj:null);
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
        {isInativaDest
          ? <><strong style={{color:"var(--text)"}}>{items.length} lead(s)</strong> serão <strong style={{color:"var(--text)"}}>movidos</strong> para a Base Inativa e removidos do quadro atual (com dados, responsáveis e atualizações).</>
          : <><strong style={{color:"var(--text)"}}>{items.length} item(s)</strong> serão copiados para o destino. Colunas com o <strong>mesmo nome e tipo</strong> serão preenchidas automaticamente.</>}
      </div>

      <label style={T.lbl}>Quadro de destino</label>
      <select value={destBoardId} onChange={e=>setDestBoardId(e.target.value)}
        style={{...T.inp,marginBottom:18}}>
        <option value="">— Selecionar quadro —</option>
        {otherBoards.map(b=><option key={b.id} value={b.id}>{b.icon||"📋"} {b.nome}</option>)}
      </select>

      {destBoardId&&(isInativaDest
        ? <p style={{fontSize:12.5,color:"var(--text3)",marginTop:4,lineHeight:1.6}}>
            📁 Os leads selecionados serão movidos para a Base Inativa e <strong style={{color:"var(--text2)"}}>removidos do quadro atual</strong>. O grupo de destino é definido automaticamente.
          </p>
        : <>
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
function Sidebar({boards,currentBoardId,currentUser,wsNome,dark,onToggleDark,perms,onSelectBoard,onAddBoard,onEditBoard,onDelBoard,onReorderBoards,onBoardSettings,onDuplicateBoard,onProfile,onAdmin,onDashboard,onAtividade=null,onMetas=null,onLogout,open,onClose}) {
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
        {/* Metas e Recompensas: todos veem (SDR e Closer só os próprios números) */}
        {onMetas&&<div onClick={()=>{onMetas();if(isMobile)onClose();}}
          style={{display:"flex",alignItems:"center",gap:9,padding:"9px 11px",borderRadius:9,cursor:"pointer",
            color:"rgba(255,255,255,.65)",fontSize:13,marginBottom:onDashboard?2:10,background:"transparent"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.1)"}
          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <span style={{fontSize:17,flexShrink:0}}>🏆</span>
          <span style={{flex:1}}>Metas e Recompensas</span>
        </div>}
        {!onDashboard&&onMetas&&<div style={{height:1,background:"rgba(255,255,255,.1)",marginBottom:10}}/>}
        {/* Dashboard link (só aparece para cargos com acesso) */}
        {onDashboard&&<>
          <div onClick={()=>{onDashboard();if(isMobile)onClose();}}
            style={{display:"flex",alignItems:"center",gap:9,padding:"9px 11px",borderRadius:9,cursor:"pointer",
              color:"rgba(255,255,255,.65)",fontSize:13,marginBottom:10,
              background:"transparent"}}
            onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.1)"}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <span style={{fontSize:17,flexShrink:0}}>📊</span>
            <span style={{flex:1}}>Dashboard</span>
          </div>
          {onAtividade&&<div onClick={()=>{onAtividade();if(isMobile)onClose();}}
            style={{display:"flex",alignItems:"center",gap:9,padding:"9px 11px",borderRadius:9,cursor:"pointer",
              color:"rgba(255,255,255,.65)",fontSize:13,marginBottom:10,background:"transparent"}}
            onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.1)"}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <span style={{fontSize:17,flexShrink:0}}>📞</span>
            <span style={{flex:1}}>Atividade SDR</span>
          </div>}
          <div style={{height:1,background:"rgba(255,255,255,.1)",marginBottom:10}}/>
        </>}
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
  const {signIn,resetPassword}=useAuth();
  const toast=useToast();
  const [view,setView]=useState("login"); // "login" | "forgot" | "sent"
  const [email,setEmail]=useState("");
  const [senha,setSenha]=useState("");
  const [busy,setBusy]=useState(false);

  const foc=e=>e.target.style.borderColor="var(--blue)";
  const blr=e=>e.target.style.borderColor="var(--border)";

  const submitLogin=async e=>{
    e.preventDefault(); setBusy(true);
    const {error}=await signIn(email.trim().toLowerCase(),senha);
    if(error) toast(
      error.message==="Invalid login credentials"?"E-mail ou senha incorretos":error.message,
      "error"
    );
    setBusy(false);
  };

  const submitReset=async e=>{
    e.preventDefault(); setBusy(true);
    const {error}=await resetPassword(email.trim().toLowerCase());
    if(error){ toast("Erro ao enviar e-mail: "+error.message,"error"); setBusy(false); return; }
    setView("sent");
    setBusy(false);
  };

  // ── Tela: e-mail enviado
  if(view==="sent") return (
    <AuthLayout>
      <div style={{textAlign:"center",padding:"8px 0 24px"}}>
        <div style={{fontSize:48,marginBottom:16}}>📬</div>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:10}}>E-mail enviado!</div>
        <div style={{fontSize:13,color:"var(--text2)",lineHeight:1.7,marginBottom:24}}>
          Enviamos um link de redefinição para<br/>
          <strong style={{color:"var(--text)"}}>{email}</strong>.<br/>
          Verifique sua caixa de entrada e spam.
        </div>
        <button onClick={()=>{setView("login");setSenha("");}}
          style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"12px",fontSize:14}}>
          Voltar para o login
        </button>
      </div>
    </AuthLayout>
  );

  // ── Tela: esqueci a senha
  if(view==="forgot") return (
    <AuthLayout>
      <div style={{marginBottom:24}}>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:4}}>Redefinir senha</div>
        <div style={{fontSize:13,color:"var(--text2)",lineHeight:1.6}}>
          Digite seu e-mail e enviaremos um link para você criar uma nova senha.
        </div>
      </div>
      <form onSubmit={submitReset}>
        <label style={T.lbl}>E-mail</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoFocus
          placeholder="nome@empresa.com"
          style={{...T.inp,marginBottom:22}} onFocus={foc} onBlur={blr}/>
        <button type="submit" disabled={busy||!email}
          style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"13px",fontSize:15,
            opacity:busy||!email?0.6:1}}>
          {busy?"Enviando…":"Enviar link de redefinição"}
        </button>
      </form>
      <p style={{textAlign:"center",fontSize:13,color:"var(--text3)",marginTop:20,marginBottom:0}}>
        <button onClick={()=>setView("login")}
          style={{background:"none",border:"none",color:"var(--blue)",cursor:"pointer",fontWeight:700,fontSize:13}}>
          ← Voltar para o login
        </button>
      </p>
    </AuthLayout>
  );

  // ── Tela: login principal
  return (
    <AuthLayout>
      <div style={{marginBottom:24}}>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:4}}>Bem-vindo de volta</div>
        <div style={{fontSize:13,color:"var(--text2)"}}>Entre com suas credenciais para acessar o CRM.</div>
      </div>
      <form onSubmit={submitLogin}>
        <label style={T.lbl}>E-mail</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoFocus
          placeholder="nome@empresa.com"
          style={{...T.inp,marginBottom:16}} onFocus={foc} onBlur={blr}/>

        <label style={T.lbl}>Senha</label>
        <input type="password" value={senha} onChange={e=>setSenha(e.target.value)} required
          placeholder="Sua senha"
          style={{...T.inp,marginBottom:10}} onFocus={foc} onBlur={blr}/>

        <div style={{textAlign:"right",marginBottom:20}}>
          <button type="button" onClick={()=>setView("forgot")}
            style={{background:"none",border:"none",color:"var(--blue)",cursor:"pointer",
              fontSize:12,fontWeight:600,padding:0}}>
            Esqueci minha senha
          </button>
        </div>

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

// Registros do quadro Vendas que NÃO são venda (decisão do Mauricio): entram no
// quadro para controle, mas ficam FORA de receita, contagem, ticket, funil e conversão.
const PLANOS_NAO_VENDA=["Potencializa Amazon"];

// ═══════════════════════════════════════════════════════════════════════════
// ATIVIDADE SDR: ligações sincronizadas do PABX VTcall (tabela vt_ligacoes).
// Fonte: Edge Function vtcall-sync (puxa /API/ligacoes do PABX). O banco deriva
// ramal → usuário, direção, telefone e o lead casado pelo telefone.
// Regras de contagem:
//   - "Ligações"  = todas as chamadas em que o ramal do SDR participou, de saída
//                   ou de entrada (internas entre ramais ficam fora).
//   - "Atendidas" = status ATENDIDA.
//   - "Conversas" = atendidas com 60 segundos ou mais (proxy de conversa real).
//   - "Leads"     = telefones distintos contatados no período.
//   - Meta diária por SDR: editável na tela (fica salva neste navegador).
//   - Todo contador, barra e número da tabela por pessoa é clicável: abre na
//     lista só as ligações que compõem aquele número (recorte).
//   - Sem campos de IA: o recurso não está contratado na VTcall.
// ═══════════════════════════════════════════════════════════════════════════
const VT_META_DIARIA_PADRAO=50;
const VT_CORES=["#3b82f6","#f59e0b","#10b981","#a855f7","#ef4444","#06b6d4","#ec4899","#84cc16"];
const segParaTexto=s=>{s=Math.max(0,Math.round(s||0));const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),x=s%60;return h?`${h}h ${String(m).padStart(2,"0")}m`:m?`${m}m ${String(x).padStart(2,"0")}s`:`${x}s`;};
const ymdLocal=d=>{const p=n=>String(n).padStart(2,"0");return`${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`;};
const fmtFoneBR=d=>{if(!d)return"";const s=String(d);if(s.length===11)return`(${s.slice(0,2)}) ${s.slice(2,7)}-${s.slice(7)}`;if(s.length===10)return`(${s.slice(0,2)}) ${s.slice(2,6)}-${s.slice(6)}`;return s;};

function AtividadeSdrPage({onBack,allUsers,perms}){
  const{isMobile,width}=useBreakpoint();
  const toast=useToast();
  const hoje=ymdLocal(new Date());
  const [preset,setPreset]=useState("7d");          // hoje | 7d | mes | custom
  const [ini,setIni]=useState(()=>{const d=new Date();d.setDate(d.getDate()-6);return ymdLocal(d);});
  const [fim,setFim]=useState(hoje);
  const [ligacoes,setLigacoes]=useState([]);
  const [leadsNome,setLeadsNome]=useState({});       // item_id → "Empresa · Nome"
  const [estado,setEstado]=useState(null);           // vt_sync_estado
  const [loading,setLoading]=useState(true);
  const [syncing,setSyncing]=useState(false);
  const [sdrSel,setSdrSel]=useState("todos");
  const [soComLead,setSoComLead]=useState(false);
  const [metaDiaria,setMetaDiaria]=useState(()=>{try{return Number(localStorage.getItem("vt_meta_diaria"))||VT_META_DIARIA_PADRAO;}catch{return VT_META_DIARIA_PADRAO;}});
  const [editMeta,setEditMeta]=useState(false);
  const [metaInput,setMetaInput]=useState("");
  const [recorte,setRecorte]=useState(null);              // {rot,fn}: ligações que compõem um número clicado
  const [sortP,setSortP]=useState({k:"total",dir:-1});     // ordenação da tabela por pessoa
  const [sortL,setSortL]=useState({k:"hora",dir:-1});      // ordenação da lista de ligações
  const [fStatus,setFStatus]=useState("todos");
  const [fDur,setFDur]=useState("todos");
  const [fLead,setFLead]=useState("todos");
  const [busca,setBusca]=useState("");
  const [rankMet,setRankMet]=useState("total");
  const [cmpSel,setCmpSel]=useState(null);                 // null = todos marcados

  // Quem aparece: usuários com ramal cadastrado (SDRs e closers com telefone)
  const comRamal=useMemo(()=>(allUsers||[]).filter(u=>u.ramal).sort((a,b)=>String(a.ramal).localeCompare(String(b.ramal))),[allUsers]);
  const sdrs=useMemo(()=>comRamal.filter(u=>u.roles?.slug==="sdr"),[comRamal]);
  const outros=useMemo(()=>comRamal.filter(u=>u.roles?.slug!=="sdr"),[comRamal]);

  const aplicarPreset=p=>{
    setPreset(p);
    const d=new Date();
    if(p==="hoje"){setIni(hoje);setFim(hoje);}
    else if(p==="7d"){const a=new Date(d);a.setDate(a.getDate()-6);setIni(ymdLocal(a));setFim(hoje);}
    else if(p==="mes"){setIni(`${hoje.slice(0,7)}-01`);setFim(hoje);}
  };

  const carregar=async()=>{
    setLoading(true);
    try{
      // Datas como texto no fuso local; a coluna é timestamptz, então compara em ISO com offset
      const de=`${ini}T00:00:00-03:00`, ate=`${fim}T23:59:59-03:00`;
      const PAG=1000, out=[];
      for(let o=0;;o+=PAG){
        const {data,error}=await db.from("vt_ligacoes")
          .select("linkedid,hora,origem,destino,status,duracao_seg,audio,ramal,direcao,telefone,profile_id,item_id,item_board_id")
          .gte("hora",de).lte("hora",ate).order("hora",{ascending:false}).range(o,o+PAG-1);
        if(error){toast("Não foi possível carregar as ligações: "+error.message,"error");break;}
        out.push(...(data||[]));
        if(!data||data.length<PAG)break;
      }
      setLigacoes(out);
      // Nomes dos leads casados (Empresa / Nome), em lotes
      const ids=[...new Set(out.map(l=>l.item_id).filter(Boolean))];
      const nomes={};
      for(let i=0;i<ids.length;i+=200){
        const chunk=ids.slice(i,i+200);
        const {data:vals}=await db.from("item_values").select("item_id,value,columns!inner(nome)").in("item_id",chunk).in("columns.nome",["Empresa","Nome","Razão Social"]);
        (vals||[]).forEach(v=>{
          const raw=(v.value&&typeof v.value==="object"&&"value" in v.value)?v.value.value:v.value;
          const t=typeof raw==="string"?raw.trim():"";
          if(!t)return;
          nomes[v.item_id]=nomes[v.item_id]||{};
          nomes[v.item_id][v.columns?.nome]=t;
        });
      }
      const mapa={};
      Object.entries(nomes).forEach(([id,n])=>{mapa[id]=[n["Empresa"]&&n["Empresa"]!=="Empresa"?n["Empresa"]:n["Razão Social"],n["Nome"]&&n["Nome"]!=="Nome"?n["Nome"]:null].filter(Boolean).join(" · ")||"(lead sem nome)";});
      setLeadsNome(mapa);
      const {data:est}=await db.from("vt_sync_estado").select("*").eq("id",1).maybeSingle();
      setEstado(est||null);
    }finally{setLoading(false);}
  };
  useEffect(()=>{carregar();},[ini,fim]); // eslint-disable-line react-hooks/exhaustive-deps

  // Sincroniza com o PABX (Edge Function) e recarrega
  const sincronizar=async(force=false)=>{
    setSyncing(true);
    try{
      const {data,error}=await db.functions.invoke("vtcall-sync",{body:force?{since:`${ini} 00:00:00`,until:`${fim} 23:59:59`}:{}});
      if(error){
        const corpo=await corpoDaResposta(error);
        toast(corpo?.message||"Não foi possível sincronizar com a VTcall.","error");
      }else if(data&&data.ok===false){
        toast(data.message||"A VTcall respondeu com erro.","error");
      }else{
        toast(data?.message||"Ligações sincronizadas.","success");
      }
      await carregar();
    }finally{setSyncing(false);}
  };
  // Ao abrir, sincroniza se a última vez foi há mais de 10 minutos
  useEffect(()=>{
    (async()=>{
      const {data:est}=await db.from("vt_sync_estado").select("ultima_sync").eq("id",1).maybeSingle();
      const velha=!est?.ultima_sync||(Date.now()-new Date(est.ultima_sync).getTime())>10*60*1000;
      if(velha) sincronizar(false);
    })();
  },[]); // eslint-disable-line react-hooks/exhaustive-deps

  // ── Métricas
  const uteis=useMemo(()=>ligacoes.filter(l=>l.direcao==="saida"||l.direcao==="entrada"),[ligacoes]);
  const filtradas=useMemo(()=>uteis.filter(l=>(sdrSel==="todos"||l.profile_id===sdrSel)&&(!soComLead||l.item_id)),[uteis,sdrSel,soComLead]);
  const diasPeriodo=useMemo(()=>{const a=new Date(ini+"T00:00:00"),b=new Date(fim+"T00:00:00");return Math.max(1,Math.round((b-a)/86400000)+1);},[ini,fim]);
  const contaUteis=(de,ate)=>{let n=0;const d=new Date(de+"T00:00:00");const b=new Date(ate+"T00:00:00");while(d<=b){const wd=d.getDay();if(wd!==0&&wd!==6)n++;d.setDate(d.getDate()+1);}return n;};
  const diasUteis=useMemo(()=>Math.max(1,contaUteis(ini,fim)),[ini,fim]);

  // Testes reutilizados pelos contadores e pelos recortes (clique)
  const ehAtendida=l=>l.status==="ATENDIDA";
  const ehConversa=l=>l.status==="ATENDIDA"&&(l.duracao_seg||0)>=60;
  const ehCurta=l=>l.status==="ATENDIDA"&&(l.duracao_seg||0)<10;

  const resumoPor=lista=>{
    const atend=lista.filter(ehAtendida);
    const tempo=atend.reduce((s,l)=>s+(l.duracao_seg||0),0);
    const fones=new Set(lista.map(l=>l.telefone).filter(Boolean));
    const comLead=lista.filter(l=>l.item_id).length;
    return {total:lista.length,saida:lista.filter(l=>l.direcao==="saida").length,entrada:lista.filter(l=>l.direcao==="entrada").length,
      atendidas:atend.length,curtas:lista.filter(ehCurta).length,taxa:lista.length?atend.length/lista.length*100:0,conversas:lista.filter(ehConversa).length,tempo,
      tempoMedio:atend.length?tempo/atend.length:0,leads:fones.size,comLead,semLead:lista.length-comLead};
  };
  const geral=useMemo(()=>resumoPor(filtradas),[filtradas]); // eslint-disable-line react-hooks/exhaustive-deps
  const porSdr=useMemo(()=>comRamal.map((u,i)=>({u,cor:VT_CORES[i%VT_CORES.length],...resumoPor(uteis.filter(l=>l.profile_id===u.id&&(!soComLead||l.item_id)))})),[comRamal,uteis,soComLead]); // eslint-disable-line react-hooks/exhaustive-deps
  const semRamal=useMemo(()=>uteis.filter(l=>!l.profile_id).length,[uteis]);

  // Meta: em "Todos", compara só as ligações dos SDRs com a meta somada dos SDRs
  // (closers com ramal aparecem na página, mas não entram na meta do time).
  const idsSdr=useMemo(()=>new Set(sdrs.map(u=>u.id)),[sdrs]);
  const nSdrMeta=sdrSel==="todos"?Math.max(1,sdrs.length):1;
  const totalMeta=useMemo(()=>sdrSel==="todos"&&sdrs.length?filtradas.filter(l=>idsSdr.has(l.profile_id)).length:filtradas.length,[filtradas,sdrSel,sdrs,idsSdr]);
  const alvoMeta=metaDiaria*diasUteis*nSdrMeta;
  const pctMeta=alvoMeta?totalMeta/alvoMeta*100:0;
  const faltaMeta=Math.max(0,alvoMeta-totalMeta);
  // Meta de HOJE (só quando o período inclui hoje e hoje é dia útil)
  const hojeNoPeriodo=ini<=hoje&&fim>=hoje&&![0,6].includes(new Date(hoje+"T00:00:00").getDay());
  const feitasHoje=useMemo(()=>filtradas.filter(l=>ymdLocal(new Date(l.hora))===hoje&&(sdrSel!=="todos"||!sdrs.length||idsSdr.has(l.profile_id))).length,[filtradas,hoje,sdrSel,sdrs,idsSdr]);
  const alvoHoje=metaDiaria*nSdrMeta;
  const corMeta=p=>p>=100?"#059669":p>=70?"#d97706":"#dc2626";

  // Por dia (barras): total e atendidas
  const porDia=useMemo(()=>{
    const m={};
    const d=new Date(ini+"T00:00:00");const b=new Date(fim+"T00:00:00");
    while(d<=b){m[ymdLocal(d)]={total:0,atend:0};d.setDate(d.getDate()+1);}
    filtradas.forEach(l=>{const k=ymdLocal(new Date(l.hora));if(m[k]){m[k].total++;if(l.status==="ATENDIDA")m[k].atend++;}});
    return Object.entries(m).map(([dia,v])=>({dia,...v}));
  },[filtradas,ini,fim]);
  const maxDia=Math.max(1,...porDia.map(d=>d.total));

  // Por hora do dia (0..23)
  const porHora=useMemo(()=>{
    const h=Array.from({length:24},()=>({total:0,atend:0}));
    filtradas.forEach(l=>{const x=new Date(l.hora).getHours();h[x].total++;if(l.status==="ATENDIDA")h[x].atend++;});
    return h;
  },[filtradas]);
  const maxHora=Math.max(1,...porHora.map(h=>h.total));

  const salvarMeta=()=>{const n=parseInt(metaInput,10);if(!n||n<=0){toast("Informe um número maior que zero","error");return;}setMetaDiaria(n);try{localStorage.setItem("vt_meta_diaria",String(n));}catch{/* sem storage */}setEditMeta(false);};

  const nomeU=id=>allUsers.find(u=>u.id===id)?.nome||"(sem ramal)";
  const primeiroNome=n=>String(n||"").split(" ")[0];
  const corStatus=s=>s==="ATENDIDA"?"#059669":s==="NÃO ATENDIDA"?"#d97706":(s==="OCUPADO"||s==="FALHOU")?"#dc2626":"var(--text3)";
  const gridCols=width>=1400?"repeat(4,1fr)":width>=1024?"repeat(4,1fr)":width>=640?"repeat(2,1fr)":"1fr";

  // ── Recorte (clique em contador, barra ou célula): mostra na lista só as
  // ligações que compõem aquele número. Um recorte por vez.
  const listaRef=useRef(null);
  const abrirRecorte=(rot,fn,porPessoa=false)=>{setRecorte({rot,fn,porPessoa});setTimeout(()=>listaRef.current?.scrollIntoView({behavior:"smooth",block:"start"}),50);};
  const recortePessoa=(u,rot,fn)=>abrirRecorte(`${primeiroNome(u.nome)}: ${rot}`,l=>l.profile_id===u.id&&(!fn||fn(l)),true);
  // O recorte por pessoa precisa enxergar a pessoa mesmo com outro filtro ativo
  const baseLista=useMemo(()=>{
    if(!recorte) return filtradas;
    const base=uteis.filter(l=>!soComLead||l.item_id);
    return base.filter(l=>(sdrSel==="todos"||l.profile_id===sdrSel||recorte.porPessoa)&&recorte.fn(l));
  },[recorte,filtradas,uteis,sdrSel,soComLead]);

  // ── Lista de ligações: filtros + ordenação
  const statusDisponiveis=useMemo(()=>[...new Set(uteis.map(l=>l.status).filter(Boolean))].sort(),[uteis]);
  const listaFinal=useMemo(()=>{
    const q=busca.trim().toLowerCase(), qd=q.replace(/\D/g,"");
    let out=baseLista.filter(l=>{
      if(fStatus!=="todos"&&l.status!==fStatus) return false;
      if(fDur==="0"&&(l.duracao_seg||0)!==0) return false;
      if(fDur==="lt10"&&!((l.duracao_seg||0)>0&&(l.duracao_seg||0)<10)) return false;
      if(fDur==="10a59"&&!((l.duracao_seg||0)>=10&&(l.duracao_seg||0)<60)) return false;
      if(fDur==="ge60"&&(l.duracao_seg||0)<60) return false;
      if(fLead==="com"&&!l.item_id) return false;
      if(fLead==="sem"&&l.item_id) return false;
      if(q){
        const nome=(l.item_id?leadsNome[l.item_id]||"":"").toLowerCase();
        const pessoa=nomeU(l.profile_id).toLowerCase();
        if(!(nome.includes(q)||pessoa.includes(q)||(qd&&String(l.telefone||"").includes(qd)))) return false;
      }
      return true;
    });
    const k=sortL.k, dir=sortL.dir;
    const val=l=>k==="hora"?new Date(l.hora).getTime():k==="pessoa"?nomeU(l.profile_id):k==="telefone"?String(l.telefone||""):k==="lead"?(l.item_id?leadsNome[l.item_id]||"":"~"):k==="status"?String(l.status||""):k==="duracao"?(l.duracao_seg||0):0;
    out=[...out].sort((a,b)=>{const x=val(a),y=val(b);return (typeof x==="number"?x-y:String(x).localeCompare(String(y),"pt-BR"))*dir;});
    return out;
  },[baseLista,fStatus,fDur,fLead,busca,sortL,leadsNome]); // eslint-disable-line react-hooks/exhaustive-deps
  const temFiltroLista=fStatus!=="todos"||fDur!=="todos"||fLead!=="todos"||busca.trim()!=="";
  const limparFiltrosLista=()=>{setFStatus("todos");setFDur("todos");setFLead("todos");setBusca("");};

  // ── Tabela por pessoa: ordenação
  const porSdrOrd=useMemo(()=>{
    const k=sortP.k, dir=sortP.dir;
    const val=r=>k==="nome"?r.u.nome:k==="ramal"?Number(r.u.ramal)||0:k==="meta"?r.total:r[k]??0;
    return [...porSdr].sort((a,b)=>{const x=val(a),y=val(b);return (typeof x==="number"?x-y:String(x).localeCompare(String(y),"pt-BR"))*dir;});
  },[porSdr,sortP]);
  const ordena=(estado,setEstadoFn,k,dirPadrao=-1)=>setEstadoFn(estado.k===k?{k,dir:-estado.dir}:{k,dir:dirPadrao});
  const seta=(estado,k)=>estado.k===k?(estado.dir===1?" ▲":" ▼"):"";

  // ── Ranking e comparação
  const VT_METRICAS=[
    {k:"total",rot:"Ligações",fmt:v=>v},
    {k:"atendidas",rot:"Atendidas",fmt:v=>v},
    {k:"conversas",rot:"Conversas ≥1min",fmt:v=>v},
    {k:"tempo",rot:"Tempo em conversa",fmt:segParaTexto},
    {k:"leads",rot:"Contatos distintos",fmt:v=>v},
    {k:"taxa",rot:"Taxa de atendimento",fmt:v=>`${v.toFixed(0)}%`},
    {k:"tempoMedio",rot:"Média por atendida",fmt:segParaTexto},
  ];
  const metRank=VT_METRICAS.find(m=>m.k===rankMet)||VT_METRICAS[0];
  const ranking=useMemo(()=>[...porSdr].sort((a,b)=>(b[rankMet]??0)-(a[rankMet]??0)),[porSdr,rankMet]);
  const maxRank=Math.max(1,...ranking.map(r=>r[rankMet]??0));
  const medalha=i=>i===0?"🥇":i===1?"🥈":i===2?"🥉":`${i+1}º`;
  const cmpLista=useMemo(()=>porSdr.filter(r=>cmpSel===null?true:cmpSel.includes(r.u.id)),[porSdr,cmpSel]);
  const alternaCmp=id=>{const atual=cmpSel===null?porSdr.map(r=>r.u.id):cmpSel;setCmpSel(atual.includes(id)?atual.filter(x=>x!==id):[...atual,id]);};
  // Série por dia de cada pessoa comparada
  const cmpDias=useMemo(()=>porDia.map(d=>({dia:d.dia,por:Object.fromEntries(cmpLista.map(r=>[r.u.id,0]))})).map(row=>{
    uteis.forEach(l=>{if(row.por[l.profile_id]!=null&&ymdLocal(new Date(l.hora))===row.dia&&(!soComLead||l.item_id))row.por[l.profile_id]++;});
    return row;
  }),[porDia,cmpLista,uteis,soComLead]);
  const maxCmpDia=Math.max(1,...cmpDias.flatMap(r=>Object.values(r.por)));

  // Estilos compartilhados
  const sCard={background:"var(--surface)",border:"1px solid var(--border)",borderRadius:14,padding:16};
  const sTh={padding:"6px 8px",borderBottom:"1px solid var(--border)",whiteSpace:"nowrap",cursor:"pointer",userSelect:"none"};
  const sNum={cursor:"pointer",textDecoration:"underline dotted",textUnderlineOffset:3};
  const sSel={...T.inp,width:"auto",padding:"6px 9px",fontSize:12,borderRadius:8};
  const cartao=(lbl,val,sub,clr,onClick)=>(
    <div onClick={onClick} title={onClick?"Clique para ver as ligações que compõem este número":undefined}
      style={{...sCard,padding:"16px 18px",cursor:onClick?"pointer":"default"}}>
      <div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.5}}>{lbl}</div>
      <div style={{fontSize:26,fontWeight:900,color:clr||"var(--text)",marginTop:6}}>{val}</div>
      {sub&&<div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>{sub}</div>}
    </div>
  );
  const linkSub=(txt,rot,fn)=><a onClick={e=>{e.stopPropagation();abrirRecorte(rot,fn);}} style={{cursor:"pointer",textDecoration:"underline dotted",textUnderlineOffset:3}}>{txt}</a>;

  return(
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",background:"var(--bg)"}}>
      {/* Cabeçalho */}
      <div style={{background:"var(--surface)",borderBottom:"1px solid var(--border)",padding:isMobile?"12px 16px":"14px 24px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",flexShrink:0}}>
        <button onClick={onBack} style={{...T.btn,background:"none",border:"1.5px solid var(--border)",color:"var(--text)",padding:"7px 14px",fontSize:12}}>← Voltar</button>
        <div style={{fontSize:isMobile?16:18,fontWeight:900,color:"var(--text)"}}>📞 Atividade SDR</div>
        {(loading||syncing)&&<Spinner size={15}/>}
        <div style={{display:"flex",borderRadius:8,overflow:"hidden",border:"1.5px solid var(--border)"}}>
          {[["hoje","Hoje"],["7d","7 dias"],["mes","Mês"],["custom","Período"]].map(([p,l])=>(
            <button key={p} onClick={()=>aplicarPreset(p)} style={{...T.btn,borderRadius:0,padding:"6px 12px",fontSize:11,background:preset===p?"var(--blue)":"var(--surface)",color:preset===p?"#fff":"var(--text2)"}}>{l}</button>
          ))}
        </div>
        {preset==="custom"&&<div style={{display:"flex",alignItems:"center",gap:6}}>
          <input type="date" value={ini} max={fim} onChange={e=>setIni(e.target.value)} style={{...T.inp,width:140,padding:"7px 10px",fontSize:12,borderRadius:8}}/>
          <span style={{color:"var(--text3)",fontSize:12}}>até</span>
          <input type="date" value={fim} min={ini} max={hoje} onChange={e=>setFim(e.target.value)} style={{...T.inp,width:140,padding:"7px 10px",fontSize:12,borderRadius:8}}/>
        </div>}
        <select value={sdrSel} onChange={e=>{setSdrSel(e.target.value);setRecorte(null);}} style={{...T.inp,width:"auto",padding:"7px 10px",fontSize:12,borderRadius:8}}>
          <option value="todos">Todos com ramal</option>
          {sdrs.length>0&&<optgroup label="SDRs">{sdrs.map(u=><option key={u.id} value={u.id}>{u.nome} · {u.ramal}</option>)}</optgroup>}
          {outros.length>0&&<optgroup label="Outros ramais">{outros.map(u=><option key={u.id} value={u.id}>{u.nome} · {u.ramal}</option>)}</optgroup>}
        </select>
        <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"var(--text2)",cursor:"pointer"}}>
          <input type="checkbox" checked={soComLead} onChange={e=>setSoComLead(e.target.checked)}/> só ligações com lead no CRM
        </label>
        <div style={{flex:1}}/>
        <div style={{fontSize:10,color:"var(--text3)"}}>
          {estado?.ultima_sync?`PABX sincronizado ${new Date(estado.ultima_sync).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}`:"PABX ainda não sincronizado"}
          {estado?.ultimo_erro&&<span style={{color:"#dc2626"}}> · último erro: {estado.ultimo_erro}</span>}
        </div>
        <button onClick={()=>sincronizar(false)} disabled={syncing} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"7px 12px",fontSize:12,border:"1px solid var(--border)"}}>↻ Sincronizar</button>
        {(perms?.all||perms?.viewAll)&&<button onClick={()=>sincronizar(true)} disabled={syncing} title="Baixa de novo todas as ligações do período escolhido" style={{...T.btn,background:"none",color:"var(--text3)",padding:"7px 10px",fontSize:11,border:"1px solid var(--border)"}}>Rebaixar período</button>}
      </div>

      <div style={{flex:1,overflowY:"auto",padding:isMobile?"12px":"20px 24px"}}>
        {comRamal.length===0&&(
          <div style={{background:"#fef3c7",color:"#92400e",border:"1px solid #fcd34d",borderRadius:12,padding:"12px 16px",marginBottom:16,fontSize:13}}>
            Nenhum usuário tem ramal cadastrado. Preencha o ramal de cada SDR em Admin › Usuários para as ligações aparecerem por pessoa.
          </div>
        )}
        {semRamal>0&&(
          <div style={{fontSize:12,color:"var(--text3)",marginBottom:12}}>⚠ {semRamal} ligação(ões) do período em ramais sem usuário no CRM (não entram nas contagens por pessoa).</div>
        )}

        {/* ── META em destaque ── */}
        <div style={{background:`linear-gradient(135deg, ${corMeta(pctMeta)}22, var(--surface) 55%)`,border:`2px solid ${corMeta(pctMeta)}`,borderRadius:18,padding:isMobile?"16px":"22px 26px",marginBottom:20}}>
          <div style={{display:"flex",alignItems:"flex-start",gap:24,flexWrap:"wrap"}}>
            <div style={{minWidth:190}}>
              <div style={{fontSize:12,fontWeight:800,color:"var(--text2)",textTransform:"uppercase",letterSpacing:1}}>🎯 Meta de ligações{sdrSel==="todos"?(sdrs.length?" · time de SDRs":""):` · ${primeiroNome(nomeU(sdrSel))}`}</div>
              <div style={{fontSize:isMobile?52:68,fontWeight:900,lineHeight:1,color:corMeta(pctMeta),marginTop:6}}>{Math.round(pctMeta)}%</div>
              <div style={{fontSize:13,color:"var(--text2)",marginTop:6,fontWeight:700}}>
                <a onClick={()=>abrirRecorte(sdrSel==="todos"&&sdrs.length?"Ligações dos SDRs (meta)":"Ligações (meta)",l=>sdrSel!=="todos"||!sdrs.length||idsSdr.has(l.profile_id))} style={sNum}>{totalMeta}</a> de {alvoMeta} ligações
              </div>
            </div>
            <div style={{flex:1,minWidth:260}}>
              <div style={{height:26,background:"var(--surface3)",borderRadius:13,overflow:"hidden",position:"relative"}}>
                <div style={{width:`${Math.min(pctMeta,100)}%`,height:"100%",background:corMeta(pctMeta),borderRadius:13,transition:"width .6s ease"}}/>
              </div>
              <div style={{display:"flex",gap:18,flexWrap:"wrap",marginTop:10,fontSize:13,color:"var(--text2)"}}>
                <span>{pctMeta>=100?<b style={{color:"#059669"}}>🎉 Meta do período batida</b>:<span>Faltam <b style={{color:"var(--text)"}}>{faltaMeta}</b> ligações</span>}</span>
                {hojeNoPeriodo&&<span>Hoje: <b style={{color:corMeta(alvoHoje?feitasHoje/alvoHoje*100:0)}}>{feitasHoje} de {alvoHoje}</b></span>}
                <span style={{color:"var(--text3)"}}>{metaDiaria}/dia útil{sdrSel==="todos"&&sdrs.length?` × ${sdrs.length} SDR(s)`:""} × {diasUteis} dia(s) útil(eis)</span>
                {editMeta
                  ?<span><input value={metaInput} onChange={e=>setMetaInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")salvarMeta();}} style={{...T.inp,width:64,padding:"3px 7px",fontSize:12}} placeholder={String(metaDiaria)}/> <button onClick={salvarMeta} style={{...T.btn,padding:"3px 10px",fontSize:12,background:"var(--blue)",color:"#fff"}}>ok</button></span>
                  :<a onClick={()=>{setMetaInput(String(metaDiaria));setEditMeta(true);}} style={{cursor:"pointer",color:"var(--blue)"}}>editar meta diária</a>}
              </div>
              {/* Meta individual de cada SDR */}
              {sdrSel==="todos"&&sdrs.length>0&&<div style={{display:"grid",gridTemplateColumns:width>=900?`repeat(${Math.min(sdrs.length,4)},1fr)`:"1fr",gap:12,marginTop:16}}>
                {porSdr.filter(r=>idsSdr.has(r.u.id)).map(r=>{const alvo=metaDiaria*diasUteis;const p=alvo?r.total/alvo*100:0;return(
                  <div key={r.u.id} onClick={()=>recortePessoa(r.u,"ligações")} title="Clique para ver as ligações" style={{background:"var(--surface)",border:"1px solid var(--border)",borderRadius:12,padding:"10px 12px",cursor:"pointer"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",gap:8}}>
                      <span style={{fontWeight:800,fontSize:13}}>{r.u.nome}</span>
                      <span style={{fontWeight:900,fontSize:18,color:corMeta(p)}}>{Math.round(p)}%</span>
                    </div>
                    <div style={{height:10,background:"var(--surface3)",borderRadius:5,overflow:"hidden",margin:"6px 0"}}><div style={{width:`${Math.min(p,100)}%`,height:"100%",background:corMeta(p)}}/></div>
                    <div style={{fontSize:11,color:"var(--text3)"}}>{r.total} de {alvo} · faltam {Math.max(0,alvo-r.total)}</div>
                  </div>);})}
              </div>}
            </div>
          </div>
        </div>

        {/* Cards gerais (clicáveis) */}
        <div style={{display:"grid",gridTemplateColumns:gridCols,gap:14,marginBottom:20}}>
          {cartao("Ligações",geral.total,`${geral.saida} de saída · ${geral.entrada} recebidas · ${(geral.total/diasUteis).toFixed(1)} por dia útil`,null,()=>abrirRecorte("Todas as ligações",()=>true))}
          {cartao("Atendidas",`${geral.atendidas} (${geral.taxa.toFixed(0)}%)`,<span>{linkSub(`${geral.conversas} conversas de 1 min ou mais`,"Conversas de 1 min ou mais",ehConversa)} · {linkSub(`${geral.curtas} atendidas com menos de 10s`,"Atendidas com menos de 10s",ehCurta)}</span>,"#059669",()=>abrirRecorte("Atendidas",ehAtendida))}
          {cartao("Tempo em conversa",segParaTexto(geral.tempo),`média de ${segParaTexto(geral.tempoMedio)} por atendida`,null,()=>{abrirRecorte("Atendidas, mais longas primeiro",ehAtendida);setSortL({k:"duracao",dir:-1});})}
          {cartao("Contatos distintos",geral.leads,<span>{linkSub(`${geral.comLead} ligações casadas com lead`,"Com lead no CRM",l=>!!l.item_id)} · {linkSub(`${geral.semLead} sem lead no CRM`,"Sem lead no CRM",l=>!l.item_id)}</span>,null,()=>{abrirRecorte("Todas, agrupadas por telefone",()=>true);setSortL({k:"telefone",dir:1});})}
        </div>

        {/* Ranking + Comparação */}
        <div style={{display:"grid",gridTemplateColumns:width>=1100?"1fr 2fr":"1fr",gap:14,marginBottom:20}}>
          <div style={sCard}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
              <div style={{fontWeight:800,fontSize:14}}>🏆 Ranking</div>
              <select value={rankMet} onChange={e=>setRankMet(e.target.value)} style={sSel}>{VT_METRICAS.map(m=><option key={m.k} value={m.k}>{m.rot}</option>)}</select>
            </div>
            {ranking.map((r,i)=>(
              <div key={r.u.id} onClick={()=>recortePessoa(r.u,"ligações")} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderTop:i?"1px solid var(--border)":"none",cursor:"pointer"}}>
                <div style={{width:34,textAlign:"center",fontSize:i<3?22:13,fontWeight:800,color:"var(--text3)"}}>{medalha(i)}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
                    <span style={{fontWeight:i===0?900:700,fontSize:13,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.u.nome}{r.u.roles?.slug&&r.u.roles.slug!=="sdr"&&<span style={{fontSize:10,color:"var(--text3)",marginLeft:6,fontWeight:400}}>{r.u.roles.nome}</span>}</span>
                    <span style={{fontWeight:900,fontSize:14}}>{metRank.fmt(r[rankMet]??0)}</span>
                  </div>
                  <div style={{height:6,background:"var(--surface3)",borderRadius:3,overflow:"hidden",marginTop:5}}><div style={{width:`${(r[rankMet]??0)/maxRank*100}%`,height:"100%",background:r.cor}}/></div>
                </div>
              </div>
            ))}
            {ranking.length===0&&<div style={{fontSize:12,color:"var(--text3)"}}>Nenhum usuário com ramal.</div>}
          </div>

          <div style={sCard}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
              <div style={{fontWeight:800,fontSize:14}}>⚖️ Comparar</div>
              {porSdr.map(r=>{const on=cmpSel===null||cmpSel.includes(r.u.id);return(
                <label key={r.u.id} style={{display:"flex",alignItems:"center",gap:5,fontSize:12,cursor:"pointer",padding:"3px 9px",borderRadius:14,border:`1.5px solid ${on?r.cor:"var(--border)"}`,color:on?"var(--text)":"var(--text3)"}}>
                  <input type="checkbox" checked={on} onChange={()=>alternaCmp(r.u.id)} style={{accentColor:r.cor}}/>{primeiroNome(r.u.nome)}
                </label>);})}
            </div>
            {cmpLista.length<2
              ?<div style={{fontSize:12,color:"var(--text3)"}}>Marque duas ou mais pessoas para comparar.</div>
              :<div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{color:"var(--text3)",textAlign:"left"}}>
                    <th style={{padding:"6px 8px",borderBottom:"1px solid var(--border)"}}>Métrica</th>
                    {cmpLista.map(r=><th key={r.u.id} style={{padding:"6px 8px",borderBottom:"1px solid var(--border)",whiteSpace:"nowrap"}}><span style={{display:"inline-block",width:9,height:9,borderRadius:5,background:r.cor,marginRight:6}}/>{primeiroNome(r.u.nome)}</th>)}
                  </tr></thead>
                  <tbody>
                    {VT_METRICAS.map(m=>{const mx=Math.max(...cmpLista.map(r=>r[m.k]??0));return(
                      <tr key={m.k}>
                        <td style={{padding:"7px 8px",color:"var(--text2)",whiteSpace:"nowrap"}}>{m.rot}</td>
                        {cmpLista.map(r=>{const v=r[m.k]??0;const lider=mx>0&&v===mx;return(
                          <td key={r.u.id} style={{padding:"7px 8px",minWidth:110}}>
                            <div style={{fontWeight:lider?900:600,color:lider?"var(--text)":"var(--text2)"}}>{m.fmt(v)}{lider&&<span style={{marginLeft:5,fontSize:10,color:"#059669"}}>▲ maior</span>}</div>
                            <div style={{height:5,background:"var(--surface3)",borderRadius:3,overflow:"hidden",marginTop:4}}><div style={{width:`${mx?v/mx*100:0}%`,height:"100%",background:r.cor}}/></div>
                          </td>);})}
                      </tr>);})}
                  </tbody>
                </table>
                {/* Ligações por dia, lado a lado */}
                <div style={{fontSize:11,color:"var(--text3)",margin:"14px 0 6px"}}>Ligações por dia, lado a lado · clique numa barra para ver as ligações</div>
                <div style={{height:110,display:"flex",alignItems:"flex-end",gap:diasPeriodo>20?2:8}}>
                  {cmpDias.map(d=>(
                    <div key={d.dia} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-end",height:"100%"}}>
                      <div style={{display:"flex",alignItems:"flex-end",gap:1,height:92,width:"100%",justifyContent:"center"}}>
                        {cmpLista.map(r=>(
                          <div key={r.u.id} onClick={()=>recortePessoa(r.u,`dia ${d.dia.split("-").reverse().slice(0,2).join("/")}`,l=>ymdLocal(new Date(l.hora))===d.dia)}
                            title={`${primeiroNome(r.u.nome)} em ${d.dia.split("-").reverse().join("/")}: ${d.por[r.u.id]} ligações`}
                            style={{flex:1,maxWidth:14,height:`${d.por[r.u.id]/maxCmpDia*90}px`,background:r.cor,borderRadius:"2px 2px 0 0",cursor:"pointer",minHeight:d.por[r.u.id]?2:0}}/>
                        ))}
                      </div>
                      {diasPeriodo<=31&&<div style={{fontSize:9,color:"var(--text3)",marginTop:4}}>{d.dia.slice(8)}</div>}
                    </div>
                  ))}
                </div>
              </div>}
          </div>
        </div>

        {/* Tabela por pessoa (ordenável, números clicáveis) */}
        <div style={{...sCard,marginBottom:20,overflowX:"auto"}}>
          <div style={{fontWeight:800,fontSize:14,marginBottom:10}}>Por pessoa <span style={{fontWeight:400,color:"var(--text3)",fontSize:12}}>· clique no título da coluna para ordenar · clique num número para ver as ligações</span></div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{color:"var(--text3)",textAlign:"left"}}>
              {[["nome","Pessoa",1],["ramal","Ramal",1],["total","Ligações"],["total","/dia útil"],["atendidas","Atendidas"],["conversas","Conversas ≥1min"],["curtas","Atendidas <10s"],["tempo","Tempo conversa"],["tempoMedio","Média"],["leads","Contatos"],["comLead","Com lead"],["meta","Meta"]].map(([k,h,d],i)=>(
                <th key={i} onClick={()=>ordena(sortP,setSortP,k,d||-1)} style={sTh}>{h}{seta(sortP,k)}</th>))}
            </tr></thead>
            <tbody>
              {porSdrOrd.map(r=>{
                const metaP=metaDiaria*diasUteis;const pct=metaP?r.total/metaP*100:0;
                return <tr key={r.u.id} style={{background:sdrSel===r.u.id?"var(--surface3)":"transparent"}}>
                  <td onClick={()=>{setSdrSel(sdrSel===r.u.id?"todos":r.u.id);setRecorte(null);}} title="Clique para filtrar a página por esta pessoa" style={{padding:"8px",fontWeight:700,whiteSpace:"nowrap",cursor:"pointer"}}><span style={{display:"inline-block",width:9,height:9,borderRadius:5,background:r.cor,marginRight:7}}/>{r.u.nome}{r.u.roles?.slug&&r.u.roles.slug!=="sdr"&&<span style={{fontSize:10,color:"var(--text3)",marginLeft:6}}>{r.u.roles.nome}</span>}</td>
                  <td style={{padding:"8px",color:"var(--text3)"}}>{r.u.ramal}</td>
                  <td style={{padding:"8px",fontWeight:800}}><span style={sNum} onClick={()=>recortePessoa(r.u,"ligações")}>{r.total}</span></td>
                  <td style={{padding:"8px"}}>{(r.total/diasUteis).toFixed(1)}</td>
                  <td style={{padding:"8px",color:"#059669"}}><span style={sNum} onClick={()=>recortePessoa(r.u,"atendidas",ehAtendida)}>{r.atendidas}</span> <span style={{color:"var(--text3)"}}>({r.taxa.toFixed(0)}%)</span></td>
                  <td style={{padding:"8px"}}><span style={sNum} onClick={()=>recortePessoa(r.u,"conversas de 1 min ou mais",ehConversa)}>{r.conversas}</span></td>
                  <td style={{padding:"8px"}}><span style={sNum} onClick={()=>recortePessoa(r.u,"atendidas com menos de 10s",ehCurta)}>{r.curtas}</span></td>
                  <td style={{padding:"8px"}}>{segParaTexto(r.tempo)}</td>
                  <td style={{padding:"8px"}}>{segParaTexto(r.tempoMedio)}</td>
                  <td style={{padding:"8px"}}>{r.leads}</td>
                  <td style={{padding:"8px"}}><span style={sNum} onClick={()=>recortePessoa(r.u,"com lead no CRM",l=>!!l.item_id)}>{r.comLead}</span><span style={{color:"var(--text3)"}}>/{r.total}</span></td>
                  <td style={{padding:"8px",minWidth:130}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <div style={{flex:1,height:10,background:"var(--surface3)",borderRadius:5,overflow:"hidden"}}><div style={{width:`${Math.min(pct,100)}%`,height:"100%",background:corMeta(pct)}}/></div>
                      <span style={{fontSize:12,fontWeight:900,color:corMeta(pct)}}>{Math.round(pct)}%</span>
                    </div>
                  </td>
                </tr>;
              })}
              {porSdrOrd.length===0&&<tr><td colSpan={12} style={{padding:12,color:"var(--text3)"}}>Nenhum usuário com ramal.</td></tr>}
            </tbody>
          </table>
        </div>

        {/* Gráficos: por dia e por hora (barras clicáveis) */}
        <div style={{display:"grid",gridTemplateColumns:width>=1000?"2fr 1fr":"1fr",gap:14,marginBottom:20}}>
          <div style={sCard}>
            <div style={{fontWeight:800,fontSize:14,marginBottom:4}}>Ligações por dia</div>
            <div style={{fontSize:11,color:"var(--text3)",marginBottom:12}}>barra cheia = total · parte escura = atendidas · linha tracejada = meta diária{sdrSel==="todos"&&sdrs.length>1?" somada dos SDRs":""} · clique numa barra para ver as ligações do dia</div>
            <div style={{position:"relative",height:160,display:"flex",alignItems:"flex-end",gap:diasPeriodo>31?1:4}}>
              {(()=>{const meta=metaDiaria*nSdrMeta;const topo=Math.max(maxDia,meta);const y=160-(meta/topo)*150;return <div style={{position:"absolute",left:0,right:0,top:y,borderTop:"2px dashed #d97706",pointerEvents:"none"}}/>;})()}
              {porDia.map(d=>{const topo=Math.max(maxDia,metaDiaria*nSdrMeta);const fds=[0,6].includes(new Date(d.dia+"T00:00:00").getDay());const rot=d.dia.split("-").reverse().join("/");return(
                <div key={d.dia} onClick={()=>d.total&&abrirRecorte(`Dia ${rot}`,l=>ymdLocal(new Date(l.hora))===d.dia)} title={`${rot}: ${d.total} ligações, ${d.atend} atendidas`} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-end",height:"100%",opacity:fds?.55:1,cursor:d.total?"pointer":"default"}}>
                  {d.total>0&&diasPeriodo<=16&&<div style={{fontSize:10,fontWeight:700,color:"var(--text2)",marginBottom:2}}>{d.total}</div>}
                  <div style={{width:"100%",maxWidth:28,height:`${d.total/topo*135}px`,background:"rgba(99,120,255,.28)",borderRadius:"4px 4px 0 0",position:"relative"}}>
                    <div style={{position:"absolute",bottom:0,left:0,right:0,height:`${d.total?d.atend/d.total*100:0}%`,background:"var(--blue)",borderRadius:"4px 4px 0 0"}}/>
                  </div>
                  {diasPeriodo<=31&&<div style={{fontSize:9,color:"var(--text3)",marginTop:4}}>{d.dia.slice(8)}</div>}
                </div>);})}
            </div>
          </div>
          <div style={sCard}>
            <div style={{fontWeight:800,fontSize:14,marginBottom:4}}>Melhor horário</div>
            <div style={{fontSize:11,color:"var(--text3)",marginBottom:12}}>ligações por hora do dia · parte escura = atendidas · clique para ver</div>
            <div style={{height:160,display:"flex",alignItems:"flex-end",gap:2}}>
              {porHora.map((h,i)=>i>=7&&i<=21&&(
                <div key={i} onClick={()=>h.total&&abrirRecorte(`Horário ${i}h`,l=>new Date(l.hora).getHours()===i)} title={`${i}h: ${h.total} ligações, ${h.atend} atendidas (${h.total?Math.round(h.atend/h.total*100):0}%)`} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-end",height:"100%",cursor:h.total?"pointer":"default"}}>
                  <div style={{width:"100%",height:`${h.total/maxHora*140}px`,background:"rgba(124,58,237,.28)",borderRadius:"3px 3px 0 0",position:"relative"}}>
                    <div style={{position:"absolute",bottom:0,left:0,right:0,height:`${h.total?h.atend/h.total*100:0}%`,background:"#7c3aed",borderRadius:"3px 3px 0 0"}}/>
                  </div>
                  <div style={{fontSize:9,color:"var(--text3)",marginTop:4}}>{i}</div>
                </div>))}
            </div>
          </div>
        </div>

        {/* Lista de ligações (filtros, ordenação, recorte) */}
        <div ref={listaRef} style={{...sCard,overflowX:"auto",scrollMarginTop:8}}>
          <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:10}}>
            <div style={{fontWeight:800,fontSize:14}}>Ligações <span style={{fontWeight:400,color:"var(--text3)",fontSize:12}}>· {listaFinal.length}{listaFinal.length>300?" (mostrando 300)":""}</span></div>
            {recorte&&<span style={{display:"inline-flex",alignItems:"center",gap:6,background:"var(--blue)",color:"#fff",borderRadius:14,padding:"3px 6px 3px 12px",fontSize:12,fontWeight:700}}>
              Mostrando: {recorte.rot}
              <button onClick={()=>setRecorte(null)} title="Voltar a mostrar todas" style={{...T.btn,background:"rgba(255,255,255,.25)",color:"#fff",padding:"0 7px",fontSize:12,borderRadius:10}}>✕</button>
            </span>}
            <div style={{flex:1}}/>
            <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar telefone, lead ou pessoa" style={{...T.inp,width:220,padding:"6px 10px",fontSize:12,borderRadius:8}}/>
            <select value={fStatus} onChange={e=>setFStatus(e.target.value)} style={sSel}>
              <option value="todos">Status: todos</option>{statusDisponiveis.map(s=><option key={s} value={s}>{s}</option>)}
            </select>
            <select value={fDur} onChange={e=>setFDur(e.target.value)} style={sSel}>
              <option value="todos">Duração: todas</option><option value="0">0s</option><option value="lt10">1 a 9s</option><option value="10a59">10 a 59s</option><option value="ge60">1 min ou mais</option>
            </select>
            <select value={fLead} onChange={e=>setFLead(e.target.value)} style={sSel}>
              <option value="todos">Lead: todos</option><option value="com">Com lead no CRM</option><option value="sem">Sem lead no CRM</option>
            </select>
            {temFiltroLista&&<a onClick={limparFiltrosLista} style={{cursor:"pointer",color:"var(--blue)",fontSize:12}}>limpar filtros</a>}
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{color:"var(--text3)",textAlign:"left"}}>
              {[["hora","Quando",-1],["pessoa","Pessoa",1],[null,"Direção"],["telefone","Telefone",1],["lead","Lead no CRM",1],["status","Status",1],["duracao","Duração",-1],[null,"Gravação"]].map(([k,h,d],i)=>(
                <th key={i} onClick={k?()=>ordena(sortL,setSortL,k,d):undefined} style={{...sTh,cursor:k?"pointer":"default"}}>{h}{k?seta(sortL,k):""}</th>))}
            </tr></thead>
            <tbody>
              {listaFinal.slice(0,300).map(l=>(
                <tr key={l.linkedid} style={{borderBottom:"1px solid var(--border)"}}>
                  <td style={{padding:"7px 8px",whiteSpace:"nowrap"}}>{new Date(l.hora).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}</td>
                  <td style={{padding:"7px 8px",whiteSpace:"nowrap"}}>{nomeU(l.profile_id)}</td>
                  <td style={{padding:"7px 8px"}}>{l.direcao==="saida"?"↗ saída":"↙ recebida"}</td>
                  <td style={{padding:"7px 8px",whiteSpace:"nowrap"}}><a onClick={()=>setBusca(String(l.telefone||""))} title="Ver todas as ligações para este número" style={{cursor:"pointer"}}>{fmtFoneBR(l.telefone)}</a></td>
                  <td style={{padding:"7px 8px",maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={l.item_id?leadsNome[l.item_id]:""}>{l.item_id?(leadsNome[l.item_id]||"…"):<span style={{color:"var(--text3)"}}>sem lead</span>}</td>
                  <td style={{padding:"7px 8px",color:corStatus(l.status),fontWeight:700}}>{l.status||"—"}</td>
                  <td style={{padding:"7px 8px"}}>{segParaTexto(l.duracao_seg)}</td>
                  <td style={{padding:"7px 8px"}}>{l.audio?<a href={l.audio} target="_blank" rel="noreferrer" style={{color:"var(--blue)"}}>▶ ouvir</a>:<span style={{color:"var(--text3)"}}>—</span>}</td>
                </tr>
              ))}
              {listaFinal.length===0&&!loading&&<tr><td colSpan={8} style={{padding:12,color:"var(--text3)"}}>{recorte||temFiltroLista?"Nenhuma ligação com este recorte e filtros.":"Nenhuma ligação no período. Se o PABX acabou de ser conectado, clique em Sincronizar."}</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════════════
// METAS E RECOMPENSAS (SDRs e Closers). Regras decididas em 17/11/2026:
//
// SDR, por comparecimento no mês:
//   Comparecimento = lead do quadro Negociações cuja coluna "Etapa" é uma de:
//   Negociando, Proposta Enviada, Negócio Futuro, Descartou, Negócio Fechado.
//   Mês = coluna "Data Reunião" (sem ela, a data em que o lead entrou em Negociações).
//   Dono = quem está na coluna "SDR" do lead.
//   Faixa (vale para TODOS os comparecimentos do mês, não é escalonado):
//     33 ou mais → R$ 35 · 30 a 32 → R$ 27 · 21 a 29 → R$ 23 · até 20 → R$ 20
//
// Closer, por MRR faturado no mês:
//   Conta venda do quadro Vendas com "Faturado" = Sim.
//   MRR = Valor do Projeto ÷ Parcelas. Mês = "Data Entrada" (sem ela, data de criação).
//   Dono = quem está na coluna "Closer" da venda.
//   MRR do mês abaixo de R$ 15.000 → 25% · R$ 15.000 ou mais → 30%.
//   A comissão incide sobre o MRR, uma única vez.
//
// Visibilidade: gestão (quem acessa o Dashboard) vê todos; SDR e Closer veem só
// a própria linha. ATENÇÃO: é filtro de tela; com o RLS desligado, não é barreira de dados.
// ═══════════════════════════════════════════════════════════════════════════
const META_ETAPAS_COMPARECIMENTO=["Negociando","Proposta Enviada","Negócio Futuro","Descartou","Negócio Fechado"];
const META_FAIXAS_SDR=[{min:33,valor:35},{min:30,valor:27},{min:21,valor:23},{min:0,valor:20}]; // do maior para o menor
const META_MRR_CORTE=15000, META_PCT_BAIXO=0.25, META_PCT_ALTO=0.30;
// Pessoas que não aparecem na tela de Metas (pedido da gestão em 18/11/2026): Theo Garcia e Mauricio Alves
const META_OCULTOS=new Set(["1aa307f7-13f2-4b8f-aabe-2f46fab70881","1dc0a183-e111-4183-903f-cb0bb410abea"]);
const faixaSdr=n=>META_FAIXAS_SDR.find(f=>n>=f.min);
const proximaFaixaSdr=n=>[...META_FAIXAS_SDR].reverse().find(f=>f.min>n)||null;

function MetasPage({onBack,allUsers,profile,verTudo}){
  const{isMobile,width}=useBreakpoint();
  const toast=useToast();
  const mesAtual=ymdLocal(new Date()).slice(0,7);
  const [mes,setMes]=useState(mesAtual);
  const [loading,setLoading]=useState(true);
  const [neg,setNeg]=useState([]);   // [{id,nome,etapa,dataRef,mes,temDataReuniao,sdrs:[ids]}]
  const [ven,setVen]=useState([]);   // [{id,nome,faturado,mes,temDataEntrada,valor,parcelas,mrr,plano,closers:[ids]}]
  const [avisos,setAvisos]=useState([]);
  const [aberto,setAberto]=useState(null); // "sdr:<id>" | "closer:<id>"

  const meses=useMemo(()=>{const out=[];const d=new Date();d.setDate(1);for(let i=0;i<12;i++){const v=ymdLocal(d).slice(0,7);out.push({v,l:d.toLocaleDateString("pt-BR",{month:"long",year:"numeric"})});d.setMonth(d.getMonth()-1);}return out;},[]);

  useEffect(()=>{(async()=>{
    setLoading(true);
    try{
      const {data:bds}=await db.from("boards").select("id,nome").in("nome",["Negociações","Vendas"]);
      const bNeg=bds?.find(b=>b.nome==="Negociações"), bVen=bds?.find(b=>b.nome==="Vendas");
      const av=[];
      if(!bNeg) av.push("Quadro Negociações não encontrado.");
      if(!bVen) av.push("Quadro Vendas não encontrado.");
      const ids=[bNeg?.id,bVen?.id].filter(Boolean);
      const {data:cols}=await db.from("columns").select("id,nome,tipo,board_id").in("board_id",ids);
      const col=(b,nome)=>(cols||[]).find(c=>c.board_id===b?.id&&c.nome===nome);
      const precisa=[[bNeg,"Etapa"],[bNeg,"Data Reunião"],[bNeg,"SDR"],[bVen,"Faturado"],[bVen,"Data Entrada"],[bVen,"Valor do Projeto"],[bVen,"Parcelas"],[bVen,"Closer"]];
      precisa.forEach(([b,n])=>{if(b&&!col(b,n))av.push(`Coluna "${n}" não encontrada no quadro ${b.nome}.`);});
      const paginado=async(montar)=>{const PAG=1000,out=[];for(let o=0;;o+=PAG){const {data,error}=await montar().range(o,o+PAG-1);if(error){av.push("Falha ao carregar dados: "+error.message);break;}out.push(...(data||[]));if(!data||data.length<PAG)break;}return out;};
      const itens=await paginado(()=>db.from("items").select("id,board_id,created_at").in("board_id",ids).order("created_at"));
      const colsValor=["Etapa","Data Reunião","Faturado","Data Entrada","Valor do Projeto","Parcelas","Plano","Empresa","Nome","Razão Social"].flatMap(n=>(cols||[]).filter(c=>c.nome===n).map(c=>c.id));
      const vals=colsValor.length?await paginado(()=>db.from("item_values").select("item_id,column_id,value").in("column_id",colsValor).order("item_id")):[];
      const colsResp=[col(bNeg,"SDR")?.id,col(bVen,"Closer")?.id].filter(Boolean);
      const resps=colsResp.length?await paginado(()=>db.from("item_responsables").select("item_id,user_id,column_id").in("column_id",colsResp).order("item_id")):[];
      const nomeCol=Object.fromEntries((cols||[]).map(c=>[c.id,c.nome]));
      const V={}; // item → {nomeColuna: texto}
      vals.forEach(v=>{const raw=(v.value&&typeof v.value==="object"&&"value" in v.value)?v.value.value:v.value;(V[v.item_id]=V[v.item_id]||{})[nomeCol[v.column_id]]=raw==null?"":String(raw).trim();});
      const R={}; resps.forEach(r=>{(R[r.item_id]=R[r.item_id]||[]).push(r.user_id);});
      const rotulo=v=>[v["Empresa"]&&v["Empresa"]!=="Empresa"?v["Empresa"]:v["Razão Social"],v["Nome"]&&v["Nome"]!=="Nome"?v["Nome"]:null].filter(Boolean).join(" · ")||"(sem nome)";
      const ehData=t=>/^\d{4}-\d{2}-\d{2}/.test(t||"");
      const num=t=>{const n=parseFloat(String(t??"").replace(/[^0-9.-]/g,""));return Number.isFinite(n)?n:0;};
      const N=[],Ve=[];
      itens.forEach(i=>{
        const v=V[i.id]||{};
        const criado=ymdLocal(new Date(i.created_at));
        if(i.board_id===bNeg?.id){
          const tem=ehData(v["Data Reunião"]);const dataRef=tem?v["Data Reunião"].slice(0,10):criado;
          N.push({id:i.id,nome:rotulo(v),etapa:v["Etapa"]||"",dataRef,mes:dataRef.slice(0,7),temDataReuniao:tem,sdrs:[...new Set(R[i.id]||[])]});
        }else if(i.board_id===bVen?.id){
          const tem=ehData(v["Data Entrada"]);const dataRef=tem?v["Data Entrada"].slice(0,10):criado;
          const valor=num(v["Valor do Projeto"]);const parcelas=num(v["Parcelas"])>0?num(v["Parcelas"]):1;
          Ve.push({id:i.id,nome:rotulo(v),faturado:v["Faturado"]||"",dataRef,mes:dataRef.slice(0,7),temDataEntrada:tem,valor,parcelas,mrr:valor/parcelas,plano:v["Plano"]||"",closers:[...new Set(R[i.id]||[])]});
        }
      });
      setNeg(N);setVen(Ve);setAvisos(av);
    }catch(e){toast("Não foi possível carregar as metas: "+(e?.message||e),"error");}
    finally{setLoading(false);}
  })();},[]); // eslint-disable-line react-hooks/exhaustive-deps

  const nomeU=id=>allUsers.find(u=>u.id===id)?.nome||"(usuário removido)";
  const papel=id=>allUsers.find(u=>u.id===id)?.roles?.slug||"";
  const meuId=profile?.id;

  // ── SDRs
  const negMes=useMemo(()=>neg.filter(n=>n.mes===mes),[neg,mes]);
  const compMes=useMemo(()=>negMes.filter(n=>META_ETAPAS_COMPARECIMENTO.includes(n.etapa)),[negMes]);
  const sdrLinhas=useMemo(()=>{
    const ids=new Set(allUsers.filter(u=>u.roles?.slug==="sdr").map(u=>u.id));
    compMes.forEach(n=>n.sdrs.forEach(id=>ids.add(id)));
    return [...ids].map(id=>{
      const leads=compMes.filter(n=>n.sdrs.includes(id));
      const emAndamento=negMes.filter(n=>n.sdrs.includes(id)&&!META_ETAPAS_COMPARECIMENTO.includes(n.etapa));
      const f=faixaSdr(leads.length), prox=proximaFaixaSdr(leads.length);
      return {id,nome:nomeU(id),n:leads.length,leads,emAndamento,valorUnit:f.valor,total:leads.length*f.valor,prox,
        ganhoProx:prox?prox.min*prox.valor-leads.length*f.valor:0};
    }).filter(l=>!META_OCULTOS.has(l.id)).filter(l=>verTudo||l.id===meuId).filter(l=>papel(l.id)==="sdr"||l.n>0).sort((a,b)=>b.n-a.n);
  },[compMes,negMes,allUsers,verTudo,meuId]); // eslint-disable-line react-hooks/exhaustive-deps
  const compSemSdr=useMemo(()=>compMes.filter(n=>n.sdrs.length===0),[compMes]);
  const compMultiSdr=useMemo(()=>compMes.filter(n=>n.sdrs.length>1),[compMes]);
  const negSemEtapa=useMemo(()=>negMes.filter(n=>!n.etapa),[negMes]);

  // ── Closers
  const venMes=useMemo(()=>ven.filter(v=>v.mes===mes),[ven,mes]);
  const closerLinhas=useMemo(()=>{
    const ids=new Set(allUsers.filter(u=>u.roles?.slug==="closer").map(u=>u.id));
    venMes.forEach(v=>v.closers.forEach(id=>ids.add(id)));
    return [...ids].map(id=>{
      const minhas=venMes.filter(v=>v.closers.includes(id));
      const fat=minhas.filter(v=>v.faturado==="Sim");
      const naoFat=minhas.filter(v=>v.faturado!=="Sim"&&v.faturado!=="Potencializa Amazon");
      const mrr=fat.reduce((s,v)=>s+v.mrr,0), mrrPend=naoFat.reduce((s,v)=>s+v.mrr,0);
      const pct=mrr>=META_MRR_CORTE?META_PCT_ALTO:META_PCT_BAIXO;
      const mrrProj=mrr+mrrPend, pctProj=mrrProj>=META_MRR_CORTE?META_PCT_ALTO:META_PCT_BAIXO;
      return {id,nome:nomeU(id),fat,naoFat,mrr,mrrPend,pct,comissao:mrr*pct,faltaCorte:Math.max(0,META_MRR_CORTE-mrr),comissaoProj:mrrProj*pctProj,pctProj};
    }).filter(l=>!META_OCULTOS.has(l.id)).filter(l=>verTudo||l.id===meuId).filter(l=>papel(l.id)==="closer"||l.fat.length>0||l.naoFat.length>0).sort((a,b)=>b.mrr-a.mrr);
  },[venMes,allUsers,verTudo,meuId]); // eslint-disable-line react-hooks/exhaustive-deps
  const fatSemCloser=useMemo(()=>venMes.filter(v=>v.faturado==="Sim"&&v.closers.length===0),[venMes]);
  const fatMultiCloser=useMemo(()=>venMes.filter(v=>v.faturado==="Sim"&&v.closers.length>1),[venMes]);

  const totalSdr=sdrLinhas.reduce((s,l)=>s+l.total,0), totalCloser=closerLinhas.reduce((s,l)=>s+l.comissao,0);
  const dataBR=d=>d?d.split("-").reverse().join("/"):"";
  const sCard={background:"var(--surface)",border:"1px solid var(--border)",borderRadius:14,padding:16};
  const sTh={padding:"6px 8px",borderBottom:"1px solid var(--border)",whiteSpace:"nowrap",textAlign:"left",color:"var(--text3)",fontWeight:600};
  const sTd={padding:"8px"};
  const corFaixa=v=>v>=35?"#059669":v>=27?"#0891b2":v>=23?"#d97706":"var(--text2)";

  return(
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",background:"var(--bg)"}}>
      <div style={{background:"var(--surface)",borderBottom:"1px solid var(--border)",padding:isMobile?"12px 16px":"14px 24px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",flexShrink:0}}>
        <button onClick={onBack} style={{...T.btn,background:"none",border:"1.5px solid var(--border)",color:"var(--text)",padding:"7px 14px",fontSize:12}}>← Voltar</button>
        <div style={{fontSize:isMobile?16:18,fontWeight:900,color:"var(--text)"}}>🏆 Metas e Recompensas</div>
        {loading&&<Spinner size={15}/>}
        <select value={mes} onChange={e=>{setMes(e.target.value);setAberto(null);}} style={{...T.inp,width:"auto",padding:"7px 10px",fontSize:12,borderRadius:8,textTransform:"capitalize"}}>
          {meses.map(m=><option key={m.v} value={m.v}>{m.l}</option>)}
        </select>
        {!verTudo&&<span style={{fontSize:11,color:"var(--text3)"}}>Você vê apenas os seus números.</span>}
      </div>

      <div style={{flex:1,overflowY:"auto",padding:isMobile?"12px":"20px 24px"}}>
        {avisos.map((a,i)=><div key={i} style={{background:"#fef3c7",color:"#92400e",border:"1px solid #fcd34d",borderRadius:12,padding:"10px 14px",marginBottom:10,fontSize:13}}>⚠ {a}</div>)}

        {verTudo&&<div style={{display:"grid",gridTemplateColumns:width>=900?"repeat(3,1fr)":"1fr",gap:14,marginBottom:20}}>
          <div style={sCard}><div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase"}}>Recompensa dos SDRs no mês</div><div style={{fontSize:26,fontWeight:900,marginTop:6}}>{fmtBRL(totalSdr)}</div><div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>{compMes.length} comparecimento(s) no mês</div></div>
          <div style={sCard}><div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase"}}>Comissão dos Closers no mês</div><div style={{fontSize:26,fontWeight:900,marginTop:6}}>{fmtBRL(totalCloser)}</div><div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>sobre {fmtBRL(closerLinhas.reduce((s,l)=>s+l.mrr,0))} de MRR faturado</div></div>
          <div style={sCard}><div style={{fontSize:11,fontWeight:700,color:"var(--text3)",textTransform:"uppercase"}}>Total a pagar</div><div style={{fontSize:26,fontWeight:900,marginTop:6,color:"#059669"}}>{fmtBRL(totalSdr+totalCloser)}</div><div style={{fontSize:11,color:"var(--text3)",marginTop:4}}>SDRs + Closers, conforme os dados de hoje</div></div>
        </div>}

        {/* ── SDRs ── */}
        {(verTudo||sdrLinhas.length>0)&&<div style={{...sCard,marginBottom:20,overflowX:"auto"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:4}}>SDRs · comparecimentos</div>
          <div style={{fontSize:11,color:"var(--text3)",marginBottom:12}}>
            Conta lead de Negociações com Etapa em {META_ETAPAS_COMPARECIMENTO.join(", ")}, no mês da Data Reunião.
            Faixas: até 20 → R$ 20 · 21 a 29 → R$ 23 · 30 a 32 → R$ 27 · 33 ou mais → R$ 35, valendo para todos os comparecimentos do mês.
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr>{["SDR","Comparecimentos","Faixa","Recompensa","Próxima faixa","Ainda sem comparecimento"].map(h=><th key={h} style={sTh}>{h}</th>)}</tr></thead>
            <tbody>
              {sdrLinhas.map(l=>{const k="sdr:"+l.id;const topo=33;return(
                <Fragment key={l.id}>
                  <tr onClick={()=>setAberto(aberto===k?null:k)} style={{cursor:"pointer",background:aberto===k?"var(--surface3)":"transparent"}}>
                    <td style={{...sTd,fontWeight:800,whiteSpace:"nowrap"}}>{aberto===k?"▼":"▶"} {l.nome}</td>
                    <td style={{...sTd,minWidth:220}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontWeight:900,fontSize:20,minWidth:30}}>{l.n}</span>
                        <div style={{flex:1,height:12,background:"var(--surface3)",borderRadius:6,position:"relative",overflow:"hidden"}}>
                          <div style={{width:`${Math.min(l.n/topo*100,100)}%`,height:"100%",background:corFaixa(l.valorUnit)==="var(--text2)"?"var(--blue)":corFaixa(l.valorUnit)}}/>
                          {[21,30,33].map(m=><div key={m} title={`${m} comparecimentos`} style={{position:"absolute",top:0,bottom:0,left:`${m/topo*100}%`,width:2,background:"var(--bg)",marginLeft:-1}}/>)}
                        </div>
                      </div>
                    </td>
                    <td style={{...sTd,fontWeight:800,color:corFaixa(l.valorUnit),whiteSpace:"nowrap"}}>{fmtBRL(l.valorUnit)} cada</td>
                    <td style={{...sTd,fontWeight:900,fontSize:16,whiteSpace:"nowrap"}}>{fmtBRL(l.total)}</td>
                    <td style={{...sTd,fontSize:12,color:"var(--text2)"}}>{l.prox?<span>faltam <b>{l.prox.min-l.n}</b> para {fmtBRL(l.prox.valor)} cada <span style={{color:"#059669"}}>(+{fmtBRL(l.ganhoProx)})</span></span>:<b style={{color:"#059669"}}>faixa máxima</b>}</td>
                    <td style={{...sTd,fontSize:12,color:"var(--text3)"}}>{l.emAndamento.length}</td>
                  </tr>
                  {aberto===k&&<tr><td colSpan={6} style={{padding:"4px 8px 14px 28px"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr>{["Lead","Etapa","Data","Conta?"].map(h=><th key={h} style={sTh}>{h}</th>)}</tr></thead>
                      <tbody>
                        {[...l.leads.map(n=>({...n,conta:true})),...l.emAndamento.map(n=>({...n,conta:false}))].sort((a,b)=>a.dataRef.localeCompare(b.dataRef)).map(n=>(
                          <tr key={n.id} style={{borderBottom:"1px solid var(--border)",opacity:n.conta?1:.6}}>
                            <td style={{padding:"6px 8px"}}>{n.nome}</td>
                            <td style={{padding:"6px 8px"}}>{n.etapa||<span style={{color:"#d97706"}}>(Etapa vazia)</span>}</td>
                            <td style={{padding:"6px 8px",whiteSpace:"nowrap"}}>{dataBR(n.dataRef)}{!n.temDataReuniao&&<span style={{color:"#d97706"}} title="Sem Data Reunião: usada a data de entrada em Negociações"> *</span>}</td>
                            <td style={{padding:"6px 8px",fontWeight:700,color:n.conta?"#059669":"var(--text3)"}}>{n.conta?"sim":"não"}</td>
                          </tr>))}
                        {l.leads.length+l.emAndamento.length===0&&<tr><td colSpan={4} style={{padding:8,color:"var(--text3)"}}>Nenhum lead deste SDR em Negociações neste mês.</td></tr>}
                      </tbody>
                    </table>
                  </td></tr>}
                </Fragment>);})}
              {sdrLinhas.length===0&&!loading&&<tr><td colSpan={6} style={{padding:12,color:"var(--text3)"}}>Nenhum SDR com dados neste mês.</td></tr>}
            </tbody>
          </table>
          {verTudo&&(compSemSdr.length>0||compMultiSdr.length>0||negSemEtapa.length>0)&&<div style={{fontSize:12,color:"#d97706",marginTop:12,lineHeight:1.6}}>
            {compSemSdr.length>0&&<div>⚠ {compSemSdr.length} comparecimento(s) do mês sem ninguém na coluna SDR; não entram para nenhum SDR: {compSemSdr.slice(0,6).map(n=>n.nome).join("; ")}{compSemSdr.length>6?"…":""}</div>}
            {compMultiSdr.length>0&&<div>⚠ {compMultiSdr.length} comparecimento(s) com mais de um SDR; contam inteiro para cada um.</div>}
            {negSemEtapa.length>0&&<div>⚠ {negSemEtapa.length} lead(s) de Negociações do mês com a Etapa vazia; não contam até a Etapa ser preenchida.</div>}
          </div>}
        </div>}

        {/* ── Closers ── */}
        {(verTudo||closerLinhas.length>0)&&<div style={{...sCard,overflowX:"auto"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:4}}>Closers · comissão sobre MRR</div>
          <div style={{fontSize:11,color:"var(--text3)",marginBottom:12}}>
            Conta venda com Faturado = Sim, no mês da Data Entrada. MRR = Valor do Projeto ÷ Parcelas.
            MRR do mês abaixo de {fmtBRL(META_MRR_CORTE)} → 25% · a partir de {fmtBRL(META_MRR_CORTE)} → 30%. Comissão sobre o MRR, paga uma vez.
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr>{["Closer","MRR faturado","Faixa","Comissão","Para os 30%","Ainda não faturado"].map(h=><th key={h} style={sTh}>{h}</th>)}</tr></thead>
            <tbody>
              {closerLinhas.map(l=>{const k="closer:"+l.id;return(
                <Fragment key={l.id}>
                  <tr onClick={()=>setAberto(aberto===k?null:k)} style={{cursor:"pointer",background:aberto===k?"var(--surface3)":"transparent"}}>
                    <td style={{...sTd,fontWeight:800,whiteSpace:"nowrap"}}>{aberto===k?"▼":"▶"} {l.nome}</td>
                    <td style={{...sTd,minWidth:240}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontWeight:900,fontSize:16,minWidth:110}}>{fmtBRL(l.mrr)}</span>
                        <div style={{flex:1,height:12,background:"var(--surface3)",borderRadius:6,overflow:"hidden"}}><div style={{width:`${Math.min(l.mrr/META_MRR_CORTE*100,100)}%`,height:"100%",background:l.mrr>=META_MRR_CORTE?"#059669":"var(--blue)"}}/></div>
                      </div>
                      <div style={{fontSize:11,color:"var(--text3)"}}>{l.fat.length} venda(s) faturada(s)</div>
                    </td>
                    <td style={{...sTd,fontWeight:800,color:l.pct===META_PCT_ALTO?"#059669":"var(--text2)"}}>{Math.round(l.pct*100)}%</td>
                    <td style={{...sTd,fontWeight:900,fontSize:16,whiteSpace:"nowrap"}}>{fmtBRL(l.comissao)}</td>
                    <td style={{...sTd,fontSize:12,color:"var(--text2)"}}>{l.faltaCorte>0?<span>faltam <b>{fmtBRL(l.faltaCorte)}</b> de MRR</span>:<b style={{color:"#059669"}}>faixa máxima</b>}</td>
                    <td style={{...sTd,fontSize:12,color:"var(--text3)"}}>{l.naoFat.length?<span>{l.naoFat.length} venda(s), {fmtBRL(l.mrrPend)} de MRR · se faturar tudo: <b style={{color:"var(--text2)"}}>{fmtBRL(l.comissaoProj)}</b> ({Math.round(l.pctProj*100)}%)</span>:"—"}</td>
                  </tr>
                  {aberto===k&&<tr><td colSpan={6} style={{padding:"4px 8px 14px 28px"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr>{["Venda","Plano","Data Entrada","Valor do Projeto","Parcelas","MRR","Faturado","Conta?"].map(h=><th key={h} style={sTh}>{h}</th>)}</tr></thead>
                      <tbody>
                        {[...l.fat.map(v=>({...v,conta:true})),...l.naoFat.map(v=>({...v,conta:false}))].sort((a,b)=>a.dataRef.localeCompare(b.dataRef)).map(v=>(
                          <tr key={v.id} style={{borderBottom:"1px solid var(--border)",opacity:v.conta?1:.6}}>
                            <td style={{padding:"6px 8px"}}>{v.nome}</td>
                            <td style={{padding:"6px 8px"}}>{v.plano||"—"}</td>
                            <td style={{padding:"6px 8px",whiteSpace:"nowrap"}}>{dataBR(v.dataRef)}{!v.temDataEntrada&&<span style={{color:"#d97706"}} title="Sem Data Entrada: usada a data de criação"> *</span>}</td>
                            <td style={{padding:"6px 8px",whiteSpace:"nowrap"}}>{fmtBRL(v.valor)}</td>
                            <td style={{padding:"6px 8px"}}>{v.parcelas}</td>
                            <td style={{padding:"6px 8px",whiteSpace:"nowrap",fontWeight:700}}>{fmtBRL(v.mrr)}</td>
                            <td style={{padding:"6px 8px"}}>{v.faturado||<span style={{color:"#d97706"}}>(vazio)</span>}</td>
                            <td style={{padding:"6px 8px",fontWeight:700,color:v.conta?"#059669":"var(--text3)"}}>{v.conta?"sim":"não"}</td>
                          </tr>))}
                        {l.fat.length+l.naoFat.length===0&&<tr><td colSpan={8} style={{padding:8,color:"var(--text3)"}}>Nenhuma venda deste closer neste mês.</td></tr>}
                      </tbody>
                    </table>
                  </td></tr>}
                </Fragment>);})}
              {closerLinhas.length===0&&!loading&&<tr><td colSpan={6} style={{padding:12,color:"var(--text3)"}}>Nenhum closer com dados neste mês.</td></tr>}
            </tbody>
          </table>
          {verTudo&&(fatSemCloser.length>0||fatMultiCloser.length>0)&&<div style={{fontSize:12,color:"#d97706",marginTop:12,lineHeight:1.6}}>
            {fatSemCloser.length>0&&<div>⚠ {fatSemCloser.length} venda(s) faturada(s) do mês sem ninguém na coluna Closer ({fmtBRL(fatSemCloser.reduce((s,v)=>s+v.mrr,0))} de MRR); não entram para nenhum closer.</div>}
            {fatMultiCloser.length>0&&<div>⚠ {fatMultiCloser.length} venda(s) faturada(s) com mais de um closer; o MRR conta inteiro para cada um.</div>}
          </div>}
        </div>}
      </div>
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
  const[groups,setGroups]=useState([]);
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
      const{data:allGroups}=await db.from("groups").select("id,board_id,nome,ordem,parent_group_id,is_parent").in("board_id",boardIds);
      setGroups(allGroups||[]);

      // Busca TODOS os itens dos 3 boards. O recorte por período é feito no cliente
      // (inPeriod), ancorado na coluna "Data Entrada" de cada board — e NÃO no created_at,
      // que reflete apenas quando a linha foi criada no quadro. Isto evita que um lead
      // com Data Entrada no período, mas criado fora dele, fique de fora da busca.
      // group_id é essencial: em Negociações o ESTADO do lead é o grupo (etapa), não a coluna Etapa.
      // PAGINADO: o Supabase corta em 1000 linhas e ignora .limit() maior. Os 3
      // quadros somavam 553 itens em nov/2026; sem paginação, ao passar de 1000 o
      // dashboard perderia itens em silêncio (e as consultas antigas por quadro,
      // que eram subconjuntos desta, tinham o mesmo teto).
      const PAG_ITENS=1000;
      const items=[];
      for(let de=0;;de+=PAG_ITENS){
        const{data,error}=await db.from("items").select("id,board_id,group_id,created_at")
          .in("board_id",boardIds).order("created_at",{ascending:true}).range(de,de+PAG_ITENS-1);
        if(error){console.error("Dashboard: falha ao carregar items",error);break;}
        items.push(...(data||[]));
        if(!data||data.length<PAG_ITENS) break;
      }

      // ── DEDUPLICAR por id antes de salvar ────────────────────────────────
      const uniqMap=new Map();
      for(const it of items) uniqMap.set(it.id,it);
      const uniqItems=[...uniqMap.values()];
      const uniqIds=uniqItems.map(x=>x.id);
      setRawItems(uniqItems);

      if(uniqIds.length){
        // ATENÇÃO: o Supabase corta a resposta em 1000 linhas e IGNORA .limit() maior.
        // Um bloco de 100 itens já chegava a 960 linhas: a 40 do teto. Por isso é
        // preciso PAGINAR com range() dentro de cada bloco, até vir página incompleta.
        const PAGINA=1000, CHUNK=50;
        const chunks=[];
        for(let i=0;i<uniqIds.length;i+=CHUNK) chunks.push(uniqIds.slice(i,i+CHUNK));
        const buscaPaginado=async(chunk)=>{
          const out=[];
          for(let de=0;;de+=PAGINA){
            const{data,error}=await db.from("item_values").select("item_id,column_id,value").in("item_id",chunk).range(de,de+PAGINA-1);
            if(error){console.error("Dashboard: falha ao carregar item_values",error);break;}
            out.push(...(data||[]));
            if(!data||data.length<PAGINA) break;
          }
          return out;
        };
        const valsArr=await Promise.all(chunks.map(buscaPaginado));
        // Deduplicar item_values por (item_id, column_id) — pega o último
        const vMap=new Map();
        for(const v of valsArr.flat()) vMap.set(v.item_id+"_"+v.column_id,v);
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

    // Índice de valores: {itemId: {colId: value}}
    const valIdx={};
    for(const row of rawVals){
      if(!valIdx[row.item_id]) valIdx[row.item_id]={};
      const v=row.value;
      valIdx[row.item_id][row.column_id]=(v!==null&&v!==undefined&&typeof v==="object"&&"value" in v)?v.value:v;
    }
    const gV=(itemId,colId)=>colId?(valIdx[itemId]?.[colId]??null):null;
    const itemsComValor=new Set(rawVals.filter(r=>r.value!==null&&r.value!==undefined&&r.value!=="").map(r=>r.item_id));
    const isPreenchido=id=>itemsComValor.has(id);

    const venB=gB("Vendas"),negB=gB("Negociações"),preB=gB("Pré - Vendas");
    const venValC=gC("Vendas","Valor do Projeto","currency");
    const venParC=gC("Vendas","Parcelas","number");
    const venPlanoC=gC("Vendas","Plano","status");
    const negValC=gC("Negociações","Valor do Projeto","currency");
    const preMqlC=gC("Pré - Vendas","MQL","status");
    const preOriC=gC("Pré - Vendas","Origem","status");

    // ── ESTADO em Negociações = GRUPO (etapa), nunca a coluna "Etapa".
    // A coluna não acompanha o arrasto do lead e ficava travada em "Negociando" para
    // leads já perdidos ou ganhos (R$ 611 mil de pipeline fantasma em out/2026).
    const grpById={};
    for(const g of groups) grpById[g.id]=g;
    const etapaDoItem=it=>{
      const g=grpById[it.group_id];
      if(!g) return null;
      if(g.is_parent) return "(sem etapa)"; // lead preso no grupo-mãe, fora de qualquer etapa
      return g.nome;
    };
    // Agrupa as etapas padronizadas (mesmo nome em todos os closers) pela ordem de exibição
    const etapaOrdem={};
    for(const g of groups){ if(g.board_id===negB?.id&&g.parent_group_id&&etapaOrdem[g.nome]==null) etapaOrdem[g.nome]=g.ordem??99; }

    // ── Âncora de período: "Data Entrada" (Vendas/Pré-Vendas); Negociações usa created_at.
    const venDeC=gC("Vendas","Data Entrada","date");
    const preDeC=gC("Pré - Vendas","Data Entrada","date");
    const deColByBoard={};
    if(venB&&venDeC) deColByBoard[venB.id]=venDeC.id;
    if(preB&&preDeC) deColByBoard[preB.id]=preDeC.id;
    const temDataEntrada=it=>{const cid=deColByBoard[it.board_id];const raw=cid?gV(it.id,cid):null;return typeof raw==="string"&&/^\d{4}-\d{2}-\d{2}/.test(raw);};
    const anchorDate=it=>{
      const cid=deColByBoard[it.board_id];
      const raw=cid?gV(it.id,cid):null;
      if(typeof raw==="string"&&/^\d{4}-\d{2}-\d{2}/.test(raw)){
        const[y,m,dd]=raw.slice(0,10).split("-").map(Number);
        return new Date(y,m-1,dd);
      }
      return new Date(it.created_at); // fallback: sem Data Entrada, usa a criação (tornado visível abaixo)
    };
    const inPeriod=it=>{
      if(filterMode==="all") return true;
      if(filterMode==="month"&&selMonth){
        const[y,m]=selMonth.split("-");
        const s=new Date(parseInt(y),parseInt(m)-1,1);
        const e=new Date(parseInt(y),parseInt(m),1);
        const d=anchorDate(it);return d>=s&&d<e;
      }
      if(filterMode==="range"&&rangeStart&&rangeEnd){
        const s=new Date(rangeStart);const e=new Date(rangeEnd+"T23:59:59Z");
        const d=anchorDate(it);return d>=s&&d<=e;
      }
      return false;
    };

    // ── Vendas: separa venda de verdade de registro que não é venda
    const ehVenda=it=>!PLANOS_NAO_VENDA.includes(String(gV(it.id,venPlanoC?.id)||"").trim());
    const venTodas   = rawItems.filter(i=>i.board_id===venB?.id && isPreenchido(i.id));
    const venFiltAll = venTodas.filter(inPeriod);
    const venFilt    = venFiltAll.filter(ehVenda);                 // vendas reais no período
    const venNaoVenda= venFiltAll.filter(i=>!ehVenda(i));           // ex.: Potencializa Amazon
    const venSemData = venFilt.filter(i=>!temDataEntrada(i));       // contadas por fallback
    const venAll8m   = venTodas.filter(ehVenda);

    const negAll  = rawItems.filter(i=>i.board_id===negB?.id && isPreenchido(i.id));
    const negFilt = negAll.filter(inPeriod);
    const preAll  = rawItems.filter(i=>i.board_id===preB?.id && isPreenchido(i.id));
    const preFilt = preAll.filter(inPeriod);

    // ── Financeiro (Vendas reais)
    const receitaEfetiva=venFilt.reduce((s,it)=>s+(parseFloat(gV(it.id,venValC?.id))||0),0);
    const mrr=venFilt.reduce((s,it)=>{
      const v=parseFloat(gV(it.id,venValC?.id))||0;
      const p=parseFloat(gV(it.id,venParC?.id))||1;
      return s+(p>0?v/p:0);
    },0);
    const venComValor=venFilt.filter(it=>(parseFloat(gV(it.id,venValC?.id))||0)>0);
    const ticketMedio=venComValor.length>0?receitaEfetiva/venComValor.length:0;

    // ── Proposta na Rua: leads NO GRUPO "Negociando" (estado atual, sem período)
    const emNegociacao=negAll.filter(it=>etapaDoItem(it)==="Negociando");
    const propostaNaRua=emNegociacao.reduce((s,it)=>s+(parseFloat(gV(it.id,negValC?.id))||0),0);

    // ── Negócios Fechados: o registro de fechamento é a venda real no período
    const negFechados=venFilt.length;

    // ── Etapas (Negociações entradas no período), por GRUPO
    const etapasMap={};
    negFilt.forEach(it=>{const e=etapaDoItem(it);if(e){etapasMap[e]=(etapasMap[e]||0)+1;}});
    const corEtapa={"Oportunidades Closer":"#3145FF","Negociando":"#d97706","Fechado/Ganho":"#059669","Não comparecido":"#7c3aed","Recusado/Negócio Futuro":"#dc2626","(sem etapa)":"#94a3b8"};
    const etapasData=Object.entries(etapasMap)
      .map(([label,value])=>({label,value,color:corEtapa[label]||"#7c3aed",ordem:etapaOrdem[label]??99}))
      .sort((a,b)=>a.ordem-b.ordem);

    // ── Origens de LEAD: Pré-Vendas no período (onde a origem é capturada)
    const oriOpts=preOriC?.config?.options||[];
    const oriMap={};
    preFilt.filter(it=>!!gV(it.id,preOriC?.id)).forEach(it=>{const o=gV(it.id,preOriC?.id);oriMap[o]=(oriMap[o]||0)+1;});
    const totalOri=Object.values(oriMap).reduce((s,v)=>s+v,0)||1;
    const origemData=Object.entries(oriMap).map(([label,value])=>({
      label,value,pct:Math.round((value/totalOri)*100),
      color:oriOpts.find(o=>o.label===label)?.color||"#94a3b8"
    })).sort((a,b)=>b.value-a.value);

    // ── MQL: Pré-Vendas NO PERÍODO (antes ignorava o filtro)
    let mqlQ=0,mqlDQ=0;
    preFilt.forEach(it=>{
      const m=gV(it.id,preMqlC?.id);
      if(m==="Qualificado")mqlQ++;
      else if(m==="Desqualificado"||m==="Não Qualificado")mqlDQ++;
    });

    // ── Receita por mês (últimos 8 meses, vendas reais)
    const mesMap={};
    const now=new Date();
    for(let i=7;i>=0;i--){
      const d=new Date(now.getFullYear(),now.getMonth()-i,1);
      mesMap[`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`]=0;
    }
    venAll8m.forEach(it=>{
      const d=anchorDate(it);
      const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      if(k in mesMap){mesMap[k]+=(parseFloat(gV(it.id,venValC?.id))||0);}
    });
    const receitaMesData=Object.entries(mesMap).map(([k,value])=>({label:k.slice(5),fullLabel:k,value,highlight:k===selMonth}));

    // ── Taxa de conversão: vendas reais ÷ leads de Pré-Vendas, ambos no período
    const totalLeads=preFilt.length;
    const taxaConversao=totalLeads>0?(venFilt.length/totalLeads*100):0;

    return{
      receitaEfetiva,mrr,ticketMedio,
      negFechados,
      propostaNaRua,emNegociacaoQtd:emNegociacao.length,
      etapasData,origemData,
      mqlQ,mqlDQ,receitaMesData,
      totalLeads,taxaConversao,
      venFiltLen:venComValor.length,
      venNaoVendaLen:venNaoVenda.length,
      venSemDataLen:venSemData.length,
      negFiltLen:negFilt.length,
      preFiltLen:preFilt.length
    };
  },[rawItems,rawVals,boards,cols,groups,filterMode,selMonth,rangeStart,rangeEnd]);

  const monthOptions=useMemo(()=>{
    const opts=[];const now=new Date();
    for(let i=11;i>=0;i--){const d=new Date(now.getFullYear(),now.getMonth()-i,1);const val=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;const label=d.toLocaleDateString("pt-BR",{month:"short",year:"numeric"}).replace("de ","");opts.push({val,label});}
    return opts;
  },[]);

  const canEditGoal=perms?.all||perms?.manageBoards||perms?.editAny||perms?.isFull;
  // A meta é MENSAL. Em "Período" e "Todos" a receita não é comparável com ela
  // (em "Todos" aparecia 554% de atingimento), então o bloco de meta só calcula no modo Mês.
  const metaAplicavel=filterMode==="month";
  const goalAtivo=metaAplicavel?goal:0;
  const progresso=goalAtivo>0?(metrics.receitaEfetiva||0)/goalAtivo*100:0;
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
              <div style={{fontSize:11,color:"var(--text3)",marginTop:2}}>{filterMode==="month"?monthOptions.find(m=>m.val===selMonth)?.label||selMonth:"Período selecionado · a meta é mensal e só é comparada no modo Mês"}</div>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
              {canEditGoal&&metaAplicavel&&!editGoal&&<button onClick={()=>{setEditGoal(true);setGoalInput(String(goal));}} style={{...T.btn,background:"var(--surface3)",color:"var(--text2)",padding:"7px 14px",fontSize:12,border:"1px solid var(--border)"}}>{goal>0?"✏️ Editar meta":"+ Definir meta"}</button>}
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
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:goalAtivo>0?16:0}}>
            {[
              {lbl:"Receita Efetiva",val:fmtBRL(metrics.receitaEfetiva||0),clr:"#059669"},
              {lbl:"Meta",val:goalAtivo>0?fmtBRL(goalAtivo):"—",clr:"var(--text)"},
              {lbl:"Restante",val:goalAtivo>0?fmtBRL(Math.max(goalAtivo-(metrics.receitaEfetiva||0),0)):"—",clr:goalAtivo>0&&(metrics.receitaEfetiva||0)>=goalAtivo?"#059669":"#dc2626"},
              {lbl:"Atingimento",val:goalAtivo>0?`${Math.round(progresso)}%`:"—",clr:progresso>=100?"#059669":progresso>=70?"#d97706":"#dc2626"},
            ].map((item,i)=>(
              <div key={i}>
                <div style={{fontSize:10,fontWeight:700,color:"var(--text3)",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>{item.lbl}</div>
                <div style={{fontSize:isMobile?18:24,fontWeight:900,color:item.clr}}>{item.val}</div>
              </div>
            ))}
          </div>
          {(metrics.venSemDataLen||0)>0&&(
            <div style={{fontSize:11,color:"#d97706",marginBottom:goalAtivo>0?10:0}}>
              ⚠ {metrics.venSemDataLen} venda(s) deste período sem "Data Entrada": contadas pela data de criação. Preencha a coluna no quadro Vendas para o número refletir a data real.
            </div>
          )}
          {goalAtivo>0&&<div>
            <div style={{height:10,background:"var(--surface3)",borderRadius:5,overflow:"hidden",marginBottom:6}}>
              <div style={{height:"100%",width:`${Math.min(progresso,100)}%`,background:progresso>=100?"#059669":progresso>=70?"#d97706":"var(--blue)",borderRadius:5,transition:"width .6s ease"}}/>
            </div>
            <div style={{fontSize:11,color:"var(--text3)"}}>{progresso>=100?"🎉 Meta atingida!":progresso>=70?`Quase lá, faltam ${fmtBRL(goalAtivo-(metrics.receitaEfetiva||0))}`:`${Math.round(100-progresso)}% abaixo da meta`}</div>
          </div>}
        </div>

        {/* KPIs */}
        <div style={{display:"grid",gridTemplateColumns:gridCols,gap:12,marginBottom:16}}>
          <KpiCard title="Receita Mensal das Vendas" value={fmtBRL(metrics.mrr||0)} icon="🔄" color="#059669" sub="Soma de valor ÷ parcelas das vendas do período"/>
          <KpiCard title="Proposta na Rua" value={fmtBRL(metrics.propostaNaRua||0)} icon="📤" color="#d97706" sub={`${metrics.emNegociacaoQtd||0} leads no grupo Negociando (hoje, sem período)`} badge={{label:"Ativo",color:"#d97706",bg:"#d9770620"}}/>
          <KpiCard title="Ticket Médio" value={fmtBRL(metrics.ticketMedio||0)} icon="💰" color="#7c3aed" sub={`${metrics.venFiltLen||0} vendas com valor no período`}/>
          <KpiCard title="Negócios Fechados" value={metrics.negFechados||0} icon="🏆" color="#0891b2" sub={`Vendas no período${metrics.venNaoVendaLen?` · +${metrics.venNaoVendaLen} Potencializa (não contam)`:""}`}/>
          <KpiCard title="Total Leads" value={metrics.totalLeads||0} icon="👥" color="#3145FF" sub="Pré-Vendas no período"/>
          <KpiCard title="Taxa de Conversão" value={`${(metrics.taxaConversao||0).toFixed(1)}%`} icon="📈" color="#059669" sub="Vendas ÷ leads, ambos no período"/>
          <KpiCard title="MQL Qualificados" value={metrics.mqlQ||0} icon="✅" color="#059669" sub={`De ${(metrics.mqlQ||0)+(metrics.mqlDQ||0)} avaliados no período`} progress={(metrics.mqlQ||0)+(metrics.mqlDQ||0)>0?((metrics.mqlQ||0)/((metrics.mqlQ||0)+(metrics.mqlDQ||0)))*100:null}/>
          <KpiCard title="Em Negociação" value={metrics.negFiltLen||0} icon="🤝" color="#7c3aed" sub="Leads que entraram em Negociações no período"/>
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

          <ChartCard title="🌐 Origens de Lead" subtitle="Pré-Vendas no período — onde o lead entrou">
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

          <ChartCard title="🔄 Etapas — Negociações" subtitle="Grupo atual dos leads que entraram no período">
            {(metrics.etapasData||[]).length>0?<HBarSVG items={metrics.etapasData||[]}/>:<div style={{color:"var(--text3)",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sem dados de etapas</div>}
          </ChartCard>

        </div>

        {/* Gráficos linha 2 */}
        <div style={{display:"grid",gridTemplateColumns:width>=1024?"repeat(2,1fr)":"1fr",gap:16,marginBottom:16}}>

          <ChartCard title="🎯 MQL — Qualificação (Pré-Vendas)" subtitle="Leads do período">
            <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
              <DonutSVG segments={[
                {label:"Qualificado",value:metrics.mqlQ||0,color:"#059669"},
                {label:"Desqualificado",value:metrics.mqlDQ||0,color:"#dc2626"},
                {label:"Sem avaliação",value:Math.max((metrics.preFiltLen||0)-(metrics.mqlQ||0)-(metrics.mqlDQ||0),0),color:"var(--border)"}
              ]} size={100}/>
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:10,minWidth:0}}>
                {[{label:"Qualificados",value:metrics.mqlQ||0,color:"#059669"},{label:"Desqualificados",value:metrics.mqlDQ||0,color:"#dc2626"},{label:"Sem avaliação",value:Math.max((metrics.preFiltLen||0)-(metrics.mqlQ||0)-(metrics.mqlDQ||0),0),color:"var(--border)"}].map((row,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:row.color,flexShrink:0}}/>
                    <div style={{flex:1,fontSize:12,color:"var(--text2)"}}>{row.label}</div>
                    <div style={{fontSize:14,fontWeight:800,color:"var(--text)"}}>{row.value}</div>
                  </div>
                ))}
              </div>
            </div>
          </ChartCard>

          <ChartCard title="🔽 Funil de Conversão" subtitle="Leads → Negociações → Vendas, todos no período">
            {(()=>{
              // Usa exatamente as mesmas contagens dos cards, para o funil nunca divergir deles.
              const steps=[
                {label:"Pré-Vendas (Leads)",value:metrics.totalLeads||0,color:"#3145FF"},
                {label:"Entraram em Negociações",value:metrics.negFiltLen||0,color:"#7c3aed"},
                {label:"Vendas",value:metrics.negFechados||0,color:"#059669"},
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
  const {session,profile,authLoading,signOut,isRecovery}=useAuth();
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
  const [sidebarHover,setSidebarHover]=useState(false);
  const [confirmM,setConfirmM]=useState(null);
  const [boardFormM,setBoardFormM]=useState(null);
  const [boardSettingsM,setBoardSettingsM]=useState(null); // board obj
  const [boardAccess,setBoardAccess]=useState({}); // {boardId:[userId,...]}

  // Permissões derivadas do cargo do perfil logado
  const perms=useMemo(()=>getPerms(profile),[profile]);
  // Dashboard: fora do alcance de Closer e SDR (decisão de 11/11). Vale para o
  // link da barra lateral (cheia e recolhida) e para a própria página.
  const canDashboard=!["closer","sdr"].includes(perms.slug);

  // Boards visíveis conforme privacidade
  const visibleBoards=useMemo(()=>{
    if(perms.all||perms.viewAll) return boards;
    // SDR não vê o board de Negociações
    if(perms.slug==="sdr") return boards.filter(b=>b.nome!=="Negociações");
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
    // Modal de confirmação com digitação obrigatória da frase
    const FRASE="Eu realmente quero deletar esse quadro inteiro!";
    setConfirmM({
      title:"⚠️ Excluir quadro permanentemente",
      danger:true,
      requirePhrase:FRASE,
      boardName:b.nome,
      message:"Esta ação é irreversível. Todos os leads, campos e anotações do quadro "+b.nome+" serão apagados para sempre.",
      onConfirm:async()=>{
        await db.from("boards").delete().eq("id",b.id);
        const rem=boards.filter(x=>x.id!==b.id);setBoards(rem);
        if(boardId===b.id)setBoardId(rem[0]?.id||null);
        toast("Quadro excluído permanentemente");setConfirmM(null);
      }
    });
  };

  if(authLoading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--bg)",flexDirection:"column",gap:16}}>
      <Spinner size={44}/><span style={{color:"var(--text3)",fontSize:14}}>Iniciando…</span>
    </div>
  );

  // Se usuário veio do link de redefinição de senha, mostrar tela de nova senha
  if(session&&isRecovery) return <ResetPasswordPage/>;

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
      onDashboard={canDashboard?()=>setPage("dashboard"):null}
      onAtividade={canDashboard?()=>setPage("atividade"):null}
      onMetas={()=>setPage("metas")}
      onLogout={signOut}
      open={sidebarOpen} onClose={()=>setSidebarOpen(false)}
    />
  );

  return (
    <div style={{display:"flex",height:"100vh",width:"100%",overflow:"hidden",background:"var(--bg)"}}>
      {!isMobile&&(
        <div
          onMouseEnter={()=>setSidebarHover(true)}
          onMouseLeave={()=>setSidebarHover(false)}
          style={{
            flexShrink:0,
            width:sidebarHover?242:56,
            transition:"width .22s cubic-bezier(.4,0,.2,1)",
            overflow:"hidden",
            position:"relative",
            zIndex:200,
            background:"var(--sidebar)",
          }}>
          {/* Mini sidebar: ícones quando recolhida */}
          {!sidebarHover&&(
            <div style={{width:56,height:"100%",display:"flex",flexDirection:"column",
              alignItems:"center",paddingTop:16,gap:4,overflowY:"auto"}}>
              <div style={{width:36,height:36,background:"var(--alt)",borderRadius:10,
                display:"flex",alignItems:"center",justifyContent:"center",
                fontWeight:900,fontSize:18,color:"#fff",marginBottom:8,flexShrink:0}}>M</div>
              <div onClick={()=>setPage("metas")} title="Metas e Recompensas"
                style={{width:36,height:36,display:"flex",alignItems:"center",justifyContent:"center",
                  borderRadius:8,cursor:"pointer",fontSize:18,color:"rgba(255,255,255,.6)",transition:"background .15s"}}
                onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.12)"}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>🏆</div>
              {canDashboard&&<>
                <div onClick={()=>setPage("dashboard")} title="Dashboard"
                  style={{width:36,height:36,display:"flex",alignItems:"center",justifyContent:"center",
                    borderRadius:8,cursor:"pointer",fontSize:18,color:"rgba(255,255,255,.6)",
                    transition:"background .15s"}}
                  onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.12)"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>📊</div>
                <div onClick={()=>setPage("atividade")} title="Atividade SDR"
                  style={{width:36,height:36,display:"flex",alignItems:"center",justifyContent:"center",
                    borderRadius:8,cursor:"pointer",fontSize:18,color:"rgba(255,255,255,.6)",
                    transition:"background .15s"}}
                  onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.12)"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>📞</div>
                <div style={{width:28,height:1,background:"rgba(255,255,255,.1)",margin:"4px 0"}}/>
              </>}
              {visibleBoards.map(b=>(
                <div key={b.id} onClick={()=>{setBoardId(b.id);setPage("board");}} title={b.nome}
                  style={{width:36,height:36,display:"flex",alignItems:"center",justifyContent:"center",
                    borderRadius:8,cursor:"pointer",fontSize:17,
                    background:b.id===boardId?"rgba(255,255,255,.18)":"transparent",
                    transition:"background .15s"}}
                  onMouseEnter={e=>{if(b.id!==boardId)e.currentTarget.style.background="rgba(255,255,255,.1)";}}
                  onMouseLeave={e=>{if(b.id!==boardId)e.currentTarget.style.background="transparent";}}>
                  {b.icon||"📋"}
                </div>
              ))}
              <div style={{flex:1}}/>
              <div title={profile?.nome||"Perfil"}
                style={{width:36,height:36,borderRadius:"50%",background:"var(--alt)",
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:13,fontWeight:800,color:"#fff",marginBottom:12,cursor:"pointer",flexShrink:0}}
                onClick={()=>setPage("profile")}>
                {(profile?.nome||"?")[0].toUpperCase()}
              </div>
            </div>
          )}
          {/* Full sidebar: expandida ao hover */}
          {sidebarHover&&<div style={{width:242,height:"100%",position:"absolute",top:0,left:0}}>{sidebar}</div>}
        </div>
      )}
      {isMobile&&sidebar}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
        {/* Mobile top bar */}
        {isMobile&&<div style={{background:"var(--surface)",borderBottom:"1px solid var(--border)",padding:"0 16px",height:52,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <button onClick={()=>setSidebarOpen(true)} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"var(--text)",lineHeight:1}}>☰</button>
          <div style={{fontWeight:800,fontSize:15,color:"var(--text)",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            {page==="dashboard"?"📊 Dashboard":page==="atividade"?"📞 Atividade SDR":page==="metas"?"🏆 Metas":page==="profile"?"👤 Perfil":page==="admin"?"⚙️ Admin":(boards.find(b=>b.id===boardId)?.icon||"📋")+" "+(boards.find(b=>b.id===boardId)?.nome||"CRM")}
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
        {page==="metas"&&(
          <MetasPage onBack={()=>setPage("board")} allUsers={allUsers} profile={profile} verTudo={canDashboard}/>
        )}
        {page==="atividade"&&canDashboard&&(
          <AtividadeSdrPage onBack={()=>setPage("board")} allUsers={allUsers} perms={perms}/>
        )}
        {page==="dashboard"&&(canDashboard
          ?<DashboardPage onBack={()=>setPage("board")} wsId={wsId} allUsers={allUsers} perms={perms} profile={profile}/>
          :<div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",color:"var(--text3)",fontSize:14,flexDirection:"column",gap:12}}>
              <div style={{fontSize:32}}>🔒</div>
              <div>Sem permissão para acessar esta área.</div>
              <button onClick={()=>setPage("board")} style={{...T.btn,background:"var(--blue)",color:"#fff",padding:"9px 20px",marginTop:8}}>Voltar</button>
            </div>
        )}
        {page==="board"&&(
          <BoardView key={boardId} boardId={boardId} boards={visibleBoards} allBoardsRaw={boards} allUsers={allUsers}
            currentUser={profile} wsId={wsId} perms={perms} onBoardCountChange={bumpCount}/>
        )}
      </div>
      {confirmM&&<ConfirmModal title={confirmM.title} message={confirmM.message} danger={confirmM.danger} onConfirm={confirmM.onConfirm} onCancel={()=>setConfirmM(null)} requirePhrase={confirmM.requirePhrase||null} boardName={confirmM.boardName||""}/>}
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
// RESET PASSWORD PAGE — exibida quando usuário volta do link de e-mail
// ─────────────────────────────────────────────────────────────────────────────
function ResetPasswordPage() {
  const {updatePassword,setIsRecovery,signOut}=useAuth();
  const toast=useToast();
  const [nova,setNova]=useState("");
  const [confirma,setConfirma]=useState("");
  const [busy,setBusy]=useState(false);
  const [done,setDone]=useState(false);

  const foc=e=>e.target.style.borderColor="var(--blue)";
  const blr=e=>e.target.style.borderColor="var(--border)";
  const match=nova&&confirma&&nova===confirma;
  const strong=nova.length>=8;

  const submit=async e=>{
    e.preventDefault();
    if(!match||!strong) return;
    setBusy(true);
    const {error}=await updatePassword(nova);
    if(error){ toast("Erro ao redefinir: "+error.message,"error"); setBusy(false); return; }
    setDone(true);
    toast("Senha redefinida com sucesso!","success");
    setTimeout(()=>{setIsRecovery(false);},2000);
    setBusy(false);
  };

  if(done) return (
    <AuthLayout>
      <div style={{textAlign:"center",padding:"8px 0 24px"}}>
        <div style={{fontSize:48,marginBottom:16}}>✅</div>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:10}}>Senha redefinida!</div>
        <div style={{fontSize:13,color:"var(--text2)"}}>Você será redirecionado para o CRM em instantes.</div>
      </div>
    </AuthLayout>
  );

  return (
    <AuthLayout>
      <div style={{marginBottom:24}}>
        <div style={{fontWeight:800,fontSize:18,color:"var(--text)",marginBottom:4}}>Nova senha</div>
        <div style={{fontSize:13,color:"var(--text2)"}}>Escolha uma senha segura para sua conta.</div>
      </div>
      <form onSubmit={submit}>
        <label style={T.lbl}>Nova senha</label>
        <input type="password" value={nova} onChange={e=>setNova(e.target.value)} required autoFocus
          placeholder="Mínimo 8 caracteres"
          style={{...T.inp,marginBottom:6,borderColor:nova&&!strong?"#dc2626":"var(--border)"}} onFocus={foc} onBlur={blr}/>
        {nova&&!strong&&<div style={{fontSize:12,color:"#dc2626",marginBottom:10}}>Use pelo menos 8 caracteres.</div>}
        {nova&&strong&&<div style={{fontSize:12,color:"#10b981",marginBottom:10}}>✓ Senha válida.</div>}

        <label style={T.lbl}>Confirmar nova senha</label>
        <input type="password" value={confirma} onChange={e=>setConfirma(e.target.value)} required
          placeholder="Repita a senha"
          style={{...T.inp,marginBottom:6,
            borderColor:confirma&&!match?"#dc2626":confirma&&match?"#10b981":"var(--border)"}} onFocus={foc} onBlur={blr}/>
        {confirma&&!match&&<div style={{fontSize:12,color:"#dc2626",marginBottom:10}}>As senhas não coincidem.</div>}
        {confirma&&match&&<div style={{fontSize:12,color:"#10b981",marginBottom:10}}>✓ As senhas coincidem.</div>}

        <div style={{marginTop:8}}>
          <button type="submit" disabled={busy||!match||!strong}
            style={{...T.btn,background:"var(--blue)",color:"#fff",width:"100%",padding:"13px",fontSize:15,
              opacity:busy||!match||!strong?0.5:1}}>
            {busy?"Redefinindo…":"Redefinir senha"}
          </button>
        </div>
      </form>
    </AuthLayout>
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
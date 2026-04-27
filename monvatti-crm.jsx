import { useState, useRef, useEffect, useCallback } from " ";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// ── SUPABASE ───────────────────────────────────────────────────────────────
const SUPA_URL = "https://hyhealogmqylciuzdmkz.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh5aGVhbG9nbXF5bGNpdXpkbWt6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYxNzQxNDgsImV4cCI6MjA5MTc1MDE0OH0.vDr-lQp6lBcT0f2bZGnTk1Jh6jjle9-oL7m2iAgw3nA";
const db = createClient(SUPA_URL, SUPA_KEY);

// ── BRAND ──────────────────────────────────────────────────────────────────
const B = { blue: "#001AD8", alt: "#3145FF", gun: "#2B333B", plat: "#DFDCDB" };
const GC = ["#3145FF","#001AD8","#0F7B6C","#B5451B","#7B2D8B","#1565C0","#475569"];
const SC = {
  Novo:"#3145FF","Em Contato":"#d97706","Sem Resposta":"#64748b",
  Qualificado:"#059669","Não Qualificado":"#dc2626",Descartado:"#6b7280",
  "Proposta na Rua":"#3145FF","Proposta Enviada":"#7c3aed","Em Análise":"#d97706",
  "Fechado/Ganho":"#059669",Recusado:"#dc2626","Negócio Futuro":"#ea580c",
  "Não compareceu":"#94a3b8",Inativo:"#94a3b8","Retomar Contato":"#d97706",
  "Prospecção Fria":"#3145FF",Tráfego:"#7c3aed",Indicação:"#059669",Outro:"#6b7280",
};

const TEAM = [
  { id:"u1", name:"João Silva",   initials:"JS", color:"#3145FF" },
  { id:"u2", name:"Maria Santos", initials:"MS", color:"#0F7B6C" },
  { id:"u3", name:"Pedro Costa",  initials:"PC", color:"#B5451B" },
  { id:"u4", name:"Ana Oliveira", initials:"AO", color:"#7B2D8B" },
];

// ── SMALL COMPONENTS ────────────────────────────────────────────────────────

function Spinner() {
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:300,flexDirection:"column",gap:16}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{width:36,height:36,border:`3px solid ${B.plat}`,borderTop:`3px solid ${B.blue}`,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
      <span style={{color:"#9ca3af",fontSize:14}}>Carregando…</span>
    </div>
  );
}

function Avatar({ userId, size=28 }) {
  const u = TEAM.find(x => x.id === userId);
  if (!u) return null;
  return (
    <div title={u.name} style={{width:size,height:size,borderRadius:"50%",background:u.color,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.35,fontWeight:500,flexShrink:0,border:"2px solid #fff",marginRight:-4}}>
      {u.initials}
    </div>
  );
}

function Badge({ value, onClick }) {
  const c = SC[value] || "#94a3b8";
  return (
    <span onClick={onClick} style={{display:"inline-flex",alignItems:"center",background:c+"22",color:c,border:`1px solid ${c}44`,borderRadius:6,padding:"2px 10px",fontSize:12,fontWeight:500,cursor:onClick?"pointer":"default",whiteSpace:"nowrap",userSelect:"none"}}>
      {value || "—"}
    </span>
  );
}

// ── CELLS ───────────────────────────────────────────────────────────────────

function EditableCell({ value, onChange, type="text", placeholder="" }) {
  const [ed, setEd] = useState(false);
  const [v, setV] = useState(value || "");
  const ref = useRef();
  useEffect(() => setV(value || ""), [value]);
  useEffect(() => { if (ed) ref.current?.focus(); }, [ed]);
  const commit = () => { setEd(false); if (v !== (value||"")) onChange(v||null); };
  if (!ed) return (
    <div onClick={() => setEd(true)} style={{padding:"3px 6px",minHeight:26,cursor:"text",fontSize:13,color:v?B.gun:"#c0ccd8",borderRadius:4}}>
      {v || placeholder || " "}
    </div>
  );
  return (
    <input ref={ref} type={type==="date"?"date":type==="number"?"number":"text"} value={v}
      onChange={e => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={e => { if(e.key==="Enter") commit(); if(e.key==="Escape"){setEd(false);setV(value||"");} }}
      style={{border:`2px solid ${B.blue}`,borderRadius:4,padding:"2px 6px",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}
    />
  );
}

function CurrencyCell({ value, onChange }) {
  const [ed, setEd] = useState(false);
  const [r, setR] = useState(value ?? "");
  useEffect(() => setR(value ?? ""), [value]);
  const fmtR = n => n == null ? "R$ —" : `R$ ${Number(n).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
  if (!ed) return (
    <div onClick={() => setEd(true)} style={{padding:"3px 6px",cursor:"text",fontSize:13,color:value!=null?B.gun:"#c0ccd8"}}>
      {fmtR(value)}
    </div>
  );
  return (
    <input autoFocus type="number" value={r}
      onChange={e => setR(e.target.value)}
      onBlur={() => { setEd(false); onChange(r!==""?parseFloat(r):null); }}
      onKeyDown={e => { if(e.key==="Enter"){setEd(false);onChange(r!==""?parseFloat(r):null);} }}
      style={{border:`2px solid ${B.blue}`,borderRadius:4,padding:"2px 6px",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}
    />
  );
}

function StatusCell({ value, options=[], onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef();
  useEffect(() => {
    const h = e => { if(ref.current&&!ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown",h); return () => document.removeEventListener("mousedown",h);
  },[]);
  return (
    <div ref={ref} style={{position:"relative"}}>
      <Badge value={value} onClick={() => setOpen(p=>!p)}/>
      {open && (
        <div style={{position:"absolute",top:"100%",left:0,zIndex:300,background:"#fff",border:"1px solid #e5e7eb",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,.12)",padding:6,minWidth:180,marginTop:4}}>
          {options.map(opt => (
            <div key={opt} onClick={() => {onChange(opt);setOpen(false);}}
              style={{padding:"6px 10px",borderRadius:6,cursor:"pointer",marginBottom:2,background:value===opt?"#f3f4f6":"transparent"}}
              onMouseEnter={e=>e.currentTarget.style.background="#f3f4f6"}
              onMouseLeave={e=>e.currentTarget.style.background=value===opt?"#f3f4f6":"transparent"}
            ><Badge value={opt}/></div>
          ))}
        </div>
      )}
    </div>
  );
}

function UserCell({ value=[], onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef();
  useEffect(() => {
    const h = e => { if(ref.current&&!ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown",h); return () => document.removeEventListener("mousedown",h);
  },[]);
  const cur = Array.isArray(value) ? value : [];
  const toggle = uid => onChange(cur.includes(uid)?cur.filter(x=>x!==uid):[...cur,uid]);
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div onClick={() => setOpen(p=>!p)} style={{display:"flex",alignItems:"center",cursor:"pointer",paddingLeft:4,minHeight:28}}>
        {!cur.length
          ? <div style={{width:26,height:26,borderRadius:"50%",border:"1.5px dashed #cbd5e1",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#94a3b8"}}>+</div>
          : cur.map(uid => <Avatar key={uid} userId={uid} size={26}/>)
        }
      </div>
      {open && (
        <div style={{position:"absolute",top:"100%",left:0,zIndex:300,background:"#fff",border:"1px solid #e5e7eb",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,.12)",padding:6,minWidth:200,marginTop:4}}>
          {TEAM.map(u => (
            <div key={u.id} onClick={() => toggle(u.id)}
              style={{display:"flex",alignItems:"center",gap:10,padding:"6px 10px",borderRadius:6,cursor:"pointer",background:cur.includes(u.id)?"#eff6ff":"transparent"}}
              onMouseEnter={e=>e.currentTarget.style.background="#f3f4f6"}
              onMouseLeave={e=>e.currentTarget.style.background=cur.includes(u.id)?"#eff6ff":"transparent"}
            >
              <Avatar userId={u.id} size={24}/>
              <span style={{fontSize:13}}>{u.name}</span>
              {cur.includes(u.id)&&<span style={{marginLeft:"auto",color:B.blue,fontSize:14}}>✓</span>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function LinkCell({ value, onChange }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:4}}>
      <EditableCell value={value} onChange={onChange} placeholder="https://"/>
      {value && <a href={value} target="_blank" rel="noopener noreferrer" style={{color:B.blue,fontSize:13,flexShrink:0,textDecoration:"none"}} onClick={e=>e.stopPropagation()}>↗</a>}
    </div>
  );
}

function CalcCell({ values }) {
  const v = parseFloat(values?.valor)||0;
  const p = parseFloat(values?.parcelas)||0;
  const r = p ? v/p : null;
  const fmt = n => `R$ ${Number(n).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
  return <div style={{padding:"3px 6px",fontSize:13,color:r!=null?"#059669":"#94a3b8",fontWeight:r!=null?500:400}}>{r!=null?fmt(r):"—"}</div>;
}

function Cell({ col, values, onChange }) {
  const v = values?.[col.id];
  const opts = col.config?.options || [];
  if (col.tipo==="calculated") return <CalcCell values={values}/>;
  if (col.tipo==="status")     return <StatusCell value={v} options={opts} onChange={onChange}/>;
  if (col.tipo==="user")       return <UserCell value={v} onChange={onChange}/>;
  if (col.tipo==="currency")   return <CurrencyCell value={v} onChange={onChange}/>;
  if (col.tipo==="link")       return <LinkCell value={v} onChange={onChange}/>;
  return <EditableCell value={v} onChange={onChange} type={col.tipo==="date"?"date":col.tipo==="number"?"number":"text"} placeholder={col.nome}/>;
}

// ── ITEM ROW ─────────────────────────────────────────────────────────────────

const iconBtn = {background:"none",border:"1px solid #e5e7eb",borderRadius:6,padding:"3px 8px",cursor:"pointer",fontSize:12,color:"#6b7280",fontFamily:"inherit"};
const rowMenu = {padding:"8px 12px",borderRadius:6,cursor:"pointer",fontSize:13,color:"#374151",display:"flex",alignItems:"center",gap:8};

function ItemRow({ item, columns, gc, onSelect, onDelete, onMoveInativa, onDupNeg, onDragStart }) {
  const [hov, setHov] = useState(false);
  const [menu, setMenu] = useState(false);
  const mref = useRef();
  useEffect(() => {
    const h = e => { if(mref.current&&!mref.current.contains(e.target)) setMenu(false); };
    document.addEventListener("mousedown",h); return () => document.removeEventListener("mousedown",h);
  },[]);
  return (
    <tr draggable onDragStart={onDragStart}
      onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)}
      style={{background:hov?"#f8f9fc":"#fff",transition:"background .1s",cursor:"grab"}}
    >
      <td style={{width:32,textAlign:"center",color:"#cbd5e1",fontSize:16,borderLeft:`3px solid ${gc}`,userSelect:"none"}}>{hov?"⋮⋮":""}</td>
      {columns.map(col => (
        <td key={col.id} style={{padding:"3px 6px",borderRight:"1px solid #f1f5f9",maxWidth:200,overflow:"hidden"}}>
          <Cell col={col} values={item.values} onChange={v => item._onChange(col.id, v)}/>
        </td>
      ))}
      <td style={{width:80,padding:"0 8px",textAlign:"right"}} ref={mref}>
        {hov && (
          <div style={{display:"flex",gap:4,justifyContent:"flex-end"}}>
            <button onClick={onSelect} style={iconBtn}>✏️</button>
            <div style={{position:"relative"}}>
              <button onClick={() => setMenu(p=>!p)} style={iconBtn}>⋯</button>
              {menu && (
                <div style={{position:"absolute",right:0,top:"100%",zIndex:400,background:"#fff",border:"1px solid #e5e7eb",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,.12)",minWidth:230,padding:6,marginTop:4}}>
                  {onDupNeg && <div onClick={()=>{setMenu(false);onDupNeg();}} style={rowMenu} onMouseEnter={e=>e.currentTarget.style.background="#f3f4f6"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>📋 Duplicar → Negociação</div>}
                  {onMoveInativa && <div onClick={()=>{setMenu(false);onMoveInativa();}} style={rowMenu} onMouseEnter={e=>e.currentTarget.style.background="#f3f4f6"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>📁 Mover para Base Inativa</div>}
                  <div style={{height:1,background:"#f1f5f9",margin:"4px 0"}}/>
                  <div onClick={()=>{setMenu(false);onDelete();}} style={{...rowMenu,color:"#dc2626"}} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>🗑️ Excluir item</div>
                </div>
              )}
            </div>
          </div>
        )}
      </td>
    </tr>
  );
}

// ── GROUP ────────────────────────────────────────────────────────────────────

const smbtn = {background:"none",border:"1px solid",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:12,fontFamily:"inherit"};

function Group({ group, columns, items, dragOver, onAddItem, onDelGroup, onRenameGroup, onToggle, onSelectItem, onUpdateValue, onDelItem, onMoveInativa, onDupNeg, onDragStart, onDragOver, onDrop }) {
  const [renaming, setRenaming] = useState(false);
  const [gname, setGname] = useState(group.nome);
  const nref = useRef();
  useEffect(() => { if(renaming) nref.current?.focus(); }, [renaming]);

  return (
    <div style={{marginTop:20,borderRadius:10,overflow:"hidden",border:dragOver===group.id?`2px solid ${group.color}`:"2px solid transparent",transition:"border .15s"}} onDragOver={onDragOver} onDrop={onDrop}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:group.color+"1A",borderLeft:`4px solid ${group.color}`}}>
        <button onClick={onToggle} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:group.color,padding:0,transform:group.collapsed?"rotate(-90deg)":"none",transition:"transform .2s",lineHeight:1}}>▼</button>
        {renaming
          ? <input ref={nref} value={gname} onChange={e=>setGname(e.target.value)}
              onBlur={()=>{setRenaming(false);onRenameGroup(gname);}}
              onKeyDown={e=>{if(e.key==="Enter"){setRenaming(false);onRenameGroup(gname);}}}
              style={{fontWeight:600,fontSize:14,color:group.color,border:"none",borderBottom:`2px solid ${group.color}`,background:"transparent",outline:"none",fontFamily:"inherit",width:200}}
            />
          : <span onDoubleClick={()=>setRenaming(true)} style={{fontWeight:600,fontSize:14,color:group.color,cursor:"pointer"}}>{group.nome}</span>
        }
        <span style={{fontSize:12,color:"#9ca3af",background:"#fff",borderRadius:20,padding:"1px 8px"}}>{items.length}</span>
        <div style={{marginLeft:"auto",display:"flex",gap:6}}>
          <button onClick={onAddItem} style={{...smbtn,color:group.color,borderColor:group.color+"55"}}>+ Item</button>
          <button onClick={onDelGroup} style={{...smbtn,color:"#dc2626",borderColor:"#fee2e2"}}>🗑️</button>
        </div>
      </div>

      {!group.collapsed && (
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,background:"#fff"}}>
            <thead>
              <tr style={{background:"#fafafa",borderBottom:"1px solid #f1f5f9"}}>
                <th style={{width:32,borderLeft:`3px solid ${group.color}`}}/>
                {columns.map(col => (
                  <th key={col.id} style={{padding:"8px 10px",textAlign:"left",fontWeight:500,fontSize:12,color:"#6b7280",whiteSpace:"nowrap",borderRight:"1px solid #f1f5f9"}}>{col.nome}</th>
                ))}
                <th style={{width:80}}/>
              </tr>
            </thead>
            <tbody>
              {items.map(item => {
                const enriched = {...item, _onChange:(cid,val)=>onUpdateValue(item.id,cid,val)};
                return (
                  <ItemRow key={item.id} item={enriched} columns={columns} gc={group.color}
                    onSelect={() => onSelectItem(item)}
                    onDelete={() => onDelItem(item.id)}
                    onMoveInativa={onMoveInativa?()=>onMoveInativa(item):null}
                    onDupNeg={onDupNeg?()=>onDupNeg(item):null}
                    onDragStart={e=>onDragStart(e,item)}
                  />
                );
              })}
            </tbody>
          </table>
          {!items.length && (
            <div style={{padding:"14px 48px",color:"#9ca3af",fontSize:13,background:"#fff",borderLeft:`3px solid ${group.color}`}}>Nenhum item neste grupo</div>
          )}
          <div style={{borderLeft:`3px solid ${group.color}`,background:"#fff"}}>
            <button onClick={onAddItem} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 16px",background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:13,fontFamily:"inherit"}}>
              + Adicionar item
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── RICH EDITOR ──────────────────────────────────────────────────────────────

function RichEditor({ onSubmit }) {
  const ref = useRef();
  const exec = cmd => { document.execCommand(cmd,false,null); ref.current?.focus(); };
  const submit = () => {
    const html = ref.current?.innerHTML;
    if (!html || html==="<br>") return;
    onSubmit(html);
    ref.current.innerHTML = "";
  };
  return (
    <div style={{border:"1px solid #e5e7eb",borderRadius:8,overflow:"hidden"}}>
      <div style={{display:"flex",gap:4,padding:"6px 10px",borderBottom:"1px solid #e5e7eb",background:"#fafafa"}}>
        {[["bold","B"],["italic","I"],["insertUnorderedList","• Lista"],["insertOrderedList","1. Lista"]].map(([cmd,lbl]) => (
          <button key={cmd} onMouseDown={e=>{e.preventDefault();exec(cmd);}}
            style={{background:"none",border:"1px solid #e5e7eb",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>{lbl}</button>
        ))}
      </div>
      <div ref={ref} contentEditable suppressContentEditableWarning
        style={{minHeight:70,padding:"10px 14px",outline:"none",fontSize:13,lineHeight:1.6,color:B.gun}}
        onKeyDown={e => { if(e.key==="Enter"&&e.metaKey) submit(); }}
      />
      <div style={{padding:"8px 10px",borderTop:"1px solid #e5e7eb",display:"flex",justifyContent:"flex-end",alignItems:"center",gap:8}}>
        <span style={{fontSize:12,color:"#9ca3af"}}>⌘↵ para publicar</span>
        <button onClick={submit} style={{background:B.blue,color:"#fff",border:"none",borderRadius:6,padding:"7px 16px",cursor:"pointer",fontSize:13,fontWeight:500,fontFamily:"inherit"}}>Publicar</button>
      </div>
    </div>
  );
}

// ── ITEM PANEL ───────────────────────────────────────────────────────────────

function ItemPanel({ item, board, onClose, onUpdateValue, onAddUpdate }) {
  const title = item.values?.empresa || item.values?.nome || "Sem título";
  return (
    <div style={{width:460,background:"#fff",borderLeft:"1px solid #e5e7eb",display:"flex",flexDirection:"column",height:"100%",flexShrink:0}}>
      <div style={{padding:"16px 20px",borderBottom:"1px solid #e5e7eb",display:"flex",alignItems:"center",gap:12,background:"#fafafa"}}>
        <div style={{flex:1}}>
          <div style={{fontSize:16,fontWeight:600,color:B.gun}}>{title}</div>
          <div style={{fontSize:12,color:"#9ca3af",marginTop:2}}>{board?.icon} {board?.nome}</div>
        </div>
        <button onClick={onClose} style={{background:"none",border:"1px solid #e5e7eb",borderRadius:6,fontSize:18,cursor:"pointer",color:"#9ca3af",width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"16px 20px"}}>
        <div style={{fontSize:11,fontWeight:500,color:"#9ca3af",marginBottom:12,textTransform:"uppercase",letterSpacing:.5}}>Campos</div>
        {board?.columns?.map(col => (
          <div key={col.id} style={{display:"flex",alignItems:"flex-start",marginBottom:10,gap:12}}>
            <div style={{width:130,fontSize:13,color:"#6b7280",flexShrink:0,paddingTop:5}}>{col.nome}</div>
            <div style={{flex:1}}>
              <Cell col={col} values={item.values} onChange={v=>onUpdateValue(col.id,v)}/>
            </div>
          </div>
        ))}
        <div style={{marginTop:24,paddingTop:20,borderTop:"1px solid #f1f5f9"}}>
          <div style={{fontSize:11,fontWeight:500,color:"#9ca3af",marginBottom:12,textTransform:"uppercase",letterSpacing:.5}}>Atualizações</div>
          <RichEditor onSubmit={onAddUpdate}/>
          <div style={{marginTop:16}}>
            {(item.updates||[]).map(u => (
              <div key={u.id} style={{padding:"12px 14px",background:"#f8f9fc",borderRadius:8,marginBottom:10,borderLeft:`3px solid ${B.alt}`}}>
                <div style={{fontSize:11,color:"#9ca3af",marginBottom:6}}>{new Date(u.created_at).toLocaleString("pt-BR")}</div>
                <div style={{fontSize:13,color:B.gun,lineHeight:1.6}} dangerouslySetInnerHTML={{__html:u.content}}/>
              </div>
            ))}
            {!(item.updates||[]).length && <div style={{fontSize:13,color:"#9ca3af",textAlign:"center",padding:"16px 0"}}>Sem atualizações ainda</div>}
          </div>
        </div>
      </div>
    </div>
  );
}

// ── SIDEBAR ──────────────────────────────────────────────────────────────────

function Sidebar({ boards, currentBoardId, onSelectBoard }) {
  return (
    <div style={{width:220,background:B.gun,color:"#fff",display:"flex",flexDirection:"column",flexShrink:0,height:"100%"}}>
      <div style={{padding:"20px 16px 14px",borderBottom:"1px solid #ffffff18"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,background:B.alt,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:15,flexShrink:0}}>M</div>
          <span style={{fontWeight:600,fontSize:15,letterSpacing:-.3}}>Monvatti CRM</span>
        </div>
      </div>
      <div style={{padding:"12px 16px 4px"}}>
        <div style={{fontSize:11,color:"#ffffff55",textTransform:"uppercase",letterSpacing:.5,marginBottom:8}}>Workspace</div>
        <div style={{fontSize:13,fontWeight:500,background:"#ffffff18",borderRadius:6,padding:"6px 10px"}}>Comercial</div>
      </div>
      <div style={{padding:"14px 16px 8px",flex:1,overflowY:"auto"}}>
        <div style={{fontSize:11,color:"#ffffff55",textTransform:"uppercase",letterSpacing:.5,marginBottom:8}}>Quadros</div>
        {boards.map(b => (
          <div key={b.id} onClick={() => onSelectBoard(b.id)}
            style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:6,cursor:"pointer",marginBottom:2,background:b.id===currentBoardId?"#ffffff22":"transparent",color:b.id===currentBoardId?"#fff":"#ffffffaa",fontSize:13}}
            onMouseEnter={e=>{ if(b.id!==currentBoardId) e.currentTarget.style.background="#ffffff12"; }}
            onMouseLeave={e=>{ if(b.id!==currentBoardId) e.currentTarget.style.background="transparent"; }}
          >
            <span style={{fontSize:15}}>{b.icon||"📋"}</span>
            <span>{b.nome}</span>
            <span style={{marginLeft:"auto",fontSize:11,color:"#ffffff44",background:"#ffffff18",borderRadius:10,padding:"1px 6px"}}>{b.totalItems||0}</span>
          </div>
        ))}
      </div>
      <div style={{padding:"12px 16px",borderTop:"1px solid #ffffff15"}}>
        <div style={{fontSize:11,color:"#ffffff55",marginBottom:8}}>Equipe</div>
        <div style={{display:"flex",alignItems:"center"}}>{TEAM.map(u=><Avatar key={u.id} userId={u.id} size={28}/>)}</div>
      </div>
    </div>
  );
}

// ── MAIN ─────────────────────────────────────────────────────────────────────

export default function CRM() {
  const [boards, setBoards]     = useState([]);
  const [boardId, setBoardId]   = useState(null);
  const [board, setBoard]       = useState(null);
  const [loading, setLoading]   = useState(true);
  const [selItem, setSelItem]   = useState(null);
  const [search, setSearch]     = useState("");
  const [dragItem, setDragItem] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  // ── Carregar lista de boards ──────────────────────────────
  useEffect(() => {
    (async () => {
      const { data } = await db.from("boards").select("*").order("ordem");
      if (!data) return;
      const counts = await Promise.all(data.map(async b => {
        const { count } = await db.from("items").select("*",{count:"exact",head:true}).eq("board_id",b.id);
        return { id:b.id, total:count||0 };
      }));
      const cmap = Object.fromEntries(counts.map(c=>[c.id,c.total]));
      const enriched = data.map(b=>({...b,totalItems:cmap[b.id]||0}));
      setBoards(enriched);
      setBoardId(enriched[0]?.id||null);
    })();
  },[]);

  // ── Carregar board completo ───────────────────────────────
  const loadBoard = useCallback(async bid => {
    if (!bid) return;
    setLoading(true);
    setSelItem(null);

    const [{ data:bData },{ data:cols },{ data:grps },{ data:itens }] = await Promise.all([
      db.from("boards").select("*").eq("id",bid).single(),
      db.from("columns").select("*").eq("board_id",bid).order("ordem"),
      db.from("groups").select("*").eq("board_id",bid).order("ordem"),
      db.from("items").select("id,group_id,ordem,created_at").eq("board_id",bid).order("ordem"),
    ]);

    // Valores dos itens
    const itemIds = (itens||[]).map(i=>i.id);
    let valMap = {};
    if (itemIds.length) {
      const { data:vals } = await db.from("item_values").select("*").in("item_id",itemIds);
      // colId → col object para encontrar o nome-slug
      const colById = Object.fromEntries((cols||[]).map(c=>[c.id,c]));
      for (const v of (vals||[])) {
        if (!valMap[v.item_id]) valMap[v.item_id] = {};
        const col = colById[v.column_id];
        // Chave = col.id (UUID) — Cell usa col.id como chave em values
        if (col) valMap[v.item_id][v.column_id] = v.value;
      }
    }

    const columns = (cols||[]).map(c=>({
      ...c,
      config: typeof c.config==="string"?JSON.parse(c.config):(c.config||{}),
    }));

    // Para CalcCell funcionar com "valor" e "parcelas", precisamos de um mapa
    // colName → colId para cada board
    const nameToId = Object.fromEntries(columns.map(c=>[c.nome.toLowerCase().replace(/ /g,"_"),c.id]));

    const groupsWithItems = (grps||[]).map(g=>({
      ...g,
      items: (itens||[])
        .filter(i=>i.group_id===g.id)
        .map(i=>({...i, values:valMap[i.id]||{}, updates:[], _nameToId:nameToId})),
    }));

    setBoard({...bData, columns, groups:groupsWithItems, nameToId});
    setLoading(false);
  },[]);

  useEffect(() => { if(boardId) loadBoard(boardId); },[boardId,loadBoard]);

  const updLocal = fn => setBoard(prev => {
    const next = JSON.parse(JSON.stringify(prev));
    fn(next);
    return next;
  });

  // ── Adicionar item ────────────────────────────────────────
  const addItem = async groupId => {
    const { data, error } = await db.from("items").insert({board_id:boardId,group_id:groupId,ordem:9999}).select().single();
    if (error||!data) return;
    updLocal(b=>{ b.groups.find(g=>g.id===groupId).items.push({...data,values:{},updates:[]}); });
    setBoards(prev=>prev.map(b=>b.id===boardId?{...b,totalItems:(b.totalItems||0)+1}:b));
  };

  // ── Atualizar valor ───────────────────────────────────────
  const updateValue = async (itemId, groupId, colId, val) => {
    updLocal(b=>{ const item=b.groups.find(g=>g.id===groupId)?.items.find(i=>i.id===itemId); if(item) item.values[colId]=val; });
    if (selItem?.id===itemId) setSelItem(p=>({...p,values:{...p.values,[colId]:val}}));
    await db.from("item_values").upsert({item_id:itemId,column_id:colId,value:val},{onConflict:"item_id,column_id"});
  };

  // ── Excluir item ──────────────────────────────────────────
  const delItem = async (groupId, itemId) => {
    updLocal(b=>{ const g=b.groups.find(g=>g.id===groupId); g.items=g.items.filter(i=>i.id!==itemId); });
    if (selItem?.id===itemId) setSelItem(null);
    setBoards(prev=>prev.map(b=>b.id===boardId?{...b,totalItems:Math.max(0,(b.totalItems||0)-1)}:b));
    await db.from("items").delete().eq("id",itemId);
  };

  // ── Criar grupo ───────────────────────────────────────────
  const addGroup = async () => {
    const nome = window.prompt("Nome do grupo:");
    if (!nome?.trim()) return;
    const color = GC[(board?.groups.length||0)%GC.length];
    const { data, error } = await db.from("groups").insert({nome:nome.trim(),board_id:boardId,color,ordem:9999}).select().single();
    if (error||!data) return;
    updLocal(b=>b.groups.push({...data,items:[]}));
  };

  // ── Excluir grupo ─────────────────────────────────────────
  const delGroup = async groupId => {
    if (!window.confirm("Excluir grupo e todos os itens?")) return;
    const qty = board?.groups.find(g=>g.id===groupId)?.items.length||0;
    updLocal(b=>{ b.groups=b.groups.filter(g=>g.id!==groupId); });
    setBoards(prev=>prev.map(b=>b.id===boardId?{...b,totalItems:Math.max(0,(b.totalItems||0)-qty)}:b));
    await db.from("groups").delete().eq("id",groupId);
  };

  // ── Renomear / toggle grupo ───────────────────────────────
  const renameGroup = async (groupId, nome) => {
    updLocal(b=>{ b.groups.find(g=>g.id===groupId).nome=nome; });
    await db.from("groups").update({nome}).eq("id",groupId);
  };
  const toggleGroup = groupId => updLocal(b=>{ const g=b.groups.find(g=>g.id===groupId); g.collapsed=!g.collapsed; });

  // ── Atualizações (notas) ──────────────────────────────────
  const loadUpdates = async (item, groupId) => {
    const { data } = await db.from("item_updates").select("*").eq("item_id",item.id).order("created_at",{ascending:false});
    setSelItem({...item, updates:data||[], _groupId:groupId});
    updLocal(b=>{ const g=b.groups.find(g=>g.id===groupId); const idx=g.items.findIndex(i=>i.id===item.id); if(idx>=0) g.items[idx]={...g.items[idx],updates:data||[]}; });
  };

  const addUpdate = async (itemId, groupId, content) => {
    const { data } = await db.from("item_updates").insert({item_id:itemId,content}).select().single();
    if (!data) return;
    updLocal(b=>{ const item=b.groups.find(g=>g.id===groupId)?.items.find(i=>i.id===itemId); if(item) item.updates.unshift(data); });
    setSelItem(p=>p?({...p,updates:[data,...(p.updates||[])]}):p);
  };

  // ── Mover para Base Inativa ───────────────────────────────
  const moveInativa = async (groupId, item) => {
    const inativoBoard = boards.find(b=>b.nome==="Base Inativa");
    if (!inativoBoard) return alert("Board 'Base Inativa' não encontrado.");
    const { data:tgtGroups } = await db.from("groups").select("*").eq("board_id",inativoBoard.id).order("ordem");
    const isTrafego = board?.nome==="Tráfego";
    const tgtGroup = isTrafego
      ? tgtGroups?.find(g=>g.nome==="Tráfego")
      : tgtGroups?.find(g=>g.nome==="Prospecção Fria") || tgtGroups?.[0];
    if (!tgtGroup) return alert("Grupo de destino não encontrado.");

    await db.from("items").update({board_id:inativoBoard.id,group_id:tgtGroup.id}).eq("id",item.id);
    updLocal(b=>{ const g=b.groups.find(g=>g.id===groupId); g.items=g.items.filter(i=>i.id!==item.id); });
    if (selItem?.id===item.id) setSelItem(null);
    setBoards(prev=>prev.map(b=>{
      if (b.id===boardId) return {...b,totalItems:Math.max(0,(b.totalItems||0)-1)};
      if (b.id===inativoBoard.id) return {...b,totalItems:(b.totalItems||0)+1};
      return b;
    }));
  };

  // ── Duplicar para Negociação ──────────────────────────────
  const dupNeg = async (groupId, item) => {
    const negBoard = boards.find(b=>b.nome==="Negociações");
    if (!negBoard) return alert("Board 'Negociações' não encontrado.");
    const { data:negGroups } = await db.from("groups").select("*").eq("board_id",negBoard.id).order("ordem");
    if (!negGroups?.length) return;

    const opts = negGroups.map((g,i)=>`${i+1}. ${g.nome}`).join("\n");
    const p = window.prompt(`Grupo de destino em Negociações:\n\n${opts}\n\nDigite o número:`);
    const idx = parseInt(p)-1;
    if (isNaN(idx)||idx<0||idx>=negGroups.length) return;

    const { data:newItem } = await db.from("items").insert({board_id:negBoard.id,group_id:negGroups[idx].id,ordem:9999}).select().single();
    if (!newItem) return;

    // Copiar values
    const { data:vals } = await db.from("item_values").select("*").eq("item_id",item.id);
    if (vals?.length) {
      await db.from("item_values").insert(vals.map(v=>({column_id:v.column_id,item_id:newItem.id,value:v.value})));
    }
    setBoards(prev=>prev.map(b=>b.id===negBoard.id?{...b,totalItems:(b.totalItems||0)+1}:b));
  };

  // ── Drag & drop ───────────────────────────────────────────
  const dstart = (e,groupId,item) => { setDragItem({groupId,item}); e.dataTransfer.effectAllowed="move"; };
  const dover  = (e,groupId)     => { e.preventDefault(); setDragOver(groupId); };
  const ddrop  = async (e,tgtGroupId) => {
    e.preventDefault();
    if (!dragItem||dragItem.groupId===tgtGroupId) { setDragItem(null);setDragOver(null); return; }
    updLocal(b=>{
      const src=b.groups.find(g=>g.id===dragItem.groupId);
      const tgt=b.groups.find(g=>g.id===tgtGroupId);
      src.items=src.items.filter(i=>i.id!==dragItem.item.id);
      tgt.items.push(dragItem.item);
    });
    setDragItem(null); setDragOver(null);
    await db.from("items").update({group_id:tgtGroupId}).eq("id",dragItem.item.id);
  };

  // ── Filtro ────────────────────────────────────────────────
  const filterItems = items => {
    if (!search.trim()) return items;
    const q = search.toLowerCase();
    return items.filter(item=>Object.values(item.values||{}).some(v=>v&&String(v).toLowerCase().includes(q)));
  };

  const canActions = board && ["Prospecção Fria","Tráfego"].includes(board.nome);

  return (
    <div style={{display:"flex",height:"100vh",fontFamily:"system-ui,-apple-system,sans-serif",overflow:"hidden",background:"#f0f2f5"}}>
      <Sidebar boards={boards} currentBoardId={boardId} onSelectBoard={id=>{setBoardId(id);setSelItem(null);setSearch("");}}/>

      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        {/* Topbar */}
        <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"0 24px",display:"flex",alignItems:"center",gap:12,height:56,flexShrink:0}}>
          <div style={{fontSize:18,fontWeight:600,color:B.gun,flex:1}}>{board?.icon||"📋"} {board?.nome||"…"}</div>
          <input placeholder="🔍 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}
            style={{border:"1px solid #e5e7eb",borderRadius:8,padding:"7px 14px",fontSize:13,width:220,outline:"none",color:B.gun,background:"#f8f9fc"}}
          />
          <button onClick={addGroup} style={{background:B.blue,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:13,fontWeight:500,cursor:"pointer",fontFamily:"inherit"}}>+ Novo Grupo</button>
        </div>

        {/* Board content */}
        <div style={{flex:1,overflowY:"auto",padding:"12px 24px 48px"}}>
          {loading
            ? <Spinner/>
            : board?.groups.map(group => (
                <Group key={group.id} group={group} columns={board.columns}
                  items={filterItems(group.items)}
                  dragOver={dragOver}
                  onAddItem={()=>addItem(group.id)}
                  onDelGroup={()=>delGroup(group.id)}
                  onRenameGroup={n=>renameGroup(group.id,n)}
                  onToggle={()=>toggleGroup(group.id)}
                  onSelectItem={item=>loadUpdates(item,group.id)}
                  onUpdateValue={(iid,cid,val)=>updateValue(iid,group.id,cid,val)}
                  onDelItem={iid=>delItem(group.id,iid)}
                  onMoveInativa={canActions?item=>moveInativa(group.id,item):null}
                  onDupNeg={canActions?item=>dupNeg(group.id,item):null}
                  onDragStart={(e,item)=>dstart(e,group.id,item)}
                  onDragOver={e=>dover(e,group.id)}
                  onDrop={e=>ddrop(e,group.id)}
                />
              ))
          }
          {!loading && (
            <button onClick={addGroup} style={{marginTop:16,display:"flex",alignItems:"center",gap:8,background:"none",border:"1.5px dashed #cbd5e1",borderRadius:8,padding:"10px 20px",cursor:"pointer",color:"#94a3b8",fontSize:13,fontFamily:"inherit"}}>
              + Adicionar grupo
            </button>
          )}
        </div>
      </div>

      {selItem && board && (
        <ItemPanel
          item={selItem} board={board}
          onClose={()=>setSelItem(null)}
          onUpdateValue={(cid,val)=>updateValue(selItem.id,selItem._groupId,cid,val)}
          onAddUpdate={content=>addUpdate(selItem.id,selItem._groupId,content)}
        />
      )}
    </div>
  );
}

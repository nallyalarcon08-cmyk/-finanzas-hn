import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

const KEY_ING  = "finanzas:ingresos";
const KEY_EG   = "finanzas:egresos";
const KEY_PRES = "finanzas:presupuesto";
const KEY_VER  = "finanzas:version";
const APP_VER  = "2.0"; // Cambia esto si quieres resetear datos en todos los dispositivos

// ── Almacenamiento local (funciona fuera de Claude) ──
const storageGet = (key) => {
  try {
    const val = localStorage.getItem(key);
    return val ? { value: val } : null;
  } catch { return null; }
};
const storageSet = (key, value) => {
  try { localStorage.setItem(key, value); }
  catch (e) { console.error("Storage lleno:", e); }
};

// ── Datos de ejemplo Mayo 2026 ──
const SEED_ING = [
  { id:"i1",  fecha:"2026-05-13", concepto:"Prestaciones",          tipo:"Laboral",           banco:"Efectivo",        monto:14992.78 },
  { id:"i2",  fecha:"2026-05-13", concepto:"BAC",                   tipo:"Banco",             banco:"BAC",             monto:4000.00  },
  { id:"i3",  fecha:"2026-05-13", concepto:"Ficohsa",               tipo:"Banco",             banco:"Ficohsa",         monto:4000.00  },
  { id:"i4",  fecha:"2026-05-13", concepto:"Allan",                 tipo:"Préstamo recibido", banco:"Efectivo",        monto:2240.00  },
  { id:"i5",  fecha:"2026-05-13", concepto:"Comida Rocío",          tipo:"Apoyo familiar",    banco:"Efectivo",        monto:232.00   },
  { id:"i6",  fecha:"2026-05-13", concepto:"Abuela",                tipo:"Apoyo familiar",    banco:"Efectivo",        monto:2000.00  },
  { id:"i7",  fecha:"2026-05-13", concepto:"Efectivo Rocío",        tipo:"Efectivo",          banco:"Efectivo",        monto:2000.00  },
  { id:"i8",  fecha:"2026-05-13", concepto:"Compra baterías/mouse", tipo:"Venta",             banco:"Efectivo",        monto:90.00    },
  { id:"i9",  fecha:"2026-05-13", concepto:"Compra Jugos",          tipo:"Venta",             banco:"Efectivo",        monto:30.00    },
  { id:"i10", fecha:"2026-05-13", concepto:"Compra número",         tipo:"Venta",             banco:"Efectivo",        monto:20.00    },
  { id:"i11", fecha:"2026-05-13", concepto:"Préstamo Yury",         tipo:"Préstamo recibido", banco:"Efectivo",        monto:300.00   },
];
const SEED_EG = [
  { id:"e1",  fecha:"2026-05-13", concepto:"Compra Ingrid",               categoria:"Personal",        banco:"BAC",            monto:272.00  },
  { id:"e2",  fecha:"2026-05-13", concepto:"Cuenta Netflix",              categoria:"Entretenimiento", banco:"BAC",            monto:374.21  },
  { id:"e3",  fecha:"2026-05-13", concepto:"Préstamo Jami",               categoria:"Préstamo pagado", banco:"BAC",            monto:500.00  },
  { id:"e4",  fecha:"2026-05-13", concepto:"Vana",                        categoria:"Personal",        banco:"BAC",            monto:1000.00 },
  { id:"e5",  fecha:"2026-05-13", concepto:"Allan",                       categoria:"Personal",        banco:"BAC",            monto:280.00  },
  { id:"e6",  fecha:"2026-05-13", concepto:"Netflix abuelo",              categoria:"Entretenimiento", banco:"BAC",            monto:186.97  },
  { id:"e7",  fecha:"2026-05-13", concepto:"Vana",                        categoria:"Personal",        banco:"BAC",            monto:1000.00 },
  { id:"e8",  fecha:"2026-05-13", concepto:"Protector y vidrio templado", categoria:"Electrónico",     banco:"Ficohsa",        monto:350.00  },
  { id:"e9",  fecha:"2026-05-13", concepto:"Farmacia",                    categoria:"Salud",           banco:"Ficohsa",        monto:208.98  },
  { id:"e10", fecha:"2026-05-13", concepto:"Cafetería Uni",               categoria:"Alimentación",    banco:"Ficohsa",        monto:98.00   },
  { id:"e11", fecha:"2026-05-13", concepto:"Super mamá",                  categoria:"Familia",         banco:"BA (Atlántida)", monto:125.00  },
  { id:"e12", fecha:"2026-05-13", concepto:"Boleto Concierto",            categoria:"Entretenimiento", banco:"BA (Atlántida)", monto:1650.00 },
  { id:"e13", fecha:"2026-05-13", concepto:"Recarga Hellen",              categoria:"Comunicación",    banco:"BA (Atlántida)", monto:67.00   },
  { id:"e14", fecha:"2026-05-13", concepto:"Depósito a BAC",              categoria:"Transferencia",   banco:"BA (Atlántida)", monto:400.00  },
  { id:"e15", fecha:"2026-05-13", concepto:"Comisión bancaria",           categoria:"Bancario",        banco:"BA (Atlántida)", monto:40.00   },
];

const TIPOS_ING = ["Salario","Netflix / Plataforma","Donación","Banco","Préstamo recibido","Venta","Apoyo familiar","Efectivo","Otro"];
const CATS_EG   = [
  // Personal y estilo de vida
  "Ropa y calzado","Maquillaje y belleza","Joyería y accesorios","Peluquería / Estética","Gym / Deporte","Higiene personal","Limpieza del hogar",
  // Alimentación
  "Supermercado","Restaurantes","Comida rápida","Bebidas / Cafetería",
  // Hogar
  "Alquiler","Electricidad","Agua","Internet","Cable / TV",
  // Transporte
  "Gasolina","Taxi / Uber","Bus / Transporte público","Mantenimiento de vehículo",
  // Salud
  "Farmacia","Doctor / Consulta","Laboratorio",
  // Entretenimiento y comunicación
  "Entretenimiento","Comunicación","Educación",
  // Familia
  "Familia",
  // Finanzas
  "Préstamo pagado","Transferencia","Bancario","Ahorro","Inversión","Deudas",
  // Otros
  "Mascotas","Viajes","Regalos","Personal","Electrónico","Otro"
];
const BANCOS    = ["BAC","Ficohsa","BA (Atlántida)","Efectivo","Otro"];

const CAT_COLOR = {
  // Personal
  "Ropa y calzado":"#f472b6","Maquillaje y belleza":"#e879f9","Joyería y accesorios":"#c084fc",
  "Peluquería / Estética":"#a78bfa","Gym / Deporte":"#34d399","Higiene personal":"#38bdf8","Limpieza del hogar":"#67e8f9",
  // Alimentación
  "Supermercado":"#fb923c","Restaurantes":"#f97316","Comida rápida":"#ef4444","Bebidas / Cafetería":"#fbbf24",
  // Hogar
  "Alquiler":"#60a5fa","Electricidad":"#fde047","Agua":"#38bdf8","Internet":"#818cf8","Cable / TV":"#6366f1",
  // Transporte
  "Gasolina":"#f87171","Taxi / Uber":"#fb923c","Bus / Transporte público":"#fca5a5","Mantenimiento de vehículo":"#f43f5e",
  // Salud
  "Farmacia":"#34d399","Doctor / Consulta":"#10b981","Laboratorio":"#6ee7b7",
  // Entretenimiento
  "Entretenimiento":"#a78bfa","Comunicación":"#38bdf8","Educación":"#c084fc","Familia":"#f472b6",
  // Finanzas
  "Préstamo pagado":"#fbbf24","Transferencia":"#6ee7b7","Bancario":"#94a3b8","Ahorro":"#34d399","Inversión":"#10b981","Deudas":"#ef4444",
  // Otros
  "Mascotas":"#fb923c","Viajes":"#60a5fa","Regalos":"#f472b6","Personal":"#60a5fa","Electrónico":"#818cf8","Otro":"#9ca3af",
};
const TIPO_COLOR = {
  Salario:"#34d399","Netflix / Plataforma":"#ef4444", Donación:"#f472b6",
  Banco:"#60a5fa","Préstamo recibido":"#fbbf24",
  Venta:"#a78bfa","Apoyo familiar":"#f472b6",Efectivo:"#fb923c",Otro:"#9ca3af",
};
const BANCO_COLOR = {
  "BAC":"#60a5fa","Ficohsa":"#f472b6","BA (Atlántida)":"#fbbf24","Efectivo":"#34d399","Otro":"#94a3b8",
};

const L  = n => `L ${Number(n).toLocaleString("es-HN",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const uid = () => Math.random().toString(36).slice(2,10);

export default function App() {
  const [ingresos,    setIngresos]    = useState([]);
  const [egresos,     setEgresos]     = useState([]);
  const [presupuesto, setPresupuesto] = useState({});
  const [loaded,      setLoaded]      = useState(false);
  const [tab,         setTab]         = useState("dash");
  const [modal,       setModal]       = useState(null);
  const [form,        setForm]        = useState({});
  const [confirm,     setConfirm]     = useState(null);
  const [toast,       setToast]       = useState(null);
  const [filterB,     setFilterB]     = useState("Todos");
  const [filterM,     setFilterM]     = useState("Todos");
  const [saving,      setSaving]      = useState(false);
  const [editId,      setEditId]      = useState(null);

  // ── Cargar datos ──
  useEffect(() => {
    const ver = localStorage.getItem(KEY_VER);
    if(ver !== APP_VER) {
      // Primera vez o versión anterior — limpiar datos de ejemplo
      localStorage.removeItem(KEY_ING);
      localStorage.removeItem(KEY_EG);
      localStorage.removeItem(KEY_PRES);
      localStorage.setItem(KEY_VER, APP_VER);
    }
    const ri = storageGet(KEY_ING);
    const re = storageGet(KEY_EG);
    const rp = storageGet(KEY_PRES);
    setIngresos(ri ? JSON.parse(ri.value) : []);
    setEgresos( re ? JSON.parse(re.value) : []);
    setPresupuesto(rp ? JSON.parse(rp.value) : {});
    setLoaded(true);
  }, []);

  const saveIng = (data) => { setSaving(true); storageSet(KEY_ING,  JSON.stringify(data)); setSaving(false); };
  const saveEg  = (data) => { setSaving(true); storageSet(KEY_EG,   JSON.stringify(data)); setSaving(false); };
  const savePres= (data) => { storageSet(KEY_PRES, JSON.stringify(data)); };

  const showToast = (msg) => { setToast(msg); setTimeout(()=>setToast(null),2800); };

  // ── Meses disponibles ──
  const meses = useMemo(()=>{
    const s = new Set([...ingresos,...egresos].map(r=>r.fecha.slice(0,7)));
    return ["Todos",...[...s].sort().reverse()];
  },[ingresos,egresos]);

  const filtI = useMemo(()=>filterM==="Todos"?ingresos:ingresos.filter(r=>r.fecha.startsWith(filterM)),[ingresos,filterM]);
  const filtE = useMemo(()=>{
    let r = filterM==="Todos"?egresos:egresos.filter(r=>r.fecha.startsWith(filterM));
    if(filterB!=="Todos") r = r.filter(e=>e.banco===filterB);
    return r;
  },[egresos,filterM,filterB]);

  const totalI  = useMemo(()=>filtI.reduce((s,r)=>s+r.monto,0),[filtI]);
  const totalE  = useMemo(()=>filtE.reduce((s,r)=>s+r.monto,0),[filtE]);
  const balance = totalI - totalE;

  const byBanco = useMemo(()=>{ const m={}; filtE.forEach(e=>{ m[e.banco]=(m[e.banco]||0)+e.monto; }); return Object.entries(m).sort((a,b)=>b[1]-a[1]); },[filtE]);
  const byCat   = useMemo(()=>{ const m={}; filtE.forEach(e=>{ m[e.categoria]=(m[e.categoria]||0)+e.monto; }); return Object.entries(m).sort((a,b)=>b[1]-a[1]); },[filtE]);

  // ── Alertas ──
  const alertas = useMemo(()=>{
    const lista = [];
    if(balance < 0)
      lista.push({ tipo:"error",   icono:"🚨", texto:`Balance negativo este mes: ${L(balance)}` });
    else if(balance < 500)
      lista.push({ tipo:"warning", icono:"⚠️", texto:`Balance muy bajo: ${L(balance)}. Considera reducir gastos.` });
    Object.entries(presupuesto).forEach(([cat,limite])=>{
      if(limite<=0) return;
      const gastado = filtE.filter(e=>e.categoria===cat).reduce((s,e)=>s+e.monto,0);
      if(gastado > limite)
        lista.push({ tipo:"error",   icono:"📛", texto:`${cat}: gastaste ${L(gastado)} de ${L(limite)} presupuestados` });
      else if(gastado > limite*0.85)
        lista.push({ tipo:"warning", icono:"⚠️", texto:`${cat}: llevas el ${((gastado/limite)*100).toFixed(0)}% del presupuesto` });
    });
    return lista;
  },[balance, filtE, presupuesto]);

  // ── Exportar Excel ──
  const exportToExcel = () => {
    const autoWidth = (ws) => {
      const data = XLSX.utils.sheet_to_json(ws,{header:1});
      ws["!cols"] = data[0].map((_,i)=>({ wch: Math.max(...data.map(r=>(r[i]?String(r[i]).length:0))) }));
    };
    const wb      = XLSX.utils.book_new();
    const wsIng   = XLSX.utils.json_to_sheet(ingresos.map(r=>({ Fecha:r.fecha, Concepto:r.concepto, Tipo:r.tipo, Banco:r.banco, "Monto (L)":r.monto })));
    const wsEg    = XLSX.utils.json_to_sheet(egresos.map(r=>({ Fecha:r.fecha, Concepto:r.concepto, Categoría:r.categoria, Banco:r.banco, "Monto (L)":r.monto })));
    autoWidth(wsIng); autoWidth(wsEg);
    XLSX.utils.book_append_sheet(wb, wsIng, "Ingresos");
    XLSX.utils.book_append_sheet(wb, wsEg,  "Gastos");
    const mes = filterM==="Todos"?"completo":filterM;
    XLSX.writeFile(wb, `FinanzasHN_${mes}.xlsx`);
  };

  const openModal = (type) => { setForm({ fecha:new Date().toISOString().split("T")[0] }); setEditId(null); setModal(type); };

  const openEdit = (type, r) => {
    setForm({ ...r });
    setEditId(r.id);
    setModal(type);
  };

  const handleSave = () => {
    if(!form.concepto?.trim()||!form.monto||!form.fecha) return;
    const monto = parseFloat(form.monto);
    if(isNaN(monto)||monto<=0) return;
    if(modal==="ing"){
      const registro = {...form,id:editId||uid(),monto,tipo:form.tipo||"Otro",banco:form.banco||"Efectivo"};
      const updated = editId ? ingresos.map(r=>r.id===editId?registro:r) : [...ingresos,registro];
      setIngresos(updated); saveIng(updated); showToast(editId?"✓ Ingreso actualizado":"✓ Ingreso guardado");
    } else {
      const registro = {...form,id:editId||uid(),monto,categoria:form.categoria||"Otro",banco:form.banco||"Efectivo"};
      const updated = editId ? egresos.map(r=>r.id===editId?registro:r) : [...egresos,registro];
      setEgresos(updated); saveEg(updated); showToast(editId?"✓ Gasto actualizado":"✓ Gasto guardado");
    }
    setModal(null); setEditId(null);
  };

  const handleDelete = () => {
    if(!confirm) return;
    if(confirm.type==="ing"){ const u=ingresos.filter(r=>r.id!==confirm.id); setIngresos(u); saveIng(u); }
    else                    { const u=egresos.filter(r=>r.id!==confirm.id);  setEgresos(u);  saveEg(u);  }
    showToast("Registro eliminado"); setConfirm(null);
  };

  const updatePresupuesto = (cat, val) => {
    const updated = {...presupuesto,[cat]:Number(val)};
    setPresupuesto(updated); savePres(updated);
  };

  if(!loaded) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#080b14",color:"#60a5fa",fontFamily:"'DM Sans',sans-serif",fontSize:16}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:40,marginBottom:12}}>💰</div><div>Cargando tus finanzas...</div></div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#080b14",color:"#e2e8f0",fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;500;700&family=Playfair+Display:wght@700;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;height:4px;}
        ::-webkit-scrollbar-thumb{background:#1e2340;border-radius:4px;}
        .card{background:#0f1320;border:1px solid #1a2035;border-radius:20px;}
        .pill{display:inline-flex;align-items:center;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;}
        .btn{cursor:pointer;border:none;font-family:inherit;font-weight:700;border-radius:12px;transition:all .15s;}
        .btn:hover{filter:brightness(1.12);}
        .btn:active{transform:scale(.97);}
        .tab{cursor:pointer;border:none;background:none;font-family:inherit;font-size:13px;font-weight:600;padding:9px 14px;border-radius:10px;color:#475569;transition:all .2s;}
        .tab.on{background:#131929;color:#e2e8f0;box-shadow:0 1px 8px #00000044;}
        .inp{width:100%;padding:11px 14px;background:#0d1019;border:1px solid #1a2035;border-radius:11px;color:#e2e8f0;font-family:inherit;font-size:14px;outline:none;transition:border-color .2s;}
        .inp:focus{border-color:#3b82f6;}
        select.inp option{background:#0d1019;}
        .row:hover{background:#0f1524!important;}
        .overlay{position:fixed;inset:0;background:rgba(0,0,0,.75);backdrop-filter:blur(8px);z-index:200;display:flex;align-items:flex-end;justify-content:center;}
        .modal{background:#0f1320;border:1px solid #1a2035;width:100%;max-width:460px;padding:28px 24px;border-radius:24px 24px 0 0;}
        .del-btn{opacity:0;background:#ff4444;color:#fff;border:none;border-radius:7px;width:28px;height:28px;cursor:pointer;font-size:12px;display:flex;align-items:center;justify-content:center;transition:opacity .2s;flex-shrink:0;}
        .edit-btn{opacity:0;background:#2563eb;color:#fff;border:none;border-radius:7px;width:28px;height:28px;cursor:pointer;font-size:12px;display:flex;align-items:center;justify-content:center;transition:opacity .2s;flex-shrink:0;}
        .row:hover .del-btn{opacity:1;}
        .row:hover .edit-btn{opacity:1;}
        .toast{position:fixed;bottom:90px;left:50%;transform:translateX(-50%);background:#131929;border:1px solid #1e2a40;border-radius:999px;padding:10px 20px;font-size:13px;font-weight:600;color:#e2e8f0;z-index:300;pointer-events:none;white-space:nowrap;}
        .stat{background:linear-gradient(135deg,#0f1320 0%,#141829 100%);border:1px solid #1a2035;border-radius:20px;padding:20px;}
        @keyframes fadeUp{from{opacity:0;transform:translateY(12px);}to{opacity:1;transform:translateY(0);}}
        .fade-up{animation:fadeUp .3s ease both;}
        @keyframes slideUp{from{transform:translateY(100%);}to{transform:translateY(0);}}
        .slide-up{animation:slideUp .25s cubic-bezier(.4,0,.2,1);}
      `}</style>

      {/* HEADER */}
      <header style={{background:"#080b14",borderBottom:"1px solid #1a2035",padding:"16px 20px",position:"sticky",top:0,zIndex:50}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap"}}>
          <div style={{fontFamily:"'Playfair Display',serif",fontWeight:900,fontSize:20,color:"#fff"}}>
            FinanzasHN <span style={{opacity:.3,fontSize:14}}>·</span>{" "}
            <span style={{fontSize:13,fontWeight:300,fontFamily:"'DM Sans',sans-serif",color:"#475569"}}>
              {saving?"guardando...":"✓ guardado"}
            </span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <select className="inp" style={{width:"auto",fontSize:13,padding:"8px 12px"}} value={filterM} onChange={e=>setFilterM(e.target.value)}>
              {meses.map(m=><option key={m}>{m}</option>)}
            </select>
            <button className="btn" onClick={()=>openModal("ing")} style={{background:"linear-gradient(135deg,#065f46,#10b981)",color:"#fff",padding:"9px 16px",fontSize:13}}>+ Ingreso</button>
            <button className="btn" onClick={()=>openModal("eg")}  style={{background:"linear-gradient(135deg,#7f1d1d,#ef4444)",color:"#fff",padding:"9px 16px",fontSize:13}}>+ Gasto</button>
            <button className="btn" onClick={exportToExcel}         style={{background:"linear-gradient(135deg,#1e3a5f,#2563eb)",color:"#fff",padding:"9px 16px",fontSize:13}}>⬇ Excel</button>
          </div>
        </div>
      </header>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"20px 16px 100px"}}>

        {/* STATS */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:14,marginBottom:22}}>
          {[
            {label:"Ingresos",  v:totalI,  color:"#10b981", sub:`${filtI.length} registros`},
            {label:"Gastos",    v:totalE,  color:"#ef4444", sub:`${filtE.length} registros`},
            {label:"Balance",   v:balance, color:balance>=0?"#10b981":"#ef4444", sub:balance>=0?"Positivo":"Déficit"},
            {label:"% Gastado", v:null,    color:"#a78bfa", pct:totalI>0?(totalE/totalI*100).toFixed(1):"0", sub:"del ingreso"},
          ].map((s,i)=>(
            <div key={i} className="stat fade-up" style={{borderLeft:`3px solid ${s.color}`,animationDelay:`${i*.06}s`}}>
              <div style={{fontSize:11,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:".6px"}}>{s.label}</div>
              <div style={{fontSize:24,fontWeight:700,color:s.color,marginTop:6,fontFamily:"'Playfair Display',serif",letterSpacing:"-1px"}}>
                {s.pct!=null?`${s.pct}%`:L(s.v)}
              </div>
              <div style={{fontSize:11,color:"#334155",marginTop:3}}>{s.sub}</div>
            </div>
          ))}
        </div>

        {/* TABS */}
        <div style={{display:"flex",gap:4,marginBottom:18,background:"#080b14",padding:4,borderRadius:13,border:"1px solid #1a2035",width:"fit-content",flexWrap:"wrap"}}>
          {[["dash","📊 Resumen"],["ing","💚 Ingresos"],["eg","🔴 Gastos"],["pres","📋 Presupuesto"],["alertas","🔔 Alertas"]].map(([k,l])=>(
            <button key={k} className={`tab ${tab===k?"on":""}`} onClick={()=>setTab(k)}>{l}</button>
          ))}
        </div>

        {/* ── DASHBOARD ── */}
        {tab==="dash" && (
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:16}}>
            <div className="card fade-up" style={{padding:22}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:15,marginBottom:18,color:"#fff"}}>Gastos por Banco</div>
              {byBanco.length===0 && <div style={{color:"#334155",fontSize:13}}>Sin datos</div>}
              {byBanco.map(([banco,tot])=>{
                const pct=(tot/Math.max(totalE,1)*100).toFixed(1);
                const color=BANCO_COLOR[banco]||"#94a3b8";
                return(<div key={banco} style={{marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:5}}>
                    <span style={{fontWeight:600}}>{banco}</span><span style={{color,fontWeight:700}}>{L(tot)}</span>
                  </div>
                  <div style={{height:6,background:"#1a2035",borderRadius:999,overflow:"hidden"}}>
                    <div style={{height:"100%",width:`${pct}%`,background:color,borderRadius:999}}/>
                  </div>
                  <div style={{fontSize:10,color:"#334155",marginTop:3,textAlign:"right"}}>{pct}%</div>
                </div>);
              })}
            </div>

            <div className="card fade-up" style={{padding:22,animationDelay:".08s"}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:15,marginBottom:18,color:"#fff"}}>Gastos por Categoría</div>
              {byCat.length===0 && <div style={{color:"#334155",fontSize:13}}>Sin datos</div>}
              {byCat.slice(0,8).map(([cat,tot])=>{
                const pct=(tot/Math.max(byCat[0][1],1)*100).toFixed(0);
                const color=CAT_COLOR[cat]||"#94a3b8";
                return(<div key={cat} style={{marginBottom:11,display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:8,height:8,borderRadius:"50%",background:color,flexShrink:0}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
                      <span style={{color:"#cbd5e1",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cat}</span>
                      <span style={{fontWeight:700,color,flexShrink:0,marginLeft:8}}>{L(tot)}</span>
                    </div>
                    <div style={{height:4,background:"#1a2035",borderRadius:999}}>
                      <div style={{height:"100%",width:`${pct}%`,background:color,borderRadius:999}}/>
                    </div>
                  </div>
                </div>);
              })}
            </div>

            <div className="card fade-up" style={{padding:22,animationDelay:".12s"}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:15,marginBottom:14,color:"#fff"}}>Últimos Ingresos</div>
              {[...filtI].reverse().slice(0,6).map(r=>(
                <div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:"1px solid #111827"}}>
                  <div style={{minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.concepto}</div>
                    <span className="pill" style={{background:(TIPO_COLOR[r.tipo]||"#9ca3af")+"20",color:TIPO_COLOR[r.tipo]||"#9ca3af",marginTop:3}}>{r.tipo}</span>
                  </div>
                  <div style={{color:"#10b981",fontWeight:700,fontSize:14,flexShrink:0,marginLeft:12}}>{L(r.monto)}</div>
                </div>
              ))}
              {filtI.length===0&&<div style={{color:"#334155",fontSize:13}}>Sin ingresos en este período</div>}
            </div>

            <div className="card fade-up" style={{padding:22,animationDelay:".16s"}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:15,marginBottom:14,color:"#fff"}}>Últimos Gastos</div>
              {[...filtE].reverse().slice(0,6).map(r=>(
                <div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:"1px solid #111827"}}>
                  <div style={{minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.concepto}</div>
                    <div style={{display:"flex",gap:5,marginTop:3,flexWrap:"wrap"}}>
                      <span className="pill" style={{background:(CAT_COLOR[r.categoria]||"#9ca3af")+"20",color:CAT_COLOR[r.categoria]||"#9ca3af"}}>{r.categoria}</span>
                      <span className="pill" style={{background:"#1a2035",color:"#64748b"}}>{r.banco}</span>
                    </div>
                  </div>
                  <div style={{color:"#ef4444",fontWeight:700,fontSize:14,flexShrink:0,marginLeft:12}}>{L(r.monto)}</div>
                </div>
              ))}
              {filtE.length===0&&<div style={{color:"#334155",fontSize:13}}>Sin gastos en este período</div>}
            </div>
          </div>
        )}

        {/* ── INGRESOS ── */}
        {tab==="ing" && (
          <div className="card fade-up" style={{overflow:"hidden"}}>
            <div style={{padding:"18px 20px",borderBottom:"1px solid #1a2035",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:16,color:"#fff"}}>Todos los Ingresos</div>
              <div style={{color:"#10b981",fontWeight:700,fontSize:18,fontFamily:"'Playfair Display',serif"}}>{L(totalI)}</div>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",minWidth:520}}>
                <thead>
                  <tr style={{background:"#080b14"}}>
                    {["Fecha","Concepto","Tipo","Banco","Monto",""].map((h,i)=>(
                      <th key={i} style={{padding:"11px 14px",textAlign:i===4?"right":"left",fontSize:10,fontWeight:700,color:"#334155",textTransform:"uppercase",letterSpacing:".6px",whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...filtI].reverse().map(r=>(
                    <tr key={r.id} className="row" style={{borderTop:"1px solid #111827",background:"transparent",transition:"background .15s"}}>
                      <td style={{padding:"12px 14px",fontSize:12,color:"#475569",whiteSpace:"nowrap"}}>{r.fecha}</td>
                      <td style={{padding:"12px 14px",fontSize:13,fontWeight:600}}>{r.concepto}</td>
                      <td style={{padding:"12px 14px"}}><span className="pill" style={{background:(TIPO_COLOR[r.tipo]||"#9ca3af")+"20",color:TIPO_COLOR[r.tipo]||"#9ca3af"}}>{r.tipo}</span></td>
                      <td style={{padding:"12px 14px",fontSize:12,color:"#64748b"}}>{r.banco}</td>
                      <td style={{padding:"12px 14px",fontSize:14,fontWeight:700,color:"#10b981",textAlign:"right",whiteSpace:"nowrap"}}>{L(r.monto)}</td>
                      <td style={{padding:"12px 8px",textAlign:"center"}}><div style={{display:"flex",gap:4,justifyContent:"center"}}><button className="edit-btn" onClick={()=>openEdit("ing",r)}>✏</button><button className="del-btn" onClick={()=>setConfirm({type:"ing",id:r.id,concepto:r.concepto})}>✕</button></div></td>
                    </tr>
                  ))}
                  {filtI.length===0&&<tr><td colSpan={6} style={{padding:"28px",textAlign:"center",color:"#334155",fontSize:13}}>No hay ingresos en este período</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── GASTOS ── */}
        {tab==="eg" && (
          <div className="card fade-up" style={{overflow:"hidden"}}>
            <div style={{padding:"18px 20px",borderBottom:"1px solid #1a2035",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
              <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:16,color:"#fff"}}>Todos los Gastos</div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <select className="inp" style={{width:"auto",fontSize:13,padding:"7px 12px"}} value={filterB} onChange={e=>setFilterB(e.target.value)}>
                  <option>Todos</option>
                  {BANCOS.map(b=><option key={b}>{b}</option>)}
                </select>
                <div style={{color:"#ef4444",fontWeight:700,fontSize:18,fontFamily:"'Playfair Display',serif"}}>{L(totalE)}</div>
              </div>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",minWidth:520}}>
                <thead>
                  <tr style={{background:"#080b14"}}>
                    {["Fecha","Concepto","Categoría","Banco","Monto",""].map((h,i)=>(
                      <th key={i} style={{padding:"11px 14px",textAlign:i===4?"right":"left",fontSize:10,fontWeight:700,color:"#334155",textTransform:"uppercase",letterSpacing:".6px",whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...filtE].reverse().map(r=>(
                    <tr key={r.id} className="row" style={{borderTop:"1px solid #111827",background:"transparent",transition:"background .15s"}}>
                      <td style={{padding:"12px 14px",fontSize:12,color:"#475569",whiteSpace:"nowrap"}}>{r.fecha}</td>
                      <td style={{padding:"12px 14px",fontSize:13,fontWeight:600}}>{r.concepto}</td>
                      <td style={{padding:"12px 14px"}}><span className="pill" style={{background:(CAT_COLOR[r.categoria]||"#9ca3af")+"20",color:CAT_COLOR[r.categoria]||"#9ca3af"}}>{r.categoria}</span></td>
                      <td style={{padding:"12px 14px",fontSize:12,color:"#64748b"}}>{r.banco}</td>
                      <td style={{padding:"12px 14px",fontSize:14,fontWeight:700,color:"#ef4444",textAlign:"right",whiteSpace:"nowrap"}}>{L(r.monto)}</td>
                      <td style={{padding:"12px 8px",textAlign:"center"}}><div style={{display:"flex",gap:4,justifyContent:"center"}}><button className="edit-btn" onClick={()=>openEdit("eg",r)}>✏</button><button className="del-btn" onClick={()=>setConfirm({type:"eg",id:r.id,concepto:r.concepto})}>✕</button></div></td>
                    </tr>
                  ))}
                  {filtE.length===0&&<tr><td colSpan={6} style={{padding:"28px",textAlign:"center",color:"#334155",fontSize:13}}>No hay gastos en este período</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── PRESUPUESTO ── */}
        {tab==="pres" && (
          <div className="card fade-up" style={{padding:24}}>
            <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:18,color:"#fff",marginBottom:6}}>Presupuesto Mensual</div>
            <div style={{fontSize:13,color:"#475569",marginBottom:24}}>Define cuánto planeas gastar en cada categoría. La barra se llena conforme gastas.</div>
            {CATS_EG.map(cat=>{
              const gastado = filtE.filter(e=>e.categoria===cat).reduce((s,e)=>s+e.monto,0);
              const limite  = presupuesto[cat]||0;
              const pct     = limite>0?Math.min((gastado/limite)*100,100):0;
              const excede  = gastado>limite && limite>0;
              const color   = CAT_COLOR[cat]||"#9ca3af";
              return(
                <div key={cat} style={{marginBottom:22,borderBottom:"1px solid #111827",paddingBottom:18}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:10,height:10,borderRadius:"50%",background:color}}/>
                      <span style={{fontSize:14,fontWeight:600,color:"#e2e8f0"}}>{cat}</span>
                    </div>
                    <span style={{fontSize:13,fontWeight:700,color:excede?"#ef4444":"#10b981"}}>
                      {L(gastado)} {limite>0?`/ ${L(limite)}`:""} 
                    </span>
                  </div>
                  {limite>0 && (
                    <div style={{height:8,background:"#1a2035",borderRadius:999,marginBottom:10,overflow:"hidden"}}>
                      <div style={{height:"100%",width:`${pct}%`,background:excede?"#ef4444":color,borderRadius:999,transition:"width .5s"}}/>
                    </div>
                  )}
                  <input className="inp" type="number" placeholder="Límite mensual en L (ej: 2000)"
                    value={presupuesto[cat]||""} min="0"
                    onChange={e=>updatePresupuesto(cat,e.target.value)}
                    style={{fontSize:13,padding:"8px 12px"}}
                  />
                </div>
              );
            })}
          </div>
        )}

        {/* ── ALERTAS ── */}
        {tab==="alertas" && (
          <div className="card fade-up" style={{padding:24}}>
            <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:18,color:"#fff",marginBottom:6}}>🔔 Alertas Activas</div>
            <div style={{fontSize:13,color:"#475569",marginBottom:24}}>Se generan automáticamente según tu presupuesto y balance mensual.</div>
            {alertas.length===0 ? (
              <div style={{textAlign:"center",padding:"40px 0",color:"#334155"}}>
                <div style={{fontSize:40,marginBottom:12}}>✅</div>
                <div style={{fontSize:15,fontWeight:600,color:"#10b981"}}>Todo en orden</div>
                <div style={{fontSize:13,marginTop:6}}>No hay alertas activas este mes</div>
              </div>
            ) : alertas.map((a,i)=>(
              <div key={i} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"14px 16px",borderRadius:12,marginBottom:10,
                background:a.tipo==="error"?"#450a0a":"#431407",
                borderLeft:`3px solid ${a.tipo==="error"?"#ef4444":"#f97316"}`}}>
                <span style={{fontSize:22,flexShrink:0}}>{a.icono}</span>
                <span style={{fontSize:13,color:"#e2e8f0",lineHeight:1.5}}>{a.texto}</span>
              </div>
            ))}
            {alertas.length>0 && (
              <div style={{marginTop:20,padding:"12px 16px",background:"#0f1320",borderRadius:12,fontSize:12,color:"#475569"}}>
                💡 Para reducir alertas, define límites en la pestaña Presupuesto o agrega más ingresos.
              </div>
            )}
          </div>
        )}
      </div>

      {/* MODAL AGREGAR */}
      {modal && (
        <div className="overlay" onClick={()=>setModal(null)}>
          <div className="modal slide-up" onClick={e=>e.stopPropagation()}>
            <div style={{width:40,height:4,background:"#1a2035",borderRadius:999,margin:"0 auto 20px"}}/>
            <div style={{fontFamily:"'Playfair Display',serif",fontWeight:900,fontSize:22,color:modal==="ing"?"#10b981":"#ef4444",marginBottom:4}}>
              {editId ? (modal==="ing"?"Editar Ingreso":"Editar Gasto") : (modal==="ing"?"Nuevo Ingreso":"Nuevo Gasto")}
            </div>
            <div style={{fontSize:12,color:"#334155",marginBottom:22}}>{editId?"Modifica los datos y toca Guardar":"Completa los datos y toca Guardar"}</div>
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div>
                <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".5px",display:"block",marginBottom:6}}>Concepto *</label>
                <input className="inp" placeholder="Ej: Salario, Farmacia..." value={form.concepto||""} onChange={e=>setForm(f=>({...f,concepto:e.target.value}))}/>
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".5px",display:"block",marginBottom:6}}>{modal==="ing"?"Tipo":"Categoría"}</label>
                <select className="inp" value={modal==="ing"?(form.tipo||""):(form.categoria||"")}
                  onChange={e=>setForm(f=>modal==="ing"?{...f,tipo:e.target.value}:{...f,categoria:e.target.value})}>
                  <option value="">Seleccionar...</option>
                  {(modal==="ing"?TIPOS_ING:CATS_EG).map(c=><option key={c}>{c}</option>)}
                </select>
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".5px",display:"block",marginBottom:6}}>Banco</label>
                <select className="inp" value={form.banco||""} onChange={e=>setForm(f=>({...f,banco:e.target.value}))}>
                  <option value="">Seleccionar...</option>
                  {BANCOS.map(b=><option key={b}>{b}</option>)}
                </select>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div>
                  <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".5px",display:"block",marginBottom:6}}>Monto (L) *</label>
                  <input className="inp" type="number" placeholder="0.00" min="0" step="0.01" value={form.monto||""} onChange={e=>setForm(f=>({...f,monto:e.target.value}))}/>
                </div>
                <div>
                  <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".5px",display:"block",marginBottom:6}}>Fecha *</label>
                  <input className="inp" type="date" value={form.fecha||""} onChange={e=>setForm(f=>({...f,fecha:e.target.value}))}/>
                </div>
              </div>
              <div style={{display:"flex",gap:10,marginTop:4}}>
                <button className="btn" onClick={()=>setModal(null)} style={{flex:1,background:"#131929",color:"#64748b",padding:"13px"}}>Cancelar</button>
                <button className="btn" onClick={handleSave}
                  style={{flex:2,background:modal==="ing"?"linear-gradient(135deg,#065f46,#10b981)":"linear-gradient(135deg,#7f1d1d,#ef4444)",color:"#fff",padding:"13px",fontSize:14}}>
                  {editId ? "Actualizar" : (modal==="ing"?"Guardar Ingreso":"Guardar Gasto")}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL CONFIRMAR BORRADO */}
      {confirm && (
        <div className="overlay" onClick={()=>setConfirm(null)}>
          <div className="modal slide-up" style={{maxWidth:360}} onClick={e=>e.stopPropagation()}>
            <div style={{width:40,height:4,background:"#1a2035",borderRadius:999,margin:"0 auto 20px"}}/>
            <div style={{fontSize:36,textAlign:"center",marginBottom:10}}>🗑️</div>
            <div style={{fontFamily:"'Playfair Display',serif",fontWeight:700,fontSize:18,textAlign:"center",marginBottom:8}}>¿Eliminar registro?</div>
            <div style={{fontSize:13,color:"#475569",textAlign:"center",marginBottom:22}}>
              Se eliminará <strong style={{color:"#e2e8f0"}}>"{confirm.concepto}"</strong>.<br/>Esta acción no se puede deshacer.
            </div>
            <div style={{display:"flex",gap:10}}>
              <button className="btn" onClick={()=>setConfirm(null)} style={{flex:1,background:"#131929",color:"#64748b",padding:"12px"}}>Cancelar</button>
              <button className="btn" onClick={handleDelete} style={{flex:1,background:"linear-gradient(135deg,#7f1d1d,#ef4444)",color:"#fff",padding:"12px"}}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}

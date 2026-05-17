import { useState, useEffect, useCallback, useMemo } from "react";
import { createClient } from "@supabase/supabase-js";
import JSZip from "jszip";
import ExcelJS from "exceljs";

// ─── CONFIGURATION SUPABASE ───────────────────────────────────────────────────
// 🔴 Remplacez ces valeurs par vos vraies clés Supabase
const SUPABASE_URL  = "https://rducxnyxgqrmtyvqwfzw.supabase.co";
const SUPABASE_ANON = "sb_publishable_2K8GkCgsLbHzSy-1FhP2cg_cWIVD-vH";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);
const EMAILJS_SERVICE  = "service_n7ugp1c";   // votre Service ID
const EMAILJS_TEMPLATE = "template_t7ro56h";  // votre Template ID
const EMAILJS_KEY      = "n2N8upDAe3RfPiddE";    // votre Public Key

// ─── ADMIN (accès illimité sans abonnement) ───────────────────────────────────
const ADMIN_EMAIL = "martin13haya@gmail.com"; // 🔑 Email admin — aucune restriction

// ─── CONTACT SUPPORT ──────────────────────────────────────────────────────────
const CONTACT_WHATSAPP = "2290196078696";        // Numéro WhatsApp (avec indicatif)
const CONTACT_EMAIL    = "martin13haya@gmail.com"; // Email de contact

// ─── PLANS D'ABONNEMENT ───────────────────────────────────────────────────────
// 💡 Modifiez les prix ci-dessous pour changer les tarifs affichés (en FCFA)
const PLANS = {
  starter:    { name:"Starter",    monthly:5000,  annual:48000,  companies:1,   color:"#1a6dd6", icon:"🚀" },
  pro:        { name:"Pro",        monthly:12000, annual:115200, companies:3,   color:"#7c3aed", icon:"⭐" },
  enterprise: { name:"Enterprise", monthly:25000, annual:240000, companies:999, color:"#d97706", icon:"👑" },
};

// ─── CSS RESPONSIVE GLOBAL ────────────────────────────────────────────────────
const injectCSS = () => {
  if(document.getElementById("rhpaie-responsive")) return;
  const style = document.createElement("style");
  style.id = "rhpaie-responsive";
  style.textContent = `
    * { box-sizing: border-box; }
    html, body, #root { max-width: 100%; overflow-x: hidden; }
    body.rh-no-scroll { overflow: hidden; }
    img { max-width: 100%; height: auto; }
    button, input, select, textarea { max-width: 100%; }

    /* ── SIDEBAR ── */
    .rh-sidebar {
      width: 220px;
      background: #1a3a6b;
      display: flex;
      flex-direction: column;
      height: 100vh;
      position: fixed;
      top: 0; left: 0;
      z-index: 400;
    }
    .rh-main {
      margin-left: 220px;
      padding: 32px;
      min-height: 100vh;
      background: #f0f4ff;
      width: calc(100% - 220px);
    }
    .rh-topbar { display: none; }
    .rh-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.55);
      z-index: 350;
    }
    .rh-mobile-close { display: none; }

    /* ── SURFACES / CONTROLS ── */
    .rh-card { width: 100%; }
    .rh-btn {
      max-width: 100%;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      text-align: center;
      white-space: normal;
    }
    .rh-control {
      min-width: 0;
      background: #fff;
      border: 1px solid #d0dff5;
      border-radius: 8px;
      padding: 8px 12px;
      color: #1a2a4a;
      font-size: 13px;
      font-family: inherit;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    /* ── LAYOUT HELPERS ── */
    .rh-grid-4 { display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin-bottom: 24px; }
    .rh-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 0 16px; }
    .rh-grid-3 { display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; }
    .rh-grid-5 { display: grid; grid-template-columns: repeat(5,1fr); gap: 8px; }
    .rh-grid-auto { display: grid; grid-template-columns: repeat(auto-fit,minmax(130px,1fr)); gap: 8px; }
    .rh-form-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 0 16px; }
    .rh-form-grid-3 { display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; }
    .rh-summary-grid-4 { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; }
    .rh-summary-grid-3 { display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; }
    .rh-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    .rh-table-wrap table { min-width: 640px; width: 100%; }
    .rh-flex-filters { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
    .rh-page-header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 24px; }
    .rh-btn-group { display: flex; gap: 8px; flex-wrap: wrap; }
    .rh-inline-actions { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
    .rh-actions-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
    .rh-checkbox-row { display: flex; gap: 16px; margin-bottom: 14px; flex-wrap: wrap; }
    .rh-item-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; }
    .rh-item-footer { margin-top: 8px; display: flex; justify-content: flex-end; }
    .rh-badges { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
    .rh-logo-preview { margin-bottom: 8px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .rh-warning-banner {
      background: #fff3cd;
      border: 1px solid #d97706;
      border-radius: 10px;
      padding: 14px 18px;
      margin-bottom: 24px;
      font-size: 13px;
      color: #92400e;
    }

    /* ── MODAL ── */
    .rh-modal-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(26,61,107,0.6);
      z-index: 1000;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .rh-modal {
      background: #ffffff;
      border: 1px solid #d0dff5;
      border-radius: 10px;
      max-width: 95vw;
      max-height: 90vh;
      overflow-y: auto;
      box-shadow: 0 8px 32px rgba(26,61,107,0.15);
    }
    .rh-modal-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 20px 24px;
      border-bottom: 1px solid #d0dff5;
    }
    .rh-modal-body { padding: 24px; }

    /* ── TABLET 768–1024px ── */
    @media (min-width: 768px) and (max-width: 1024px) {
      .rh-sidebar { width: 68px; }
      .rh-sidebar-label { display: none !important; }
      .rh-sidebar-brand-text { display: none !important; }
      .rh-main { margin-left: 68px; width: calc(100% - 68px); padding: 20px; }
      .rh-grid-4, .rh-summary-grid-4 { grid-template-columns: repeat(2,1fr); }
      .rh-grid-5 { grid-template-columns: repeat(3,1fr); }
      .rh-summary-grid-3 { grid-template-columns: repeat(2,1fr); }
    }

    /* ── MOBILE <768px ── */
    @media (max-width: 767px) {
      .rh-topbar {
        display: flex !important;
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 56px;
        background: #1a3a6b;
        z-index: 300;
        align-items: center;
        justify-content: space-between;
        padding: 0 16px;
        box-shadow: 0 2px 8px rgba(26,61,107,0.4);
      }
      .rh-sidebar {
        transform: translateX(-100%);
        transition: transform 0.28s cubic-bezier(.4,0,.2,1);
        width: min(82vw, 290px);
        z-index: 400;
        box-shadow: 4px 0 24px rgba(0,0,0,0.25);
      }
      .rh-sidebar.open { transform: translateX(0); }
      .rh-overlay.open { display: block !important; }
      .rh-mobile-close { display: block !important; }
      .rh-main {
        margin-left: 0 !important;
        width: 100% !important;
        padding: 68px 12px 24px !important;
      }
      .rh-card { padding: 16px !important; }
      .rh-grid-4, .rh-summary-grid-4 { grid-template-columns: repeat(2,1fr) !important; gap: 10px !important; }
      .rh-grid-5 { grid-template-columns: repeat(2,1fr) !important; }
      .rh-grid-3 { grid-template-columns: repeat(2,1fr) !important; }
      .rh-grid-2, .rh-form-grid-2, .rh-form-grid-3, .rh-summary-grid-3 { grid-template-columns: 1fr !important; }
      .rh-page-header, .rh-inline-actions, .rh-actions-row, .rh-item-header {
        flex-direction: column;
        align-items: stretch !important;
      }
      .rh-btn-group, .rh-inline-actions, .rh-actions-row, .rh-badges { width: 100%; }
      .rh-badges { justify-content: flex-start; }
      .rh-btn-group button, .rh-inline-actions button, .rh-actions-row button {
        min-width: 0;
        flex: 1 1 auto;
        font-size: 12px !important;
        padding: 8px 10px !important;
      }
      .rh-flex-filters { flex-direction: column; align-items: stretch; }
      .rh-flex-filters select,
      .rh-flex-filters input,
      .rh-inline-actions select,
      .rh-inline-actions input,
      .rh-control { width: 100% !important; }
      .rh-checkbox-row { flex-direction: column; gap: 10px; }
      .rh-item-footer { justify-content: stretch; }
      .rh-item-footer > * { width: 100%; }
      .rh-modal-backdrop { padding: 8px; align-items: flex-end; }
      .rh-modal {
        width: 100% !important;
        max-width: 100% !important;
        max-height: calc(100vh - 8px) !important;
        border-radius: 14px 14px 0 0;
      }
      .rh-modal-header { padding: 16px; }
      .rh-modal-body { padding: 16px; }
    }

    @media (max-width: 480px) {
      .rh-grid-4, .rh-grid-5, .rh-grid-3, .rh-summary-grid-4 { grid-template-columns: 1fr !important; }
      .rh-table-wrap table { min-width: 560px; }
    }
  `;
  document.head.appendChild(style);
};

// Inject CSS immediately
if(typeof document !== "undefined") injectCSS();

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
const MOIS = ["","Janvier","Février","Mars","Avril","Mai","Juin",
               "Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
const CURRENT_YEAR = new Date().getFullYear();

// ─── CALCULS PAIE (portage exact du Python) ───────────────────────────────────
const roundTo5 = v => Math.ceil((parseFloat(v)||0)/5)*5;

const calculateIts = brut => {
  const bf = Math.floor((parseFloat(brut)||0)/1000)*1000;
  if(bf<=60000) return 0;
  let its=0;
  [[60000,150000,0.10],[150000,250000,0.15],[250000,500000,0.19],[500000,Infinity,0.30]]
    .forEach(([lo,hi,r])=>{ if(bf>lo) its+=(Math.min(bf,hi)-lo)*r; });
  return Math.round(its);
};

const taxeRadioAuto = (base,mois) => {
  const b=parseFloat(base||0), m=parseInt(mois||0);
  if(b<=60000) return m===3?1000:0;
  if(m===3) return 1000;
  if(m===6) return 3000;
  return 0;
};

const vpsExonereCheck = (nationalite,premierEmploi,dateEmbauche,
                          nouvelleEntreprise,datePremierExercice,mois,annee) => {
  if(premierEmploi && dateEmbauche){
    try{
      const emb=new Date(dateEmbauche);
      const paie=new Date(annee,mois-1,1);
      const anc=((paie.getFullYear()-emb.getFullYear())*12)+(paie.getMonth()-emb.getMonth());
      if(anc<24) return {exonere:true,motif:`Premier emploi — ancienneté ${anc} mois`,restants:24-anc};
    }catch(e){}
  }
  if(nouvelleEntreprise && datePremierExercice){
    try{
      const nat=(nationalite||"").trim().toUpperCase();
      if(nat==="BÉNINOISE"||nat==="BENINOISE"){
        const dp=new Date(datePremierExercice);
        const cur=new Date(annee,mois-1,1);
        if(cur.getFullYear()===dp.getFullYear())
          return {exonere:true,motif:`Nouvelle entreprise — 1er exercice ${dp.getFullYear()} (Art. 192 CGI)`,restants:0};
      }
    }catch(e){}
  }
  return {exonere:false,motif:"",restants:0};
};

const calculatePayroll = d => {
  const base=parseFloat(d.salaire_base||0),
        primes=parseFloat(d.primes||0), indem=parseFloat(d.indemnites||0),
        avances=parseFloat(d.avances||0), saisie=parseFloat(d.saisie_arret||0),
        assurance=parseFloat(d.assurance_sante||0), mois=parseInt(d.mois||0),
        annee=parseInt(d.annee||CURRENT_YEAR), tauxPat=parseFloat(d.taux_cnss_patronale||0.194);
  // Calcul heures supplémentaires par catégorie
  const tauxHoraire = base / 173.33;
  const h12=parseFloat(d.h_sup_12||0), h35=parseFloat(d.h_sup_35||0),
        h50=parseFloat(d.h_sup_50||0), h100=parseFloat(d.h_sup_100||0);
  const montant_h12=Math.round(h12*tauxHoraire*1.12);
  const montant_h35=Math.round(h35*tauxHoraire*1.35);
  const montant_h50=Math.round(h50*tauxHoraire*1.50);
  const montant_h100=Math.round(h100*tauxHoraire*2.00);
  const hsup=montant_h12+montant_h35+montant_h50+montant_h100;
  const taxeRadio=taxeRadioAuto(base,mois);
  const brut=base+hsup+primes+indem;
  const cnssOuvriere=Math.round(brut*0.036);
  const its=calculateIts(brut);
  const {exonere,motif,restants}=vpsExonereCheck(
    d.nationalite,d.premier_emploi,d.date_embauche,
    d.nouvelle_entreprise,d.date_premier_exercice,mois,annee);
  const vps=exonere?0:roundTo5(brut*0.04);
  const cnssPatronale=Math.round(brut*tauxPat);
  const salaireNet=brut-cnssOuvriere-its-taxeRadio;
  const remunerationDue=salaireNet-avances-saisie-assurance;
  return {
    salaire_brut:Math.round(brut), cnss_ouvriere:cnssOuvriere, its, taxe_radio:taxeRadio,
    salaire_net:Math.round(salaireNet), avances:Math.round(avances),
    saisie_arret:Math.round(saisie), assurance_sante:Math.round(assurance),
    remuneration_due:Math.round(remunerationDue), vps, vps_exonere:exonere,
    motif_exoneration:motif, mois_restants_exo:restants,
    cnss_patronale:cnssPatronale, taux_cnss_patronale:tauxPat,
    montant_heures_sup:hsup, h_sup_12:h12, h_sup_35:h35, h_sup_50:h50, h_sup_100:h100,
    montant_h12, montant_h35, montant_h50, montant_h100
  };
};

const fmt = v => new Intl.NumberFormat("fr-FR").format(Math.round(parseFloat(v)||0))+" FCFA";
const fmtN = v => new Intl.NumberFormat("fr-FR").format(Math.round(parseFloat(v)||0));

// ─── STYLES GLOBAUX ───────────────────────────────────────────────────────────
const G = {
  bg:"#f0f4ff", sidebar:"#1a3a6b", card:"#ffffff", border:"#d0dff5",
  accent:"#1a6dd6", accentDim:"#1558b0", red:"#e53e3e", yellow:"#d97706",
  text:"#1a2a4a", textDim:"#5a7a9a", input:"#e8f0fe",
  radius:"10px", shadow:"0 4px 24px rgba(26,61,107,0.12)"
};


// ─── COMPOSANTS UI ────────────────────────────────────────────────────────────

// ─── TOAST NOTIFICATIONS ─────────────────────────────────────────────────────
let _setToasts = null;
const toast = {
  success: (msg) => _setToasts && _setToasts(t=>[...t,{id:Date.now(),msg,type:"success"}]),
  error:   (msg) => _setToasts && _setToasts(t=>[...t,{id:Date.now(),msg,type:"error"}]),
  info:    (msg) => _setToasts && _setToasts(t=>[...t,{id:Date.now(),msg,type:"info"}]),
  warn:    (msg) => _setToasts && _setToasts(t=>[...t,{id:Date.now(),msg,type:"warn"}]),
};

const ToastContainer = () => {
  const [toasts,setToasts] = useState([]);
  useEffect(()=>{ _setToasts=setToasts; return ()=>{_setToasts=null;}; },[]);
  useEffect(()=>{
    if(!toasts.length) return;
    const t=setTimeout(()=>setToasts(ts=>ts.slice(1)),4000);
    return ()=>clearTimeout(t);
  },[toasts]);
  const colors={success:"#1a6dd6",error:"#e53e3e",info:"#1a3a6b",warn:"#d97706"};
  const icons={success:"✅",error:"❌",info:"ℹ️",warn:"⚠️"};
  return (
    <div style={{position:"fixed",bottom:"24px",right:"24px",zIndex:9999,
                 display:"flex",flexDirection:"column",gap:"8px",maxWidth:"340px"}}>
      {toasts.map(t=>(
        <div key={t.id} style={{background:"#fff",border:`2px solid ${colors[t.type]}`,
          borderRadius:"10px",padding:"12px 16px",
          boxShadow:"0 4px 20px rgba(26,61,107,0.18)",
          fontSize:"13px",color:"#1a2a4a",display:"flex",alignItems:"center",gap:"10px"}}>
          <span style={{fontSize:"18px"}}>{icons[t.type]}</span>
          <span>{t.msg}</span>
        </div>
      ))}
    </div>
  );
};

// ─── BOUTON FLOTTANT CONTACT ──────────────────────────────────────────────────
const FloatingContact = () => {
  const [open,setOpen] = useState(false);
  return (
    <div style={{position:"fixed",bottom:"28px",right:"28px",zIndex:9000,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:"10px"}}>
      {open && (
        <>
          <a href={`https://wa.me/${CONTACT_WHATSAPP}?text=Bonjour%2C%20j%27ai%20une%20question%20concernant%20RH-Paie%20Pro.`}
             target="_blank" rel="noopener noreferrer"
             style={{display:"flex",alignItems:"center",gap:"10px",background:"#25D366",
                     color:"#fff",padding:"10px 16px",borderRadius:"50px",boxShadow:"0 4px 18px rgba(37,211,102,0.45)",
                     textDecoration:"none",fontSize:"13px",fontWeight:700,fontFamily:"system-ui",
                     whiteSpace:"nowrap",animation:"rhSlideIn 0.2s ease"}}>
            <span style={{fontSize:"18px"}}>💬</span> WhatsApp
          </a>
          <a href={`mailto:${CONTACT_EMAIL}?subject=Question%20RH-Paie%20Pro`}
             style={{display:"flex",alignItems:"center",gap:"10px",background:"#1a6dd6",
                     color:"#fff",padding:"10px 16px",borderRadius:"50px",boxShadow:"0 4px 18px rgba(26,109,214,0.40)",
                     textDecoration:"none",fontSize:"13px",fontWeight:700,fontFamily:"system-ui",
                     whiteSpace:"nowrap",animation:"rhSlideIn 0.25s ease"}}>
            <span style={{fontSize:"18px"}}>✉️</span> Email
          </a>
        </>
      )}
      <button onClick={()=>setOpen(o=>!o)}
        style={{width:"52px",height:"52px",borderRadius:"50%",border:"none",cursor:"pointer",
                background:"linear-gradient(135deg,#1a3a6b,#1a6dd6)",
                boxShadow:"0 4px 20px rgba(26,61,107,0.4)",
                display:"flex",alignItems:"center",justifyContent:"center",
                fontSize:"22px",transition:"transform 0.2s",transform:open?"rotate(45deg)":"none"}}
        title="Nous contacter">
        {open ? "✕" : "💬"}
      </button>
      <style>{`@keyframes rhSlideIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}`}</style>
    </div>
  );
};

const Btn = ({children,onClick,variant="primary",small=false,type="button",disabled=false}) => {
  const bg=variant==="primary"?G.accent:variant==="danger"?G.red:variant==="ghost"?"transparent":"#e8f0fe";
  const col=variant==="ghost"?G.textDim:variant==="secondary"?G.accent:"#fff";
  const border=variant==="ghost"?`1px solid ${G.border}`:variant==="secondary"?`1px solid ${G.accent}`:"none";
  return (
    <button className="rh-btn" type={type} onClick={onClick} disabled={disabled}
      style={{background:bg,color:col,border,borderRadius:"8px",cursor:"pointer",
              padding:small?"6px 14px":"9px 20px",fontSize:small?"12px":"14px",
              fontWeight:600,fontFamily:"inherit",transition:"opacity 0.15s",
              opacity:disabled?0.5:1}}>
      {children}
    </button>
  );
};

const Input = ({label,value,onChange,type="text",placeholder="",required=false,small=false}) => (
  <div style={{marginBottom:small?0:"14px"}}>
    {label&&<label style={{display:"block",fontSize:"12px",color:"#1a3a6b",marginBottom:"5px",fontWeight:600}}>{label}</label>}
    <input type={type} value={value} onChange={e=>onChange(e.target.value)}
      placeholder={placeholder} required={required}
      style={{width:"100%",background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
              padding:"9px 12px",color:G.text,fontSize:"14px",fontFamily:"inherit",
              boxSizing:"border-box",outline:"none",boxShadow:"inset 0 1px 3px rgba(0,0,0,0.06)"}}/>
  </div>
);

const Select = ({label,value,onChange,options,small=false}) => (
  <div style={{marginBottom:small?0:"14px"}}>
    {label&&<label style={{display:"block",fontSize:"12px",color:"#1a3a6b",marginBottom:"5px",fontWeight:600}}>{label}</label>}
    <select className="rh-control" value={value} onChange={e=>onChange(e.target.value)}
      style={{width:"100%",background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
              padding:"9px 12px",color:G.text,fontSize:"14px",fontFamily:"inherit",outline:"none"}}>
      {options.map(o=><option key={o.value||o} value={o.value||o}>{o.label||o}</option>)}
    </select>
  </div>
);

const Card = ({children,style={}}) => (
  <div className="rh-card" style={{background:"#ffffff",border:"1px solid #d0dff5",borderRadius:"10px",
               padding:"24px",boxShadow:"0 2px 12px rgba(26,61,107,0.07)",...style}}>{children}</div>
);

const Badge = ({label,color=G.accent}) => (
  <span style={{background:color+"22",color,padding:"3px 10px",borderRadius:"20px",
                fontSize:"11px",fontWeight:600}}>{label}</span>
);

const Modal = ({title,onClose,children,width="500px"}) => {
  useEffect(()=>{
    if(typeof document==="undefined") return;
    document.body.classList.add("rh-no-scroll");
    return ()=>document.body.classList.remove("rh-no-scroll");
  },[]);

  return (
    <div className="rh-modal-backdrop" onClick={onClose}>
      <div className="rh-modal" style={{width}} onClick={e=>e.stopPropagation()}>
        <div className="rh-modal-header">
          <span style={{fontSize:"16px",fontWeight:700,color:G.text}}>{title}</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:G.textDim,
                  cursor:"pointer",fontSize:"20px",lineHeight:1}}>×</button>
        </div>
        <div className="rh-modal-body">{children}</div>
      </div>
    </div>
  );
};

const Alert = ({msg,type="success"}) => msg?(
  <div style={{padding:"12px 16px",borderRadius:"8px",marginBottom:"16px",fontSize:"13px",
               background:type==="success"?"rgba(255,255,255,0.15)":G.red+"22",
               color:type==="success"?G.accent:G.red,
               border:`1px solid ${type==="success"?G.accent:G.red}33`}}>{msg}</div>
):null;


const Table = ({cols,rows,onRow}) => (
  <div className="rh-table-wrap">
    <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
      <thead>
        <tr style={{borderBottom:"2px solid #1a6dd6"}}>
          {cols.map(c=><th key={c.key||c.label}
            style={{padding:"10px 12px",textAlign:"left",color:"#1a3a6b",
                    fontWeight:700,fontSize:"11px",textTransform:"uppercase",
                    letterSpacing:"0.5px"}}>{c.label}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.length===0&&(
          <tr><td colSpan={cols.length} style={{padding:"32px",textAlign:"center",color:G.textDim}}>
            Aucune donnée
          </td></tr>
        )}
        {rows.map((row,i)=>(
          <tr key={row.id||i} onClick={()=>onRow&&onRow(row)}
            style={{borderBottom:"1px solid #e8f0fe",cursor:onRow?"pointer":"default",
                    transition:"background 0.1s"}}
            onMouseEnter={e=>{if(onRow)e.currentTarget.style.background='#f0f4ff'}}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            {cols.map(c=>(
              <td key={c.key||c.label} style={{padding:"10px 12px",color:c.dim?G.textDim:G.text}}>
                {c.render?c.render(row):row[c.key]}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  </div>
);

// ─── SIDEBAR ──────────────────────────────────────────────────────────────────
const NAV = [
  {id:"dashboard",icon:"⊞",label:"Tableau de bord"},
  {id:"companies",icon:"🏢",label:"Sociétés"},
  {id:"employees",icon:"👥",label:"Employés"},
  {id:"payroll",icon:"💰",label:"Fiches de Paie"},
  {id:"historique",icon:"📜",label:"Historique"},
  {id:"declarations",icon:"📋",label:"Déclarations"},
  {id:"rapport",icon:"📊",label:"Rapport Cabinet"},
  {id:"settings",icon:"⚙",label:"Paramètres"},
];

const Sidebar = ({page,setPage,user,onLogout,open=false,setOpen=()=>{}}) => (
  <>
    <div className={`rh-overlay ${open?"open":""}`} onClick={()=>setOpen(false)}/>
    <aside className={`rh-sidebar ${open?"open":""}`} style={{background:G.sidebar,borderRight:`1px solid ${G.border}`}}>
      <div style={{padding:"24px 20px",borderBottom:`1px solid ${G.border}`,position:"relative"}}>
        <button className="rh-mobile-close" onClick={()=>setOpen(false)}
          style={{position:"absolute",top:"14px",right:"14px",background:"none",border:"none",color:"#fff",fontSize:"24px",cursor:"pointer",lineHeight:1}}>×</button>
        <div className="rh-sidebar-brand-text" style={{fontSize:"18px",fontWeight:800,color:"#ffffff",letterSpacing:"-0.5px"}}>RH-Paie Pro</div>
        <div className="rh-sidebar-brand-text" style={{fontSize:"11px",color:"#a0c0e8",marginTop:"3px"}}>Gestion de la Paie</div>
      </div>
      <nav style={{flex:1,padding:"16px 12px",overflowY:"auto"}}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>{setPage(n.id);setOpen(false);}}
            style={{width:"100%",display:"flex",alignItems:"center",gap:"10px",padding:"10px 12px",
                    borderRadius:"8px",border:"none",cursor:"pointer",textAlign:"left",
                    fontFamily:"inherit",fontSize:"13px",fontWeight:500,marginBottom:"3px",
                    background:page===n.id?"rgba(255,255,255,0.2)":"transparent",
                    color:page===n.id?"#ffffff":"#a0c0e8",transition:"all 0.15s"}}>
            <span style={{fontSize:"15px",minWidth:"18px",textAlign:"center"}}>{n.icon}</span>
            <span className="rh-sidebar-label">{n.label}</span>
          </button>
        ))}
      </nav>
      <div style={{padding:"16px",borderTop:"1px solid rgba(255,255,255,0.15)"}}>
        <div style={{fontSize:"12px",color:G.textDim,marginBottom:"8px",
                     overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
          {user?.email}
        </div>
        <Btn onClick={()=>{setOpen(false);onLogout();}} variant="ghost" small>Déconnexion</Btn>
      </div>
    </aside>
  </>
);
// ─── PAGES ────────────────────────────────────────────────────────────────────

// LOGIN + LANDING PAGE (avec avis et contact)
const LoginPage = ({onLogin}) => {
  const [email,setEmail]=useState(""), [pass,setPass]=useState(""),
        [err,setErr]=useState(""), [loading,setLoading]=useState(false),
        [isSignup,setIsSignup]=useState(false);

  // ── Avis utilisateurs ──
  const [reviews,setReviews]=useState([]);
  const [showReviewForm,setShowReviewForm]=useState(false);
  const [rvNom,setRvNom]=useState(""), [rvEntreprise,setRvEntreprise]=useState(""),
        [rvAvis,setRvAvis]=useState(""), [rvNote,setRvNote]=useState(5),
        [rvMsg,setRvMsg]=useState(""), [rvLoading,setRvLoading]=useState(false);

  useEffect(()=>{
    (async()=>{
      const {data}=await supabase.from("reviews").select("*")
        .eq("approved",true).order("created_at",{ascending:false}).limit(6);
      setReviews(data||[]);
    })();
  },[]);

  const submitReview = async () => {
    if(!rvNom.trim()||!rvAvis.trim()){setRvMsg("Veuillez remplir votre nom et votre avis.");return;}
    setRvLoading(true); setRvMsg("");
    const {error}=await supabase.from("reviews").insert({
      nom:rvNom.trim(), entreprise:rvEntreprise.trim()||null,
      avis:rvAvis.trim(), note:rvNote, approved:true
    });
    setRvLoading(false);
    if(error){setRvMsg("Erreur : "+error.message);}
    else{
      setRvMsg("✅ Merci pour votre avis !");
      setRvNom(""); setRvEntreprise(""); setRvAvis(""); setRvNote(5);
      setTimeout(()=>{setShowReviewForm(false);setRvMsg("");},2000);
      const {data}=await supabase.from("reviews").select("*")
        .eq("approved",true).order("created_at",{ascending:false}).limit(6);
      setReviews(data||[]);
    }
  };

  const submit = async e => {
    e.preventDefault(); setErr(""); setLoading(true);
    try{
      const {error} = isSignup
        ? await supabase.auth.signUp({email,password:pass})
        : await supabase.auth.signInWithPassword({email,password:pass});
      if(error) setErr(error.message);
      else { if(isSignup) setErr("Vérifiez votre email pour confirmer le compte."); else onLogin(); }
    }finally{setLoading(false);}
  };

  const stars = n => "★".repeat(n)+"☆".repeat(5-n);

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,#0f2557 0%,#1a3a6b 40%,#1a6dd6 100%)",
                 fontFamily:"system-ui,sans-serif",color:"#fff"}}>

      {/* ── HERO + LOGIN ── */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",
                   padding:"48px 20px 32px",flexWrap:"wrap",gap:"48px"}}>

        {/* Bloc texte gauche */}
        <div style={{maxWidth:"440px",flex:"1 1 280px"}}>
          <div style={{fontSize:"13px",background:"rgba(255,255,255,0.12)",display:"inline-block",
                       padding:"4px 14px",borderRadius:"20px",marginBottom:"16px",fontWeight:600}}>
            🇧🇯 Solution 100% adaptée au Bénin & UEMOA
          </div>
          <div style={{fontSize:"40px",fontWeight:900,letterSpacing:"-1.5px",lineHeight:1.15,marginBottom:"16px"}}>
            RH-Paie Pro
          </div>
          <div style={{fontSize:"16px",color:"#a0c8f0",lineHeight:1.6,marginBottom:"28px"}}>
            Gérez la paie de vos employés en toute conformité — CNSS, VPS, ITS — simplement et rapidement.
          </div>
          <div style={{display:"flex",gap:"16px",flexWrap:"wrap"}}>
            {[["⚡","Calcul automatique","ITS, CNSS, VPS"],["📄","Bulletins PDF","3 modèles inclus"],["🏢","Multi-sociétés","Toutes vos entités"]].map(([ic,t,s])=>(
              <div key={t} style={{background:"rgba(255,255,255,0.1)",borderRadius:"10px",padding:"12px 16px",flex:"1 1 100px"}}>
                <div style={{fontSize:"22px",marginBottom:"4px"}}>{ic}</div>
                <div style={{fontWeight:700,fontSize:"13px"}}>{t}</div>
                <div style={{color:"#a0c8f0",fontSize:"11px"}}>{s}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Carte login droite */}
        <div style={{width:"100%",maxWidth:"380px"}}>
          <Card>
            <div style={{textAlign:"center",marginBottom:"20px"}}>
              <div style={{fontSize:"20px",fontWeight:800,color:"#1a3a6b"}}>
                {isSignup ? "Créer un compte" : "Se connecter"}
              </div>
              <div style={{fontSize:"12px",color:"#5a7a9a",marginTop:"4px"}}>14 jours gratuits — sans carte bancaire</div>
            </div>
            <Alert msg={err} type={err.includes("rifi") ? "success" : "error"}/>
            <Input label="Email" value={email} onChange={setEmail} type="email" required/>
            <Input label="Mot de passe" value={pass} onChange={setPass} type="password" required/>
            <div style={{marginTop:"8px"}}>
              <Btn onClick={submit} disabled={loading}>
                {loading?"...":(isSignup?"Créer le compte":"Se connecter")}
              </Btn>
            </div>
            <div style={{marginTop:"16px",textAlign:"center"}}>
              <button onClick={()=>{setIsSignup(!isSignup);setErr("");}}
                style={{background:"none",border:"none",color:G.accent,cursor:"pointer",
                        fontSize:"13px",fontFamily:"inherit"}}>
                {isSignup?"Déjà un compte ? Se connecter":"Pas de compte ? S'inscrire"}
              </button>
            </div>
          </Card>
        </div>
      </div>

      {/* ── SECTION AVIS ── */}
      <div style={{background:"rgba(255,255,255,0.06)",borderTop:"1px solid rgba(255,255,255,0.1)",
                   padding:"48px 24px"}}>
        <div style={{maxWidth:"1000px",margin:"0 auto"}}>
          <div style={{textAlign:"center",marginBottom:"32px"}}>
            <div style={{fontSize:"28px",fontWeight:800,marginBottom:"8px"}}>Ce que disent nos utilisateurs</div>
            <div style={{color:"#a0c8f0",fontSize:"14px"}}>Ils font confiance à RH-Paie Pro pour leur gestion de la paie</div>
          </div>

          {/* Grille avis */}
          {reviews.length > 0 ? (
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:"16px",marginBottom:"32px"}}>
              {reviews.map(r=>(
                <div key={r.id} style={{background:"rgba(255,255,255,0.1)",borderRadius:"12px",padding:"20px",
                                        backdropFilter:"blur(4px)",border:"1px solid rgba(255,255,255,0.15)"}}>
                  <div style={{color:"#fbbf24",fontSize:"16px",marginBottom:"10px"}}>{stars(r.note||5)}</div>
                  <div style={{fontSize:"14px",lineHeight:1.6,marginBottom:"14px",color:"#e0ecff",fontStyle:"italic"}}>
                    "{r.avis}"
                  </div>
                  <div style={{borderTop:"1px solid rgba(255,255,255,0.15)",paddingTop:"10px"}}>
                    <div style={{fontWeight:700,fontSize:"13px"}}>{r.nom}</div>
                    {r.entreprise&&<div style={{fontSize:"11px",color:"#a0c8f0"}}>{r.entreprise}</div>}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div style={{textAlign:"center",color:"#a0c8f0",marginBottom:"24px",fontSize:"14px"}}>
              Soyez le premier à laisser un avis !
            </div>
          )}

          {/* Bouton / formulaire avis */}
          <div style={{textAlign:"center"}}>
            {!showReviewForm ? (
              <button onClick={()=>setShowReviewForm(true)}
                style={{background:"rgba(255,255,255,0.15)",border:"2px solid rgba(255,255,255,0.4)",
                        color:"#fff",padding:"12px 28px",borderRadius:"50px",cursor:"pointer",
                        fontSize:"14px",fontWeight:700,fontFamily:"system-ui",
                        transition:"all 0.2s"}}>
                ✍️ Laisser un avis
              </button>
            ) : (
              <div style={{background:"rgba(255,255,255,0.12)",borderRadius:"16px",padding:"28px",
                           maxWidth:"480px",margin:"0 auto",border:"1px solid rgba(255,255,255,0.2)"}}>
                <div style={{fontSize:"16px",fontWeight:700,marginBottom:"20px"}}>Votre avis</div>
                {rvMsg && (
                  <div style={{background:rvMsg.includes("✅")?"rgba(34,197,94,0.2)":"rgba(239,68,68,0.2)",
                               border:"1px solid",borderColor:rvMsg.includes("✅")?"#22c55e":"#ef4444",
                               borderRadius:"8px",padding:"10px 14px",fontSize:"13px",marginBottom:"14px"}}>
                    {rvMsg}
                  </div>
                )}
                <div style={{marginBottom:"12px"}}>
                  <label style={{display:"block",fontSize:"12px",fontWeight:600,marginBottom:"5px",color:"#a0c8f0"}}>Votre nom *</label>
                  <input value={rvNom} onChange={e=>setRvNom(e.target.value)} placeholder="Ex: Koffi Adjovi"
                    style={{width:"100%",padding:"9px 12px",borderRadius:"8px",border:"1px solid rgba(255,255,255,0.25)",
                            background:"rgba(255,255,255,0.1)",color:"#fff",fontSize:"14px",fontFamily:"system-ui",
                            outline:"none",boxSizing:"border-box"}}/>
                </div>
                <div style={{marginBottom:"12px"}}>
                  <label style={{display:"block",fontSize:"12px",fontWeight:600,marginBottom:"5px",color:"#a0c8f0"}}>Entreprise (optionnel)</label>
                  <input value={rvEntreprise} onChange={e=>setRvEntreprise(e.target.value)} placeholder="Ex: PSARIZ SARL"
                    style={{width:"100%",padding:"9px 12px",borderRadius:"8px",border:"1px solid rgba(255,255,255,0.25)",
                            background:"rgba(255,255,255,0.1)",color:"#fff",fontSize:"14px",fontFamily:"system-ui",
                            outline:"none",boxSizing:"border-box"}}/>
                </div>
                <div style={{marginBottom:"12px"}}>
                  <label style={{display:"block",fontSize:"12px",fontWeight:600,marginBottom:"5px",color:"#a0c8f0"}}>Note</label>
                  <div style={{display:"flex",gap:"6px"}}>
                    {[1,2,3,4,5].map(n=>(
                      <button key={n} onClick={()=>setRvNote(n)}
                        style={{background:"none",border:"none",cursor:"pointer",fontSize:"24px",
                                color:n<=rvNote?"#fbbf24":"rgba(255,255,255,0.3)",transition:"color 0.15s"}}>★</button>
                    ))}
                  </div>
                </div>
                <div style={{marginBottom:"16px"}}>
                  <label style={{display:"block",fontSize:"12px",fontWeight:600,marginBottom:"5px",color:"#a0c8f0"}}>Votre avis *</label>
                  <textarea value={rvAvis} onChange={e=>setRvAvis(e.target.value)}
                    placeholder="Partagez votre expérience avec RH-Paie Pro..." rows={4}
                    style={{width:"100%",padding:"9px 12px",borderRadius:"8px",border:"1px solid rgba(255,255,255,0.25)",
                            background:"rgba(255,255,255,0.1)",color:"#fff",fontSize:"14px",fontFamily:"system-ui",
                            outline:"none",resize:"vertical",boxSizing:"border-box"}}/>
                </div>
                <div style={{display:"flex",gap:"10px"}}>
                  <button onClick={submitReview} disabled={rvLoading}
                    style={{flex:1,background:"#1a6dd6",color:"#fff",border:"none",borderRadius:"8px",
                            padding:"10px",cursor:"pointer",fontWeight:700,fontSize:"14px",fontFamily:"system-ui",
                            opacity:rvLoading?0.6:1}}>
                    {rvLoading?"Envoi...":"Publier l'avis"}
                  </button>
                  <button onClick={()=>{setShowReviewForm(false);setRvMsg("");}}
                    style={{background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px solid rgba(255,255,255,0.25)",
                            borderRadius:"8px",padding:"10px 16px",cursor:"pointer",fontSize:"14px",fontFamily:"system-ui"}}>
                    Annuler
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ── FOOTER ── */}
      <div style={{textAlign:"center",padding:"24px",color:"rgba(255,255,255,0.5)",fontSize:"12px",
                   borderTop:"1px solid rgba(255,255,255,0.08)"}}>
        © {new Date().getFullYear()} RH-Paie Pro — Bénin &nbsp;|&nbsp;
        <a href={`https://wa.me/${CONTACT_WHATSAPP}`} target="_blank" rel="noopener noreferrer"
          style={{color:"#25D366",textDecoration:"none",fontWeight:600}}>💬 WhatsApp</a>
        &nbsp;|&nbsp;
        <a href={`mailto:${CONTACT_EMAIL}`} style={{color:"#63a8f0",textDecoration:"none",fontWeight:600}}>✉️ Email</a>
      </div>

      {/* Bouton flottant sur la page de login aussi */}
      <FloatingContact/>
    </div>
  );
};

// DASHBOARD
const Dashboard = ({companies}) => {
  const [stats,setStats]=useState({employees:0,payrolls:0,monthlyCost:0});
  useEffect(()=>{
    (async()=>{
      const ids=companies.map(c=>c.id);
      if(!ids.length) return;
      const {count:ec}=await supabase.from("employees").select("*",{count:"exact",head:true}).in("company_id",ids);
      const now=new Date();
      const {data:pData}=await supabase.from("payrolls").select("remuneration_due")
        .in("company_id",ids).eq("mois",now.getMonth()+1).eq("annee",now.getFullYear());
      const total=(pData||[]).reduce((s,p)=>s+parseFloat(p.remuneration_due||0),0);
      setStats({employees:ec||0,payrolls:(pData||[]).length,monthlyCost:total});
    })();
  },[companies]);
  const cards=[
    {label:"Sociétés",value:companies.length,icon:"🏢",color:G.accent},
    {label:"Employés actifs",value:stats.employees,icon:"👥",color:"#6366f1"},
    {label:"Fiches ce mois",value:stats.payrolls,icon:"📄",color:G.yellow},
    {label:"Masse salariale",value:fmt(stats.monthlyCost),icon:"💰",color:"#f472b6"},
  ];
  return (
    <div>
      <h2 style={{marginBottom:"24px",fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Tableau de bord</h2>
      <div className="rh-grid-4">
        {cards.map(c=>(
          <Card key={c.label}>
            <div style={{fontSize:"24px",marginBottom:"8px"}}>{c.icon}</div>
            <div style={{fontSize:"24px",fontWeight:800,color:c.color}}>{c.value}</div>
            <div style={{fontSize:"12px",color:G.textDim,marginTop:"4px"}}>{c.label}</div>
          </Card>
        ))}
      </div>
      <Card>
        <div style={{color:G.textDim,fontSize:"13px"}}>
          Bienvenue dans <strong style={{color:G.text}}>RH-Paie Pro</strong> — Votre solution de gestion de la paie conforme aux règles CNSS, VPS et ITS du Bénin.
          <br/><br/>Commencez par créer votre <strong style={{color:G.accent}}>société</strong>, ajoutez vos <strong style={{color:G.accent}}>employés</strong>, puis générez les <strong style={{color:G.accent}}>fiches de paie</strong>.
        </div>
      </Card>
    </div>
  );
};

// COMPANIES
// ─── LOGO UPLOAD ─────────────────────────────────────────────────────────────
const LogoUpload = ({value, onChange}) => {
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState("");
  const upload = async e => {
    const file = e.target.files[0];
    if(!file) return;
    if(file.size > 1024*1024) { setError("Fichier trop grand (max 1 Mo)"); return; }
    setUploading(true); setError("");
    try {
      const ext = file.name.split(".").pop();
      const path = `logos/${Date.now()}.${ext}`;
      const {error:upErr} = await supabase.storage.from("logo").upload(path, file,
        {cacheControl:"3600", upsert:true});
      if(upErr) throw upErr;
      const {data} = supabase.storage.from("logo").getPublicUrl(path);
      onChange(data.publicUrl);
    } catch(err) {
      setError(err.message||"Erreur upload");
    } finally { setUploading(false); }
  };
  return (
    <div style={{marginBottom:"14px"}}>
      <label style={{display:"block",fontSize:"12px",color:"#1a3a6b",marginBottom:"5px",fontWeight:600}}>
        Logo entreprise (optionnel, max 1 Mo)
      </label>
      {value&&(
        <div className="rh-logo-preview">
          <img src={value} alt="Logo" onError={e=>e.target.style.display="none"}
            style={{height:"48px",borderRadius:"6px",objectFit:"contain",background:"#fff",padding:"4px"}}/>
          <button onClick={()=>onChange("")}
            style={{background:"none",border:"none",color:G.red,cursor:"pointer",fontSize:"12px"}}>
            Supprimer
          </button>
        </div>
      )}
      <label style={{display:"inline-block",cursor:"pointer",background:"#fff",
                     border:`1px solid ${G.border}`,borderRadius:"8px",
                     padding:"9px 16px",fontSize:"13px",color:G.textDim}}>
        {uploading?"Envoi en cours...":"📁 Choisir un fichier"}
        <input type="file" accept="image/*" onChange={upload} style={{display:"none"}} disabled={uploading}/>
      </label>
      {error&&<div style={{color:G.red,fontSize:"12px",marginTop:"4px"}}>{error}</div>}
    </div>
  );
};

// ─── COMPANIES
const Companies = ({companies,reload,userId}) => {
  const [modal,setModal]=useState(null);
  const [form,setForm]=useState({raison_sociale:"",rccm:"",adresse:"",tel:"",email:"",
    nouvelle_entreprise:false,date_premier_exercice:"",taux_cnss_patronale:"0.194",
    cnss_employeur:"",ifu:"",
    signataire:"",fonction_signataire:""});
  const [msg,setMsg]=useState("");
  const open=(c=null)=>{
    setForm(c?{...c,taux_cnss_patronale:String(c.taux_cnss_patronale||0.194),
      nouvelle_entreprise:Boolean(c.nouvelle_entreprise)}:
      {raison_sociale:"",rccm:"",adresse:"",tel:"",email:"",nouvelle_entreprise:false,
       date_premier_exercice:"",taux_cnss_patronale:"0.194",cnss_employeur:"",ifu:"",
       signataire:"",fonction_signataire:""});
    setModal(c||"new"); setMsg("");
  };
  const save = async () => {
    const payload={...form,user_id:userId,taux_cnss_patronale:parseFloat(form.taux_cnss_patronale||0.194)};
    const {error}=modal==="new"
      ? await supabase.from("companies").insert(payload)
      : await supabase.from("companies").update(payload).eq("id",modal.id);
    if(error) setMsg(error.message);
    else{ setModal(null); reload(); }
  };
  const del = async id => {
    if(!confirm("Supprimer cette société et toutes ses données ?")) return;
    await supabase.from("companies").delete().eq("id",id);
    reload();
  };
  const f=(k,v)=>setForm(p=>({...p,[k]:v}));
  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Sociétés</h2>
        <div className="rh-btn-group"><Btn onClick={()=>open()}>+ Nouvelle société</Btn></div>
      </div>
      <Card>
        <Table cols={[
          {key:"raison_sociale",label:"Raison sociale"},
          {key:"rccm",label:"RCCM",dim:true},
          {key:"tel",label:"Téléphone",dim:true},
          {label:"CNSS Pat.",render:r=>`${(parseFloat(r.taux_cnss_patronale||0.194)*100).toFixed(1)}%`},
          {label:"Actions",render:r=>(
            <div className="rh-btn-group">
              <Btn small onClick={e=>{e.stopPropagation();open(r)}} variant="secondary">Modifier</Btn>
              <Btn small onClick={e=>{e.stopPropagation();del(r.id)}} variant="danger">Suppr.</Btn>
            </div>
          )}
        ]} rows={companies}/>
      </Card>
      {modal&&(
        <Modal title={modal==="new"?"Nouvelle société":"Modifier société"} onClose={()=>setModal(null)}>
          <Alert msg={msg} type="error"/>
          <Input label="Raison sociale *" value={form.raison_sociale} onChange={v=>f("raison_sociale",v)} required/>
          <Input label="RCCM" value={form.rccm} onChange={v=>f("rccm",v)}/>
          <Input label="Adresse" value={form.adresse} onChange={v=>f("adresse",v)}/>
          <Input label="Téléphone" value={form.tel} onChange={v=>f("tel",v)}/>
          <Input label="Email" value={form.email} onChange={v=>f("email",v)} type="email"/>
          <Input label="Taux CNSS patronale (ex: 0.194)" value={form.taux_cnss_patronale} onChange={v=>f("taux_cnss_patronale",v)}/>
          <Input label="N° CNSS Employeur" value={form.cnss_employeur||""} onChange={v=>f("cnss_employeur",v)} placeholder="ex: RB/NAT/2020-B-321"/>
          <Input label="N° IFU" value={form.ifu||""} onChange={v=>f("ifu",v)} placeholder="ex: 3201641471613"/>
          <div style={{gridColumn:"1/-1"}}><LogoUpload value={form.logo_url||""} onChange={v=>f("logo_url",v)}/></div>
          <Input label="Signataire" value={form.signataire} onChange={v=>f("signataire",v)}/>
          <Input label="Fonction signataire" value={form.fonction_signataire} onChange={v=>f("fonction_signataire",v)}/>
          <div style={{marginBottom:"14px"}}>
            <label style={{display:"flex",alignItems:"center",gap:"8px",color:G.textDim,fontSize:"13px",cursor:"pointer"}}>
              <input type="checkbox" checked={form.nouvelle_entreprise} onChange={e=>f("nouvelle_entreprise",e.target.checked)}/>
              Nouvelle entreprise (exonération VPS art.192)
            </label>
          </div>
          {form.nouvelle_entreprise&&(
            <Input label="Date premier exercice" value={form.date_premier_exercice} onChange={v=>f("date_premier_exercice",v)} type="date"/>
          )}
          <div className="rh-actions-row" style={{marginTop:"8px"}}>
            <Btn onClick={save}>Enregistrer</Btn>
            <Btn onClick={()=>setModal(null)} variant="ghost">Annuler</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

// EMPLOYEES
const Employees = ({companies}) => {
  const [employees,setEmployees]=useState([]);
  const [compId,setCompId]=useState(companies[0]?.id||"");
  const [modal,setModal]=useState(null);
  const [msg,setMsg]=useState("");
  const initForm={matricule:"",email:"",nom:"",prenoms:"",emploi:"",categorie:"",
    situation_matrimoniale:"Célibataire",nb_enfants:"0",cnss:"",ifu:"",
    salaire_base:"",date_embauche:"",nationalite:"Béninoise",premier_emploi:false,email:"",actif:true};
  const [form,setForm]=useState(initForm);
  const load = useCallback(async()=>{
    if(!compId) return;
    const {data}=await supabase.from("employees").select("*").eq("company_id",compId).order("nom");
    setEmployees(data||[]);
  },[compId]);
  useEffect(()=>{load();},[load]);
  useEffect(()=>{if(!compId&&companies.length) setCompId(companies[0].id);},[companies]);
  const open=(e=null)=>{ setForm(e?{...e,nb_enfants:String(e.nb_enfants||0),
    salaire_base:String(e.salaire_base||""),premier_emploi:Boolean(e.premier_emploi)}:initForm);
    setModal(e||"new"); setMsg(""); };
  const save = async()=>{
    const p={...form,company_id:compId,nb_enfants:parseInt(form.nb_enfants||0),
      salaire_base:parseFloat(form.salaire_base||0)};
    const {error}=modal==="new"
      ?await supabase.from("employees").insert(p)
      :await supabase.from("employees").update(p).eq("id",modal.id);
    if(error) setMsg(error.message); else{setModal(null);load();}
  };
  const del = async id=>{
    if(!confirm("Supprimer cet employé ?")) return;
    await supabase.from("employees").delete().eq("id",id); load();
  };
  const f=(k,v)=>setForm(p=>({...p,[k]:v}));
  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Employés</h2>
        <div className="rh-inline-actions">
          <select className="rh-control" value={compId} onChange={e=>setCompId(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {companies.map(c=><option key={c.id} value={c.id}>{c.raison_sociale}</option>)}
          </select>
          <Btn onClick={()=>open()}>+ Ajouter</Btn>
        </div>
      </div>
      <Card>
        <Table cols={[
          {key:"matricule",label:"Matricule",dim:true},
          {label:"Nom complet",render:r=>`${r.nom} ${r.prenoms}`},
          {key:"emploi",label:"Emploi"},
          {key:"categorie",label:"Catégorie",dim:true},
          {label:"Salaire base",render:r=>fmt(r.salaire_base)},
          {label:"Statut",render:r=><Badge label={r.actif?"Actif":"Inactif"} color={r.actif?G.accent:G.red}/>},
          {label:"Actions",render:r=>(
            <div className="rh-btn-group">
              <Btn small onClick={e=>{e.stopPropagation();open(r)}} variant="secondary">Modifier</Btn>
              <Btn small onClick={e=>{e.stopPropagation();del(r.id)}} variant="danger">Suppr.</Btn>
            </div>
          )}
        ]} rows={employees}/>
      </Card>
      {modal&&(
        <Modal title={modal==="new"?"Nouvel employé":"Modifier employé"} onClose={()=>setModal(null)} width="600px">
          <Alert msg={msg} type="error"/>
          <div className="rh-form-grid-2">
            <Input label="Matricule" value={form.matricule} onChange={v=>f("matricule",v)}/>
            <Input label="Nom *" value={form.nom} onChange={v=>f("nom",v)} required/>
            <Input label="Prénoms *" value={form.prenoms} onChange={v=>f("prenoms",v)} required/>
            <Input label="Emploi" value={form.emploi} onChange={v=>f("emploi",v)}/>
            <Input label="Catégorie" value={form.categorie} onChange={v=>f("categorie",v)}/>
            <Select label="Situation matrimoniale" value={form.situation_matrimoniale}
              onChange={v=>f("situation_matrimoniale",v)}
              options={["Célibataire","Marié(e)","Divorcé(e)","Veuf/Veuve"].map(x=>({value:x,label:x}))}/>
            <Input label="Nb enfants" value={form.nb_enfants} onChange={v=>f("nb_enfants",v)} type="number"/>
            <Input label="N° CNSS" value={form.cnss} onChange={v=>f("cnss",v)}/>
            <Input label="N° IFU" value={form.ifu} onChange={v=>f("ifu",v)}/>
            <Input label="Salaire de base (FCFA)" value={form.salaire_base} onChange={v=>f("salaire_base",v)} type="number"/>
            <Input label="Date d'embauche" value={form.date_embauche} onChange={v=>f("date_embauche",v)} type="date"/>
            <Input label="Nationalité" value={form.nationalite} onChange={v=>f("nationalite",v)}/>
            <Input label="Email" value={form.email} onChange={v=>f("email",v)} type="email"/>
          </div>
          <div className="rh-checkbox-row">
            <label style={{display:"flex",alignItems:"center",gap:"8px",color:G.textDim,fontSize:"13px",cursor:"pointer"}}>
              <input type="checkbox" checked={form.premier_emploi} onChange={e=>f("premier_emploi",e.target.checked)}/>
              Premier emploi
            </label>
            <label style={{display:"flex",alignItems:"center",gap:"8px",color:G.textDim,fontSize:"13px",cursor:"pointer"}}>
              <input type="checkbox" checked={form.actif} onChange={e=>f("actif",e.target.checked)}/>
              Actif
            </label>
          </div>
          <div className="rh-actions-row">
            <Btn onClick={save}>Enregistrer</Btn>
            <Btn onClick={()=>setModal(null)} variant="ghost">Annuler</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

// PAYROLL
const Payroll = ({companies}) => {
  const [payrolls,setPayrolls]=useState([]);
  const [employees,setEmployees]=useState([]);
  const [compId,setCompId]=useState(companies[0]?.id||"");
  const [filterMois,setFilterMois]=useState(String(new Date().getMonth()+1));
  const [filterAnnee,setFilterAnnee]=useState(String(CURRENT_YEAR));
  const [modal,setModal]=useState(null);
  const [viewModal,setViewModal]=useState(null);
  const [msg,setMsg]=useState("");
  const [loading,setLoading]=useState(false);
  const [modele,setModele]=useState('officiel');
  const comp=companies.find(c=>c.id===compId);
  const initForm={employee_id:"",mois:String(new Date().getMonth()+1),annee:String(CURRENT_YEAR),
    salaire_base:"",h_sup_12:"0",h_sup_35:"0",h_sup_50:"0",h_sup_100:"0",
    primes:"0",indemnites:"0",avances:"0",saisie_arret:"0",assurance_sante:"0",notes:""};
  const [form,setForm]=useState(initForm);
  const [calc,setCalc]=useState(null);

  const loadEmployees = useCallback(async()=>{
    if(!compId) return;
    const {data}=await supabase.from("employees").select("*").eq("company_id",compId).eq("actif",true).order("nom");
    setEmployees(data||[]);
  },[compId]);
  const load = useCallback(async()=>{
    if(!compId) return;
    const {data}=await supabase.from("payrolls")
      .select("*, employees(nom,prenoms,emploi,matricule,email,cnss,ifu,situation_matrimoniale,nb_enfants)")
      .eq("company_id",compId).eq("mois",parseInt(filterMois)).eq("annee",parseInt(filterAnnee))
      .order("created_at");
    setPayrolls(data||[]);
  },[compId,filterMois,filterAnnee]);
  useEffect(()=>{loadEmployees();load();},[loadEmployees,load]);
  useEffect(()=>{if(!compId&&companies.length) setCompId(companies[0].id);},[companies]);

  const doCalc = useCallback(()=>{
    if(!form.employee_id||!form.salaire_base) return;
    const emp=employees.find(e=>e.id===form.employee_id);
    if(!emp) return;
    const d={...form,...comp,
      nationalite:emp.nationalite,premier_emploi:emp.premier_emploi,
      date_embauche:emp.date_embauche,
      nouvelle_entreprise:comp?.nouvelle_entreprise,
      date_premier_exercice:comp?.date_premier_exercice,
      h_sup_12:form.h_sup_12||0, h_sup_35:form.h_sup_35||0,
      h_sup_50:form.h_sup_50||0, h_sup_100:form.h_sup_100||0};
    setCalc(calculatePayroll(d));
  },[form,employees,comp]);
  useEffect(()=>{doCalc();},[doCalc]);

  const f=(k,v)=>{
    const base=k==="employee_id"?employees.find(e=>e.id===v)?.salaire_base||"":"";
    setForm(p=>({...p,[k]:v,...(k==="employee_id"?{salaire_base:String(base)}:{})}));
  };

  const save = async()=>{
    if(!calc) return;
    setLoading(true);
    const {motif_exoneration, mois_restants_exo, ...calcClean} = calc;
    const payload={...form,...calcClean,company_id:compId,
      employee_id:form.employee_id,mois:parseInt(form.mois),annee:parseInt(form.annee),
      salaire_base:parseFloat(form.salaire_base||0),montant_heures_sup:parseFloat(form.montant_heures_sup||0),
      primes:parseFloat(form.primes||0),indemnites:parseFloat(form.indemnites||0)};
    // Vérifier si fiche existe déjà
    const {data:exist}=await supabase.from("payrolls").select("id")
      .eq("employee_id",form.employee_id).eq("mois",parseInt(form.mois)).eq("annee",parseInt(form.annee)).maybeSingle();
    if(exist) toast.info("ℹ️ Fiche déjà existante — mise à jour effectuée");
    const {error}=await supabase.from("payrolls").upsert(payload,{onConflict:"employee_id,mois,annee"});
    setLoading(false);
    if(error){ setMsg(error.message); toast.error("Erreur : "+error.message); }
    else{ setModal(null); toast.success("✅ Fiche de paie enregistrée avec succès !"); load(); }
  };

  const genererTout=async()=>{
    if(!employees.length){toast.warn("Aucun employé actif pour cette société.");return;}
    if(!confirm(`Générer et imprimer les fiches de ${MOIS[parseInt(filterMois)]} ${filterAnnee} pour ${employees.length} employé(s) ?`)) return;
    setLoading(true);
    let ok=0,err=0,deja=0;
    const fiches=[];

    for(const emp of employees){
      // Vérifier si fiche déjà existante
      const {data:exist}=await supabase.from("payrolls").select("id")
        .eq("employee_id",emp.id).eq("mois",parseInt(filterMois)).eq("annee",parseInt(filterAnnee)).maybeSingle();
      if(exist){
        deja++;
        const {data:ficheExist}=await supabase.from("payrolls").select("*, employees(*)")
          .eq("id",exist.id).maybeSingle();
        if(ficheExist) fiches.push({p:ficheExist,emp:ficheExist.employees||emp});
        continue;
      }
      const d={...emp,...comp,employee_id:emp.id,mois:filterMois,annee:filterAnnee,
               h_sup_12:0,h_sup_35:0,h_sup_50:0,h_sup_100:0,
               primes:0,indemnites:0,avances:0,saisie_arret:0,assurance_sante:0};
      const c=calculatePayroll(d);
      const {motif_exoneration,mois_restants_exo,...calcClean}=c;
      const payload={...calcClean,company_id:compId,employee_id:emp.id,
        mois:parseInt(filterMois),annee:parseInt(filterAnnee),
        salaire_base:parseFloat(emp.salaire_base||0),montant_heures_sup:0,
        h_sup_12:0,h_sup_35:0,h_sup_50:0,h_sup_100:0,
        montant_h12:0,montant_h35:0,montant_h50:0,montant_h100:0,
        primes:0,indemnites:0,avances:0,saisie_arret:0,assurance_sante:0};
      const {error}=await supabase.from("payrolls")
        .upsert(payload,{onConflict:"employee_id,mois,annee"}).select().maybeSingle();
      if(error) err++;
      else{ ok++; fiches.push({p:{...payload,...calcClean},emp}); }
    }

    setLoading(false);
    load();

    // Notifications
    if(deja>0) toast.info(`ℹ️ ${deja} fiche(s) déjà existante(s) — incluses`);
    if(err>0) toast.error(`${err} erreur(s) lors de la génération`);
    if(ok>0) toast.success(`✅ ${ok} nouvelle(s) fiche(s) générée(s)`);

    // Ouvrir TOUTES les fiches dans une seule fenêtre d'impression
    if(fiches.length>0){
      const compNom=(comp?.raison_sociale||"").toUpperCase();
      const periode=`${MOIS[parseInt(filterMois)]} ${filterAnnee}`;
      // Construire un seul HTML avec toutes les fiches
      const allHtml=fiches.map(({p,emp},i)=>{
        const ficheHtml=buildBulletinHTML(p,emp,comp,modele);
        // Extraire uniquement le contenu du body
        const bodyMatch=ficheHtml.match(/<body[^>]*>([\s\S]*)<\/body>/i);
        const bodyContent=bodyMatch?bodyMatch[1]:"";
        const pageBreak=i<fiches.length-1
          ?'<div style="page-break-after:always;"></div>':"";
        return `<div class="fiche-page">${bodyContent}</div>${pageBreak}`;
      }).join("");

      // Extraire le style de la première fiche
      const firstHtml=buildBulletinHTML(fiches[0].p,fiches[0].emp,comp,modele);
      const styleMatch=firstHtml.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
      const styles=styleMatch?styleMatch[1]:"";

      const win=window.open("","_blank");
      win.document.write(`<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<title>Bulletins de paie — ${compNom} — ${periode}</title>
<style>
${styles}
.fiche-page { page-break-inside: avoid; }
@media print {
  @page { margin: 10mm 12mm; size: A4 portrait; }
  .fiche-page { page-break-inside: avoid; }
}
</style>
</head><body>
<div style="text-align:center;font-family:Arial,sans-serif;font-size:12pt;
     padding:8px;background:#d6d6d6;border:2px solid #888;margin-bottom:12px;
     font-weight:bold;letter-spacing:1px;">
  ${compNom} — BULLETINS DE PAIE — ${periode.toUpperCase()} — ${fiches.length} EMPLOYÉ(S)
</div>
${allHtml}
<script>
  window.onload = function() {
    setTimeout(function(){ window.print(); }, 800);
  };
<\/script>
</body></html>`);
      win.document.close();
      toast.success(`🖨 Impression de ${fiches.length} fiche(s) lancée — choisissez "Microsoft Print to PDF" pour sauvegarder`);
    }
  };

  const envoyerTout=async()=>{
    if(!payrolls.length){alert("Aucune fiche à envoyer.");return;}
    const avecTel=payrolls.filter(p=>p.employees?.email);
    if(!avecTel.length){
      alert("⚠️ Aucun employé n'a de numéro WhatsApp renseigné.\nRenseignez le numéro dans le champ 'Email' de chaque employé (ex: 2290197000000).");
      return;
    }
    const {data:cfg}=await supabase.from("email_settings").select("*")
      .eq("user_id",(await supabase.auth.getUser()).data.user?.id).maybeSingle();
    const nomExpediteur=cfg?.smtp_user||comp?.raison_sociale||"Service RH";
    const msgPerso=cfg?.smtp_pass||"";
    const liste=avecTel.map(p=>`• ${p.employees?.nom} ${p.employees?.prenoms} → ${p.employees?.email}`).join("\n");
    if(!confirm(`Envoyer les fiches via WhatsApp à ${avecTel.length} employé(s) ?\n\n${liste}`)) return;
    setLoading(true);
    let ok=0, err=0;
    for(const p of avecTel){
      try{
        const emp=p.employees||{};
        const html=buildBulletinHTML(p,emp,comp,modele);
        // Uploader HTML dans Supabase Storage (s'ouvre bien sur mobile)
        const htmlComplet = html.replace("</body>","<script>window.onload=function(){document.title='Bulletin "+emp.nom+" "+MOIS[parseInt(p.mois)]+" "+p.annee+"';}<\/script></body>");
        const blob=new Blob([htmlComplet],{type:"text/html;charset=utf-8"});
        const fileName=`bulletin_${(emp.nom||"emp").replace(/\s+/g,"_")}_${p.mois}_${p.annee}_${Date.now()}.html`;
        const {error:upErr}=await supabase.storage.from("bulletins")
          .upload(fileName,blob,{cacheControl:"3600",upsert:true,contentType:"text/html"});
        if(upErr) throw upErr;
        const {data:urlData}=supabase.storage.from("bulletins").getPublicUrl(fileName);
        const lienFiche=urlData.publicUrl;
        const tel=(emp.email||"").replace(/[^0-9+]/g,"");
        // Casser le lien pour éviter l'aperçu en double dans WhatsApp
        const lienCasse=lienFiche.replace("https://","https:// ").replace(".pdf"," .pdf");
        const message=encodeURIComponent(
          `Bonjour ${emp.prenoms||""} ${emp.nom||""},

`+
          `Votre bulletin de paie de *${MOIS[parseInt(p.mois)]} ${p.annee}* est disponible.

`+
          `💰 *Net à payer : ${fmt(p.remuneration_due)}*
`+
          `📊 Salaire brut : ${fmt(p.salaire_brut)}
`+
          `🏛 CNSS : ${fmt(p.cnss_ouvriere)} | ITS : ${fmt(p.its)}

`+
          `📄 Votre bulletin est disponible ici :
${lienCasse}

`+
          `💡 _Le fichier PDF se trouve aussi dans votre dossier Téléchargements_

`+
          (msgPerso?`${msgPerso}

`:"")+
          `Cordialement,
*${nomExpediteur}*`
        );
        const waUrl=tel
          ?`https://wa.me/${tel}?text=${message}`
          :`https://wa.me/?text=${message}`;
        window.open(waUrl,"_blank");
        ok++;
        await new Promise(r=>setTimeout(r,2000));
      }catch(e){
        err++;
        console.error("Erreur WhatsApp:",e);
      }
    }
    setLoading(false);
    alert(err>0
      ?`✅ ${ok} fiche(s) envoyée(s) | ⚠️ ${err} erreur(s)`
      :`✅ ${ok} fiche(s) envoyée(s) avec succès via WhatsApp !`);
  };

  const del=async id=>{
    if(!confirm("Supprimer cette fiche ?")) return;
    const {error}=await supabase.from("payrolls").delete().eq("id",id);
    if(error) toast.error("Erreur lors de la suppression");
    else toast.success("✅ Fiche supprimée");
    load();
  };

  const printPdf=(p,mod=modele)=>{
    const emp=p.employees||{};
    const html=buildBulletinHTML(p,emp,comp,mod);
    const nomFichier=`Bulletin_${(emp.nom||"").replace(/\s+/g,"_")}_${(emp.prenoms||"").replace(/\s+/g,"_")}_${MOIS[parseInt(p.mois)]}_${p.annee}.html`;
    // Télécharger le fichier HTML (ouvrable + imprimable en PDF)
    const blob=new Blob([html],{type:"text/html;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=nomFichier;
    a.click();
    URL.revokeObjectURL(url);
    toast.success(`📄 Fiche téléchargée : ${nomFichier}`);
  };
  const printPdfWindow=(p,mod=modele)=>{
    const emp=p.employees||{};
    const html=buildBulletinHTML(p,emp,comp,mod);
    const win=window.open("","_blank");
    win.document.write(html);
    win.document.close();
    win.onload=()=>win.print();
  };

  const years=Array.from({length:5},(_,i)=>String(CURRENT_YEAR-i));

  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Fiches de Paie</h2>
        <div className="rh-btn-group">
          <Btn onClick={genererTout} variant="secondary" small>⚡ Générer tout le mois</Btn>
          <Btn onClick={envoyerTout} variant="ghost" small>📱 Envoyer via WhatsApp</Btn>
          <Btn onClick={()=>{setForm(initForm);setCalc(null);setMsg("");setModal("new")}}>+ Nouvelle fiche</Btn>
        </div>
      </div>

      {/* Filtres */}
      <Card style={{marginBottom:"16px",padding:"16px"}}>
        <div className="rh-flex-filters">
          <select className="rh-control" value={compId} onChange={e=>setCompId(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {companies.map(c=><option key={c.id} value={c.id}>{c.raison_sociale}</option>)}
          </select>
          <select value={filterMois} onChange={e=>setFilterMois(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {MOIS.slice(1).map((m,i)=><option key={i+1} value={String(i+1)}>{m}</option>)}
          </select>
          <select value={filterAnnee} onChange={e=>setFilterAnnee(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {years.map(y=><option key={y} value={y}>{y}</option>)}
          </select>
          <select value={modele} onChange={e=>setModele(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            <option value="officiel">📄 Officiel</option>
            <option value="simple">📋 Simple</option>
            <option value="moderne">✨ Moderne</option>
          </select>
        </div>
      </Card>

      <Card>
        {/* Totaux */}
        {payrolls.length>0&&(
          <div className="rh-summary-grid-4" style={{marginBottom:"16px",paddingBottom:"16px",borderBottom:`1px solid ${G.border}`}}>
            {[
              {label:"Brut total",val:payrolls.reduce((s,p)=>s+parseFloat(p.salaire_brut||0),0)},
              {label:"CNSS ouvrière",val:payrolls.reduce((s,p)=>s+parseFloat(p.cnss_ouvriere||0),0)},
              {label:"ITS total",val:payrolls.reduce((s,p)=>s+parseFloat(p.its||0),0)},
              {label:"Net à payer",val:payrolls.reduce((s,p)=>s+parseFloat(p.remuneration_due||0),0)},
            ].map(t=>(
              <div key={t.label} style={{background:"#f0f4ff",borderRadius:"8px",padding:"12px",border:"1px solid #d0dff5"}}>
                <div style={{fontSize:"11px",color:G.textDim,marginBottom:"4px"}}>{t.label}</div>
                <div style={{fontSize:"15px",fontWeight:700,color:G.accent}}>{fmt(t.val)}</div>
              </div>
            ))}
          </div>
        )}
        <Table cols={[
          {label:"Employé",render:r=>`${r.employees?.nom||""} ${r.employees?.prenoms||""}`},
          {key:"emploi",label:"Emploi",render:r=>r.employees?.emploi||"",dim:true},
          {label:"Brut",render:r=>fmt(r.salaire_brut)},
          {label:"CNSS",render:r=>fmt(r.cnss_ouvriere)},
          {label:"ITS",render:r=>fmt(r.its)},
          {label:"Net à payer",render:r=><span style={{color:G.accent,fontWeight:700}}>{fmt(r.remuneration_due)}</span>},
          {label:"Actions",render:r=>(
            <div style={{display:"flex",gap:"6px"}}>
              <Btn small onClick={e=>{e.stopPropagation();setViewModal(r)}} variant="secondary">Voir</Btn>
              <Btn small onClick={e=>{e.stopPropagation();printPdf(r)}} variant="ghost" title="Télécharger">⬇️</Btn>
              <Btn small onClick={e=>{e.stopPropagation();printPdfWindow(r)}} variant="ghost" title="Imprimer">🖨</Btn>
              <Btn small onClick={e=>{e.stopPropagation();del(r.id)}} variant="danger">✕</Btn>
            </div>
          )}
        ]} rows={payrolls}/>
      </Card>

      {/* Modal génération fiche */}
      {modal&&(
        <Modal title="Générer une fiche de paie" onClose={()=>setModal(null)} width="700px">
          <Alert msg={msg} type="error"/>
          <div className="rh-form-grid-2">
            <Select label="Employé *" value={form.employee_id} onChange={v=>f("employee_id",v)}
              options={[{value:"",label:"-- Choisir --"},...employees.map(e=>({value:e.id,label:`${e.nom} ${e.prenoms}`}))]}/>
            <Select label="Mois" value={form.mois} onChange={v=>f("mois",v)}
              options={MOIS.slice(1).map((m,i)=>({value:String(i+1),label:m}))}/>
            <Select label="Année" value={form.annee} onChange={v=>f("annee",v)}
              options={years.map(y=>({value:y,label:y}))}/>
            <Input label="Salaire base (FCFA)" value={form.salaire_base} onChange={v=>f("salaire_base",v)} type="number"/>
            <div style={{gridColumn:"1/-1",background:G.bg,borderRadius:"8px",padding:"12px",marginBottom:"4px"}}>
              <div style={{fontSize:"12px",fontWeight:600,color:G.accent,marginBottom:"10px"}}>
                Heures supplémentaires (taux horaire = salaire base ÷ 173,33)
              </div>
              <div className="rh-grid-4">
                <Input label="+12% (nb h)" value={form.h_sup_12} onChange={v=>f("h_sup_12",v)} type="number" small/>
                <Input label="+35% (nb h)" value={form.h_sup_35} onChange={v=>f("h_sup_35",v)} type="number" small/>
                <Input label="+50% (nb h)" value={form.h_sup_50} onChange={v=>f("h_sup_50",v)} type="number" small/>
                <Input label="+100% (nb h)" value={form.h_sup_100} onChange={v=>f("h_sup_100",v)} type="number" small/>
              </div>
              {calc&&calc.montant_heures_sup>0&&(
                <div style={{fontSize:"11px",color:G.textDim,marginTop:"6px"}}>
                  Total H.Sup : {fmt(calc.montant_heures_sup)}
                  {calc.montant_h12>0&&` | +12%: ${fmtN(calc.montant_h12)}`}
                  {calc.montant_h35>0&&` | +35%: ${fmtN(calc.montant_h35)}`}
                  {calc.montant_h50>0&&` | +50%: ${fmtN(calc.montant_h50)}`}
                  {calc.montant_h100>0&&` | +100%: ${fmtN(calc.montant_h100)}`}
                </div>
              )}
            </div>
            <Input label="Primes (FCFA)" value={form.primes} onChange={v=>f("primes",v)} type="number"/>
            <Input label="Indemnités (FCFA)" value={form.indemnites} onChange={v=>f("indemnites",v)} type="number"/>
            <Input label="Avances (FCFA)" value={form.avances} onChange={v=>f("avances",v)} type="number"/>
            <Input label="Saisie-arrêt (FCFA)" value={form.saisie_arret} onChange={v=>f("saisie_arret",v)} type="number"/>
            <Input label="Assurance santé (FCFA)" value={form.assurance_sante} onChange={v=>f("assurance_sante",v)} type="number"/>
          </div>

          {/* Aperçu calcul */}
          {calc&&(
            <div style={{background:"#f0f4ff",borderRadius:"10px",padding:"16px",marginTop:"8px",marginBottom:"16px",border:"1px solid #d0dff5"}}>
              <div style={{fontSize:"12px",fontWeight:700,color:G.textDim,marginBottom:"12px",
                           textTransform:"uppercase",letterSpacing:"0.5px"}}>Aperçu du calcul</div>
              <div className="rh-summary-grid-3">
                {[
                  {l:"Salaire Brut",v:calc.salaire_brut,c:G.text},
                  {l:"CNSS Ouvrière (3.6%)",v:calc.cnss_ouvriere,c:G.red},
                  {l:"ITS",v:calc.its,c:G.red},
                  {l:"Taxe Radio",v:calc.taxe_radio,c:G.yellow},
                  {l:"Salaire Net",v:calc.salaire_net,c:G.accent},
                  {l:"Net à payer",v:calc.remuneration_due,c:G.accent},
                  {l:"CNSS Patronale",v:calc.cnss_patronale,c:G.textDim},
                  {l:"VPS (4%)",v:calc.vps,c:calc.vps_exonere?G.yellow:G.textDim},
                ].map(x=>(
                  <div key={x.l} style={{background:G.card,borderRadius:"6px",padding:"8px 10px"}}>
                    <div style={{fontSize:"10px",color:G.textDim}}>{x.l}</div>
                    <div style={{fontSize:"13px",fontWeight:700,color:x.c}}>{fmt(x.v)}</div>
                  </div>
                ))}
              </div>
              {calc.vps_exonere&&(
                <div style={{marginTop:"10px",fontSize:"11px",color:G.yellow,background:G.yellow+"11",
                             padding:"6px 10px",borderRadius:"6px"}}>
                  ⚡ VPS exonéré : {calc.motif_exoneration}
                </div>
              )}
            </div>
          )}
          <div className="rh-actions-row">
            <Btn onClick={save} disabled={!calc||loading}>{loading?"...":"Enregistrer"}</Btn>
            <Btn onClick={()=>setModal(null)} variant="ghost">Annuler</Btn>
          </div>
        </Modal>
      )}

      {/* Modal détail fiche */}
      {viewModal&&(
        <Modal title={`Fiche — ${viewModal.employees?.nom} ${viewModal.employees?.prenoms}`}
               onClose={()=>setViewModal(null)} width="600px">
          <PayrollDetail p={viewModal} comp={comp}/>
          <div className="rh-actions-row" style={{marginTop:"16px"}}>
            <div className="rh-btn-group">
              <Btn onClick={()=>printPdf(viewModal,"officiel")} variant="secondary" small>📄 Officiel</Btn>
              <Btn onClick={()=>printPdf(viewModal,"simple")} variant="secondary" small>📋 Simple</Btn>
              <Btn onClick={()=>printPdf(viewModal,"moderne")} variant="primary" small>✨ Moderne</Btn>
            </div>
            <Btn onClick={()=>setViewModal(null)} variant="ghost">Fermer</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

const PayrollDetail = ({p,comp}) => {
  const rows=[
    {l:"Salaire de base",v:p.salaire_base,section:"GAINS"},
    {l:"Heures supp.",v:p.montant_heures_sup},
    {l:"Primes",v:p.primes},
    {l:"Indemnités",v:p.indemnites},
    {l:"SALAIRE BRUT",v:p.salaire_brut,bold:true},
    {l:"CNSS ouvrière (3.6%)",v:-p.cnss_ouvriere,section:"RETENUES",neg:true},
    {l:"ITS",v:-p.its,neg:true},
    {l:"Taxe radio",v:-p.taxe_radio,neg:true},
    {l:"SALAIRE NET",v:p.salaire_net,bold:true},
    {l:"Avances",v:-p.avances,neg:true},
    {l:"Saisie-arrêt",v:-p.saisie_arret,neg:true},
    {l:"Assurance santé",v:-p.assurance_sante,neg:true},
    {l:"RÉMUNÉRATION DUE",v:p.remuneration_due,bold:true,accent:true},
    {l:"CNSS patronale",v:p.cnss_patronale,section:"CHARGES PATRONALES"},
    {l:`VPS (4%)${p.vps_exonere?" [EXONÉRÉ]":""}`,v:p.vps},
  ];
  return (
    <div>
      <div style={{marginBottom:"16px",padding:"12px",background:G.bg,borderRadius:"8px",
                   fontSize:"13px",color:G.textDim}}>
        <strong style={{color:G.text}}>{comp?.raison_sociale}</strong> — {MOIS[p.mois]} {p.annee}
      </div>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
        <tbody>
          {rows.filter(r=>(Math.abs(r.v)||r.bold)).map((r,i)=>(
            <tr key={i} style={{borderBottom:`1px solid ${G.border}22`}}>
              <td style={{padding:"7px 10px",color:r.section?G.accent:r.bold?G.text:G.textDim,
                          fontWeight:r.bold?700:400,fontSize:r.section?"10px":"13px",
                          textTransform:r.section?"uppercase":"none"}}>
                {r.section||r.l}
              </td>
              {!r.section&&<td style={{padding:"7px 10px",textAlign:"right",
                color:r.accent?G.accent:r.neg&&r.v<0?G.red:G.text,fontWeight:r.bold?700:400}}>
                {fmt(Math.abs(r.v))}
              </td>}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// DECLARATIONS

// ─── GÉNÉRATION FICHES DÉCLARATION EXCEL ─────────────────────────────────────

const getTauxRisque = (taux) => {
  const t = Math.round(parseFloat(taux||0.194)*1000)/10;
  if(t<=16.4) return 1;
  if(t<=17.4) return 2;
  if(t<=18.4) return 3;
  return 4;
};

const genDeclarationCNSS = async (comp, employes, mois, annee) => {
  const ExcelJSLib = await import("exceljs");
  const wb = new ExcelJSLib.default.Workbook();
  wb.creator = "RH-Paie Pro";

  const moisNum = parseInt(mois);
  const joursFin = [4,6,9,11].includes(moisNum)?30:moisNum===2?28:31;
  const periodeDebut = `01/${String(moisNum).padStart(2,'0')}/${annee}`;
  const periodeFin   = `${joursFin}/${String(moisNum).padStart(2,'0')}/${annee}`;
  const taux = parseFloat(comp?.taux_cnss_patronale||0.194);
  const tauxRisque = getTauxRisque(taux);
  const ifu = comp?.ifu||"";
  const cnssEmployeur = comp?.cnss_employeur||comp?.rccm||"";
  const societe = comp?.raison_sociale||"";

  // Calculs
  const masse    = employes.reduce((s,e)=>s+parseFloat(e.salaire_brut||0),0);
  const nbEmp    = employes.length;
  const prestFam = Math.round(masse*0.09);
  const risques  = Math.round(masse*tauxRisque/100);
  const partPat  = Math.round(masse*0.064);
  const partOuv  = employes.reduce((s,e)=>s+parseFloat(e.cnss_ouvriere||0),0);
  const totAV    = partPat+partOuv;
  const totCot   = prestFam+risques+totAV;

  const thin = {style:'thin'};
  const border = {top:thin,left:thin,bottom:thin,right:thin};
  const tnr = (bold=false,sz=10) => ({name:'Times New Roman',bold,size:sz});
  const al = (h,v='middle',wrap=false) => ({horizontal:h,vertical:v,wrapText:wrap});
  const numFmt = '#,##0';

  // ── FEUILLE DECLARATION ──
  const ws = wb.addWorksheet("declaration", {
    pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:1},
    printArea:'A1:I57'
  });

  // Largeurs colonnes exactes
  ws.columns = [
    {width:36.57},{width:23.57},{width:2.71},{width:3.57},
    {width:9.43},{width:8.43},{width:3.57},{width:11.43},{width:11.43}
  ];

  // Hauteurs lignes exactes
  const rowH = {4:22.5,6:7.5,8:10.5,10:20.25,11:33.75,12:18.75,13:4.5,
    14:18.75,15:6.75,16:18.75,17:6.75,18:18.75,19:6.75,20:18.75,21:6.75,
    22:24,23:6.75,24:29.25,25:3.75,26:15,28:6.75,29:15,30:9.75,
    31:15,32:15,33:112.5,34:7.5,36:25.5,37:19.5,38:25.5,39:6.75,
    40:25.5,41:6.75,42:25.5,43:25.5,44:25.5,45:25.5,46:25.5,47:25.5,
    48:25.5,49:25.5,50:25.5,51:25.5,52:25.5,53:12,54:4.5,55:30,56:22.5,57:36};
  Object.entries(rowH).forEach(([r,h])=>{ ws.getRow(parseInt(r)).height=h; });

  const setCell = (row,col,val,bold=false,sz=10,halign='left',bord=false,fmt=null) => {
    const c = ws.getCell(row,col);
    c.value = val;
    c.font = tnr(bold,sz);
    c.alignment = al(halign);
    if(bord) c.border = border;
    if(fmt) c.numFmt = fmt;
  };

  // Fusions
  const merges = [
    'A1:B7','G1:I1','G2:I3','A10:I10','A11:E11',
    'A12:B12','D12:G12','H12:I12',
    'A14:B14','D14:F14','H14:I14',
    'A16:B16','D16:F16','H16:I16',
    'A18:B19','D18:G19','H18:I18',
    'A20:B20','D20:F20','H20:I20',
    'A22:B22','D22:G22','H22:I22',
    'A24:B24','D24:F24','H24:I24',
    'A26:B26','D26:H26',
    'A27:B33',
    'C35:E35','H35:I35',
    'A36:B36','C36:F36','H36:I36',
    'A38:B38','C38:F38','H38:I38',
    'A40:B40','C40:F40','H40:I40',
    'A42:B42','C42:F42','H42:I42',
    'A44:B44','C44:F44','H44:I44',
    'A46:B46','C46:F46','H46:I46',
    'A48:B48','C48:F48','H48:I48',
    'A50:B50','C50:F50','H50:I50',
    'A52:B52','C52:F52',
    'E55:I55','E56:I57',
  ];
  merges.forEach(m=>ws.mergeCells(m));

  // Contenu
  setCell(1,7,'N°  DOCUMENT',false,10,'center');
  setCell(2,7,'BENIN - DGI',false,11,'left');
  setCell(10,1,'Cotisation CNSS',true,15,'center');
  setCell(11,1,'CONTRIBUABLE  ET  RENSEIGNEMENTS',true,12,'left');
  setCell(11,6,'PERIODE DE COTISATION',true,11,'left');

  setCell(12,1,'IFU',true,10); setCell(12,4,ifu,false,10,'left',true);
  setCell(12,8,'NOM DU CONTRIBUABLE',true,10); setCell(12,8,societe,false,10,'left',true);
  ws.mergeCells('H12:I12'); setCell(12,8,societe,false,10,'left',true);

  setCell(14,1,'NUMERO COMPTE COTISATION',true,10);
  setCell(14,4,'',false,10,'left',true);
  setCell(14,8,'COTISATION',true,10);
  ws.mergeCells('H14:I14'); setCell(14,8,'CNSS',false,10,'left',true);

  setCell(16,1,'N° CNSS EMPLOYEUR',true,10);
  setCell(16,4,cnssEmployeur,false,10,'left',true);
  setCell(16,8,'BUR/SERV DE GESTION',true,10);
  ws.mergeCells('H16:I16'); setCell(16,8,'',false,10,'left',true);

  setCell(18,1,'CENTRE DE RECOUVREMENT',true,10);
  setCell(18,4,'',false,10,'left',true);
  ws.mergeCells('H18:I18'); setCell(18,8,'',false,10,'left',true);

  setCell(20,1,'AGENCE',true,10);
  setCell(20,4,'',false,10,'left',true);
  setCell(20,8,'OBJET  IMPOSABLE',true,10);
  ws.mergeCells('H20:I20'); setCell(20,8,'Salaires & accessoires',false,10,'left',true);

  setCell(22,1,'DEBUT DE LA PERIODE DE COTISATION',true,10);
  setCell(22,4,periodeDebut,false,10,'left',true);
  setCell(22,8,'FIN DE LA PERIODE DE COTISATION',true,10);
  ws.mergeCells('H22:I22'); setCell(22,8,periodeFin,false,10,'left',true);

  setCell(24,1,'DATE LIMITE DE DEPÔT',true,10);
  setCell(24,4,'',false,10,'left',true);
  setCell(24,8,'DATE LIMITE DE PAIEMENT',true,10);
  ws.mergeCells('H24:I24'); setCell(24,8,'',false,10,'left',true);

  setCell(26,1,'ADRESSE DE CORRESPONDANCE',true,10);
  setCell(26,4,'LOCALISATION',true,10,'center');

  const contact = `Pour nous contacter:
Direction Générale des Impôts
01 BP 369, Cotonou - République du Bénin
Téléphone: 21-30-57-27 Fax: 21-30-57-36
Site web: http://www.impots.finances.gouv.bj
Caisse Nationale de Sécurité Sociale
BP 374 - Cotonou Téléphone: 21-30-27-65`;
  const cContact = ws.getCell(27,1);
  cContact.value = contact; cContact.font = tnr(false,10);
  cContact.alignment = {horizontal:'left',vertical:'top',wrapText:true};

  setCell(35,3,'Annexe  Fiscale',true,10);
  setCell(35,7,'Ligne',true,10);
  setCell(35,8,'Montant',true,10,'right');

  const lignes = [
    [36,'Effectif Total',nbEmp,'3'],
    [38,'Masse salariale',masse,'5'],
    [40,'Prestations familiales (9%)',prestFam,'10'],
    [42,'Risques Professionnels (1 à 4%)',risques,'20'],
    [44,'Assurance vieillesse - Part patronale (6,4%) ',partPat,'30'],
    [46,'Assurance vieillesse - Part ouvrière (3,6%)',partOuv,'40'],
    [48,'Total assurance vieillesse (L30+L40)',totAV,'50'],
    [50,'Total des cotisations (L10+L20+L50) ',totCot,'60'],
    [52,'Total des cotisations à payer ',totCot,'70'],
  ];

  lignes.forEach(([row,lib,val,lig])=>{
    setCell(row,1,lib,false,10,'left');
    setCell(row,3,'……………………………………………….',false,10,'center');
    setCell(row,7,lig,false,10,'right');
    const cm = ws.getCell(row,8);
    cm.value=val; cm.font=tnr(false,10);
    cm.alignment=al('right'); cm.border=border; cm.numFmt=numFmt;
    if(lig==='70'){
      const c52=ws.getCell(row,3);
      c52.value='ANNEXE COTISATION CNSS'; c52.font=tnr(true,10);
      c52.alignment=al('center');
    }
  });

  setCell(55,5,`Cotonou, le ${periodeFin}`,false,11);
  setCell(56,5,'  Nom, prénom et signature du déclarant',false,11,'center');

  // ── FEUILLE ANNEXE ──
  const ws2 = wb.addWorksheet("annexe", {
    pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1}
  });

  const entetes = [
    "N° d'ordre","Numéro d'assurance du travailleur","N° IFU de l'employé",
    "N° employeur",`Taux risques prof. (A) = ${tauxRisque}`,
    "Nom du travailleur","Prénoms du travailleur","Date embauche","Date débauchage",
    "Jours ouvrables","Jours assimilés","Rémunération F CFA (1)",
    "Rappel salaires (2)","Période rappels","IDR (3)","Indemnité licenciement (4)",
    "Autres primes (5)","Total rémunérations (6)=(1+2+3+4+5)",
    "Prest. familiales (7)=6×9%",`Risques prof. (8)=6×${tauxRisque}%`,
    "Ass. vieillesse pat. (9)=6×6,4%","Ass. vieillesse ouv. (10)=6×3,6%",
    "Total cotisations (11)=7+8+9+10"
  ];
  const colW2 = [4,16,12,10,12,14,14,10,10,8,8,14,12,13,9,11,11,15,9,9,11,11,12];
  ws2.columns = colW2.map(w=>({width:w}));
  ws2.getRow(1).height=18; ws2.getRow(4).height=120;

  // Titre annexe
  ws2.mergeCells(1,1,1,23);
  const cTitre=ws2.getCell(1,1);
  cTitre.value='ANNEXE COTISATION CNSS';
  cTitre.font=tnr(true,12); cTitre.alignment=al('center'); cTitre.border=border;

  ws2.mergeCells(2,1,2,23);
  const cSub=ws2.getCell(2,1);
  cSub.value=`${societe}  —  Période : ${periodeDebut} au ${periodeFin}  —  N° CNSS : ${cnssEmployeur}`;
  cSub.font={name:'Times New Roman',size:9,italic:true};
  cSub.alignment=al('center');

  ws2.getRow(3).height=5;

  // En-têtes
  entetes.forEach((h,i)=>{
    const c=ws2.getCell(4,i+1);
    c.value=h; c.font=tnr(true,8);
    c.alignment={horizontal:'center',vertical:'middle',wrapText:true};
    c.border=border;
  });

  // Données
  employes.forEach((emp,i)=>{
    const row=5+i;
    const brut=parseFloat(emp.salaire_brut||0);
    const ouv=parseFloat(emp.cnss_ouvriere||0);
    const pat=Math.round(brut*0.064);
    const fam=Math.round(brut*0.09);
    const risp=Math.round(brut*tauxRisque/100);
    const tot=fam+risp+pat+ouv;
    const vals=[i+1,emp.cnss||'',emp.ifu||'',cnssEmployeur,tauxRisque,
      emp.nom||'',emp.prenoms||'',emp.date_embauche||'','',26,0,
      brut,0,'',0,0,0,brut,fam,risp,pat,ouv,tot];
    ws2.getRow(row).height=18;
    vals.forEach((v,j)=>{
      const c=ws2.getCell(row,j+1);
      c.value=v; c.font=tnr(false,9);
      c.alignment=al(j>=11?'right':j<=4?'center':'left');
      c.border=border;
      if(j>=11&&typeof v==='number') c.numFmt=numFmt;
    });
  });

  // Totaux
  const tr=5+employes.length;
  ws2.mergeCells(tr,1,tr,11);
  const cTot=ws2.getCell(tr,1);
  cTot.value='TOTAUX'; cTot.font=tnr(true,9);
  cTot.alignment=al('center'); cTot.border=border;
  ws2.getRow(tr).height=18;
  for(let j=12;j<=23;j++){
    const cl=String.fromCharCode(64+j); // approx col letter
    const c=ws2.getCell(tr,j);
    c.value=employes.reduce((s,e)=>{
      const brut=parseFloat(e.salaire_brut||0);
      const ouv=parseFloat(e.cnss_ouvriere||0);
      const vals=[0,0,0,0,0,0,0,brut,Math.round(brut*0.09),Math.round(brut*tauxRisque/100),Math.round(brut*0.064),ouv,Math.round(brut*0.09)+Math.round(brut*tauxRisque/100)+Math.round(brut*0.064)+ouv];
      return s+(vals[j-11]||0);
    },0);
    c.font=tnr(true,9); c.alignment=al('right');
    c.border=border; c.numFmt=numFmt;
  }

  ws2.printArea=`A1:W${tr+1}`;

  // Téléchargement
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href=url;
  a.download=`Declaration_CNSS_${societe.replace(/\s+/g,'_')}_${MOIS[moisNum]}_${annee}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
};

const genDeclarationITS = async (comp, employes, mois, annee) => {
  const ExcelJSLib = await import("exceljs");
  const wb = new ExcelJSLib.default.Workbook();

  const moisNum = parseInt(mois);
  const joursFin = [4,6,9,11].includes(moisNum)?30:moisNum===2?28:31;
  const periodeDebut = `01/${String(moisNum).padStart(2,'0')}/${annee}`;
  const periodeFin   = `${joursFin}/${String(moisNum).padStart(2,'0')}/${annee}`;
  const ifu = comp?.ifu||"";
  const cnssEmployeur = comp?.cnss_employeur||comp?.rccm||"";
  const societe = comp?.raison_sociale||"";

  const nbEmp  = employes.length;
  const masse  = employes.reduce((s,e)=>s+parseFloat(e.salaire_brut||0),0);
  const itsTot = employes.reduce((s,e)=>s+parseFloat(e.its||0),0);

  const thin={style:'thin'};
  const border={top:thin,left:thin,bottom:thin,right:thin};
  const tnr=(bold=false,sz=10)=>({name:'Times New Roman',bold,size:sz});
  const al=(h,v='middle',wrap=false)=>({horizontal:h,vertical:v,wrapText:wrap});

  const ws = wb.addWorksheet("declaration",{
    pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:1},
    printArea:'A1:I46'
  });

  ws.columns=[{width:36.57},{width:23.57},{width:2.71},{width:3.57},
    {width:9.43},{width:8.43},{width:3.57},{width:11.43},{width:11.43}];

  const rowH={4:22.5,6:7.5,8:10.5,10:20.25,11:33.75,12:18.75,13:4.5,
    14:18.75,15:6.75,16:18.75,17:6.75,18:18.75,19:6.75,20:18.75,21:6.75,
    22:24,23:6.75,24:29.25,25:3.75,26:15,28:6.75,29:15,30:9.75,
    31:15,32:15,33:112.5,34:7.5,36:25.5,37:19.5,38:25.5,39:6.75,
    40:25.5,41:6.75,42:12,43:4.5,44:30,45:22.5,46:36};
  Object.entries(rowH).forEach(([r,h])=>{ws.getRow(parseInt(r)).height=h;});

  const setCell=(row,col,val,bold=false,sz=10,halign='left',bord=false,fmt=null)=>{
    const c=ws.getCell(row,col);
    c.value=val; c.font=tnr(bold,sz); c.alignment=al(halign);
    if(bord) c.border=border;
    if(fmt) c.numFmt=fmt;
  };

  // Mêmes fusions que CNSS sauf lignes 42-52
  ['A1:B7','G1:I1','G2:I3','A10:I10','A11:E11',
   'A12:B12','D12:G12','H12:I12',
   'A14:B14','D14:F14','H14:I14',
   'A16:B16','D16:F16','H16:I16',
   'A18:B19','D18:G19','H18:I18',
   'A20:B20','D20:F20','H20:I20',
   'A22:B22','D22:G22','H22:I22',
   'A24:B24','D24:F24','H24:I24',
   'A26:B26','D26:H26','A27:B33',
   'C35:E35','H35:I35',
   'A36:B36','C36:F36','H36:I36',
   'A38:B38','C38:F38','H38:I38',
   'A40:B40','C40:F40','H40:I40',
   'E44:I44','E45:I46',
  ].forEach(m=>ws.mergeCells(m));

  setCell(1,7,'N°  DOCUMENT',false,10,'center');
  setCell(2,7,'BENIN - DGI',false,11);
  setCell(10,1,'IRPP TS',true,15,'center');
  setCell(11,1,'CONTRIBUABLE  ET  RENSEIGNEMENTS',true,12);
  setCell(11,6,'PERIODE DE COTISATION',true,11);

  setCell(12,1,'IFU',true,10); setCell(12,4,ifu,false,10,'left',true);
  setCell(12,8,'NOM DU CONTRIBUABLE',true,10); setCell(12,8,societe,false,10,'left',true);
  ws.mergeCells('H12:I12'); setCell(12,8,societe,false,10,'left',true);

  setCell(14,1,'NUMERO COMPTE COTISATION',true,10);
  setCell(14,4,'',false,10,'left',true);
  setCell(14,8,'COTISATION',true,10);
  ws.mergeCells('H14:I14'); setCell(14,8,'IRPP-TS',false,10,'left',true);

  setCell(16,1,'N° CNSS EMPLOYEUR',true,10);
  setCell(16,4,cnssEmployeur,false,10,'left',true);
  setCell(16,8,'BUR/SERV DE GESTION',true,10);
  ws.mergeCells('H16:I16'); setCell(16,8,'',false,10,'left',true);

  setCell(18,1,'CENTRE DE RECOUVREMENT',true,10);
  setCell(18,4,'',false,10,'left',true);
  ws.mergeCells('H18:I18'); setCell(18,8,'',false,10,'left',true);

  setCell(20,1,'AGENCE',true,10); setCell(20,4,'',false,10,'left',true);
  setCell(20,8,'OBJET  IMPOSABLE',true,10);
  ws.mergeCells('H20:I20'); setCell(20,8,'Salaires & accessoires',false,10,'left',true);

  setCell(22,1,'DEBUT DE LA PERIODE DE COTISATION',true,10);
  setCell(22,4,periodeDebut,false,10,'left',true);
  setCell(22,8,'FIN DE LA PERIODE DE COTISATION',true,10);
  ws.mergeCells('H22:I22'); setCell(22,8,periodeFin,false,10,'left',true);

  setCell(24,1,'DATE LIMITE DE DEPÔT',true,10);
  setCell(24,4,'',false,10,'left',true);
  setCell(24,8,'DATE LIMITE DE PAIEMENT',true,10);
  ws.mergeCells('H24:I24'); setCell(24,8,'',false,10,'left',true);

  setCell(26,1,'ADRESSE DE CORRESPONDANCE',true,10);
  setCell(26,4,'LOCALISATION',true,10,'center');

  const cContact=ws.getCell(27,1);
  cContact.value=`Pour nous contacter:
Direction Générale des Impôts
01 BP 369, Cotonou - République du Bénin
Téléphone: 21-30-57-27 Fax: 21-30-57-36
Site web: http://www.impots.finances.gouv.bj
Heures d'ouvertures: Lundi à Vendredi de 7h00 à 12h30 et 15h00 à 18h30`;
  cContact.font=tnr(false,10);
  cContact.alignment={horizontal:'left',vertical:'top',wrapText:true};

  setCell(35,3,'Annexe  Fiscale',true,10);
  setCell(35,7,'Ligne',true,10);
  setCell(35,8,'Montant',true,10,'right');

  [[36,'Nombre de salariés',nbEmp,'5'],
   [38,'Montant brut des salaires',masse,'10'],
   [40,'IRPP/TS à reverser',itsTot,'15'],
  ].forEach(([row,lib,val,lig])=>{
    setCell(row,1,lib,false,10,'left');
    setCell(row,3,'……………………………………………….',false,10,'center');
    setCell(row,7,lig,false,10,'right');
    const cm=ws.getCell(row,8);
    cm.value=val; cm.font=tnr(false,10);
    cm.alignment=al('right'); cm.border=border; cm.numFmt='#,##0';
  });

  setCell(44,5,`Cotonou, le ${periodeFin}`,false,11);
  setCell(45,5,'  Nom, prénom et signature du déclarant',false,11,'center');

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download=`Declaration_ITS_${societe.replace(/\s+/g,'_')}_${MOIS[moisNum]}_${annee}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
};

const Declarations = ({companies}) => {
  const [compId,setCompId]=useState(companies[0]?.id||"");
  const [mois,setMois]=useState(String(new Date().getMonth()+1));
  const [annee,setAnnee]=useState(String(CURRENT_YEAR));
  const [data,setData]=useState(null);
  const [decl,setDecl]=useState(null);
  const [msg,setMsg]=useState("");
  const years=Array.from({length:5},(_,i)=>String(CURRENT_YEAR-i));

  const load = useCallback(async()=>{
    if(!compId) return;
    const {data:p}=await supabase.from("payrolls")
      .select("*, employees(nom,prenoms,cnss,ifu,date_embauche,emploi)")
      .eq("company_id",compId).eq("mois",parseInt(mois)).eq("annee",parseInt(annee));
    const payrolls=p||[];
    const agg={
      cnss_employe:payrolls.reduce((s,x)=>s+parseFloat(x.cnss_ouvriere||0),0),
      cnss_employeur:payrolls.reduce((s,x)=>s+parseFloat(x.cnss_patronale||0),0),
      vps:payrolls.reduce((s,x)=>s+parseFloat(x.vps||0),0),
      its_total:payrolls.reduce((s,x)=>s+parseFloat(x.its||0),0),
      nb_fiches:payrolls.length
    };
    agg.cnss_total=agg.cnss_employe+agg.cnss_employeur;
    setData({...agg, payrolls_detail: p||[]});
    const {data:d}=await supabase.from("declarations").select("*")
      .eq("company_id",compId).eq("mois",parseInt(mois)).eq("annee",parseInt(annee)).maybeSingle();
    setDecl(d);
  },[compId,mois,annee]);

  useEffect(()=>{load();},[load]);
  useEffect(()=>{if(!compId&&companies.length) setCompId(companies[0].id);},[companies]);

  const updateStatut = async (field,statut) => {
    const now=new Date().toISOString().split("T")[0];
    const dateField=field==="statut_cnss"?"date_declaration_cnss":
                    field==="statut_vps"?"date_declaration_vps":"date_declaration_its";
    const payload={company_id:compId,mois:parseInt(mois),annee:parseInt(annee),
      cnss_employe:data?.cnss_employe||0,cnss_employeur:data?.cnss_employeur||0,cnss_total:data?.cnss_total||0,vps:data?.vps||0,its_total:data?.its_total||0,updated_at:new Date().toISOString(),[field]:statut,
      ...(statut==="Déclaré"?{[dateField]:now}:{})};
    const {error}=await supabase.from("declarations").upsert(payload,{onConflict:"company_id,mois,annee"});
    if(error) setMsg(error.message); else load();
  };

  const StatutRow=({label,field,montant,color=G.accent})=>{
    const st=decl?.[field]||"Non déclaré";
    return (
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
                   padding:"14px 0",borderBottom:`1px solid ${G.border}`}}>
        <div>
          <div style={{fontWeight:600,fontSize:"14px"}}>{label}</div>
          <div style={{fontSize:"18px",fontWeight:800,color,marginTop:"4px"}}>{fmt(montant)}</div>
        </div>
        <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
          <Badge label={st} color={st==="Déclaré"?G.accent:G.yellow}/>
          {st!=="Déclaré"
            ?<Btn small onClick={()=>updateStatut(field,"Déclaré")} variant="primary">Marquer déclaré</Btn>
            :decl?.[field.replace("statut_","date_declaration_")]&&(
              <span style={{fontSize:"11px",color:G.textDim}}>
                le {decl[field.replace("statut_","date_declaration_")]}
              </span>
            )
          }
        </div>
      </div>
    );
  };

  const exportExcel=()=>{
    if(!data) return;
    const comp=companies.find(c=>c.id===compId);
    const rows=[
      ["RAPPORT DE DÉCLARATIONS SOCIALES ET FISCALES"],
      ["Société :", comp?.raison_sociale||""],
      ["Période :", `${MOIS[parseInt(mois)]} ${annee}`],
      ["Nombre de fiches :", data.nb_fiches||0],
      [],
      ["RUBRIQUE","MONTANT (FCFA)","STATUT","DATE DÉCLARATION"],
      ["CNSS Salarié (3,6%)", data.cnss_employe||0, decl?.statut_cnss||"Non déclaré", decl?.date_declaration_cnss||""],
      ["CNSS Patronale", data.cnss_employeur||0, "", ""],
      ["Total CNSS", data.cnss_total||0, decl?.statut_cnss||"Non déclaré", ""],
      ["VPS (4%)", data.vps||0, decl?.statut_vps||"Non déclaré", decl?.date_declaration_vps||""],
      ["ITS", data.its_total||0, decl?.statut_its||"Non déclaré", decl?.date_declaration_its||""],
      [],
      ["TOTAL À VERSER", (data.cnss_total||0)+(data.vps||0)+(data.its_total||0), "", ""],
    ];
    // Créer CSV
    const csv=rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,"\"")}"`).join(",")).join("\n");
    const bom="\uFEFF";
    const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8;"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=`Declarations_${comp?.raison_sociale||"societe"}_${MOIS[parseInt(mois)]}_${annee}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const printDecl=()=>{
    const html=buildDeclarationHTML(data,decl,companies.find(c=>c.id===compId),mois,annee);
    const win=window.open("","_blank"); win.document.write(html); win.document.close();
    win.onload=()=>win.print();
  };

  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Déclarations CNSS / VPS / ITS</h2>
        {data?.nb_fiches>0&&(
          <div className="rh-btn-group">
            <Btn onClick={printDecl} variant="secondary" small>🖨 Imprimer</Btn>
            <Btn onClick={async()=>{
              const comp=companies.find(c=>c.id===compId);
              const emps=(data?.payrolls_detail||[]).map(p=>({
                ...p,...(p.employees||{}),
                nom:p.employees?.nom||'',prenoms:p.employees?.prenoms||'',
                cnss:p.employees?.cnss||'',ifu:p.employees?.ifu||'',
                date_embauche:p.employees?.date_embauche||'',
              }));
              await genDeclarationCNSS(comp,emps,mois,annee);
              toast.success("✅ Déclaration CNSS téléchargée !");
            }} variant="primary" small>📊 CNSS Excel</Btn>
            <Btn onClick={async()=>{
              const comp=companies.find(c=>c.id===compId);
              const emps=(data?.payrolls_detail||[]).map(p=>({
                ...p,...(p.employees||{}),
                nom:p.employees?.nom||'',prenoms:p.employees?.prenoms||'',
              }));
              await genDeclarationITS(comp,emps,mois,annee);
              toast.success("✅ Déclaration ITS téléchargée !");
            }} variant="ghost" small>📊 ITS Excel</Btn>
          </div>
        )}
      </div>
      <Card style={{marginBottom:"16px",padding:"16px"}}>
        <div className="rh-flex-filters">
          <select className="rh-control" value={compId} onChange={e=>setCompId(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {companies.map(c=><option key={c.id} value={c.id}>{c.raison_sociale}</option>)}
          </select>
          <select value={mois} onChange={e=>setMois(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {MOIS.slice(1).map((m,i)=><option key={i+1} value={String(i+1)}>{m}</option>)}
          </select>
          <select value={annee} onChange={e=>setAnnee(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {years.map(y=><option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </Card>
      <Alert msg={msg} type="error"/>
      {data&&(
        data.nb_fiches===0
          ?<Card><div style={{color:G.textDim,textAlign:"center",padding:"24px"}}>
            Aucune fiche de paie pour {MOIS[parseInt(mois)]} {annee}
          </div></Card>
          :<Card>
            <div style={{marginBottom:"16px",fontSize:"13px",color:G.textDim}}>
              {data.nb_fiches} fiche(s) de paie — {MOIS[parseInt(mois)]} {annee}
            </div>
            <StatutRow label="CNSS Employé + Employeur" field="statut_cnss"
              montant={data.cnss_total} color={G.accent}/>
            <StatutRow label="VPS (Versement Patronal Social)" field="statut_vps"
              montant={data.vps} color="#6366f1"/>
            <StatutRow label="ITS (Impôt sur Traitements et Salaires)" field="statut_its"
              montant={data.its_total} color={G.yellow}/>
            {/* Récap */}
            <div style={{marginTop:"20px",padding:"16px",background:G.bg,borderRadius:"8px"}}>
              <div style={{fontSize:"11px",fontWeight:700,color:G.textDim,
                           textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:"12px"}}>
                Récapitulatif
              </div>
              <div className="rh-summary-grid-3">
                {[
                  {l:"CNSS Salarié",v:data.cnss_employe},
                  {l:"CNSS Patronale",v:data.cnss_employeur},
                  {l:"Total CNSS",v:data.cnss_total},
                  {l:"VPS",v:data.vps},
                  {l:"ITS",v:data.its_total},
                  {l:"Total à verser",v:data.cnss_total+data.vps+data.its_total},
                ].map(x=>(
                  <div key={x.l} style={{background:G.card,borderRadius:"6px",padding:"10px"}}>
                    <div style={{fontSize:"10px",color:G.textDim,marginBottom:"4px"}}>{x.l}</div>
                    <div style={{fontSize:"14px",fontWeight:700,color:G.accent}}>{fmt(x.v)}</div>
                  </div>
                ))}
              </div>
            </div>
          </Card>
      )}
    </div>
  );
};



// ─── RAPPORT CABINET ─────────────────────────────────────────────────────────
const RapportCabinet = ({companies}) => {
  const [mois,setMois]=useState(String(new Date().getMonth()+1));
  const [annee,setAnnee]=useState(String(CURRENT_YEAR));
  const [rapports,setRapports]=useState([]);
  const [loading,setLoading]=useState(false);
  const years=Array.from({length:5},(_,i)=>String(CURRENT_YEAR-i));

  const generer=useCallback(async()=>{
    if(!companies.length) return;
    setLoading(true);
    const res=[];
    for(const comp of companies){
      const {data:payrolls}=await supabase.from("payrolls").select("*")
        .eq("company_id",comp.id).eq("mois",parseInt(mois)).eq("annee",parseInt(annee));
      const p=payrolls||[];
      const {data:decl}=await supabase.from("declarations").select("*")
        .eq("company_id",comp.id).eq("mois",parseInt(mois)).eq("annee",parseInt(annee)).maybeSingle();
      res.push({
        comp,
        nb:p.length,
        masse_brute:p.reduce((s,x)=>s+parseFloat(x.salaire_brut||0),0),
        masse_nette:p.reduce((s,x)=>s+parseFloat(x.remuneration_due||0),0),
        cnss_salarie:p.reduce((s,x)=>s+parseFloat(x.cnss_ouvriere||0),0),
        cnss_patronale:p.reduce((s,x)=>s+parseFloat(x.cnss_patronale||0),0),
        vps:p.reduce((s,x)=>s+parseFloat(x.vps||0),0),
        its:p.reduce((s,x)=>s+parseFloat(x.its||0),0),
        decl,
      });
    }
    setRapports(res);
    setLoading(false);
  },[companies,mois,annee]);

  useEffect(()=>{generer();},[generer]);

  const exportCSV=()=>{
    const rows=[
      [`RAPPORT CABINET — ${MOIS[parseInt(mois)]} ${annee}`],
      [],
      ["SOCIÉTÉ","NB FICHES","MASSE BRUTE","MASSE NETTE","CNSS SALARIÉ","CNSS PATRONAL","TOTAL CNSS","VPS","ITS","TOTAL À VERSER"],
      ...rapports.map(r=>[
        r.comp.raison_sociale,
        r.nb,
        Math.round(r.masse_brute),
        Math.round(r.masse_nette),
        Math.round(r.cnss_salarie),
        Math.round(r.cnss_patronale),
        Math.round(r.cnss_salarie+r.cnss_patronale),
        Math.round(r.vps),
        Math.round(r.its),
        Math.round(r.cnss_salarie+r.cnss_patronale+r.vps+r.its),
      ]),
      [],
      ["TOTAUX",
        rapports.reduce((s,r)=>s+r.nb,0),
        Math.round(rapports.reduce((s,r)=>s+r.masse_brute,0)),
        Math.round(rapports.reduce((s,r)=>s+r.masse_nette,0)),
        Math.round(rapports.reduce((s,r)=>s+r.cnss_salarie,0)),
        Math.round(rapports.reduce((s,r)=>s+r.cnss_patronale,0)),
        Math.round(rapports.reduce((s,r)=>s+r.cnss_salarie+r.cnss_patronale,0)),
        Math.round(rapports.reduce((s,r)=>s+r.vps,0)),
        Math.round(rapports.reduce((s,r)=>s+r.its,0)),
        Math.round(rapports.reduce((s,r)=>s+r.cnss_salarie+r.cnss_patronale+r.vps+r.its,0)),
      ],
    ];
    const csv=rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,"\"")}"`).join(",")).join("\n");
    const blob=new Blob(["﻿"+csv],{type:"text/csv;charset=utf-8;"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=`Rapport_Cabinet_${MOIS[parseInt(mois)]}_${annee}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const printRapport=()=>{
    const html=`<!DOCTYPE html><html><head><meta charset="UTF-8">
    <title>Rapport Cabinet</title>
    <style>
    *{margin:0;padding:0;box-sizing:border-box}
    @page{margin:10mm;size:A4 landscape}
    body{font-family:Arial,sans-serif;font-size:10px;padding:10px}
    h1{font-size:14px;font-weight:bold;text-align:center;margin-bottom:4px}
    h2{font-size:11px;color:#555;text-align:center;margin-bottom:12px}
    table{width:100%;border-collapse:collapse}
    th{background:#1a1a2e;color:#fff;padding:6px 8px;text-align:right;font-size:10px}
    th:first-child{text-align:left}
    td{padding:5px 8px;border-bottom:1px solid #eee;text-align:right}
    td:first-child{text-align:left;font-weight:500}
    .total{background:#e8f5f1;font-weight:bold;border-top:2px solid #00c896}
    .status{font-size:9px;padding:2px 6px;border-radius:10px}
    .ok{background:#e8f5e9;color:#2e7d32}
    .nok{background:#fff3e0;color:#e65100}
    </style></head><body>
    <h1>RAPPORT CABINET — ${MOIS[parseInt(mois)]} ${annee}</h1>
    <h2>Généré le ${new Date().toLocaleDateString("fr-FR")}</h2>
    <table>
    <thead><tr>
      <th>Société</th><th>Fiches</th><th>Masse Brute</th><th>Masse Nette</th>
      <th>CNSS Sal.</th><th>CNSS Pat.</th><th>Total CNSS</th>
      <th>VPS</th><th>ITS</th><th>Total à verser</th>
      <th>Statuts</th>
    </tr></thead>
    <tbody>
    ${rapports.map(r=>`<tr>
      <td>${r.comp.raison_sociale}</td>
      <td>${r.nb}</td>
      <td>${fmtN(r.masse_brute)}</td>
      <td>${fmtN(r.masse_nette)}</td>
      <td>${fmtN(r.cnss_salarie)}</td>
      <td>${fmtN(r.cnss_patronale)}</td>
      <td>${fmtN(r.cnss_salarie+r.cnss_patronale)}</td>
      <td>${fmtN(r.vps)}</td>
      <td>${fmtN(r.its)}</td>
      <td><strong>${fmtN(r.cnss_salarie+r.cnss_patronale+r.vps+r.its)}</strong></td>
      <td>
        <span class="status ${r.decl?.statut_cnss==="Déclaré"?"ok":"nok"}">CNSS:${r.decl?.statut_cnss||"Non déclaré"}</span>
        <span class="status ${r.decl?.statut_vps==="Déclaré"?"ok":"nok"}">VPS:${r.decl?.statut_vps||"Non déclaré"}</span>
        <span class="status ${r.decl?.statut_its==="Déclaré"?"ok":"nok"}">ITS:${r.decl?.statut_its||"Non déclaré"}</span>
      </td>
    </tr>`).join("")}
    <tr class="total">
      <td>TOTAUX</td>
      <td>${rapports.reduce((s,r)=>s+r.nb,0)}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.masse_brute,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.masse_nette,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.cnss_salarie,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.cnss_patronale,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.cnss_salarie+r.cnss_patronale,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.vps,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.its,0))}</td>
      <td>${fmtN(rapports.reduce((s,r)=>s+r.cnss_salarie+r.cnss_patronale+r.vps+r.its,0))}</td>
      <td></td>
    </tr>
    </tbody></table>
    </body></html>`;
    const win=window.open("","_blank");
    win.document.write(html); win.document.close();
    win.onload=()=>win.print();
  };

  const totaux={
    nb:rapports.reduce((s,r)=>s+r.nb,0),
    brute:rapports.reduce((s,r)=>s+r.masse_brute,0),
    nette:rapports.reduce((s,r)=>s+r.masse_nette,0),
    cnss:rapports.reduce((s,r)=>s+r.cnss_salarie+r.cnss_patronale,0),
    vps:rapports.reduce((s,r)=>s+r.vps,0),
    its:rapports.reduce((s,r)=>s+r.its,0),
  };

  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Rapport Cabinet</h2>
        <div className="rh-btn-group">
          <Btn onClick={printRapport} variant="secondary" small>🖨 Rapport PDF</Btn>
          <Btn onClick={exportCSV} variant="ghost" small>📊 Export Excel/CSV</Btn>
        </div>
      </div>
      <Card style={{marginBottom:"16px",padding:"16px"}}>
        <div className="rh-inline-actions">
          <select value={mois} onChange={e=>setMois(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {MOIS.slice(1).map((m,i)=><option key={i+1} value={String(i+1)}>{m}</option>)}
          </select>
          <select value={annee} onChange={e=>setAnnee(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {years.map(y=><option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </Card>

      {/* Totaux globaux */}
      {rapports.length>0&&(
        <div className="rh-summary-grid-4" style={{marginBottom:"16px"}}>
          {[
            {l:"Fiches totales",v:totaux.nb,c:G.text,unit:""},
            {l:"Masse brute totale",v:totaux.brute,c:G.accent,unit:" FCFA"},
            {l:"Total CNSS",v:totaux.cnss,c:"#6366f1",unit:" FCFA"},
            {l:"Total à verser",v:totaux.cnss+totaux.vps+totaux.its,c:G.red,unit:" FCFA"},
          ].map(x=>(
            <Card key={x.l} style={{padding:"14px"}}>
              <div style={{fontSize:"11px",color:G.textDim,marginBottom:"4px"}}>{x.l}</div>
              <div style={{fontSize:"16px",fontWeight:800,color:x.c}}>
                {x.unit?fmt(x.v):x.v}
              </div>
            </Card>
          ))}
        </div>
      )}

      {loading&&<div style={{color:G.textDim,padding:"24px",textAlign:"center"}}>Chargement...</div>}
      
      {rapports.map((r,i)=>(
        <Card key={i} style={{marginBottom:"10px",padding:"16px"}}>
          <div className="rh-item-header" style={{marginBottom:"10px"}}>
            <div>
              <div style={{fontWeight:700,fontSize:"14px"}}>{r.comp.raison_sociale}</div>
              <div style={{fontSize:"12px",color:G.textDim}}>{r.nb} fiche(s) — {MOIS[parseInt(mois)]} {annee}</div>
            </div>
            <div className="rh-badges">
              <Badge label={"CNSS: "+(r.decl?.statut_cnss||"Non déclaré")}
                     color={r.decl?.statut_cnss==="Déclaré"?G.accent:G.yellow}/>
              <Badge label={"VPS: "+(r.decl?.statut_vps||"Non déclaré")}
                     color={r.decl?.statut_vps==="Déclaré"?G.accent:G.yellow}/>
              <Badge label={"ITS: "+(r.decl?.statut_its||"Non déclaré")}
                     color={r.decl?.statut_its==="Déclaré"?G.accent:G.yellow}/>
            </div>
          </div>
          <div className="rh-grid-5">
            {[
              {l:"Masse brute",v:r.masse_brute},
              {l:"CNSS salarié",v:r.cnss_salarie},
              {l:"CNSS patronal",v:r.cnss_patronale},
              {l:"VPS",v:r.vps},
              {l:"ITS",v:r.its},
            ].map(x=>(
              <div key={x.l} style={{background:G.bg,borderRadius:"6px",padding:"8px 10px"}}>
                <div style={{fontSize:"10px",color:G.textDim,marginBottom:"3px"}}>{x.l}</div>
                <div style={{fontSize:"13px",fontWeight:700,color:G.accent}}>{fmt(x.v)}</div>
              </div>
            ))}
          </div>
          <div className="rh-item-footer">
            <div style={{background:G.red+"22",color:G.red,padding:"6px 14px",
                         borderRadius:"8px",fontSize:"13px",fontWeight:700}}>
              Total à verser : {fmt(r.cnss_salarie+r.cnss_patronale+r.vps+r.its)}
            </div>
          </div>
        </Card>
      ))}

      {!loading&&rapports.every(r=>r.nb===0)&&(
        <Card><div style={{color:G.textDim,textAlign:"center",padding:"24px"}}>
          Aucune fiche pour {MOIS[parseInt(mois)]} {annee}
        </div></Card>
      )}
    </div>
  );
};

// ─── HISTORIQUE ───────────────────────────────────────────────────────────────
const Historique = ({companies}) => {
  const [compId,setCompId]=useState(companies[0]?.id||"");
  const [annee,setAnnee]=useState(String(CURRENT_YEAR));
  const [data,setData]=useState([]);
  const [loading,setLoading]=useState(false);
  const [modele,setModele]=useState("officiel");
  const years=Array.from({length:5},(_,i)=>String(CURRENT_YEAR-i));
  const comp=companies.find(c=>c.id===compId);

  const load=useCallback(async()=>{
    if(!compId) return; setLoading(true);
    const {data:rows}=await supabase.from("payrolls")
      .select("*, employees(nom,prenoms,emploi,matricule,email,cnss,ifu,situation_matrimoniale,nb_enfants)")
      .eq("company_id",compId).eq("annee",parseInt(annee))
      .order("mois").order("created_at");
    setData(rows||[]); setLoading(false);
  },[compId,annee]);

  useEffect(()=>{load();},[load]);
  useEffect(()=>{if(!compId&&companies.length) setCompId(companies[0].id);},[companies]);

  const byMois=useMemo(()=>{
    const m={};
    (data||[]).forEach(p=>{
      if(!m[p.mois]) m[p.mois]={fiches:[],brut:0,net:0,cnss:0,its:0};
      m[p.mois].fiches.push(p);
      m[p.mois].brut+=parseFloat(p.salaire_brut||0);
      m[p.mois].net+=parseFloat(p.remuneration_due||0);
      m[p.mois].cnss+=parseFloat(p.cnss_ouvriere||0)+parseFloat(p.cnss_patronale||0);
      m[p.mois].its+=parseFloat(p.its||0);
    });
    return m;
  },[data]);

  const printFiche=(p,mod)=>{
    const emp=p.employees||{};
    const html=buildBulletinHTML(p,emp,comp,mod||modele);
    const win=window.open("","_blank");
    win.document.write(html); win.document.close();
    win.onload=()=>win.print();
  };

  return (
    <div>
      <div className="rh-page-header">
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Historique des Fiches de Paie</h2>
      </div>
      <Card style={{marginBottom:"16px",padding:"16px"}}>
        <div className="rh-flex-filters">
          <select className="rh-control" value={compId} onChange={e=>setCompId(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {companies.map(c=><option key={c.id} value={c.id}>{c.raison_sociale}</option>)}
          </select>
          <select value={annee} onChange={e=>setAnnee(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            {years.map(y=><option key={y} value={y}>{y}</option>)}
          </select>
          <select value={modele} onChange={e=>setModele(e.target.value)}
            style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"8px",
                    padding:"8px 12px",color:G.text,fontSize:"13px",fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            <option value="officiel">📄 Officiel</option>
            <option value="simple">📋 Simple</option>
            <option value="moderne">✨ Moderne</option>
          </select>
        </div>
      </Card>
      {loading&&<div style={{color:G.textDim,padding:"24px",textAlign:"center"}}>Chargement...</div>}
      {Object.keys(byMois).sort((a,b)=>parseInt(a)-parseInt(b)).map(m=>(
        <Card key={m} style={{marginBottom:"12px"}}>
          <div className="rh-item-header" style={{marginBottom:"12px",paddingBottom:"12px",borderBottom:`1px solid ${G.border}`}}>
            <div>
              <div style={{fontSize:"16px",fontWeight:700,color:G.accent}}>{MOIS[parseInt(m)]} {annee}</div>
              <div style={{fontSize:"12px",color:G.textDim,marginTop:"2px"}}>
                {byMois[m].fiches.length} fiche(s) — Brut : {fmt(byMois[m].brut)} — Net : {fmt(byMois[m].net)}
              </div>
            </div>
            <div className="rh-btn-group">
              <Badge label={"CNSS : "+fmt(byMois[m].cnss)} color={G.accent}/>
              <Badge label={"ITS : "+fmt(byMois[m].its)} color={G.yellow}/>
            </div>
          </div>
          <Table cols={[
            {label:"Employé",render:r=>r.employees?.nom+" "+r.employees?.prenoms},
            {label:"Emploi",render:r=>r.employees?.emploi||"",dim:true},
            {label:"Brut",render:r=>fmt(r.salaire_brut)},
            {label:"Net à payer",render:r=><span style={{color:G.accent,fontWeight:700}}>{fmt(r.remuneration_due)}</span>},
            {label:"",render:r=>(
              <div style={{display:"flex",gap:"6px"}}>
                <Btn small onClick={e=>{e.stopPropagation();printFiche(r);}} variant="ghost">🖨 PDF</Btn>
              </div>
            )}
          ]} rows={byMois[m].fiches}/>
        </Card>
      ))}
      {!loading&&Object.keys(byMois).length===0&&(
        <Card><div style={{color:G.textDim,textAlign:"center",padding:"24px"}}>
          Aucune fiche pour {annee}
        </div></Card>
      )}
    </div>
  );
};

// SETTINGS
const Settings = ({user}) => {
 const [form,setForm]=useState({smtp_host:"",smtp_port:"587",
    smtp_user:"",smtp_pass:"",email_from:""});
  const [msg,setMsg]=useState("");
  useEffect(()=>{
    (async()=>{
      const {data}=await supabase.from("email_settings").select("*").eq("user_id",user.id).maybeSingle();
      if(data) setForm({...data,smtp_port:String(data.smtp_port||587)});
    })();
  },[user]);
  const save=async()=>{
    const {error}=await supabase.from("email_settings").upsert({
      ...form,user_id:user.id,smtp_port:parseInt(form.smtp_port||587),
      updated_at:new Date().toISOString()},{onConflict:"user_id"});
    setMsg(error?error.message:"Paramètres enregistrés !");
  };
  const f=(k,v)=>setForm(p=>({...p,[k]:v}));
  return (
    <div>
      <h2 style={{marginBottom:"24px",fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Paramètres</h2>
      <Card style={{maxWidth:"500px"}}>
        <div style={{fontSize:"14px",fontWeight:600,marginBottom:"16px",color:G.accent}}>
          ⚙️ Configuration WhatsApp — Expéditeur
        </div>
        <div style={{background:"#f0f4ff",borderRadius:"8px",padding:"12px",marginBottom:"16px",
                     fontSize:"12px",color:G.textDim,border:"1px solid #d0dff5"}}>
          💡 Renseignez les coordonnées de l'expéditeur qui apparaîtront dans les messages WhatsApp envoyés aux employés.
        </div>
        <Alert msg={msg} type={msg.includes("enregistr")?"success":"error"}/>
        <Input label="Numéro WhatsApp expéditeur (avec indicatif pays)" 
          value={form.smtp_host} onChange={v=>f("smtp_host",v)} 
          placeholder="22997000000"/>
        <Input label="Nom affiché (expéditeur)" 
          value={form.smtp_user} onChange={v=>f("smtp_user",v)} 
          placeholder="PSARIZ SARL - Service RH"/>
        <Input label="Message personnalisé (optionnel)" 
          value={form.smtp_pass} onChange={v=>f("smtp_pass",v)} 
          placeholder="Pour toute question, contactez le service RH."/>
        <div style={{background:"#f0f4ff",borderRadius:"8px",padding:"12px",marginBottom:"16px",
                     fontSize:"12px",color:"#1a3a6b",border:"1px solid #d0dff5"}}>
          📱 <strong>Numéro employé :</strong> renseignez le numéro WhatsApp de chaque employé dans le champ <strong>"Email"</strong> de sa fiche (ex: 22997000000).
        </div>
        <Btn onClick={save}>Enregistrer</Btn>
      </Card>
    </div>
  );
};

// ─── GÉNÉRATEURS HTML POUR IMPRESSION ─────────────────────────────────────────

// MODÈLE 1 — OFFICIEL (fidèle au template slip_pdf.html)
const buildBulletinOfficiel = (p,emp,comp) => {
  const moisNb = parseInt(p.mois||1);
  const joursFin = [4,6,9,11].includes(moisNb)?30:moisNb===2?28:31;
  const brutFiscal = Math.floor((parseFloat(p.salaire_brut||0))/1000)*1000;
  const logo = comp?.logo_url ? `<img src="${comp.logo_url}" style="max-height:50px;max-width:120px;object-fit:contain;margin-bottom:6px;display:block">` : "";
  const notes = p.notes ? `<div style="background:#f0f0f0;border-left:3px solid #888;padding:5px 10px;font-size:9.5pt;margin-bottom:4px"><strong>Observations :</strong> ${p.notes}</div>` : "";
  return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<style>
@page {
  size: 21cm 29.7cm;
  margin-top:    1.9cm;
  margin-bottom: 1.9cm;
  margin-left:   1.8cm;
  margin-right:  1.8cm;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body {
  font-family: 'Times New Roman', Times, serif;
  font-size: 11pt;
  color: #1a1a1a;
  height: 100%;
}
.page-wrapper {
  display: flex;
  flex-direction: column;
  min-height: calc(29.7cm - 3.8cm);
  height: calc(29.7cm - 3.8cm);
}
.titre-bulletin {
  text-align: center;
  font-size: 14pt;
  font-weight: bold;
  letter-spacing: 2px;
  text-transform: uppercase;
  border: 2px solid #888;
  padding: 6px;
  margin-bottom: 8px;
  background: #d6d6d6;
  color: #000000;
}
.header-grid {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 0 12px;
  margin-bottom: 8px;
  border: 1px solid #ccc;
  padding: 8px 10px;
}
.period-col {
  text-align: right;
  display: flex;
  flex-direction: column;
  justify-content: center;
  min-width: 140px;
}
.info-row { display: flex; gap: 6px; font-size: 10.5pt; margin-bottom: 2px; }
.info-lbl { color: #444; min-width: 90px; }
.info-val { font-weight: 600; }
.period-badge {
  background: #d6d6d6;
  color: #000000;
  font-weight: bold;
  padding: 7px 14px;
  text-align: center;
  font-size: 12pt;
  letter-spacing: 0.5px;
  border-radius: 3px;
  border: 1.5px solid #888;
  margin-bottom: 4px;
}
.periode-line { font-size: 9pt; color: #444; text-align: center; }
.emp-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 2px 16px;
  border: 1px solid #ccc;
  padding: 7px 10px;
  margin-bottom: 8px;
  background: #fafafa;
}
.emp-row { display: flex; gap: 6px; font-size: 10.5pt; padding: 1px 0; }
.emp-lbl { color: #444; min-width: 110px; }
.emp-val { font-weight: 600; }
table { width: 100%; border-collapse: collapse; font-size: 10.5pt; margin-bottom: 0; }
th, td { border: 1px solid #bbb; padding: 4px 8px; vertical-align: middle; }
th { font-weight: bold; text-align: left; }
td.text-right { text-align: right; }
.th-section { background: #c8c8c8; color: #000; font-weight: bold; font-size: 10pt; letter-spacing: 0.5px; text-transform: uppercase; border-color: #888; }
.th-montant { width: 150px; text-align: right; }
.row-normal td { background: #ffffff; }
.row-alt    td { background: #f5f5f5; }
.row-hsup-sub td { background: #fafafa; font-size: 9.5pt; color: #555; padding: 3px 8px 3px 20px; border-top: none; }
.row-brut td { background: #e8e8e8; font-weight: bold; font-size: 11pt; border-top: 1.5px solid #888; color: #000; }
.row-net  td { background: #e8e8e8; font-weight: bold; font-size: 11pt; border-top: 1.5px solid #888; color: #000; }
.net-box {
  background: #d6d6d6;
  color: #000000;
  font-weight: bold;
  padding: 10px 14px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin: 6px 0;
  border: 2px solid #888;
}
.net-label  { font-size: 12pt; font-weight: bold; letter-spacing: 0.5px; color: #000; }
.net-amount { font-size: 16pt; font-weight: bold; color: #000; }
.signatures { display: flex; justify-content: space-between; margin-top: auto; padding-top: 10px; }
.sig-box  { width: 44%; text-align: center; }
.sig-space { height: 36px; }
.sig-line { border-top: 1px solid #666; padding-top: 4px; font-size: 9.5pt; color: #333; }
.sig-name { font-weight: bold; font-size: 10pt; }
.sig-func { font-size: 8.5pt; color: #666; }
.sep { height: 4px; }
@media print {
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
</style>
</head>
<body>
<div class="page-wrapper">

  <div class="titre-bulletin">BULLETIN DE PAIE</div>

  <div class="header-grid">
    <div class="company-col">
      ${logo}
      <div class="info-row"><span class="info-lbl">Raison sociale :</span><span class="info-val">${comp?.raison_sociale||'—'}</span></div>
      <div class="info-row"><span class="info-lbl">RCCM :</span><span class="info-val">${comp?.rccm||'—'}</span></div>
      <div class="info-row"><span class="info-lbl">Adresse :</span><span class="info-val">${comp?.adresse||'—'}</span></div>
      <div class="info-row"><span class="info-lbl">Tél :</span><span class="info-val">${comp?.tel||'—'}</span></div>
      <div class="info-row"><span class="info-lbl">Email :</span><span class="info-val">${comp?.email||'—'}</span></div>
    </div>
    <div class="period-col">
      <div class="period-badge">${MOIS[moisNb]} ${p.annee}</div>
      <div class="periode-line">Période : 01 au ${joursFin} ${MOIS[moisNb]} ${p.annee}</div>
    </div>
  </div>

  <div class="emp-grid">
    <div class="emp-row"><span class="emp-lbl">Nom :</span><span class="emp-val">${emp.nom||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">Emploi :</span><span class="emp-val">${emp.emploi||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">Prénoms :</span><span class="emp-val">${emp.prenoms||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">Sit. matrimoniale :</span><span class="emp-val">${emp.situation_matrimoniale||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">Matricule :</span><span class="emp-val">${emp.matricule||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">N° CNSS :</span><span class="emp-val">${emp.cnss||'—'}</span></div>
    <div class="emp-row"><span class="emp-lbl">Nb enfants :</span><span class="emp-val">${emp.nb_enfants||0}</span></div>
    <div class="emp-row"><span class="emp-lbl">IFU :</span><span class="emp-val">${emp.ifu||'—'}</span></div>
  </div>

  <table>
    <thead>
      <tr>
        <th class="th-section">Éléments de rémunération</th>
        <th class="th-section th-montant">Montant (FCFA)</th>
      </tr>
    </thead>
    <tbody>
      <tr class="row-normal">
        <td>Salaire de base</td>
        <td class="text-right">${fmtN(p.salaire_base||0)}</td>
      </tr>
      <tr class="row-alt">
        <td>Heures supplémentaires</td>
        <td class="text-right">${fmtN(p.montant_heures_sup||0)}</td>
      </tr>
      <tr class="row-hsup-sub">
        <td colspan="2">&nbsp;&nbsp;↳ &nbsp;
          à +12% : ${fmtN(p.h_sup_12||0)} h &emsp;
          à +35% : ${fmtN(p.h_sup_35||0)} h &emsp;
          à +50% : ${fmtN(p.h_sup_50||0)} h &emsp;
          à +100% : ${fmtN(p.h_sup_100||0)} h
        </td>
      </tr>
      <tr class="row-normal">
        <td>Primes diverses</td>
        <td class="text-right">${fmtN(p.primes||0)}</td>
      </tr>
      <tr class="row-alt">
        <td>Indemnités diverses</td>
        <td class="text-right">${fmtN(p.indemnites||0)}</td>
      </tr>
      <tr class="row-brut">
        <td><strong>SALAIRE BRUT</strong></td>
        <td class="text-right"><strong>${fmtN(p.salaire_brut||0)}</strong></td>
      </tr>
    </tbody>
  </table>

  <div class="sep"></div>

  <table>
    <thead>
      <tr>
        <th class="th-section">Retenues obligatoires</th>
        <th class="th-section th-montant">Montant (FCFA)</th>
      </tr>
    </thead>
    <tbody>
      <tr class="row-normal">
        <td>* Cotisations ouvrières CNSS &nbsp;<span style="font-size:9pt;color:#555">(3,6% du brut)</span></td>
        <td class="text-right">${fmtN(p.cnss_ouvriere||0)}</td>
      </tr>
      <tr class="row-alt">
        <td>* IRPP-TS (ITS) &nbsp;<span style="font-size:9pt;color:#555">(base fiscale : ${fmtN(brutFiscal)} FCFA)</span></td>
        <td class="text-right">${fmtN(p.its||0)}</td>
      </tr>
      <tr class="row-normal">
        <td>* Taxe radiophonique / télévisuelle${parseInt(p.taxe_radio||0)===0?' <span style="font-size:9pt;color:#666">&nbsp;— Non applicable</span>':''}</td>
        <td class="text-right">${fmtN(p.taxe_radio||0)}</td>
      </tr>
      <tr class="row-net">
        <td><strong>SALAIRE NET</strong></td>
        <td class="text-right"><strong>${fmtN(p.salaire_net||0)}</strong></td>
      </tr>
    </tbody>
  </table>

  <div class="sep"></div>

  <table>
    <thead>
      <tr>
        <th class="th-section">Retenues facultatives</th>
        <th class="th-section th-montant">Montant (FCFA)</th>
      </tr>
    </thead>
    <tbody>
      <tr class="row-normal">
        <td>* Avances et acomptes</td>
        <td class="text-right">${fmtN(p.avances||0)}</td>
      </tr>
      <tr class="row-alt">
        <td>* Opposition / Saisie-arrêt</td>
        <td class="text-right">${fmtN(p.saisie_arret||0)}</td>
      </tr>
      <tr class="row-normal">
        <td>* Assurance santé</td>
        <td class="text-right">${fmtN(p.assurance_sante||0)}</td>
      </tr>
    </tbody>
  </table>

  <div class="net-box">
    <span class="net-label">RÉMUNÉRATION NETTE À PAYER</span>
    <span class="net-amount">${fmtN(p.remuneration_due||0)} FCFA</span>
  </div>

  ${notes}

  <div class="signatures">
    <div class="sig-box">
      <div class="sig-space"></div>
      <div class="sig-line">
        <div class="sig-name">${comp?.signataire||'Signature de l\'employeur'}</div>
        <div class="sig-func">${comp?.fonction_signataire||''}</div>
      </div>
    </div>
    <div class="sig-box">
      <div class="sig-space"></div>
      <div class="sig-line">
        <div style="font-size:9.5pt;">Signature de l'employé(e)</div>
        <div style="font-size:9pt;color:#666;">${emp.nom||''} ${emp.prenoms||''}</div>
      </div>
    </div>
  </div>

</div>
</body>
</html>`;
};

// MODÈLE 2 — SIMPLE (compact, minimaliste)
const buildBulletinSimple = (p,emp,comp) => `
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Bulletin de Paie</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Arial,sans-serif;font-size:12px;color:#333;padding:24px;max-width:600px;margin:auto}
.header{text-align:center;margin-bottom:16px}
.header h1{font-size:18px;font-weight:bold;color:#222}
.header p{font-size:11px;color:#666}
.company{background:#f5f5f5;padding:10px;border-radius:4px;margin-bottom:12px;font-size:11px}
.emp{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:12px;font-size:11px}
.emp span{color:#666}
table{width:100%;border-collapse:collapse;margin-bottom:8px}
td{padding:6px 8px;border-bottom:1px solid #eee}
td:last-child{text-align:right;font-weight:500}
.cat{background:#f0f0f0;font-weight:bold;font-size:11px;text-transform:uppercase;letter-spacing:0.5px}
.subtotal{background:#e8e8e8;font-weight:bold}
.net{background:#222;color:#fff;font-size:14px;font-weight:bold}
.net td{color:#fff;padding:10px 8px}
.sign{display:flex;justify-content:space-between;margin-top:24px;font-size:11px}
.sign div{text-align:center;width:40%}
.sign .line{border-top:1px solid #999;padding-top:6px;margin-top:30px}
@media print{body{padding:0}}
</style></head><body>
<div class="header">
  <h1>BULLETIN DE PAIE</h1>
  <p>${MOIS[p.mois]} ${p.annee}</p>
</div>
<div class="company">
  <strong>${comp?.raison_sociale||""}</strong> — RCCM : ${comp?.rccm||""}<br>
  ${comp?.adresse||""} | Tél : ${comp?.tel||""}
</div>
<div class="emp">
  <div><span>Employé : </span><strong>${emp.nom||""} ${emp.prenoms||""}</strong></div>
  <div><span>Emploi : </span><strong>${emp.emploi||""}</strong></div>
  <div><span>Matricule : </span>${emp.matricule||"—"}</div>
  <div><span>N° CNSS : </span>${emp.cnss||""}</div>
</div>
<table>
<tr><td class="cat" colspan="2">Gains</td></tr>
<tr><td>Salaire de base</td><td>${fmtN(p.salaire_base||0)}</td></tr>
<tr><td>Heures supplémentaires</td><td>${fmtN(p.montant_heures_sup||0)}</td></tr>
<tr><td>Primes diverses</td><td>${fmtN(p.primes||0)}</td></tr>
<tr><td>Indemnités diverses</td><td>${fmtN(p.indemnites||0)}</td></tr>
<tr class="subtotal"><td>Salaire Brut</td><td>${fmtN(p.salaire_brut||0)}</td></tr>
<tr><td class="cat" colspan="2">Retenues obligatoires</td></tr>
<tr><td>CNSS ouvrière (3,6%)</td><td>- ${fmtN(p.cnss_ouvriere||0)}</td></tr>
<tr><td>ITS</td><td>- ${fmtN(p.its||0)}</td></tr>
<tr><td>Taxe radio</td><td>- ${fmtN(p.taxe_radio||0)}</td></tr>
<tr class="subtotal"><td>Salaire Net</td><td>${fmtN(p.salaire_net||0)}</td></tr>
<tr><td class="cat" colspan="2">Retenues facultatives</td></tr>
<tr><td>Avances et acomptes</td><td>- ${fmtN(p.avances||0)}</td></tr>
<tr><td>Opposition / Saisie-arrêt</td><td>- ${fmtN(p.saisie_arret||0)}</td></tr>
<tr><td>Assurance santé</td><td>- ${fmtN(p.assurance_sante||0)}</td></tr>
<tr class="net"><td>Net à payer</td><td>${fmtN(p.remuneration_due||0)} FCFA</td></tr>
</table>
<div class="sign">
  <div><div class="line">L'Employeur<br><small>${comp?.signataire||""}</small></div></div>
  <div><div class="line">L'Employé(e)<br><small>${emp.nom||""} ${emp.prenoms||""}</small></div></div>
</div>
</body></html>`;

// MODÈLE 3 — MODERNE (coloré, avec logo)
const buildBulletinModerne = (p,emp,comp) => `
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Bulletin de Paie</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Arial,sans-serif;font-size:11px;color:#333;padding:20px;max-width:750px;margin:auto}
.header{background:#1a1a2e;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0;
        display:flex;justify-content:space-between;align-items:center}
.header-left{display:flex;align-items:center;gap:12px}
.header-logo{height:48px;width:48px;background:#fff;border-radius:6px;
             display:flex;align-items:center;justify-content:center;overflow:hidden}
.header-logo img{height:44px;object-fit:contain}
.header-title{font-size:20px;font-weight:800;color:#00c896}
.header-sub{font-size:11px;color:#aaa;margin-top:2px}
.header-right{text-align:right}
.badge{background:#00c896;color:#fff;padding:4px 14px;border-radius:20px;
       font-weight:700;font-size:13px;display:inline-block}
.period{font-size:10px;color:#aaa;margin-top:4px}
.company-bar{background:#f8f9fa;border:1px solid #e0e0e0;border-top:none;
             padding:10px 16px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.company-bar .item{font-size:10px;color:#666}
.company-bar .item strong{color:#333;display:block}
.emp-card{border:1px solid #e0e0e0;border-top:none;padding:12px 16px;
          display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.emp-field .lbl{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.3px}
.emp-field .val{font-weight:700;color:#1a1a2e;font-size:12px}
.section{margin-top:0}
.section-title{background:#1a1a2e;color:#fff;padding:6px 16px;
               display:flex;justify-content:space-between;font-weight:600;font-size:11px;
               border-top:2px solid #00c896}
.line{display:flex;justify-content:space-between;padding:5px 16px;
      border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;
      border-bottom:1px solid #f0f0f0}
.line:hover{background:#f9f9f9}
.line-sub{display:flex;justify-content:space-between;padding:3px 16px 3px 28px;
          border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;
          border-bottom:1px solid #f5f5f5;color:#888;font-size:10px}
.subtotal{display:flex;justify-content:space-between;padding:6px 16px;
          background:#e8f5f1;border:1px solid #00c896;font-weight:700;color:#00a87c}
.net-box{background:#1a1a2e;color:#fff;padding:14px 16px;border-radius:0 0 8px 8px;
         display:flex;justify-content:space-between;align-items:center;margin-top:2px}
.net-label{font-size:13px;font-weight:600;color:#aaa}
.net-amount{font-size:22px;font-weight:900;color:#00c896}
.charges{background:#fff8e1;border:1px solid #ffd54f;border-radius:6px;
         padding:10px 16px;margin-top:12px;font-size:10px;color:#666}
.charges strong{color:#f57f17}
.sign{display:flex;justify-content:space-between;margin-top:20px}
.sign-box{width:44%;border:1px solid #e0e0e0;border-radius:6px;padding:10px;text-align:center}
.sign-box .title{font-size:10px;color:#888;margin-bottom:24px}
.sign-box .name{font-size:11px;font-weight:600;border-top:1px solid #ccc;padding-top:6px}
@media print{body{padding:0}.header{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>
<div class="header">
  <div class="header-left">
    ${comp?.logo_url?`<div class="header-logo"><img src="${comp.logo_url}" alt="logo"/></div>`:""}
    <div>
      <div class="header-title">BULLETIN DE PAIE</div>
      <div class="header-sub">${comp?.raison_sociale||""}</div>
    </div>
  </div>
  <div class="header-right">
    <div class="badge">${MOIS[p.mois]} ${p.annee}</div>
    <div class="period">RCCM : ${comp?.rccm||""}</div>
  </div>
</div>
<div class="company-bar">
  <div class="item"><span>Adresse</span><strong>${comp?.adresse||""}</strong></div>
  <div class="item"><span>Téléphone</span><strong>${comp?.tel||""}</strong></div>
  <div class="item"><span>Email</span><strong>${comp?.email||""}</strong></div>
</div>
<div class="emp-card">
  <div class="emp-field"><div class="lbl">Nom & Prénoms</div><div class="val">${emp.nom||""} ${emp.prenoms||""}</div></div>
  <div class="emp-field"><div class="lbl">Emploi</div><div class="val">${emp.emploi||""}</div></div>
  <div class="emp-field"><div class="lbl">Sit. Matrimoniale</div><div class="val">${emp.situation_matrimoniale||""}</div></div>
  <div class="emp-field"><div class="lbl">Matricule</div><div class="val">${emp.matricule||"—"}</div></div>
  <div class="emp-field"><div class="lbl">N° CNSS</div><div class="val">${emp.cnss||""}</div></div>
  <div class="emp-field"><div class="lbl">IFU</div><div class="val">${emp.ifu||""}</div></div>
</div>
<div class="section">
  <div class="section-title"><span>ÉLÉMENTS DE RÉMUNÉRATION</span><span>MONTANT (FCFA)</span></div>
  <div class="line"><span>Salaire de base</span><span>${fmtN(p.salaire_base||0)}</span></div>
  <div class="line"><span>Heures supplémentaires</span><span>${fmtN(p.montant_heures_sup||0)}</span></div>
  ${(p.h_sup_12>0||p.h_sup_35>0||p.h_sup_50>0||p.h_sup_100>0)?`<div class="line-sub"><span>↳ +12%: ${p.h_sup_12||0}h | +35%: ${p.h_sup_35||0}h | +50%: ${p.h_sup_50||0}h | +100%: ${p.h_sup_100||0}h</span></div>`:""}
  <div class="line"><span>Primes diverses</span><span>${fmtN(p.primes||0)}</span></div>
  <div class="line"><span>Indemnités diverses</span><span>${fmtN(p.indemnites||0)}</span></div>
  <div class="subtotal"><span>SALAIRE BRUT</span><span>${fmtN(p.salaire_brut||0)}</span></div>
</div>
<div class="section">
  <div class="section-title"><span>RETENUES OBLIGATOIRES</span><span>MONTANT (FCFA)</span></div>
  <div class="line"><span>Cotisations ouvrières CNSS (3,6% du brut)</span><span>- ${fmtN(p.cnss_ouvriere||0)}</span></div>
  <div class="line"><span>IRPP-TS / ITS (base : ${fmtN(p.salaire_brut||0)} FCFA)</span><span>- ${fmtN(p.its||0)}</span></div>
  <div class="line"><span>Taxe radiophonique / télévisuelle${p.taxe_radio===0?" — Non applicable":""}</span><span>- ${fmtN(p.taxe_radio||0)}</span></div>
  <div class="subtotal"><span>SALAIRE NET</span><span>${fmtN(p.salaire_net||0)}</span></div>
</div>
<div class="section">
  <div class="section-title"><span>RETENUES FACULTATIVES</span><span>MONTANT (FCFA)</span></div>
  <div class="line"><span>Avances et acomptes</span><span>- ${fmtN(p.avances||0)}</span></div>
  <div class="line"><span>Opposition / Saisie-arrêt</span><span>- ${fmtN(p.saisie_arret||0)}</span></div>
  <div class="line"><span>Assurance santé</span><span>- ${fmtN(p.assurance_sante||0)}</span></div>
</div>
<div class="net-box">
  <div class="net-label">RÉMUNÉRATION NETTE À PAYER</div>
  <div class="net-amount">${fmtN(p.remuneration_due||0)} FCFA</div>
</div>
<div class="charges">
  <strong>Charges patronales :</strong> CNSS patronale (${((parseFloat(p.taux_cnss_patronale||0.194))*100).toFixed(1)}%) : ${fmtN(p.cnss_patronale||0)} FCFA
  ${p.vps_exonere?` | VPS : EXONÉRÉ`:`| VPS (4%) : ${fmtN(p.vps||0)} FCFA`}
</div>
<div class="sign">
  <div class="sign-box"><div class="title">Signature de l'employeur</div>
    <div class="name">${comp?.signataire||""}<br><small>${comp?.fonction_signataire||""}</small></div></div>
  <div class="sign-box"><div class="title">Signature de l'employé(e)</div>
    <div class="name">${emp.nom||""} ${emp.prenoms||""}</div></div>
</div>
</body></html>`;

// Sélecteur de modèle
const buildBulletinHTML = (p,emp,comp,modele="officiel") => {
  if(modele==="simple") return buildBulletinSimple(p,emp,comp);
  if(modele==="moderne") return buildBulletinModerne(p,emp,comp);
  return buildBulletinOfficiel(p,emp,comp);
};


const buildDeclarationHTML = (data,decl,comp,mois,annee) => `
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Déclaration Sociale</title>
<style>
  body{font-family:Arial,sans-serif;font-size:12px;padding:20px}
  h1{font-size:16px;margin-bottom:16px}
  .company{font-weight:bold;font-size:14px;margin-bottom:8px}
  table{width:100%;border-collapse:collapse;margin-bottom:16px}
  td,th{border:1px solid #ccc;padding:8px;text-align:left}
  th{background:#f0f0f0}
  .right{text-align:right;font-weight:bold}
  .total{background:#e8f5e9;font-weight:bold}
</style></head><body>
<div class="company">${comp?.raison_sociale||""}</div>
<div>RCCM : ${comp?.rccm||""} — ${comp?.adresse||""}</div>
<h1>DÉCLARATION SOCIALE — ${MOIS[parseInt(mois)]} ${annee}</h1>
<table>
<thead><tr><th>Rubrique</th><th class="right">Montant (FCFA)</th><th>Statut</th></tr></thead>
<tbody>
<tr><td>CNSS salarié (3,6%)</td><td class="right">${fmtN(data?.cnss_employe||0)}</td><td>${decl?.statut_cnss||"Non déclaré"}</td></tr>
<tr><td>CNSS patronale</td><td class="right">${fmtN(data?.cnss_employeur||0)}</td><td></td></tr>
<tr class="total"><td>Total CNSS</td><td class="right">${fmtN(data?.cnss_total||0)}</td><td>${decl?.statut_cnss||"—"}</td></tr>
<tr><td>VPS (4%)</td><td class="right">${fmtN(data?.vps||0)}</td><td>${decl?.statut_vps||"Non déclaré"}</td></tr>
<tr><td>ITS</td><td class="right">${fmtN(data?.its_total||0)}</td><td>${decl?.statut_its||"Non déclaré"}</td></tr>
<tr class="total"><td>TOTAL À VERSER</td>
  <td class="right">${fmtN((data?.cnss_total||0)+(data?.vps||0)+(data?.its_total||0))}</td><td></td></tr>
</tbody></table>
<div style="margin-top:40px;display:flex;justify-content:space-between">
  <div>Visa CNSS :<br><br><br>__________________</div>
  <div>Fait à __________, le __________<br>${comp?.signataire||""}<br>${comp?.fonction_signataire||""}</div>
</div>
</body></html>`;

// ─── APPLICATION PRINCIPALE ───────────────────────────────────────────────────
export default function App() {
  const [user,setUser]=useState(null);
  const [page,setPage]=useState("dashboard");
  const [companies,setCompanies]=useState([]);
  const [loading,setLoading]=useState(true);
  const [sidebarOpen,setSidebarOpen]=useState(false);

  // ── Vérification admin (accès illimité sans abonnement) ──
  const isAdmin = user?.email === ADMIN_EMAIL;

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{
      setUser(session?.user||null); setLoading(false);
    });
    const {data:{subscription}}=supabase.auth.onAuthStateChange((_,s)=>setUser(s?.user||null));
    return ()=>subscription.unsubscribe();
  },[]);

  const loadCompanies = useCallback(async()=>{
    if(!user) return;
    const {data}=await supabase.from("companies").select("*").eq("user_id",user.id).order("raison_sociale");
    setCompanies(data||[]);
  },[user]);

  useEffect(()=>{loadCompanies();},[loadCompanies]);

  const logout=async()=>{ await supabase.auth.signOut(); setUser(null); };

  if(loading) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1a3a6b,#1a6dd6)",display:"flex",alignItems:"center",
                 justifyContent:"center",color:"#ffffff",fontFamily:"system-ui",fontSize:"18px"}}>
      Chargement…
    </div>
  );

  if(!user) return <LoginPage onLogin={()=>loadCompanies()}/>;

  const pages={
    dashboard:<Dashboard companies={companies}/>,
    companies:<Companies companies={companies} reload={loadCompanies} userId={user.id}/>,
    employees:<Employees companies={companies}/>,
    payroll:<Payroll companies={companies}/>,
    historique:<Historique companies={companies}/>,
    declarations:<Declarations companies={companies}/>,
    rapport:<RapportCabinet companies={companies}/>,
    settings:<Settings user={user}/>,
  };


  return (
    <div style={{minHeight:"100vh",background:"#f0f4ff",color:"#1a2a4a",fontFamily:"system-ui,sans-serif"}}>
      <ToastContainer/>
      <FloatingContact/>
      {/* Mobile top bar */}
      <div className="rh-topbar">
        <div style={{fontSize:"16px",fontWeight:800,color:"#fff"}}>RH-Paie Pro</div>
        <button onClick={()=>setSidebarOpen(true)}
          style={{background:"none",border:"none",color:"#fff",fontSize:"24px",cursor:"pointer",
                  padding:"4px 8px",lineHeight:1}}>☰</button>
      </div>
      <Sidebar page={page} setPage={setPage} user={user} onLogout={logout} isAdmin={isAdmin}
               open={sidebarOpen} setOpen={setSidebarOpen}/>
      <div className="rh-main">
        {isAdmin && (
          <div style={{background:"linear-gradient(90deg,#1a3a6b,#1a6dd6)",color:"#fff",
                       borderRadius:"10px",padding:"8px 16px",marginBottom:"16px",
                       fontSize:"12px",fontWeight:600,display:"inline-flex",alignItems:"center",gap:"8px"}}>
            👑 Compte Admin — Accès illimité
          </div>
        )}
        {companies.length===0&&page!=="companies"&&(
          <div className="rh-warning-banner">
            ⚠️ Commencez par créer votre <strong
              style={{cursor:"pointer",textDecoration:"underline"}} onClick={()=>setPage("companies")}>
              première société
            </strong>.
          </div>
        )}
        {pages[page]}
      </div>
    </div>
  );
}

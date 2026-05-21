import { useState, useEffect, useCallback, useMemo } from "react";
import { createClient } from "@supabase/supabase-js";
import JSZip from "jszip";
import ExcelJS from "exceljs";

// ─── CONFIGURATION SUPABASE ───────────────────────────────────────────────────
// 🔴 Remplacez ces valeurs par vos vraies clés Supabase
const SUPABASE_URL  = "https://rducxnyxgqrmtyvqwfzw.supabase.co";
const SUPABASE_ANON = "sb_publishable_2K8GkCgsLbHzSy-1FhP2cg_cWIVD-vH";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);
const ADMIN_EMAIL = "martin13haya@gmail.com";
const EMAILJS_SERVICE  = "service_n7ugp1c";   // votre Service ID
const EMAILJS_TEMPLATE = "template_t7ro56h";  // votre Template ID
const EMAILJS_KEY      = "n2N8upDAe3RfPiddE";    // votre Public Key

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
// ─── ABONNEMENT ──────────────────────────────────────────────────────────────
const PLANS = {
  free:       {label:"Essai gratuit", color:"#f59e0b", maxComp:1},
  starter:    {label:"Starter",       color:"#3b82f6", maxComp:1},
  pro:        {label:"Pro",           color:"#8b5cf6", maxComp:3},
  enterprise: {label:"Enterprise",    color:"#10b981", maxComp:999},
};

const getPlanInfo = (sub) => {
  if(!sub) return {plan:"free", label:"Essai gratuit", color:"#f59e0b", maxComp:1, isActive:false, trialDaysLeft:0};
  const p = PLANS[sub.plan]||PLANS.free;
  const now = new Date();
  const trialEnd = sub.trial_ends_at ? new Date(sub.trial_ends_at) : null;
  const trialDaysLeft = trialEnd ? Math.max(0, Math.ceil((trialEnd-now)/(1000*60*60*24))) : 0;
  const isActive = sub.status==="active" || (sub.status==="trial" && trialDaysLeft>0);
  return {...p, plan:sub.plan||"free", isActive, trialDaysLeft, status:sub.status};
};

const SubscriptionBanner = ({sub, onUpgrade}) => {
  const info = getPlanInfo(sub);
  if(!sub || sub.status==="active") return null;
  const expired = sub.status==="trial" && info.trialDaysLeft===0;
  const bg = expired ? "#fee2e2" : info.trialDaysLeft<=3 ? "#fff3cd" : "#eff6ff";
  const color = expired ? "#dc2626" : info.trialDaysLeft<=3 ? "#92400e" : "#1d4ed8";
  const msg = expired
    ? "⛔ Votre période d'essai est terminée. Abonnez-vous pour continuer."
    : `⏳ Essai gratuit : ${info.trialDaysLeft} jour(s) restant(s).`;
  return (
    <div style={{background:bg,color,padding:"10px 20px",fontSize:"13px",
                 display:"flex",justifyContent:"space-between",alignItems:"center",
                 borderBottom:"1px solid rgba(0,0,0,0.08)"}}>
      <span>{msg}</span>
      <button onClick={onUpgrade}
        style={{background:color,color:"#fff",border:"none",borderRadius:"6px",
                padding:"5px 14px",cursor:"pointer",fontSize:"12px",fontWeight:600}}>
        Voir les plans
      </button>
    </div>
  );
};

const PricingPage = ({currentSub, userId, onClose}) => {
  const [billing,setBilling]=useState("monthly");
  const prices = {
    starter:  {monthly:1500,  annual:1200},
    pro:      {monthly:5000,  annual:4000},
    enterprise:{monthly:10000,annual:8000},
  };
  const pay = async (plan) => {
    if(typeof window.FedaPay==="undefined"){alert("FedaPay non chargé. Vérifiez votre connexion."); return;}
    window.FedaPay.init({public_key:"pk_live_Ca1JjsEAdsaJiGf29jSOdA91",
      transaction:{amount:prices[plan][billing],description:`RH-Paie Pro ${plan} - ${billing}`},
      customer:{email:currentSub?.email||""},
      onComplete:async(t)=>{
        if(t.reason==="APPROVED"){
          const endDate = new Date();
          endDate.setMonth(endDate.getMonth()+(billing==="annual"?12:1));
          await supabase.from("subscriptions").upsert({
            user_id:userId, plan, status:"active",
            billing_cycle:billing, expires_at:endDate.toISOString(),
            updated_at:new Date().toISOString()
          },{onConflict:"user_id"});
          toast.success("✅ Abonnement activé !");
          if(onClose) onClose();
        }
      }
    }).open();
  };
  const planCards = [
    {key:"starter", name:"Starter", desc:"1 société", icon:"🌱"},
    {key:"pro",     name:"Pro",     desc:"3 sociétés", icon:"🚀"},
    {key:"enterprise",name:"Enterprise",desc:"Illimité",icon:"🏢"},
  ];
  return (
    <div style={{padding:"24px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"24px"}}>
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>💳 Plans & Abonnements</h2>
        {onClose&&<Btn onClick={onClose} variant="ghost" small>✕ Fermer</Btn>}
      </div>
      <div style={{display:"flex",gap:"8px",marginBottom:"24px",background:"#f0f4ff",
                   padding:"4px",borderRadius:"8px",width:"fit-content"}}>
        {["monthly","annual"].map(b=>(
          <button key={b} onClick={()=>setBilling(b)}
            style={{padding:"6px 16px",borderRadius:"6px",border:"none",cursor:"pointer",
                    fontWeight:600,fontSize:"13px",transition:"all 0.15s",
                    background:billing===b?"#1a3a6b":"transparent",
                    color:billing===b?"#fff":"#64748b"}}>
            {b==="monthly"?"Mensuel":"Annuel"}{b==="annual"&&<span style={{color:"#10b981",marginLeft:"4px",fontSize:"11px"}}>-20%</span>}
          </button>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:"16px"}}>
        {planCards.map(({key,name,desc,icon})=>{
          const isCurrent = currentSub?.plan===key && currentSub?.status==="active";
          const price = prices[key][billing];
          return (
            <div key={key} style={{background:"#fff",borderRadius:"12px",padding:"24px",
                                   border:`2px solid ${isCurrent?"#1a3a6b":"#e2e8f0"}`,
                                   boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
              <div style={{fontSize:"28px",marginBottom:"8px"}}>{icon}</div>
              <div style={{fontSize:"18px",fontWeight:700,color:"#1a3a6b"}}>{name}</div>
              <div style={{fontSize:"12px",color:"#64748b",marginBottom:"12px"}}>{desc}</div>
              <div style={{fontSize:"24px",fontWeight:800,color:"#1a3a6b",marginBottom:"4px"}}>
                {price.toLocaleString("fr-FR")} <span style={{fontSize:"13px",fontWeight:400}}>FCFA/mois</span>
              </div>
              {isCurrent
                ? <div style={{background:"#d1fae5",color:"#065f46",padding:"8px",borderRadius:"6px",
                               textAlign:"center",fontSize:"12px",fontWeight:600}}>✅ Plan actuel</div>
                : <button onClick={()=>pay(key)}
                    style={{width:"100%",background:"#1a3a6b",color:"#fff",border:"none",
                            borderRadius:"8px",padding:"10px",cursor:"pointer",fontWeight:600,fontSize:"13px"}}>
                    Choisir {name}
                  </button>
              }
            </div>
          );
        })}
      </div>
    </div>
  );
};

// ─── ADMIN ────────────────────────────────────────────────────────────────────
const AdminPanel = () => {
  const [users,setUsers]=useState([]);
  const [loading,setLoading]=useState(true);
  const [search,setSearch]=useState("");

  useEffect(()=>{
    (async()=>{
      const {data}=await supabase.from("subscriptions").select("*").order("created_at",{ascending:false});
      setUsers(data||[]); setLoading(false);
    })();
  },[]);

  const updatePlan = async (userId, plan, status) => {
    await supabase.from("subscriptions").update({plan,status,updated_at:new Date().toISOString()}).eq("user_id",userId);
    setUsers(u=>u.map(x=>x.user_id===userId?{...x,plan,status}:x));
    toast.success("Plan mis à jour !");
  };

  const filtered = users.filter(u=>
    (u.user_id||"").toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"24px"}}>
        <h2 style={{fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>🛡️ Panneau Admin</h2>
        <div style={{background:"#fee2e2",color:"#dc2626",padding:"4px 12px",borderRadius:"20px",fontSize:"12px",fontWeight:600}}>ADMIN</div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:"12px",marginBottom:"24px"}}>
        {[
          {label:"Total utilisateurs",val:users.length,icon:"👤"},
          {label:"Actifs",val:users.filter(u=>u.status==="active").length,icon:"✅"},
          {label:"En essai",val:users.filter(u=>u.status==="trial").length,icon:"⏳"},
          {label:"Expirés",val:users.filter(u=>u.status==="expired"||(!u.status)).length,icon:"⛔"},
        ].map(s=>(
          <div key={s.label} style={{background:"#fff",borderRadius:"10px",padding:"16px",
                                     border:"1px solid #e2e8f0",textAlign:"center"}}>
            <div style={{fontSize:"24px"}}>{s.icon}</div>
            <div style={{fontSize:"22px",fontWeight:800,color:"#1a3a6b"}}>{s.val}</div>
            <div style={{fontSize:"11px",color:"#64748b"}}>{s.label}</div>
          </div>
        ))}
      </div>
      <Card>
        <Input label="Rechercher par User ID" value={search} onChange={setSearch} placeholder="UUID..."/>
        {loading ? <div style={{padding:"20px",textAlign:"center",color:"#64748b"}}>Chargement…</div> : (
          <div style={{overflowX:"auto",marginTop:"16px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {["User ID","Plan","Statut","Expire le","Facturation","Actions"].map(h=>(
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:600,
                                        color:"#64748b",borderBottom:"1px solid #e2e8f0"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map(u=>{
                  const info = getPlanInfo(u);
                  return (
                    <tr key={u.user_id} style={{borderBottom:"1px solid #f1f5f9"}}>
                      <td style={{padding:"10px 12px",fontFamily:"monospace",fontSize:"11px",color:"#64748b"}}>
                        {u.user_id?.substring(0,16)}…
                      </td>
                      <td style={{padding:"10px 12px"}}>
                        <span style={{background:info.color+"22",color:info.color,padding:"2px 8px",
                                      borderRadius:"20px",fontSize:"11px",fontWeight:600}}>
                          {info.label}
                        </span>
                      </td>
                      <td style={{padding:"10px 12px"}}>
                        <span style={{color:u.status==="active"?"#10b981":u.status==="trial"?"#f59e0b":"#ef4444",fontWeight:600}}>
                          {u.status||"—"}
                        </span>
                      </td>
                      <td style={{padding:"10px 12px",color:"#64748b"}}>
                        {u.trial_ends_at ? new Date(u.trial_ends_at).toLocaleDateString("fr-FR") :
                         u.expires_at   ? new Date(u.expires_at).toLocaleDateString("fr-FR") : "—"}
                      </td>
                      <td style={{padding:"10px 12px",color:"#64748b"}}>{u.billing_cycle||"—"}</td>
                      <td style={{padding:"10px 12px"}}>
                        <div style={{display:"flex",gap:"4px"}}>
                          {["starter","pro","enterprise"].map(p=>(
                            <button key={p} onClick={()=>updatePlan(u.user_id,p,"active")}
                              style={{padding:"3px 8px",fontSize:"10px",borderRadius:"4px",border:"1px solid #e2e8f0",
                                      cursor:"pointer",background:u.plan===p?"#1a3a6b":"#fff",
                                      color:u.plan===p?"#fff":"#1a3a6b",fontWeight:600}}>
                              {p}
                            </button>
                          ))}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {filtered.length===0&&<div style={{padding:"20px",textAlign:"center",color:"#64748b"}}>Aucun utilisateur trouvé</div>}
          </div>
        )}
      </Card>
    </div>
  );
};

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

const Sidebar = ({page,setPage,user,onLogout,open=false,setOpen=()=>{},sub,onUpgrade}) => {
  const info = getPlanInfo(sub);
  const isAdmin = user?.email===ADMIN_EMAIL;
  const nav = isAdmin ? [...NAV,{id:"admin",icon:"🛡️",label:"Admin"}] : NAV;
  return (
  <>
    <div className={`rh-overlay ${open?"open":""}`} onClick={()=>setOpen(false)}/>
    <aside className={`rh-sidebar ${open?"open":""}`} style={{background:G.sidebar,borderRight:`1px solid ${G.border}`}}>
      <div style={{padding:"24px 20px",borderBottom:`1px solid ${G.border}`,position:"relative"}}>
        <button className="rh-mobile-close" onClick={()=>setOpen(false)}
          style={{position:"absolute",top:"14px",right:"14px",background:"none",border:"none",color:"#fff",fontSize:"24px",cursor:"pointer",lineHeight:1}}>×</button>
        <div className="rh-sidebar-brand-text" style={{fontSize:"18px",fontWeight:800,color:"#ffffff",letterSpacing:"-0.5px"}}>RH-Paie Pro</div>
        <div className="rh-sidebar-brand-text" style={{fontSize:"11px",color:"#a0c0e8",marginTop:"3px"}}>Gestion de la Paie</div>
        {/* Badge plan */}
        <div onClick={onUpgrade}
          style={{marginTop:"8px",display:"inline-flex",alignItems:"center",gap:"5px",
                  background:info.color+"33",border:`1px solid ${info.color}66`,
                  borderRadius:"20px",padding:"3px 10px",cursor:"pointer"}}>
          <span style={{width:"7px",height:"7px",borderRadius:"50%",background:info.color,display:"inline-block"}}/>
          <span style={{fontSize:"11px",fontWeight:700,color:info.color}}>{info.label}</span>
          {info.status==="trial"&&info.trialDaysLeft>0&&
            <span style={{fontSize:"10px",color:"#fbbf24"}}>({info.trialDaysLeft}j)</span>}
        </div>
      </div>
      <nav style={{flex:1,padding:"16px 12px",overflowY:"auto"}}>
        {nav.map(n=>(
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
};
// ─── PAGES ────────────────────────────────────────────────────────────────────

// LOGIN
const LoginPage = ({onLogin}) => {
  const [email,setEmail]=useState(""), [pass,setPass]=useState(""),
        [err,setErr]=useState(""), [loading,setLoading]=useState(false),
        [isSignup,setIsSignup]=useState(false), [showForgot,setShowForgot]=useState(false),
        [showPass,setShowPass]=useState(false);

  const submit = async () => {
    if(!email||!pass){setErr("Veuillez remplir tous les champs."); return;}
    setErr(""); setLoading(true);
    try{
      const {error} = isSignup
        ? await supabase.auth.signUp({email,password:pass})
        : await supabase.auth.signInWithPassword({email,password:pass});
      if(error) setErr(error.message);
      else { if(isSignup) setErr("✅ Vérifiez votre email pour confirmer le compte."); else onLogin(); }
    }finally{setLoading(false);}
  };

  const isSuccess = err.startsWith("✅");

  const pageStyle={
    minHeight:"100vh",
    background:"linear-gradient(135deg,#0f2d6b 0%,#1a4fa8 55%,#2d6fd4 100%)",
    display:"flex",alignItems:"center",justifyContent:"center",
    fontFamily:"system-ui,sans-serif",padding:"20px",position:"relative",overflow:"hidden"
  };
  const cardStyle={
    background:"#ffffff",borderRadius:"20px",padding:"2.5rem 2.5rem 2rem",
    width:"100%",maxWidth:"420px",position:"relative",zIndex:1,
    boxShadow:"0 24px 64px rgba(0,0,0,0.18)"
  };
  const inputWrapStyle={position:"relative"};
  const inputStyle={
    width:"100%",padding:"11px 14px 11px 42px",
    border:`1.5px solid ${err&&!isSuccess?"#fca5a5":"#e2e8f0"}`,
    borderRadius:"10px",fontSize:"14px",color:"#1e293b",
    background:"#f8fafc",outline:"none",fontFamily:"inherit",
    transition:"border-color .2s, box-shadow .2s"
  };
  const iconStyle={
    position:"absolute",left:"13px",top:"50%",transform:"translateY(-50%)",
    color:"#94a3b8",fontSize:"17px",pointerEvents:"none"
  };
  const labelStyle={display:"block",fontSize:"13px",fontWeight:600,color:"#334155",marginBottom:"6px"};
  const btnPrimaryStyle={
    width:"100%",padding:"13px",
    background:loading?"#93c5fd":"linear-gradient(135deg,#1a4fa8,#2d6fd4)",
    color:"#fff",border:"none",borderRadius:"10px",fontSize:"15px",fontWeight:600,
    cursor:loading?"not-allowed":"pointer",letterSpacing:"0.2px",
    transition:"opacity .15s, transform .15s",fontFamily:"inherit"
  };

  return (
    <div style={pageStyle}>
      {/* déco cercles */}
      <div style={{position:"absolute",width:"500px",height:"500px",borderRadius:"50%",
                   background:"rgba(255,255,255,0.04)",top:"-130px",right:"-100px",zIndex:0}}/>
      <div style={{position:"absolute",width:"300px",height:"300px",borderRadius:"50%",
                   background:"rgba(255,255,255,0.04)",bottom:"-80px",left:"-70px",zIndex:0}}/>
      {/* grille */}
      <div style={{position:"absolute",inset:0,zIndex:0,
                   backgroundImage:"linear-gradient(rgba(255,255,255,0.03) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,0.03) 1px,transparent 1px)",
                   backgroundSize:"40px 40px"}}/>

      <div style={cardStyle}>
        {/* badge */}
        <div style={{position:"absolute",top:"-13px",right:"24px",
                     background:"linear-gradient(135deg,#1a4fa8,#2d6fd4)",color:"#fff",
                     fontSize:"11px",fontWeight:700,padding:"4px 14px",
                     borderRadius:"20px",letterSpacing:"0.3px"}}>
          Bénin / UEMOA
        </div>

        {/* logo */}
        <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"1.75rem"}}>
          <div style={{width:"44px",height:"44px",background:"linear-gradient(135deg,#1a4fa8,#2d6fd4)",
                       borderRadius:"12px",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="2" y="7" width="20" height="14" rx="2"/>
              <path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/>
              <line x1="12" y1="12" x2="12" y2="16"/>
              <line x1="10" y1="14" x2="14" y2="14"/>
            </svg>
          </div>
          <div>
            <div style={{fontSize:"18px",fontWeight:800,color:"#0f2d6b",letterSpacing:"-0.3px"}}>RH-Paie Pro</div>
            <div style={{fontSize:"12px",color:"#64748b",marginTop:"1px"}}>Gestion de la Paie</div>
          </div>
        </div>

        {/* titre */}
        <div style={{marginBottom:"1.5rem"}}>
          <h2 style={{fontSize:"22px",fontWeight:700,color:"#0f2d6b",marginBottom:"4px"}}>
            {isSignup?"Créer un compte":"Bienvenue 👋"}
          </h2>
          <p style={{fontSize:"14px",color:"#64748b"}}>
            {isSignup?"Rejoignez RH-Paie Pro dès aujourd'hui":"Connectez-vous à votre espace de gestion"}
          </p>
        </div>

        {/* alerte */}
        {err&&(
          <div style={{background:isSuccess?"#f0fdf4":"#eff6ff",
                       border:`1px solid ${isSuccess?"#bbf7d0":"#bfdbfe"}`,
                       borderRadius:"8px",padding:"9px 13px",fontSize:"13px",
                       color:isSuccess?"#166534":"#1e40af",marginBottom:"1rem",
                       display:"flex",alignItems:"flex-start",gap:"8px"}}>
            <span style={{flexShrink:0}}>{isSuccess?"✅":"ℹ️"}</span>
            <span>{err.replace("✅ ","")}</span>
          </div>
        )}

        {/* champ email */}
        <div style={{marginBottom:"1rem"}}>
          <label style={labelStyle}>Adresse email</label>
          <div style={inputWrapStyle}>
            <span style={iconStyle}>✉</span>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)}
              placeholder="exemple@email.com" style={inputStyle}
              onFocus={e=>{e.target.style.borderColor="#1a4fa8";e.target.style.background="#fff";e.target.style.boxShadow="0 0 0 3px rgba(26,79,168,0.1)";}}
              onBlur={e=>{e.target.style.borderColor="#e2e8f0";e.target.style.background="#f8fafc";e.target.style.boxShadow="none";}}
              onKeyDown={e=>{if(e.key==="Enter") submit();}}
            />
          </div>
        </div>

        {/* champ mot de passe */}
        <div style={{marginBottom:"0.75rem"}}>
          <label style={labelStyle}>Mot de passe</label>
          <div style={inputWrapStyle}>
            <span style={iconStyle}>🔒</span>
            <input type={showPass?"text":"password"} value={pass} onChange={e=>setPass(e.target.value)}
              placeholder="••••••••" style={{...inputStyle,paddingRight:"42px"}}
              onFocus={e=>{e.target.style.borderColor="#1a4fa8";e.target.style.background="#fff";e.target.style.boxShadow="0 0 0 3px rgba(26,79,168,0.1)";}}
              onBlur={e=>{e.target.style.borderColor="#e2e8f0";e.target.style.background="#f8fafc";e.target.style.boxShadow="none";}}
              onKeyDown={e=>{if(e.key==="Enter") submit();}}
            />
            <button onClick={()=>setShowPass(p=>!p)}
              style={{position:"absolute",right:"12px",top:"50%",transform:"translateY(-50%)",
                      background:"none",border:"none",cursor:"pointer",color:"#94a3b8",
                      fontSize:"16px",padding:"2px",lineHeight:1}}>
              {showPass?"🙈":"👁"}
            </button>
          </div>
        </div>

        {/* mot de passe oublié */}
        {!isSignup&&(
          <div style={{textAlign:"right",marginBottom:"1.4rem"}}>
            <button onClick={()=>setShowForgot(true)}
              style={{background:"none",border:"none",color:"#1a4fa8",cursor:"pointer",
                      fontSize:"13px",fontWeight:500,fontFamily:"inherit",textDecoration:"none"}}>
              Mot de passe oublié ?
            </button>
          </div>
        )}

        {/* bouton connexion */}
        <button onClick={submit} disabled={loading} style={btnPrimaryStyle}
          onMouseOver={e=>{if(!loading){e.target.style.opacity="0.9";e.target.style.transform="translateY(-1px)";}}}
          onMouseOut={e=>{e.target.style.opacity="1";e.target.style.transform="translateY(0)";}}>
          {loading?"⏳ Connexion en cours…":(isSignup?"Créer mon compte":"Se connecter")}
        </button>

        {/* divider */}
        <div style={{display:"flex",alignItems:"center",gap:"10px",margin:"1.2rem 0"}}>
          <div style={{flex:1,height:"1px",background:"#e2e8f0"}}/>
          <span style={{fontSize:"13px",color:"#94a3b8"}}>ou</span>
          <div style={{flex:1,height:"1px",background:"#e2e8f0"}}/>
        </div>

        {/* switch signup/login */}
        <button onClick={()=>{setIsSignup(s=>!s);setErr("");}}
          style={{width:"100%",padding:"11px",background:"#f1f5f9",color:"#475569",
                  border:"1.5px solid #e2e8f0",borderRadius:"10px",fontSize:"14px",
                  fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                  transition:"background .15s"}}
          onMouseOver={e=>e.target.style.background="#e2e8f0"}
          onMouseOut={e=>e.target.style.background="#f1f5f9"}>
          {isSignup?"Déjà un compte ? Se connecter":"Pas de compte ? S'inscrire"}
        </button>

        {showForgot&&<ForgotPasswordModal onClose={()=>setShowForgot(false)}/>}
      </div>
    </div>
  );
};

// FORGOT PASSWORD MODAL
const ForgotPasswordModal = ({onClose}) => {
  const [email,setEmail]=useState("");
  const [loading,setLoading]=useState(false);
  const [sent,setSent]=useState(false);
  const [err,setErr]=useState("");

  const handle = async () => {
    if(!email){setErr("Veuillez entrer votre email."); return;}
    setLoading(true); setErr("");
    const {error}=await supabase.auth.resetPasswordForEmail(email,{
      redirectTo: window.location.origin+window.location.pathname
    });
    setLoading(false);
    if(error) setErr(error.message); else setSent(true);
  };

  return (
    <Modal title="🔑 Mot de passe oublié" onClose={onClose} width="400px">
      {sent
        ?<div style={{textAlign:"center",padding:"16px 0"}}>
          <div style={{fontSize:"36px",marginBottom:"12px"}}>📧</div>
          <div style={{fontWeight:700,fontSize:"15px",color:G.text,marginBottom:"8px"}}>Email envoyé !</div>
          <div style={{fontSize:"13px",color:G.textDim,marginBottom:"20px"}}>
            Vérifiez votre boîte mail et cliquez le lien de réinitialisation.
          </div>
          <Btn onClick={onClose}>Fermer</Btn>
        </div>
        :<>
          <div style={{fontSize:"13px",color:G.textDim,marginBottom:"16px"}}>
            Entrez votre email — vous recevrez un lien pour créer un nouveau mot de passe.
          </div>
          <Alert msg={err} type="error"/>
          <Input label="Email" value={email} onChange={setEmail} type="email" placeholder="votre@email.com"/>
          <div style={{display:"flex",gap:"8px",marginTop:"8px"}}>
            <Btn onClick={handle} disabled={loading}>{loading?"Envoi...":"Envoyer le lien"}</Btn>
            <Btn onClick={onClose} variant="ghost">Annuler</Btn>
          </div>
        </>
      }
    </Modal>
  );
};

// RESET PASSWORD PAGE
const ResetPasswordPage = ({onDone}) => {
  const [pass,setPass]=useState("");
  const [confirm,setConfirm]=useState("");
  const [loading,setLoading]=useState(false);
  const [err,setErr]=useState("");
  const [ok,setOk]=useState(false);

  const handle = async () => {
    if(!pass){setErr("Entrez un nouveau mot de passe."); return;}
    if(pass.length<6){setErr("Minimum 6 caractères."); return;}
    if(pass!==confirm){setErr("Les mots de passe ne correspondent pas."); return;}
    setLoading(true); setErr("");
    const {error}=await supabase.auth.updateUser({password:pass});
    setLoading(false);
    if(error) setErr(error.message);
    else{setOk(true); setTimeout(onDone,2000);}
  };

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1a3a6b 0%,#1a6dd6 50%,#63a8f0 100%)",
                 display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif"}}>
      <div style={{width:"100%",maxWidth:"380px",padding:"20px"}}>
        <div style={{textAlign:"center",marginBottom:"32px"}}>
          <div style={{fontSize:"32px",fontWeight:900,color:"#fff",letterSpacing:"-1px"}}>RH-Paie Pro</div>
          <div style={{color:"rgba(255,255,255,0.75)",fontSize:"14px",marginTop:"6px"}}>Réinitialisation du mot de passe</div>
        </div>
        <Card>
          {ok
            ?<div style={{textAlign:"center",padding:"16px 0"}}>
              <div style={{fontSize:"36px",marginBottom:"12px"}}>✅</div>
              <div style={{fontWeight:700,color:G.accent}}>Mot de passe mis à jour !</div>
              <div style={{fontSize:"13px",color:G.textDim,marginTop:"8px"}}>Redirection en cours…</div>
            </div>
            :<>
              <div style={{fontSize:"14px",fontWeight:700,color:G.text,marginBottom:"16px"}}>🔒 Nouveau mot de passe</div>
              <Alert msg={err} type="error"/>
              <Input label="Nouveau mot de passe" value={pass} onChange={setPass} type="password" placeholder="Minimum 6 caractères"/>
              <Input label="Confirmer" value={confirm} onChange={setConfirm} type="password" placeholder="Répétez le mot de passe"/>
              <Btn onClick={handle} disabled={loading}>{loading?"Mise à jour...":"Changer le mot de passe"}</Btn>
            </>
          }
        </Card>
      </div>
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
    if(error) { setMsg(error.message); toast.error("Erreur : "+error.message); }
    else{ setModal(null); reload(); toast.success(modal==="new"?"✅ Société créée !":"✅ Société mise à jour !"); }
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
  const initForm={matricule:"",whatsApp:"",nom:"",prenoms:"",emploi:"",categorie:"",
    situation_matrimoniale:"Célibataire",nb_enfants:"0",cnss:"",ifu:"",
    salaire_base:"",date_embauche:"",nationalite:"Béninoise",premier_emploi:false,whatsApp:"",actif:true};
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
    if(error) { setMsg(error.message); toast.error("Erreur : "+error.message); }
    else{ setModal(null); load(); toast.success(modal==="new"?"✅ Employé ajouté !":"✅ Employé mis à jour !"); }
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
// ─── TEMPLATES BASE64 ─────────────────────────────────────────────────────────
const CNSS_TEMPLATE_B64 = "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";
const ITS_TEMPLATE_B64 = "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";

const b64ToBuffer = (b64) => {
  const bin = atob(b64);
  const buf = new Uint8Array(bin.length);
  for (let i=0; i<bin.length; i++) buf[i]=bin.charCodeAt(i);
  return buf.buffer;
};

const getTauxRisque = (taux) => {
  const t = Math.round(parseFloat(taux||0.194)*1000)/10;
  if(t<=16.4) return 1;
  if(t<=17.4) return 2;
  if(t<=18.4) return 3;
  return 4;
};

const genDeclarationCNSS = async (comp, employes, mois, annee) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(b64ToBuffer(CNSS_TEMPLATE_B64));

  const ws  = wb.getWorksheet('declaration');
  const ws2 = wb.getWorksheet('annexe');

  const moisNum  = parseInt(mois);
  const joursFin = [4,6,9,11].includes(moisNum)?30:moisNum===2?28:31;
  const periodeDebut = `01/${String(moisNum).padStart(2,'0')}/${annee}`;
  const periodeFin   = `${joursFin}/${String(moisNum).padStart(2,'0')}/${annee}`;
  const taux       = parseFloat(comp?.taux_cnss_patronale||0.194);
  const tauxRisque = getTauxRisque(taux);
  const societe    = comp?.raison_sociale||'';
  const ifu        = comp?.ifu||'';
  const cnssEmp    = comp?.cnss_employeur||comp?.rccm||'';

  const masse    = employes.reduce((s,e)=>s+parseFloat(e.salaire_brut||0),0);
  const nbEmp    = employes.length;
  const prestFam = Math.round(masse*0.09);
  const risques  = Math.round(masse*tauxRisque/100);
  const partPat  = Math.round(masse*0.064);
  const partOuv  = Math.round(employes.reduce((s,e)=>s+parseFloat(e.cnss_ouvriere||0),0));
  const totAV    = partPat+partOuv;
  const totCot   = prestFam+risques+totAV;

  const numFmt = '#,##0';
  const setV = (cell, val, numF=false) => {
    cell.value = val;
    if(numF && typeof val==='number') cell.numFmt = numFmt;
  };

  // Remplissage IFU chiffre par chiffre dans les cellules C12 à O12
  const ifuStr = String(ifu).replace(/\D/g,'');
  const ifuCols = ['C','D','E','F','G','H','I','J','K','L','M','N','O'];
  ifuCols.forEach((col,idx) => {
    ws.getCell(`${col}12`).value = ifuStr[idx] !== undefined ? ifuStr[idx] : '';
  });

  // N° CNSS Employeur
  ws.getCell('C14').value = cnssEmp;
  // Debut de la période
  ws.getCell('C20').value = periodeDebut;
  // Date limite de dépôt = 10 du mois suivant
  const moisSuivant = moisNum===12?1:moisNum+1;
  const anneeSuivant = moisNum===12?parseInt(annee)+1:parseInt(annee);
  const dateLimite = `10/${String(moisSuivant).padStart(2,'0')}/${anneeSuivant}`;
  ws.getCell('C22').value = dateLimite;

  // Remplissage côté droit (colonne U pour les valeurs, T pour les labels fixes)
  setV(ws.getCell('U10'), `${MOIS[moisNum]} ${annee}`);  // PERIODE DE COTISATION
  setV(ws.getCell('U12'), societe);                        // NOM DU CONTRIBUABLE
  setV(ws.getCell('U14'), 'CNSS');                         // COTISATION
  // U16 = CENTRE DE RECOUVREMENT (manuel)
  setV(ws.getCell('U18'), 'Salaires');                     // OBJET IMPOSABLE
  setV(ws.getCell('U20'), periodeFin);                     // FIN DE LA PERIODE
  setV(ws.getCell('U22'), dateLimite);                     // DATE LIMITE DE PAIEMENT

  // Bas du formulaire (montants)
  setV(ws.getCell('U34'), nbEmp, true);
  setV(ws.getCell('U36'), Math.round(masse), true);
  setV(ws.getCell('U38'), prestFam, true);
  setV(ws.getCell('U40'), risques, true);
  setV(ws.getCell('U42'), partPat, true);
  setV(ws.getCell('U44'), partOuv, true);
  setV(ws.getCell('U46'), totAV, true);
  setV(ws.getCell('U48'), totCot, true);
  setV(ws.getCell('U50'), totCot, true);

  // Signature
  ws.getCell('R53').value = `Cotonou, le ${periodeFin}`;

  // ── FEUILLE ANNEXE ──
  const thin = {style:'thin'};
  const brd  = {top:thin,left:thin,bottom:thin,right:thin};
  const cal  = (bold=false,sz=9) => ({name:'Calibri',bold,size:sz});

  // Sous-titre
  ws2.mergeCells(2,1,2,23);
  const cSub = ws2.getCell(2,1);
  cSub.value = `${societe}  —  Période : ${periodeDebut} au ${periodeFin}  —  N° CNSS : ${cnssEmp}`;
  cSub.font  = {name:'Calibri',size:9,italic:true};
  cSub.alignment = {horizontal:'center',vertical:'middle'};

  employes.forEach((emp,i) => {
    const row  = 5+i;
    const brut = parseFloat(emp.salaire_brut||0);
    const ouv  = parseFloat(emp.cnss_ouvriere||0);
    const pat  = Math.round(brut*0.064);
    const fam  = Math.round(brut*0.09);
    const risp = Math.round(brut*tauxRisque/100);
    const tot  = fam+risp+pat+ouv;
    const vals = [i+1,emp.cnss||'',emp.ifu||'',cnssEmp,tauxRisque,
      emp.nom||'',emp.prenoms||'',emp.date_embauche||'','',26,0,
      Math.round(brut),0,'',0,0,0,Math.round(brut),fam,risp,pat,Math.round(ouv),tot];
    ws2.getRow(row).height = 18;
    vals.forEach((v,j) => {
      const c = ws2.getCell(row,j+1);
      c.value = v; c.font = cal(false,9);
      c.alignment = {horizontal:j>=11?'right':j<=4?'center':'left',vertical:'middle'};
      c.border = brd;
      if(j>=11 && typeof v==='number') c.numFmt = numFmt;
    });
  });

  const tr = 5+employes.length;
  ws2.mergeCells(tr,1,tr,11);
  const cTot = ws2.getCell(tr,1);
  cTot.value = 'TOTAUX'; cTot.font = cal(true,9);
  cTot.alignment = {horizontal:'center',vertical:'middle'}; cTot.border = brd;
  ws2.getRow(tr).height = 18;
  for(let j=12;j<=23;j++){
    const c = ws2.getCell(tr,j);
    c.value = employes.reduce((s,e)=>{
      const brut=parseFloat(e.salaire_brut||0), ouv=parseFloat(e.cnss_ouvriere||0);
      const vals=[0,0,0,0,0,0,0,Math.round(brut),Math.round(brut*0.09),
        Math.round(brut*tauxRisque/100),Math.round(brut*0.064),Math.round(ouv),
        Math.round(brut*0.09)+Math.round(brut*tauxRisque/100)+Math.round(brut*0.064)+Math.round(ouv)];
      return s+(vals[j-11]||0);
    },0);
    c.font=cal(true,9); c.alignment={horizontal:'right',vertical:'middle'};
    c.border=brd; c.numFmt=numFmt;
  }
  ws2.printArea = `A1:W${tr+1}`;

  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=`Declaration_CNSS_${societe.replace(/\s+/g,'_')}_${MOIS[moisNum]}_${annee}.xlsx`;
  a.click(); URL.revokeObjectURL(url);
};

const genDeclarationITS = async (comp, employes, mois, annee) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(b64ToBuffer(ITS_TEMPLATE_B64));

  const ws = wb.getWorksheet('declaration');

  const moisNum  = parseInt(mois);
  const joursFin = [4,6,9,11].includes(moisNum)?30:moisNum===2?28:31;
  const periodeDebut = `01/${String(moisNum).padStart(2,'0')}/${annee}`;
  const periodeFin   = `${joursFin}/${String(moisNum).padStart(2,'0')}/${annee}`;
  const societe  = comp?.raison_sociale||'';
  const ifu      = comp?.ifu||'';
  const cnssEmp  = comp?.cnss_employeur||comp?.rccm||'';

  const nbEmp  = employes.length;
  const masse  = employes.reduce((s,e)=>s+parseFloat(e.salaire_brut||0),0);
  const itsTot = employes.reduce((s,e)=>s+parseFloat(e.its||0),0);

  const numFmt = '#,##0';
  const setV = (cell, val, numF=false) => {
    cell.value = val;
    if(numF && typeof val==='number') cell.numFmt = numFmt;
  };

  // IFU chiffre par chiffre dans les cellules B14 à N14 (ITS)
  const ifuStr = String(ifu).replace(/\D/g,'');
  const ifuCols = ['B','C','D','E','F','G','H','I','J','K','L','M','N'];
  ifuCols.forEach((col,idx) => {
    ws.getCell(`${col}14`).value = ifuStr[idx] !== undefined ? ifuStr[idx] : '';
  });

  // Calcul date limite = 10 du mois suivant
  const moisSuivant = moisNum===12?1:moisNum+1;
  const anneeSuivant = moisNum===12?parseInt(annee)+1:parseInt(annee);
  const dateLimite = `10/${String(moisSuivant).padStart(2,'0')}/${anneeSuivant}`;

  // Côté gauche
  ws.getCell('B16').value = cnssEmp;       // NUMERO COMPTE COTISATION
  ws.getCell('B22').value = periodeDebut;   // DEBUT DE LA PERIODE
  ws.getCell('B24').value = dateLimite;     // DATE LIMITE DE DEPÔT

  // Côté droit (colonnes T/U)
  setV(ws.getCell('T12'), `${MOIS[moisNum]} ${annee}`);   // PERIODE DE COTISATION
  setV(ws.getCell('T14'), societe);                          // NOM DU CONTRIBUABLE
  setV(ws.getCell('T16'), 'IRPP-TS');                        // COTISATION
  // T18 = CENTRE DE RECOUVREMENT (manuel)
  setV(ws.getCell('T20'), 'Salaires & traitements');         // OBJET IMPOSABLE
  setV(ws.getCell('T22'), periodeFin);                       // FIN DE LA PERIODE
  setV(ws.getCell('T24'), dateLimite);                       // DATE LIMITE DE PAIEMENT

  // Montants bas du formulaire (colonne U pour les valeurs)
  setV(ws.getCell('U36'), nbEmp, true);              // Nombre de salariés
  setV(ws.getCell('U38'), Math.round(masse), true);   // Montant brut des salaires
  setV(ws.getCell('U40'), Math.round(itsTot), true);  // IRPP/TS à reverser

  // Signature
  ws.getCell('Q44').value = `Cotonou, le ${dateLimite}`;

  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=`Declaration_ITS_${societe.replace(/\s+/g,'_')}_${MOIS[moisNum]}_${annee}.xlsx`;
  a.click(); URL.revokeObjectURL(url);
};

// ─── GÉNÉRATION DEPUIS MODÈLE ─────────────────────────────────────────────────
const generateFromTemplate = async (template, variables) => {
  const {data:fileData, error}=await supabase.storage
    .from("declaration-templates").download(template.file_path);
  if(error){toast.error("Erreur téléchargement : "+error.message); return;}

  const arrayBuffer=await fileData.arrayBuffer();
  const ft=(template.file_type||"").toLowerCase();

  if(ft==="xlsx"||ft==="xls"){
    const wb=new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    wb.eachSheet(sheet=>{
      sheet.eachRow(row=>{
        row.eachCell(cell=>{
          if(typeof cell.value==="string"){
            let v=cell.value;
            Object.entries(variables).forEach(([k,r])=>{
              v=v.split(k).join(String(r??""));
            });
            if(v!==cell.value) cell.value=v;
          }
          if(cell.value&&cell.value.richText){
            cell.value={richText:cell.value.richText.map(rt=>({
              ...rt,
              text:Object.entries(variables).reduce((t,[k,r])=>t.split(k).join(String(r??"")),rt.text||"")
            }))};
          }
        });
      });
    });
    const buf=await wb.xlsx.writeBuffer();
    const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=`${template.name}_rempli.xlsx`;
    a.click(); URL.revokeObjectURL(url);
    toast.success("✅ Fiche générée et téléchargée !");
  } else {
    const blob=new Blob([arrayBuffer],{type:"application/pdf"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=template.file_name;
    a.click(); URL.revokeObjectURL(url);
    toast.info("📄 Modèle PDF téléchargé — à remplir manuellement.");
  }
};

// ─── LIGNE DE TEMPLATE ────────────────────────────────────────────────────────
const TemplateRow = ({t,onGen,onDel,userId}) => {
  const icons={xlsx:"📊",xls:"📊",pdf:"📄"};
  return (
    <Card style={{marginBottom:"8px",padding:"14px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
          <span style={{fontSize:"24px"}}>{icons[t.file_type]||"📁"}</span>
          <div>
            <div style={{fontWeight:600,fontSize:"14px",color:G.text}}>{t.name}</div>
            <div style={{fontSize:"11px",color:G.textDim,marginTop:"2px",display:"flex",alignItems:"center",gap:"6px"}}>
              {t.file_name} • {(t.file_type||"").toUpperCase()}
              {t.is_global&&<Badge label="Global" color={G.accent}/>}
            </div>
            {t.description&&<div style={{fontSize:"12px",color:G.textDim,marginTop:"3px"}}>{t.description}</div>}
          </div>
        </div>
        <div style={{display:"flex",gap:"8px"}}>
          <Btn small onClick={onGen} variant="primary">⚡ Générer</Btn>
          {t.user_id===userId&&<Btn small onClick={onDel} variant="danger">🗑</Btn>}
        </div>
      </div>
    </Card>
  );
};

// ─── GESTIONNAIRE DE MODÈLES ─────────────────────────────────────────────────
const TemplateManager = ({companies, user, compId, mois, annee, data}) => {
  const [templates,setTemplates]=useState([]);
  const [loading,setLoading]=useState(false);
  const [uploading,setUploading]=useState(false);
  const [showUpload,setShowUpload]=useState(false);
  const [form,setForm]=useState({name:"",description:"",is_global:false,company_id:""});

  const comp=companies.find(c=>c.id===compId);

  const loadTemplates=useCallback(async()=>{
    setLoading(true);
    const {data:rows}=await supabase.from("declaration_templates")
      .select("*").order("created_at",{ascending:false});
    setTemplates(rows||[]);
    setLoading(false);
  },[]);

  useEffect(()=>{loadTemplates();},[loadTemplates]);

  const upload=async(e)=>{
    const file=e.target.files?.[0]; if(!file) return;
    const ext=file.name.split(".").pop().toLowerCase();
    if(!["xlsx","xls","pdf"].includes(ext)){
      toast.error("Format non supporté. Utilisez .xlsx, .xls ou .pdf"); return;
    }
    if(!form.name.trim()){toast.error("Donnez un nom au modèle."); return;}
    setUploading(true);
    const path=`${user.id}/${Date.now()}_${file.name.replace(/\s+/g,"_")}`;
    const {error:upErr}=await supabase.storage.from("declaration-templates").upload(path,file);
    if(upErr){toast.error("Erreur upload : "+upErr.message); setUploading(false); return;}
    const {error:dbErr}=await supabase.from("declaration_templates").insert({
      name:form.name.trim(), description:form.description.trim()||null,
      file_path:path, file_name:file.name, file_type:ext,
      is_global:form.is_global,
      company_id:form.is_global?null:(form.company_id||null),
      user_id:user.id
    });
    if(dbErr){toast.error("Erreur DB : "+dbErr.message);}
    else{toast.success("✅ Modèle uploadé !"); setShowUpload(false); setForm({name:"",description:"",is_global:false,company_id:""}); loadTemplates();}
    setUploading(false);
  };

  const deleteTemplate=async(t)=>{
    if(!window.confirm(`Supprimer "${t.name}" ?`)) return;
    await supabase.storage.from("declaration-templates").remove([t.file_path]);
    await supabase.from("declaration_templates").delete().eq("id",t.id);
    toast.success("Modèle supprimé"); loadTemplates();
  };

  const buildVariables=()=>{
    if(!comp||!data) return {};
    return {
      "{{societe}}":comp.raison_sociale||"",
      "{{rccm}}":comp.rccm||"",
      "{{adresse}}":comp.adresse||"",
      "{{tel}}":comp.tel||"",
      "{{mois}}":MOIS[parseInt(mois)]||"",
      "{{annee}}":annee,
      "{{nb_employes}}":data.nb_fiches||0,
      "{{cnss_salarie}}":Math.round(data.cnss_employe||0),
      "{{cnss_patronale}}":Math.round(data.cnss_employeur||0),
      "{{cnss_total}}":Math.round(data.cnss_total||0),
      "{{vps}}":Math.round(data.vps||0),
      "{{its}}":Math.round(data.its_total||0),
      "{{total_a_verser}}":Math.round((data.cnss_total||0)+(data.vps||0)+(data.its_total||0)),
      "{{date}}":new Date().toLocaleDateString("fr-FR"),
      "{{signataire}}":comp.signataire||"",
    };
  };

  const globalTmpl=templates.filter(t=>t.is_global);
  const myTmpl=templates.filter(t=>!t.is_global);

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px"}}>
        <div style={{fontSize:"13px",color:G.textDim}}>
          Stockez vos modèles Excel/PDF et générez des fiches pré-remplies avec les données du mois sélectionné.
        </div>
        <Btn onClick={()=>setShowUpload(true)} small>📤 Uploader un modèle</Btn>
      </div>

      <Card style={{marginBottom:"16px",padding:"14px",background:"#f0f4ff"}}>
        <div style={{fontSize:"11px",fontWeight:700,color:G.accent,marginBottom:"8px",textTransform:"uppercase",letterSpacing:"0.5px"}}>
          Variables disponibles dans vos modèles Excel
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:"6px"}}>
          {["{{societe}}","{{rccm}}","{{mois}}","{{annee}}","{{nb_employes}}",
            "{{cnss_salarie}}","{{cnss_patronale}}","{{cnss_total}}",
            "{{vps}}","{{its}}","{{total_a_verser}}","{{date}}","{{signataire}}"].map(v=>(
            <code key={v} style={{background:"#fff",border:`1px solid ${G.border}`,borderRadius:"4px",
                                   padding:"2px 7px",fontSize:"11px",color:G.accent,fontFamily:"monospace"}}>{v}</code>
          ))}
        </div>
      </Card>

      {loading&&<div style={{textAlign:"center",color:G.textDim,padding:"24px"}}>Chargement…</div>}

      {globalTmpl.length>0&&(
        <div style={{marginBottom:"16px"}}>
          <div style={{fontSize:"12px",fontWeight:700,color:G.textDim,textTransform:"uppercase",
                       letterSpacing:"0.5px",marginBottom:"8px"}}>🌐 Modèles globaux</div>
          {globalTmpl.map(t=>(
            <TemplateRow key={t.id} t={t} onGen={()=>generateFromTemplate(t,buildVariables())}
                         onDel={()=>deleteTemplate(t)} userId={user?.id}/>
          ))}
        </div>
      )}

      <div style={{marginBottom:"16px"}}>
        <div style={{fontSize:"12px",fontWeight:700,color:G.textDim,textTransform:"uppercase",
                     letterSpacing:"0.5px",marginBottom:"8px"}}>📁 Mes modèles</div>
        {myTmpl.length===0&&!loading&&(
          <Card><div style={{color:G.textDim,textAlign:"center",padding:"24px",fontSize:"13px"}}>
            Aucun modèle — uploadez votre premier modèle Excel ou PDF.
          </div></Card>
        )}
        {myTmpl.map(t=>(
          <TemplateRow key={t.id} t={t} onGen={()=>generateFromTemplate(t,buildVariables())}
                       onDel={()=>deleteTemplate(t)} userId={user?.id}/>
        ))}
      </div>

      {showUpload&&(
        <Modal title="📤 Uploader un modèle" onClose={()=>setShowUpload(false)} width="480px">
          <Input label="Nom du modèle *" value={form.name} onChange={v=>setForm(p=>({...p,name:v}))}
                 placeholder="ex: Déclaration CNSS mensuelle"/>
          <Input label="Description (optionnel)" value={form.description}
                 onChange={v=>setForm(p=>({...p,description:v}))} placeholder="ex: Formulaire officiel CNSS"/>
          <div style={{marginBottom:"14px"}}>
            <label style={{display:"block",fontSize:"12px",color:"#1a3a6b",marginBottom:"5px",fontWeight:600}}>
              Portée du modèle
            </label>
            <div style={{display:"flex",gap:"16px"}}>
              {[{v:false,l:"📁 Mon espace (privé)"},{v:true,l:"🌐 Global (toutes mes sociétés)"}].map(opt=>(
                <label key={String(opt.v)} style={{display:"flex",alignItems:"center",gap:"6px",
                         cursor:"pointer",fontSize:"13px",color:G.text}}>
                  <input type="radio" checked={form.is_global===opt.v}
                         onChange={()=>setForm(p=>({...p,is_global:opt.v}))} style={{cursor:"pointer"}}/>
                  {opt.l}
                </label>
              ))}
            </div>
          </div>
          {!form.is_global&&companies.length>0&&(
            <Select label="Associer à une société (optionnel)"
              value={form.company_id} onChange={v=>setForm(p=>({...p,company_id:v}))}
              options={[{value:"",label:"— Aucune société spécifique —"},...companies.map(c=>({value:c.id,label:c.raison_sociale}))]}/>
          )}
          <div style={{marginBottom:"14px"}}>
            <label style={{display:"block",fontSize:"12px",color:"#1a3a6b",marginBottom:"5px",fontWeight:600}}>
              Fichier (.xlsx, .xls, .pdf) *
            </label>
            <input type="file" accept=".xlsx,.xls,.pdf" onChange={upload} disabled={uploading}
              style={{width:"100%",padding:"8px",border:`1px dashed ${G.border}`,borderRadius:"8px",
                      background:G.bg,fontSize:"13px",cursor:"pointer"}}/>
          </div>
          {uploading&&<div style={{textAlign:"center",color:G.accent,fontSize:"13px",padding:"8px"}}>
            ⏳ Upload en cours…
          </div>}
          <Btn onClick={()=>setShowUpload(false)} variant="ghost" small>Annuler</Btn>
        </Modal>
      )}
    </div>
  );
};

const Declarations = ({companies, user}) => {
  const [compId,setCompId]=useState(companies[0]?.id||"");
  const [mois,setMois]=useState(String(new Date().getMonth()+1));
  const [annee,setAnnee]=useState(String(CURRENT_YEAR));
  const [data,setData]=useState(null);
  const [decl,setDecl]=useState(null);
  const [msg,setMsg]=useState("");
  const [tab,setTab]=useState("suivi");
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

      {/* Onglets */}
      <div style={{display:"flex",gap:"4px",marginBottom:"16px",borderBottom:`2px solid ${G.border}`}}>
        {[{id:"suivi",label:"📋 Suivi déclarations"},{id:"modeles",label:"📁 Modèles de fiches"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)}
            style={{padding:"9px 18px",background:"none",border:"none",
                    borderBottom:tab===t.id?`2px solid ${G.accent}`:"2px solid transparent",
                    marginBottom:"-2px",cursor:"pointer",fontFamily:"inherit",fontSize:"13px",fontWeight:600,
                    color:tab===t.id?G.accent:G.textDim,transition:"all 0.15s"}}>
            {t.label}
          </button>
        ))}
      </div>

      {tab==="modeles"&&(
        <TemplateManager companies={companies} user={user}
          compId={compId} mois={mois} annee={annee} data={data}/>
      )}

      {tab==="suivi"&&data&&(
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
  const [saving,setSaving]=useState(false);

  useEffect(()=>{
    (async()=>{
      const {data}=await supabase.from("email_settings").select("*").eq("user_id",user.id).maybeSingle();
      if(data) setForm({...data,smtp_port:String(data.smtp_port||587)});
    })();
  },[user]);

  const save=async()=>{
    setSaving(true);
    const {error}=await supabase.from("email_settings").upsert({
      ...form,user_id:user.id,smtp_port:parseInt(form.smtp_port||587),
      updated_at:new Date().toISOString()},{onConflict:"user_id"});
    setSaving(false);
    if(error) toast.error("❌ Erreur : "+error.message);
    else toast.success("✅ Paramètres enregistrés avec succès !");
  };

  const f=(k,v)=>setForm(p=>({...p,[k]:v}));

  const ADMIN_PHONE = "2290195808385"; // numéro WhatsApp admin
  const ADMIN_EMAIL_CONTACT = ADMIN_EMAIL;
  const waMsg = encodeURIComponent(`Bonjour, j'ai besoin d'aide sur RH-Paie Pro. Mon compte : ${user?.email}`);

  return (
    <div>
      <h2 style={{marginBottom:"24px",fontSize:"22px",fontWeight:700,color:"#1a3a6b"}}>Paramètres</h2>

      {/* Config WhatsApp expéditeur */}
      <Card style={{maxWidth:"520px",marginBottom:"20px"}}>
        <div style={{fontSize:"14px",fontWeight:600,marginBottom:"16px",color:G.accent}}>
          ⚙️ Configuration WhatsApp — Expéditeur
        </div>
        <div style={{background:"#f0f4ff",borderRadius:"8px",padding:"12px",marginBottom:"16px",
                     fontSize:"12px",color:G.textDim,border:"1px solid #d0dff5"}}>
          💡 Renseignez les coordonnées de l'expéditeur qui apparaîtront dans les messages WhatsApp envoyés aux employés.
        </div>
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
        <Btn onClick={save} disabled={saving}>
          {saving ? "⏳ Enregistrement…" : "💾 Enregistrer"}
        </Btn>
      </Card>

      {/* Contact administrateur */}
      <Card style={{maxWidth:"520px"}}>
        <div style={{fontSize:"14px",fontWeight:600,marginBottom:"4px",color:G.accent}}>
          🆘 Contacter l'Administrateur
        </div>
        <div style={{fontSize:"12px",color:G.textDim,marginBottom:"16px"}}>
          Besoin d'aide, d'une mise à niveau de plan ou d'un accès spécial ?
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"12px"}}>
          {/* WhatsApp */}
          <a href={`https://wa.me/${ADMIN_PHONE}?text=${waMsg}`} target="_blank" rel="noreferrer"
            style={{display:"flex",alignItems:"center",gap:"12px",padding:"14px 16px",
                    background:"#dcfce7",borderRadius:"10px",textDecoration:"none",
                    border:"1px solid #bbf7d0",transition:"opacity 0.15s"}}
            onMouseOver={e=>e.currentTarget.style.opacity="0.85"}
            onMouseOut={e=>e.currentTarget.style.opacity="1"}>
            <span style={{fontSize:"24px"}}>💬</span>
            <div>
              <div style={{fontWeight:700,fontSize:"13px",color:"#15803d"}}>WhatsApp</div>
              <div style={{fontSize:"12px",color:"#166534"}}>+{ADMIN_PHONE} — Réponse rapide</div>
            </div>
            <span style={{marginLeft:"auto",fontSize:"18px",color:"#15803d"}}>→</span>
          </a>
          {/* Email */}
          <a href={`mailto:${ADMIN_EMAIL_CONTACT}?subject=Support RH-Paie Pro&body=Bonjour,%0D%0AMon compte : ${user?.email}%0D%0A%0D%0AMon message :`}
            style={{display:"flex",alignItems:"center",gap:"12px",padding:"14px 16px",
                    background:"#eff6ff",borderRadius:"10px",textDecoration:"none",
                    border:"1px solid #bfdbfe",transition:"opacity 0.15s"}}
            onMouseOver={e=>e.currentTarget.style.opacity="0.85"}
            onMouseOut={e=>e.currentTarget.style.opacity="1"}>
            <span style={{fontSize:"24px"}}>✉️</span>
            <div>
              <div style={{fontWeight:700,fontSize:"13px",color:"#1d4ed8"}}>Email</div>
              <div style={{fontSize:"12px",color:"#1e40af"}}>{ADMIN_EMAIL_CONTACT}</div>
            </div>
            <span style={{marginLeft:"auto",fontSize:"18px",color:"#1d4ed8"}}>→</span>
          </a>
        </div>
        <div style={{marginTop:"16px",padding:"10px 12px",background:"#f8fafc",borderRadius:"8px",
                     fontSize:"11px",color:G.textDim,border:"1px solid #e2e8f0"}}>
          ⏰ Disponible du lundi au vendredi, 8h–18h (heure de Cotonou)
        </div>
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
  const [isRecovery,setIsRecovery]=useState(false);
  const [sub,setSub]=useState(null);
  const [showPricing,setShowPricing]=useState(false);

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{
      setUser(session?.user||null); setLoading(false);
    });
    const {data:{subscription}}=supabase.auth.onAuthStateChange((event,s)=>{
      setUser(s?.user||null);
      if(event==="PASSWORD_RECOVERY") setIsRecovery(true);
    });
    return ()=>subscription.unsubscribe();
  },[]);

  const loadCompanies = useCallback(async()=>{
    if(!user) return;
    const {data}=await supabase.from("companies").select("*").eq("user_id",user.id).order("raison_sociale");
    setCompanies(data||[]);
  },[user]);

  const loadSub = useCallback(async()=>{
    if(!user) return;
    try {
      const {data, error}=await supabase.from("subscriptions").select("*").eq("user_id",user.id).maybeSingle();
      if(error) { setSub({status:"trial",plan:"free",user_id:user.id,trial_ends_at:null}); return; }
      setSub(data||{status:"trial",plan:"free",user_id:user.id,trial_ends_at:null});
    } catch(e) { setSub({status:"trial",plan:"free",user_id:user.id,trial_ends_at:null}); }
  },[user]);

  useEffect(()=>{loadCompanies();},[loadCompanies]);
  useEffect(()=>{loadSub();},[loadSub]);

  const logout=async()=>{ await supabase.auth.signOut(); setUser(null); };

  if(loading) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1a3a6b,#1a6dd6)",display:"flex",alignItems:"center",
                 justifyContent:"center",color:"#ffffff",fontFamily:"system-ui",fontSize:"18px"}}>
      Chargement…
    </div>
  );

  if(!user) return <LoginPage onLogin={()=>loadCompanies()}/>;
  if(isRecovery) return <ResetPasswordPage onDone={async()=>{setIsRecovery(false); await supabase.auth.signOut(); setUser(null);}}/>;

  const pages={
    dashboard:<Dashboard companies={companies}/>,
    companies:<Companies companies={companies} reload={loadCompanies} userId={user.id}/>,
    employees:<Employees companies={companies}/>,
    payroll:<Payroll companies={companies}/>,
    historique:<Historique companies={companies}/>,
    declarations:<Declarations companies={companies} user={user}/>,
    rapport:<RapportCabinet companies={companies}/>,
    settings:<Settings user={user}/>,
    pricing:<PricingPage currentSub={sub} userId={user.id} onClose={()=>setPage("dashboard")}/>,
    admin:<AdminPanel/>,
  };

  const handleUpgrade = () => setPage("pricing");

  return (
    <div style={{minHeight:"100vh",background:"#f0f4ff",color:"#1a2a4a",fontFamily:"system-ui,sans-serif"}}>
      <ToastContainer/>
      <div className="rh-topbar">
        <div style={{fontSize:"16px",fontWeight:800,color:"#fff"}}>RH-Paie Pro</div>
        <button onClick={()=>setSidebarOpen(true)}
          style={{background:"none",border:"none",color:"#fff",fontSize:"24px",cursor:"pointer",
                  padding:"4px 8px",lineHeight:1}}>☰</button>
      </div>
      <Sidebar page={page} setPage={setPage} user={user} onLogout={logout}
               open={sidebarOpen} setOpen={setSidebarOpen} sub={sub} onUpgrade={handleUpgrade}/>
      <div className="rh-main">
        <SubscriptionBanner sub={sub} onUpgrade={handleUpgrade}/>
        {companies.length===0&&page!=="companies"&&page!=="pricing"&&page!=="admin"&&(
          <div className="rh-warning-banner">
            ⚠️ Commencez par créer votre <strong
              style={{cursor:"pointer",textDecoration:"underline"}} onClick={()=>setPage("companies")}>
              première société
            </strong>.
          </div>
        )}
        {pages[page]||pages["dashboard"]}
      </div>
    </div>
  );
}

import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const PIN = "2723";
const PHONE = "08022324523";
const FULL_NAME = "Sholarin Samuel";

const TXS = [
  { id: 1,  date: "Jul 4, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 50000,  balance: 50000,   tag: "deposit" },
  { id: 2,  date: "Jul 31, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 1250,   balance: 51250,   tag: "interest" },
  { id: 3,  date: "Aug 5, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 20000,  balance: 71250,   tag: "deposit" },
  { id: 4,  date: "Aug 31, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 1781,   balance: 73031,   tag: "interest" },
  { id: 5,  date: "Sep 3, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 20000,  balance: 93031,   tag: "deposit" },
  { id: 6,  date: "Sep 30, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 2325,   balance: 95356,   tag: "interest" },
  { id: 7,  date: "Oct 2, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 50000,  balance: 145356,  tag: "deposit" },
  { id: 8,  date: "Oct 31, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 3634,   balance: 148990,  tag: "interest" },
  { id: 9,  date: "Nov 1, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 60000,  balance: 208990,  tag: "deposit" },
  { id: 10, date: "Nov 30, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 5224,   balance: 214214,  tag: "interest" },
  { id: 11, date: "Dec 3, 2025",  type: "credit", label: "Target Savings — Funded",   amount: 50000,  balance: 264214,  tag: "deposit" },
  { id: 12, date: "Dec 31, 2025", type: "credit", label: "Monthly Interest Credit",   amount: 6605,   balance: 270819,  tag: "interest" },
  { id: 13, date: "Jan 18, 2026", type: "debit",  label: "Withdrawal — Target Break", amount: 30000,  balance: 240819,  tag: "withdrawal" },
  { id: 14, date: "Jan 18, 2026", type: "debit",  label: "Interest Forfeited",        amount: 20819,  balance: 220000,  tag: "forfeited" },
  { id: 15, date: "Feb 14, 2026", type: "debit",  label: "Withdrawal — Target Break", amount: 20000,  balance: 200000,  tag: "withdrawal" },
  { id: 16, date: "Feb 14, 2026", type: "debit",  label: "Interest Forfeited",        amount: 5000,   balance: 195000,  tag: "forfeited" },
  { id: 17, date: "Mar 5, 2026",  type: "debit",  label: "Withdrawal — Target Break", amount: 30000,  balance: 165000,  tag: "withdrawal" },
  { id: 18, date: "Mar 5, 2026",  type: "debit",  label: "Interest Forfeited",        amount: 4875,   balance: 160125,  tag: "forfeited" },
  { id: 19, date: "Mar 13, 2026", type: "credit", label: "Monthly Interest Credit",   amount: 22875,  balance: 183000,  tag: "interest" },
];

const fmt = (n: number) => "₦" + Number(n).toLocaleString("en-NG", { minimumFractionDigits: 2 });

// ── SVG Icons ─────────────────────────────────────────────────────────────────
const Icon = ({ name, size = 22, color = "currentColor" }: { name: string, size?: number, color?: string }) => {
  const icons: Record<string, React.JSX.Element> = {
    home: <><path d="M3 9.5L12 3l9 6.5V20a1 1 0 01-1 1H4a1 1 0 01-1-1V9.5z" stroke={color} strokeWidth="1.8" fill="none"/><rect x="9" y="13" width="6" height="8" rx="1" stroke={color} strokeWidth="1.8" fill="none"/></>,
    piggy: <><path d="M19 11c0-4-3.13-7-7-7S5 7 5 11c0 1.85.68 3.54 1.79 4.85L6 19h2.5l.5-1.5h6l.5 1.5H18l-.79-3.15A6.95 6.95 0 0019 11z" stroke={color} strokeWidth="1.8" fill={color === "#fff" ? "#fff" : "none"}/><circle cx="15" cy="10" r="1.5" fill={color === "#fff" ? "#0D60D8" : color}/><path d="M19 10h2c1 0 2 1 2 2v1c0 1-1 2-2 2h-1" stroke={color} strokeWidth="1.8" strokeLinecap="round" fill="none"/><path d="M12 4v-1c0-1 1-1 1-1" stroke={color} strokeWidth="1.5" strokeLinecap="round" fill="none"/></>,
    save: <><path d="M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2z" stroke={color} strokeWidth="1.8" fill="none"/><polyline points="17 21 17 13 7 13 7 21" stroke={color} strokeWidth="1.8" fill="none"/><polyline points="7 3 7 8 15 8" stroke={color} strokeWidth="1.8" fill="none"/></>,
    user: <><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><circle cx="12" cy="7" r="4" stroke={color} strokeWidth="1.8" fill="none"/></>,
    card: <><rect x="2" y="5" width="20" height="14" rx="2" stroke={color} strokeWidth="1.8" fill="none"/><line x1="2" y1="10" x2="22" y2="10" stroke={color} strokeWidth="1.8"/><line x1="6" y1="15" x2="10" y2="15" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    trend: <><polyline points="22 7 13.5 15.5 8.5 10.5 2 17" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/><polyline points="16 7 22 7 22 13" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    pay: <><circle cx="12" cy="12" r="9" stroke={color} strokeWidth="1.8" fill="none"/><path d="M8 12h8M14 9l3 3-3 3" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    more: <><circle cx="5" cy="12" r="1.5" fill={color}/><circle cx="12" cy="12" r="1.5" fill={color}/><circle cx="19" cy="12" r="1.5" fill={color}/></>,
    eye: <><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" stroke={color} strokeWidth="1.8" fill="none"/><circle cx="12" cy="12" r="3" stroke={color} strokeWidth="1.8" fill="none"/></>,
    eyeoff: <><path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19m-6.72-1.07a3 3 0 11-4.24-4.24" stroke={color} strokeWidth="1.8" strokeLinecap="round" fill="none"/><line x1="1" y1="1" x2="23" y2="23" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    bell: <><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0" stroke={color} strokeWidth="1.8" fill="none" strokeLinecap="round"/></>,
    up: <><line x1="12" y1="19" x2="12" y2="5" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><polyline points="5 12 12 5 19 12" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    down: <><line x1="12" y1="5" x2="12" y2="19" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><polyline points="19 12 12 19 5 12" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    back: <><polyline points="15 18 9 12 15 6" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></>,
    target: <><circle cx="12" cy="12" r="10" stroke={color} strokeWidth="1.8" fill="none"/><circle cx="12" cy="12" r="6" stroke={color} strokeWidth="1.8" fill="none"/><circle cx="12" cy="12" r="2" fill={color}/></>,
    lock: <><rect x="3" y="11" width="18" height="11" rx="2" stroke={color} strokeWidth="1.8" fill="none"/><path d="M7 11V7a5 5 0 0110 0v4" stroke={color} strokeWidth="1.8" fill="none" strokeLinecap="round"/></>,
    gift: <><polyline points="20 12 20 22 4 22 4 12" stroke={color} strokeWidth="1.8" fill="none" strokeLinecap="round"/><rect x="2" y="7" width="20" height="5" rx="1" stroke={color} strokeWidth="1.8" fill="none"/><line x1="12" y1="22" x2="12" y2="7" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    info: <><circle cx="12" cy="12" r="10" stroke={color} strokeWidth="1.8" fill="none"/><line x1="12" y1="8" x2="12" y2="12" stroke={color} strokeWidth="2" strokeLinecap="round"/><line x1="12" y1="16" x2="12.01" y2="16" stroke={color} strokeWidth="2" strokeLinecap="round"/></>,
    add: <><circle cx="12" cy="12" r="10" stroke={color} strokeWidth="1.8" fill="none"/><line x1="12" y1="8" x2="12" y2="16" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><line x1="8" y1="12" x2="16" y2="12" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    minus: <><circle cx="12" cy="12" r="10" stroke={color} strokeWidth="1.8" fill="none"/><line x1="8" y1="12" x2="16" y2="12" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    send: <><line x1="22" y1="2" x2="11" y2="13" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><polygon points="22 2 15 22 11 13 2 9 22 2" stroke={color} strokeWidth="1.8" fill="none" strokeLinejoin="round"/></>,
    refer: <><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><circle cx="9" cy="7" r="4" stroke={color} strokeWidth="1.8" fill="none"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    bank: <><line x1="3" y1="22" x2="21" y2="22" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><line x1="6" y1="18" x2="6" y2="11" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><line x1="10" y1="18" x2="10" y2="11" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><line x1="14" y1="18" x2="14" y2="11" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><line x1="18" y1="18" x2="18" y2="11" stroke={color} strokeWidth="1.8" strokeLinecap="round"/><polygon points="12 2 20 7 4 7" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    ussd: <><rect x="5" y="2" width="14" height="20" rx="2" stroke={color} strokeWidth="1.8" fill="none"/><line x1="9" y1="18" x2="15" y2="18" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    check_circle: <><circle cx="12" cy="12" r="10" fill="#0D60D8" stroke="none"/><polyline points="8 12 11 15 16 9" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></>,
    warn_tri: <><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" fill="#fff3cd" stroke="#f59e0b" strokeWidth="1.8"/><line x1="12" y1="9" x2="12" y2="13" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round"/><line x1="12" y1="17" x2="12.01" y2="17" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round"/></>,
    bulb: <><path d="M9 18h6M10 22h4M12 2v1M5.22 5.22l.71.71M18.78 5.22l-.71.71M2 12h1M21 12h1M12 7a5 5 0 00-5 5c0 2.5 2 4.5 4 5h2c2-.5 4-2.5 4-5a5 5 0 00-5-5z" stroke={color} strokeWidth="1.8" fill="none"/></>,
    phone: <><rect x="5" y="2" width="14" height="20" rx="2" stroke={color} strokeWidth="1.8" fill="none"/><line x1="12" y1="18" x2="12.01" y2="18" stroke={color} strokeWidth="2" strokeLinecap="round"/></>,
    tv: <><rect x="2" y="7" width="20" height="13" rx="2" stroke={color} strokeWidth="1.8" fill="none"/><polyline points="7 2 12 7 17 2" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    water: <><path d="M12 22a7 7 0 007-7c0-2-1-3.9-3-5.5s-4-4-4-4-2 2.4-4 4-3 3.5-3 5.5a7 7 0 007 7z" stroke={color} strokeWidth="1.8" fill="none"/></>,
    search: <><circle cx="11" cy="11" r="8" stroke={color} strokeWidth="1.8" fill="none"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke={color} strokeWidth="1.8" strokeLinecap="round"/></>,
    topup: <><path d="M12 2v20M5 9l7-7 7 7" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
    history: <><circle cx="12" cy="12" r="10" stroke={color} strokeWidth="1.8" fill="none"/><polyline points="12 6 12 12 16 14" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">{icons[name]}</svg>;
};

// ── Styles ────────────────────────────────────────────────────────────────────
const S: Record<string, React.CSSProperties> = {
  modalBg: { position: "fixed", top: 0, left: 0, width: "100%", height: "100%", background: "rgba(0,0,0,0.55)", zIndex: 200, display: "flex", alignItems: "flex-end", justifyContent: "center" },
  modal: { background: "#fff", borderRadius: "24px 24px 0 0", width: "100%", maxWidth: 420, padding: "24px 20px 32px", maxHeight: "90vh", overflowY: "auto" },
  btnPrimary: { background: "linear-gradient(135deg,#0D60D8,#1E88E5)", border: "none", borderRadius: 14, color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer", padding: "14px", width: "100%", fontFamily: "inherit" },
  btnGhost: { background: "#fff", border: "2px solid #f0f0f0", borderRadius: 14, color: "#111", fontWeight: 600, fontSize: 14, cursor: "pointer", padding: "14px", width: "100%", fontFamily: "inherit" },
  input: { width: "100%", padding: "13px 14px", borderRadius: 12, border: "1.5px solid #e5e7eb", fontSize: 14, fontFamily: "inherit", background: "#fff", color: "#111", outline: "none" },
  card: { background: "#fff", borderRadius: 20, border: "1px solid #f0f0f0", boxShadow: "0 2px 12px rgba(0,0,0,0.04)" },
  header: { background: "linear-gradient(145deg,#0A4BB3,#0D60D8 55%,#1E88E5)", padding: "48px 18px 24px", position: "relative", overflow: "hidden" },
  headerBlue: { background: "linear-gradient(145deg,#0A4BB3,#0D60D8 55%,#1E88E5)", padding: "48px 18px 24px", position: "relative", overflow: "hidden" },
};

// ── Splash Screen ─────────────────────────────────────────────────────────────
function SplashScreen() {
  return (
    <motion.div 
      initial={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.5 }}
      style={{ 
        position: "fixed", 
        top: 0, 
        left: 0, 
        width: "100%", 
        height: "100%", 
        background: "#0D60D8", 
        zIndex: 1000, 
        display: "flex", 
        flexDirection: "column", 
        alignItems: "center", 
        justifyContent: "center",
        color: "#fff",
        textAlign: "center",
        padding: 20
      }}
    >
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
          <div style={{ width: 100, height: 100, borderRadius: 32, background: "#fff", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 20px 40px rgba(0,0,0,0.2)" }}>
            <Icon name="piggy" size={64} color="#0D60D8" />
          </div>
          <span style={{ fontSize: 48, fontWeight: 900, letterSpacing: -2.5, color: "#fff" }}>piggyvest</span>
        </div>
      </div>
      
      <div style={{ paddingBottom: 60 }}>
        <div style={{ fontSize: 24, fontWeight: 800, marginBottom: 12 }}>Save & invest with ease 🚀</div>
        <div style={{ fontSize: 13, color: "rgba(255,255,255,0.8)", lineHeight: 1.5, maxWidth: 300 }}>
          People have saved over <span style={{ fontWeight: 800, color: "#fff" }}>4 Trillion Naira</span> on Piggyvest in the last <span style={{ fontWeight: 800, color: "#fff" }}>10 years</span> - Let's get started!
        </div>
      </div>
    </motion.div>
  );
}

// ── PIN Screen ────────────────────────────────────────────────────────────────
function PinScreen({ onUnlock }: { onUnlock: () => void }) {
  const [pin, setPin] = useState("");
  const [shake, setShake] = useState(false);
  const [dots, setDots] = useState([false, false, false, false]);

  const press = (k: string) => {
    if (pin.length >= 4) return;
    const next = pin + k;
    setPin(next);
    setDots([false, false, false, false].map((_, i) => i < next.length));
    if (next.length === 4) {
      setTimeout(() => {
        if (next === PIN) onUnlock();
        else {
          setShake(true);
          setTimeout(() => { setShake(false); setPin(""); setDots([false, false, false, false]); }, 600);
        }
      }, 200);
    }
  };

  const del = () => {
    const n = pin.slice(0, -1);
    setPin(n);
    setDots([false, false, false, false].map((_, i) => i < n.length));
  };

  const keys = [1, 2, 3, 4, 5, 6, 7, 8, 9, "", 0, "del"];

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg,#001533,#0A4BB3 50%,#0D60D8)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "0 28px", fontFamily: "system-ui,sans-serif" }}>
      <div style={{ marginBottom: 40, textAlign: "center" }}>
        <div style={{ width: 72, height: 72, borderRadius: 24, background: "#fff", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", boxShadow: "0 12px 32px rgba(0,0,0,0.3)" }}>
          <Icon name="piggy" size={40} color="#0D60D8" />
        </div>
        <div style={{ color: "rgba(255,255,255,0.65)", fontSize: 12, marginBottom: 4 }}>Welcome back</div>
        <div style={{ color: "#fff", fontSize: 18, fontWeight: 700, letterSpacing: 0.5 }}>{PHONE}</div>
      </div>
      <div style={{ display: "flex", gap: 20, marginBottom: 46, animation: shake ? "shake 0.5s ease" : "none" }}>
        {dots.map((f, i) => <div key={i} style={{ width: 15, height: 15, borderRadius: "50%", background: f ? "#fff" : "rgba(255,255,255,0.25)", border: "2px solid rgba(255,255,255,0.5)", transition: "all 0.15s", transform: f ? "scale(1.15)" : "scale(1)" }} />)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "14px 38px", width: 272 }}>
        {keys.map((k, i) => (
          <button key={i} onClick={() => k === "del" ? del() : k !== "" && press(String(k))}
            style={{ height: 68, borderRadius: 18, border: "none", cursor: k === "" ? "default" : "pointer", background: k === "" ? "transparent" : "rgba(255,255,255,0.12)", color: "#fff", fontSize: k === "del" ? 14 : 24, fontFamily: "inherit", outline: "none", transition: "background 0.1s" }}>
            {k === "del" ? "⌫" : k}
          </button>
        ))}
      </div>
      <div style={{ color: "rgba(255,255,255,0.4)", fontSize: 12, marginTop: 36 }}>Forgot PIN? • Use Fingerprint</div>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}20%{transform:translateX(-12px)}40%{transform:translateX(12px)}60%{transform:translateX(-8px)}80%{transform:translateX(8px)}}`}</style>
    </div>
  );
}

// ── Withdraw Modal ────────────────────────────────────────────────────────────
function WithdrawModal({ onClose }: { onClose: () => void }) {
  const [step, setStep] = useState(0);
  const [bank, setBank] = useState("");
  const [account, setAccount] = useState("");
  const [amount, setAmount] = useState("183000");
  const [processing, setProcessing] = useState(false);

  const handleWithdraw = () => {
    setProcessing(true);
    setTimeout(() => {
      setProcessing(false);
      setStep(2);
    }, 2500);
  };

  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} transition={{ type: "spring", damping: 25, stiffness: 300 }}
        style={S.modal} onClick={e => e.stopPropagation()}>
        
        {step === 0 && (
          <>
            <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
              <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Icon name="piggy" size={40} color="#0D60D8" />
              </div>
            </div>
            <div style={{ textAlign: "center", marginBottom: 20 }}>
              <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 10 }}>Target Maturity Reached! 🥳</div>
              <div style={{ fontSize: 13, color: "#6b7280", lineHeight: 1.7 }}>
                Congratulations! Your savings target has matured as of <strong style={{ color: "#111" }}>April 28, 2026</strong>. 
                You can now withdraw your funds to your preferred bank account without any penalties.
              </div>
            </div>
            <div style={{ background: "#f0fdf4", border: "1.5px solid #bbf7d0", borderRadius: 16, padding: "16px", marginBottom: 20 }}>
              <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                <div style={{ fontSize: 24, flexShrink: 0 }}>💰</div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#166534", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 2 }}>Available for Withdrawal</div>
                  <div style={{ fontSize: 22, fontWeight: 900, color: "#15803d" }}>{fmt(parseFloat(amount))}</div>
                </div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={onClose} style={{ ...S.btnGhost, flex: 1 }}>Not Now</button>
              <button onClick={() => setStep(1)} style={{ ...S.btnPrimary, flex: 1 }}>Withdraw Now</button>
            </div>
          </>
        )}

        {step === 1 && (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>Withdrawal Details</div>
              <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={{ fontSize: 12, fontWeight: 700, color: "#6b7280", marginBottom: 8, display: "block" }}>Amount to Withdraw</label>
              <div style={{ position: "relative" }}>
                <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontWeight: 700, color: "#111" }}>₦</span>
                <input type="number" readOnly value={amount} style={{ ...S.input, paddingLeft: 30, background: "#f9fafb" }} />
              </div>
            </div>
            <div style={{ marginBottom: 16 }}>
              <label style={{ fontSize: 12, fontWeight: 700, color: "#6b7280", marginBottom: 8, display: "block" }}>Select Beneficiary Bank</label>
              <select value={bank} onChange={(e) => setBank(e.target.value)} style={S.input}>
                <option value="">Select a bank</option>
                <option value="access">Access Bank</option>
                <option value="gtb">GTBank</option>
                <option value="zenith">Zenith Bank</option>
                <option value="kuda">Kuda MFB</option>
                <option value="opay">OPay</option>
              </select>
            </div>
            <div style={{ marginBottom: 24 }}>
              <label style={{ fontSize: 12, fontWeight: 700, color: "#6b7280", marginBottom: 8, display: "block" }}>Account Number</label>
              <input type="tel" maxLength={10} placeholder="0123456789" value={account} onChange={(e) => setAccount(e.target.value.replace(/\D/g, ""))} style={S.input} />
              {account.length === 10 && <div style={{ fontSize: 11, color: "#15803d", fontWeight: 600, marginTop: 6 }}>✅ Account Name: {FULL_NAME.toUpperCase()}</div>}
            </div>
            <button disabled={!bank || account.length < 10 || processing} onClick={handleWithdraw} style={{ ...S.btnPrimary, opacity: (!bank || account.length < 10) ? 0.6 : 1 }}>
              {processing ? "Initiating Withdrawal..." : "Withdraw Funds"}
            </button>
          </>
        )}

        {step === 2 && (
          <div style={{ textAlign: "center", padding: "10px 0" }}>
            <div style={{ display: "flex", justifyContent: "center", marginBottom: 20 }}>
              <div style={{ width: 72, height: 72, borderRadius: "50%", background: "#fef9c3", display: "flex", alignItems: "center", justifyContent: "center", animation: "pulse 2s infinite" }}>
                <div style={{ fontSize: 32 }}>⏳</div>
              </div>
            </div>
            <div style={{ fontSize: 20, fontWeight: 800, color: "#111", marginBottom: 12 }}>Withdrawal Processing</div>
            <div style={{ fontSize: 14, color: "#6b7280", lineHeight: 1.6, marginBottom: 24 }}>
              Your withdrawal of <strong style={{ color: "#111" }}>{fmt(parseFloat(amount))}</strong> to <strong style={{ color: "#111" }}>{account} ({(bank || "").toUpperCase()})</strong> has been initiated successfully.
              <br /><br />
              Please note that funds will be released in <strong style={{ color: "#0D60D8" }}>less than or up to 24 hours</strong>. You will receive a notification once the transfer is completed.
            </div>
            <button style={S.btnPrimary} onClick={onClose}>I Understand</button>
            <style>{`@keyframes pulse { 0% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.1); opacity: 0.8; } 100% { transform: scale(1); opacity: 1; } }`}</style>
          </div>
        )}
      </motion.div>
    </div>
  );
}

// ── Fund Modal ────────────────────────────────────────────────────────────────
function FundModal({ onClose }: { onClose: () => void }) {
  const [step, setStep] = useState(0);
  const [method, setMethod] = useState<string | null>(null);
  const [amount, setAmount] = useState("");
  const [card, setCard] = useState({ num: "", exp: "", cvv: "", name: "" });
  const [processing, setProcessing] = useState(false);
  const [done, setDone] = useState(false);

  const fmtCard = (v: string) => v.replace(/\D/g, "").slice(0, 16).replace(/(.{4})/g, "$1 ").trim();
  const fmtExp = (v: string) => v.replace(/\D/g, "").slice(0, 4).replace(/(.{2})/, "$1/").slice(0, 5);

  const handlePay = () => {
    setProcessing(true);
    setTimeout(() => {
      setProcessing(false);
      setDone(true);
    }, 2000);
  };

  const methods = [
    { id: "card", icon: <Icon name="card" size={20} color="#0D60D8" />, label: "Debit / Credit Card", sub: "Visa, Mastercard, Verve" },
    { id: "bank", icon: <Icon name="bank" size={20} color="#0D60D8" />, label: "Bank Transfer", sub: "Pay from any bank app instantly" },
    { id: "ussd", icon: <Icon name="ussd" size={20} color="#0D60D8" />, label: "USSD", sub: "Dial *901# or *737# on your phone" },
  ];

  if (done) return (
    <div style={S.modalBg}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} style={{ ...S.modal, textAlign: "center" }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}><Icon name="check_circle" size={56} /></div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 6 }}>Payment Successful! 🎉</div>
        <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 4 }}>{fmt(parseInt(amount || "0"))} has been added to your savings</div>
        <div style={{ fontSize: 12, color: "#9ca3af", marginBottom: 24 }}>Your money is growing at 30% p.a.</div>
        <button style={S.btnPrimary} onClick={onClose}>Done</button>
      </motion.div>
    </div>
  );

  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {step > 0 && <button onClick={() => setStep(s => s - 1)} style={{ background: "none", border: "none", cursor: "pointer", color: "#111", display: "flex" }}><Icon name="back" size={20} /></button>}
          </div>
          <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>
            {step === 0 ? "Fund Target Savings" : step === 1 ? "Enter Amount" : method === "card" ? "Card Details" : "Payment Details"}
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
        </div>

        {/* Step 0: Choose method */}
        {step === 0 && (
          <div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 16 }}>Choose how you'd like to fund your savings</div>
            {methods.map(m => (
              <div key={m.id} onClick={() => { setMethod(m.id); setStep(1); }}
                style={{ display: "flex", alignItems: "center", gap: 14, padding: "14px 16px", background: method === m.id ? "#eff6ff" : "#f9fafb", borderRadius: 16, marginBottom: 10, cursor: "pointer", border: `2px solid ${method === m.id ? "#0D60D8" : "transparent"}`, transition: "all 0.15s" }}>
                <div style={{ width: 44, height: 44, borderRadius: 13, background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{m.icon}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: "#111" }}>{m.label}</div>
                  <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>{m.sub}</div>
                </div>
                <div style={{ fontSize: 20, color: method === m.id ? "#0D60D8" : "#9ca3af" }}>›</div>
              </div>
            ))}
          </div>
        )}

        {/* Step 1: Enter Amount */}
        {step === 1 && (
          <div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 16 }}>How much would you like to save today?</div>
            <div style={{ position: "relative", marginBottom: 24 }}>
              <div style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontWeight: 700, color: "#111" }}>₦</div>
              <input type="number" placeholder="5,000" value={amount} onChange={e => setAmount(e.target.value)}
                style={{ ...S.input, paddingLeft: 32, fontSize: 18, fontWeight: 700 }} />
            </div>
            <div style={{ display: "flex", gap: 8, marginBottom: 24 }}>
              {["5000", "10000", "20000", "50000"].map(v => (
                <button key={v} onClick={() => setAmount(v)} style={{ flex: 1, padding: "8px", borderRadius: 10, border: "1.5px solid #f0f0f0", background: amount === v ? "#eff6ff" : "#fff", color: amount === v ? "#0D60D8" : "#6b7280", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                  {fmt(parseInt(v)).split(".")[0]}
                </button>
              ))}
            </div>
            <button style={S.btnPrimary} onClick={() => setStep(2)} disabled={!amount}>Continue</button>
          </div>
        )}

        {/* Step 2: Payment Details */}
        {step === 2 && (
          <div>
            {method === "card" ? (
              <div>
                <div style={{ background: "linear-gradient(135deg,#111,#333)", borderRadius: 16, padding: 20, color: "#fff", marginBottom: 24, boxShadow: "0 8px 20px rgba(0,0,0,0.15)" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 32 }}>
                    <div style={{ fontSize: 12, opacity: 0.7 }}>Debit Card</div>
                    <div style={{ fontWeight: 800, fontStyle: "italic" }}>VISA</div>
                  </div>
                  <div style={{ fontSize: 18, letterSpacing: 2, marginBottom: 24 }}>{card.num || "**** **** **** ****"}</div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ fontSize: 9, textTransform: "uppercase", opacity: 0.6 }}>Card Holder</div>
                      <div style={{ fontSize: 13, fontWeight: 600 }}>{card.name || "YOUR NAME"}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 9, textTransform: "uppercase", opacity: 0.6 }}>Expires</div>
                      <div style={{ fontSize: 13, fontWeight: 600 }}>{card.exp || "MM/YY"}</div>
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 14, marginBottom: 24 }}>
                  <input placeholder="Card Number" value={card.num} onChange={e => setCard({ ...card, num: fmtCard(e.target.value) })} style={S.input} />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <input placeholder="MM/YY" value={card.exp} onChange={e => setCard({ ...card, exp: fmtExp(e.target.value) })} style={S.input} />
                    <input placeholder="CVV" maxLength={3} value={card.cvv} onChange={e => setCard({ ...card, cvv: e.target.value.replace(/\D/g, "") })} style={S.input} />
                  </div>
                  <input placeholder="Card Holder Name" value={card.name} onChange={e => setCard({ ...card, name: e.target.value.toUpperCase() })} style={S.input} />
                </div>

                <button style={S.btnPrimary} onClick={handlePay}>{processing ? "Processing..." : `Pay ${fmt(parseInt(amount || "0"))}`}</button>
              </div>
            ) : (
              <div style={{ textAlign: "center", padding: "20px 0" }}>
                <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px" }}>
                  <Icon name={method === "bank" ? "bank" : "ussd"} size={32} color="#0D60D8" />
                </div>
                <div style={{ fontSize: 16, fontWeight: 700, color: "#111", marginBottom: 8 }}>
                  {method === "bank" ? "Bank Transfer" : "USSD Payment"}
                </div>
                <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 24, lineHeight: 1.5 }}>
                  {method === "bank" 
                    ? "Transfer to the account number shown in your bank app to complete this transaction." 
                    : "Dial the USSD code on your registered phone number to authorize this payment."}
                </div>
                <button style={S.btnPrimary} onClick={handlePay}>{processing ? "Processing..." : "Confirm Payment"}</button>
              </div>
            )}
          </div>
        )}
      </motion.div>
    </div>
  );
}

// ── Send Modal ────────────────────────────────────────────────────────────────
function SendModal({ onClose }: { onClose: () => void }) {
  const [step, setStep] = useState(0);
  const [amount, setAmount] = useState("");
  const [recipient, setRecipient] = useState("");
  const [done, setDone] = useState(false);

  if (done) return (
    <div style={S.modalBg}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} style={{ ...S.modal, textAlign: "center" }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}><Icon name="check_circle" size={56} /></div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 6 }}>Money Sent! 🚀</div>
        <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 24 }}>You've successfully sent {fmt(parseInt(amount || "0"))} to {recipient}</div>
        <button style={S.btnPrimary} onClick={onClose}>Done</button>
      </motion.div>
    </div>
  );

  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>Send Money</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
        </div>

        {step === 0 && (
          <div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 16 }}>Who are you sending to?</div>
            <input placeholder="Username or Phone Number" value={recipient} onChange={e => setRecipient(e.target.value)} style={{ ...S.input, marginBottom: 20 }} />
            <button style={S.btnPrimary} onClick={() => setStep(1)} disabled={!recipient}>Next</button>
          </div>
        )}

        {step === 1 && (
          <div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 16 }}>How much would you like to send?</div>
            <div style={{ position: "relative", marginBottom: 24 }}>
              <div style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontWeight: 700, color: "#111" }}>₦</div>
              <input type="number" placeholder="0.00" value={amount} onChange={e => setAmount(e.target.value)} style={{ ...S.input, paddingLeft: 32, fontSize: 18, fontWeight: 700 }} />
            </div>
            <button style={S.btnPrimary} onClick={() => setDone(true)} disabled={!amount}>Send {fmt(parseInt(amount || "0"))}</button>
          </div>
        )}
      </motion.div>
    </div>
  );
}

// ── Refer Modal ───────────────────────────────────────────────────────────────
function ReferModal({ onClose }: { onClose: () => void }) {
  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={{ textAlign: "center", padding: "10px 0" }}>
          <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px" }}>
            <Icon name="refer" size={32} color="#0D60D8" />
          </div>
          <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 10 }}>Refer & Earn ₦1,000</div>
          <div style={{ fontSize: 13, color: "#6b7280", lineHeight: 1.6, marginBottom: 24 }}>
            Invite your friends to PiggyVest and you both get <strong style={{ color: "#111" }}>₦1,000</strong> when they fund their account with at least ₦2,000.
          </div>
          <div style={{ background: "#f9fafb", border: "1.5px dashed #e5e7eb", borderRadius: 12, padding: "14px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontWeight: 700, color: "#111", letterSpacing: 1 }}>SHOLARIN13</span>
            <span style={{ fontSize: 12, color: "#0D60D8", fontWeight: 700, cursor: "pointer" }}>COPY</span>
          </div>
          <button style={S.btnPrimary} onClick={onClose}>Invite Friends</button>
        </div>
      </motion.div>
    </div>
  );
}

// ── Main Dashboard ────────────────────────────────────────────────────────────
function Dashboard({ onNavigate }: { onNavigate: (s: string) => void }) {
  const [showBal, setShowBal] = useState(true);
  const [showFund, setShowFund] = useState(false);
  const [showWithdraw, setShowWithdraw] = useState(false);
  const [showSend, setShowSend] = useState(false);
  const [showRefer, setShowRefer] = useState(false);

  return (
    <div style={{ paddingBottom: 100 }}>
      <div style={S.header}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: "50%", background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Icon name="user" color="#fff" size={24} />
            </div>
            <div>
              <div style={{ color: "#fff", fontSize: 18, fontWeight: 800 }}>Hello, {FULL_NAME.split(" ")[0]}</div>
              <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, fontWeight: 600 }}>Good morning ☀️</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ width: 40, height: 40, borderRadius: 12, background: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Icon name="piggy" color="#0D60D8" size={24} />
            </div>
            <div style={{ width: 40, height: 40, borderRadius: 12, background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Icon name="bell" color="#fff" size={20} />
            </div>
          </div>
        </div>

        <div style={{ marginBottom: 24 }}>
          <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>Total Balance</div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ color: "#fff", fontSize: 32, fontWeight: 900, letterSpacing: -0.5 }}>
              {showBal ? fmt(183000) : "₦ ****"}
            </div>
            <button onClick={() => setShowBal(!showBal)} style={{ background: "rgba(255,255,255,0.2)", border: "none", borderRadius: 8, padding: 6, cursor: "pointer", display: "flex" }}>
              <Icon name={showBal ? "eye" : "eyeoff"} size={16} color="#fff" />
            </button>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 24 }}>
          <div style={{ background: "rgba(255,255,255,0.1)", borderRadius: 16, padding: "16px 14px", border: "1px solid rgba(255,255,255,0.1)" }}>
            <div style={{ fontSize: 10, color: "rgba(255,255,255,0.7)", fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>Savings</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{fmt(183000).split(".")[0]}</div>
          </div>
          <div style={{ background: "rgba(255,255,255,0.1)", borderRadius: 16, padding: "16px 14px", border: "1px solid rgba(255,255,255,0.1)" }}>
            <div style={{ fontSize: 10, color: "rgba(255,255,255,0.7)", fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>Interest Earned</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{fmt(22875).split(".")[0]}</div>
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
          {[
            { id: "fund", label: "Add Money", icon: "add", onClick: () => setShowFund(true) },
            { id: "withdraw", label: "Withdraw", icon: "minus", onClick: () => setShowWithdraw(true) },
            { id: "send", label: "Send", icon: "send", onClick: () => setShowSend(true) },
            { id: "refer", label: "Refer", icon: "refer", onClick: () => setShowRefer(true) },
          ].map(btn => (
            <button key={btn.id} onClick={btn.onClick} style={{ flex: 1, background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 8, cursor: "pointer" }}>
              <div style={{ width: 48, height: 48, borderRadius: 16, background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid rgba(255,255,255,0.1)" }}>
                <Icon name={btn.icon} color="#fff" size={20} />
              </div>
              <span style={{ fontSize: 11, fontWeight: 700, color: "#fff" }}>{btn.label}</span>
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: "24px 18px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontSize: 16, fontWeight: 800, color: "#111" }}>My Savings</div>
          <div onClick={() => onNavigate("save")} style={{ fontSize: 13, color: "#0D60D8", fontWeight: 700, cursor: "pointer" }}>See all</div>
        </div>

        <div style={{ ...S.card, padding: 20, marginBottom: 14, cursor: "pointer" }} onClick={() => onNavigate("target")}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
            <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
              <div style={{ width: 44, height: 44, borderRadius: 14, background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Icon name="target" color="#0D60D8" size={24} />
              </div>
              <div>
                <div style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>Target Savings</div>
                <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>Emergency Saving</div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontWeight: 900, fontSize: 16, color: "#111" }}>₦183,000</div>
              <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>of ₦200,000</div>
            </div>
          </div>
          <div style={{ height: 6, background: "#f3f4f6", borderRadius: 3, overflow: "hidden", marginBottom: 12 }}>
            <div style={{ width: "91.5%", height: "100%", background: "#0D60D8", borderRadius: 3 }} />
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#6b7280" }}>91.5% of goal 🎯</div>
            <div style={{ background: "#eff6ff", color: "#0D60D8", padding: "4px 10px", borderRadius: 10, fontSize: 11, fontWeight: 800 }}>30% p.a.</div>
          </div>
        </div>

        <div style={{ ...S.card, padding: 20, cursor: "pointer" }} onClick={() => onNavigate("flex")}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
              <div style={{ width: 44, height: 44, borderRadius: 14, background: "#fffbeb", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Icon name="piggy" color="#d97706" size={24} />
              </div>
              <div>
                <div style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>PiggyFlex</div>
                <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>Flexible Savings</div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontWeight: 900, fontSize: 16, color: "#111" }}>₦0.00</div>
              <div style={{ fontSize: 11, color: "#d97706", fontWeight: 700, marginTop: 2 }}>10% p.a.</div>
            </div>
          </div>
        </div>
      </div>

      <AnimatePresence>
        {showFund && <FundModal onClose={() => setShowFund(false)} />}
        {showWithdraw && <WithdrawModal onClose={() => setShowWithdraw(false)} />}
        {showSend && <SendModal onClose={() => setShowSend(false)} />}
        {showRefer && <ReferModal onClose={() => setShowRefer(false)} />}
      </AnimatePresence>
    </div>
  );
}

// ── Save Screen (All Savings Plans) ───────────────────────────────────────────
function SaveScreen({ onNavigate }: { onNavigate: (s: string) => void }) {
  return (
    <div style={{ paddingBottom: 100 }}>
      <div style={S.header}>
        <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Savings</div>
        <div style={{ color: "#fff", fontSize: 28, fontWeight: 900, marginBottom: 8 }}>All Savings Plans</div>
        <div style={{ color: "rgba(255,255,255,0.8)", fontSize: 14, fontWeight: 600 }}>Total Balance: {fmt(183000)}</div>
      </div>

      <div style={{ padding: "24px 18px", display: "grid", gap: 16 }}>
        {[
          { id: "target", label: "Target Savings", sub: "Emergency Saving", bal: 183000, rate: "30% p.a.", icon: "target", color: "#0D60D8", bg: "#eff6ff" },
          { id: "flex", label: "PiggyFlex", sub: "Flexible • Withdraw anytime", bal: 0, rate: "10% p.a.", icon: "piggy", color: "#d97706", bg: "#fffbeb" },
          { id: "lock", label: "SafeLock", sub: "Fixed • Up to 12 months", bal: 0, rate: "up to 16%", icon: "lock", color: "#7c3aed", bg: "#f5f3ff" },
        ].map(plan => (
          <div key={plan.id} onClick={() => onNavigate(plan.id)} style={{ ...S.card, padding: 20, display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
            <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
              <div style={{ width: 48, height: 48, borderRadius: 16, background: plan.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Icon name={plan.icon} color={plan.color} size={24} />
              </div>
              <div>
                <div style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>{plan.label}</div>
                <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>{plan.sub}</div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontWeight: 900, fontSize: 16, color: "#111" }}>{fmt(plan.bal).split(".")[0]}</div>
              <div style={{ fontSize: 11, color: plan.color, fontWeight: 800, marginTop: 2 }}>{plan.rate}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Invest Screen ─────────────────────────────────────────────────────────────
function InvestScreen({ onSelect }: { onSelect: (p: any) => void }) {
  return (
    <div style={{ paddingBottom: 100 }}>
      <div style={S.headerBlue}>
        <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Investments</div>
        <div style={{ color: "#fff", fontSize: 28, fontWeight: 900, marginBottom: 8 }}>Grow Your Money</div>
        <div style={{ color: "rgba(255,255,255,0.8)", fontSize: 13, fontWeight: 600 }}>SEC regulated • Invest from ₦5,000</div>
      </div>

      <div style={{ padding: "24px 18px" }}>
        <div style={{ ...S.card, padding: 24, marginBottom: 32 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontSize: 14, color: "#6b7280", fontWeight: 600 }}>Portfolio Value</div>
            <div style={{ background: "#eff6ff", color: "#3b82f6", padding: "4px 10px", borderRadius: 10, fontSize: 11, fontWeight: 800 }}>0 active</div>
          </div>
          <div style={{ fontSize: 32, fontWeight: 900, color: "#111", marginBottom: 12 }}>₦0.00</div>
          <div style={{ fontSize: 13, color: "#9ca3af" }}>Start investing to see your portfolio grow here</div>
        </div>

        <div style={{ fontSize: 16, fontWeight: 800, color: "#111", marginBottom: 16 }}>Available Products</div>
        <div style={{ display: "grid", gap: 16 }}>
          {[
            { label: "Treasury Bills", sub: "100% government backed", min: "Min: ₦5,000", rate: "18-21% p.a.", icon: "bank" },
            { label: "Dollar Investment", sub: "USD-denominated returns", min: "Min: $10", rate: "8-12% p.a.", icon: "trend" },
            { label: "Real Estate", sub: "Fractional property ownership", min: "Min: ₦10,000", rate: "15-20% p.a.", icon: "home" },
          ].map(p => (
            <div key={p.label} onClick={() => onSelect(p)} style={{ ...S.card, padding: 20, display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
              <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
                <div style={{ width: 44, height: 44, borderRadius: 14, background: "#f3f4f6", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Icon name={p.icon} color="#111" size={20} />
                </div>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>{p.label}</div>
                  <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>{p.sub}</div>
                  <div style={{ fontSize: 10, color: "#6b7280", marginTop: 4 }}>{p.min}</div>
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#0D60D8" }}>{p.rate}</div>
                <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 2 }}>returns</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Pay Screen ────────────────────────────────────────────────────────────────
function PayScreen({ onSelect, onComingSoon }: { onSelect: (b: any) => void; onComingSoon: (t: string) => void }) {
  const [showTopUp, setShowTopUp] = useState(false);

  return (
    <div style={{ paddingBottom: 100 }}>
      <div style={S.header}>
        <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Payments</div>
        <div style={{ color: "#fff", fontSize: 28, fontWeight: 900, marginBottom: 8 }}>Pay Bills & Send Money</div>
        <div style={{ color: "rgba(255,255,255,0.8)", fontSize: 13, fontWeight: 600 }}>Instant payments • Zero fees on bill payments</div>
      </div>

      <div style={{ padding: "0 18px", marginTop: -24 }}>
        <div style={{ ...S.card, padding: "12px 16px", display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
          <Icon name="search" color="#9ca3af" size={20} />
          <input placeholder="Search for a service..." style={{ border: "none", outline: "none", fontSize: 14, flex: 1, color: "#111" }} />
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 32 }}>
          {[
            { label: "Send", icon: "send", onClick: () => onComingSoon("Send Money") },
            { label: "Pay Card", icon: "card", onClick: () => onComingSoon("Card Payment") },
            { label: "Top Up", icon: "topup", onClick: () => setShowTopUp(true) },
            { label: "History", icon: "history", onClick: () => onComingSoon("Payment History") },
          ].map(item => (
            <div key={item.label} onClick={item.onClick} style={{ ...S.card, padding: "16px 8px", textAlign: "center", cursor: "pointer" }}>
              <div style={{ width: 36, height: 36, borderRadius: 10, background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 10px" }}>
                <Icon name={item.icon} color="#3b82f6" size={18} />
              </div>
              <div style={{ fontWeight: 700, fontSize: 11, color: "#6b7280" }}>{item.label}</div>
            </div>
          ))}
        </div>

        <div style={{ fontSize: 16, fontWeight: 800, color: "#111", marginBottom: 16 }}>Bill Payments</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          {[
            { label: "Electricity", sub: "EKEDC, IKEDC, AEDC", icon: "bulb" },
            { label: "Airtime & Data", sub: "MTN, Airtel, Glo, 9Mobile", icon: "phone" },
            { label: "Cable TV", sub: "DStv, GOtv, Startimes", icon: "tv" },
            { label: "Water", sub: "LSWC, Abuja Water Board", icon: "water" },
          ].map(p => (
            <div key={p.label} onClick={() => onSelect(p)} style={{ ...S.card, padding: 20, cursor: "pointer" }}>
              <div style={{ width: 40, height: 40, borderRadius: 12, background: "#f3f4f6", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 12 }}>
                <Icon name={p.icon} color="#111" size={20} />
              </div>
              <div style={{ fontWeight: 800, fontSize: 14, color: "#111" }}>{p.label}</div>
              <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 4 }}>{p.sub}</div>
            </div>
          ))}
        </div>
      </div>

      <AnimatePresence>
        {showTopUp && <TopUpModal onClose={() => setShowTopUp(false)} />}
      </AnimatePresence>
    </div>
  );
}

// ── Top Up Modal ──────────────────────────────────────────────────────────────
function TopUpModal({ onClose }: { onClose: () => void }) {
  const [amount, setAmount] = useState("");
  const [done, setDone] = useState(false);

  if (done) return (
    <div style={S.modalBg}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} style={{ ...S.modal, textAlign: "center" }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}><Icon name="check_circle" size={56} /></div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 6 }}>Top Up Successful! 📱</div>
        <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 24 }}>You've successfully topped up {fmt(parseInt(amount || "0"))}</div>
        <button style={S.btnPrimary} onClick={onClose}>Done</button>
      </motion.div>
    </div>
  );

  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>Airtime Top Up</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
        </div>
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 8 }}>Phone Number</div>
          <input placeholder="0802 232 4523" style={S.input} />
        </div>
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 8 }}>Amount</div>
          <div style={{ position: "relative" }}>
            <div style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontWeight: 700, color: "#111" }}>₦</div>
            <input type="number" placeholder="500" value={amount} onChange={e => setAmount(e.target.value)} style={{ ...S.input, paddingLeft: 32 }} />
          </div>
        </div>
        <button style={S.btnPrimary} onClick={() => setDone(true)} disabled={!amount}>Top Up Now</button>
      </motion.div>
    </div>
  );
}

// ── Export Modal ──────────────────────────────────────────────────────────────
function ExportModal({ onClose }: { onClose: () => void }) {
  const [format, setFormat] = useState("pdf");
  const [exporting, setExporting] = useState(false);
  const [success, setSuccess] = useState(false);

  const handleExport = () => {
    setExporting(true);
    
    setTimeout(() => {
      try {
        if (format === "excel") {
          const data = TXS.map(tx => ({
            "Date": tx.date,
            "Description": tx.label,
            "Type": tx.type.toUpperCase(),
            "Amount (NGN)": tx.amount,
            "Balance (NGN)": tx.balance
          }));
          const ws = XLSX.utils.json_to_sheet(data);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Transactions");
          XLSX.writeFile(wb, `PiggyVest_Statement_${new Date().toISOString().split('T')[0]}.xlsx`);
        } else {
          const doc = new jsPDF();
          
          // Header
          doc.setFontSize(20);
          doc.setTextColor(13, 96, 216); // #0D60D8
          doc.text("piggyvest", 14, 20);
          
          doc.setFontSize(16);
          doc.setTextColor(0, 0, 0);
          doc.text("Transaction Statement", 14, 30);
          
          doc.setFontSize(10);
          doc.setTextColor(100, 100, 100);
          doc.text(`Account Holder: ${FULL_NAME}`, 14, 40);
          doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 45);
          
          const tableData = TXS.map(tx => [
            tx.date,
            tx.label,
            tx.type.toUpperCase(),
            `N${tx.amount.toLocaleString()}`,
            `N${tx.balance.toLocaleString()}`
          ]);

          autoTable(doc, {
            startY: 55,
            head: [['Date', 'Description', 'Type', 'Amount', 'Balance']],
            body: tableData,
            theme: 'striped',
            headStyles: { fillColor: [13, 96, 216] },
            styles: { fontSize: 9 },
          });
          
          doc.save(`PiggyVest_Statement_${new Date().toISOString().split('T')[0]}.pdf`);
        }
        setSuccess(true);
      } catch (error) {
        console.error("Export error:", error);
      } finally {
        setExporting(false);
      }
    }, 1500);
  };

  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        {success ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px" }}>
              <Icon name="check_circle" color="#0D60D8" size={32} />
            </div>
            <h3 style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 8 }}>Export Successful!</h3>
            <p style={{ fontSize: 14, color: "#6b7280", marginBottom: 24, lineHeight: 1.5 }}>
              Your transaction statement has been generated and downloaded as a <strong>{format.toUpperCase()}</strong> file.
            </p>
            <button style={S.btnPrimary} onClick={onClose}>Done</button>
          </div>
        ) : (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>Export Statement</div>
              <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
            </div>
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 12 }}>Select Date Range</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div>
                  <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>Start Date</div>
                  <input type="date" defaultValue="2025-07-01" style={S.input} />
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>End Date</div>
                  <input type="date" defaultValue="2026-03-13" style={S.input} />
                </div>
              </div>
            </div>
            <div style={{ marginBottom: 24 }}>
              <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 12 }}>Export Format</div>
              <div style={{ display: "flex", gap: 12 }}>
                {["pdf", "excel"].map(f => (
                  <button key={f} onClick={() => setFormat(f)} style={{ flex: 1, padding: "12px", borderRadius: 12, border: `2px solid ${format === f ? "#0D60D8" : "#f0f0f0"}`, background: format === f ? "#eff6ff" : "#fff", color: format === f ? "#0D60D8" : "#6b7280", fontWeight: 700, fontSize: 13, cursor: "pointer", textTransform: "uppercase" }}>
                    {f}
                  </button>
                ))}
              </div>
            </div>
            <button style={S.btnPrimary} onClick={handleExport} disabled={exporting}>
              {exporting ? "Generating..." : "Download Statement"}
            </button>
          </>
        )}
      </motion.div>
    </div>
  );
}

// ── Target Savings Screen ─────────────────────────────────────────────────────
function TargetScreen({ onBack }: { onBack: () => void }) {
  const [showFund, setShowFund] = useState(false);
  const [showWithdraw, setShowWithdraw] = useState(false);
  const [showExport, setShowExport] = useState(false);

  return (
    <div style={{ minHeight: "100vh", background: "#f9fafb", paddingBottom: 40 }}>
      <div style={{ ...S.header, paddingBottom: 40 }}>
        <button onClick={onBack} style={{ background: "none", border: "none", cursor: "pointer", color: "#fff", marginBottom: 20, display: "flex", alignItems: "center", gap: 6 }}>
          <Icon name="back" color="#fff" />
          <span style={{ fontSize: 14, fontWeight: 600 }}>Back to Home</span>
        </button>
        <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 12, textTransform: "uppercase", letterSpacing: 1, fontWeight: 700, marginBottom: 8 }}>Target Savings</div>
        <div style={{ color: "#fff", fontSize: 28, fontWeight: 900, marginBottom: 24 }}>Emergency Saving 🚨</div>

        <div style={{ background: "rgba(255,255,255,0.1)", borderRadius: 16, padding: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", color: "#fff", fontSize: 13, marginBottom: 8 }}>
            <div style={{ fontWeight: 600 }}>Progress: 91.5%</div>
            <div style={{ fontWeight: 800 }}>{fmt(183000)} / {fmt(200000)}</div>
          </div>
          <div style={{ height: 8, background: "rgba(255,255,255,0.2)", borderRadius: 4, overflow: "hidden" }}>
            <div style={{ width: "91.5%", height: "100%", background: "#fff", borderRadius: 4 }} />
          </div>
        </div>
      </div>

      <div style={{ padding: "0 18px", marginTop: -20 }}>
        <div style={{ ...S.card, padding: 20, display: "flex", gap: 12 }}>
          <button onClick={() => setShowFund(true)} style={{ ...S.btnPrimary, flex: 1 }}>Fund Target</button>
          <button onClick={() => setShowWithdraw(true)} style={{ ...S.btnGhost, flex: 1 }}>Withdraw</button>
        </div>

        <div style={{ marginTop: 24 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div style={{ fontSize: 16, fontWeight: 800, color: "#111" }}>Transaction History</div>
            <button onClick={() => setShowExport(true)} style={{ background: "#eff6ff", border: "none", borderRadius: 8, padding: "6px 12px", color: "#0D60D8", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
              Export Statement
            </button>
          </div>
          <div style={{ ...S.card, padding: "0 16px" }}>
            {TXS.slice().reverse().map(tx => (
              <div key={tx.id} style={{ display: "flex", alignItems: "center", gap: 14, padding: "16px 0", borderBottom: "1px solid #f3f4f6" }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: tx.type === "credit" ? "#eff6ff" : "#fef2f2", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Icon name={tx.type === "credit" ? "up" : "down"} size={16} color={tx.type === "credit" ? "#0D60D8" : "#ef4444"} />
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: "#111" }}>{tx.label}</div>
                  <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 2 }}>{tx.date}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontWeight: 800, fontSize: 13, color: tx.type === "credit" ? "#0D60D8" : "#ef4444" }}>
                    {tx.type === "credit" ? "+" : "-"}{fmt(tx.amount).split(".")[0]}
                  </div>
                  <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 2 }}>{fmt(tx.balance).split(".")[0]}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <AnimatePresence>
        {showFund && <FundModal onClose={() => setShowFund(false)} />}
        {showWithdraw && <WithdrawModal onClose={() => setShowWithdraw(false)} />}
        {showExport && <ExportModal onClose={() => setShowExport(false)} />}
      </AnimatePresence>
    </div>
  );
}

// ── Coming Soon Modal ─────────────────────────────────────────────────────────
function ComingSoonModal({ onClose, title }: { onClose: () => void; title: string }) {
  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={{ ...S.modal, textAlign: "center" }} onClick={e => e.stopPropagation()}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🚧</div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#111", marginBottom: 8 }}>{title} Coming Soon</div>
        <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 24, lineHeight: 1.5 }}>
          We're working hard to bring this feature to you. Stay tuned for updates!
        </div>
        <button style={S.btnPrimary} onClick={onClose}>Got it</button>
      </motion.div>
    </div>
  );
}

// ── Linked Banks Modal ────────────────────────────────────────────────────────
function LinkedBanksModal({ onClose }: { onClose: () => void }) {
  return (
    <div style={S.modalBg} onClick={onClose}>
      <motion.div initial={{ y: "100%" }} animate={{ y: 0 }} exit={{ y: "100%" }} style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: "#111" }}>Linked Bank Accounts</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: "#9ca3af", cursor: "pointer" }}>✕</button>
        </div>
        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ ...S.card, padding: 16, display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "#f3f4f6", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Icon name="bank" color="#111" size={20} />
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#111" }}>GT Bank</div>
              <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>0952729029 • Sholarin Samuel</div>
            </div>
            <div style={{ color: "#0D60D8", fontSize: 11, fontWeight: 700 }}>VERIFIED</div>
          </div>
          <div style={{ ...S.card, padding: 16, display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "#f3f4f6", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Icon name="pay" color="#111" size={20} />
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#111" }}>OPay / Play</div>
              <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>8022324523 • Sholarin Samuel</div>
            </div>
            <div style={{ color: "#0D60D8", fontSize: 11, fontWeight: 700 }}>VERIFIED</div>
          </div>
        </div>
        <button style={{ ...S.btnGhost, marginTop: 24, borderColor: "#0D60D8", color: "#0D60D8" }}>+ Add New Bank</button>
      </motion.div>
    </div>
  );
}

// ── App Root ──────────────────────────────────────────────────────────────────
export default function App() {
  const [screen, setScreen] = useState("pin");
  const [nav, setNav] = useState("home");
  const [comingSoon, setComingSoon] = useState<string | null>(null);
  const [showBanks, setShowBanks] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallBtn, setShowInstallBtn] = useState(false);
  const [showSplash, setShowSplash] = useState(true);

  useEffect(() => {
    const timer = setTimeout(() => setShowSplash(false), 2500);
    return () => clearTimeout(timer);
  }, []);

  useEffect(() => {
    const handler = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
      setShowInstallBtn(true);
    };
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  const handleInstall = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === "accepted") {
      setShowInstallBtn(false);
    }
    setDeferredPrompt(null);
  };

  if (screen === "pin") return <PinScreen onUnlock={() => setScreen("main")} />;

  return (
    <div style={{ minHeight: "100vh", background: "#f9fafb", fontFamily: "system-ui,sans-serif", maxWidth: 420, margin: "0 auto", position: "relative" }}>
      <AnimatePresence>
        {showSplash && <SplashScreen />}
      </AnimatePresence>
      
      {nav === "home" && <Dashboard onNavigate={(s) => setNav(s)} />}
      {nav === "save" && <SaveScreen onNavigate={(s) => setNav(s)} />}
      {nav === "target" && <TargetScreen onBack={() => setNav("home")} />}
      {nav === "flex" && (
        <div style={{ padding: 20 }}>
          <button onClick={() => setNav("home")} style={{ background: "none", border: "none", cursor: "pointer", marginBottom: 20 }}><Icon name="back" /></button>
          <h2 style={{ fontWeight: 800, fontSize: 24, marginBottom: 12 }}>PiggyFlex</h2>
          <p style={{ color: "#6b7280", fontSize: 14 }}>Your flexible savings account for emergency funds.</p>
          <div style={{ ...S.card, padding: 24, marginTop: 24, textAlign: "center" }}>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 8 }}>Flex Balance</div>
            <div style={{ fontSize: 32, fontWeight: 900 }}>₦0.00</div>
          </div>
          <button onClick={() => setComingSoon("Withdrawal")} style={{ ...S.btnPrimary, marginTop: 24 }}>Withdraw Funds</button>
        </div>
      )}
      {nav === "lock" && (
        <div style={{ padding: 20 }}>
          <button onClick={() => setNav("home")} style={{ background: "none", border: "none", cursor: "pointer", marginBottom: 20 }}><Icon name="back" /></button>
          <h2 style={{ fontWeight: 800, fontSize: 24, marginBottom: 12 }}>SafeLock</h2>
          <p style={{ color: "#6b7280", fontSize: 14 }}>Lock funds for a fixed period and earn upfront interest.</p>
          <div style={{ ...S.card, padding: 24, marginTop: 24, background: "#f5f3ff", border: "1px solid #ddd6fe" }}>
            <div style={{ fontWeight: 800, color: "#7c3aed" }}>Up to 16% p.a.</div>
            <p style={{ fontSize: 12, color: "#6d28d9", marginTop: 4 }}>Fixed interest paid instantly to your Flex account.</p>
          </div>
          <button onClick={() => setComingSoon("SafeLock")} style={{ ...S.btnPrimary, marginTop: 24, background: "#7c3aed" }}>Create a SafeLock</button>
        </div>
      )}
      {nav === "invest" && <InvestScreen onSelect={(p) => setNav("invest_detail")} />}
      {nav === "invest_detail" && (
        <div style={{ padding: 20, paddingBottom: 100 }}>
          <button onClick={() => setNav("invest")} style={{ background: "none", border: "none", cursor: "pointer", marginBottom: 20 }}><Icon name="back" /></button>
          <div style={{ ...S.card, padding: 24, background: "linear-gradient(135deg,#1e3a8a,#1e40af)", color: "#fff", marginBottom: 24 }}>
            <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 8 }}>Treasury Bills</div>
            <div style={{ fontSize: 24, fontWeight: 900, marginBottom: 16 }}>18-21% p.a.</div>
            <div style={{ fontSize: 13, opacity: 0.9, lineHeight: 1.6 }}>Invest in low-risk government-backed securities and earn guaranteed returns.</div>
          </div>
          <div style={{ display: "grid", gap: 16 }}>
            <div style={{ ...S.card, padding: 20 }}>
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>Minimum Investment</div>
              <div style={{ fontWeight: 700, color: "#111" }}>₦5,000.00</div>
            </div>
            <div style={{ ...S.card, padding: 20 }}>
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>Investment Duration</div>
              <div style={{ fontWeight: 700, color: "#111" }}>30 - 365 Days</div>
            </div>
          </div>
          <button onClick={() => setComingSoon("Investment")} style={{ ...S.btnPrimary, marginTop: 32, background: "#1e40af" }}>Invest Now</button>
        </div>
      )}
      {nav === "pay" && <PayScreen onSelect={(b) => setNav("pay_detail")} onComingSoon={(t) => setComingSoon(t)} />}
      {nav === "pay_detail" && (
        <div style={{ padding: 20, paddingBottom: 100 }}>
          <button onClick={() => setNav("pay")} style={{ background: "none", border: "none", cursor: "pointer", marginBottom: 20 }}><Icon name="back" /></button>
          <h2 style={{ fontWeight: 800, fontSize: 24, marginBottom: 24 }}>Electricity Bill</h2>
          <div style={{ display: "grid", gap: 20 }}>
            <div>
              <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 8 }}>Select Provider</div>
              <select style={S.input}>
                <option>EKEDC (Eko Electricity)</option>
                <option>IKEDC (Ikeja Electricity)</option>
                <option>AEDC (Abuja Electricity)</option>
              </select>
            </div>
            <div>
              <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 8 }}>Meter Number</div>
              <input placeholder="Enter meter number" style={S.input} />
            </div>
            <div>
              <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 8 }}>Amount</div>
              <input type="number" placeholder="₦0.00" style={S.input} />
            </div>
            <button onClick={() => setComingSoon("Bill Payment")} style={S.btnPrimary}>Continue</button>
          </div>
        </div>
      )}
      {nav === "more" && (
        <div style={{ padding: 20, paddingBottom: 100 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 32 }}>
            <div style={{ width: 60, height: 60, borderRadius: 20, background: "#0D60D8", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 24, fontWeight: 800 }}>SS</div>
            <div>
              <div style={{ fontSize: 18, fontWeight: 800 }}>{FULL_NAME}</div>
              <div style={{ fontSize: 13, color: "#6b7280" }}>{PHONE}</div>
            </div>
          </div>
          <div style={{ display: "grid", gap: 1 }}>
            {showInstallBtn && (
              <div onClick={handleInstall} style={{ padding: "18px 0", borderBottom: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <div style={{ width: 32, height: 32, borderRadius: 10, background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Icon name="topup" color="#0D60D8" size={18} />
                  </div>
                  <span style={{ fontWeight: 700, color: "#0D60D8" }}>Install PiggyVest App</span>
                </div>
                <Icon name="back" size={16} color="#9ca3af" />
              </div>
            )}
            {["My Profile", "Linked Bank Accounts", "Security Settings", "Refer & Earn", "Help Center", "Logout"].map(m => (
              <div key={m} onClick={() => {
                if (m === "Logout") setScreen("pin");
                else if (m === "My Profile") setNav("profile");
                else if (m === "Linked Bank Accounts") setShowBanks(true);
                else setComingSoon(m);
              }} style={{ padding: "18px 0", borderBottom: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <span style={{ fontWeight: 600, color: m === "Logout" ? "#ef4444" : "#111" }}>{m}</span>
                <Icon name="back" size={16} color="#9ca3af" />
              </div>
            ))}
          </div>
        </div>
      )}
      {nav === "profile" && (
        <div style={{ padding: 20, paddingBottom: 100 }}>
          <button onClick={() => setNav("more")} style={{ background: "none", border: "none", cursor: "pointer", marginBottom: 20 }}><Icon name="back" /></button>
          <h2 style={{ fontWeight: 800, fontSize: 24, marginBottom: 24 }}>My Profile</h2>
          <div style={{ ...S.card, padding: 20, marginBottom: 20 }}>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>Full Name</div>
            <div style={{ fontWeight: 700, color: "#111", marginBottom: 16 }}>{FULL_NAME}</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>Phone Number</div>
            <div style={{ fontWeight: 700, color: "#111", marginBottom: 16 }}>{PHONE}</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>Email Address</div>
            <div style={{ fontWeight: 700, color: "#111", marginBottom: 16 }}>ayomidesholarin13@gmail.com</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>BVN Status</div>
            <div style={{ fontWeight: 700, color: "#0D60D8" }}>Verified ✅</div>
          </div>
          <button onClick={() => setComingSoon("Profile Editing")} style={S.btnGhost}>Edit Profile</button>
        </div>
      )}

      {/* Bottom Navigation */}
      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 420, background: "#fff", borderTop: "1px solid #f0f0f0", display: "flex", padding: "10px 0 24px", zIndex: 100 }}>
        {[
          { id: "home", label: "Home", icon: "home" },
          { id: "save", label: "Save", icon: "save" },
          { id: "invest", label: "Invest", icon: "trend" },
          { id: "pay", label: "Pay", icon: "pay" },
          { id: "more", label: "More", icon: "more" },
        ].map(item => (
          <button key={item.id} onClick={() => setNav(item.id)}
            style={{ flex: 1, background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 4, cursor: "pointer", color: nav === item.id || (nav === "target" && item.id === "home") ? "#0D60D8" : "#9ca3af" }}>
            <Icon name={item.icon} size={22} color="currentColor" />
            <span style={{ fontSize: 10, fontWeight: 700 }}>{item.label}</span>
          </button>
        ))}
      </div>

      <AnimatePresence>
        {comingSoon && <ComingSoonModal title={comingSoon} onClose={() => setComingSoon(null)} />}
        {showBanks && <LinkedBanksModal onClose={() => setShowBanks(false)} />}
      </AnimatePresence>
    </div>
  );
}

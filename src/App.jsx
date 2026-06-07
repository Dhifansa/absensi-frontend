import { useState, useEffect, useRef, useCallback } from "react";

// ─── Config ─────────────────────────────────────────────────────────────────
const API = "https://absensi-backend-pjxt.vercel.app/api";// Ganti sesuai URL Laravel Anda

// ─── QR Library via CDN (qrcode.react alternative: pure JS) ─────────────────
// Pakai qrcode CDN via script tag, injected once
let qrLib = null;
function loadQR() {
  return new Promise((res) => {
    if (qrLib) return res(qrLib);
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js";
    s.onload = () => { qrLib = window.QRCode; res(window.QRCode); };
    document.head.appendChild(s);
  });
}

// ─── jsQR for scanning ───────────────────────────────────────────────────────
let jsQRLib = null;
function loadJsQR() {
  return new Promise((res) => {
    if (jsQRLib) return res(jsQRLib);
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js";
    s.onload = () => { jsQRLib = window.jsQR; res(window.jsQR); };
    document.head.appendChild(s);
  });
}

// ─── SheetJS for XLSX export ─────────────────────────────────────────────────
let xlsxLib = null;
function loadXLSX() {
  return new Promise((res) => {
    if (xlsxLib) return res(xlsxLib);
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    s.onload = () => { xlsxLib = window.XLSX; res(window.XLSX); };
    document.head.appendChild(s);
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
const api = async (path, opts = {}, token = null) => {
  const headers = { "Content-Type": "application/json", "Accept": "application/json" };
  if (token) headers["Authorization"] = `Bearer ${token}`;
  const r = await fetch(`${API}${path}`, { ...opts, headers });
  const json = await r.json().catch(() => ({}));
  return { ok: r.ok, status: r.status, data: json };
};

// ─── Styles ──────────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background: #0f1117;
    color: #e8eaf0;
    min-height: 100vh;
  }

  .app {
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 24px 16px;
    background: radial-gradient(ellipse at 20% 20%, #1a1f35 0%, #0f1117 60%);
  }

  /* ── Cards ── */
  .card {
    background: #161b2e;
    border: 1px solid #252d45;
    border-radius: 20px;
    padding: 40px;
    width: 100%;
    max-width: 440px;
    box-shadow: 0 24px 64px rgba(0,0,0,.5);
  }

  .card-wide { max-width: 860px; }

  /* ── Logo / Header ── */
  .logo {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 32px;
  }
  .logo-dot {
    width: 36px; height: 36px;
    border-radius: 10px;
    background: linear-gradient(135deg, #4f8ef7, #a855f7);
    display: flex; align-items: center; justify-content: center;
    font-size: 18px;
  }
  .logo-text { font-size: 18px; font-weight: 700; color: #fff; }
  .logo-sub  { font-size: 12px; color: #6b7280; margin-top: 1px; }

  /* ── Form ── */
  .form-title  { font-size: 22px; font-weight: 700; color: #fff; margin-bottom: 6px; }
  .form-sub    { font-size: 14px; color: #6b7280; margin-bottom: 28px; }
  .form-label  { display: block; font-size: 13px; color: #9ca3af; margin-bottom: 6px; font-weight: 500; }
  .form-group  { margin-bottom: 18px; }

  .input, .select {
    width: 100%;
    padding: 12px 16px;
    background: #0f1117;
    border: 1px solid #252d45;
    border-radius: 12px;
    color: #e8eaf0;
    font-size: 15px;
    font-family: inherit;
    outline: none;
    transition: border-color .2s;
  }
  .input:focus, .select:focus { border-color: #4f8ef7; }
  .select option { background: #161b2e; }

  /* ── Buttons ── */
  .btn {
    width: 100%;
    padding: 13px;
    border-radius: 12px;
    font-size: 15px;
    font-weight: 600;
    font-family: inherit;
    cursor: pointer;
    border: none;
    transition: all .15s;
  }
  .btn:active { transform: scale(0.98); }
  .btn-primary {
    background: linear-gradient(135deg, #4f8ef7, #a855f7);
    color: #fff;
  }
  .btn-primary:hover { opacity: 0.9; }
  .btn-secondary {
    background: #1e2640;
    border: 1px solid #252d45;
    color: #c9d1e8;
  }
  .btn-secondary:hover { background: #252d45; }
  .btn-danger {
    background: #3b1219;
    border: 1px solid #7f1d1d;
    color: #fca5a5;
  }
  .btn-danger:hover { background: #4b1a22; }
  .btn-sm {
    width: auto;
    padding: 8px 18px;
    font-size: 13px;
  }
  .btn-success {
    background: linear-gradient(135deg, #059669, #10b981);
    color: #fff;
  }

  /* ── Alerts ── */
  .alert {
    padding: 12px 16px;
    border-radius: 10px;
    font-size: 13px;
    margin-bottom: 18px;
    display: flex;
    align-items: flex-start;
    gap: 10px;
  }
  .alert-error   { background: #3b1219; border: 1px solid #7f1d1d; color: #fca5a5; }
  .alert-success { background: #052e16; border: 1px solid #166534; color: #86efac; }
  .alert-info    { background: #172554; border: 1px solid #1e40af; color: #93c5fd; }

  /* ── Role chooser ── */
  .role-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
    margin-top: 28px;
  }
  .role-card {
    padding: 24px 16px;
    border: 1px solid #252d45;
    border-radius: 16px;
    text-align: center;
    cursor: pointer;
    transition: all .2s;
    background: #0f1117;
  }
  .role-card:hover { border-color: #4f8ef7; background: #101828; }
  .role-card-icon { font-size: 32px; margin-bottom: 10px; }
  .role-card-title { font-size: 15px; font-weight: 600; color: #fff; }
  .role-card-sub   { font-size: 12px; color: #6b7280; margin-top: 4px; }

  /* ── QR Code container ── */
  .qr-wrap {
    background: #fff;
    border-radius: 16px;
    padding: 20px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    margin: 16px auto;
  }
  .qr-outer {
    text-align: center;
    padding: 24px 0;
  }
  .qr-timer {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin-top: 14px;
    font-size: 13px;
    color: #6b7280;
  }
  .qr-timer-bar {
    height: 4px;
    border-radius: 2px;
    background: #252d45;
    overflow: hidden;
    margin-top: 8px;
    width: 200px;
    display: inline-block;
  }
  .qr-timer-fill {
    height: 100%;
    background: linear-gradient(90deg, #4f8ef7, #a855f7);
    transition: width .1s linear;
  }

  /* ── Sesi selector ── */
  .sesi-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin-bottom: 24px;
  }
  .sesi-card {
    padding: 18px;
    border: 1px solid #252d45;
    border-radius: 14px;
    cursor: pointer;
    transition: all .2s;
    text-align: center;
  }
  .sesi-card:hover     { border-color: #4f8ef7; }
  .sesi-card.active    { border-color: #4f8ef7; background: #0d1a3a; }
  .sesi-card-icon  { font-size: 24px; margin-bottom: 8px; }
  .sesi-card-title { font-size: 14px; font-weight: 600; color: #fff; }

  /* ── Camera ── */
  .camera-wrap {
    position: relative;
    border-radius: 16px;
    overflow: hidden;
    background: #0a0c10;
    width: 100%;
    aspect-ratio: 4/3;
    max-height: 300px;
  }
  .camera-wrap video { width: 100%; height: 100%; object-fit: cover; }
  .scanner-overlay {
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    pointer-events: none;
  }
  .scanner-frame {
    width: 160px; height: 160px;
    border: 2px solid #4f8ef7;
    border-radius: 12px;
    box-shadow: 0 0 0 9999px rgba(0,0,0,.45);
  }
  .scanner-line {
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, #4f8ef7, transparent);
    animation: scan 2s ease-in-out infinite;
  }
  @keyframes scan {
    0%  { top: 0; }
    50% { top: calc(100% - 2px); }
    100%{ top: 0; }
  }

  /* ── Table ── */
  .table-wrap {
    overflow-x: auto;
    border-radius: 12px;
    border: 1px solid #252d45;
  }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead tr { background: #1a2035; }
  th {
    padding: 12px 16px;
    text-align: left;
    color: #9ca3af;
    font-weight: 600;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .04em;
    white-space: nowrap;
  }
  td { padding: 11px 16px; border-top: 1px solid #1a2035; color: #c9d1e8; }
  tr:hover td { background: #131929; }

  /* ── Badge ── */
  .badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 600;
  }
  .badge-blue   { background: #172554; color: #93c5fd; }
  .badge-purple { background: #2e1065; color: #d8b4fe; }

  /* ── Dashboard header ── */
  .dash-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 24px;
    flex-wrap: wrap;
    gap: 12px;
  }
  .dash-title { font-size: 20px; font-weight: 700; color: #fff; }
  .dash-actions { display: flex; gap: 10px; flex-wrap: wrap; }

  /* ── Stats ── */
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 14px;
    margin-bottom: 24px;
  }
  .stat-card {
    background: #0f1117;
    border: 1px solid #252d45;
    border-radius: 14px;
    padding: 18px;
  }
  .stat-num   { font-size: 28px; font-weight: 700; color: #fff; }
  .stat-label { font-size: 12px; color: #6b7280; margin-top: 4px; }

  /* ── Filter bar ── */
  .filter-bar {
    display: flex;
    gap: 10px;
    margin-bottom: 18px;
    flex-wrap: wrap;
  }
  .filter-bar .select { width: auto; min-width: 140px; font-size: 13px; padding: 8px 12px; }

  /* ── Success screen ── */
  .success-icon {
    width: 80px; height: 80px;
    border-radius: 50%;
    background: #052e16;
    border: 2px solid #166534;
    display: flex; align-items: center; justify-content: center;
    font-size: 36px;
    margin: 0 auto 20px;
  }
  .success-title { font-size: 22px; font-weight: 700; color: #4ade80; text-align: center; }
  .success-sub   { font-size: 14px; color: #6b7280; text-align: center; margin-top: 6px; }
  .success-detail {
    background: #0f1117;
    border: 1px solid #252d45;
    border-radius: 12px;
    padding: 16px;
    margin: 20px 0;
  }
  .detail-row {
    display: flex; justify-content: space-between;
    padding: 6px 0;
    font-size: 14px;
    border-bottom: 1px solid #1a2035;
  }
  .detail-row:last-child { border-bottom: none; }
  .detail-key { color: #6b7280; }
  .detail-val { color: #e8eaf0; font-weight: 500; }

  /* ── Nav ── */
  .nav {
    display: flex;
    gap: 8px;
    margin-bottom: 24px;
    background: #0f1117;
    border: 1px solid #252d45;
    border-radius: 14px;
    padding: 6px;
  }
  .nav-item {
    flex: 1;
    padding: 10px;
    border-radius: 10px;
    text-align: center;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    color: #6b7280;
    transition: all .2s;
    border: none;
    background: none;
    font-family: inherit;
  }
  .nav-item.active { background: #1e2640; color: #93c5fd; }

  /* ── Divider ── */
  .divider {
    height: 1px;
    background: #252d45;
    margin: 20px 0;
  }

  .text-muted { color: #6b7280; font-size: 13px; }
  .mt-12 { margin-top: 12px; }
  .mt-20 { margin-top: 20px; }
  .gap-8  { display: flex; gap: 8px; }
`;

// ─── QR Display Component ─────────────────────────────────────────────────────
function QRDisplay({ token }) {
  const containerRef = useRef(null);
  const qrRef = useRef(null);

  useEffect(() => {
    if (!token) return;
    loadQR().then((QRCode) => {
      if (containerRef.current) {
        containerRef.current.innerHTML = "";
        qrRef.current = new QRCode(containerRef.current, {
          text: token,
          width: 200,
          height: 200,
          colorDark: "#000000",
          colorLight: "#ffffff",
          correctLevel: QRCode.CorrectLevel.H,
        });
      }
    });
  }, [token]);

  return (
    <div className="qr-wrap">
      <div ref={containerRef} />
    </div>
  );
}

// ─── Camera Scanner Component ────────────────────────────────────────────────
function CameraScanner({ onScan, active }) {
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const streamRef = useRef(null);
  const rafRef = useRef(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState(null);

  const stopCamera = useCallback(() => {
    if (rafRef.current) cancelAnimationFrame(rafRef.current);
    if (streamRef.current) {
      streamRef.current.getTracks().forEach((t) => t.stop());
      streamRef.current = null;
    }
  }, []);

  useEffect(() => {
    if (!active) { stopCamera(); return; }

    loadJsQR().then(() => {
      navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
        .then((stream) => {
          streamRef.current = stream;
          if (videoRef.current) {
            videoRef.current.srcObject = stream;
            videoRef.current.play();
            setReady(true);
          }
        })
        .catch(() => setError("Tidak dapat mengakses kamera. Pastikan izin kamera diberikan."));
    });

    return stopCamera;
  }, [active, stopCamera]);

  useEffect(() => {
    if (!ready || !active) return;

    const tick = () => {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      if (!video || !canvas || video.readyState !== video.HAVE_ENOUGH_DATA) {
        rafRef.current = requestAnimationFrame(tick);
        return;
      }
      const ctx = canvas.getContext("2d");
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const code = jsQRLib(imageData.data, imageData.width, imageData.height, {
        inversionAttempts: "dontInvert",
      });
      if (code) {
        onScan(code.data);
        return;
      }
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [ready, active, onScan]);

  if (error) return <div className="alert alert-error">⚠ {error}</div>;

  return (
    <div className="camera-wrap">
      <video ref={videoRef} playsInline muted style={{ display: active ? "block" : "none" }} />
      <canvas ref={canvasRef} style={{ display: "none" }} />
      {active && (
        <div className="scanner-overlay">
          <div style={{ position: "relative", width: 160, height: 160 }}>
            <div className="scanner-frame" />
            <div className="scanner-line" />
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE: Role Chooser
// ═══════════════════════════════════════════════════════════════════════════════
function PageChooser({ onChoose }) {
  return (
    <div className="card">
      <div className="logo">
        <div className="logo-dot">📋</div>
        <div>
          <div className="logo-text">AbsenQR REM Music</div>
          <div className="logo-sub">Tinggal Scan</div>
        </div>
      </div>
      <div className="form-title">Selamat Datang</div>
      <div className="form-sub">Pilih peran untuk melanjutkan</div>
      <div className="role-grid">
        <div className="role-card" onClick={() => onChoose("peserta")}>
          <div className="role-card-icon">👤</div>
          <div className="role-card-title">Peserta</div>
          <div className="role-card-sub">Login & Scan QR</div>
        </div>
        <div className="role-card" onClick={() => onChoose("admin")}>
          <div className="role-card-icon">🛡️</div>
          <div className="role-card-title">Admin</div>
          <div className="role-card-sub">Kelola Presensi</div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE: Login Peserta
// ═══════════════════════════════════════════════════════════════════════════════
const DIVISI_LIST = [
  "Sekretaris","Bendahara","ITC","PR","Kosinus",
  "Keamanan","Properti","Sponsor","Acara","Ketupel","Waketupel"
];

function PageLoginPeserta({ onBack, onLogin }) {
  const [nic, setNic] = useState("");
  const [divisi, setDivisi] = useState("");
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");

  const submit = async () => {
    if (!nic || !divisi) { setErr("NIC dan Divisi wajib diisi."); return; }
    setLoading(true); setErr("");
    const { ok, data } = await api("/auth/peserta/login", {
      method: "POST",
      body: JSON.stringify({ nic, divisi }),
    });
    setLoading(false);
    if (ok) onLogin(data);
    else setErr(data.message || "Login gagal.");
  };

  return (
    <div className="card">
      <div className="logo">
        <div className="logo-dot">📋</div>
        <div><div className="logo-text">AbsenQR REM Music</div></div>
      </div>
      <div className="form-title">Login Peserta</div>
      <div className="form-sub">Masukkan NIC dan Divisi Anda</div>
      {err && <div className="alert alert-error">⚠ {err}</div>}
      <div className="form-group">
        <label className="form-label">NIC (Nomor Induk)</label>
        <input className="input" placeholder="NIC Kalian" value={nic}
          onChange={e => setNic(e.target.value)} onKeyDown={e => e.key==="Enter" && submit()} />
      </div>
      <div className="form-group">
        <label className="form-label">Divisi</label>
        <select className="select" value={divisi} onChange={e => setDivisi(e.target.value)}>
          <option value="">-- Pilih Divisi --</option>
          {DIVISI_LIST.map(d => <option key={d} value={d}>{d}</option>)}
        </select>
      </div>
      <button className="btn btn-primary" onClick={submit} disabled={loading}>
        {loading ? "Memproses..." : "Masuk"}
      </button>
      <button className="btn btn-secondary mt-12" onClick={onBack}>← Kembali</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE: Dashboard Peserta — Scan QR
// ═══════════════════════════════════════════════════════════════════════════════
function PagePeserta({ user, onLogout }) {
  const [scanning, setScanning] = useState(false);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null); // { success, message, ... }
  const [err, setErr] = useState("");
  const scannedRef = useRef(false);

  const handleScan = useCallback(async (rawToken) => {
    if (scannedRef.current || loading) return;
    scannedRef.current = true;
    setScanning(false);
    setLoading(true); setErr("");

    const { ok, data } = await api("/presensi/scan", {
      method: "POST",
      body: JSON.stringify({ token: rawToken, nic: user.nic, divisi: user.divisi }),
    }, user.token);
    setLoading(false);

    if (ok && data.success) {
      setResult(data);
    } else {
      setErr(data.message || "Scan gagal.");
      scannedRef.current = false;
    }
  }, [user, loading]);

  const reset = () => {
    setResult(null); setErr(""); scannedRef.current = false;
  };

  if (result && result.success) {
    return (
      <div className="card">
        <div className="success-icon">✅</div>
        <div className="success-title">Presensi Berhasil!</div>
        <div className="success-sub">Data Anda telah tercatat</div>
        <div className="success-detail">
          <div className="detail-row"><span className="detail-key">NIC</span><span className="detail-val">{result.nic}</span></div>
          <div className="detail-row"><span className="detail-key">Divisi</span><span className="detail-val">{result.divisi}</span></div>
          <div className="detail-row"><span className="detail-key">Sesi</span><span className="detail-val">{result.jenis_sesi}</span></div>
          <div className="detail-row"><span className="detail-key">Waktu</span><span className="detail-val">{result.waktu}</span></div>
        </div>
        <button className="btn btn-secondary" onClick={reset}>Scan Lagi</button>
        <button className="btn btn-danger mt-12" onClick={onLogout}>Keluar</button>
      </div>
    );
  }

  return (
    <div className="card">
      <div className="dash-header" style={{marginBottom:16}}>
        <div>
          <div style={{fontSize:17,fontWeight:700,color:"#fff"}}>Halo, NIC {user.nic}</div>
          <div className="text-muted">{user.divisi}</div>
        </div>
        <button className="btn btn-danger btn-sm" onClick={onLogout}>Keluar</button>
      </div>
      <div className="divider" />
      {err && <div className="alert alert-error" style={{marginBottom:14}}>⚠ {err}</div>}
      {loading && <div className="alert alert-info">⏳ Memproses presensi...</div>}

      {!scanning && !loading && (
        <>
          <div className="text-muted" style={{textAlign:"center",marginBottom:20,lineHeight:1.6}}>
            Tekan tombol di bawah untuk membuka kamera, lalu arahkan ke QR Code dari Admin.
          </div>
          <button className="btn btn-primary" onClick={() => { scannedRef.current=false; setScanning(true); }}>
            📷 Buka Kamera & Scan QR
          </button>
        </>
      )}

      {scanning && (
        <>
          <div className="alert alert-info" style={{marginBottom:12}}>
            📸 Arahkan kamera ke QR Code dari Admin
          </div>
          <CameraScanner active={scanning} onScan={handleScan} />
          <button className="btn btn-secondary mt-12" onClick={() => setScanning(false)}>
            Batal
          </button>
        </>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE: Login Admin
// ═══════════════════════════════════════════════════════════════════════════════
function PageLoginAdmin({ onBack, onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");

  const submit = async () => {
    if (!username || !password) { setErr("Username dan password wajib diisi."); return; }
    setLoading(true); setErr("");
    const { ok, data } = await api("/auth/admin/login", {
      method: "POST",
      body: JSON.stringify({ username, password }),
    });
    setLoading(false);
    if (ok) onLogin(data);
    else setErr(data.message || "Login gagal.");
  };

  return (
    <div className="card">
      <div className="logo">
        <div className="logo-dot">🛡️</div>
        <div><div className="logo-text">AbsenQR Admin</div></div>
      </div>
      <div className="form-title">Login Admin</div>
      <div className="form-sub">Default: ytta / ytta</div>
      {err && <div className="alert alert-error">⚠ {err}</div>}
      <div className="form-group">
        <label className="form-label">Username</label>
        <input className="input" placeholder="admin" value={username}
          onChange={e => setUsername(e.target.value)} />
      </div>
      <div className="form-group">
        <label className="form-label">Password</label>
        <input className="input" type="password" placeholder="••••••••" value={password}
          onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key==="Enter" && submit()} />
      </div>
      <button className="btn btn-primary" onClick={submit} disabled={loading}>
        {loading ? "Memproses..." : "Masuk sebagai Admin"}
      </button>
      <button className="btn btn-secondary mt-12" onClick={onBack}>← Kembali</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE: Dashboard Admin
// ═══════════════════════════════════════════════════════════════════════════════
function PageAdmin({ admin, onLogout }) {
  const [tab, setTab] = useState("qr"); // "qr" | "presensi"
  const [sesi, setSesi] = useState(null);
  const [token, setToken] = useState(null);
  const [progress, setProgress] = useState(100);
  const [generating, setGenerating] = useState(false);
  const [qrActive, setQrActive] = useState(false);

  const [presensi, setPresensi] = useState([]);
  const [filterSesi, setFilterSesi] = useState("");
  const [filterDivisi, setFilterDivisi] = useState("");
  const [loadingPresensi, setLoadingPresensi] = useState(false);

  // ── Generate QR & auto-rotate every 5s ──────────────────────────────────────
  const generateToken = useCallback(async (jenis) => {
    const { ok, data } = await api("/qr/generate", {
      method: "POST",
      body: JSON.stringify({ jenis_sesi: jenis }),
    }, admin.token);
    if (ok) setToken(data.token);
  }, [admin.token]);

  useEffect(() => {
    if (!qrActive || !sesi) return;
    generateToken(sesi);
    setProgress(100);

    const INTERVAL = 5000;
    const TICK = 50;
    let elapsed = 0;

    const rotateTimer = setInterval(() => {
      generateToken(sesi);
      elapsed = 0;
      setProgress(100);
    }, INTERVAL);

    const progressTimer = setInterval(() => {
      elapsed += TICK;
      setProgress(Math.max(0, 100 - (elapsed / INTERVAL) * 100));
    }, TICK);

    return () => { clearInterval(rotateTimer); clearInterval(progressTimer); };
  }, [qrActive, sesi, generateToken]);

  const startQR = async (jenis) => {
    setSesi(jenis);
    setGenerating(true);
    await generateToken(jenis);
    setGenerating(false);
    setQrActive(true);
  };

  const stopQR = () => { setQrActive(false); setToken(null); setSesi(null); };

  // ── Presensi ─────────────────────────────────────────────────────────────────
  const fetchPresensi = useCallback(async () => {
    setLoadingPresensi(true);
    const params = new URLSearchParams();
    if (filterSesi)  params.set("jenis_sesi", filterSesi);
    if (filterDivisi) params.set("divisi", filterDivisi);
    const { ok, data } = await api(`/presensi?${params}`, {}, admin.token);
    setLoadingPresensi(false);
    if (ok) setPresensi(data);
  }, [admin.token, filterSesi, filterDivisi]);

  useEffect(() => { if (tab === "presensi") fetchPresensi(); }, [tab, fetchPresensi]);

  // ── Export CSV ────────────────────────────────────────────────────────────────
  const exportCSV = () => {
    const params = new URLSearchParams();
    if (filterSesi) params.set("jenis_sesi", filterSesi);
    window.open(`${API}/presensi/export?${params}&token=${admin.token}`, "_blank");
  };

  // ── Export XLSX (client-side) ─────────────────────────────────────────────────
  const exportXLSX = async () => {
    const XLSX = await loadXLSX();
    const rows = presensi.map(p => ({
      NIC: p.nic,
      Divisi: p.divisi,
      "Jenis Sesi": p.jenis_sesi,
      "Waktu Presensi": p.waktu,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 12 }, { wch: 22 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Presensi");
    XLSX.writeFile(wb, `presensi_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const totalHariH   = presensi.filter(p => p.jenis_sesi === "Hari H").length;
  const totalGladi   = presensi.filter(p => p.jenis_sesi === "Gladi").length;
  const allDivisi    = [...new Set(presensi.map(p => p.divisi))].sort();

  return (
    <div className="card card-wide">
      {/* ── Header ── */}
      <div className="dash-header">
        <div>
          <div className="dash-title">Dashboard Admin</div>
          <div className="text-muted">Sistem Presensi QR Code</div>
        </div>
        <button className="btn btn-danger btn-sm" onClick={onLogout}>Keluar</button>
      </div>

      {/* ── Tabs ── */}
      <div className="nav">
        <button className={`nav-item${tab==="qr"?" active":""}`} onClick={() => setTab("qr")}>
          📱 Generate QR
        </button>
        <button className={`nav-item${tab==="presensi"?" active":""}`} onClick={() => setTab("presensi")}>
          📊 Data Presensi
        </button>
      </div>

      {/* ════════════════════ TAB: QR ════════════════════ */}
      {tab === "qr" && (
        <>
          {!qrActive && (
            <>
              <div className="form-sub" style={{marginBottom:16}}>Pilih jenis sesi untuk generate QR Code</div>
              <div className="sesi-grid">
                <div className={`sesi-card${sesi==="Hari H"?" active":""}`} onClick={() => !qrActive && setSesi("Hari H")}>
                  <div className="sesi-card-icon">🎯</div>
                  <div className="sesi-card-title">Hari H</div>
                </div>
                <div className={`sesi-card${sesi==="Gladi"?" active":""}`} onClick={() => !qrActive && setSesi("Gladi")}>
                  <div className="sesi-card-icon">🎭</div>
                  <div className="sesi-card-title">Gladi</div>
                </div>
              </div>
              <button
                className="btn btn-primary"
                disabled={!sesi || generating}
                onClick={() => startQR(sesi)}
              >
                {generating ? "Membuat QR..." : "▶ Mulai Generate QR"}
              </button>
            </>
          )}

          {qrActive && token && (
            <div className="qr-outer">
              <div className="alert alert-info" style={{marginBottom:16,justifyContent:"center"}}>
                🔄 QR untuk sesi <strong style={{margin:"0 4px"}}>{sesi}</strong> — berputar otomatis tiap 5 detik
              </div>
              <QRDisplay token={token} />
              <div className="qr-timer">
                <span>Refresh dalam</span>
                <strong style={{color:"#93c5fd"}}>{Math.ceil(progress / 20)}s</strong>
              </div>
              <div className="qr-timer-bar">
                <div className="qr-timer-fill" style={{width:`${progress}%`}} />
              </div>
              <div className="mt-20">
                <button className="btn btn-danger btn-sm" style={{width:"100%"}} onClick={stopQR}>
                  ⏹ Hentikan QR
                </button>
              </div>
            </div>
          )}
        </>
      )}

      {/* ════════════════════ TAB: PRESENSI ════════════════════ */}
      {tab === "presensi" && (
        <>
          <div className="stats-grid">
            <div className="stat-card">
              <div className="stat-num">{presensi.length}</div>
              <div className="stat-label">Total Presensi</div>
            </div>
            <div className="stat-card">
              <div className="stat-num">{totalHariH}</div>
              <div className="stat-label">Hari H</div>
            </div>
            <div className="stat-card">
              <div className="stat-num">{totalGladi}</div>
              <div className="stat-label">Gladi</div>
            </div>
            <div className="stat-card">
              <div className="stat-num">{allDivisi.length}</div>
              <div className="stat-label">Divisi Hadir</div>
            </div>
          </div>

          <div className="filter-bar">
            <select className="select" value={filterSesi} onChange={e => setFilterSesi(e.target.value)}>
              <option value="">Semua Sesi</option>
              <option>Hari H</option>
              <option>Gladi</option>
            </select>
            <select className="select" value={filterDivisi} onChange={e => setFilterDivisi(e.target.value)}>
              <option value="">Semua Divisi</option>
              {DIVISI_LIST.map(d => <option key={d}>{d}</option>)}
            </select>
            <button className="btn btn-secondary btn-sm" onClick={fetchPresensi}>
              🔄 Refresh
            </button>
            <div style={{marginLeft:"auto"}} className="gap-8">
              <button className="btn btn-secondary btn-sm" onClick={exportCSV}>
                ↓ CSV
              </button>
              <button className="btn btn-success btn-sm" onClick={exportXLSX}>
                ↓ XLSX
              </button>
            </div>
          </div>

          {loadingPresensi ? (
            <div className="alert alert-info">⏳ Memuat data presensi...</div>
          ) : presensi.length === 0 ? (
            <div className="alert alert-info">📭 Belum ada data presensi.</div>
          ) : (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th>NIC</th>
                    <th>Divisi</th>
                    <th>Jenis Sesi</th>
                    <th>Waktu Presensi</th>
                  </tr>
                </thead>
                <tbody>
                  {presensi.map((p, i) => (
                    <tr key={p.id}>
                      <td style={{color:"#4b5563"}}>{i + 1}</td>
                      <td><strong style={{color:"#c9d1e8"}}>{p.nic}</strong></td>
                      <td>{p.divisi}</td>
                      <td>
                        <span className={`badge ${p.jenis_sesi==="Hari H" ? "badge-blue" : "badge-purple"}`}>
                          {p.jenis_sesi}
                        </span>
                      </td>
                      <td style={{fontVariantNumeric:"tabular-nums"}}>
                        {new Date(p.waktu).toLocaleString("id-ID", {
                          day:"2-digit", month:"short", year:"numeric",
                          hour:"2-digit", minute:"2-digit", second:"2-digit"
                        })}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ROOT
// ═══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [page, setPage] = useState("choose"); // choose | loginPeserta | loginAdmin | peserta | admin
  const [user, setUser]   = useState(null);
  const [admin, setAdmin] = useState(null);

  const handlePesertaLogin = (data) => { setUser(data);  setPage("peserta"); };
  const handleAdminLogin   = (data) => { setAdmin(data); setPage("admin"); };
  const logout = () => { setUser(null); setAdmin(null); setPage("choose"); };

  return (
    <>
      <style>{css}</style>
      <div className="app">
        {page === "choose"       && <PageChooser onChoose={r => setPage(r === "peserta" ? "loginPeserta" : "loginAdmin")} />}
        {page === "loginPeserta" && <PageLoginPeserta onBack={() => setPage("choose")} onLogin={handlePesertaLogin} />}
        {page === "loginAdmin"   && <PageLoginAdmin   onBack={() => setPage("choose")} onLogin={handleAdminLogin} />}
        {page === "peserta"      && user  && <PagePeserta user={user}   onLogout={logout} />}
        {page === "admin"        && admin && <PageAdmin   admin={admin} onLogout={logout} />}
      </div>
    </>
  );
}

import { useState, useRef, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Responsive hook ──────────────────────────────────────────────────────────
function useBreakpoint() {
  const [w, setW] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(() => {
    const h = () => setW(window.innerWidth);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return { w, mobile: w < 640, tablet: w >= 640 && w < 1024, desktop: w >= 1024 };
}

// ─── Electrical constants ─────────────────────────────────────────────────────
const VOLTAGE_PRESETS = [
  { label: "480/277V", factor: 0.831 },
  { label: "120/208V", factor: 0.360 },
];
const DEFAULT_CATEGORIES = [
  { name: "Retail", wattsPerSqft: 30, isHousePanel: false },
  { name: "Restaurant", wattsPerSqft: 50, isHousePanel: false },
  { name: "Office", wattsPerSqft: 15, isHousePanel: false },
  { name: "House Panel", wattsPerSqft: 0, isHousePanel: true },
];
const SERVICE_SIZES = [100, 200, 400, 600, 800, 1000, 1200, 2000];
const SELECTABLE_SIZES = [100, 200, 400, 600, 800];
const getServiceSize = (a) => SERVICE_SIZES.find((s) => a <= s) || 2000;
const calcKVA = (sqft, w) => (Number(sqft) * Number(w)) / 1000;
const calcAmps = (kva, f) => (kva > 0 ? kva / f : 0);
const DEFAULT_TAP_CAN_IN = 30;
const POST_TAP_GAP = 6; // 6" gap between tap can and first disconnect

const DEFAULT_WIREWAY_RULES = [
  { id: "r1", label: "≤ 200A", maxAmps: 200, type: "Stacked", disconnectIn: 19, meterIn: 0, divIn: 0, gapIn: 6 },
  { id: "r2", label: "400–600A", maxAmps: 600, type: "Side by Side", disconnectIn: 30, meterIn: 30, divIn: 6, gapIn: 6 },
  { id: "r3", label: "≥ 800A", maxAmps: 99999, type: "Side by Side Large", disconnectIn: 36, meterIn: 36, divIn: 6, gapIn: 6 },
];
const WIREWAY_TYPES = ["Stacked", "Side by Side", "Side by Side Large"];

function calcSlotTotal(r) {
  return r.type === "Stacked"
    ? r.disconnectIn + r.gapIn
    : r.disconnectIn + r.divIn + r.meterIn + r.gapIn;
}

function getSlot(serviceSize, rules) {
  for (const r of rules) {
    if (serviceSize <= r.maxAmps) {
      return { type: r.type, disconnectIn: r.disconnectIn, meterIn: r.meterIn, divIn: r.divIn, gapIn: r.gapIn, totalIn: calcSlotTotal(r) };
    }
  }
  const last = rules[rules.length - 1];
  return { type: last.type, disconnectIn: last.disconnectIn, meterIn: last.meterIn, divIn: last.divIn, gapIn: last.gapIn, totalIn: calcSlotTotal(last) };
}

let _id = 1;
const uid = () => _id++ + "-" + Math.random().toString(36).slice(2);
const blankTenant = (b = 1, cat = "Retail") => ({
  id: uid(), building: b, name: "", category: cat, sqft: "", fixedKVA: "", serviceOverride: "",
});
const BLANK_PROJECT = { name: "", number: "", buildings: 1 };
const BLANK_WIREWAY = { availableWallFeet: "" };

// ─── Color tokens ─────────────────────────────────────────────────────────────
const C = {
  bg: "#0e1c30",
  surface: "#162840",
  raised: "#1c344f",
  border: "#2a4e72",
  borderB: "#f59e0b",
  text: "#ddeeff",
  muted: "#7fa8c9",
  amber: "#f59e0b",
  blue: "#38bdf8",
  green: "#4ade80",
  purple: "#a78bfa",
  red: "#f87171",
  lime: "#a3e635",
};

// ─── Design system ────────────────────────────────────────────────────────────
const S = {
  app: {
    minHeight: "100vh", background: C.bg, color: C.text,
    fontFamily: "'Courier New',monospace", fontSize: 13,
  },
  hdr: {
    background: C.surface, borderBottom: `2px solid ${C.amber}`,
    padding: "12px 24px", display: "flex", alignItems: "center",
    gap: 14, flexWrap: "wrap", position: "sticky", top: 0, zIndex: 50,
  },
  logo: { color: C.amber, fontSize: 18, fontWeight: "bold", letterSpacing: 4 },
  sub: { fontSize: 9, color: C.muted, letterSpacing: 3, textTransform: "uppercase", marginTop: 2 },
  nav: { display: "flex", gap: 6, marginLeft: "auto", flexWrap: "wrap", alignItems: "center" },
  main: { maxWidth: 1120, margin: "0 auto", padding: "20px 16px" },

  card: {
    background: C.surface, border: `1px solid ${C.border}`,
    padding: 18, marginBottom: 16, borderRadius: 2,
  },
  cardH: {
    color: C.amber, fontSize: 9, letterSpacing: 3, textTransform: "uppercase",
    marginBottom: 14, borderBottom: `1px solid ${C.border}`, paddingBottom: 8,
    display: "flex", justifyContent: "space-between", alignItems: "center",
  },

  lbl: {
    fontSize: 9, color: C.muted, letterSpacing: 2, textTransform: "uppercase",
    display: "block", marginBottom: 4,
  },
  inp: {
    background: C.raised, border: `1px solid ${C.border}`, color: C.text,
    padding: "7px 10px", fontFamily: "monospace", fontSize: 12,
    width: "100%", outline: "none", boxSizing: "border-box", borderRadius: 2,
  },

  btn: (bg = C.amber, fg = C.bg) => ({
    background: bg, color: fg, border: "none", padding: "7px 16px",
    cursor: "pointer", fontFamily: "monospace", fontSize: 11,
    fontWeight: "bold", letterSpacing: 1, borderRadius: 2,
    transition: "opacity .15s",
  }),
  navB: (active) => ({
    background: active ? C.amber : "transparent",
    color: active ? C.bg : C.muted,
    border: `1px solid ${active ? C.amber : C.border}`,
    padding: "6px 14px", cursor: "pointer", fontFamily: "monospace",
    fontSize: 11, letterSpacing: 1, borderRadius: 2, transition: "all .15s",
  }),
  ghost: {
    background: "transparent", color: C.muted, border: `1px solid ${C.border}`,
    padding: "6px 12px", cursor: "pointer", fontFamily: "monospace",
    fontSize: 11, borderRadius: 2,
  },

  tbl: { width: "100%", borderCollapse: "collapse", fontSize: 11 },
  th: {
    background: C.raised, color: C.muted, padding: "8px 10px", textAlign: "left",
    fontSize: 9, letterSpacing: 2, borderBottom: `1px solid ${C.border}`,
    whiteSpace: "nowrap",
  },
  td: { padding: "8px 10px", borderBottom: `1px solid #0a1525`, verticalAlign: "middle" },

  pill: (col = C.green, bg = "#0e2a1a", br = "#166534") => ({
    display: "inline-block", padding: "2px 9px", fontSize: 9, fontWeight: "bold",
    letterSpacing: 1, background: bg, color: col, border: `1px solid ${br}`, borderRadius: 2,
  }),

  stat: {
    background: C.bg, border: `1px solid ${C.border}`, padding: "12px 16px",
    textAlign: "center", borderRadius: 2,
  },
  sVal: (col = C.amber) => ({ fontSize: 20, color: col, fontWeight: "bold", fontFamily: "monospace" }),
  sLbl: { fontSize: 8, color: C.muted, letterSpacing: 2, textTransform: "uppercase", marginTop: 4 },

  g2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  g3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 },
  g4: { display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12 },
  row: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },

  // Responsive grid helpers (called with bp)
  rg2: (m) => ({ display: "grid", gridTemplateColumns: m ? "1fr" : "1fr 1fr", gap: 12 }),
  rg3: (m) => ({ display: "grid", gridTemplateColumns: m ? "1fr" : "1fr 1fr 1fr", gap: 12 }),
  rg4: (m, t) => ({ display: "grid", gridTemplateColumns: m ? "1fr 1fr" : t ? "repeat(2,1fr)" : "repeat(4,1fr)", gap: 12 }),

  catPill: (active) => ({
    display: "inline-block", padding: "3px 10px", fontSize: 10,
    fontWeight: active ? "bold" : "normal", letterSpacing: 1,
    background: active ? C.amber : C.raised, color: active ? C.bg : C.muted,
    border: `1px solid ${active ? C.amber : C.border}`, borderRadius: 2,
    cursor: "pointer", transition: "all .12s", whiteSpace: "nowrap",
  }),

  sel: {
    background: C.raised, border: `1px solid ${C.border}`, color: C.text,
    padding: "7px 10px", fontFamily: "monospace", fontSize: 12,
    width: "100%", outline: "none", borderRadius: 2,
  },
};

// ─── Print styles (injected once) ─────────────────────────────────────────────
const INJECTED_CSS = `
/* ── Responsive ── */
@media (max-width: 639px) {
  [data-r-hdr] { padding: 8px 12px !important; gap: 8px !important; }
  [data-r-hdr] [data-r-logo] { font-size: 14px !important; letter-spacing: 2px !important; }
  [data-r-hdr] [data-r-nav] { gap: 4px !important; margin-left: 0 !important; width: 100%; justify-content: flex-start; }
  [data-r-hdr] [data-r-nav] button { padding: 5px 8px !important; font-size: 9px !important; }
  [data-r-main] { padding: 12px 8px !important; }
  [data-r-card] { padding: 12px !important; }
  [data-r-card] [data-r-cardh] { font-size: 8px !important; flex-wrap: wrap; gap: 6px; }
  [data-r-tblwrap] { overflow-x: auto; -webkit-overflow-scrolling: touch; margin: 0 -12px; padding: 0 12px; }
  [data-r-tblwrap] table { min-width: 560px; }
  [data-r-stat] { padding: 8px 10px !important; }
  [data-r-stat] [data-r-sval] { font-size: 15px !important; }
  [data-r-home-grid] { grid-template-columns: 1fr !important; }
  [data-r-home-card] { padding: 20px !important; }
  [data-r-ww-inline] { flex-direction: column; align-items: flex-start !important; gap: 10px !important; }
}
@media (min-width: 640px) and (max-width: 1023px) {
  [data-r-hdr] { padding: 10px 16px !important; }
  [data-r-hdr] [data-r-nav] button { padding: 5px 10px !important; font-size: 10px !important; }
  [data-r-tblwrap] { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  [data-r-tblwrap] table { min-width: 500px; }
}

/* ── Print ── */
@media print {
  body, html { background: #fff !important; color: #000 !important; font-size: 10pt !important; }
  [data-print-hide] { display: none !important; }
  [data-print-only] { display: block !important; }
  * {
    color: #000 !important; background: #fff !important;
    border-color: #ccc !important; box-shadow: none !important;
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }
  table { page-break-inside: avoid; }
  svg text { fill: #333 !important; }
  svg rect { stroke: #999 !important; }
  svg line { stroke: #666 !important; }
}
`;

// ─── Inline confirm dialog ────────────────────────────────────────────────────
function ConfirmInline({ message, onConfirm, onCancel }) {
  return (
    <div style={{
      background: C.raised, border: `1px solid ${C.red}`, borderRadius: 2,
      padding: "10px 14px", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap",
    }}>
      <span style={{ color: C.red, fontSize: 11, flex: 1 }}>⚠ {message}</span>
      <button style={S.btn(C.red, "#fff")} onClick={onConfirm}>Yes, delete</button>
      <button style={S.ghost} onClick={onCancel}>Cancel</button>
    </div>
  );
}

// ─── Wireway Diagram ──────────────────────────────────────────────────────────
function WirewayDiagram({ tenantSlots, totalIn, totalFt, availFt, totalAmps, tapCanIn }) {
  const bp = useBreakpoint();
  const ok = availFt > 0 && totalFt <= availFt;
  const W = bp.mobile ? 400 : bp.tablet ? 560 : 700;
  const PAD = 48;
  const scale = totalIn > 0 ? Math.min((W - PAD * 2) / totalIn, 5) : 1;

  // Build tenant segments
  const tenantSegs = [];
  let tx = 0;
  const pushT = (seg) => { tenantSegs.push({ ...seg, rx: tx }); tx += seg.w; };

  // Initial 6″ gap after tap can
  pushT({ w: POST_TAP_GAP * scale, kind: "gap", col: "transparent", label: `${POST_TAP_GAP}"` });

  tenantSlots.forEach((ts) => {
    const sl = ts.slot;
    if (sl.type === "Stacked") {
      pushT({ w: sl.disconnectIn * scale, kind: "equip", label: ts.name || ts.category, sub: `${sl.disconnectIn}"`, col: C.blue, serviceSize: ts.serviceSize, showMeter: true });
      pushT({ w: sl.gapIn * scale, kind: "gap", col: "transparent" });
    } else {
      // Side by side: meter + divider + disconnect
      pushT({ w: sl.meterIn * scale, kind: "meter-only", label: "METER", sub: `${sl.meterIn}"`, col: C.green });
      pushT({ w: sl.divIn * scale, kind: "gap", col: "transparent" });
      pushT({ w: sl.disconnectIn * scale, kind: "equip", label: ts.name || ts.category, sub: `${sl.disconnectIn}"`, col: C.amber, serviceSize: ts.serviceSize, showMeter: false });
      pushT({ w: sl.gapIn * scale, kind: "gap", col: "transparent" });
    }
  });

  const tapW = tapCanIn * scale;
  const chanW = tx;
  const totalDrawW = tapW + chanW;

  // Layout Y
  const LBL_Y = 2;                                 // tenant name
  const BOX_H = 46;                                // equipment box height
  const TOP_Y = 16;                                // top row (stacked disconnect)
  const BOX_GAP = 12;                              // gap between stacked boxes
  const BOT_Y = TOP_Y + BOX_H + BOX_GAP;           // bottom row (meter/side-by-side disconnect)
  const CONN_GAP = 8;
  const CHAN_H = 30;                               // gutter
  const CHAN_Y = BOT_Y + BOX_H + CONN_GAP;
  const TAP_H = CHAN_Y + CHAN_H - TOP_Y;           // tap spans full height
  const TAP_Y2 = TOP_Y;
  const BOTTOM = CHAN_Y + CHAN_H;
  const DIM_Y = BOTTOM + 22;
  const SVG_H = DIM_Y + 44;

  const TAP_X = PAD;
  const CHAN_X = TAP_X + tapW;
  // Gutter length = total inches of all services (without tap can and initial gap)
  const gutterIn = totalIn - tapCanIn - POST_TAP_GAP;
  const gutterFt = (gutterIn / 12);

  return (
    <div style={{ overflowX: "auto", background: C.bg, borderRadius: 2, padding: "16px 8px" }}>
      <svg width={Math.max(W, PAD + totalDrawW + PAD)} height={SVG_H}
        style={{ display: "block", fontFamily: "monospace", overflow: "visible" }}>

        {/* ── Tap Can box ── */}
        <rect x={TAP_X} y={TAP_Y2} width={tapW} height={TAP_H}
          fill={C.purple} fillOpacity={0.15} stroke={C.purple} strokeWidth={2} rx={2} />
        <text x={TAP_X + tapW / 2} y={TAP_Y2 + TAP_H / 2 - 6}
          fill={C.purple} fontSize={9} fontWeight="bold" textAnchor="middle">TAP CAN</text>
        <text x={TAP_X + tapW / 2} y={TAP_Y2 + TAP_H / 2 + 8}
          fill={C.purple} fontSize={7} textAnchor="middle" opacity={0.8}>{tapCanIn}"</text>

        {/* ── Connection: tap can → metering gutter ── */}
        <line x1={TAP_X + tapW} y1={CHAN_Y + CHAN_H / 2}
          x2={CHAN_X} y2={CHAN_Y + CHAN_H / 2}
          stroke={C.purple} strokeWidth={2} />

        {/* ── Equipment segments ── */}
        {tenantSegs.map((seg, i) => {
          const sx = CHAN_X + seg.rx;
          const cx = sx + seg.w / 2;

          if (seg.kind === "gap") {
            const gapTop = TOP_Y;
            const gapBot = CHAN_Y;
            return (
              <g key={i}>
                <line x1={sx} y1={gapTop} x2={sx} y2={gapBot}
                  stroke={C.border} strokeWidth={1} strokeDasharray="3 2" />
                {seg.w > 10 && (
                  <text x={cx} y={(gapTop + gapBot) / 2 + 3}
                    fill={C.muted} fontSize={7} textAnchor="middle">6"</text>
                )}
              </g>
            );
          }

          const isSmall = seg.w < 28;

          if (seg.kind === "equip") {
            const circR = Math.min(seg.w / 3, 14);
            const discTop = seg.showMeter ? TOP_Y : BOT_Y;
            const meterTop = BOT_Y;

            return (
              <g key={i}>
                {/* ── Tenant name above everything ── */}
                {!isSmall && (
                  <text x={cx} y={LBL_Y + 8}
                    fill={C.text} fontSize={Math.min(seg.w / 6, 8)} fontWeight="bold" textAnchor="middle">{seg.label}</text>
                )}

                {/* ── DISCONNECT SWITCH box ── */}
                <rect x={sx + 1} y={discTop} width={seg.w - 2} height={BOX_H}
                  fill={seg.col} fillOpacity={0.15} stroke={seg.col} strokeWidth={1.5} rx={1} />
                {/* Handle: straight vertical line extending down from center */}
                <line x1={cx} y1={discTop + BOX_H} x2={cx} y2={discTop + BOX_H + 8}
                  stroke={seg.col} strokeWidth={2} strokeLinecap="round" />
                {!isSmall && (
                  <text x={cx} y={discTop + BOX_H / 2 + 3}
                    fill={seg.col} fontSize={5} fontWeight="bold" textAnchor="middle" letterSpacing={0.3}>DISCONNECT SW.</text>
                )}

                {/* ── METER box (only for stacked / showMeter) ── */}
                {seg.showMeter && (<>
                  <rect x={sx + 1} y={meterTop} width={seg.w - 2} height={BOX_H}
                    fill={C.green} fillOpacity={0.08} stroke={C.green} strokeWidth={1.5} rx={1} />
                  <circle cx={cx} cy={meterTop + BOX_H * 0.42} r={circR}
                    fill="none" stroke={C.green} strokeWidth={1.5} />
                  <circle cx={cx} cy={meterTop + BOX_H * 0.42} r={2}
                    fill={C.green} fillOpacity={0.6} />
                  {!isSmall && (
                    <text x={cx} y={meterTop + BOX_H - 3}
                      fill={C.green} fontSize={6} fontWeight="bold" textAnchor="middle" letterSpacing={1}>METER</text>
                  )}
                  {/* Connector: meter → gutter */}
                  <line x1={cx} y1={meterTop + BOX_H} x2={cx} y2={CHAN_Y}
                    stroke={C.green} strokeWidth={1} strokeOpacity={0.4} />

                  {/* Connector: disconnect → meter (connector line under handle) */}
                  <line x1={cx} y1={discTop + BOX_H + 8} x2={cx} y2={meterTop}
                    stroke={C.muted} strokeWidth={1} strokeOpacity={0.5} />
                </>)}

                {/* ── Connector: disconnect → gutter (only if no meter below it) ── */}
                {!seg.showMeter && (
                  <line x1={cx} y1={discTop + BOX_H + 8} x2={cx} y2={CHAN_Y}
                    stroke={seg.col} strokeWidth={1} strokeOpacity={0.4} />
                )}
              </g>
            );
          }

          if (seg.kind === "meter-only") {
            // Standalone meter (for side-by-side layouts) — aligned with disconnect row
            const circR = Math.min(seg.w / 3, 14);
            return (
              <g key={i}>
                <rect x={sx + 1} y={BOT_Y} width={seg.w - 2} height={BOX_H}
                  fill={C.green} fillOpacity={0.08} stroke={C.green} strokeWidth={1.5} rx={1} />
                <circle cx={cx} cy={BOT_Y + BOX_H * 0.42} r={circR}
                  fill="none" stroke={C.green} strokeWidth={1.5} />
                <circle cx={cx} cy={BOT_Y + BOX_H * 0.42} r={2}
                  fill={C.green} fillOpacity={0.6} />
                {!isSmall && (
                  <text x={cx} y={BOT_Y + BOX_H - 3}
                    fill={C.green} fontSize={6} fontWeight="bold" textAnchor="middle" letterSpacing={1}>METER</text>
                )}
                <line x1={cx} y1={BOT_Y + BOX_H} x2={cx} y2={CHAN_Y}
                  stroke={C.green} strokeWidth={1} strokeOpacity={0.4} />
              </g>
            );
          }

          return null;
        })}

        {/* ── Metering Gutter (with gutter length) ── */}
        {chanW > 0 && (
          <>
            <rect x={CHAN_X} y={CHAN_Y} width={chanW} height={CHAN_H}
              fill={C.raised} stroke={C.border} strokeWidth={1.5} rx={1} />
            <text x={CHAN_X + chanW / 2} y={CHAN_Y + 11}
              fill={C.muted} fontSize={7} textAnchor="middle" letterSpacing={2}>
              METERING GUTTER
            </text>
            <text x={CHAN_X + chanW / 2} y={CHAN_Y + 24}
              fill={C.amber} fontSize={8} fontWeight="bold" textAnchor="middle">
              {gutterIn}" = {gutterFt.toFixed(2)} ft
            </text>
          </>
        )}

        {/* ── Dimension line ── */}
        <line x1={PAD} y1={DIM_Y} x2={PAD + totalDrawW} y2={DIM_Y}
          stroke={ok ? C.green : C.red} strokeWidth={1.5} />
        <polygon points={`${PAD},${DIM_Y - 4} ${PAD},${DIM_Y + 4} ${PAD - 5},${DIM_Y}`}
          fill={ok ? C.green : C.red} />
        <polygon points={`${PAD + totalDrawW},${DIM_Y - 4} ${PAD + totalDrawW},${DIM_Y + 4} ${PAD + totalDrawW + 5},${DIM_Y}`}
          fill={ok ? C.green : C.red} />
        <line x1={PAD} y1={DIM_Y - 8} x2={PAD} y2={DIM_Y + 8}
          stroke={ok ? C.green : C.red} strokeWidth={1} />
        <line x1={PAD + totalDrawW} y1={DIM_Y - 8} x2={PAD + totalDrawW} y2={DIM_Y + 8}
          stroke={ok ? C.green : C.red} strokeWidth={1} />
        <text x={PAD + totalDrawW / 2} y={DIM_Y + 16}
          fill={ok ? C.green : C.red} fontSize={10} textAnchor="middle" fontWeight="bold">
          {totalIn}" = {totalFt.toFixed(3)} ft
          {availFt > 0 ? (ok ? `  ✓ OK (${availFt} ft avail.)` : `  ✗ EXCEEDS ${availFt} ft`) : ""}
        </text>

        {/* ── Legend ── */}
        {[
          { col: C.purple, lbl: "Tap Can" },
          { col: C.green, lbl: "Meter" },
          { col: C.blue, lbl: "Disc. Switch ≤200A" },
          { col: C.amber, lbl: "Disc. Switch 400–800A" },
        ].map((l, i) => (
          <g key={i} transform={`translate(${PAD + i * 140}, ${SVG_H - 14})`}>
            <rect width={8} height={8} fill={l.col} opacity={0.8} rx={1} />
            <text x={11} y={8} fill={C.muted} fontSize={7}>{l.lbl}</text>
          </g>
        ))}
      </svg>
    </div>
  );
}

const slotPill = (slot) => {
  if (!slot) return {};
  if (slot.type === "Stacked") return S.pill(C.blue, "#062030", "#0369a1");
  if (slot.type === "Side by Side") return S.pill(C.amber, "#1e1000", "#b45309");
  return S.pill(C.red, "#1e0505", "#b91c1c");
};

// ─── Print Cover Sheet ────────────────────────────────────────────────────────
function PrintCover({ project, voltage, totalKVA, totalAmps, buildingCount }) {
  return (
    <div data-print-only style={{ display: "none", pageBreakAfter: "always", padding: 40, textAlign: "center" }}>
      <div style={{ fontSize: 36, marginBottom: 20 }}>⚡</div>
      <h1 style={{ fontSize: 28, letterSpacing: 4, marginBottom: 8 }}>ELECTRICAL ESTIMATE</h1>
      <h2 style={{ fontSize: 18, fontWeight: "normal", marginBottom: 40 }}>{project.name || "Untitled"}</h2>
      <table style={{ margin: "0 auto", borderCollapse: "collapse", fontSize: 13, textAlign: "left" }}>
        <tbody>
          {[
            ["Project #", project.number || "—"], ["Buildings", buildingCount],
            ["Global Voltage", `${voltage.label} (÷ ${voltage.factor})`],
            ["Total KVA", totalKVA.toFixed(2)], ["Total Amps", totalAmps.toFixed(2)],
          ].map(([k, v], i) => (
            <tr key={i}>
              <td style={{ padding: "6px 20px 6px 0", fontWeight: "bold" }}>{k}</td>
              <td style={{ padding: "6px 0" }}>{v}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ marginTop: 60, fontSize: 10, color: "#999" }}>
        Generated {new Date().toLocaleDateString()} — ELEC-ESTIMATOR
      </div>
    </div>
  );
}

// ─── Inline read/write text helper ────────────────────────────────────────────
const readVal = { color: C.text, fontSize: 12, fontFamily: "monospace", letterSpacing: 0.5 };
const readMuted = { color: C.muted, fontSize: 11, fontFamily: "monospace" };
const editBtn = (editing) => ({
  background: editing ? C.green : "rgba(245,158,11,0.12)",
  color: editing ? "#000" : C.amber,
  border: `1px solid ${editing ? C.green : C.amber}`,
  padding: "4px 12px", cursor: "pointer", fontFamily: "monospace",
  fontSize: 10, fontWeight: "bold", letterSpacing: 1, borderRadius: 2,
  transition: "all .15s",
});

// ─── Settings Panel ───────────────────────────────────────────────────────────
function SettingsPanel({ voltage, setVoltage, categories, setCategories, tenants, setTenants }) {
  const bp = useBreakpoint();
  const [editVolt, setEditVolt] = useState(false);
  const [editCats, setEditCats] = useState(false);
  const [newCatName, setNewCatName] = useState("");
  const [addingCat, setAddingCat] = useState(false);
  const [confirmDel, setConfirmDel] = useState(null);

  const updateCategory = (id, f, v) => setCategories((p) => p.map((c) => (c.id === id ? { ...c, [f]: v } : c)));
  const doDeleteCat = (id, name) => {
    if (tenants.some((t) => t.category === name)) { setConfirmDel(null); return; }
    setCategories((p) => p.filter((c) => c.id !== id));
    setConfirmDel(null);
  };
  const submitNewCat = () => {
    const name = newCatName.trim();
    if (!name || categories.find((c) => c.name === name)) return;
    setCategories((p) => [...p, { id: uid(), name, wattsPerSqft: 20, isHousePanel: false }]);
    setNewCatName(""); setAddingCat(false);
  };

  const VButtons = ({ v, setV }) => (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
      {VOLTAGE_PRESETS.map((p) => (
        <button key={p.label} style={S.btn(v.label === p.label ? C.amber : C.raised, v.label === p.label ? C.bg : C.muted)}
          onClick={() => setV(p)}>{p.label}</button>
      ))}
    </div>
  );

  return (
    <div data-print-hide>
      {/* ── Voltage System ── */}
      <div style={S.card} data-r-card>
        <div style={S.cardH} data-r-cardh>
          <span>VOLTAGE SYSTEM</span>
          <button style={editBtn(editVolt)} onClick={() => setEditVolt((e) => !e)}>
            {editVolt ? "✓ SAVE" : "✎ EDIT FACTORS"}
          </button>
        </div>

        <div style={S.rg2(bp.mobile)}>
          <div>
            <label style={S.lbl}>Global Voltage</label>
            <VButtons v={voltage} setV={setVoltage} />
            <div style={{ marginTop: 6, display: "flex", gap: 8, alignItems: "center" }}>
              <span style={readMuted}>factor:</span>
              {editVolt
                ? <input style={{ ...S.inp, width: 80 }} type="number" step="0.001" value={voltage.factor}
                  onChange={(e) => setVoltage({ label: "Custom", factor: Number(e.target.value) })} />
                : <span style={readVal}>{voltage.factor}</span>}
            </div>
          </div>
        </div>
      </div>

      {/* ── Categories ── */}
      <div style={S.card} data-r-card>
        <div style={S.cardH} data-r-cardh>
          <span>CATEGORIES &amp; W/SQFT</span>
          <button style={editBtn(editCats)} onClick={() => { setEditCats((e) => !e); setAddingCat(false); setConfirmDel(null); }}>
            {editCats ? "✓ SAVE" : "✎ EDIT"}
          </button>
        </div>

        {editCats ? (
          <div style={{ overflowX: "auto" }} data-r-tblwrap>
            <table style={S.tbl}>
              <thead>
                <tr>{["Name", "W / ft²", "House Panel (fixed KVA)", ""].map((h) => <th key={h} style={S.th}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {categories.map((cat) => (
                  <CatRow key={cat.id} cat={cat} tenants={tenants} confirmDel={confirmDel}
                    setConfirmDel={setConfirmDel} updateCategory={updateCategory}
                    setCategories={setCategories} setTenants={setTenants} doDeleteCat={doDeleteCat} />
                ))}
                {addingCat ? (
                  <tr><td style={S.td} colSpan={4}>
                    <div style={S.row}>
                      <input style={{ ...S.inp, width: 200 }} value={newCatName} autoFocus placeholder="Category name"
                        onChange={(e) => setNewCatName(e.target.value)}
                        onKeyDown={(e) => { if (e.key === "Enter") submitNewCat(); if (e.key === "Escape") setAddingCat(false); }} />
                      <button style={S.btn()} onClick={submitNewCat}>Add</button>
                      <button style={S.ghost} onClick={() => { setAddingCat(false); setNewCatName(""); }}>Cancel</button>
                    </div>
                  </td></tr>
                ) : (
                  <tr><td colSpan={4} style={S.td}>
                    <button style={S.ghost} onClick={() => setAddingCat(true)}>+ New Category</button>
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 16 }}>
            {categories.map((cat) => (
              <div key={cat.id} style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
                <span style={readVal}>{cat.name}</span>
                <span style={readMuted}>
                  {cat.isHousePanel ? "fixed KVA" : `${cat.wattsPerSqft} W/ft²`}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Category row ─────────────────────────────────────────────────────────────
function CatRow({ cat, tenants, confirmDel, setConfirmDel, updateCategory, setCategories, setTenants, doDeleteCat }) {
  return (
    <>
      <tr>
        <td style={S.td}>
          <input style={{ ...S.inp, width: 160 }} value={cat.name}
            onChange={(e) => {
              const nv = e.target.value;
              setCategories((prev) => prev.map((c) => (c.id === cat.id ? { ...c, name: nv } : c)));
              setTenants((prev) => prev.map((t) => (t.category === cat.name ? { ...t, category: nv } : t)));
            }} />
        </td>
        <td style={S.td}>
          {cat.isHousePanel ? <span style={{ color: C.muted, fontSize: 10 }}>Fixed KVA per tenant</span> : (
            <div style={S.row}>
              <input style={{ ...S.inp, width: 80 }} type="number" value={cat.wattsPerSqft}
                onChange={(e) => updateCategory(cat.id, "wattsPerSqft", Number(e.target.value))} />
              <span style={{ color: C.muted, fontSize: 10 }}>W/ft²</span>
            </div>
          )}
        </td>
        <td style={S.td}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
            <input type="checkbox" checked={cat.isHousePanel}
              onChange={(e) => updateCategory(cat.id, "isHousePanel", e.target.checked)} />
            <span style={{ color: C.muted, fontSize: 11 }}>Yes — uses HP voltage</span>
          </label>
        </td>
        <td style={S.td}>
          {tenants.some((t) => t.category === cat.name)
            ? <span style={{ color: C.muted, fontSize: 9 }}>in use</span>
            : <button style={{ ...S.ghost, color: C.red, borderColor: "#450a0a" }}
              onClick={() => setConfirmDel({ id: cat.id, name: cat.name })}>✕ Delete</button>}
        </td>
      </tr>
      {confirmDel?.id === cat.id && (
        <tr><td colSpan={4} style={{ ...S.td, padding: 0 }}>
          <ConfirmInline message={`Delete category "${cat.name}"?`}
            onConfirm={() => doDeleteCat(cat.id, cat.name)} onCancel={() => setConfirmDel(null)} />
        </td></tr>
      )}
    </>
  );
}

// ─── Tenant Row ───────────────────────────────────────────────────────────────
function TenantRow({ t, computed, catMap, categories, updateTenant, removeTenant }) {
  const co = computed.find((c) => c.id === t.id) || {};
  const cat = catMap[t.category] || {};
  return (
    <tr>
      <td style={S.td}>
        <input style={S.inp} value={t.name} onChange={(e) => updateTenant(t.id, "name", e.target.value)} placeholder="Tenant name" />
      </td>
      <td style={S.td}>
        <select style={{ ...S.sel, width: 130 }} value={t.category}
          onChange={(e) => updateTenant(t.id, "category", e.target.value)}>
          {categories.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
        </select>
      </td>
      <td style={S.td}>
        {cat.isHousePanel ? <span style={{ color: C.muted }}>—</span>
          : <input style={{ ...S.inp, width: 80 }} type="number" value={t.sqft}
            onChange={(e) => updateTenant(t.id, "sqft", e.target.value)} placeholder="sqft" />}
      </td>
      <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>{cat.isHousePanel ? "—" : cat.wattsPerSqft}</td>
      <td style={S.td}>
        {cat.isHousePanel
          ? <input style={{ ...S.inp, width: 68 }} type="number" value={t.fixedKVA}
            onChange={(e) => updateTenant(t.id, "fixedKVA", e.target.value)} placeholder="KVA" />
          : <strong style={{ color: C.amber }}>{co.kva?.toFixed(2) || "—"}</strong>}
      </td>
      <td style={{ ...S.td, color: C.blue }}>{co.amps?.toFixed(2) || "—"}</td>
      <td style={S.td}>
        <select style={{ ...S.sel, width: 110, color: t.serviceOverride ? C.amber : C.green, fontWeight: "bold", fontSize: 11 }}
          value={t.serviceOverride || ""}
          onChange={(e) => updateTenant(t.id, "serviceOverride", e.target.value)}>
          <option value="">Auto ({co.autoServiceSize || "—"}A)</option>
          {SELECTABLE_SIZES.map((sz) => <option key={sz} value={sz}>{sz}A</option>)}
        </select>
      </td>
      <td style={S.td}>{co.slot && <span style={slotPill(co.slot)}>{co.slot.totalIn}"</span>}</td>
      <td style={S.td} data-print-hide>
        <button style={{ ...S.ghost, color: C.red, borderColor: "#450a0a" }} onClick={() => removeTenant(t.id)}>✕</button>
      </td>
    </tr>
  );
}

// ─── Wireway Rules Card (view/edit pattern) ──────────────────────────────────
function WirewayRulesCard({ wirewayRules, setWirewayRules, tapCanIn, setTapCanIn }) {
  const [editing, setEditing] = useState(false);

  const updateRule = (id, field, value) =>
    setWirewayRules((prev) => prev.map((r) => (r.id === id ? { ...r, [field]: value } : r)));

  return (
    <div style={S.card} data-r-card>
      <div style={S.cardH} data-r-cardh>
        <span>WIREWAY RULES</span>
        <button style={editBtn(editing)} onClick={() => setEditing((e) => !e)}>
          {editing ? "✓ SAVE" : "✎ EDIT"}
        </button>
      </div>

      <div style={{ overflowX: "auto" }} data-r-tblwrap>
        <table style={S.tbl}>
          <thead>
            <tr>
              {["Service Size", "Type", "Disconnect", "Divider", "Meter", "Next Gap", "TOTAL"].map((h) => (
                <th key={h} style={S.th}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {wirewayRules.map((r) => {
              const total = calcSlotTotal(r);
              return (
                <tr key={r.id}>
                  <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{r.label}</td>
                  {editing ? (
                    <>
                      <td style={S.td}>
                        <select style={{ ...S.sel, width: 140 }} value={r.type}
                          onChange={(e) => updateRule(r.id, "type", e.target.value)}>
                          {WIREWAY_TYPES.map((wt) => <option key={wt} value={wt}>{wt}</option>)}
                        </select>
                      </td>
                      <td style={S.td}>
                        <input style={{ ...S.inp, width: 50 }} type="number" value={r.disconnectIn}
                          onChange={(e) => updateRule(r.id, "disconnectIn", Number(e.target.value))} />
                      </td>
                      <td style={S.td}>
                        {r.type === "Stacked"
                          ? <span style={readMuted}>—</span>
                          : <input style={{ ...S.inp, width: 50 }} type="number" value={r.divIn}
                            onChange={(e) => updateRule(r.id, "divIn", Number(e.target.value))} />}
                      </td>
                      <td style={S.td}>
                        {r.type === "Stacked"
                          ? <span style={readMuted}>stacked</span>
                          : <input style={{ ...S.inp, width: 50 }} type="number" value={r.meterIn}
                            onChange={(e) => updateRule(r.id, "meterIn", Number(e.target.value))} />}
                      </td>
                      <td style={S.td}>
                        <input style={{ ...S.inp, width: 50 }} type="number" value={r.gapIn}
                          onChange={(e) => updateRule(r.id, "gapIn", Number(e.target.value))} />
                      </td>
                    </>
                  ) : (
                    <>
                      <td style={{ ...S.td, color: C.muted }}>{r.type}</td>
                      <td style={{ ...S.td, color: C.blue }}>{r.disconnectIn}"</td>
                      <td style={{ ...S.td, color: C.muted }}>{r.type === "Stacked" ? "—" : r.divIn + '"'}</td>
                      <td style={{ ...S.td, color: C.green }}>{r.type === "Stacked" ? "stacked" : r.meterIn + '"'}</td>
                      <td style={{ ...S.td, color: C.muted }}>{r.gapIn}"</td>
                    </>
                  )}
                  <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{total}"</td>
                </tr>
              );
            })}
            {/* Tap Can row */}
            <tr>
              <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>Tap Can</td>
              {editing ? (
                <>
                  <td style={S.td}>Fixed at start</td>
                  <td style={S.td}>—</td>
                  <td style={S.td}>—</td>
                  <td style={S.td}>—</td>
                  <td style={S.td}>—</td>
                  <td style={S.td}>
                    <input style={{ ...S.inp, width: 50 }} type="number" value={tapCanIn}
                      onChange={(e) => setTapCanIn(Number(e.target.value) || 0)} />
                  </td>
                </>
              ) : (
                <>
                  <td style={{ ...S.td, color: C.muted }}>Fixed at start</td>
                  <td style={{ ...S.td, color: C.muted }}>—</td>
                  <td style={{ ...S.td, color: C.muted }}>—</td>
                  <td style={{ ...S.td, color: C.muted }}>—</td>
                  <td style={{ ...S.td, color: C.muted }}>—</td>
                  <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{tapCanIn}"</td>
                </>
              )}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [screen, setScreen] = useState("home");
  const [project, setProject] = useState(BLANK_PROJECT);
  const [voltage, setVoltage] = useState({ label: "120/208V", factor: 0.360 });
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES.map((c) => ({ ...c, id: uid() })));
  const [tenants, setTenants] = useState([{ ...blankTenant(1, "House Panel"), name: "House Panel", fixedKVA: "10" }]);
  const [wireway, setWireway] = useState(BLANK_WIREWAY);
  const [tapCanIn, setTapCanIn] = useState(DEFAULT_TAP_CAN_IN);
  const [wirewayRules, setWirewayRules] = useState(DEFAULT_WIREWAY_RULES.map((r) => ({ ...r })));
  const fileRef = useRef();
  const bp = useBreakpoint();

  const printStyleRef = useRef(false);
  if (!printStyleRef.current && typeof document !== "undefined") {
    const style = document.createElement("style");
    style.textContent = INJECTED_CSS;
    document.head.appendChild(style);
    printStyleRef.current = true;
  }

  const catMap = useMemo(() => Object.fromEntries(categories.map((c) => [c.name, c])), [categories]);

  const computed = useMemo(() =>
    tenants.map((t) => {
      const cat = catMap[t.category] || {};
      const isHP = !!cat.isHousePanel;
      const factor = voltage.factor;
      const kva = isHP && t.fixedKVA !== "" ? Number(t.fixedKVA) : calcKVA(t.sqft, cat.wattsPerSqft);
      const amps = calcAmps(kva, factor);
      const autoSvcSz = getServiceSize(amps);
      const svcSz = t.serviceOverride ? Number(t.serviceOverride) : autoSvcSz;
      return { ...t, isHP, factor, kva, amps, autoServiceSize: autoSvcSz, serviceSize: svcSz, slot: getSlot(svcSz, wirewayRules) };
    }), [tenants, catMap, voltage.factor, wirewayRules]);

  const totalKVA = computed.reduce((s, t) => s + t.kva, 0);
  const totalAmps = totalKVA / voltage.factor;

  const catSummary = useMemo(() =>
    categories.map((cat) => {
      const rows = computed.filter((t) => t.category === cat.name);
      return {
        name: cat.name, isHP: cat.isHousePanel,
        totalSqft: rows.reduce((s, t) => s + (Number(t.sqft) || 0), 0),
        totalKVA: rows.reduce((s, t) => s + t.kva, 0)
      };
    }).filter((c) => c.totalKVA > 0 || c.totalSqft > 0), [categories, computed]);

  const allBuildings = useMemo(() =>
    [...new Set([...Array.from({ length: project.buildings }, (_, i) => i + 1), ...tenants.map((t) => t.building)])].sort((a, b) => a - b),
    [project.buildings, tenants]);

  const buildingWireway = useCallback((b) => {
    const raw = computed.filter((t) => t.building === b);
    // Sort: house panels first, then others
    const slots = [...raw].sort((a, bb) => {
      if (a.isHP && !bb.isHP) return -1;
      if (!a.isHP && bb.isHP) return 1;
      return 0;
    });
    const tenantsIn = slots.reduce((s, t) => s + t.slot.totalIn, 0);
    const totalIn = Number(tapCanIn) + POST_TAP_GAP + tenantsIn;
    const totalFt = totalIn / 12;
    const availFt = Number(wireway.availableWallFeet) || 0;
    return {
      tenantSlots: slots, tenantsIn, totalIn, totalFt, availFt,
      ok: availFt > 0 && totalFt <= availFt
    };
  }, [computed, wireway, tapCanIn]);

  // ── Download helper (works in sandboxed iframes) ──
  const downloadRef = useRef();
  const triggerDownload = useCallback((blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = downloadRef.current;
    if (a) {
      a.href = url;
      a.download = filename;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  }, []);

  const handleSave = useCallback(() => {
    const data = {
      project, voltage,
      categories: categories.map(({ id, ...r }) => r),
      tenants: tenants.map(({ id, ...r }) => r),
      wireway, tapCanIn, wirewayRules: wirewayRules.map(({ id, ...r }) => r),
      savedAt: new Date().toISOString()
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    triggerDownload(blob, `${project.name || "estimate"}_${project.number || "v1"}.json`);
  }, [project, voltage, categories, tenants, wireway, tapCanIn, wirewayRules, triggerDownload]);

  const handleExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const rows = [
      ["ELECTRICAL ESTIMATE"],
      ["Project", project.name, "Number", project.number],
      ["Global Voltage", voltage.label, "Factor", voltage.factor],
      [],
      ["Building", "Tenant", "Category", "S.F.", "W/sqft", "KVA", "Amps", "Service Size", "Wireway Type", "Wireway (in)"],
      ...computed.map((t) => [t.building, t.name, t.category, t.isHP ? "—" : (t.sqft || 0),
      t.isHP ? "fixed" : (catMap[t.category]?.wattsPerSqft || 0),
      +t.kva.toFixed(2), +t.amps.toFixed(4), t.serviceSize, t.slot.type, t.slot.totalIn]),
      [],
      ["SUMMARY BY CATEGORY"], ["Category", "Total S.F.", "KVA"],
      ...catSummary.map((c) => [c.name + " S.F.:", c.isHP ? "—" : c.totalSqft, +c.totalKVA.toFixed(2)]),
      ["Total KVA", "", +totalKVA.toFixed(2)],
      ["Total Amps", `= ${totalKVA.toFixed(2)} / ${voltage.factor}`, +totalAmps.toFixed(6)],
      [], ["WIREWAY PER BUILDING"],
      ...allBuildings.flatMap((b) => {
        const bw = buildingWireway(b);
        return [[`Building ${b}`], ["Tap Can", tapCanIn + '"'], ["Gap (tap→disc)", POST_TAP_GAP + '"'],
        ...bw.tenantSlots.map((t) => [t.name || t.category, t.serviceSize + "A", t.slot.type, t.slot.totalIn + '"']),
        ["Total (in)", bw.totalIn, "Total (ft)", +bw.totalFt.toFixed(4)],
        ["Available (ft)", bw.availFt, "Status", bw.ok ? "OK" : "EXCEEDS"], []];
      }),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Estimate");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    triggerDownload(blob, `${project.name || "estimate"}_${project.number || "v1"}.xlsx`);
  }, [project, voltage, computed, catMap, catSummary, totalKVA, totalAmps, allBuildings, buildingWireway, tapCanIn, triggerDownload]);

  const handleLoad = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const d = JSON.parse(ev.target.result);
        if (d.project) setProject(d.project);
        if (d.voltage) setVoltage(d.voltage);
        if (d.categories) setCategories(d.categories.map((c) => ({ ...c, id: uid() })));
        if (d.tenants) setTenants(d.tenants.map((t) => ({ ...t, id: uid() })));
        if (d.wireway) setWireway(d.wireway);
        if (d.tapCanIn !== undefined) setTapCanIn(Number(d.tapCanIn));
        if (d.wirewayRules) setWirewayRules(d.wirewayRules.map((r) => ({ ...r, id: uid() })));
        setScreen("editor");
      } catch { /* silent */ }
    };
    r.readAsText(file); e.target.value = "";
  };

  const handlePrint = useCallback(() => {
    // Collect all the HTML we need: summary tables + wireway SVGs
    // We render into an off-screen div, serialize, then open in new window
    const printRoot = document.createElement("div");
    printRoot.id = "print-root";
    printRoot.style.cssText = "position:absolute;left:-9999px;top:0;";
    document.body.appendChild(printRoot);

    // Wait a tick then grab content from the current DOM
    setTimeout(() => {
      // Grab the main content area
      const mainEl = document.querySelector("[data-r-main]");
      if (!mainEl) { document.body.removeChild(printRoot); return; }

      // Build print HTML
      const css = `
        body { font-family: 'Courier New', monospace; font-size: 11px; color: #000; background: #fff; margin: 20px; }
        h1 { font-size: 22px; letter-spacing: 3px; margin: 0 0 4px 0; }
        h2 { font-size: 14px; font-weight: normal; color: #666; margin: 0 0 20px 0; }
        h3 { font-size: 11px; letter-spacing: 2px; text-transform: uppercase; color: #666; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin: 20px 0 10px 0; }
        table { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 16px; page-break-inside: avoid; }
        th { background: #f0f0f0; color: #666; padding: 6px 8px; text-align: left; font-size: 8px; letter-spacing: 1px; border-bottom: 1px solid #ccc; }
        td { padding: 6px 8px; border-bottom: 1px solid #eee; }
        .stat-row { display: flex; gap: 20px; margin-bottom: 16px; }
        .stat { text-align: center; flex: 1; border: 1px solid #ddd; padding: 10px; }
        .stat-val { font-size: 18px; font-weight: bold; }
        .stat-lbl { font-size: 8px; color: #888; letter-spacing: 2px; text-transform: uppercase; margin-top: 4px; }
        .bold { font-weight: bold; }
        .muted { color: #888; }
        svg { display: block; margin: 10px 0; }
        svg text { fill: #333 !important; }
        svg rect { stroke: #666 !important; }
        svg line { stroke: #666 !important; }
        .page-break { page-break-before: always; }
        @media print { body { margin: 10px; } }
      `;

      // Build summary content
      let html = '<div class="stat-row">';
      [{ v: project.name || "—", l: "Project" },
      { v: totalKVA.toFixed(2) + " KVA", l: "Total KVA" },
      { v: totalAmps.toFixed(2) + " A", l: "Total Amps ÷ " + voltage.factor },
      { v: project.buildings, l: "Buildings" }
      ].forEach(s => {
        html += '<div class="stat"><div class="stat-val">' + s.v + '</div><div class="stat-lbl">' + s.l + '</div></div>';
      });
      html += '</div>';

      // Voltage system info
      html += '<div class="stat-row">';
      html += '<div class="stat"><div class="stat-lbl">GLOBAL VOLTAGE</div><div class="stat-val" style="font-size:14px">' + voltage.label + '</div></div>';
      html += '</div>';

      // Building tables
      allBuildings.forEach(b => {
        const bComp = computed.filter(t => t.building === b);
        const bw = buildingWireway(b);
        html += '<h3>BUILDING ' + b + '</h3>';
        html += '<table><thead><tr>';
        ["Tenant", "Category", "S.F.", "W/ft²", "KVA", "Amps", "Service", "Wireway"].forEach(h => { html += '<th>' + h + '</th>'; });
        html += '</tr></thead><tbody>';
        bComp.forEach(t => {
          html += '<tr>';
          html += '<td>' + (t.name || '(unnamed)') + '</td>';
          html += '<td class="muted">' + t.category + '</td>';
          html += '<td>' + (t.isHP ? '—' : (t.sqft || 0)) + '</td>';
          html += '<td class="muted">' + (t.isHP ? 'fixed' : (catMap[t.category]?.wattsPerSqft || '')) + '</td>';
          html += '<td class="bold">' + t.kva.toFixed(2) + '</td>';
          html += '<td>' + t.kva.toFixed(2) + ' / ' + t.factor + ' = <strong>' + t.amps.toFixed(4) + '</strong></td>';
          html += '<td>' + t.serviceSize + 'A</td>';
          html += '<td>' + t.slot.totalIn + '"</td>';
          html += '</tr>';
        });
        html += '<tr style="background:#f5f5f5"><td colspan="4" class="muted" style="font-size:8px;letter-spacing:2px">SUBTOTAL B' + b + '</td>';
        html += '<td class="bold">' + bComp.reduce((s, t) => s + t.kva, 0).toFixed(2) + '</td>';
        html += '<td class="bold">' + bComp.reduce((s, t) => s + t.amps, 0).toFixed(4) + '</td>';
        html += '<td></td>';
        html += '<td class="bold">' + bw.totalIn + '" = ' + bw.totalFt.toFixed(2) + ' ft ' + (bw.availFt > 0 ? (bw.ok ? '✓' : '✗') : '') + '</td>';
        html += '</tr></tbody></table>';
      });

      // Category summary
      html += '<h3>SUMMARY BY CATEGORY</h3><table><thead><tr><th>Category</th><th>Total S.F.</th><th>KVA</th></tr></thead><tbody>';
      catSummary.forEach(c => {
        html += '<tr><td>' + c.name + ' S.F.:</td><td>' + (c.isHP ? '—' : c.totalSqft.toLocaleString()) + '</td><td class="bold">' + c.totalKVA.toFixed(2) + '</td></tr>';
      });
      html += '<tr style="background:#f5f5f5;border-top:2px solid #ccc"><td class="muted" style="font-size:8px;letter-spacing:2px">TOTAL KVA</td><td></td><td class="bold" style="font-size:14px">' + totalKVA.toFixed(2) + '</td></tr>';
      html += '<tr style="background:#f5f5f5"><td class="muted" style="font-size:8px;letter-spacing:2px">TOTAL AMPS</td><td></td><td class="bold" style="font-size:14px">' + totalAmps.toFixed(2) + '</td></tr>';
      html += '</tbody></table>';

      // Wireway diagrams - grab SVGs from DOM
      html += '<div class="page-break"></div>';
      html += '<h3>WIREWAY DIAGRAMS</h3>';
      const svgs = document.querySelectorAll("svg");
      svgs.forEach((svg, idx) => {
        html += '<div style="margin:16px 0;overflow:visible">' + svg.outerHTML + '</div>';
      });

      // Open print window
      const win = window.open("", "_blank", "width=900,height=700");
      if (win) {
        win.document.write('<!DOCTYPE html><html><head><title>Electrical Estimate - ' + (project.name || 'Print') + '</title><style>' + css + '</style></head><body>');
        win.document.write('<h1>⚡ ELECTRICAL ESTIMATE</h1>');
        win.document.write('<h2>' + (project.name || 'Untitled') + ' · ' + (project.number || '—') + '</h2>');
        win.document.write(html);
        win.document.write('</body></html>');
        win.document.close();
        setTimeout(() => { win.print(); }, 500);
      }

      document.body.removeChild(printRoot);
    }, 100);
  }, [project, voltage, totalKVA, totalAmps, allBuildings, computed, buildingWireway, catMap, catSummary]);

  const updateTenant = useCallback((id, f, v) => setTenants((p) => p.map((t) => (t.id === id ? { ...t, [f]: v } : t))), []);
  const removeTenant = useCallback((id) => setTenants((p) => p.filter((t) => t.id !== id)), []);
  const addTenant = (b) => setTenants((p) => [...p, blankTenant(b, categories.find((c) => !c.isHousePanel)?.name || "Retail")]);
  const addHP = (b) => setTenants((p) => [...p, { ...blankTenant(b, "House Panel"), name: "House Panel", fixedKVA: "10" }]);
  const syncBuildings = (n) => {
    setProject((p) => ({ ...p, buildings: Number(n) }));
    setTenants((p) => { const f = p.filter((t) => t.building <= Number(n)); return f.length ? f : [{ ...blankTenant(1, "House Panel"), name: "House Panel", fixedKVA: "10" }]; });
  };

  // ── HOME ─────────────────────────────────────────────────────────────────
  if (screen === "home") return (
    <div style={S.app}>
      <a ref={downloadRef} style={{ display: "none" }} />
      <div style={S.hdr} data-r-hdr><div><div style={S.logo} data-r-logo>⚡ ELEC-ESTIMATOR</div></div></div>
      <div style={{ ...S.main, maxWidth: 580, paddingTop: bp.mobile ? 32 : 56 }} data-r-main>
        <div style={{ textAlign: "center", marginBottom: bp.mobile ? 24 : 40 }}>
          <div style={{ fontSize: bp.mobile ? 32 : 42, marginBottom: 12 }}>⚡</div>
          <h1 style={{ color: C.amber, fontSize: bp.mobile ? 18 : 24, margin: 0, letterSpacing: 4, fontWeight: "bold" }}>ELECTRICAL ESTIMATOR</h1>
          <p style={{ color: C.muted, fontSize: 11, marginTop: 8, letterSpacing: 1 }}>KVA · AMPS · SERVICE SIZE · WIREWAY AUTO-CALC</p>
        </div>

        <div data-r-home-grid style={{ display: "grid", gridTemplateColumns: bp.mobile ? "1fr" : "1fr 1fr", gap: 14, marginBottom: 14 }}>
          <div style={{ ...S.card, cursor: "pointer", border: `1px solid ${C.amber}`, textAlign: "center", padding: 32, transition: "opacity .15s" }}
            onClick={() => {
              setProject(BLANK_PROJECT); setCategories(DEFAULT_CATEGORIES.map((c) => ({ ...c, id: uid() })));
              setTenants([{ ...blankTenant(1, "House Panel"), name: "House Panel", fixedKVA: "10" }]); setWireway(BLANK_WIREWAY);
              setTapCanIn(DEFAULT_TAP_CAN_IN);
              setWirewayRules(DEFAULT_WIREWAY_RULES.map((r) => ({ ...r })));
              setVoltage({ label: "120/208V", factor: 0.360 });
              setScreen("editor");
            }}>
            <div style={{ fontSize: 30, marginBottom: 10 }}>📋</div>
            <div style={{ color: C.amber, fontWeight: "bold", letterSpacing: 2, fontSize: 12 }}>NEW PROJECT</div>
            <div style={{ color: C.muted, fontSize: 10, marginTop: 6 }}>Start from scratch</div>
          </div>
          <div style={{ ...S.card, cursor: "pointer", textAlign: "center", padding: bp.mobile ? 20 : 32 }} onClick={() => fileRef.current.click()}>
            <div style={{ fontSize: 30, marginBottom: 10 }}>📂</div>
            <div style={{ color: C.text, fontWeight: "bold", letterSpacing: 2, fontSize: 12 }}>LOAD FILE</div>
            <div style={{ color: C.muted, fontSize: 10, marginTop: 6 }}>Open saved .json</div>
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".json" style={{ display: "none" }} onChange={handleLoad} />

        <div style={{ ...S.card, borderColor: C.border }}>
          <div style={{ color: C.muted, fontSize: 11, lineHeight: 2 }}>
            💾 Save as <strong style={{ color: C.amber }}>.json</strong> — reload the full project<br />
            📊 Export to <strong style={{ color: C.blue }}>.xlsx</strong> — for sharing<br />
            🖨 <strong style={{ color: C.green }}>Print / PDF</strong> — from the summary view<br />
            🚫 No server · No database · Everything local
          </div>
        </div>
      </div>
    </div>
  );

  // ── EDITOR ───────────────────────────────────────────────────────────────
  return (
    <div style={S.app}>
      <a ref={downloadRef} style={{ display: "none" }} />
      <PrintCover project={project} voltage={voltage}
        totalKVA={totalKVA} totalAmps={totalAmps} buildingCount={project.buildings} />

      <div style={S.hdr} data-print-hide data-r-hdr>
        <div>
          <div style={S.logo} data-r-logo>⚡ ELEC-ESTIMATOR</div>
          <div style={S.sub}>{project.name || "Untitled"} · {project.number || "—"}</div>
        </div>
        <div style={S.nav} data-r-nav>
          {["editor", "wireway", "results"].map((sc) => (
            <button key={sc} style={S.navB(screen === sc)} onClick={() => setScreen(sc)}>
              {sc === "editor" ? "DATA" : sc === "wireway" ? "WIREWAY" : "SUMMARY"}
            </button>
          ))}
          <button style={S.btn()} onClick={handleSave}>💾 .JSON</button>
          <button style={S.btn(C.blue, "#fff")} onClick={handleExcel}>📊 EXCEL</button>
          <button style={S.btn(C.green, "#000")} onClick={handlePrint}>🖨 PRINT</button>
          <button style={S.ghost} onClick={() => setScreen("home")}>← HOME</button>
        </div>
      </div>

      <div style={S.main} data-r-main>
        {screen === "editor" && (<>
          <div style={S.card} data-r-card>
            <div style={S.cardH} data-r-cardh><span>PROJECT</span></div>
            <div style={S.rg3(bp.mobile)}>
              <div><label style={S.lbl}>Project Name</label>
                <input style={S.inp} value={project.name} onChange={(e) => setProject((p) => ({ ...p, name: e.target.value }))} placeholder="Plaza North" /></div>
              <div><label style={S.lbl}>Project Number</label>
                <input style={S.inp} value={project.number} onChange={(e) => setProject((p) => ({ ...p, number: e.target.value }))} placeholder="2026-001" /></div>
              <div><label style={S.lbl}>Number of Buildings</label>
                <input style={S.inp} type="number" min={1} value={project.buildings} onChange={(e) => syncBuildings(e.target.value)} /></div>
            </div>
          </div>

          <SettingsPanel voltage={voltage} setVoltage={setVoltage}
            categories={categories} setCategories={setCategories} tenants={tenants} setTenants={setTenants} />

          {allBuildings.map((b) => {
            const bComp = computed.filter((t) => t.building === b);
            const bKVA = bComp.reduce((s, t) => s + t.kva, 0);
            const bAmps = bComp.reduce((s, t) => s + t.amps, 0);
            return (
              <div key={b} style={S.card} data-r-card>
                <div style={S.cardH} data-r-cardh>
                  <span>BUILDING {b}</span>
                  <span style={{ color: C.muted, fontSize: 9 }}>
                    KVA: <strong style={{ color: C.amber }}>{bKVA.toFixed(2)}</strong>
                    &nbsp;·&nbsp;Amps: <strong style={{ color: C.blue }}>{bAmps.toFixed(2)}</strong>
                  </span>
                </div>
                <div style={{ overflowX: "auto" }} data-r-tblwrap>
                  <table style={S.tbl}>
                    <thead><tr>{["Name", "Category", "S.F.", "W/ft²", "KVA", "Amps", "Service", "Wireway", ""].map((h) => (
                      <th key={h} style={S.th}>{h}</th>
                    ))}</tr></thead>
                    <tbody>
                      {tenants.filter((t) => t.building === b).map((t) => (
                        <TenantRow key={t.id} t={t} computed={computed} catMap={catMap}
                          categories={categories} updateTenant={updateTenant} removeTenant={removeTenant} />
                      ))}
                    </tbody>
                  </table>
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 12 }} data-print-hide>
                  <button style={S.ghost} onClick={() => addTenant(b)}>+ Tenant</button>
                  <button style={{ ...S.ghost, color: C.lime }} onClick={() => addHP(b)}>+ House Panel</button>
                </div>
              </div>
            );
          })}
        </>)}

        {screen === "wireway" && (<>
          <WirewayRulesCard wirewayRules={wirewayRules} setWirewayRules={setWirewayRules} tapCanIn={tapCanIn} setTapCanIn={setTapCanIn} />

          {allBuildings.map((b) => {
            const bw = buildingWireway(b);
            return (
              <div key={b} style={S.card} data-r-card>
                <div style={S.cardH} data-r-cardh>
                  <span>BUILDING {b} — WIREWAY</span>
                  <span style={{ color: bw.availFt > 0 ? (bw.ok ? C.green : C.red) : C.amber, fontWeight: "bold", fontSize: 11 }}>
                    {bw.totalIn}" = {bw.totalFt.toFixed(3)} ft{bw.availFt > 0 ? (bw.ok ? " ✓ OK" : " ✗ EXCEEDS") : ""}
                  </span>
                </div>

                {/* Inline wall space check */}
                <div style={{ display: "flex", gap: 16, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }} data-r-ww-inline>
                  <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <label style={{ ...S.lbl, margin: 0, whiteSpace: "nowrap" }}>Available Wall Space (ft)</label>
                    <input style={{ ...S.inp, width: 70 }} type="number" value={wireway.availableWallFeet}
                      onChange={(e) => setWireway((w) => ({ ...w, availableWallFeet: e.target.value }))} placeholder="ft" />
                    {bw.availFt > 0 && (
                      <span style={{
                        fontSize: 11, fontWeight: "bold", letterSpacing: 1,
                        color: bw.ok ? C.green : C.red,
                      }}>
                        {bw.ok ? "✓ FITS" : `✗ OVER by ${(bw.totalFt - bw.availFt).toFixed(2)} ft`}
                      </span>
                    )}
                  </div>
                </div>
                <div style={{ overflowX: "auto" }} data-r-tblwrap>
                  <table style={{ ...S.tbl, marginBottom: 16 }}>
                    <thead><tr>{["Element", "Service", "Type", "Calculation", "Inches"].map((h) => (
                      <th key={h} style={S.th}>{h}</th>
                    ))}</tr></thead>
                    <tbody>
                      <tr>
                        <td style={{ ...S.td, color: C.purple, fontWeight: "bold" }}>Tap Can</td>
                        <td style={S.td}>—</td><td style={S.td}>—</td>
                        <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>Fixed at start</td>
                        <td style={{ ...S.td, color: C.purple, fontWeight: "bold" }}>{tapCanIn}"</td>
                      </tr>
                      <tr>
                        <td style={{ ...S.td, color: C.muted }}>Gap</td>
                        <td style={S.td}>—</td><td style={S.td}>—</td>
                        <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>Tap can → first disconnect</td>
                        <td style={{ ...S.td, color: C.muted, fontWeight: "bold" }}>{POST_TAP_GAP}"</td>
                      </tr>
                      {bw.tenantSlots.map((t, i) => {
                        const sl = t.slot;
                        const formula = sl.type === "Stacked"
                          ? `${sl.disconnectIn}" (stacked) + ${sl.gapIn}" gap`
                          : `${sl.disconnectIn}" disc + ${sl.divIn}" + ${sl.meterIn}" meter + ${sl.gapIn}" gap`;
                        return (
                          <tr key={i}>
                            <td style={S.td}>{t.name || <em style={{ color: C.muted }}>(unnamed)</em>}</td>
                            <td style={S.td}><span style={S.pill()}>{t.serviceSize}A</span></td>
                            <td style={S.td}><span style={slotPill(sl)}>{sl.type}</span></td>
                            <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>{formula}</td>
                            <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{sl.totalIn}"</td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: C.raised, borderTop: `2px solid ${C.border}` }}>
                        <td colSpan={3} style={{ ...S.td, color: C.muted, fontSize: 9, letterSpacing: 2 }}>TOTAL WALL SPACE NEEDED</td>
                        <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>{tapCanIn}" (tap) + {POST_TAP_GAP}" (gap) + {bw.tenantsIn}" (tenants)</td>
                        <td style={{
                          ...S.td, fontWeight: "bold", fontSize: 13,
                          color: bw.availFt > 0 ? (bw.ok ? C.green : C.red) : C.amber
                        }}>{bw.totalIn}" = {bw.totalFt.toFixed(3)} ft</td>
                      </tr>
                      {bw.availFt > 0 && (
                        <tr style={{ background: C.raised }}>
                          <td colSpan={4} style={{ ...S.td, color: C.muted, fontSize: 9, letterSpacing: 2 }}>TOTAL AVAILABLE WALL SPACE</td>
                          <td style={{ ...S.td, color: C.muted, fontWeight: "bold" }}>{bw.availFt} ft</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
                <WirewayDiagram tenantSlots={bw.tenantSlots} totalIn={bw.totalIn}
                  totalFt={bw.totalFt} availFt={bw.availFt}
                  totalAmps={bw.tenantSlots.reduce((s, t) => s + (t.amps || 0), 0)} tapCanIn={tapCanIn} />
              </div>
            );
          })}
        </>)}

        {screen === "results" && (<>
          <div style={S.rg4(bp.mobile, bp.tablet)}>
            {[
              { v: project.name || "—", l: "Project" },
              { v: `${totalKVA.toFixed(2)} KVA`, l: "Total KVA" },
              { v: `${totalAmps.toFixed(2)} A`, l: `Total Amps ÷ ${voltage.factor}` },
              { v: project.buildings, l: "Buildings" },
            ].map((st, i) => (
              <div key={i} style={{ ...S.stat, marginBottom: 14 }} data-r-stat>
                <div style={S.sVal()} data-r-sval>{st.v}</div>
                <div style={S.sLbl}>{st.l}</div>
              </div>
            ))}
          </div>
          <div style={{ ...S.rg2(bp.mobile), marginBottom: 14 }}>
            <div style={{ ...S.stat, textAlign: "left", display: "flex", gap: 12, alignItems: "center", justifyContent: "center" }} data-r-stat>
              <span style={{ fontSize: 9, color: C.muted, letterSpacing: 2, textTransform: "uppercase" }}>Global Voltage:</span>
              <span style={{ color: C.amber, fontWeight: "bold", fontSize: 14 }}>{voltage.label}</span>
            </div>
          </div>

          {allBuildings.map((b) => {
            const bComp = computed.filter((t) => t.building === b);
            const bw = buildingWireway(b);
            return (
              <div key={b} style={S.card} data-r-card>
                <div style={S.cardH} data-r-cardh><span>BUILDING {b}</span></div>
                <div style={{ overflowX: "auto" }} data-r-tblwrap>
                  <table style={S.tbl}>
                    <thead><tr>{["Tenant", "Category", "S.F.", "W/ft²", "KVA", "Amps = KVA / factor", "Service", "Wireway"].map((h) => (
                      <th key={h} style={S.th}>{h}</th>
                    ))}</tr></thead>
                    <tbody>
                      {bComp.map((t) => (
                        <tr key={t.id}>
                          <td style={S.td}>{t.name || <em style={{ color: C.muted }}>(unnamed)</em>}</td>
                          <td style={{ ...S.td, color: C.muted }}>{t.category}</td>
                          <td style={S.td}>{t.isHP ? "—" : t.sqft || 0}</td>
                          <td style={{ ...S.td, color: C.muted }}>{t.isHP ? "fixed" : catMap[t.category]?.wattsPerSqft}</td>
                          <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{t.kva.toFixed(2)}</td>
                          <td style={{ ...S.td, color: C.blue, fontSize: 10 }}>
                            {t.kva.toFixed(2)} / {t.factor} = <strong>{t.amps.toFixed(4)}</strong>
                          </td>
                          <td style={S.td}><span style={S.pill()}>{t.serviceSize}A</span></td>
                          <td style={S.td}><span style={slotPill(t.slot)}>{t.slot.totalIn}"</span></td>
                        </tr>
                      ))}
                      <tr style={{ background: C.raised }}>
                        <td colSpan={4} style={{ ...S.td, color: C.muted, fontSize: 9, letterSpacing: 2 }}>SUBTOTAL B{b}</td>
                        <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{bComp.reduce((s, t) => s + t.kva, 0).toFixed(2)}</td>
                        <td style={{ ...S.td, color: C.blue, fontWeight: "bold" }}>{bComp.reduce((s, t) => s + t.amps, 0).toFixed(4)}</td>
                        <td style={S.td} />
                        <td style={{ ...S.td, color: bw.availFt > 0 ? (bw.ok ? C.green : C.red) : C.amber, fontWeight: "bold" }}>
                          {bw.totalIn}" = {bw.totalFt.toFixed(2)} ft {bw.availFt > 0 ? (bw.ok ? "✓" : "✗") : ""}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                {/* Wireway diagram */}
                <WirewayDiagram tenantSlots={bw.tenantSlots} totalIn={bw.totalIn}
                  totalFt={bw.totalFt} availFt={bw.availFt}
                  totalAmps={bw.tenantSlots.reduce((s, t) => s + (t.amps || 0), 0)} tapCanIn={tapCanIn} />
              </div>
            );
          })}

          <div style={S.card} data-r-card>
            <div style={S.cardH} data-r-cardh><span>SUMMARY BY CATEGORY</span></div>
            <div style={{ overflowX: "auto" }} data-r-tblwrap>
              <table style={{ ...S.tbl, maxWidth: 500 }}>
                <thead><tr>{["Category", "Total S.F.", "KVA"].map((h) => <th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {catSummary.map((c, i) => (
                    <tr key={i}>
                      <td style={{ ...S.td, color: "#94a3b8" }}>{c.name} S.F.:</td>
                      <td style={S.td}>{c.isHP ? "—" : c.totalSqft.toLocaleString()}</td>
                      <td style={{ ...S.td, color: C.amber, fontWeight: "bold" }}>{c.totalKVA.toFixed(2)}</td>
                    </tr>
                  ))}
                  <tr style={{ background: C.raised, borderTop: `2px solid ${C.border}` }}>
                    <td style={{ ...S.td, color: C.muted, fontSize: 9, letterSpacing: 2 }}>TOTAL KVA</td>
                    <td style={S.td} />
                    <td style={{ ...S.td, color: C.amber, fontWeight: "bold", fontSize: 16 }}>{totalKVA.toFixed(2)}</td>
                  </tr>
                  <tr style={{ background: C.raised }}>
                    <td style={{ ...S.td, color: C.muted, fontSize: 9, letterSpacing: 2 }}>TOTAL AMPS</td>
                    <td style={{ ...S.td, color: C.muted, fontSize: 10 }}>{totalKVA.toFixed(2)} ÷ {voltage.factor}</td>
                    <td style={{ ...S.td, color: C.blue, fontWeight: "bold", fontSize: 16 }}>{totalAmps.toFixed(6)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 4 }} data-print-hide>
            <button style={S.btn()} onClick={handleSave}>💾 SAVE .JSON</button>
            <button style={S.btn(C.blue, "#fff")} onClick={handleExcel}>📊 EXPORT EXCEL</button>
            <button style={S.btn(C.green, "#000")} onClick={handlePrint}>🖨 PRINT / PDF</button>
          </div>
        </>)}
      </div>
    </div>
  );
}

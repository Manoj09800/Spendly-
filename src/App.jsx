import { useState, useEffect, useMemo, useRef } from "react";
import {
  Home, PlusCircle, PieChart as PieIcon, Wallet, Utensils, Car, ShoppingBag,
  Zap, Heart, Film, MoreHorizontal, Briefcase, Gift, TrendingUp, PiggyBank,
  Trash2, ArrowUpRight, ArrowDownRight, X, UploadCloud, Loader2, FileStack
} from "lucide-react";
import {
  PieChart, Pie, Cell, ResponsiveContainer, BarChart, Bar, XAxis, YAxis,
  Tooltip, CartesianGrid
} from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";

const COLORS = {
  void: "#0A0F0D",
  panel: "#121913",
  panelRaised: "#1A231C",
  line: "#263129",
  text: "#EDEFEA",
  muted: "#8B9C90",
  jade: "#2BAE81",
  jadeDim: "#1D7A5B",
  gold: "#E3A63E",
  coral: "#E8604C",
  sky: "#5B93D6",
};

const EXPENSE_CATS = [
  { id: "food", label: "Food", icon: Utensils },
  { id: "transport", label: "Transport", icon: Car },
  { id: "shopping", label: "Shopping", icon: ShoppingBag },
  { id: "bills", label: "Bills", icon: Zap },
  { id: "health", label: "Health", icon: Heart },
  { id: "fun", label: "Entertainment", icon: Film },
  { id: "other_exp", label: "Other", icon: MoreHorizontal },
];
const INCOME_CATS = [
  { id: "salary", label: "Salary", icon: Briefcase },
  { id: "gift", label: "Gift", icon: Gift },
  { id: "freelance", label: "Freelance", icon: TrendingUp },
  { id: "other_inc", label: "Other", icon: MoreHorizontal },
];
const SAVING_CATS = [{ id: "savings", label: "Savings", icon: PiggyBank }];

const CAT_LOOKUP = [...EXPENSE_CATS, ...INCOME_CATS, ...SAVING_CATS].reduce(
  (acc, c) => ({ ...acc, [c.id]: c }),
  {}
);

const PIE_COLORS = [COLORS.jade, COLORS.gold, COLORS.coral, COLORS.sky, "#8B7FD6", "#4FB8C4", COLORS.muted];

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function fmt(n) {
  const sign = n < 0 ? "-" : "";
  return sign + Math.abs(n).toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function monthKey(dateStr) {
  const d = new Date(dateStr);
  return d.toLocaleString("en-US", { month: "short", year: "numeric" });
}

function monthSortValue(dateStr) {
  const d = new Date(dateStr);
  return d.getFullYear() * 12 + d.getMonth();
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve();
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Failed to load " + src));
    document.head.appendChild(s);
  });
}

async function ensurePdfJs() {
  if (window.pdfjsLib) return window.pdfjsLib;
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.min.js");
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.js";
  return window.pdfjsLib;
}

async function ensureTesseract() {
  if (window.Tesseract) return window.Tesseract;
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/5.0.4/tesseract.min.js");
  return window.Tesseract;
}

function parseDateGuess(s) {
  const parts = s.split(/[\/\-.]/);
  if (parts.length === 3) {
    let [a, b, c] = parts.map((p) => parseInt(p, 10));
    if (c < 100) c += 2000;
    const d = new Date(c, b - 1, a);
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  return new Date().toISOString();
}

// Parses a spreadsheet already split into rows-of-cells (e.g. from
// XLSX.utils.sheet_to_json with header:1). Looks for a header row naming
// its columns (Date, Debit, Credit, Amount, Balance, Details...) and reads
// each transaction directly from the right column — far more reliable
// than guessing from flattened text, and handles the common bank-statement
// layout of separate Debit/Credit columns.
function parseStatementRows(grid) {
  if (!grid || grid.length === 0) return null;

  const norm = (c) => String(c ?? "").trim().toLowerCase();
  let headerIdx = -1;
  let headers = [];
  for (let i = 0; i < Math.min(grid.length, 15); i++) {
    const row = (grid[i] || []).map(norm);
    const hasDate = row.some((c) => c.includes("date"));
    const hasAmountish = row.some(
      (c) => c.includes("debit") || c.includes("credit") || c.includes("amount") || c.includes("withdrawal") || c.includes("deposit")
    );
    if (hasDate && hasAmountish) {
      headerIdx = i;
      headers = row;
      break;
    }
  }
  if (headerIdx === -1) return null;

  const findCol = (patterns) => headers.findIndex((h) => patterns.some((p) => h.includes(p)));
  const dateCol = findCol(["value date", "date"]);
  const debitCol = findCol(["debit", "withdrawal"]);
  const creditCol = findCol(["credit", "deposit"]);
  const amountCol = findCol(["amount"]);
  const descCol = findCol(["detail", "description", "narration", "particular"]);

  if (dateCol === -1 || (debitCol === -1 && creditCol === -1 && amountCol === -1)) return null;

  const toNumber = (v) => {
    if (v === "" || v == null) return NaN;
    if (typeof v === "number") return v;
    return parseFloat(String(v).replace(/[^\d.-]/g, ""));
  };

  const results = [];
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const row = grid[i] || [];
    const rawDate = row[dateCol];
    if (rawDate === "" || rawDate == null) continue;

    let dateVal;
    if (rawDate instanceof Date) dateVal = rawDate.toISOString();
    else dateVal = parseDateGuess(String(rawDate));

    let amount = 0;
    let type = "expense";
    const debitVal = debitCol >= 0 ? toNumber(row[debitCol]) : NaN;
    const creditVal = creditCol >= 0 ? toNumber(row[creditCol]) : NaN;

    if (!isNaN(debitVal) && debitVal > 0) {
      amount = debitVal;
      type = "expense";
    } else if (!isNaN(creditVal) && creditVal > 0) {
      amount = creditVal;
      type = "income";
    } else if (amountCol >= 0) {
      const raw = toNumber(row[amountCol]);
      if (!isNaN(raw) && raw !== 0) {
        amount = Math.abs(raw);
        type = raw < 0 ? "expense" : "income";
      }
    }
    if (!amount || amount <= 0) continue;

    const desc = descCol >= 0 ? String(row[descCol] ?? "").slice(0, 60) : "Statement entry";
    results.push({ date: dateVal, description: desc || "Statement entry", amount, type });
  }
  return results;
}

// Groups a PDF page's text items into visual rows by their y-position,
// rather than trusting text-flow line breaks — table cells often don't
// carry reliable end-of-line flags, but their vertical position is exact.
function pageTextToRows(content) {
  const tolerance = 2.5;
  const rows = []; // { y, items: [{x, str}] }
  for (const it of content.items) {
    if (!it.str || !it.str.trim()) continue;
    const y = it.transform[5];
    const x = it.transform[4];
    let row = rows.find((r) => Math.abs(r.y - y) <= tolerance);
    if (!row) {
      row = { y, items: [] };
      rows.push(row);
    }
    row.items.push({ x, str: it.str });
  }
  rows.sort((a, b) => b.y - a.y); // top of page first
  return rows
    .map((r) =>
      r.items
        .sort((a, b) => a.x - b.x)
        .map((i) => i.str)
        .join(" ")
    )
    .join("\n");
}

function parseStatementText(text) {
  const dateRe = /(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})/;
  const amtRe = /(?:₹|rs\.?|inr)?\s?-?[\d,]+\.\d{2}\b/gi;
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const results = [];

  for (const line of lines) {
    const dateMatch = line.match(dateRe);
    if (!dateMatch) continue;
    const amounts = line.match(amtRe);
    if (!amounts || amounts.length === 0) continue;

    const clean = (s) => parseFloat(s.replace(/[^\d.-]/g, ""));
    const amount = amounts.length >= 2 ? clean(amounts[amounts.length - 2]) : clean(amounts[0]);
    if (!amount || isNaN(amount) || amount <= 0) continue;

    const lower = line.toLowerCase();
    // Common Indian bank-statement shorthand: WDL/DR = debit (expense),
    // DEP/CR = credit (income). Checked as priority signals first, since
    // generic words like "credit" can appear inside reference numbers.
    let isIncome;
    if (/\bwdl\b|\bdr\b|withdrawal|debited/.test(lower)) {
      isIncome = false;
    } else if (/\bdep\b|\bcr\b|credit|deposit|received|salary|refund/.test(lower)) {
      isIncome = true;
    } else {
      isIncome = false;
    }

    let desc = line.replace(dateMatch[0], "");
    amounts.forEach((a) => (desc = desc.replace(a, "")));
    desc = desc.replace(/[-|]+/g, " ").replace(/\s{2,}/g, " ").trim().slice(0, 60);

    results.push({
      date: parseDateGuess(dateMatch[0]),
      description: desc || "Statement entry",
      amount,
      type: isIncome ? "income" : "expense",
    });
  }
  return results;
}

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = checking, null = signed out

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <Shell>
        <div className="flex-1 flex items-center justify-center" style={{ color: COLORS.muted }}>
          Loading…
        </div>
      </Shell>
    );
  }

  if (!session) return <AuthView />;

  return <MoneyManager session={session} />;
}

function AuthView() {
  const [mode, setMode] = useState("signin"); // signin | signup
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit() {
    setError(""); setInfo(""); setBusy(true);
    try {
      if (mode === "signup") {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setInfo("Account created — check your email to confirm, then sign in.");
        setMode("signin");
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      }
    } catch (e) {
      setError(e.message || "Something went wrong.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <Shell>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');
        .f-display { font-family: 'Fraunces', serif; }
        .f-body { font-family: 'Inter', sans-serif; }
        .f-mono { font-family: 'JetBrains Mono', monospace; }
      `}</style>
      <div className="flex-1 flex flex-col justify-center px-6">
        <div className="f-display text-3xl font-semibold mb-1">Spendly</div>
        <div className="f-body text-sm mb-8" style={{ color: COLORS.muted }}>
          {mode === "signup" ? "Create an account to sync across your devices" : "Sign in to your account"}
        </div>

        <label className="f-body text-xs mb-1 block" style={{ color: COLORS.muted }}>Email</label>
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          className="w-full rounded-xl px-3 py-2.5 mb-3 outline-none text-sm"
          style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
        />

        <label className="f-body text-xs mb-1 block" style={{ color: COLORS.muted }}>Password</label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          className="w-full rounded-xl px-3 py-2.5 mb-5 outline-none text-sm"
          style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
        />

        {error && <div className="f-body text-sm mb-4" style={{ color: COLORS.coral }}>{error}</div>}
        {info && <div className="f-body text-sm mb-4" style={{ color: COLORS.jade }}>{info}</div>}

        <button
          onClick={submit}
          disabled={busy || !email || !password}
          className="w-full py-3 rounded-xl font-medium text-sm disabled:opacity-40 mb-4"
          style={{ background: COLORS.jade, color: COLORS.void }}
        >
          {busy ? "Please wait…" : mode === "signup" ? "Create account" : "Sign in"}
        </button>

        <button
          onClick={() => { setMode(mode === "signup" ? "signin" : "signup"); setError(""); setInfo(""); }}
          className="f-body text-sm"
          style={{ color: COLORS.muted }}
        >
          {mode === "signup" ? "Already have an account? Sign in" : "New here? Create an account"}
        </button>
      </div>
    </Shell>
  );
}

function MoneyManager({ session }) {
  const user = session.user;
  const [tab, setTab] = useState("home");
  const [txns, setTxns] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      const { data, error } = await supabase
        .from("transactions")
        .select("*")
        .eq("user_id", user.id)
        .order("date", { ascending: false });
      if (!error && data) {
        setTxns(
          data.map((r) => ({
            id: r.id,
            type: r.type,
            amount: Number(r.amount),
            category: r.category,
            note: r.note || "",
            date: r.date,
          }))
        );
      } else if (error) {
        showToast("Couldn't load your data");
      }
      setLoading(false);
    })();
  }, [user.id]);

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  }

  async function addTxn(t) {
    const { data, error } = await supabase
      .from("transactions")
      .insert({ ...t, user_id: user.id })
      .select()
      .single();
    if (error) { showToast("Couldn't save — try again"); return; }
    setTxns((prev) => [{ ...t, id: data.id }, ...prev]);
    setShowAdd(false);
    setTab("home");
    showToast("Added");
  }

  async function deleteTxn(id) {
    const { error } = await supabase.from("transactions").delete().eq("id", id).eq("user_id", user.id);
    if (error) { showToast("Couldn't delete — try again"); return; }
    setTxns((prev) => prev.filter((t) => t.id !== id));
    showToast("Deleted");
  }

  async function addManyTxns(list) {
    const rows = list.map((t) => ({ ...t, user_id: user.id }));
    const { data, error } = await supabase.from("transactions").insert(rows).select();
    if (error) { showToast("Couldn't save — try again"); return; }
    setTxns((prev) => [...data.map((r) => ({ id: r.id, type: r.type, amount: Number(r.amount), category: r.category, note: r.note || "", date: r.date })), ...prev]);
    setTab("home");
    showToast(`Added ${data.length} entries`);
  }

  const totals = useMemo(() => {
    let income = 0, expense = 0, savings = 0;
    for (const t of txns) {
      if (t.type === "income") income += t.amount;
      else if (t.type === "expense") expense += t.amount;
      else if (t.type === "saving") savings += t.amount;
    }
    return { income, expense, savings, balance: income - expense - savings };
  }, [txns]);

  const sorted = useMemo(
    () => [...txns].sort((a, b) => new Date(b.date) - new Date(a.date)),
    [txns]
  );

  if (loading) {
    return (
      <Shell>
        <div className="flex-1 flex items-center justify-center" style={{ color: COLORS.muted }}>
          Loading…
        </div>
      </Shell>
    );
  }

  return (
    <Shell>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');
        .f-display { font-family: 'Fraunces', serif; }
        .f-body { font-family: 'Inter', sans-serif; }
        .f-mono { font-family: 'JetBrains Mono', monospace; }
        .no-scrollbar::-webkit-scrollbar { display: none; }
      `}</style>

      <div className="flex items-center justify-between px-5 pt-4">
        <span className="f-body text-xs truncate" style={{ color: COLORS.muted }}>{user.email}</span>
        <button
          onClick={() => supabase.auth.signOut()}
          className="f-body text-xs"
          style={{ color: COLORS.muted }}
        >
          Sign out
        </button>
      </div>

      <div className="flex-1 overflow-y-auto no-scrollbar pb-24">
        {tab === "home" && (
          <HomeView totals={totals} txns={sorted.slice(0, 6)} onDelete={deleteTxn} onSeeAll={() => setTab("reports")} />
        )}
        {tab === "reports" && <ReportsView txns={sorted} onDelete={deleteTxn} />}
        {tab === "statement" && <StatementView onBulkAdd={addManyTxns} />}
      </div>

      {toast && (
        <div
          className="f-body absolute bottom-24 left-1/2 -translate-x-1/2 px-4 py-2 rounded-full text-sm shadow-lg"
          style={{ background: COLORS.panelRaised, color: COLORS.text, border: `1px solid ${COLORS.line}` }}
        >
          {toast}
        </div>
      )}

      <NavBar tab={tab} setTab={setTab} onAdd={() => setShowAdd(true)} />

      {showAdd && <AddSheet onClose={() => setShowAdd(false)} onSave={addTxn} />}
    </Shell>
  );
}

function Shell({ children }) {
  return (
    <div
      className="relative w-full mx-auto flex flex-col"
      style={{
        maxWidth: 430,
        height: "100vh",
        minHeight: 640,
        background: COLORS.void,
        color: COLORS.text,
      }}
    >
      {children}
    </div>
  );
}

function HomeView({ totals, txns, onDelete, onSeeAll }) {
  return (
    <div className="px-5 pt-8">
      <div className="f-body text-sm mb-1" style={{ color: COLORS.muted }}>
        Your balance
      </div>
      <div className="flex items-baseline gap-2 mb-6">
        <span className="f-display text-5xl font-semibold">₹{fmt(totals.balance)}</span>
      </div>

      <div className="grid grid-cols-3 gap-3 mb-8">
        <StatCard label="Income" value={totals.income} color={COLORS.sky} Icon={ArrowUpRight} />
        <StatCard label="Expense" value={totals.expense} color={COLORS.coral} Icon={ArrowDownRight} />
        <StatCard label="Savings" value={totals.savings} color={COLORS.gold} Icon={PiggyBank} />
      </div>

      <div className="flex items-center justify-between mb-3">
        <span className="f-body text-sm font-medium" style={{ color: COLORS.muted }}>
          Recent entries
        </span>
        {txns.length > 0 && (
          <button onClick={onSeeAll} className="f-body text-sm" style={{ color: COLORS.jade }}>
            See all
          </button>
        )}
      </div>

      {txns.length === 0 ? (
        <EmptyState />
      ) : (
        <Ledger txns={txns} onDelete={onDelete} />
      )}
    </div>
  );
}

function StatCard({ label, value, color, Icon }) {
  return (
    <div
      className="rounded-2xl p-3 flex flex-col gap-2"
      style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}
    >
      <div className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: color + "22" }}>
        <Icon size={14} color={color} />
      </div>
      <div>
        <div className="f-body text-xs" style={{ color: COLORS.muted }}>{label}</div>
        <div className="f-mono text-sm font-medium">₹{fmt(value)}</div>
      </div>
    </div>
  );
}

function EmptyState() {
  return (
    <div
      className="f-body rounded-2xl p-8 text-center text-sm"
      style={{ background: COLORS.panel, border: `1px dashed ${COLORS.line}`, color: COLORS.muted }}
    >
      Nothing logged yet. Tap the + below to add your first entry.
    </div>
  );
}

function Ledger({ txns, onDelete }) {
  return (
    <div className="relative">
      <div className="absolute left-[15px] top-1 bottom-1 w-px" style={{ background: COLORS.line }} />
      <div className="flex flex-col">
        {txns.map((t) => {
          const cat = CAT_LOOKUP[t.category] || { label: t.category, icon: MoreHorizontal };
          const Icon = cat.icon;
          const color = t.type === "income" ? COLORS.sky : t.type === "saving" ? COLORS.gold : COLORS.coral;
          const sign = t.type === "income" ? "+" : "-";
          return (
            <div key={t.id} className="relative flex items-center gap-3 py-2.5 group">
              <div
                className="relative z-10 w-8 h-8 rounded-full flex items-center justify-center shrink-0"
                style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}` }}
              >
                <Icon size={14} color={color} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="f-body text-sm truncate">{t.note || cat.label}</div>
                <div className="f-mono text-[11px]" style={{ color: COLORS.muted }}>
                  {new Date(t.date).toLocaleDateString("en-IN", { day: "numeric", month: "short" })} · {cat.label}
                </div>
              </div>
              <div className="f-mono text-sm font-medium" style={{ color }}>
                {sign}₹{fmt(t.amount)}
              </div>
              <button
                onClick={() => onDelete(t.id)}
                className="opacity-0 group-hover:opacity-100 transition-opacity"
                style={{ color: COLORS.muted }}
              >
                <Trash2 size={13} />
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ReportsView({ txns, onDelete }) {
  const [selectedMonth, setSelectedMonth] = useState("all");

  const availableMonths = useMemo(() => {
    const map = new Map();
    for (const t of txns) {
      const key = monthKey(t.date);
      if (!map.has(key)) map.set(key, monthSortValue(t.date));
    }
    return Array.from(map.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([key]) => key);
  }, [txns]);

  const monthly = useMemo(() => {
    const map = {};
    for (const t of txns) {
      const k = monthKey(t.date);
      if (!map[k]) map[k] = { month: k, income: 0, expense: 0, sort: monthSortValue(t.date) };
      if (t.type === "income") map[k].income += t.amount;
      if (t.type === "expense") map[k].expense += t.amount;
    }
    return Object.values(map).sort((a, b) => a.sort - b.sort).slice(-12);
  }, [txns]);

  const highestSpendMonth = useMemo(() => {
    if (monthly.length === 0) return null;
    return monthly.reduce((max, m) => (m.expense > (max?.expense ?? -1) ? m : max), null);
  }, [monthly]);

  const activeMonth = selectedMonth === "all" ? availableMonths[0] : selectedMonth;

  const filteredTxns = useMemo(() => {
    if (selectedMonth === "all") return txns;
    return txns.filter((t) => monthKey(t.date) === selectedMonth);
  }, [txns, selectedMonth]);

  const recap = useMemo(() => {
    if (!activeMonth) return null;
    const inMonth = txns.filter((t) => monthKey(t.date) === activeMonth);
    const spent = inMonth.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

    const catMap = {};
    for (const t of inMonth) {
      if (t.type !== "expense") continue;
      catMap[t.category] = (catMap[t.category] || 0) + t.amount;
    }
    const topCats = Object.entries(catMap)
      .map(([id, value]) => ({ name: CAT_LOOKUP[id]?.label || id, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 3);

    const idx = availableMonths.indexOf(activeMonth);
    const prevMonthKey = idx >= 0 ? availableMonths[idx + 1] : null;
    let change = null;
    if (prevMonthKey) {
      const prevSpent = txns
        .filter((t) => monthKey(t.date) === prevMonthKey && t.type === "expense")
        .reduce((s, t) => s + t.amount, 0);
      if (prevSpent > 0) change = ((spent - prevSpent) / prevSpent) * 100;
    }

    return { month: activeMonth, spent, topCats, change };
  }, [txns, activeMonth, availableMonths]);

  const byCategory = useMemo(() => {
    const map = {};
    for (const t of filteredTxns) {
      if (t.type !== "expense") continue;
      map[t.category] = (map[t.category] || 0) + t.amount;
    }
    return Object.entries(map)
      .map(([id, value]) => ({ name: CAT_LOOKUP[id]?.label || id, value }))
      .sort((a, b) => b.value - a.value);
  }, [filteredTxns]);

  return (
    <div className="px-5 pt-8">
      <div className="flex items-center justify-between mb-6">
        <div className="f-display text-2xl font-semibold">Reports</div>
        {availableMonths.length > 0 && (
          <select
            value={selectedMonth}
            onChange={(e) => setSelectedMonth(e.target.value)}
            className="f-body text-xs rounded-lg px-2 py-1.5"
            style={{ background: COLORS.panelRaised, color: COLORS.text, border: `1px solid ${COLORS.line}` }}
          >
            <option value="all">All time</option>
            {availableMonths.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        )}
      </div>

      {txns.length === 0 ? (
        <EmptyState />
      ) : (
        <>
          {recap && (
            <div className="rounded-2xl p-4 mb-5" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
              <div className="f-body text-sm mb-1" style={{ color: COLORS.muted }}>
                {recap.month} recap
              </div>
              <div className="f-display text-2xl font-semibold mb-1">₹{fmt(recap.spent)}</div>
              {recap.change !== null && (
                <div
                  className="f-body text-xs mb-3"
                  style={{ color: recap.change > 0 ? COLORS.coral : COLORS.jade }}
                >
                  {recap.change > 0 ? "▲" : "▼"} {Math.abs(recap.change).toFixed(0)}% vs previous month
                </div>
              )}
              {recap.topCats.length > 0 && (
                <div className="flex flex-col gap-1.5 mt-2">
                  {recap.topCats.map((c) => (
                    <div key={c.name} className="f-body flex items-center justify-between text-sm">
                      <span style={{ color: COLORS.muted }}>{c.name}</span>
                      <span className="f-mono">₹{fmt(c.value)}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {highestSpendMonth && monthly.length > 1 && (
            <div
              className="rounded-2xl p-4 mb-5 flex items-center gap-3"
              style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}
            >
              <div className="w-8 h-8 rounded-full flex items-center justify-center shrink-0" style={{ background: COLORS.gold + "22" }}>
                <TrendingUp size={15} color={COLORS.gold} />
              </div>
              <div className="f-body text-sm">
                <span style={{ color: COLORS.muted }}>Highest spending month: </span>
                <span>{highestSpendMonth.month} (₹{fmt(highestSpendMonth.expense)})</span>
              </div>
            </div>
          )}

          {byCategory.length > 0 && (
            <div className="rounded-2xl p-4 mb-5" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
              <div className="f-body text-sm mb-2" style={{ color: COLORS.muted }}>
                Spending by category {selectedMonth !== "all" ? `— ${selectedMonth}` : ""}
              </div>
              <div style={{ width: "100%", height: 200 }}>
                <ResponsiveContainer>
                  <PieChart>
                    <Pie data={byCategory} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} paddingAngle={3}>
                      {byCategory.map((_, i) => (
                        <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} stroke="none" />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, borderRadius: 8, fontSize: 12 }}
                      formatter={(v) => [`₹${fmt(v)}`, ""]}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="flex flex-wrap gap-x-4 gap-y-1.5 mt-2">
                {byCategory.map((c, i) => (
                  <div key={c.name} className="f-body flex items-center gap-1.5 text-xs" style={{ color: COLORS.muted }}>
                    <span className="w-2 h-2 rounded-full" style={{ background: PIE_COLORS[i % PIE_COLORS.length] }} />
                    {c.name}
                  </div>
                ))}
              </div>
            </div>
          )}

          {monthly.length > 1 && (
            <div className="rounded-2xl p-4 mb-5" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
              <div className="f-body text-sm mb-2" style={{ color: COLORS.muted }}>Income vs expense</div>
              <div style={{ width: "100%", height: 180 }}>
                <ResponsiveContainer>
                  <BarChart data={monthly}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} vertical={false} />
                    <XAxis dataKey="month" stroke={COLORS.muted} fontSize={11} tickLine={false} axisLine={false} />
                    <YAxis stroke={COLORS.muted} fontSize={11} tickLine={false} axisLine={false} width={30} />
                    <Tooltip
                      contentStyle={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, borderRadius: 8, fontSize: 12 }}
                      formatter={(v) => [`₹${fmt(v)}`, ""]}
                    />
                    <Bar dataKey="income" fill={COLORS.sky} radius={[3, 3, 0, 0]} />
                    <Bar dataKey="expense" fill={COLORS.coral} radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          <div className="f-body text-sm mb-3" style={{ color: COLORS.muted }}>
            {selectedMonth === "all" ? "All entries" : `${selectedMonth} entries`}
          </div>
          <Ledger txns={filteredTxns} onDelete={onDelete} />
        </>
      )}
    </div>
  );
}

function NavBar({ tab, setTab, onAdd }) {
  return (
    <div
      className="absolute bottom-0 left-0 right-0 flex items-center justify-around py-3"
      style={{ background: COLORS.panel, borderTop: `1px solid ${COLORS.line}` }}
    >
      <NavBtn active={tab === "home"} onClick={() => setTab("home")} Icon={Home} label="Home" />
      <NavBtn active={tab === "statement"} onClick={() => setTab("statement")} Icon={FileStack} label="Statement" />
      <button onClick={onAdd} className="flex flex-col items-center gap-1">
        <div className="w-10 h-10 rounded-full flex items-center justify-center" style={{ background: COLORS.jade }}>
          <PlusCircle size={20} color={COLORS.void} />
        </div>
      </button>
      <NavBtn active={tab === "reports"} onClick={() => setTab("reports")} Icon={PieIcon} label="Reports" />
    </div>
  );
}

function NavBtn({ active, onClick, Icon, label }) {
  return (
    <button onClick={onClick} className="f-body flex flex-col items-center gap-1 text-[11px]" style={{ color: active ? COLORS.jade : COLORS.muted }}>
      <Icon size={20} />
      {label}
    </button>
  );
}

function AddSheet({ onClose, onSave }) {
  const [type, setType] = useState("expense");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState(EXPENSE_CATS[0].id);
  const [note, setNote] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));

  const cats = type === "expense" ? EXPENSE_CATS : type === "income" ? INCOME_CATS : SAVING_CATS;

  function switchType(t) {
    setType(t);
    const list = t === "expense" ? EXPENSE_CATS : t === "income" ? INCOME_CATS : SAVING_CATS;
    setCategory(list[0].id);
  }

  function submit() {
    const n = parseFloat(amount);
    if (!n || n <= 0) return;
    onSave({ type, amount: n, category, note: note.trim(), date: new Date(date).toISOString() });
  }

  return (
    <div className="absolute inset-0 z-20 flex items-end" style={{ background: "#000000aa" }} onClick={onClose}>
      <div
        className="f-body w-full rounded-t-3xl p-5 pb-8"
        style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}`, maxWidth: 430, margin: "0 auto" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-5">
          <span className="f-display text-lg font-semibold">Add entry</span>
          <button onClick={onClose} style={{ color: COLORS.muted }}><X size={18} /></button>
        </div>

        <div className="flex gap-2 mb-5">
          {[
            { id: "expense", label: "Expense" },
            { id: "income", label: "Income" },
            { id: "saving", label: "Savings" },
          ].map((o) => (
            <button
              key={o.id}
              onClick={() => switchType(o.id)}
              className="flex-1 py-2 rounded-xl text-sm font-medium"
              style={{
                background: type === o.id ? COLORS.jade : COLORS.panelRaised,
                color: type === o.id ? COLORS.void : COLORS.muted,
              }}
            >
              {o.label}
            </button>
          ))}
        </div>

        <label className="text-xs mb-1 block" style={{ color: COLORS.muted }}>Amount</label>
        <div className="flex items-center rounded-xl px-3 mb-4" style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}` }}>
          <span className="f-mono mr-1" style={{ color: COLORS.muted }}>₹</span>
          <input
            autoFocus
            inputMode="decimal"
            value={amount}
            onChange={(e) => setAmount(e.target.value.replace(/[^0-9.]/g, ""))}
            placeholder="0"
            className="f-mono w-full bg-transparent py-3 outline-none text-lg"
            style={{ color: COLORS.text }}
          />
        </div>

        <label className="text-xs mb-2 block" style={{ color: COLORS.muted }}>Category</label>
        <div className="grid grid-cols-4 gap-2 mb-4">
          {cats.map((c) => {
            const Icon = c.icon;
            const active = category === c.id;
            return (
              <button
                key={c.id}
                onClick={() => setCategory(c.id)}
                className="flex flex-col items-center gap-1 py-2 rounded-xl text-[11px]"
                style={{
                  background: active ? COLORS.jade + "22" : "transparent",
                  border: `1px solid ${active ? COLORS.jade : COLORS.line}`,
                  color: active ? COLORS.jade : COLORS.muted,
                }}
              >
                <Icon size={16} />
                {c.label}
              </button>
            );
          })}
        </div>

        <label className="text-xs mb-1 block" style={{ color: COLORS.muted }}>Note (optional)</label>
        <input
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="e.g. Swiggy lunch"
          className="w-full rounded-xl px-3 py-2.5 mb-4 outline-none text-sm"
          style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
        />

        <label className="text-xs mb-1 block" style={{ color: COLORS.muted }}>Date</label>
        <input
          type="date"
          value={date}
          onChange={(e) => setDate(e.target.value)}
          className="f-mono w-full rounded-xl px-3 py-2.5 mb-6 outline-none text-sm"
          style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
        />

        <button
          onClick={submit}
          disabled={!amount || parseFloat(amount) <= 0}
          className="w-full py-3 rounded-xl font-medium text-sm disabled:opacity-40"
          style={{ background: COLORS.jade, color: COLORS.void }}
        >
          Save entry
        </button>
      </div>
    </div>
  );
}

function StatementView({ onBulkAdd }) {
  const [mode, setMode] = useState("pdf");
  const [password, setPassword] = useState("");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [rows, setRows] = useState([]);
  const [pasteText, setPasteText] = useState("");

  function finishParse(parsed) {
    setStatus("");
    if (parsed.length === 0) {
      setError("Couldn't find any transactions automatically. You can still add them manually from the + button.");
    }
    setRows(
      parsed.map((p) => ({
        ...p,
        id: uid(),
        include: true,
        category: p.type === "income" ? "other_inc" : "other_exp",
      }))
    );
  }

  async function handlePdfFile(file) {
    setError(""); setRows([]); setStatus("Reading PDF…");
    try {
      const pdfjsLib = await ensurePdfJs();
      const buf = await file.arrayBuffer();
      let doc;
      try {
        doc = await pdfjsLib.getDocument({ data: buf, password: password || undefined }).promise;
      } catch (e) {
        if (e && e.name === "PasswordException") {
          setError("This PDF needs the correct password — enter it above and try again.");
          setStatus("");
          return;
        }
        throw e;
      }
      let fullText = "";
      for (let i = 1; i <= doc.numPages; i++) {
        setStatus(`Reading page ${i} of ${doc.numPages}…`);
        const page = await doc.getPage(i);
        const content = await page.getTextContent();
        fullText += pageTextToRows(content) + "\n";
      }
      finishParse(parseStatementText(fullText));
    } catch (e) {
      setError("Couldn't read that PDF. Try a different file, or add entries manually.");
      setStatus("");
    }
  }

  async function handleExcelFile(file) {
    setError(""); setRows([]); setStatus("Reading file…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      const byColumns = parseStatementRows(grid);
      if (byColumns && byColumns.length > 0) {
        finishParse(byColumns);
      } else {
        // Fall back to the generic text parser for files with no clear header row
        const csv = XLSX.utils.sheet_to_csv(sheet);
        finishParse(parseStatementText(csv));
      }
    } catch (e) {
      setError("Couldn't read that spreadsheet.");
      setStatus("");
    }
  }

  async function handleImageFile(file) {
    setError(""); setRows([]); setStatus("Scanning image — this can take a minute…");
    try {
      const Tesseract = await ensureTesseract();
      const { data } = await Tesseract.recognize(file, "eng");
      finishParse(parseStatementText(data.text));
    } catch (e) {
      setError("Couldn't scan that image.");
      setStatus("");
    }
  }

  function updateRow(id, patch) {
    setRows((rs) => rs.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  }

  function save() {
    const toAdd = rows
      .filter((r) => r.include && r.amount > 0)
      .map((r) => ({ type: r.type, amount: r.amount, category: r.category, note: r.description, date: r.date }));
    if (toAdd.length === 0) return;
    onBulkAdd(toAdd);
    setRows([]);
  }

  const included = rows.filter((r) => r.include).length;

  return (
    <div className="px-5 pt-8">
      <div className="f-display text-2xl font-semibold mb-1">Import statement</div>
      <div className="f-body text-sm mb-5" style={{ color: COLORS.muted }}>
        Upload a PDF, Excel/CSV, or a photo of your statement. Nothing is added until you review it below.
      </div>

      <div className="flex gap-2 mb-5 flex-wrap">
        {[["pdf", "PDF"], ["excel", "Excel / CSV"], ["image", "Photo"], ["paste", "Paste SMS/text"]].map(([id, label]) => (
          <button
            key={id}
            onClick={() => { setMode(id); setRows([]); setError(""); setStatus(""); }}
            className="flex-1 py-2 rounded-xl text-sm font-medium"
            style={{ background: mode === id ? COLORS.jade : COLORS.panelRaised, color: mode === id ? COLORS.void : COLORS.muted }}
          >
            {label}
          </button>
        ))}
      </div>

      {mode === "pdf" && (
        <div className="rounded-2xl p-4 mb-4" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
          <label className="text-xs mb-1 block" style={{ color: COLORS.muted }}>PDF password (if it has one)</label>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Leave blank if none"
            className="w-full rounded-xl px-3 py-2.5 mb-3 outline-none text-sm"
            style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
          />
          <FilePicker accept=".pdf" onFile={handlePdfFile} label="Choose PDF" />
        </div>
      )}
      {mode === "excel" && (
        <div className="rounded-2xl p-4 mb-4" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
          <FilePicker accept=".xlsx,.xls,.csv" onFile={handleExcelFile} label="Choose Excel or CSV file" />
        </div>
      )}
      {mode === "image" && (
        <div className="rounded-2xl p-4 mb-4" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
          <FilePicker accept="image/*" onFile={handleImageFile} label="Choose photo" />
          <div className="f-body text-xs mt-2" style={{ color: COLORS.muted }}>
            Photo scanning is the least reliable mode — double-check every row below.
          </div>
        </div>
      )}
      {mode === "paste" && (
        <div className="rounded-2xl p-4 mb-4" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
          <label className="text-xs mb-1 block" style={{ color: COLORS.muted }}>
            Paste bank SMS alerts or any statement text
          </label>
          <textarea
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            rows={6}
            placeholder={"e.g. 01-05-2026 Rs.332.00 debited from your a/c...\n(paste several messages at once — one per line works too)"}
            className="w-full rounded-xl px-3 py-2.5 mb-3 outline-none text-sm"
            style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}`, color: COLORS.text }}
          />
          <button
            onClick={() => finishParse(parseStatementText(pasteText))}
            disabled={!pasteText.trim()}
            className="w-full py-2.5 rounded-xl text-sm font-medium disabled:opacity-40"
            style={{ background: COLORS.jade, color: COLORS.void }}
          >
            Find transactions
          </button>
          <div className="f-body text-xs mt-2" style={{ color: COLORS.muted }}>
            No app can read your SMS inbox automatically for privacy/security reasons — but copying and
            pasting a batch here is fast and stays private, nothing leaves your browser.
          </div>
        </div>
      )}

      {status && (
        <div className="f-body text-sm mb-4 flex items-center gap-2" style={{ color: COLORS.gold }}>
          <Loader2 size={14} className="animate-spin" />
          {status}
        </div>
      )}
      {error && <div className="f-body text-sm mb-4" style={{ color: COLORS.coral }}>{error}</div>}

      {rows.length > 0 && (
        <>
          <div className="f-body text-sm mb-3" style={{ color: COLORS.muted }}>
            {rows.length} found — review and edit before adding
          </div>
          <div className="flex flex-col gap-2 mb-5">
            {rows.map((r) => (
              <ReviewRow key={r.id} row={r} onChange={(patch) => updateRow(r.id, patch)} />
            ))}
          </div>
          <button
            onClick={save}
            disabled={included === 0}
            className="w-full py-3 rounded-xl font-medium text-sm disabled:opacity-40 mb-8"
            style={{ background: COLORS.jade, color: COLORS.void }}
          >
            Add {included} to ledger
          </button>
        </>
      )}
    </div>
  );
}

function FilePicker({ accept, onFile, label }) {
  const inputRef = useRef(null);
  return (
    <div>
      <button
        type="button"
        onClick={() => inputRef.current && inputRef.current.click()}
        className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm"
        style={{ border: `1px dashed ${COLORS.line}`, color: COLORS.jade, background: "transparent" }}
      >
        <UploadCloud size={16} />
        {label}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        style={{
          position: "absolute",
          width: 1,
          height: 1,
          padding: 0,
          margin: -1,
          overflow: "hidden",
          clip: "rect(0,0,0,0)",
          whiteSpace: "nowrap",
          border: 0,
        }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onFile(f);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function ReviewRow({ row, onChange }) {
  const cats = row.type === "income" ? INCOME_CATS : EXPENSE_CATS;
  return (
    <div className="rounded-xl p-3" style={{ background: COLORS.panelRaised, border: `1px solid ${COLORS.line}` }}>
      <div className="flex items-center gap-2 mb-2">
        <input type="checkbox" checked={row.include} onChange={(e) => onChange({ include: e.target.checked })} />
        <input
          value={row.description}
          onChange={(e) => onChange({ description: e.target.value })}
          className="f-body flex-1 bg-transparent text-sm outline-none"
          style={{ color: COLORS.text }}
        />
      </div>
      <div className="flex gap-2 items-center">
        <select
          value={row.type}
          onChange={(e) => onChange({ type: e.target.value, category: e.target.value === "income" ? "other_inc" : "other_exp" })}
          className="f-body text-xs rounded-lg px-2 py-1.5"
          style={{ background: COLORS.panel, color: COLORS.text, border: `1px solid ${COLORS.line}` }}
        >
          <option value="expense">Expense</option>
          <option value="income">Income</option>
        </select>
        <select
          value={row.category}
          onChange={(e) => onChange({ category: e.target.value })}
          className="f-body text-xs rounded-lg px-2 py-1.5 flex-1"
          style={{ background: COLORS.panel, color: COLORS.text, border: `1px solid ${COLORS.line}` }}
        >
          {cats.map((c) => (
            <option key={c.id} value={c.id}>{c.label}</option>
          ))}
        </select>
        <div className="flex items-center f-mono text-sm">
          ₹
          <input
            value={row.amount}
            onChange={(e) => onChange({ amount: parseFloat(e.target.value) || 0 })}
            className="w-16 bg-transparent outline-none"
            style={{ color: COLORS.text }}
          />
        </div>
      </div>
    </div>
  );
                    }


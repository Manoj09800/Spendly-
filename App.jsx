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
  return d.toLocaleString("en-US", { month: "short" });
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

// Heuristic line parser: finds a date + one or two amounts per line.
// If two amounts are found, treats the last as a running balance and the
// second-to-last as the transaction amount (common bank-statement layout).
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
    const isIncome =
      /(credit|cr\b|deposit|received|salary|refund)/.test(lower) && !/(debit|dr\b)/.test(lower);

    let desc = line.replace(dateMatch[0], "");
    amounts.forEach((a) => (desc = desc.replace(a, "")));
    desc = desc.replace(/\s{2,}/g, " ").trim().slice(0, 60);

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
  const byCategory = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (t.type !== "expense") continue;
      map[t.category] = (map[t.category] || 0) + t.amount;
    }
    return Object.entries(map)
      .map(([id, value]) => ({ name: CAT_LOOKUP[id]?.label || id, value }))
      .sort((a, b) => b.value - a.value);
  }, [txns]);

  const monthly = useMemo(() => {
    const map = {};
    for (const t of txns) {
      const k = monthKey(t.date);
      if (!map[k]) map[k] = { month: k, income: 0, expense: 0 };
      if (t.type === "income") map[k].income += t.amount;
      if (t.type === "expense") map[k].expense += t.amount;
    }
    return Object.values(map).slice(-6);
  }, [txns]);

  return (
    <div className="px-5 pt-8">
      <div className="f-display text-2xl font-semibold mb-6">Reports</div>

      {txns.length === 0 ? (
        <EmptyState />
      ) : (
        <>
          {byCategory.length > 0 && (
            <div className="rounded-2xl p-4 mb-5" style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}` }}>
              <div className="f-body text-sm mb-2" style={{ color: COLORS.muted }}>Spending by category</div>
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

          <div className="f-body text-sm mb-3" style={{ color: COLORS.muted }}>All entries</div>
          <Ledger txns={txns} onDelete={onDelete} />
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
        fullText += content.items.map((it) => it.str).join(" ") + "\n";
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
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const csv = XLSX.utils.sheet_to_csv(sheet);
      finishParse(parseStatementText(csv.replace(/,/g, " ")));
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

      <div className="flex gap-2 mb-5">
        {[["pdf", "PDF"], ["excel", "Excel / CSV"], ["image", "Photo"]].map(([id, label]) => (
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

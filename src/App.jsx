import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  ArrowClockwise,
  ArrowRight,
  BellSimple,
  BookOpenText,
  CaretDown,
  Check,
  CheckCircle,
  ClockCounterClockwise,
  CloudArrowUp,
  Database,
  DownloadSimple,
  FileArrowUp,
  FileMagnifyingGlass,
  GlobeHemisphereWest,
  House,
  MagnifyingGlass,
  PencilSimpleLine,
  Plus,
  SealCheck,
  ShieldCheck,
  SpinnerGap,
  Tag,
  Trash,
  UploadSimple,
  UserCircle,
  WarningCircle,
  X,
  XCircle,
} from "@phosphor-icons/react";
import { api } from "./api";
import * as XLSX from "xlsx";
import { extractInquiryRows, isSupportedQuoteFile, prepareQuoteFile, quoteFileType } from "./quote-file-parser";
import { prepareKnowledgeFile } from "./knowledge-file-parser";

const NAV_ITEMS = [
  { id: "overview", label: "概览", icon: House },
  { id: "models", label: "型号查询", icon: MagnifyingGlass },
  { id: "quote", label: "批量报价", icon: FileMagnifyingGlass },
  { id: "reviews", label: "二次校对", icon: SealCheck },
  { id: "knowledge", label: "知识库", icon: Database },
];

const PAGE_META = {
  overview: ["工作概览", "统一查看型号、价格、校对和报价进度"],
  models: ["型号查询", "输入内部型号、OE 号或 UF 号，自动补齐关联信息"],
  quote: ["批量报价", "上传客户询价表，选择价格类型并导出报价"],
  reviews: ["二次校对", "审核外网补充关系，确认后再进入正式知识库"],
  knowledge: ["知识库", "查看版本并上传新的型号表、价格表或产品目录"],
};

function formatDate(value) {
  if (!value) return "—";
  return new Intl.DateTimeFormat("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" }).format(new Date(value));
}

function evidenceSummary(value = "") {
  const text = String(value)
    .replace(/!\[[^\]]*\]\([^)]*\)/g, " ")
    .replace(/\[([^\]]+)\]\(https?:\/\/[^)]+\)/g, "$1")
    .replace(/<[^>]+>/g, " ")
    .replace(/(?:Markdown Content|URL Source|Published Time|Title):/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
  return text.length > 220 ? `${text.slice(0, 220)}…` : text;
}

function useStoredState(key, initialValue) {
  const [value, setValue] = useState(() => {
    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : typeof initialValue === "function" ? initialValue() : initialValue;
    } catch { return typeof initialValue === "function" ? initialValue() : initialValue; }
  });
  useEffect(() => {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* 浏览器空间不足时保留当前会话 */ }
  }, [key, value]);
  return [value, setValue];
}

function splitModelQueries(value) {
  return [...new Set(String(value || "").split(/[\n,，;；]+/).map((item) => item.trim()).filter(Boolean))].slice(0, 50);
}

function normalizeTextInquiry(value) {
  return String(value || "").replace(/[，,;；\t]+/g, "\n").trim();
}

function isLegacyMisdetectedQuote(result) {
  const column = String(result?.detectedColumn || "").toUpperCase().replace(/[\s._/\\-]+/g, "");
  if (!/^(?:ITEMNO|ITEMNUMBER|NO|序号)$/.test(column)) return false;
  const inputs = (result?.lines || []).map((line) => String(line.input || "").trim()).filter(Boolean);
  return inputs.length > 0 && inputs.every((input) => /^\d{1,4}$/.test(input));
}

const HISTORY_TRASH_TTL = 7 * 24 * 60 * 60 * 1000;

function validTrashItems(items = []) {
  const now = Date.now();
  return items.filter((item) => Number(item.expiresAt) > now);
}

function remainingTrashDays(item) {
  return Math.max(1, Math.ceil((Number(item.expiresAt) - Date.now()) / (24 * 60 * 60 * 1000)));
}

async function extractModelQueriesFromFile(file, onProgress) {
  const prepared = await prepareQuoteFile(file, onProgress);
  const workbook = XLSX.read(await prepared.file.arrayBuffer(), { type: "array", raw: false });
  const text = workbook.SheetNames.map((name) => XLSX.utils.sheet_to_csv(workbook.Sheets[name])).join("\n");
  return [...new Set(extractInquiryRows(text).map((row) => row.客户型号).filter(Boolean))].slice(0, 50);
}

function RecycleBin({ items, onRestore, onDelete }) {
  if (!items.length) return <div className="recycle-empty">回收站为空，删除记录将在这里保留 7 天。</div>;
  return <div className="recycle-list">{items.map((item) => <div key={item.trashId}><span><strong>{item.item.data?.filename || item.item.label || "历史记录"}</strong><small>{remainingTrashDays(item)} 天后自动删除</small></span><button className="restore-button" onClick={() => onRestore(item)}><ClockCounterClockwise size={14} />恢复</button><button className="icon-button history-delete" title="永久删除" onClick={() => onDelete(item)}><Trash size={15} /></button></div>)}</div>;
}

function downloadXlsx(rows, filename, sheetName = "结果") {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 结果: "暂无数据" }]);
  worksheet["!cols"] = Object.keys(rows[0] || { 结果: "" }).map((key) => ({ wch: Math.min(42, Math.max(12, key.length * 2 + 4)) }));
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
  XLSX.writeFile(workbook, filename);
}

function quoteDownloadRows(result) {
  return (result?.lines || []).map((line) => {
    const priceColumns = {};
    const prices = line.prices?.length ? line.prices : [{ label: result.pricingLabel, price: line.price, source: line.priceSource }];
    prices.forEach((item, index) => {
      const label = item.label || result.pricingLabels?.[index] || result.pricingLabel || `价格类型${index + 1}`;
      priceColumns[`${label}报价`] = item.price ?? "";
      priceColumns[`${label}价格来源`] = item.source || "";
    });
    return {
      ...line.original,
      匹配内部型号: line.internal,
      "属性（自制/调货）": line.attribute || "",
      匹配OE号: line.oe,
      匹配UF号: line.uf,
      适用车型: line.vehicle,
      匹配依据: line.selectionReason || (line.matchedKind ? `通过${line.matchedKind}命中` : ""),
      ...priceColumns,
      型号来源: line.modelSource,
      校对状态: line.reviewStatus,
      异常说明: line.issue,
    };
  });
}

function quotePricingLabels(result) {
  if (result?.pricingLabels?.length) return result.pricingLabels;
  if (result?.pricings?.length) return result.pricings.map((item) => `${item.tier} ${item.packaging}${item.tax}`);
  return [result?.pricingLabel || "报价"];
}

function quoteLinePrices(line, result) {
  if (line?.prices?.length) return line.prices;
  return [{ label: result?.pricingLabel || "报价", price: line?.price ?? null, source: line?.priceSource || "" }];
}

function Badge({ tone = "blue", children, icon: Icon }) {
  return <span className={`badge badge-${tone}`}>{Icon && <Icon size={13} weight="fill" />}{children}</span>;
}

function Status({ value }) {
  const tone = value === "已确认" ? "success" : value === "已驳回" ? "danger" : value === "待校对" ? "warning" : "muted";
  return <span className={`status status-${tone}`}><span className="status-dot" />{value || "待处理"}</span>;
}

function EmptyState({ icon: Icon = FileMagnifyingGlass, title, description, action }) {
  return (
    <div className="empty-state">
      <div className="empty-icon"><Icon size={28} weight="duotone" /></div>
      <h3>{title}</h3>
      <p>{description}</p>
      {action}
    </div>
  );
}

function Toast({ message, tone = "success", onClose }) {
  if (!message) return null;
  const Icon = tone === "error" ? WarningCircle : CheckCircle;
  return <div className={`toast toast-${tone}`}><Icon size={20} weight="fill" /><span>{message}</span><button onClick={onClose} aria-label="关闭"><X size={16} /></button></div>;
}

function Sidebar({ page, setPage, counts }) {
  return (
    <aside className="sidebar">
      <button className="brand" onClick={() => setPage("overview")}>
        <span className="brand-logo-shell"><img className="brand-logo" src="/assets/hifly-logo.png" alt="HIFLY" /></span>
      </button>
      <nav>
        {NAV_ITEMS.map(({ id, label, icon: Icon }) => (
          <button key={id} className={`nav-item ${page === id ? "active" : ""}`} aria-current={page === id ? "page" : undefined} onClick={() => setPage(id)}>
            <Icon size={20} weight={page === id ? "fill" : "regular"} />
            <span>{label}</span>
            {id === "reviews" && counts?.reviewCount > 0 && <em>{counts.reviewCount}</em>}
          </button>
        ))}
      </nav>
      <div className="sidebar-foot">
        <div className="sync-state"><span className="online-dot" /><span><strong>知识库在线</strong><small>多人共享数据已同步</small></span></div>
        <div className="version-line"><span>型号 {counts?.modelVersion || "—"}</span><span>价格 {counts?.priceVersion || "—"}</span><span>目录 {counts?.catalogVersion || "—"}</span></div>
      </div>
    </aside>
  );
}

function Header({ page, setPage, counts }) {
  const [title, subtitle] = PAGE_META[page];
  const [openMenu, setOpenMenu] = useState("");
  const actionsRef = useRef(null);
  useEffect(() => {
    const close = (event) => { if (!actionsRef.current?.contains(event.target)) setOpenMenu(""); };
    const closeOnEscape = (event) => { if (event.key === "Escape") setOpenMenu(""); };
    document.addEventListener("pointerdown", close);
    document.addEventListener("keydown", closeOnEscape);
    return () => { document.removeEventListener("pointerdown", close); document.removeEventListener("keydown", closeOnEscape); };
  }, []);
  const pending = counts?.reviewCount || 0;
  return (
    <header className="page-header">
      <div><h1>{title}</h1><p>{subtitle}</p></div>
      <div className="header-actions" ref={actionsRef}>
        <button className="icon-button" aria-label="通知" aria-expanded={openMenu === "notifications"} onClick={() => setOpenMenu((value) => value === "notifications" ? "" : "notifications")}><BellSimple size={20} />{pending > 0 && <span className="notify-dot" />}</button>
        <button className="user-button" aria-expanded={openMenu === "user"} onClick={() => setOpenMenu((value) => value === "user" ? "" : "user")}><UserCircle size={26} weight="duotone" /><span><strong>当前用户</strong><small>协作成员</small></span><CaretDown className={openMenu === "user" ? "rotated" : ""} size={14} /></button>
        {openMenu === "notifications" && <div className="header-popover notification-popover" role="dialog" aria-label="通知中心"><div className="popover-title"><strong>通知中心</strong><span>{pending} 条待处理</span></div>{pending > 0 ? <button onClick={() => { setPage("reviews"); setOpenMenu(""); }}><span className="popover-icon warning"><SealCheck size={19} /></span><span><strong>二次校对待处理</strong><small>有 {pending} 条外网候选需要确认</small></span><ArrowRight size={16} /></button> : <div className="popover-empty"><CheckCircle size={22} weight="duotone" /><span><strong>暂无新通知</strong><small>所有校对项均已处理</small></span></div>}</div>}
        {openMenu === "user" && <div className="header-popover user-popover" role="dialog" aria-label="当前用户信息"><div className="user-summary"><UserCircle size={34} weight="duotone" /><span><strong>当前用户</strong><small>协作成员</small></span></div><div className="shared-state"><span className="online-dot" /><span><strong>共享知识库已连接</strong><small>型号、价格、目录与校对记录多人同步</small></span></div></div>}
      </div>
    </header>
  );
}

function Overview({ data, loading, setPage }) {
  const cards = [
    { label: "正式型号", value: data?.modelCount ?? "—", detail: `型号库 ${data?.modelVersion || "—"}`, icon: Tag, tone: "blue" },
    { label: "有效价格", value: data?.priceCount ?? "—", detail: `价格库 ${data?.priceVersion || "—"}`, icon: Database, tone: "cyan" },
    { label: "待二次校对", value: data?.reviewCount ?? "—", detail: "外网补充候选", icon: ShieldCheck, tone: "purple" },
    { label: "报价任务", value: data?.quoteCount ?? "—", detail: "累计生成记录", icon: FileMagnifyingGlass, tone: "orange" },
  ];
  return (
    <div className="content-stack">
      <section className="metric-grid">
        {cards.map(({ label, value, detail, icon: Icon, tone }) => (
          <article className="metric-card" key={label}>
            <div className={`metric-icon metric-${tone}`}><Icon size={23} weight="duotone" /></div>
            <div><p>{label}</p><strong>{loading ? "…" : Number(value).toLocaleString("zh-CN")}</strong><small>{detail}</small></div>
          </article>
        ))}
      </section>

      <section className="split-grid">
        <div className="panel quick-panel">
          <div className="panel-heading"><div><h2>开始工作</h2><p>选择一个常用任务快速进入</p></div></div>
          <div className="quick-actions">
            <button onClick={() => setPage("models")}><span className="quick-icon blue"><MagnifyingGlass size={22} /></span><span><strong>查型号</strong><small>内部 / OE / UF 号</small></span><ArrowRight size={18} /></button>
            <button onClick={() => setPage("quote")}><span className="quick-icon cyan"><FileArrowUp size={22} /></span><span><strong>生成报价</strong><small>上传客户询价 Excel</small></span><ArrowRight size={18} /></button>
            <button onClick={() => setPage("reviews")}><span className="quick-icon purple"><SealCheck size={22} /></span><span><strong>二次校对</strong><small>{data?.reviewCount || 0} 条外网候选待处理</small></span><ArrowRight size={18} /></button>
          </div>
        </div>
        <div className="panel source-panel">
          <div className="panel-heading"><div><h2>当前数据源</h2><p>报价与匹配均可追溯到源文件</p></div><Badge tone="green" icon={CheckCircle}>状态正常</Badge></div>
          <div className="source-list">
            <div><span className="file-icon"><BookOpenText size={21} /></span><span><strong>大厂型号匹配260809.xlsx</strong><small>{data?.modelCount || 0} 个标准型号</small></span><Badge>原始匹配表</Badge></div>
            <div><span className="file-icon"><Database size={21} /></span><span><strong>外贸价格体系60319.xlsx</strong><small>{data?.priceCount || 0} 条价格记录</small></span><Badge tone="cyan">价格源</Badge></div>
            <div><span className="file-icon"><FileMagnifyingGlass size={21} /></span><span><strong>海福来点火线圈目录60625.xlsx</strong><small>{data?.catalogCount || 0} 个产品图片</small></span><Badge tone="green">产品目录</Badge></div>
          </div>
        </div>
      </section>

      <section className="panel activity-panel">
        <div className="panel-heading"><div><h2>知识库动态</h2><p>最近的版本更新记录</p></div><button className="text-button" onClick={() => setPage("knowledge")}>查看全部 <ArrowRight size={15} /></button></div>
        <div className="table-wrap">
          <table><thead><tr><th>类型</th><th>版本</th><th>源文件</th><th>更新方式</th><th>更新时间</th><th>状态</th></tr></thead>
          <tbody>{(data?.recentVersions || []).map((item) => <tr key={item.id}><td><strong>{item.type}</strong></td><td className="mono">{item.version}</td><td>{item.filename}</td><td>{item.operator}</td><td>{formatDate(item.created_at)}</td><td><Status value="已确认" /></td></tr>)}</tbody></table>
        </div>
      </section>
    </div>
  );
}

function CodeValues({ values = [], webValues = [] }) {
  const web = new Set(webValues.map((item) => String(item).toUpperCase().replace(/[^A-Z0-9]/g, "")));
  if (!values.length) return <span className="muted">—</span>;
  return <span className="code-values">{values.slice(0, 8).map((item) => { const isWeb = web.has(String(item).toUpperCase().replace(/[^A-Z0-9]/g, "")); return <span key={item} className={isWeb ? "web-confirmed-code" : ""}>{item}{isWeb && <em>外网确认</em>}</span>; })}{values.length > 8 && <small>另有 {values.length - 8} 个</small>}</span>;
}

function classifiedModelCodes(model = {}) {
  const oeNos = [...(model.oeNos || [])];
  const ufNos = [...(model.ufNos || []), ...oeNos.filter((item) => /^UF/i.test(String(item).trim()))];
  const webOeNos = [...(model.webOeNos || [])];
  const webUfNos = [...(model.webUfNos || []), ...webOeNos.filter((item) => /^UF/i.test(String(item).trim()))];
  const uniqueCodes = (values) => [...new Set(values.map((item) => String(item || "").trim()).filter(Boolean))];
  return {
    oeNos: uniqueCodes(oeNos.filter((item) => !/^UF/i.test(String(item).trim()))),
    ufNos: uniqueCodes(ufNos),
    webOeNos: uniqueCodes(webOeNos.filter((item) => !/^UF/i.test(String(item).trim()))),
    webUfNos: uniqueCodes(webUfNos),
  };
}

function ModelSearch({ onReviewCreated, showToast }) {
  const [query, setQuery] = useStoredState("hifly-model-query", "");
  const [result, setResult] = useStoredState("hifly-model-current", null);
  const [history, setHistory] = useStoredState("hifly-model-history", []);
  const [trash, setTrash] = useStoredState("hifly-history-trash", []);
  const [loading, setLoading] = useState(false);
  const [draggingFile, setDraggingFile] = useState(false);
  const [fileState, setFileState] = useState({ stage: "", progress: 0 });
  const [trashOpen, setTrashOpen] = useState(false);
  const [manualInputs, setManualInputs] = useState({});
  const [confirming, setConfirming] = useState("");
  const modelTrash = useMemo(() => validTrashItems(trash).filter((item) => item.type === "model"), [trash]);
  useEffect(() => { setTrash((items) => validTrashItems(items)); }, [setTrash]);
  const search = async (value = query) => {
    const queries = splitModelQueries(value);
    if (!queries.length) return showToast("请输入至少一个型号", "error");
    setLoading(true);
    try {
      const data = await api.searchModelsBatch(queries);
      setResult(data);
      setHistory((items) => [{ id: Date.now(), createdAt: data.searchedAt, label: queries.join("、"), data }, ...items].slice(0, 20));
      if (data.ignoredQueries?.length) showToast(`已忽略 ${data.ignoredQueries.length} 个数量、表头或年份等非型号内容`, "error");
      if (data.results.some((item) => item.webSupplement?.status === "待校对")) onReviewCreated();
    } catch (error) { showToast(error.message, "error"); }
    finally { setLoading(false); }
  };
  const downloadResult = (data = result) => {
    if (!data) return;
    const rows = data.results.flatMap((item) => item.local.length ? item.local.map((model) => { const codes = classifiedModelCodes(model); return ({
      查询输入: item.query, 内部型号: model.internal, OE号: codes.oeNos.join(" / "), UF号: codes.ufNos.join(" / "), 适用车型: model.vehicle,
      命中依据: model.matchedKind, 选择规则: item.preferredModelId === model.id ? item.preferenceReason : "", 来源: model.webOeNos.length || model.webUfNos.length ? "原始表 + 外网确认" : "原始匹配表", 状态: model.reviewStatus,
    }); }) : [{ 查询输入: item.query, 内部型号: "未匹配", OE号: item.webSupplement?.oeNos?.filter((value) => !/^UF/i.test(value))?.join(" / ") || "", UF号: [...(item.webSupplement?.ufNos || []), ...(item.webSupplement?.oeNos || []).filter((value) => /^UF/i.test(value))].join(" / "), 来源: item.webSupplement?.sources?.length ? "外网候选" : "未匹配", 状态: item.webSupplement?.status || "" }]);
    downloadXlsx(rows, `型号查询-${new Date().toISOString().slice(0, 10)}.xlsx`, "型号查询");
  };
  const loadQueryFile = async (file) => {
    if (!file) return;
    setLoading(true);
    try {
      const queries = await extractModelQueriesFromFile(file, setFileState);
      if (!queries.length) throw new Error("文件中没有识别到可查询的型号");
      const value = queries.join("\n");
      setQuery(value);
      setFileState({ stage: `已识别 ${queries.length} 个型号`, progress: 100 });
      await search(value);
    } catch (error) { showToast(error.message, "error"); }
    finally { setLoading(false); }
  };
  const deleteHistory = (item) => {
    setHistory((items) => items.filter((historyItem) => historyItem.id !== item.id));
    setTrash((items) => [{ trashId: `model-${item.id}-${Date.now()}`, type: "model", deletedAt: Date.now(), expiresAt: Date.now() + HISTORY_TRASH_TTL, item }, ...validTrashItems(items)]);
    if (result?.searchedAt === item.data?.searchedAt) setResult(null);
    showToast("查询记录已移入回收站，7天内可恢复");
  };
  const restoreHistory = (trashItem) => {
    setHistory((items) => [trashItem.item, ...items.filter((item) => item.id !== trashItem.item.id)].slice(0, 20));
    setTrash((items) => items.filter((item) => item.trashId !== trashItem.trashId));
    showToast("查询记录已恢复");
  };
  const permanentlyDeleteHistory = (trashItem) => {
    if (!window.confirm("确认永久删除这条查询记录吗？")) return;
    setTrash((items) => items.filter((item) => item.trashId !== trashItem.trashId));
  };
  const confirmUnmatched = async (item) => {
    const targetInternal = String(manualInputs[item.query] || "").trim();
    if (!targetInternal) return showToast("请输入要关联的内部型号", "error");
    setConfirming(item.query);
    try {
      const confirmation = await api.confirmMapping(item.query, targetInternal, "型号查询");
      const nextResult = { ...result, results: result.results.map((resultItem) => resultItem.query === item.query ? { ...resultItem, local: [confirmation.model], preferredModelId: confirmation.model.id, preferenceReason: "用户手动确认" } : resultItem) };
      setResult(nextResult);
      setHistory((items) => items.map((historyItem) => historyItem.data?.searchedAt === result.searchedAt ? { ...historyItem, data: nextResult } : historyItem));
      setManualInputs((values) => ({ ...values, [item.query]: "" }));
      onReviewCreated();
      showToast("已写入知识库并生成校对记录");
    } catch (error) { showToast(error.message, "error"); }
    finally { setConfirming(""); }
  };
  const confirmExistingCandidate = async (item) => {
    const key = `candidate-${item.query}`;
    const targetInternal = String(manualInputs[key] || "").trim();
    if (!targetInternal) return showToast("请选择要确认的内部型号", "error");
    setConfirming(key);
    try {
      const confirmation = await api.confirmMapping(item.query, targetInternal, "型号查询");
      const nextResult = { ...result, results: result.results.map((resultItem) => resultItem.query === item.query ? { ...resultItem, local: [confirmation.model], preferredModelId: confirmation.model.id, preferenceReason: "用户已选择通用型号关系" } : resultItem) };
      setResult(nextResult);
      setHistory((items) => items.map((historyItem) => historyItem.data?.searchedAt === result.searchedAt ? { ...historyItem, data: nextResult } : historyItem));
      setManualInputs((values) => ({ ...values, [key]: "" }));
      onReviewCreated();
      showToast("已确认并同步到大厂型号匹配知识库");
    } catch (error) { showToast(error.message, "error"); }
    finally { setConfirming(""); }
  };
  return (
    <div className="content-stack">
      <section className="search-hero panel">
        <div className="search-copy"><Badge tone="blue" icon={MagnifyingGlass}>批量智能匹配</Badge><h2>多个型号，一次查全关联关系</h2><p>每行一个型号，或用逗号、分号分隔；单次最多 50 个。</p></div>
        <form className="model-search batch-search" onSubmit={(event) => { event.preventDefault(); search(); }}>
          <MagnifyingGlass size={22} />
          <textarea value={query} onChange={(event) => setQuery(event.target.value)} placeholder={"输入内部型号、OE 号或 UF 号\n例如：DQ910844；90919-02244；UF333"} autoFocus />
          <button className="primary-button" disabled={loading}>{loading ? <><SpinnerGap className="spin" size={18} /> 正在匹配</> : <>批量查询 <ArrowRight size={17} /></>}</button>
        </form>
        <label className={`model-file-upload ${draggingFile ? "is-dragging" : ""}`} role="button" tabIndex={0} aria-label="上传型号文件" onKeyDown={(event) => { if (["Enter", " "].includes(event.key)) { event.preventDefault(); event.currentTarget.querySelector("input")?.click(); } }} onDragEnter={(event) => { event.preventDefault(); setDraggingFile(true); }} onDragOver={(event) => { event.preventDefault(); setDraggingFile(true); }} onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget)) setDraggingFile(false); }} onDrop={(event) => { event.preventDefault(); setDraggingFile(false); loadQueryFile(event.dataTransfer.files?.[0]); }}>
          <input type="file" accept=".xlsx,.xls,.csv,.txt,.pdf,.png,.jpg,.jpeg,application/pdf,image/png,image/jpeg,text/plain,text/csv" onChange={(event) => { loadQueryFile(event.target.files?.[0]); event.target.value = ""; }} />
          <CloudArrowUp size={17} /><span>{draggingFile ? "松开即可识别并查询" : "拖拽或选择文字、Excel、PDF、图片"}</span>{fileState.stage && <small>{fileState.stage}</small>}
        </label>
        <div className="search-examples"><span>试试：</span>{["DQ910844\n90919-02244", "DQ1212X\n0221503004", "DQ910852"].map((item) => <button key={item} onClick={() => { setQuery(item); search(item); }}>{item.replace("\n", " + ")}</button>)}</div>
      </section>

      {loading && <section className="panel searching-panel"><div className="searching-steps"><span className="done"><Check size={15} />规范化输入</span><span className="done"><Check size={15} />检索本地型号库</span><span className="active"><SpinnerGap className="spin" size={15} />联网补齐缺失字段</span><span>生成可追溯结果</span></div></section>}

      {(history.length > 0 || modelTrash.length > 0) && <section className="panel compact-history"><div className="panel-heading"><div><h2>最近型号查询</h2><p>保存在当前浏览器，最多 20 条</p></div><button className="secondary-button" onClick={() => setTrashOpen((value) => !value)}><Trash size={15} />回收站 {modelTrash.length || ""}</button></div>{history.length > 0 && <div className="history-chips">{history.slice(0, 8).map((item) => <div key={item.id}><button onClick={() => { setResult(item.data); setQuery(item.data.queries.join("\n")); }}><ClockCounterClockwise size={15} /><span><strong>{item.label}</strong><small>{formatDate(item.createdAt)}</small></span></button><button className="icon-button history-delete" title="移入回收站" onClick={() => deleteHistory(item)}><Trash size={15} /></button></div>)}</div>}{trashOpen && <RecycleBin items={modelTrash} onRestore={restoreHistory} onDelete={permanentlyDeleteHistory} />}</section>}

      {result && <section className="panel results-panel batch-results">
        <div className="panel-heading"><div><h2>批量匹配结果</h2><p>已处理 {result.results.length} 个输入</p></div><button className="secondary-button" onClick={() => downloadResult()}><DownloadSimple size={16} />下载型号表</button></div>
        {result.results.map((item) => <div className="batch-result-block" key={item.query}>
          <div className="batch-result-title"><div><Badge tone="blue">查询输入</Badge><strong className="mono">{item.query}</strong>{item.preferenceReason && <small>{item.preferenceReason}</small>}</div><span>{item.local.length} 条本地匹配</span></div>
          {item.local.length ? <>
            <div className="table-wrap"><table className="result-table"><thead><tr><th>内部型号</th><th>OE 号</th><th>UF 号</th><th>适用车型</th><th>来源</th><th>状态</th></tr></thead><tbody>{item.local.map((model) => { const codes = classifiedModelCodes(model); const matchedKind = /^UF/i.test(item.query) ? "UF号" : model.matchedKind; return <tr key={model.id} className={item.preferredModelId === model.id ? "selected" : ""}><td><strong className="mono">{model.internal}</strong><small>命中：{matchedKind}{item.preferredModelId === model.id && item.preferenceReason ? " · 系统优先" : ""}</small></td><td className="codes"><CodeValues values={codes.oeNos} webValues={codes.webOeNos} /></td><td className="codes"><CodeValues values={codes.ufNos} webValues={codes.webUfNos} /></td><td>{model.vehicle || "—"}</td><td><Badge tone={model.sourceType === "user_confirmed" ? "green" : "blue"} icon={model.sourceType === "user_confirmed" ? CheckCircle : BookOpenText}>{model.sourceType === "user_confirmed" ? "用户确认" : "原始匹配表"}</Badge></td><td><Status value="已确认" /></td></tr>; })}</tbody></table></div>
            {item.local.length > 1 && !item.preferredModelId && <div className="manual-confirm"><span><strong>请选择正确的内部型号</strong><small>通用型号匹配结果不一致，确认后会同步更新知识库</small></span><select value={manualInputs[`candidate-${item.query}`] || ""} onChange={(event) => setManualInputs((values) => ({ ...values, [`candidate-${item.query}`]: event.target.value }))}><option value="">请选择内部型号</option>{item.local.map((model) => <option key={model.id} value={model.internal}>{model.internal}</option>)}</select><button className="primary-button" disabled={confirming === `candidate-${item.query}`} onClick={() => confirmExistingCandidate(item)}>{confirming === `candidate-${item.query}` ? <SpinnerGap className="spin" size={16} /> : <Check size={16} />}确认并入库</button></div>}
          </> : <><EmptyState icon={GlobeHemisphereWest} title="本地知识库暂无记录" description="可在下方选择已有内部型号或填写新型号，确认后同步进入知识库。" /><div className="manual-confirm"><span><strong>手动确认内部型号</strong><small>可填写已有型号或创建新的内部型号</small></span><input value={manualInputs[item.query] || ""} onChange={(event) => setManualInputs((values) => ({ ...values, [item.query]: event.target.value }))} placeholder="例如 DQ910844" /><button className="primary-button" disabled={confirming === item.query} onClick={() => confirmUnmatched(item)}>{confirming === item.query ? <SpinnerGap className="spin" size={16} /> : <Check size={16} />}确认并入库</button></div></>}
          {item.webSupplement && <WebSupplement data={item.webSupplement} compact />}
        </div>)}
      </section>}
    </div>
  );
}

function WebSupplement({ data, compact = false }) {
  return (
    <section className={`${compact ? "embedded-web-panel" : "panel web-panel"}`}>
      <div className="panel-heading"><div className="web-heading"><span className="web-icon"><GlobeHemisphereWest size={22} weight="duotone" /></span><div><h2>外网自动补充</h2><p>结果与本地数据分开保存，需二次校对后才能进入正式知识库</p></div></div><Badge tone="purple" icon={GlobeHemisphereWest}>外网补充</Badge></div>
      {data.sources?.length ? <>
        <div className="platform-proof"><span>已检索平台</span>{["Google", "Amazon", "eBay"].map((platform) => <Badge key={platform} tone={data.platforms?.includes(platform) ? "green" : "gray"}>{platform}{data.platforms?.includes(platform) ? " ✓" : " 无结果"}</Badge>)}</div>
        <div className="candidate-strip"><div><span>候选 OE 号</span><strong className="mono">{data.oeNos?.join(" / ") || "网页未提取到明确编号"}</strong><small>{data.verifiedOeNos?.length ? `跨平台确认：${data.verifiedOeNos.join(" / ")}` : "暂无两个平台共同支持的 OE"}</small></div><div><span>候选 UF 号</span><strong className="mono">{data.ufNos?.join(" / ") || "网页未提取到明确编号"}</strong><small>{data.verifiedUfNos?.length ? `跨平台确认：${data.verifiedUfNos.join(" / ")}` : "暂无两个平台共同支持的 UF"}</small></div><div><span>匹配置信度</span><strong>{data.confidence}%</strong></div><div><span>校对状态</span><Status value="待校对" /></div></div>
        <div className="source-links"><h3>引用来源</h3>{data.sources.slice(0, 6).map((source, index) => <a href={source.url} target="_blank" rel="noreferrer" key={`${source.url}-${index}`}><span>{source.platform?.slice(0, 1) || index + 1}</span><div><strong>{source.title}</strong><small>{source.platform || source.provider} · {source.snippet || source.url}</small></div><ArrowRight size={16} /></a>)}</div>
      </> : <div className="inline-alert warning"><WarningCircle size={20} /><span><strong>本次联网补充未找到可靠来源</strong><small>{data.error || "可稍后重试，现有本地结果不受影响。"}</small></span></div>}
    </section>
  );
}

function Segmented({ value, options, onChange }) {
  return <div className="segmented">{options.map((option) => <button key={option} type="button" aria-pressed={value === option} className={value === option ? "active" : ""} onClick={() => onChange(option)}>{option}</button>)}</div>;
}

function QuotePage({ showToast, refreshOverview }) {
  const [file, setFile] = useState(null);
  const [inputMode, setInputMode] = useStoredState("hifly-quote-input-mode", "file");
  const [textInquiry, setTextInquiry] = useStoredState("hifly-quote-text-inquiry", "");
  const [dragging, setDragging] = useState(false);
  const [parseState, setParseState] = useState({ stage: "", progress: 0 });
  const [pricings, setPricings] = useStoredState("hifly-quote-pricings", () => {
    try {
      const legacy = JSON.parse(localStorage.getItem("hifly-quote-pricing") || "null");
      return legacy ? [legacy] : [{ packaging: "彩盒", tax: "未税", tier: "4S" }];
    } catch { return [{ packaging: "彩盒", tax: "未税", tier: "4S" }]; }
  });
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useStoredState("hifly-quote-current", null);
  const [history, setHistory] = useStoredState("hifly-quote-history", []);
  const [trash, setTrash] = useStoredState("hifly-history-trash", []);
  const [trashOpen, setTrashOpen] = useState(false);
  const textInquiryCount = useMemo(() => extractInquiryRows(normalizeTextInquiry(textInquiry)).length, [textInquiry]);
  const quoteTrash = useMemo(() => validTrashItems(trash).filter((item) => item.type === "quote"), [trash]);
  useEffect(() => { setTrash((items) => validTrashItems(items)); }, [setTrash]);
  useEffect(() => {
    if (!isLegacyMisdetectedQuote(result)) return;
    setResult(null);
    setParseState({ stage: "旧版结果误识别了序号列，请重新上传表格", progress: 0 });
    showToast("已拦截旧版错误结果，请重新上传原表生成报价", "error");
  }, [result, setResult, showToast]);
  const chooseFile = (next) => {
    if (!next) return;
    if (!isSupportedQuoteFile(next)) return showToast("仅支持文字、Excel、PDF、PNG 和 JPG 文件", "error");
    if (next.size > 25 * 1024 * 1024) return showToast("文件不能超过 25MB", "error");
    setInputMode("file"); setFile(next); setResult(null); setParseState({ stage: "已准备解析", progress: 0 });
  };
  const dropFile = (event) => {
    event.preventDefault();
    setDragging(false);
    const dropped = event.dataTransfer.files?.[0];
    if (dropped) chooseFile(dropped);
  };
  const createQuote = async () => {
    const normalizedText = normalizeTextInquiry(textInquiry);
    const activeFile = inputMode === "text" ? (normalizedText ? new File([normalizedText], "文字询价.txt", { type: "text/plain" }) : null) : file;
    if (!activeFile) return showToast(inputMode === "text" ? "请先填写客户询价内容" : "请先上传客户询价表", "error");
    if (inputMode === "text" && !textInquiryCount) return showToast("没有识别到有效型号，请检查输入内容", "error");
    setLoading(true);
    try {
      setParseState({ stage: inputMode === "text" ? "正在解析文字询价" : quoteFileType(activeFile) === "Excel" ? "正在解析 Excel" : `正在识别${quoteFileType(activeFile)}`, progress: 2 });
      const prepared = await prepareQuoteFile(activeFile, setParseState);
      let data = await api.previewQuote(prepared.file, pricings);
      const attachedImages = data.lines.map((line, lineIndex) => ({
        lineIndex,
        image: prepared.customerImages?.[data.sheetName]?.[line.sourceRow] || prepared.customerImagesByIndex?.[lineIndex] || "",
        validated: prepared.imageValidation === "embedded-image-column",
        status: prepared.imageStatus || "",
      })).filter((item) => item.image && item.validated);
      if (attachedImages.length) data = await api.attachQuoteImages(data.id, attachedImages);
      data = { ...data, recognition: { mode: prepared.mode, extractedRows: prepared.extractedRows } };
      setResult(data);
      setHistory((items) => [{ id: data.id, createdAt: new Date().toISOString(), data }, ...items.filter((item) => item.id !== data.id)].slice(0, 20));
      refreshOverview();
      setParseState({ stage: prepared.extractedRows ? `识别完成，共提取 ${prepared.extractedRows} 条型号` : "Excel 解析完成", progress: 100 });
      showToast("报价已生成，可以下载结果");
    }
    catch (error) { setParseState({ stage: "解析失败", progress: 0 }); showToast(error.message, "error"); }
    finally { setLoading(false); }
  };
  const updateResult = (data) => {
    setResult(data);
    setHistory((items) => items.map((item) => item.id === data.id ? { ...item, data } : item));
  };
  const deleteHistory = (item) => {
    setHistory((items) => items.filter((historyItem) => historyItem.id !== item.id));
    setTrash((items) => [{ trashId: `quote-${item.id}-${Date.now()}`, type: "quote", deletedAt: Date.now(), expiresAt: Date.now() + HISTORY_TRASH_TTL, item }, ...validTrashItems(items)]);
    if (result?.id === item.id) setResult(null);
    showToast("历史报价已移入回收站，7天内可恢复");
  };
  const restoreHistory = (trashItem) => {
    setHistory((items) => [trashItem.item, ...items.filter((item) => item.id !== trashItem.item.id)].slice(0, 20));
    setTrash((items) => items.filter((item) => item.trashId !== trashItem.trashId));
    showToast("历史报价已恢复");
  };
  const permanentlyDeleteHistory = (trashItem) => {
    if (!window.confirm("确认永久删除这份历史报价吗？")) return;
    setTrash((items) => items.filter((item) => item.trashId !== trashItem.trashId));
  };
  const updatePricing = (index, patch) => {
    const next = { ...pricings[index], ...patch };
    const duplicated = pricings.some((item, itemIndex) => itemIndex !== index && item.tier === next.tier && item.packaging === next.packaging && item.tax === next.tax);
    if (duplicated) return showToast("这个价格类型已经添加过了", "error");
    setPricings((items) => items.map((item, itemIndex) => itemIndex === index ? next : item));
  };
  const addPricing = () => {
    const used = new Set(pricings.map((item) => `${item.tier}|${item.packaging}|${item.tax}`));
    const next = ["NP", "LP", "SLP", "2S", "3S", "4S"].flatMap((tier) => ["彩盒", "中性"].flatMap((packaging) => ["含税", "未税"].map((tax) => ({ tier, packaging, tax })))).find((item) => !used.has(`${item.tier}|${item.packaging}|${item.tax}`));
    if (!next) return showToast("24 种价格组合均已添加", "error");
    setPricings((items) => [...items, next]);
  };
  const removePricing = (index) => setPricings((items) => items.length > 1 ? items.filter((_, itemIndex) => itemIndex !== index) : items);
  return (
    <div className="content-stack">
      <section className="quote-layout">
        <div className="panel quote-config">
          <div className="panel-heading"><div><Badge tone="blue">步骤 1</Badge><h2>上传客户询价表</h2><p>支持文字、Excel、PDF、PNG 和 JPG</p></div></div>
          <div className="quote-input-tabs" role="tablist" aria-label="询价输入方式">
            <button type="button" role="tab" aria-selected={inputMode === "file"} className={inputMode === "file" ? "active" : ""} onClick={() => setInputMode("file")}><UploadSimple size={17} />文件上传</button>
            <button type="button" role="tab" aria-selected={inputMode === "text"} className={inputMode === "text" ? "active" : ""} onClick={() => setInputMode("text")}><PencilSimpleLine size={17} />文字询价</button>
          </div>
          {inputMode === "file" ? <label
            className={`dropzone ${file ? "has-file" : ""} ${dragging ? "is-dragging" : ""}`}
            role="button"
            tabIndex={0}
            aria-label="上传客户询价表"
            onKeyDown={(event) => { if (["Enter", " "].includes(event.key)) { event.preventDefault(); event.currentTarget.querySelector("input")?.click(); } }}
            onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
            onDragOver={(event) => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; setDragging(true); }}
            onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget)) setDragging(false); }}
            onDrop={dropFile}
          >
            <input type="file" accept=".xlsx,.xls,.csv,.txt,.pdf,.png,.jpg,.jpeg,application/pdf,image/png,image/jpeg,text/plain,text/csv" onChange={(event) => { chooseFile(event.target.files?.[0]); event.target.value = ""; }} />
            {file ? <><span className="upload-icon success"><CheckCircle size={28} weight="fill" /></span><strong>{file.name}</strong><small>{quoteFileType(file)} · {(file.size / 1024).toFixed(1)} KB · {parseState.stage || "已准备解析"}</small>{loading && parseState.progress > 0 ? <span className="upload-progress"><i style={{ width: `${parseState.progress}%` }} /></span> : null}<em>点击或拖拽更换文件</em></> : <><span className="upload-icon"><UploadSimple size={28} /></span><strong>{dragging ? "松开鼠标即可上传" : "拖拽询价表到这里，或点击上传"}</strong><small>支持 Excel、PDF 和图片，系统会自动识别型号内容</small><em>选择文件</em></>}
          </label> : <div className="text-inquiry-box">
            <textarea value={textInquiry} onChange={(event) => { setTextInquiry(event.target.value); setResult(null); setParseState({ stage: "已填写文字询价", progress: 0 }); }} placeholder={'直接粘贴客户询价，例如：\n90919-02248   100\n33400-51K80   200\nUF333'} autoFocus />
            <div className="text-inquiry-meta"><span>支持 OE、UF、内部型号；每行一条，也可用逗号或分号分隔</span><strong className={textInquiryCount ? "has-items" : ""}>已识别 {textInquiryCount} 条</strong></div>
          </div>}
        </div>
        <div className="panel quote-config">
          <div className="panel-heading"><div><Badge tone="blue">步骤 2</Badge><h2>限定价格类型</h2><p>当前组合将映射到价格表中的唯一列</p></div></div>
          <div className="pricing-list">
            {pricings.map((pricing, index) => <div className="pricing-choice" key={`${index}-${pricing.tier}-${pricing.packaging}-${pricing.tax}`}>
              <div className="pricing-choice-head"><span><strong>价格类型 {index + 1}</strong><small>{pricing.tier} · {pricing.packaging}{pricing.tax}</small></span>{pricings.length > 1 && <button title="删除此价格类型" onClick={() => removePricing(index)}><X size={15} /></button>}</div>
              <div className="config-row"><label>包装方式</label><Segmented value={pricing.packaging} options={["彩盒", "中性"]} onChange={(packaging) => updatePricing(index, { packaging })} /></div>
              <div className="config-row"><label>税别</label><Segmented value={pricing.tax} options={["含税", "未税"]} onChange={(tax) => updatePricing(index, { tax })} /></div>
              <div className="config-row"><label>价格档位</label><Segmented value={pricing.tier} options={["NP", "LP", "SLP", "2S", "3S", "4S"]} onChange={(tier) => updatePricing(index, { tier })} /></div>
            </div>)}
          </div>
          <button className="add-pricing-button" onClick={addPricing}><Plus size={17} weight="bold" />添加价格类型</button>
          <div className="pricing-summary"><span className="pricing-summary-icon"><Database size={18} /></span><div><small>本次共使用 {pricings.length} 个价格类型</small><strong>{pricings.map((item) => `${item.tier} · ${item.packaging}${item.tax}`).join("；")}</strong></div><Badge tone="green" icon={CheckCircle}>已定位价格列</Badge></div>
          <button className="primary-button wide" onClick={createQuote} disabled={(inputMode === "file" ? !file : !textInquiry.trim()) || loading}>{loading ? <><SpinnerGap className="spin" size={19} />{parseState.stage || "正在匹配并联网补充"}</> : <><FileMagnifyingGlass size={19} />生成报价预览</>}</button>
        </div>
      </section>
      {(history.length > 0 || quoteTrash.length > 0) && <section className="panel compact-history"><div className="panel-heading"><div><h2>历史报价</h2><p>保存在当前浏览器，删除后在回收站保留 7 天</p></div><button className="secondary-button" onClick={() => setTrashOpen((value) => !value)}><Trash size={15} />回收站 {quoteTrash.length || ""}</button></div>{history.length > 0 && <div className="history-list">{history.map((item) => { const legacy = isLegacyMisdetectedQuote(item.data); return <div key={item.id}><button onClick={() => legacy ? showToast("这份旧记录误识别了序号列，请重新上传原表", "error") : setResult(item.data)}><ClockCounterClockwise size={17} /><span><strong>{item.data.filename}</strong><small>{legacy ? "旧版识别失效 · 请重新上传" : `${item.data.pricingLabel} · ${formatDate(item.createdAt)}`}</small></span></button>{!legacy && <a className="icon-button" href={`/api/quotes/${item.id}/export`} title="下载含图片的历史报价" aria-label={`下载 ${item.data.filename}`}><DownloadSimple size={16} /></a>}<button className="icon-button history-delete" title="移入回收站" aria-label={`删除 ${item.data.filename}`} onClick={() => deleteHistory(item)}><Trash size={16} /></button></div>; })}</div>}{trashOpen && <RecycleBin items={quoteTrash} onRestore={restoreHistory} onDelete={permanentlyDeleteHistory} />}</section>}
      {result && <QuoteResult result={result} onUpdate={updateResult} showToast={showToast} onKnowledgeUpdated={refreshOverview} />}
    </div>
  );
}

function QuoteImageCell({ src, status, alt, blankWhenEmpty = false }) {
  if (src) return <a className="quote-image" href={src} target="_blank" rel="noreferrer" title="点击查看大图"><img src={src} alt={alt} /></a>;
  if (status) return <span className="image-pending"><WarningCircle size={14} />{status}</span>;
  if (blankWhenEmpty) return null;
  return <span className="muted">—</span>;
}

function QuoteResult({ result, onUpdate, showToast, onKnowledgeUpdated }) {
  const [selecting, setSelecting] = useState("");
  const [manualInputs, setManualInputs] = useState({});
  const pricingLabels = quotePricingLabels(result);
  const selectCandidate = async (lineIndex, modelId) => {
    const key = `${lineIndex}-${modelId}`;
    setSelecting(key);
    try {
      onUpdate(await api.selectQuoteCandidate(result.id, lineIndex, modelId));
      onKnowledgeUpdated();
      showToast("已按所选内部型号重新匹配价格，并同步更新知识库");
    } catch (error) {
      showToast(error.message, "error");
    } finally {
      setSelecting("");
    }
  };
  const confirmLine = async (lineIndex) => {
    const targetInternal = String(manualInputs[lineIndex] || "").trim();
    if (!targetInternal) return showToast("请输入要关联的内部型号", "error");
    setSelecting(`manual-${lineIndex}`);
    try {
      onUpdate(await api.confirmQuoteLine(result.id, lineIndex, targetInternal));
      setManualInputs((values) => ({ ...values, [lineIndex]: "" }));
      onKnowledgeUpdated();
      showToast("已重新匹配价格，并写入知识库和校对记录");
    } catch (error) { showToast(error.message, "error"); }
    finally { setSelecting(""); }
  };
  return (
    <section className="panel quote-result">
      <div className="panel-heading"><div><Badge tone="green" icon={CheckCircle}>报价完成</Badge><h2>{result.filename}</h2><p>识别型号列：{result.detectedColumn} · 使用价格：{result.pricingLabel}</p></div><a className="primary-button download" href={`/api/quotes/${result.id}/export`}><DownloadSimple size={18} />下载含图片报价</a></div>
      <div className="result-summary"><div><span>总行数</span><strong>{result.total}</strong></div><div><span>已匹配价格</span><strong className="green-text">{result.matched}</strong></div><div><span>待选内部型号</span><strong className="blue-text">{result.needsSelection || 0}</strong></div><div><span>待校对</span><strong className="orange-text">{result.needsReview}</strong></div><div><span>未匹配 / 无价格</span><strong className="red-text">{result.unmatched}</strong></div></div>
      <div className="table-wrap"><table><thead><tr><th>客户输入</th><th>客户询价图片</th><th>内部型号</th><th>属性</th><th>产品目录图片</th><th>OE / UF</th>{pricingLabels.map((label) => <th key={label}>{label}</th>)}<th>型号来源</th><th>状态</th></tr></thead><tbody>{result.lines.slice(0, 60).map((line, index) => {
        const needsManualConfirm = line.reviewStatus !== "已确认" && !line.selectedModelId && !(line.candidateOptions?.length > 1);
        return <tr key={`${line.input}-${index}`}><td className="mono">{line.input || "—"}</td><td><QuoteImageCell src={line.customerImageValidated ? line.customerImage : ""} status={line.customerImageValidated ? line.customerImageStatus : ""} alt={`${line.input} 客户询价图片`} blankWhenEmpty /></td><td>{line.candidateOptions?.length > 1 && !line.selectedModelId ? <select className="candidate-select" value="" disabled={selecting.startsWith(`${index}-`)} onChange={(event) => selectCandidate(index, Number(event.target.value))}><option value="">请选择内部型号</option>{line.candidateOptions.map((option) => <option key={option.modelId} value={option.modelId}>{option.internal} · {quoteLinePrices(option, result).map((item) => item.price === null ? "—" : `¥${Number(item.price).toFixed(2)}`).join(" / ")}</option>)}</select> : needsManualConfirm ? <div className="manual-quote-confirm"><input value={manualInputs[index] || ""} onChange={(event) => setManualInputs((values) => ({ ...values, [index]: event.target.value }))} placeholder="填写已有或新内部型号" /><button disabled={selecting === `manual-${index}`} onClick={() => confirmLine(index)}>{selecting === `manual-${index}` ? <SpinnerGap className="spin" size={14} /> : <Check size={14} />}确认</button></div> : <strong className="mono">{line.internal || "—"}</strong>}<small>{line.selectionReason || (line.matchedKind ? `通过${line.matchedKind}命中` : line.vehicle)}</small></td><td>{line.attribute ? <Badge tone={line.attribute.includes("自制") ? "green" : "warning"}>{line.attribute}</Badge> : "—"}</td><td><QuoteImageCell src={line.catalogImage} alt={`${line.internal} 产品目录图片`} /></td><td className="codes">{line.oe || line.uf || "—"}</td>{quoteLinePrices(line, result).map((item, priceIndex) => <td className="price-cell" key={`${item.label || pricingLabels[priceIndex]}-${priceIndex}`} title={item.source || "暂无价格来源"}>{item.price === null ? "—" : `¥ ${Number(item.price).toFixed(2)}`}</td>)}<td><Badge tone={line.modelSource === "外网补充" ? "purple" : line.modelSource === "原始匹配表" ? "blue" : line.modelSource === "用户确认" ? "green" : "gray"}>{line.modelSource}</Badge></td><td>{line.reviewStatus === "待选择" ? <Status value="待选择" /> : line.issue ? <span className="issue" title={line.issue}><WarningCircle size={16} />需处理</span> : <Status value="已确认" />}</td></tr>;
      })}</tbody></table></div>
    </section>
  );
}

function ReviewsPage({ refreshKey, showToast, refreshOverview, onPendingCountChange }) {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [targetInputs, setTargetInputs] = useState({});
  const [candidateSelections, setCandidateSelections] = useState({});
  const load = useCallback(async () => {
    setLoading(true);
    try {
      const nextItems = await api.reviews();
      setItems(nextItems);
      onPendingCountChange(nextItems.filter((item) => item.status === "待校对").length);
    } catch (error) {
      showToast(error.message, "error");
    } finally {
      setLoading(false);
    }
  }, [onPendingCountChange, showToast]);
  useEffect(() => { load(); }, [load, refreshKey]);
  const toggleCandidate = (id, kind, value) => setCandidateSelections((current) => {
    const selected = new Set(current[id]?.[kind] || []);
    if (selected.has(value)) selected.delete(value); else selected.add(value);
    return { ...current, [id]: { ...(current[id] || {}), [kind]: [...selected] } };
  });
  const act = async (item, action) => {
    const targetInternal = targetInputs[item.id] || "";
    if (action === "approve" && item.internal === "待识别" && !targetInternal.trim()) return showToast("请先填写要关联的内部型号", "error");
    const selected = candidateSelections[item.id] || {};
    try { await api.review(item.id, action, "当前用户", targetInternal, selected.oe || [], selected.uf || []); showToast(action === "approve" ? "已确认查询关系；仅勾选的候选编号已写入知识库" : "候选关系已驳回"); await refreshOverview(); await load(); }
    catch (error) { showToast(error.message, "error"); }
  };
  const pending = items.filter((item) => item.status === "待校对");
  const confirmed = items.filter((item) => item.status === "已确认");
  return (
    <div className="content-stack">
      <section className="review-banner panel"><div><span className="review-banner-icon"><ShieldCheck size={27} weight="duotone" /></span><div><h2>外网结果与正式知识库严格隔离</h2><p>确认前只用于当前任务提示；确认后才会成为所有人可用的正式映射。</p></div></div><Badge tone="warning" icon={WarningCircle}>{pending.length} 条待处理</Badge></section>
      <section className="panel review-panel">
        <div className="panel-heading"><div><h2>待校对队列</h2><p>建议优先处理高置信度且包含两个独立来源的候选</p></div><button className="secondary-button" onClick={load}><ArrowClockwise size={16} />刷新</button></div>
        {loading ? <div className="loading-block"><SpinnerGap className="spin" size={25} />正在读取校对队列</div> : pending.length ? <div className="review-list">{pending.map((item) => {
          const queryKind = /^UF/i.test(item.query) ? "UF号" : /^[A-Z]{1,4}\d/i.test(item.query) && !/-/.test(item.query) ? "查询型号" : "OE号";
          return <article className="review-card" key={item.id}>
          <div className="review-main"><div className="review-title"><div><Badge tone="purple" icon={GlobeHemisphereWest}>外网补充</Badge><h3>{item.internal}</h3><div className="review-query-highlight"><span>{queryKind}</span><strong className="mono">{item.query}</strong>{item.vehicle && <small>{item.vehicle}</small>}</div>{/[A-Z]O[A-Z0-9]/i.test(item.query) && <small className="query-character-hint">请核对型号中的字母 O 与数字 0</small>}</div><div className="confidence"><span>置信度</span><strong>{item.confidence}%</strong></div></div>
          <div className="review-platforms">{["Google", "Amazon", "eBay"].map((platform) => <Badge key={platform} tone={item.platforms?.includes(platform) ? "green" : "gray"}>{platform}{item.platforms?.includes(platform) ? " ✓" : ""}</Badge>)}</div>
          <div className="review-fields"><div><span>候选 OE 号（勾选后才写入）</span>{item.proposedOe.length ? <div className="review-candidate-options">{item.proposedOe.slice(0, 2).map((value) => <label key={value}><input type="checkbox" checked={(candidateSelections[item.id]?.oe || []).includes(value)} onChange={() => toggleCandidate(item.id, "oe", value)} /><strong className="mono">{value}</strong>{item.verifiedOe?.includes(value) && <em>跨平台</em>}</label>)}</div> : <strong className="mono">未提取</strong>}<small>{item.verifiedOe?.length ? "跨平台候选优先" : "未勾选则不会写入"}</small></div><div><span>候选 UF 号（最多 2 个，逐个确认）</span>{item.proposedUf.length ? <div className="review-candidate-options">{item.proposedUf.slice(0, 2).map((value) => <label key={value}><input type="checkbox" checked={(candidateSelections[item.id]?.uf || []).includes(value)} onChange={() => toggleCandidate(item.id, "uf", value)} /><strong className="mono">{value}</strong>{item.verifiedUf?.includes(value) && <em>跨平台</em>}</label>)}</div> : <strong className="mono">未提取</strong>}<small>{item.verifiedUf?.length ? "跨平台候选优先" : "未勾选则不会写入"}</small></div><div><span>抓取时间</span><strong>{formatDate(item.searchedAt)}</strong><small>查询编号本身将与内部型号绑定</small></div></div>
          <label className="target-model-input"><span>{item.internal === "待识别" ? "确认前指定内部型号" : "如需改选内部型号"}</span><input value={targetInputs[item.id] || ""} onChange={(event) => setTargetInputs((values) => ({ ...values, [item.id]: event.target.value }))} placeholder={item.internal === "待识别" ? "填写已有或新型号，例如 DQ910844" : `当前 ${item.internal}，留空则直接确认`} /></label>
          <div className="review-sources">{item.sources.slice(0, 3).map((source, index) => <a href={source.url} target="_blank" rel="noreferrer" key={`${source.url}-${index}`}><GlobeHemisphereWest size={16} /><span><strong>{source.title}</strong><small>{source.platform || source.provider} · {evidenceSummary(source.snippet || source.url)}</small></span><ArrowRight size={15} /></a>)}</div></div>
          <div className="review-actions"><button className="approve-button" onClick={() => act(item, "approve")}><Check size={17} />确认所选关系</button><button className="reject-button" onClick={() => act(item, "reject")}><X size={17} />驳回</button><small>查询编号必确认；候选 UF/OE 仅写入已勾选项</small></div>
        </article>})}</div> : <EmptyState icon={SealCheck} title="没有待校对项目" description="所有外网补充候选均已处理。查询缺失型号后，新候选会自动出现在这里。" />}
      </section>
      <section className="panel history-panel"><div className="panel-heading"><div><h2>二次校对确认记录</h2><p>固定格式：对外型号、OE号、UF号、确认时间</p></div><a className="secondary-button" href="/api/reviews/export"><DownloadSimple size={16} />下载确认记录</a></div><div className="table-wrap"><table><thead><tr><th>对外型号</th><th>OE号</th><th>UF号</th><th>确认时间</th></tr></thead><tbody>{confirmed.length ? confirmed.slice(0, 20).map((item) => <tr key={item.id}><td className="mono"><strong>{item.internal}</strong></td><td className="mono">{item.proposedOe.join(" / ") || "—"}</td><td className="mono">{item.proposedUf.join(" / ") || "—"}</td><td>{formatDate(item.reviewedAt)}</td></tr>) : <tr><td className="empty-table-cell" colSpan="4">暂无已确认记录；完成校对后会自动出现在这里。</td></tr>}</tbody></table></div></section>
    </div>
  );
}

function KnowledgePage({ showToast, refreshOverview }) {
  const [type, setType] = useState("models");
  const [entries, setEntries] = useState([]);
  const [versions, setVersions] = useState([]);
  const [loading, setLoading] = useState(false);
  const [dragging, setDragging] = useState(false);
  const loadVersions = useCallback(async () => { try { setVersions(await api.versions()); } catch (error) { showToast(error.message, "error"); } }, [showToast]);
  useEffect(() => { loadVersions(); }, [loadVersions]);
  const updateEntry = (id, patch) => setEntries((items) => items.map((item) => item.id === id ? { ...item, ...patch } : item));
  const choose = async (nextFiles) => {
    const selected = Array.from(nextFiles || []).filter(Boolean);
    if (!selected.length) return;
    const accepted = selected.filter(isSupportedQuoteFile);
    if (!accepted.length) return showToast("仅支持文字、Excel、PDF、PNG 和 JPG 文件", "error");
    if (accepted.length !== selected.length) showToast("已忽略不支持的文件", "error");
    const batch = accepted.map((file, index) => ({ id: `${file.name}-${file.lastModified}-${index}`, file, status: "parsing", progress: 0, prepared: null, preview: null, error: "" }));
    setEntries(batch);
    setLoading(true);
    for (const entry of batch) {
      try {
        const prepared = await prepareKnowledgeFile(entry.file, type, ({ stage, progress }) => updateEntry(entry.id, { stage, progress }));
        const preview = await api.previewKnowledge(prepared, type);
        updateEntry(entry.id, { prepared, preview, status: "ready", progress: 100, stage: "差异解析完成" });
      } catch (error) {
        updateEntry(entry.id, { status: "error", error: error.message, stage: "解析失败" });
      }
    }
    setLoading(false);
  };
  const readyEntries = entries.filter((entry) => entry.status === "ready" && entry.preview && entry.prepared);
  const summary = readyEntries.reduce((total, entry) => ({ incoming: total.incoming + entry.preview.incoming, added: total.added + entry.preview.added, updated: total.updated + entry.preview.updated, conflicts: total.conflicts + entry.preview.conflicts }), { incoming: 0, added: 0, updated: 0, conflicts: 0 });
  const apply = async () => {
    if (!readyEntries.length) return;
    setLoading(true);
    try {
      const results = [];
      for (const entry of readyEntries) {
        updateEntry(entry.id, { status: "importing", stage: "正在写入知识库" });
        results.push(await api.importKnowledge(entry.prepared, type));
        updateEntry(entry.id, { status: "imported", stage: "已导入" });
      }
      showToast(`${results.length} 个文件已更新知识库`);
      setEntries([]);
      await loadVersions();
      refreshOverview();
    }
    catch (error) { showToast(error.message, "error"); }
    finally { setLoading(false); }
  };
  const switchType = (next) => { setType(next); setEntries([]); };
  return (
    <div className="content-stack">
      <section className="knowledge-grid">
        <div className="panel update-panel">
          <div className="panel-heading"><div><h2>更新源知识库</h2><p>上传后先解析差异，确认后才生成新版本</p></div><Badge tone="green" icon={ShieldCheck}>安全更新</Badge></div>
          <div className="knowledge-tabs" role="tablist"><button role="tab" aria-selected={type === "models"} className={type === "models" ? "active" : ""} onClick={() => switchType("models")}><Tag size={18} />型号匹配表</button><button role="tab" aria-selected={type === "prices"} className={type === "prices" ? "active" : ""} onClick={() => switchType("prices")}><Database size={18} />价格表</button><button role="tab" aria-selected={type === "catalog"} className={type === "catalog" ? "active" : ""} onClick={() => switchType("catalog")}><FileMagnifyingGlass size={18} />产品目录</button></div>
          <label className={`dropzone compact ${entries.length ? "has-file" : ""} ${dragging ? "is-dragging" : ""}`} role="button" tabIndex={0} aria-label="上传知识库文件" onKeyDown={(event) => { if (["Enter", " "].includes(event.key)) { event.preventDefault(); event.currentTarget.querySelector("input")?.click(); } }} onDragEnter={(event) => { event.preventDefault(); setDragging(true); }} onDragOver={(event) => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; setDragging(true); }} onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget)) setDragging(false); }} onDrop={(event) => { event.preventDefault(); setDragging(false); choose(event.dataTransfer.files); }}><input type="file" multiple accept=".xlsx,.xls,.csv,.txt,.pdf,.png,.jpg,.jpeg,application/pdf,image/png,image/jpeg,text/plain,text/csv" onChange={(event) => { choose(event.target.files); event.target.value = ""; }} />{loading && !readyEntries.length ? <><SpinnerGap className="spin" size={28} /><strong>正在逐个解析 {entries.length} 个文件</strong><small>{type === "catalog" ? "正在读取目录内容并压缩可识别的产品图片" : "不会修改当前正式版本"}</small></> : entries.length ? <><span className="upload-icon success"><CheckCircle size={26} weight="fill" /></span><strong>已选择 {entries.length} 个文件</strong><small>{readyEntries.length} 个可导入 · {entries.filter((entry) => entry.status === "error").length} 个需处理</small><em>点击或拖拽重新选择</em></> : <><span className="upload-icon"><CloudArrowUp size={28} /></span><strong>{dragging ? "松开即可上传多个文件" : `上传新版${type === "models" ? "型号匹配表" : type === "catalog" ? "产品目录" : "价格表"}`}</strong><small>支持文字、Excel、PDF、PNG 和 JPG，可多选或批量拖拽</small><em>选择多个文件</em></>}</label>
          {entries.length > 0 && <div className="knowledge-file-list">{entries.map((entry) => <div key={entry.id} className={entry.status === "error" ? "has-error" : ""}><FileArrowUp size={17} /><span><strong>{entry.file.name}</strong><small>{entry.error || entry.stage || "等待解析"}</small></span>{entry.status === "parsing" || entry.status === "importing" ? <SpinnerGap className="spin" size={16} /> : entry.status === "error" ? <WarningCircle size={16} /> : <CheckCircle size={16} weight="fill" />}<button className="icon-button" title="移除文件" onClick={() => setEntries((items) => items.filter((item) => item.id !== entry.id))}><X size={14} /></button></div>)}</div>}
          {readyEntries.length > 0 && <div className="diff-preview"><div className="diff-heading"><span><PencilSimpleLine size={18} /></span><div><strong>{readyEntries.length} 个文件差异预览</strong><small>确认后才会逐个写入共享知识库</small></div></div><div className="diff-grid"><div><span>文件记录</span><strong>{summary.incoming}</strong></div><div><span>新增</span><strong className="green-text">+{summary.added}</strong></div><div><span>更新</span><strong className="blue-text">{summary.updated}</strong></div><div><span>冲突</span><strong className="orange-text">{summary.conflicts}</strong></div></div><button className="primary-button wide" onClick={apply} disabled={loading}>{loading ? <><SpinnerGap className="spin" />正在写入</> : <><Check size={18} />确认导入 {readyEntries.length} 个文件</>}</button></div>}
        </div>
        <div className="panel rules-panel">
          <div className="panel-heading"><div><h2>更新规则</h2><p>多人使用时避免误覆盖正式数据</p></div></div>
          <ol><li><span>1</span><div><strong>自动识别结构</strong><small>识别型号、OE、UF、价格列及产品目录图片</small></div></li><li><span>2</span><div><strong>先看差异</strong><small>新增、更新与冲突清晰分开，不直接覆盖</small></div></li><li><span>3</span><div><strong>保留版本</strong><small>每次更新记录文件、时间和操作者</small></div></li></ol>
          <div className="inline-alert info"><ShieldCheck size={20} /><span><strong>外网数据不会随文件更新混入</strong><small>仍需在“二次校对”页面单独确认。</small></span></div>
        </div>
      </section>
      <section className="panel versions-panel"><div className="panel-heading"><div><h2>知识库版本</h2><p>当前共保留 {versions.length} 条更新记录，下载内容为最新生效版本</p></div><div className="version-actions"><a className="secondary-button" href="/api/knowledge/export?type=models"><DownloadSimple size={16} />下载大厂型号匹配</a><a className="secondary-button" href="/api/knowledge/export?type=prices"><DownloadSimple size={16} />下载价格库</a><a className="secondary-button" href="/api/knowledge/export?type=catalog"><DownloadSimple size={16} />下载产品目录</a><button className="secondary-button" onClick={loadVersions}><ClockCounterClockwise size={16} />刷新记录</button></div></div><div className="table-wrap"><table><thead><tr><th>数据类型</th><th>版本</th><th>源文件</th><th>记录数</th><th>操作人</th><th>更新时间</th><th>状态</th></tr></thead><tbody>{versions.map((item) => <tr key={item.id}><td><strong>{item.type}</strong></td><td className="mono">{item.version}</td><td>{item.filename}</td><td>{item.summary?.total || item.summary?.added || "—"}</td><td>{item.operator}</td><td>{formatDate(item.created_at)}</td><td><Status value="已确认" /></td></tr>)}</tbody></table></div></section>
    </div>
  );
}

export function App() {
  const [page, setPage] = useState("overview");
  const [overview, setOverview] = useState(null);
  const [overviewLoading, setOverviewLoading] = useState(true);
  const [reviewKey, setReviewKey] = useState(0);
  const [toast, setToast] = useState({ message: "", tone: "success" });
  const showToast = useCallback((message, tone = "success") => {
    setToast({ message, tone });
    window.clearTimeout(window.__quoteToastTimer);
    window.__quoteToastTimer = window.setTimeout(() => setToast({ message: "", tone: "success" }), 4200);
  }, []);
  const refreshOverview = useCallback(async () => {
    try { setOverview(await api.overview()); }
    catch (error) { showToast(error.message, "error"); }
    finally { setOverviewLoading(false); }
  }, [showToast]);
  const syncReviewCount = useCallback((reviewCount) => {
    setOverview((current) => ({ ...(current || {}), reviewCount }));
  }, []);
  useEffect(() => { refreshOverview(); }, [refreshOverview]);
  const pageContent = useMemo(() => {
    if (page === "overview") return <Overview data={overview} loading={overviewLoading} setPage={setPage} />;
    if (page === "models") return <ModelSearch onReviewCreated={() => { setReviewKey((key) => key + 1); refreshOverview(); }} showToast={showToast} />;
    if (page === "quote") return <QuotePage showToast={showToast} refreshOverview={refreshOverview} />;
    if (page === "reviews") return <ReviewsPage refreshKey={reviewKey} showToast={showToast} refreshOverview={refreshOverview} onPendingCountChange={syncReviewCount} />;
    return <KnowledgePage showToast={showToast} refreshOverview={refreshOverview} />;
  }, [page, overview, overviewLoading, reviewKey, refreshOverview, showToast, syncReviewCount]);
  return (
    <div className="app-shell">
      <Sidebar page={page} setPage={setPage} counts={overview} />
      <main className="workspace"><Header page={page} setPage={setPage} counts={overview} />{pageContent}<footer><span>HIFLY 外贸报价系统</span><span>所有外网补充均保留来源并需二次校对</span></footer></main>
      <Toast message={toast.message} tone={toast.tone} onClose={() => setToast({ message: "", tone: "success" })} />
    </div>
  );
}

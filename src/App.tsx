import React, { useEffect, useMemo, useState } from "react";
import PRISALogo from "./assets/PRISA.png";
import {
  ArrowLeft,
  ArrowRight,
  BarChart3,
  BriefcaseBusiness,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleHelp,
  FileText,
  Layers3,
  MapPin,
  Menu,
  Plus,
  RotateCcw,
  Search,
  Settings,
  Share2,
  Sparkles,
  Target,
  TrendingUp,
  User,
  Users,
  UsersRound,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import "./App.css";

type PageKey =
  | "campaigns"
  | "summary"
  | "brief"
  | "discovery"
  | "insights"
  | "opportunities"
  | "formats"
  | "proposal"
  | "simulation"
  | "presentation"
  | "configuration";

type ExcelRow = Record<string, unknown>;

type BriefData = {
  brand: string;
  website: string;
  category: string;
  product: string;
  objective: string;
  audience: string;
  timingStart: string;
  budget: string;

  clientType: "B2B" | "B2C" | "";

  businessSector: string;
  companySize: string;

  gender: string;
  ageRanges: string[];
  socioeconomic: string[];
  regions: string[];
  context: string;
  campaignName: string;
};

type Product = {
  id: string;
  name: string;
  category: string;
  description: string;
  howWorks: string;
  application: string;
  objectives: string;
  audiences: string;
  pains: string;
  channels: string;
  tags: string;
  sheet: string;
};

type Broadcaster = {
  id: string;
  name: string;
  profile: string;
  genderProfile: string;
  ages: string;
  socioeconomic: string;
  interests: string;
  coverage: string;
};

type RegionInventory = {
  name: string;
  value: number;
};

type Insight = {
  title: string;
  description: string;
  status: "empty" | "ready";
  count: number;
};

type Opportunity = {
  id: string;
  title: string;
  description: string;
  count: number;
  kind: "pain" | "territory" | "audience" | "space";
};

type RecommendedFormat = {
  id: string;
  name: string;
  category: string;
  affinity: number;
  description: string;
  reason: string;
  source: string;
  breakdown: {
    objective: number;
    audience: number;
    context: number;
    category: number;
    keyword: number;
  };
};

type RecommendedBroadcaster = {
  id: string;
  name: string;
  affinity: number;
  profile: string;
  coverage: string;
  breakdown: {
    objective: number;
    audience: number;
    geography: number;
    context: number;
    clientType: number;
  };
  reasons: string[];
};

type Franchise = {
  id: string;
  name: string;
  category: string;
  description: string;
  application: string;
  objectives: string;
  audiences: string;
  tags: string;
  territories: string;
  sheet: string;
};

type RecommendedFranchise = {
  id: string;
  name: string;
  affinity: number;
  description: string;
  category: string;
  breakdown: {
    objective: number;
    audience: number;
    context: number;
    category: number;
    geography: number;
  };
  reasons: string[];
};

type Project = {
  id: string;
  name: string;
  status: "DRAFT" | "LIVE";
  updatedAt: string;
  brief: BriefData;
  discovery: DiscoveryData;
  selectedOpportunity: string | null;
  selectedBroadcasterId: string | null;
  selectedFranchiseId: string | null;
  selectedFormats: string[];
};

const CATALOG_FILE = "/data/Catalogo_PRISA_Normalizado_Actualizado.xlsx";
const TARIFF_FILE = "/data/Tarifario 2026 1.7 Comercial.xlsx";

const EMPTY_BRIEF: BriefData = {
  brand: "",
  website: "",
  category: "",
  product: "",
  objective: "",
  audience: "",
  timingStart: "",
  budget: "",
  clientType: "",
  businessSector: "",
  companySize: "",
  gender: "",
  ageRanges: [],
  socioeconomic: [],
  regions: [],
  context: "",
  campaignName: "",
};

function createEmptyBrief(): BriefData {
  return {
    ...EMPTY_BRIEF,
    ageRanges: [],
    socioeconomic: [],
    regions: [],
  };
}

type DiscoveryData = {
  challenge: string[];
  challengeMore: string;
  result: string[];
  resultMore: string;
  audience: string[];
  audienceMore: string;
  reaction: string[];
  reactionMore: string;
  moreInformation: string;
};

const EMPTY_DISCOVERY: DiscoveryData = {
  challenge: [],
  challengeMore: "",
  result: [],
  resultMore: "",
  audience: [],
  audienceMore: "",
  reaction: [],
  reactionMore: "",
  moreInformation: "",
};

function createEmptyDiscovery(): DiscoveryData {
  return {
    ...EMPTY_DISCOVERY,
    challenge: [],
    result: [],
    audience: [],
    reaction: [],
  };
}

const PROJECTS_STORAGE_KEY = "prisa_projects";
const SELECTED_BROADCASTER_STORAGE_KEY = "prisa_selected_broadcaster";
const SELECTED_FRANCHISE_STORAGE_KEY = "prisa_selected_franchise";

function readSelectionStorage(key: string): string | null {
  try {
    return sessionStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeSelectionStorage(key: string, value: string | null): void {
  try {
    if (value) sessionStorage.setItem(key, value);
    else sessionStorage.removeItem(key);
  } catch {
    // React mantiene el estado aunque sessionStorage no esté disponible.
  }
}

function loadSavedProjects(): Project[] {
  try {
    const raw = localStorage.getItem(PROJECTS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(
      (item): item is Project =>
        typeof item === "object" &&
        item !== null &&
        typeof (item as Project).id === "string" &&
        typeof (item as Project).name === "string" &&
        typeof (item as Project).brief === "object"
    );
  } catch {
    return [];
  }
}


const NAV_ITEMS: Array<{
  key: PageKey;
  label: string;
  icon: React.ElementType;
}> = [
  { key: "campaigns", label: "Campañas", icon: BriefcaseBusiness },
  { key: "summary", label: "Resumen", icon: BarChart3 },
  { key: "brief", label: "Brief / Inputs", icon: FileText },
  {
    key: "discovery",
    label: "Descubrimiento",
    icon: CircleHelp,
  },
  { key: "insights", label: "Insights", icon: Sparkles },
  { key: "opportunities", label: "Oportunidades", icon: Target },
  { key: "formats", label: "Formatos", icon: Layers3 },
  { key: "proposal", label: "Propuesta", icon: FileText },
  { key: "simulation", label: "Simulación", icon: TrendingUp },
  { key: "presentation", label: "Presentación", icon: BarChart3 },
];

function normalize(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}

function findColumn(row: ExcelRow, candidates: string[]): string {
  const keys = Object.keys(row);

  for (const candidate of candidates) {
    const exact = keys.find(
      (key) => normalize(key) === normalize(candidate)
    );
    if (exact) return exact;
  }

  for (const candidate of candidates) {
    const partial = keys.find((key) =>
      normalize(key).includes(normalize(candidate))
    );
    if (partial) return partial;
  }

  return keys[0] ?? "";
}

function rowValue(row: ExcelRow, candidates: string[]): string {
  const key = findColumn(row, candidates);
  return text(row[key]);
}

function parseNumber(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const raw = String(value ?? "")
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(",", ".");

  const result = Number(raw);
  return Number.isFinite(result) ? result : 0;
}

async function readWorkbook(url: string): Promise<ExcelRow[]> {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`No se pudo cargar ${url}`);
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const rows: ExcelRow[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json<ExcelRow>(sheet, {
      defval: "",
      raw: false,
    });

    data.forEach((row) => {
      rows.push({ ...row, __sheet: sheetName });
    });
  });

  return rows;
}

async function readGeolocationWorkbook(url: string): Promise<ExcelRow[]> {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`No se pudo cargar ${url}`);
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets["Geofocalización"];

  if (!sheet) return [];

  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });

  const headerIndex = matrix.findIndex((row) =>
    Array.isArray(row) &&
    row.some((cell) => normalize(cell) === "formato publicitario")
  );

  if (headerIndex < 0) return [];

  const headerRow = (matrix[headerIndex] as unknown[]).map((cell) =>
    text(cell)
  );

  const result: ExcelRow[] = [];

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] as unknown[];
    if (!row || !row.some((cell) => text(cell) !== "")) continue;

    const object: ExcelRow = {};

    headerRow.forEach((header, columnIndex) => {
      if (!header) return;
      object[header] = row[columnIndex] ?? "";
    });

    result.push(object);
  }

  return result;
}

function tokenSet(value: unknown): Set<string> {
  return new Set(
    normalize(value)
      .split(/\s+/)
      .map((word) => word.replace(/[^a-z0-9áéíóúüñ]/gi, ""))
      .filter((word) => word.length > 3)
  );
}

function overlapScore(source: unknown, target: unknown): number {
  const sourceTokens = tokenSet(source);
  const targetTokens = tokenSet(target);
  if (!sourceTokens.size || !targetTokens.size) return 0;

  let matches = 0;
  sourceTokens.forEach((token) => {
    if (targetTokens.has(token)) matches += 1;
  });

  return Math.min(100, Math.round((matches / sourceTokens.size) * 100));
}

function capScore(value: number): number {
  return Math.max(0, Math.min(98, Math.round(value)));
}

function ReferenceStyles() {
  return (
    <style>{`
      .app-reference { min-height: 100vh; background:#080a10; color:#eef2f7; }
      .reference-sidebar { width:187px !important; background:#121621 !important; }
      .reference-main { margin-left:187px !important; width:calc(100% - 187px) !important; }
      .reference-header { height:62px; border-bottom:1px solid #262e40; background:#080a10; display:flex; align-items:center; justify-content:space-between; padding:0 24px; }
      .reference-head-left { display:flex; align-items:center; gap:12px; min-width:0; }
      .reference-head-search { width:332px; height:37px; display:flex; align-items:center; gap:8px; background:#171d2a; border:1px solid #293249; border-radius:9px; padding:0 11px; }
      .reference-head-search svg{color:#6b758a; flex:none}
      .reference-head-search input{width:100%; background:none; border:0; outline:0; color:#eef2f7; font-size:11px;}
      .reference-project { height:37px; display:flex; align-items:center; gap:9px; padding:0 12px; border:1px solid #293249; border-radius:9px; background:#171d2a; color:#dfe4ed; font-size:10px; }
      .reference-live { color:#25d978; font-size:9px; font-weight:800; }
      .reference-head-actions{display:flex;gap:5px;align-items:center}
      .reference-icon{width:31px;height:31px;display:grid;place-items:center;color:#778297;background:transparent;border:0;border-radius:8px}
      .reference-icon:hover{background:#171d2a;color:#fff}
      .reference-content{padding:27px 28px 48px; max-width:1230px; margin:0 auto;}
      .reference-title-row{position:relative;display:flex;justify-content:space-between;align-items:flex-start;gap:18px;margin-bottom:18px}
      .reference-title-row > div:first-child{flex:1;text-align:center}
      .reference-title-row h1{margin:5px 0;color:#f1f4f8;font-size:26px;letter-spacing:-.04em}
      .reference-title-row p{margin:0;color:#8691a8;font-size:10px;line-height:1.65}
      .reference-title-row .reference-complete{position:absolute;right:0;top:0}
      .discovery-more-button:hover{background:transparent!important;color:#ff5bb8!important}
      .reference-eyebrow{color:#ff2b9d;font-size:8px;font-weight:800;letter-spacing:.16em;text-transform:uppercase}
      .reference-card{background:#151a27;border:1px solid #293248;border-radius:14px;}
      .reference-info{padding:17px 19px;display:flex;gap:10px;align-items:flex-start;color:#9ba5b8;font-size:10px;line-height:1.6}
      .reference-info svg{color:#4292ff;flex:none;margin-top:1px}
      .reference-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:15px 0}
      .reference-kpi{padding:16px;min-height:104px;background:#151a27;border:1px solid #293248;border-radius:12px}
      .reference-kpi small{display:block;color:#7d889d;font-size:8px;margin-bottom:8px}
      .reference-kpi strong{display:block;color:#f0f3f8;font-size:19px;line-height:1.1}
      .reference-kpi span{display:block;color:#68748b;font-size:8px;margin-top:5px}
      .reference-grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
      .reference-panel{padding:17px;border:1px solid #293248;border-radius:14px;background:#121722}
      .reference-panel h2{margin:5px 0 0;font-size:15px;color:#eef2f7}
      .reference-panel-head{display:flex;align-items:flex-start;justify-content:space-between;gap:10px}
      .reference-map{position:relative;height:235px;margin-top:14px;border-radius:10px;border:1px solid #242d40;background:radial-gradient(circle at center,#171d2a 0%,#101520 72%);overflow:hidden}
      .reference-map:before,.reference-map:after{content:"";position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);background:#242e43}
      .reference-map:before{width:78%;height:1px}.reference-map:after{width:1px;height:78%}
      .reference-node{position:absolute;width:70px;height:70px;border-radius:50%;display:flex;flex-direction:column;align-items:center;justify-content:center;border:1px solid #b51172;background:#32172f;box-shadow:0 0 35px rgba(236,15,130,.1);z-index:2}
      .reference-node span{font-size:7px;color:#8994a8}.reference-node strong{font-size:17px;color:#ff2b9d;margin-top:2px}
      .reference-node.center{left:50%;top:50%;transform:translate(-50%,-50%);width:78px;height:78px;background:#ec0f82;border-color:#ff53b1;color:#fff}.reference-node.center strong{color:#fff;font-size:11px}.reference-node.center span{color:#ffe4f3}
      .reference-node.n1{left:11%;top:13%}.reference-node.n2{right:11%;top:18%}.reference-node.n3{left:12%;bottom:13%}.reference-node.n4{right:12%;bottom:15%}
      .reference-list{display:flex;flex-direction:column;margin-top:10px}
      .reference-list-row{display:grid;grid-template-columns:34px minmax(0,1fr) auto;gap:10px;align-items:center;padding:11px 0;border-bottom:1px solid #242c3b}
      .reference-list-row:last-child{border-bottom:0}
      .reference-list-icon{width:31px;height:31px;display:grid;place-items:center;border-radius:8px;background:#2b1730;color:#ff2b9d}
      .reference-list-row strong{display:block;font-size:10px;color:#eaf0f7}.reference-list-row span{display:block;margin-top:3px;color:#708097;font-size:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .reference-score{color:#ff2b9d;font-size:10px;font-weight:800}
      .reference-tabs{display:flex;gap:4px;background:#171d2a;padding:3px;border-radius:8px;width:max-content;margin:0 0 14px}
      .reference-tab{padding:8px 12px;background:transparent;border:0;border-radius:6px;color:#78849a;font-size:9px}.reference-tab.active{background:#263047;color:#fff}
      .reference-brief-card{max-width:775px;margin:0 auto;padding:18px}
      .reference-brief-head{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:1px solid #2a3244;padding-bottom:14px;margin-bottom:12px;position:relative}.reference-brief-head>div:first-child{flex:1;text-align:center}.reference-brief-head .reference-complete{position:absolute;right:0;top:0}
      .reference-complete{display:flex;flex-direction:column;align-items:flex-end}.reference-complete strong{font-size:21px}.reference-complete span{font-size:8px;color:#78849a}
      .reference-progress{height:5px;background:#20283a;border-radius:99px;overflow:hidden;margin-bottom:15px}.reference-progress>div{height:100%;background:#ec0f82}
      .reference-brief-field{margin-bottom:10px}.reference-brief-field label{display:block;color:#a6afc1;font-size:9px;margin:0 0 5px}.reference-brief-field input,.reference-brief-field textarea{width:100%;background:#181e2d;color:#e9eef5;border:1px solid #293248;border-radius:8px;padding:9px 10px;outline:0;font-size:10px}.reference-brief-field textarea{min-height:65px;resize:vertical}.reference-brief-field small{display:block;margin-top:4px;color:#66728a;font-size:7px}
      .reference-chip-wrap{display:flex;flex-wrap:wrap;gap:6px;margin-top:5px}.reference-chip{padding:6px 8px;background:#121827;border:1px solid #2a3348;border-radius:7px;color:#9aa5b8;font-size:8px}.reference-chip.active{background:#36142f;border-color:#ec0f82;color:#fff}
      .reference-brief-footer{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;border-top:1px solid #293248;margin-top:12px;padding-top:14px}.reference-info-block span{display:block;color:#6f7a91;font-size:7px;text-transform:uppercase;letter-spacing:.08em}.reference-info-block strong{display:block;margin-top:6px;color:#eef2f7;font-size:9px}.reference-info-block .next{color:#ff2b9d;display:flex;align-items:center;gap:2px}
      .reference-insights-grid{display:grid;gap:11px}.reference-insight-card{padding:16px 18px;background:#151a27;border:1px solid #293248;border-radius:13px}.reference-insight-card .row{display:flex;justify-content:space-between;align-items:center}.reference-insight-card h3{margin:0;font-size:12px}.reference-insight-card p{margin:8px 0 0;color:#718097;font-size:10px}.reference-warning{padding:13px 15px;border:1px solid #805d16;background:#231d12;color:#e1b946;border-radius:11px;font-size:9px;line-height:1.55;margin-top:11px}
      .reference-op-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:13px}.reference-op-kpi{text-align:center;padding:20px;background:#151a27;border:1px solid #293248;border-radius:13px}.reference-op-kpi .big{display:block;font-size:21px;font-weight:800;margin-top:8px}.reference-op-kpi h3{margin:4px 0 3px;font-size:10px}.reference-op-kpi small{color:#6f7b91;font-size:8px}.reference-accordion{display:flex;flex-direction:column;gap:10px}.reference-accordion-row{background:#151a27;border:1px solid #293248;border-radius:11px;overflow:hidden}.reference-accordion-button{width:100%;border:0;background:transparent;color:#eef2f7;padding:12px 14px;display:flex;align-items:center;gap:9px;text-align:left}.reference-accordion-button span{flex:1;font-size:10px;font-weight:700}.reference-accordion-body{padding:0 14px 13px;color:#718098;font-size:9px}
      .reference-formats-top{display:flex;gap:10px;align-items:center;padding:14px 16px;border:1px solid #293248;border-radius:12px;background:#151a27;margin-bottom:12px}.reference-summary-pill{display:flex;gap:6px;color:#778398;font-size:9px}.reference-format-tabs{display:flex;gap:3px;background:#171d2a;padding:4px;border-radius:8px;width:max-content;margin-bottom:12px}.reference-format-list{display:flex;flex-direction:column;gap:8px}.reference-format-row{display:grid;grid-template-columns:44px 44px minmax(180px,1fr) 120px 78px;gap:12px;align-items:center;justify-items:center;text-align:center;padding:11px 12px;background:#151a27;border:1px solid #293248;border-radius:11px}.reference-format-row.selected{border-color:#ec0f82}.reference-radio{width:18px;height:18px;border:1px solid #3a455e;border-radius:50%;display:grid;place-items:center}.reference-radio.active:after{content:"";width:8px;height:8px;border-radius:50%;background:#ec0f82}.reference-emitter-icon{width:31px;height:31px;display:grid;place-items:center;background:#172843;color:#4e98ff;border-radius:8px}.reference-format-row strong{font-size:10px;white-space:nowrap}.reference-format-row small{display:block;color:#6f7b91;font-size:8px;margin-top:3px}.reference-why{color:#ff2b9d;font-size:9px;text-align:center;white-space:nowrap;min-width:68px}.reference-progress-bar{height:5px;background:#20283a;border-radius:99px;overflow:hidden}.reference-progress-bar>div{height:100%;background:#ec0f82}
      .reference-proposal-list{display:flex;flex-direction:column;gap:10px}.reference-proposal-header{padding:18px 19px;display:flex;align-items:center;justify-content:space-between;background:#151a27;border:1px solid #293248;border-radius:14px}.reference-proposal-section{padding:16px 18px;position:relative;background:#151a27;border:1px solid #293248;border-radius:13px}.reference-proposal-section h3{margin:0;font-size:11px}.reference-proposal-section p{margin:7px 0 0;color:#718098;font-size:9px}.reference-edit{position:absolute;right:15px;top:15px;color:#ff2b9d;font-size:8px;background:none;border:0}
      .reference-sim-chart{padding:18px;background:#151a27;border:1px solid #293248;border-radius:14px;margin-top:14px}.reference-chart{height:300px;display:flex;flex-direction:column;justify-content:space-around;padding:10px 0}.reference-bar-row{display:grid;grid-template-columns:115px minmax(0,1fr);gap:10px;align-items:center}.reference-bar-label{color:#8f9aaf;font-size:8px;text-align:right}.reference-bar-track{height:18px;background:#20273a;border-radius:3px;overflow:hidden}.reference-bar{height:100%;background:#2b3244;border-radius:3px}.reference-sim-stats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}.reference-stat{padding:15px;background:#151a27;border:1px solid #293248;border-radius:13px}.reference-stat span{color:#78849a;font-size:8px;display:block}.reference-stat strong{display:block;font-size:17px;margin-top:5px}
      .reference-presentation-shell{background:#080a10;min-height:calc(100vh - 62px);padding:26px 28px 42px}.reference-presentation-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:13px}.reference-presentation-top h2{margin:0;font-size:15px}.reference-presentation-meta{display:flex;align-items:center;gap:9px}.reference-slide-counter{color:#a0a9b9;font-size:9px}.reference-presentation-warning{padding:12px 14px;border:1px solid #805d16;background:#231d12;color:#e2b744;border-radius:10px;font-size:9px;margin-bottom:13px}.reference-slide{max-width:770px;aspect-ratio:16/9;background:radial-gradient(circle at 85% 55%,rgba(236,15,130,.16),transparent 24%),#111520;border:1px solid #293248;border-radius:14px;margin:0 auto;position:relative;overflow:hidden;padding:48px}.reference-slide:before{content:"";position:absolute;left:0;top:0;width:92px;height:4px;background:#ec0f82}.reference-slide-logo{display:flex;align-items:center;gap:7px;margin-bottom:31px}.reference-slide-logo img{width:85px;height:31px;object-fit:contain;object-position:left center}.reference-slide-label{color:#ff2b9d;font-size:8px;font-weight:800;letter-spacing:.16em}.reference-slide h1{font-size:36px;letter-spacing:-.05em;margin:12px 0 7px}.reference-slide p{color:#8f99ae;font-size:10px}.reference-thumbnails{display:grid;grid-template-columns:repeat(6,1fr);gap:7px;max-width:840px;margin:13px auto 0}.reference-thumb{min-height:54px;padding:8px;background:#121827;border:1px solid #293248;border-radius:7px;color:#a9b1c1;text-align:left;font-size:8px}.reference-thumb.active{border-color:#ec0f82;color:#fff}.reference-source-note{max-width:840px;margin:11px auto 0;color:#69758c;font-size:8px;line-height:1.5}.reference-presentation-footer{max-width:840px;margin:16px auto 0;border-top:1px solid #293248;padding-top:15px}.reference-presentation-footer button{display:inline-flex}
      .presentation-top-centered{align-items:flex-start}
      .presentation-top-centered>div:first-child{text-align:center;flex:1}
      .presentation-top-centered .reference-presentation-meta{position:absolute;right:28px;top:24px}
      .presentation-centered-content{text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:flex-start;padding-top:52px}
      .story-list-centered{text-align:left;width:100%;max-width:620px}
      .slide-centered-extra{max-width:640px;margin:20px auto 0!important;color:#8f99ae!important;font-size:9px!important;line-height:1.6!important}
      .insight-slide-centered{width:100%;max-width:640px;display:flex;flex-direction:column;align-items:center}
      .concept-box-centered{width:100%;text-align:center}
      .audience-slide-centered-grid{display:grid;grid-template-columns:1fr 1fr;gap:22px;width:100%;max-width:680px;margin-top:24px}
      .audience-center-card{padding:18px;border:1px solid #293248;background:#151a27;border-radius:11px;text-align:center;min-height:150px}
      .audience-center-card p{margin:8px 0 15px!important}
      .affinity-list-centered{justify-content:center}
      .journey-area-centered{text-align:center;width:100%;display:flex;flex-direction:column;align-items:center}
      .journey-centered{justify-content:center}
      .presentation-centered-content .solution-flow{width:100%;max-width:760px}
      .presentation-centered-content .flow-card{text-align:center}
      @media(max-width:1050px){.reference-kpis{grid-template-columns:1fr 1fr}.reference-grid2,.reference-op-grid{grid-template-columns:1fr}.reference-format-row{grid-template-columns:40px 40px minmax(150px,1fr) 90px 74px}.reference-brief-footer{grid-template-columns:1fr 1fr 1fr}}
      @media(max-width:760px){
        .presentation-top-centered .reference-presentation-meta{position:static}
        .audience-slide-centered-grid{grid-template-columns:1fr}
.reference-sidebar{width:76px!important}.reference-main{margin-left:76px!important;width:calc(100% - 76px)!important}.reference-head-search{width:220px}.reference-project{display:none}.reference-content{padding:20px 14px 35px}.reference-kpis{grid-template-columns:1fr 1fr}.reference-grid2,.reference-op-grid{grid-template-columns:1fr}.reference-brief-card{max-width:none}.reference-presentation-shell{padding:17px 13px 32px}.reference-slide{padding:25px}.reference-slide h1{font-size:28px}.reference-thumbnails{grid-template-columns:repeat(4,1fr)}.reference-format-row{grid-template-columns:34px 34px minmax(110px,1fr) 72px 64px;gap:7px;padding:10px 8px}.reference-format-row .ref-score{display:block}.reference-format-row .reference-why{display:block;white-space:nowrap;font-size:9px!important}}

      /* --- SIMULACIÓN: replica de la referencia izquierda --- */
      .simulation-reference-page{max-width:1180px}
      .simulation-info-banner{max-width:100%;font-size:10px}
      .simulation-kpis{margin-top:15px}
      .simulation-kpi{min-height:104px}
      .simulation-kpi strong{font-size:18px}
      .simulation-chart-card{margin-top:14px}
      .simulation-chart-title{color:#9eabc0;font-size:9px;margin-bottom:2px}
      .simulation-chart{height:auto;min-height:330px;padding:7px 0 11px;gap:2px}
      .simulation-bar-row{grid-template-columns:150px minmax(0,1fr) 85px;gap:10px}
      .simulation-bar-label{font-size:8px}
      .simulation-bar-track{height:18px;background:#20273a}
      .simulation-bar{background:#2b3244}
      .simulation-bar-value{color:#667289;font-size:8px;text-align:right}
      .simulation-footnote{margin:4px 0 0;color:#6e7a91;font-size:8px}

      /* --- PRESENTACIÓN: 6 láminas de referencia --- */
      .presentation-subtitle{margin:6px auto 0;color:#78849a;font-size:10px;line-height:1.6;max-width:760px;text-align:center}
      .reference-slide{max-width:840px;min-height:472px;aspect-ratio:16/9;padding:52px 56px;background:radial-gradient(circle at 85% 55%,rgba(236,15,130,.16),transparent 24%),#111520}
      .reference-slide-logo{display:none}
      .reference-slide-logo img{width:85px;height:31px;object-fit:contain;object-position:left center}
      .reference-slide-logo span{color:#9aa5b8;font-size:8px}
      .presentation-slide-content{position:relative;z-index:2;min-height:300px}
      .reference-slide h1{font-size:39px;margin:12px 0 8px}
      .reference-slide h2{font-size:24px;margin:8px 0 0;letter-spacing:-.035em}
      .reference-slide p{margin:0;color:#93a0b5;font-size:10px;line-height:1.65}
      .slide-large-subtitle{font-size:14px!important;color:#b8c1d0!important}
      .presentation-slide-content>small{display:block;margin-top:28px;color:#5f6b81;font-size:8px}
      .story-list{display:flex;flex-direction:column;gap:17px;margin-top:34px;max-width:610px}
      .story-item{display:grid;grid-template-columns:25px minmax(0,1fr);gap:11px;align-items:start}
      .story-number{color:#2d8cff;font-size:9px;font-weight:800;padding-top:2px}
      .story-red{color:#ff2b9d}.story-green{color:#29d875}
      .story-item strong{display:block;color:#eff2f7;font-size:11px}
      .story-item p{margin-top:4px}
      .insight-slide-body{margin-top:42px}
      .insight-slide-body>span,.concept-box span,.slide-caption{display:block;color:#6f7b91;font-size:8px;text-transform:uppercase;letter-spacing:.08em}
      .insight-slide-body h2{max-width:520px;font-size:25px;margin-top:8px}
      .concept-box{margin-top:26px;padding:16px 18px;background:#34152f;border:1px solid #a90d66;border-radius:10px;max-width:520px}
      .concept-box strong{display:block;color:#f3dff0;margin-top:7px;font-size:11px}
      .audience-slide-grid{display:grid;grid-template-columns:1fr 1fr;gap:38px;margin-top:21px}
      .audience-slide-grid p{margin:6px 0 15px}
      .split-meter{height:6px;display:flex;background:#20283a;border-radius:99px;overflow:hidden;margin-top:8px}
      .split-meter>div:first-child{background:#ff2b9d}
      .split-meter>div:last-child{background:#3086ff}
      .split-labels{display:flex;gap:15px;font-size:8px;margin:6px 0 14px;color:#b7c0ce}
      .split-labels span:first-child::before{content:"•";color:#ff2b9d;margin-right:4px}
      .split-labels span:last-child::before{content:"•";color:#3086ff;margin-right:4px}
      .mini-bar{display:grid;grid-template-columns:49px minmax(0,1fr) 29px;gap:7px;align-items:center;margin:7px 0}
      .mini-bar>span{color:#8590a5;font-size:8px}
      .mini-bar>div{height:5px;background:#20283a;border-radius:99px;overflow:hidden}
      .mini-bar i{display:block;height:100%;background:#ec0f82;border-radius:99px}
      .mini-bar strong{font-size:8px;color:#c5cdd9;text-align:right}
      .region-bar{grid-template-columns:105px minmax(0,1fr) 29px}.region-bar>span{text-align:left}
      .affinity-list{display:flex;width:100%;gap:8px;flex-wrap:wrap;justify-content:center!important;align-items:center;margin:10px auto 0!important;max-width:520px}.affinity-list span{padding:5px 7px;border:1px solid #30394e;background:#111826;border-radius:7px;color:#a9b4c5;font-size:8px}
      .slide-bottom-note{position:absolute;left:0;bottom:0!important;margin:0!important;color:#606d83!important;font-size:8px!important}
      .solution-flow{display:grid;grid-template-columns:1fr auto 1fr auto 1fr auto 1fr;align-items:center;gap:7px;margin-top:46px}
      .flow-card{min-height:112px;padding:13px;border:1px solid #293248;background:#151a27;border-radius:10px}
      .flow-card small{display:block;margin-top:7px;color:#7d889d;font-size:8px}.flow-card strong{display:block;color:#eef2f7;font-size:10px;margin-top:5px}.flow-card p{font-size:8px;margin-top:11px}
      .flow-icon{width:25px;height:25px;display:grid;place-items:center;border-radius:7px;background:#172843;color:#4d97ff;font-size:11px}
      .flow-icon.green{background:#153024;color:#2bdd7b}.flow-icon.pink{background:#34162f;color:#ff2b9d}.flow-icon.yellow{background:#30291a;color:#f3bd38}
      .flow-arrow{color:#6d7890}
      .journey-area{margin-top:66px}
      .journey{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin:12px 0 31px}
      .journey span{padding:8px 11px;background:#171e2d;border:1px solid #2b3448;border-radius:7px;color:#dce3ed;font-size:9px}.journey svg{color:#6b768d}
      .measure-text{margin-top:11px!important}
      .final-slide-content{display:grid;grid-template-columns:1fr 1fr;gap:42px;padding-top:28px;align-content:start}
      .final-label{display:block;font-size:8px;font-weight:800;letter-spacing:.1em;text-transform:uppercase}.final-label.orange{color:#f0b52b}.final-label.green{color:#27d978}
      .final-slide-content h3{margin:8px 0;color:#edf1f6;font-size:11px}.final-slide-content p{font-size:9px}
      .final-message{grid-column:1/-1;align-self:end;margin-top:80px;color:#edf2f7;font-size:25px;font-weight:700;letter-spacing:-.035em}
      .slide-index{position:absolute;top:16px;right:18px;color:#6e788d;font-size:8px}
      .reference-thumb span{display:block;color:#6f7a91;margin-bottom:5px}
      /* ===== AJUSTES FINALES DE UI ===== */
      .reference-content, .reference-presentation-shell{font-size:11px;}
      .sidebar-nav .nav-item{font-size:11px!important;}
      .sidebar-section-label{font-size:9px!important;}
      .meeting{font-size:11px!important;}
      .sidebar-user{justify-content:center!important;align-items:center!important;flex-direction:column!important;text-align:center!important;gap:6px!important;}
      .sidebar-user .avatar{margin:0 auto!important;}
      .sidebar-user strong{font-size:10px!important;}
      .sidebar-user span{font-size:8px!important;}
      .reference-head-search input{font-size:12px!important;}
      .reference-project{font-size:11px!important;}
      .reference-title-row h1{font-size:29px!important;}
      .reference-title-row p{font-size:11px!important;}
      .reference-eyebrow{font-size:9px!important;}
      .reference-brief-field label{font-size:10px!important;}
      .reference-brief-field input,.reference-brief-field textarea{font-size:11px!important;}
      .reference-brief-field small{font-size:8px!important;}
      .reference-chip{font-size:9px!important;padding:7px 9px!important;}
      .reference-tab{font-size:10px!important;padding:9px 13px!important;}
      .reference-format-row strong{font-size:11px!important;white-space:nowrap;}
      .reference-format-row small{font-size:9px!important;text-align:center;}
      .reference-why{
        color:#ff2b9d!important;
        background:transparent!important;
        border:1px solid transparent!important;
        box-shadow:none!important;
        border-radius:7px!important;
        padding:6px 9px!important;
        font-size:10px!important;
        font-weight:700!important;
        line-height:1.2!important;
        cursor:pointer!important;
        transition:background .15s ease,border-color .15s ease,color .15s ease!important;
      }
      .reference-why:hover{background:#35162f!important;border-color:#8d185e!important;color:#fff!important;}
      .discovery-more-button{padding:5px 0!important;border:0!important;}
      .reference-presentation-shell{text-align:center;}
      .reference-presentation-top{justify-content:center!important;position:relative;top:-4px;}
      .reference-presentation-top>div:first-child{flex:1;text-align:center;}
      .reference-presentation-top .reference-presentation-meta{position:absolute;right:28px;}
      .reference-format-row > div:nth-child(3), .reference-format-row > div:nth-child(4){width:100%;text-align:center;}
      .reference-format-row > .reference-why{justify-self:center;}
      .reference-slide{text-align:center!important;}
      .presentation-centered-logo{justify-content:center!important;}
      .presentation-centered-content{display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center!important;min-height:340px;width:100%;position:relative;top:-22px;}
      .presentation-centered-content h1,.presentation-centered-content h2,.presentation-centered-content p,.presentation-centered-content small{max-width:760px;text-align:center;}
      .story-list{margin:24px auto 0!important;width:100%;max-width:690px;text-align:center!important;}
      .story-item{grid-template-columns:28px minmax(0,1fr)!important;text-align:center!important;}
      .story-item>div{text-align:center!important;}
      .insight-slide-body{width:100%;text-align:center!important;}
      .insight-slide-body h2{margin-left:auto!important;margin-right:auto!important;text-align:center!important;}
      .concept-box{margin-left:auto!important;margin-right:auto!important;text-align:center!important;}
      .audience-slide-grid{width:100%;max-width:690px;margin:18px auto 0;text-align:center!important;}
      .audience-slide-grid>div{text-align:center!important;}
      .mini-bar{grid-template-columns:70px minmax(0,1fr) 40px!important;}
      .mini-bar>span{text-align:right!important;}
      .region-bar{grid-template-columns:115px minmax(0,1fr) 40px!important;}
      .solution-flow{width:100%;max-width:760px;justify-content:center;margin-left:auto;margin-right:auto;}
      .flow-card{text-align:center;}
      .journey{justify-content:center!important;max-width:720px;margin-left:auto;margin-right:auto;}
      .measure-text{text-align:center!important;}
      .final-slide-content{text-align:center!important;max-width:700px;margin-left:auto;margin-right:auto;}
      .final-slide-content>div{text-align:center!important;}
      .final-message{text-align:center!important;}
      /* ===== PRESENTACIÓN: centrado vertical real de las 6 láminas ===== */
      /* ===== MOBILE: OVERRIDE DE REFERENCIA (se declara después de los estilos anteriores) ===== */
      @media screen and (max-width:760px){
        html,body,#root{width:100%!important;min-width:0!important;max-width:100%!important;overflow-x:hidden!important;}
        .app-reference{width:100%!important;min-width:0!important;overflow-x:hidden!important;}
        .app-reference .reference-main{margin-left:0!important;width:100%!important;max-width:100%!important;}
        .app-reference .reference-sidebar{position:fixed!important;left:0!important;top:0!important;bottom:0!important;width:280px!important;min-width:280px!important;max-width:86vw!important;height:100dvh!important;margin:0!important;z-index:100!important;transform:translate3d(-110%,0,0)!important;visibility:hidden!important;opacity:0!important;pointer-events:none!important;box-sizing:border-box!important;transition:transform .22s ease,visibility 0s linear .22s,opacity .18s ease!important;}
        .app-reference .reference-sidebar.mobile-open{transform:translate3d(0,0,0)!important;visibility:visible!important;opacity:1!important;pointer-events:auto!important;}
        .app-reference .mobile-sidebar-overlay{display:block!important;position:fixed!important;inset:0!important;width:100%!important;height:100%!important;padding:0!important;margin:0!important;border:0!important;background:rgba(0,0,0,.62)!important;z-index:90!important;}
        .app-reference .mobile-menu-button{display:grid!important;width:42px!important;height:42px!important;flex:0 0 42px!important;place-items:center!important;}
        .app-reference .reference-header{width:100%!important;box-sizing:border-box!important;padding:0 10px!important;}
        .app-reference .reference-head-left{width:100%!important;min-width:0!important;gap:8px!important;}
        .app-reference .reference-head-search{flex:1!important;min-width:0!important;width:auto!important;}
        .app-reference .reference-head-actions .reference-icon{display:none!important;}
        .app-reference .reference-head-actions{flex:0 0 auto!important;}
        .app-reference .reference-content{width:100%!important;max-width:none!important;box-sizing:border-box!important;padding:18px 10px 34px!important;margin:0!important;}
      }

      .reference-slide{
        position:relative!important;
        overflow:hidden!important;
      }
      .presentation-centered-content{
        position:absolute!important;
        inset:0!important;
        box-sizing:border-box!important;
        width:100%!important;
        height:100%!important;
        min-height:0!important;
        padding:68px 56px 58px!important;
        margin:0!important;
        top:0!important;
        left:0!important;
        display:flex!important;
        flex-direction:column!important;
        align-items:center!important;
        justify-content:center!important;
        text-align:center!important;
      }
      .presentation-centered-content .story-list,
      .presentation-centered-content .insight-slide-body,
      .presentation-centered-content .audience-slide-grid,
      .presentation-centered-content .audience-slide-centered-grid,
      .presentation-centered-content .solution-flow,
      .presentation-centered-content .journey-area,
      .presentation-centered-content .final-slide-content{
        margin-top:18px!important;
        margin-bottom:0!important;
      }
      .presentation-centered-content .story-list{
        max-width:680px!important;
      }
      .presentation-centered-content .solution-flow{
        max-width:760px!important;
      }
      .presentation-centered-content .affinity-list,
      .presentation-centered-content .affinity-list-centered{
        width:100%!important;
        max-width:520px!important;
        justify-content:center!important;
        align-items:center!important;
        margin-left:auto!important;
        margin-right:auto!important;
      }
      .presentation-centered-content .flow-card{
        min-height:96px!important;
      }

      /* ===== RESUMEN: ajuste visual basado en el diseño de referencia ===== */
      .summary-progress-heading{
        text-align:center;
        margin:0 0 9px;
      }
      .summary-progress-card{
        padding:17px 20px!important;
        margin-bottom:19px!important;
      }
      .summary-progress-track{
        display:flex;
        align-items:flex-start;
        justify-content:center;
        width:100%;
      }
      .summary-step{
        flex:1;
        max-width:190px;
        border:0;
        background:transparent;
        color:#778399;
        padding:0 5px;
        text-align:center;
        cursor:pointer;
      }
      .summary-step-number{
        width:26px;
        height:26px;
        border-radius:50%;
        margin:0 auto 7px;
        display:grid;
        place-items:center;
        background:#111827;
        border:1px solid #2a3348;
        color:#778399;
        font-size:9px;
        font-weight:800;
      }
      .summary-step.done .summary-step-number{
        background:#ec0f82;
        border-color:#ec0f82;
        color:#fff;
      }
      .summary-step strong{
        display:block;
        color:#eef2f7;
        font-size:9px;
      }
      .summary-step.pending strong{color:#778399}
      .summary-step small{
        display:block;
        margin-top:4px;
        color:#ec0f82;
        font-size:7px;
      }
      .summary-step.pending small{color:#5e6a80}
      .summary-step-line{
        flex:0 0 48px;
        height:1px;
        margin-top:13px;
        background:#293248;
      }
      .summary-step-line.done{background:#7e1f5a}

      .summary-section-header{
        display:flex;
        justify-content:space-between;
        align-items:flex-end;
        gap:16px;
        margin:0 0 11px;
      }
      .summary-section-header>div{
        flex:1;
        text-align:center;
      }
      .summary-section-header h1{
        margin:5px 0 0;
        color:#f1f4f8;
        font-size:26px!important;
        letter-spacing:-.04em;
      }

      .summary-map-card{
        padding:10px!important;
        margin-bottom:20px!important;
      }
      .summary-map{
        position:relative;
        min-height:330px;
        border:1px solid #242d40;
        border-radius:10px;
        background:radial-gradient(circle at center,#171d2a 0%,#101520 74%);
        overflow:hidden;
      }
      .summary-map-side{
        position:absolute;
        top:18px;
        bottom:18px;
        width:31%;
        display:flex;
        flex-direction:column;
        justify-content:space-between;
        z-index:2;
      }
      .summary-map-side-left{left:18px}
      .summary-map-side-right{right:18px}
      .summary-map-node{
        display:flex;
        gap:7px;
        align-items:center;
        min-height:48px;
        padding:7px 8px;
        background:#121827;
        border:1px solid #2a3348;
        border-radius:9px;
        box-shadow:0 7px 18px rgba(0,0,0,.14);
      }
      .summary-map-node-right{
        justify-content:flex-end;
        text-align:right;
      }
      .summary-map-node-icon{
        flex:0 0 27px;
        width:27px;
        height:27px;
        border-radius:7px;
        display:grid;
        place-items:center;
        background:#2b1730;
        color:#ff2b9d;
      }
      .summary-map-node-copy{
        min-width:0;
        flex:1;
      }
      .summary-map-node-copy strong{
        display:block;
        color:#dfe6ef;
        font-size:7px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
      }
      .summary-map-node-copy span{
        display:block;
        margin-top:3px;
        color:#ff2b9d;
        font-size:7px;
      }
      .summary-map-node-copy small{
        display:block;
        margin-top:3px;
        color:#65718a;
        font-size:6px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
      }
      .summary-map-center{
        position:absolute;
        left:50%;
        top:50%;
        transform:translate(-50%,-50%);
        width:72px;
        height:72px;
        border-radius:50%;
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        background:#ec0f82;
        border:1px solid #ff53b1;
        color:#fff;
        z-index:4;
        box-shadow:0 0 34px rgba(236,15,130,.16);
      }
      .summary-map-center div{
        font-size:11px;
        line-height:1;
        letter-spacing:.08em;
      }
      .summary-map-center span{
        margin-top:4px;
        color:#ffe5f3;
        font-size:6px;
        letter-spacing:.08em;
      }
      .summary-map-connectors{
        position:absolute;
        inset:0;
        z-index:1;
        pointer-events:none;
      }
      .summary-map-connectors span{
        position:absolute;
        left:50%;
        top:50%;
        width:25%;
        height:1px;
        background:#4a2b62;
        transform-origin:left center;
      }
      .summary-map-connectors span:nth-child(1){transform:rotate(205deg)}
      .summary-map-connectors span:nth-child(2){transform:rotate(188deg)}
      .summary-map-connectors span:nth-child(3){transform:rotate(172deg)}
      .summary-map-connectors span:nth-child(4){transform:rotate(155deg)}
      .summary-map-connectors span:nth-child(5){transform:rotate(335deg)}
      .summary-map-connectors span:nth-child(6){transform:rotate(352deg)}
      .summary-map-connectors span:nth-child(7){transform:rotate(8deg)}
      .summary-map-connectors span:nth-child(8){transform:rotate(25deg)}

      .summary-formats-heading{
        display:flex;
        justify-content:space-between;
        align-items:center;
        margin:0 0 10px;
      }
      .summary-format-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:12px;
      }
      .summary-format-card{
        position:relative;
        display:grid;
        grid-template-columns:31px minmax(0,1fr) auto;
        gap:10px;
        align-items:start;
        text-align:left;
        width:100%;
        padding:13px;
        background:#151a27;
        border:1px solid #293248;
        border-radius:11px;
        color:#eef2f7;
        cursor:pointer;
      }
      .summary-format-card.selected{
        border-color:#ec0f82;
        background:#1b1320;
      }
      .summary-format-icon{
        width:31px;
        height:31px;
        display:grid;
        place-items:center;
        background:#32182f;
        color:#ff2b9d;
        border-radius:8px;
      }
      .summary-format-copy{
        min-width:0;
        padding-right:22px;
      }
      .summary-format-copy strong{
        display:block;
        color:#eef2f7;
        font-size:10px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
      }
      .summary-format-copy>span{
        display:inline-block;
        margin-top:8px;
        padding:3px 6px;
        border-radius:5px;
        background:#101624;
        color:#ff2b9d;
        font-size:7px;
      }
      .summary-format-progress{
        height:5px;
        margin-top:10px;
        background:#20283a;
        border-radius:99px;
        overflow:hidden;
      }
      .summary-format-progress>div{
        height:100%;
        background:#ec0f82;
        border-radius:99px;
      }
      .summary-format-copy small{
        display:block;
        margin-top:7px;
        color:#ff2b9d;
        font-size:7px;
      }
      .summary-format-card>b{
        color:#ff2b9d;
        font-size:10px;
      }
      .summary-format-radio{
        position:absolute;
        right:12px;
        top:12px;
        width:15px;
        height:15px;
        border-radius:50%;
        border:1px solid #47546d;
        color:#fff;
        display:grid;
        place-items:center;
        font-size:8px;
      }
      .summary-format-radio.active{
        background:#ec0f82;
        border-color:#ec0f82;
      }

      @media(max-width:1050px){
        .summary-format-grid{grid-template-columns:1fr}
        .summary-map-side{width:36%}
        .summary-map-connectors{opacity:.45}
      }
      @media(max-width:760px){
        .presentation-centered-content{
          padding:48px 24px 40px!important;
        }
        .summary-progress-track{flex-wrap:wrap;gap:9px}
        .summary-step{min-width:90px}
        .summary-step-line{display:none}
        .summary-section-header,.summary-formats-heading{
          flex-direction:column;
          align-items:center;
        }
        .summary-section-header>div{text-align:center}
        .summary-map{min-height:560px}
        .summary-map-side{
          position:static;
          width:auto;
          margin:12px;
          gap:8px;
        }
        .summary-map-center{
          position:static;
          transform:none;
          margin:10px auto;
        }
        .summary-map-connectors{display:none}
      }

      @media(max-width:1050px){
        .simulation-bar-row{grid-template-columns:120px minmax(0,1fr) 70px}
        .solution-flow{grid-template-columns:1fr 1fr;gap:12px}
        .flow-arrow{display:none}
      }
      @media(max-width:760px){

        .simulation-bar-row{grid-template-columns:92px minmax(0,1fr);gap:7px}.simulation-bar-value{display:none}
        .reference-slide{max-width:none;min-height:420px;aspect-ratio:auto;padding:30px 24px}
        .audience-slide-grid{grid-template-columns:1fr;gap:18px}
        .solution-flow{grid-template-columns:1fr}
        .final-slide-content{grid-template-columns:1fr}
        .final-message{grid-column:auto;margin-top:40px;font-size:21px}
      }
    `}</style>
  );
}

function Logo() {
  return (
    <div className="brand-logo" style={{ width: 145, height: 43 }}>
      <img
        src={PRISALogo}
        alt="PRISA Media"
        className="prisa-logo"
        style={{ width: 110, height: 39, objectFit: "contain", objectPosition: "left center" }}
      />
    </div>
  );
}

function NameModal({
  value,
  onConfirm,
}: {
  value: string;
  onConfirm: (value: string) => void;
}) {
  const [name, setName] = useState(value);

  return (
    <div className="modal-overlay">
      <div className="name-modal">
        <Logo />
        <div className="modal-icon"><Users size={22} /></div>
        <span className="eyebrow">PRISA MEDIA</span>
        <h1>¿Cómo te llamas?</h1>
        <p>
          Va junto a cada cambio para que el equipo sepa quién tocó qué.
          Se guarda solo en este navegador, no se comparte con el enlace.
        </p>
        <input
          autoFocus
          value={name}
          placeholder="Tu nombre"
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter" && name.trim()) onConfirm(name.trim());
          }}
        />
        <div className="modal-actions">
          <button className="secondary-button" onClick={() => onConfirm("")}>Ahora no</button>
          <button className="primary-button" onClick={() => name.trim() && onConfirm(name.trim())}>Listo</button>
        </div>
      </div>
    </div>
  );
}

function Sidebar({
  currentPage,
  setCurrentPage,
  onNewCampaign,
  userName,
  collapsed,
  setCollapsed,
  mobileMenuOpen,
  setMobileMenuOpen,
}: {
  currentPage: PageKey;
  setCurrentPage: (page: PageKey) => void;
  onNewCampaign: () => void;
  userName: string;
  collapsed: boolean;
  setCollapsed: (value: boolean) => void;
  mobileMenuOpen: boolean;
  setMobileMenuOpen: (value: boolean) => void;
}) {
  return (
    <>
      {mobileMenuOpen && (
        <button
          className="mobile-sidebar-overlay"
          aria-label="Cerrar menú"
          onClick={() => setMobileMenuOpen(false)}
        />
      )}
      <aside className={`sidebar reference-sidebar ${collapsed ? "collapsed" : ""} ${mobileMenuOpen ? "mobile-open" : ""}`}>
      <div className="sidebar-top">
        <Logo />
        <button className="sidebar-toggle" style={{ display: "grid" }} onClick={() => setCollapsed(!collapsed)}>
          {collapsed ? <Menu size={17} /> : <X size={17} />}
        </button>
      </div>
      <button
        className="sidebar-new-campaign"
        type="button"
        onClick={() => { onNewCampaign(); setMobileMenuOpen(false); }}
        title="Iniciar una campaña nueva desde cero"
      >
        <span className="sidebar-new-icon"><Plus size={15} strokeWidth={2.4} /></span>
        <span>Nueva campaña</span>
      </button>
      <div className="sidebar-section-label">WORKSPACE</div>
      <nav className="sidebar-nav">
        {NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          return (
            <button
              key={item.key}
              className={currentPage === item.key ? "nav-item active" : "nav-item"}
              onClick={() => { setCurrentPage(item.key); setMobileMenuOpen(false); }}
              title={collapsed ? item.label : undefined}
            >
              <Icon size={16} />
              {!collapsed && <span>{item.label}</span>}
            </button>
          );
        })}
      </nav>
      {!collapsed && <div className="meeting">● Modo reunión</div>}
      <div className="sidebar-bottom">
        <button className="nav-item" onClick={() => { setCurrentPage("configuration"); setMobileMenuOpen(false); }}>
          <Settings size={16} />
          {!collapsed && <span>Configuración</span>}
        </button>
        {!collapsed && (
          <div className="sidebar-user" style={{ marginTop: 9 }}>
            <div className="avatar">{userName ? userName[0].toUpperCase() : "U"}</div>
            <div><strong>{userName || "Usuario"}</strong><span>Equipo comercial</span></div>
          </div>
        )}
      </div>
      </aside>
    </>
  );
}

function Header({
  currentPage,
  userName,
  searchTerm,
  setSearchTerm,
  onOpenMobileMenu,
}: {
  currentPage: PageKey;
  userName: string;
  searchTerm: string;
  setSearchTerm: (value: string) => void;
  onOpenMobileMenu: () => void;
}) {
  const label = NAV_ITEMS.find((item) => item.key === currentPage)?.label ?? "Campañas";
  return (
    <header className="reference-header">
      <div className="reference-head-left">
        <button className="mobile-menu-button" type="button" aria-label="Abrir menú" onClick={onOpenMobileMenu}>
          <Menu size={20} />
        </button>
        <div className="reference-head-search">
          <Search size={15} />
          <input value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} placeholder="Buscar proyectos, marcas..." />
        </div>
        <div className="reference-project">
          <span>{label === "Campañas" ? "Proyecto sin nombre" : "Proyecto sin nombre"}</span>
          <span className="reference-live">● LIVE</span>
        </div>
      </div>
      <div className="reference-head-actions">
        <button className="reference-icon" title="Actualizar"><RotateCcw size={15} /></button>
        <button className="reference-icon" title="Compartir"><Share2 size={15} /></button>
        <button className="reference-icon" title="Ayuda"><CircleHelp size={17} /></button>
        <div className="avatar" style={{ marginLeft: 4 }}>{userName ? userName[0].toUpperCase() : "U"}</div>
      </div>
    </header>
  );
}

function ToggleChip({ active, children, onClick }: { active: boolean; children: React.ReactNode; onClick: () => void }) {
  return <button className={active ? "reference-chip active" : "reference-chip"} onClick={onClick}>{children}</button>;
}

function BriefScreen({
  brief,
  completion,
  setBrief,
  onInsights,
}: {
  brief: BriefData;
  completion: number;
  setBrief: React.Dispatch<React.SetStateAction<BriefData>>;
  onInsights: () => void;
}) {
  const update = (field: keyof BriefData, value: string) => setBrief((prev) => ({ ...prev, [field]: value }));
  const toggleArray = (field: "ageRanges" | "socioeconomic" | "regions", value: string) => {
    setBrief((prev) => ({
      ...prev,
      [field]: prev[field].includes(value) ? prev[field].filter((item) => item !== value) : [...prev[field], value],
    }));
  };

  const fieldCount = [brief.brand, brief.website, brief.category, brief.product, brief.objective, brief.audience, brief.timingStart, brief.budget].filter(Boolean).length;
  const ageOptions = ["12-24", "18-34", "25-34", "35-44", "35-54", "45-54", "55-99"];
  const socioOptions = ["Bajo", "Medio", "Alto"];
  const regions = ["Amazonía", "Antioquia", "Bogotá / Centro", "Centro Oriente", "Costa Caribe", "Eje Cafetero", "Insular", "Llanos Orientales", "Pacífico", "Santanderes", "Sur de Colombia"];

  const toggleGender = (value: string) => {
    setBrief((prev) => {
      const current = prev.gender
        ? prev.gender.split(" y ").filter(Boolean)
        : [];

      const next = current.includes(value)
        ? current.filter((item) => item !== value)
        : [...current, value];

      return {
        ...prev,
        gender: next.join(" y "),
      };
    });
  };

  return (
  <div className="reference-content">
    <div className="reference-brief-card reference-card">

      <div className="reference-brief-head">
        <div>
          <h2 style={{ margin: 0, fontSize: 16 }}>
            Brief de la marca
          </h2>

          <p
            style={{
              margin: "5px 0 0",
              color: "#778398",
              fontSize: 9,
            }}
          >
            Lo que sabes antes de entrar a la reunión.
            Puede quedar incompleto.
          </p>
        </div>

        <div className="reference-complete">
          <strong>{completion}%</strong>
          <span>completitud</span>
        </div>
      </div>

      <div className="reference-progress">
        <div style={{ width: `${completion}%` }} />
      </div>

      {/* =====================================================
          TIPO DE CLIENTE: B2B / B2C
         ===================================================== */}

      <div className="reference-brief-field">
        <label>Tipo de cliente objetivo</label>

        <div className="reference-chip-wrap">

          <ToggleChip
            active={brief.clientType === "B2C"}
            onClick={() =>
              setBrief((prev) => ({
                ...prev,
                clientType:
                  prev.clientType === "B2C"
                    ? ""
                    : "B2C",

                // Al pasar a B2C limpiamos los campos B2B
                businessSector: "",
                companySize: "",
              }))
            }
          >
            B2C · Consumidor final
          </ToggleChip>

          <ToggleChip
            active={brief.clientType === "B2B"}
            onClick={() =>
              setBrief((prev) => ({
                ...prev,
                clientType:
                  prev.clientType === "B2B"
                    ? ""
                    : "B2B",

                // Al pasar a B2B limpiamos los
                // datos demográficos B2C
                gender: "",
                ageRanges: [],
                socioeconomic: [],
              }))
            }
          >
            B2B · Empresa
          </ToggleChip>

        </div>

        <small>
          Define si la comunicación está dirigida a
          consumidores finales o a otras empresas.
        </small>
      </div>

      {/* =====================================================
          INFORMACIÓN GENERAL
         ===================================================== */}

      {[
        ["Marca", "brand", "Nombre del cliente"],
        ["Sitio web", "website", "marca.com.co"],
        ["Categoría", "category", "Retail, banca, consumo masivo..."],
        ["Producto o servicio", "product", ""],
        [
          "Objetivo declarado",
          "objective",
          "Con las palabras del cliente, no traducido a jerga."
        ],
        ["Audiencia declarada", "audience", ""],
      ].map(([label, key, placeholder]) => (
        <div
          className="reference-brief-field"
          key={key}
        >
          <label>{label}</label>

          <input
            value={
              brief[key as keyof BriefData] as string
            }
            placeholder={placeholder}
            onChange={(e) =>
              update(
                key as keyof BriefData,
                e.target.value
              )
            }
          />

          {label === "Objetivo declarado" && (
            <small>
              Con las palabras del cliente,
              no traducido a jerga.
            </small>
          )}
        </div>
      ))}

      {/* =====================================================
          B2B
         ===================================================== */}

      {brief.clientType === "B2B" && (
        <>
          <div className="reference-brief-field">
            <label>Sector económico</label>

            <input
              value={brief.businessSector}
              placeholder="Ej. Tecnología, banca, retail..."
              onChange={(e) =>
                update(
                  "businessSector",
                  e.target.value
                )
              }
            />
          </div>

          <div className="reference-brief-field">
            <label>Tamaño de empresa</label>

            <div className="reference-chip-wrap">

              {[
                "Pequeña",
                "Mediana",
                "Grande",
              ].map((size) => (
                <ToggleChip
                  key={size}
                  active={
                    brief.companySize === size
                  }
                  onClick={() =>
                    update(
                      "companySize",
                      brief.companySize === size
                        ? ""
                        : size
                    )
                  }
                >
                  {size}
                </ToggleChip>
              ))}

            </div>

            <small>
              Selecciona el tamaño de la empresa
              a la que se dirige la campaña.
            </small>
          </div>
        </>
      )}

      {/* =====================================================
          B2C
         ===================================================== */}

      {brief.clientType === "B2C" && (
        <>
          <div className="reference-brief-field">
            <label>Género</label>

            <div className="reference-chip-wrap">

              {["Mujeres", "Hombres"].map((item) => (
                <ToggleChip
                  key={item}
                  active={
                    brief.gender === item ||
                    brief.gender.split(" y ").includes(item)
                  }
                  onClick={() => toggleGender(item)}
                >
                  {item}
                </ToggleChip>
              ))}

            </div>

            <small>
              Puedes elegir Mujeres, Hombres o ambos.
            </small>
          </div>

          <div className="reference-brief-field">
            <label>Rango de edad</label>

            <div className="reference-chip-wrap">

              {ageOptions.map((item) => (
                <ToggleChip
                  key={item}
                  active={brief.ageRanges.includes(item)}
                  onClick={() =>
                    toggleArray(
                      "ageRanges",
                      item
                    )
                  }
                >
                  {item}
                </ToggleChip>
              ))}

            </div>

            <small>
              Puede seleccionarse más de uno.
            </small>
          </div>

          <div className="reference-brief-field">
            <label>
              Nivel socioeconómico
            </label>

            <div className="reference-chip-wrap">

              {socioOptions.map((item) => (
                <ToggleChip
                  key={item}
                  active={
                    brief.socioeconomic.includes(item)
                  }
                  onClick={() =>
                    toggleArray(
                      "socioeconomic",
                      item
                    )
                  }
                >
                  {item}
                </ToggleChip>
              ))}

            </div>
          </div>
        </>
      )}

      {/* =====================================================
          REGIONES
         ===================================================== */}

      <div className="reference-brief-field">
        <label>Regiones objetivo</label>

        <div className="reference-chip-wrap">

          {regions.map((item) => (
            <ToggleChip
              key={item}
              active={brief.regions.includes(item)}
              onClick={() =>
                toggleArray(
                  "regions",
                  item
                )
              }
            >
              {item}
            </ToggleChip>
          ))}

        </div>

        <small>
          La fuente llega a departamento/región.
          No hay desglose por ciudad.
        </small>
      </div>

{/* =====================================================
    CONTEXTO
   ===================================================== */}

<div className="reference-brief-field">
  <label>Contexto del cliente</label>

  <textarea
    value={brief.context}
    placeholder="Situación, competencia, antecedentes, restricciones..."
    onChange={(e) =>
      update(
        "context",
        e.target.value
      )
    }
  />
</div>

{/* =====================================================
    FECHA Y PRESUPUESTO
   ===================================================== */}

<div className="reference-brief-field">
  <label>Inicio de campaña</label>

  <input
    type="date"
    value={brief.timingStart}
    onChange={(e) =>
      update(
        "timingStart",
        e.target.value
      )
    }
  />
</div>

<div className="reference-brief-field">
  <label>Presupuesto</label>

  <input
    type="text"
    inputMode="numeric"
    pattern="[0-9]*"
    value={brief.budget}
    placeholder="Opcional"
    onChange={(e) =>
      update(
        "budget",
        e.target.value.replace(/\D/g, "")
      )
    }
  />

  <small>
    Se utilizará posteriormente para relacionarlo
    con las tarifas del tarifario comercial.
  </small>
</div>

{/* =====================================================
    ACCIONES
   ===================================================== */}

<div
  style={{
    display: "flex",
    gap: 8,
    marginTop: 12,
  }}
>
  <button
    className="primary-button"
    onClick={onInsights}
  >
    <Sparkles size={13} />
    Empezar descubrimiento
  </button>


</div>

{/* =====================================================
    ESTADO DEL BRIEF
   ===================================================== */}

<div className="reference-brief-footer">

  <div className="reference-info-block">
    <span>Estado del brief</span>

    <strong>
      {completion === 0
        ? "Sin información"
        : completion === 100
        ? "Completo"
        : "En construcción"}
    </strong>
  </div>

  <div className="reference-info-block">
    <span>Campos diligenciados</span>

    <strong>
      {fieldCount} de 8
    </strong>
  </div>

  <div className="reference-info-block">
    <span>Siguiente paso</span>

    <strong className="next">
      Insights
      <ChevronRight size={12} />
    </strong>
  </div>

</div>

    </div>

    <div
      style={{
        display: "flex",
        justifyContent: "flex-end",
        marginTop: 18,
      }}
    >
      <button
        className="primary-button"
        onClick={onInsights}
      >
        Continuar a Insights
        <ArrowRight size={14} />
      </button>
    </div>

  </div>
);
}


function DiscoveryScreen({
  discovery,
  setDiscovery,
  onContinue,
  onBack,
}: {
  discovery: DiscoveryData;
  setDiscovery: React.Dispatch<React.SetStateAction<DiscoveryData>>;
  onContinue: () => void;
  onBack: () => void;
}) {
  const [openMoreInfo, setOpenMoreInfo] = useState<Record<string, boolean>>({
    challenge: false,
    result: false,
    audience: false,
    reaction: false,
  });

  const toggleOption = (
    field: "challenge" | "result" | "audience" | "reaction",
    value: string
  ) => {
    setDiscovery((prev) => ({
      ...prev,
      [field]: prev[field].includes(value)
        ? prev[field].filter((item) => item !== value)
        : [...prev[field], value],
    }));
  };

  const questions = [
    {
      number: "01",
      title: "¿Cuál es el principal reto que enfrenta la marca?",
      description:
        "¿Qué es lo que más necesita resolver la marca en este momento?",
      field: "challenge" as const,
      options: [
        "Generar conocimiento",
        "Conseguir nuevos clientes",
        "Aumentar ventas",
        "Diferenciarse",
        "Fidelizar clientes",
        "Lanzar un nuevo producto o servicio",
        "Enfrentar a la competencia",
      ],
      moreField: "challengeMore" as const,
    },
    {
      number: "02",
      title: "¿Qué resultado necesita conseguir la campaña?",
      description:
        "¿Qué te gustaría lograr con esta campaña o qué tendría que cambiar para que haya sido exitosa?",
      field: "result" as const,
      options: [
        "Generar conocimiento",
        "Aumentar consideración",
        "Generar tráfico",
        "Generar interacción",
        "Conseguir conversiones",
        "Aumentar ventas",
        "Conseguir nuevos registros o contactos",
      ],
      moreField: "resultMore" as const,
    },
    {
      number: "03",
      title: "¿A quién necesita llegar?",
      description:
        "¿Cuál es el público al que realmente queremos llegar con esta campaña?",
      field: "audience" as const,
      options:
        discoveryChallengeAudienceOptions(discovery),
      moreField: "audienceMore" as const,
    },
    {
      number: "04",
      title:
        "¿Qué debería hacer o sentir esa audiencia después de conocer la marca?",
      description:
        "Después de ver o escuchar la campaña, ¿qué te gustaría que las personas hicieran, pensaran o sintieran?",
      field: "reaction" as const,
      options: [
        "Conocer mejor la marca",
        "Recordar la marca",
        "Considerarla",
        "Interactuar",
        "Comprar",
        "Compartir",
        "Recomendar",
      ],
      moreField: "reactionMore" as const,
    },
  ];

  const answeredQuestions =
    questions.filter(
      (question) => discovery[question.field].length > 0
    ).length +
    (discovery.moreInformation.trim() ? 1 : 0);

  return (
    <div className="reference-content">
      <div className="reference-title-row">
        <div>
          <span className="reference-eyebrow">DESCUBRIMIENTO</span>
          <h1>Entendamos la necesidad real</h1>
          <p>
            Cinco preguntas para pasar del brief a una necesidad clara de negocio.
          </p>
        </div>

        <div className="reference-complete">
          <strong>{answeredQuestions}/5</strong>
          <span>respondidas</span>
        </div>
      </div>

      <div
        className="reference-card"
        style={{
          padding: "15px 17px",
          marginBottom: 12,
        }}
      >
        <div
          style={{
            color: "#8f9aae",
            fontSize: 9,
            lineHeight: 1.6,
          }}
        >
          Estas preguntas ayudan a entender qué necesita realmente la marca
          antes de recomendar emisoras, franquicias y formatos.
        </div>
      </div>

      <div className="reference-insights-grid">
        {questions.map((question) => {
          const options = question.options;

          return (
            <div
              className="reference-insight-card"
              key={question.number}
            >
              <div
                style={{
                  display: "flex",
                  gap: 12,
                  alignItems: "flex-start",
                }}
              >
                <div
                  style={{
                    minWidth: 26,
                    color: "#ff2b9d",
                    fontSize: 9,
                    fontWeight: 800,
                    paddingTop: 2,
                  }}
                >
                  {question.number}
                </div>

                <div style={{ flex: 1 }}>
                  <h3
                    style={{
                      margin: 0,
                      color: "#eef2f7",
                      fontSize: 12,
                    }}
                  >
                    {question.title}
                  </h3>

                  <p style={{ margin: "6px 0 10px" }}>
                    {question.description}
                  </p>

                  <div className="reference-chip-wrap">
                    {options.map((option) => (
                      <ToggleChip
                        key={option}
                        active={discovery[
                          question.field
                        ].includes(option)}
                        onClick={() =>
                          toggleOption(
                            question.field,
                            option
                          )
                        }
                      >
                        {option}
                      </ToggleChip>
                    ))}
                  </div>

                  <button
                    type="button"
                    className="reference-why discovery-more-button"
                    style={{
                      marginTop: 10,
                      textAlign: "left",
                      background: "transparent",
                      border: 0,
                      boxShadow: "none",
                      padding: 0,
                      color: "#ff2b9d",
                    }}
                    onClick={() =>
                      setOpenMoreInfo((prev) => ({
                        ...prev,
                        [question.field]: !prev[question.field],
                      }))
                    }
                  >
                    {openMoreInfo[question.field]
                      ? "− Ocultar más información"
                      : "+ Más información"}
                  </button>

                  {openMoreInfo[question.field] && (
                    <textarea
                      value={discovery[question.moreField]}
                      onChange={(event) =>
                        setDiscovery((prev) => ({
                          ...prev,
                          [question.moreField]:
                            event.target.value,
                        }))
                      }
                      style={{
                        width: "100%",
                        marginTop: 9,
                        minHeight: 58,
                        background: "#181e2d",
                        color: "#e9eef5",
                        border: "1px solid #293248",
                        borderRadius: 8,
                        padding: "9px 10px",
                        outline: 0,
                        fontSize: 9,
                        resize: "vertical",
                      }}
                      placeholder="Agrega más información sobre esta respuesta..."
                    />
                  )}
                </div>
              </div>
            </div>
          );
        })}
      </div>

      <div
        className="reference-card"
        style={{
          marginTop: 12,
          padding: 18,
        }}
      >
        <div
          style={{
            display: "flex",
            gap: 12,
            alignItems: "flex-start",
          }}
        >
          <div
            style={{
              minWidth: 26,
              color: "#ff2b9d",
              fontSize: 9,
              fontWeight: 800,
              paddingTop: 2,
            }}
          >
            05
          </div>

          <div style={{ flex: 1 }}>
            <h3
              style={{
                margin: 0,
                color: "#eef2f7",
                fontSize: 12,
              }}
            >
              ¿Hay algo más que debamos saber sobre la necesidad de la marca?
            </h3>

            <p style={{ margin: "6px 0 10px" }}>
              Cuéntanos cualquier detalle, contexto o información que consideres
              importante para entender mejor el reto de la campaña.
            </p>

            <textarea
              value={discovery.moreInformation}
              placeholder="Escribe aquí cualquier detalle adicional que haya compartido el cliente..."
              onChange={(event) =>
                setDiscovery((prev) => ({
                  ...prev,
                  moreInformation: event.target.value,
                }))
              }
              style={{
                width: "100%",
                minHeight: 100,
                background: "#181e2d",
                color: "#e9eef5",
                border: "1px solid #293248",
                borderRadius: 8,
                padding: "10px 11px",
                outline: 0,
                fontSize: 9,
                resize: "vertical",
              }}
            />

            <small
              style={{
                display: "block",
                marginTop: 5,
                color: "#66728a",
                fontSize: 7,
              }}
            >
              Puedes escribir libremente. No es necesario usar lenguaje técnico.
            </small>
          </div>
        </div>
      </div>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginTop: 17,
          paddingTop: 14,
          borderTop: "1px solid #293248",
        }}
      >
        <button
          className="secondary-button"
          onClick={onBack}
        >
          <ArrowLeft size={13} />
          Brief
        </button>

        <button
          className="primary-button"
          onClick={onContinue}
        >
          Continuar a Insights
          <ArrowRight size={13} />
        </button>
      </div>
    </div>
  );
}

function discoveryChallengeAudienceOptions(discovery: DiscoveryData): string[] {
  void discovery;
  return [
    "Clientes actuales",
    "Nuevos clientes",
    "Personas que conocen la categoría",
    "Personas que todavía no conocen la marca",
    "Público joven",
    "Adultos",
    "Familias",
    "Tomadores de decisión",
  ];
}


function CampaignsScreen({
  projects,
  onNewCampaign,
  onOpenProject,
  onRefresh,
  searchTerm,
}: {
  projects: Project[];
  onNewCampaign: () => void;
  onOpenProject: (project: Project) => void;
  onRefresh: () => void;
  searchTerm: string;
}) {
  const filtered = projects.filter((project) =>
    normalize(`${project.name} ${project.brief.brand} ${project.brief.category}`).includes(
      normalize(searchTerm)
    )
  );

  return (
    <div className="reference-content">
      <div className="reference-card reference-info">
        <UsersRound size={18} />
        <div>
          <strong>Campañas guardadas en este navegador</strong>
          <div>
            Aquí puedes conservar briefs, respuestas, selecciones y propuestas para
            continuar trabajando en la misma campaña.
          </div>
        </div>
        <button
          className="secondary-button"
          style={{ marginLeft: "auto" }}
          onClick={onRefresh}
          type="button"
        >
          Actualizar
        </button>
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 15 }}>
        <div
          className="reference-head-search"
          style={{ flex: 1, width: "auto" }}
        >
          <Search size={15} />
          <input
            value={searchTerm}
            readOnly
            placeholder="Usa el buscador superior para filtrar..."
          />
        </div>

        <button
          className="primary-button"
          onClick={onNewCampaign}
          type="button"
        >
          <Plus size={14} /> Nueva campaña
        </button>
      </div>

      <div style={{ marginTop: 17 }}>
        {filtered.length === 0 ? (
          <div
            className="reference-card"
            style={{ padding: 26, textAlign: "center", color: "#7f8ba0" }}
          >
            <BriefcaseBusiness size={24} style={{ opacity: 0.7 }} />
            <h3 style={{ margin: "10px 0 5px", color: "#eef2f7", fontSize: 13 }}>
              Aún no hay campañas guardadas
            </h3>
            <p style={{ margin: 0, fontSize: 9 }}>
              Completa una campaña y pulsa «Guardar campaña» en Presentación.
            </p>
          </div>
        ) : (
          filtered.map((project) => (
            <button
              key={project.id}
              type="button"
              onClick={() => onOpenProject(project)}
              className="campaign-card"
              style={{
                width: "100%",
                marginBottom: 10,
                textAlign: "left",
                cursor: "pointer",
              }}
            >
              <div className="campaign-card-top">
                <span className="status-badge">
                  <span className="status-dot" />
                  {project.status}
                </span>
                <ChevronRight size={16} />
              </div>
              <h3>{project.name}</h3>
              <p>{project.brief.brand || "Sin marca"}</p>
              <small>
                {project.discovery.challenge.length +
                  project.discovery.result.length +
                  project.discovery.audience.length +
                  project.discovery.reaction.length}{" "}
                respuestas &nbsp; {project.selectedFormats.length} formatos &nbsp;
                {project.brief.clientType || "Tipo sin definir"} &nbsp;{" "}
                {project.updatedAt}
              </small>
            </button>
          ))
        )}
      </div>
    </div>
  );
}

function SummaryScreen({
  brief,
  formats,
  opportunities,
  regions,
  discovery,
  selectedOpportunity,
  selectedBroadcasterId,
  selectedFranchiseId,
  selectedFormats,
  onNavigate,
}: {
  brief: BriefData;
  formats: RecommendedFormat[];
  opportunities: Opportunity[];
  regions: RegionInventory[];
  discovery: DiscoveryData;
  selectedOpportunity: string | null;
  selectedBroadcasterId: string | null;
  selectedFranchiseId: string | null;
  selectedFormats: string[];
  onNavigate: (page: PageKey) => void;
}) {
  const stageState = [
    Boolean(
      brief.brand &&
      brief.objective &&
      brief.audience &&
      brief.timingStart &&
      brief.clientType
    ),
    Boolean(
      discovery.challenge.length ||
      discovery.result.length ||
      discovery.audience.length ||
      discovery.reaction.length ||
      discovery.moreInformation.trim()
    ),
    Boolean(selectedOpportunity),
    Boolean(selectedBroadcasterId || selectedFranchiseId || selectedFormats.length),
    Boolean(selectedFormats.length),
  ];

  const stagePages: PageKey[] = [
    "brief",
    "discovery",
    "opportunities",
    "formats",
    "proposal",
  ];

  const stageLabels = [
    ["Brief", "En progreso"],
    ["Descubrimiento", "En progreso"],
    ["Oportunidades", "Pendiente"],
    ["Formatos", "Pendiente"],
    ["Propuesta", "Pendiente"],
  ];

  const mapLeft = [
    {
      title: "Marca & categoría",
      count: brief.brand ? 1 : 0,
      detail: brief.category || "Sin señales",
    },
    {
      title: "Audiencia & comportamiento",
      count: brief.audience || discovery.audience.length ? 1 : 0,
      detail:
        discovery.audience.length > 0
          ? discovery.audience.join(" · ")
          : brief.audience || "Sin señales",
    },
    {
      title: "Objetivos declarados",
      count: discovery.result.length || (brief.objective ? 1 : 0),
      detail:
        discovery.result.length > 0
          ? discovery.result.join(" · ")
          : brief.objective || "Sin señales",
    },
    {
      title: "Contexto & geografía",
      count: brief.regions.length,
      detail:
        brief.regions.length > 0
          ? brief.regions.join(" · ")
          : "Sin regiones",
    },
  ];

  const mapRight = [
    {
      title: "Dolores clave",
      count: opportunities.find((item) => item.kind === "pain")?.count ?? 0,
      detail:
        discovery.challenge.length > 0
          ? discovery.challenge.join(" · ")
          : brief.context
          ? "Contexto capturado"
          : "Sin señales",
    },
    {
      title: "Territorios de oportunidad",
      count: brief.regions.length,
      detail:
        brief.regions.length > 0
          ? brief.regions.join(" · ")
          : "Sin territorios",
    },
    {
      title: "Audiencias clave",
      count: brief.audience || discovery.audience.length ? 1 : 0,
      detail:
        brief.audience || discovery.audience.length > 0
          ? [brief.audience, ...discovery.audience].filter(Boolean).join(" · ")
          : "Sin señales",
    },
    {
      title: "Espacios estratégicos",
      count: formats.length,
      detail:
        formats.length > 0 ? `${formats.length} formatos recomendados` : "Sin señales",
    },
  ];

  const recommendedFormats = formats
    .slice()
    .sort((a, b) => b.affinity - a.affinity)
    .slice(0, 6);

  return (
    <div className="reference-content summary-reference-page">
      <div className="summary-progress-heading">
        <span className="reference-eyebrow">PROGRESO DE LA CAMPAÑA</span>
      </div>

      <div className="summary-progress-card reference-card">
        <div className="summary-progress-track">
          {stageLabels.map(([label, pendingLabel], index) => {
            const done = stageState[index];
            return (
              <React.Fragment key={label}>
                <button
                  type="button"
                  className={`summary-step ${done ? "done" : "pending"}`}
                  onClick={() => onNavigate(stagePages[index])}
                >
                  <span className="summary-step-number">
                    {done ? <Check size={12} /> : index + 1}
                  </span>
                  <strong>{label}</strong>
                  <small>{done ? "Listo" : pendingLabel}</small>
                </button>
                {index < stageLabels.length - 1 && (
                  <span
                    className={`summary-step-line ${
                      done ? "done" : ""
                    }`}
                  />
                )}
              </React.Fragment>
            );
          })}
        </div>
      </div>

      <div className="summary-section-header">
        <div>
          <span className="reference-eyebrow">MAPA DE OPORTUNIDADES</span>
          <h1>Mapa de oportunidades</h1>
        </div>

        <button
          type="button"
          className="secondary-button"
          onClick={() => onNavigate("opportunities")}
        >
          Ver mapa completo <ChevronRight size={13} />
        </button>
      </div>

      <div className="summary-map-card reference-card">
        <div className="summary-map">
          <div className="summary-map-side summary-map-side-left">
            {mapLeft.map((node) => (
              <div className="summary-map-node" key={node.title}>
                <div className="summary-map-node-icon">
                  <Target size={12} />
                </div>
                <div className="summary-map-node-copy">
                  <strong>{node.title}</strong>
                  <span>{node.count} señales</span>
                  <small>{node.detail}</small>
                </div>
              </div>
            ))}
          </div>

          <div className="summary-map-center">
            <div>•••</div>
            <span>ANÁLISIS</span>
          </div>

          <div className="summary-map-side summary-map-side-right">
            {mapRight.map((node) => (
              <div
                className="summary-map-node summary-map-node-right"
                key={node.title}
              >
                <div className="summary-map-node-copy">
                  <strong>{node.title}</strong>
                  <span>{node.count} señales</span>
                  <small>{node.detail}</small>
                </div>
                <div className="summary-map-node-icon">
                  <Layers3 size={12} />
                </div>
              </div>
            ))}
          </div>

          <div className="summary-map-connectors" aria-hidden="true">
            <span />
            <span />
            <span />
            <span />
            <span />
            <span />
            <span />
            <span />
          </div>
        </div>
      </div>

      <div className="summary-formats-heading">
        <span className="reference-eyebrow">FORMATOS RECOMENDADOS</span>
        <button
          type="button"
          className="secondary-button"
          onClick={() => onNavigate("formats")}
        >
          Ver todos los formatos <ChevronDown size={13} />
        </button>
      </div>

      <div className="summary-format-grid">
        {recommendedFormats.length > 0 ? (
          recommendedFormats.map((format) => {
            const active = selectedFormats.includes(format.id);

            return (
              <button
                type="button"
                className={`summary-format-card ${active ? "selected" : ""}`}
                key={format.id}
                onClick={() => onNavigate("formats")}
              >
                <div className="summary-format-icon">
                  <Layers3 size={15} />
                </div>

                <div className="summary-format-copy">
                  <strong>{format.name}</strong>
                  <span>
                    {format.affinity >= 66
                      ? "Afinidad MEDIA"
                      : "Afinidad BAJA"}
                  </span>

                  <div className="summary-format-progress">
                    <div
                      style={{
                        width: `${Math.max(
                          0,
                          Math.min(100, format.affinity)
                        )}%`,
                      }}
                    />
                  </div>

                  <small>Ver detalle →</small>
                </div>

                <b>{format.affinity}%</b>

                <span className={`summary-format-radio ${active ? "active" : ""}`}>
                  {active ? "✓" : ""}
                </span>
              </button>
            );
          })
        ) : (
          <div
            className="reference-panel"
            style={{
              gridColumn: "1 / -1",
              textAlign: "center",
              padding: 22,
            }}
          >
            <h2>Sin formatos recomendados todavía</h2>
            <p style={{ marginTop: 8 }}>
              Completa Descubrimiento para generar recomendaciones comerciales.
            </p>
          </div>
        )}
      </div>

      {regions.length > 0 && (
        <div className="reference-info" style={{ marginTop: 15 }}>
          <MapPin size={16} />
          <div>
            Inventario disponible en geofocalización: {regions.length} regiones
            con datos.
          </div>
        </div>
      )}
    </div>
  );
}

function InsightsScreen({
  brief,
  discovery,
  onBack,
  onContinue,
}: {
  brief: BriefData;
  discovery: DiscoveryData;
  onBack: () => void;
  onContinue: () => void;
}) {
  const hasChallenge = discovery.challenge.length > 0;
  const hasResult = discovery.result.length > 0;
  const hasAudience = discovery.audience.length > 0;
  const hasReaction = discovery.reaction.length > 0;
  const hasMoreInformation = discovery.moreInformation.trim().length > 0;

  const coveredDimensions = [
    hasChallenge,
    hasResult,
    hasAudience,
    hasReaction,
    hasMoreInformation,
  ].filter(Boolean).length;

  const missingDimensions = [
    !hasChallenge ? "Reto de la marca" : null,
    !hasResult ? "Resultado esperado" : null,
    !hasAudience ? "Audiencia prioritaria" : null,
    !hasReaction ? "Reacción esperada" : null,
    !hasMoreInformation ? "Contexto adicional" : null,
  ].filter((item): item is string => Boolean(item));

  const clientTypeLabel =
    brief.clientType === "B2B"
      ? "B2B · Empresa"
      : brief.clientType === "B2C"
      ? "B2C · Consumidor final"
      : "Tipo de cliente no definido";

  const insights: Insight[] = [
    {
      title: "Reto de la marca",
      description: hasChallenge
        ? discovery.challenge.join(" · ")
        : brief.context || "Sin señales todavía. Falta información del reto.",
      status: hasChallenge || Boolean(brief.context) ? "ready" : "empty",
      count: discovery.challenge.length,
    },
    {
      title: "Resultado esperado",
      description: hasResult
        ? discovery.result.join(" · ")
        : brief.objective || "Sin señales todavía. Falta definir qué debe cambiar.",
      status: hasResult || Boolean(brief.objective) ? "ready" : "empty",
      count: discovery.result.length,
    },
    {
      title: "Audiencia prioritaria",
      description: hasAudience
        ? discovery.audience.join(" · ")
        : brief.audience || "Sin señales todavía. Falta definir a quién necesitamos llegar.",
      status: hasAudience || Boolean(brief.audience) ? "ready" : "empty",
      count: discovery.audience.length,
    },
    {
      title: "Reacción esperada",
      description: hasReaction
        ? discovery.reaction.join(" · ")
        : "Sin señales todavía. Falta definir qué queremos provocar en la audiencia.",
      status: hasReaction ? "ready" : "empty",
      count: discovery.reaction.length,
    },
    {
      title: "Información adicional",
      description: hasMoreInformation
        ? discovery.moreInformation
        : "Sin información adicional capturada.",
      status: hasMoreInformation ? "ready" : "empty",
      count: hasMoreInformation ? 1 : 0,
    },
  ];

  return (
    <div className="reference-content">
      <div className="reference-title-row">
        <div>
          <span className="reference-eyebrow">INSIGHTS</span>
          <h1>De las respuestas a las señales</h1>
          <p>
            El sistema resume el Brief y las respuestas de Descubrimiento
            para identificar qué sabemos, qué hipótesis tenemos y qué falta cubrir.
          </p>
        </div>

        <div className="reference-complete">
          <strong>{coveredDimensions}/5</strong>
          <span>dimensiones cubiertas</span>
        </div>
      </div>

      <div className="reference-card" style={{ padding: 15, marginBottom: 14 }}>
        <div
          style={{
            display: "flex",
            gap: 14,
            flexWrap: "wrap",
            alignItems: "center",
            color: "#8c96a9",
            fontSize: 9,
          }}
        >
          <strong style={{ color: "#dde3ec" }}>Contexto de lectura</strong>
          <span>{clientTypeLabel}</span>
          {brief.brand && <span>Marca: {brief.brand}</span>}
          {brief.category && <span>Categoría: {brief.category}</span>}
        </div>

        <p style={{ margin: "9px 0 0", color: "#68748b", fontSize: 8 }}>
          Las señales se construyen a partir del Brief y de las respuestas capturadas
          en Descubrimiento. No reemplazan la información original del cliente.
        </p>
      </div>

      <div className="reference-insights-grid">
        {insights.map((insight) => (
          <div
            className="reference-insight-card"
            key={insight.title}
            style={{
              borderColor:
                insight.status === "ready"
                  ? "#293248"
                  : "#805d16",
            }}
          >
            <div className="row">
              <div>
                <h3>{insight.title}</h3>
                <span
                  style={{
                    display: "inline-block",
                    marginTop: 4,
                    color:
                      insight.status === "ready"
                        ? "#29d875"
                        : "#f0b52b",
                    fontSize: 7,
                    textTransform: "uppercase",
                    letterSpacing: ".08em",
                  }}
                >
                  {insight.status === "ready"
                    ? "Señal disponible"
                    : "Sin suficiente información"}
                </span>
              </div>

              <span style={{ color: "#8994a8", fontSize: 8 }}>
                {insight.count}
              </span>
            </div>

            <p>{insight.description}</p>
          </div>
        ))}
      </div>

      {missingDimensions.length > 0 ? (
        <div className="reference-warning">
          <strong>⚠ Dimensiones aún sin cubrir</strong>
          <div style={{ marginTop: 5 }}>
            {missingDimensions.join(" · ")}
          </div>
          <div style={{ marginTop: 5, color: "#bd9b45" }}>
            Puedes continuar y completar estas señales más adelante.
          </div>
        </div>
      ) : (
        <div
          className="reference-card"
          style={{
            padding: 14,
            marginTop: 11,
            borderColor: "#245a3d",
            background: "#111f18",
          }}
        >
          <strong style={{ color: "#29d875", fontSize: 9 }}>
            ✓ Las cinco dimensiones de Descubrimiento están cubiertas.
          </strong>
        </div>
      )}

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginTop: 17,
          paddingTop: 14,
          borderTop: "1px solid #293248",
        }}
      >
        <button className="secondary-button" onClick={onBack}>
          <ArrowLeft size={13} />
          Descubrimiento
        </button>

        <button className="primary-button" onClick={onContinue}>
          Continuar a Oportunidades
          <ArrowRight size={13} />
        </button>
      </div>
    </div>
  );
}

function OpportunitiesScreen({
  brief,
  discovery,
  selected,
  onSelect,
  onBack,
  onContinue,
}: {
  brief: BriefData;
  discovery: DiscoveryData;
  selected: string | null;
  onSelect: (id: string) => void;
  onBack: () => void;
  onContinue: () => void;
}) {
  const hasChallenge = discovery.challenge.length > 0 || Boolean(discovery.challengeMore.trim());
  const hasResult = discovery.result.length > 0 || Boolean(discovery.resultMore.trim());
  const hasAudience = discovery.audience.length > 0 || Boolean(discovery.audienceMore.trim());
  const hasReaction = discovery.reaction.length > 0 || Boolean(discovery.reactionMore.trim());
  const hasAdditionalInfo = Boolean(discovery.moreInformation.trim());

  const audienceContext = [
    ...discovery.audience,
    ...(discovery.audienceMore.trim() ? [discovery.audienceMore.trim()] : []),
  ].filter(Boolean);

  const opportunities: Opportunity[] = [
    {
      id: "pain",
      title: "Dolores identificados",
      description: hasChallenge
        ? [
            discovery.challenge.join(" · "),
            discovery.challengeMore.trim(),
          ].filter(Boolean).join(" — ")
        : brief.context
        ? `Contexto declarado por el cliente: ${brief.context}`
        : "Ningún dolor o reto capturado todavía.",
      count: discovery.challenge.length + (discovery.challengeMore.trim() ? 1 : 0),
      kind: "pain",
    },
    {
      id: "territory",
      title: "Territorios de oportunidad",
      description: brief.regions.length
        ? `La campaña tiene ${brief.regions.length} región(es) objetivo: ${brief.regions.join(" · ")}.`
        : "No hay regiones seleccionadas. El sistema podrá trabajar con cobertura disponible en el catálogo.",
      count: brief.regions.length,
      kind: "territory",
    },
    {
      id: "audience",
      title: "Audiencias clave",
      description: audienceContext.length
        ? audienceContext.join(" · ")
        : brief.audience || "No se ha definido una audiencia prioritaria.",
      count: audienceContext.length || (brief.audience ? 1 : 0),
      kind: "audience",
    },
    {
      id: "space",
      title: "Espacios estratégicos",
      description: hasResult || hasReaction
        ? [
            discovery.result.join(" · "),
            discovery.resultMore.trim(),
            discovery.reaction.join(" · "),
            discovery.reactionMore.trim(),
          ].filter(Boolean).join(" — ")
        : brief.product
        ? `Activación alrededor de ${brief.product}.`
        : "Todavía no hay suficiente información para definir el espacio estratégico.",
      count:
        discovery.result.length +
        discovery.reaction.length +
        (discovery.resultMore.trim() ? 1 : 0) +
        (discovery.reactionMore.trim() ? 1 : 0),
      kind: "space",
    },
  ];

  const sourceLabels: Record<Opportunity["kind"], string> = {
    pain: "Discovery · Reto",
    territory: "Brief · Geografía",
    audience: brief.clientType
      ? `Brief + Discovery · ${brief.clientType}`
      : "Brief + Discovery",
    space: "Discovery · Resultado + reacción",
  };

  const selectedOpportunityData = opportunities.find(
    (item) => item.id === selected
  );

  const covered = [
    hasChallenge,
    hasResult,
    hasAudience,
    hasReaction,
    hasAdditionalInfo,
  ].filter(Boolean).length;

  return (
    <div className="reference-content">
      <div className="reference-title-row">
        <div>
          <span className="reference-eyebrow">OPORTUNIDADES</span>
          <h1>Del diagnóstico a la oportunidad</h1>
          <p>
            Convertimos las señales del Brief y Descubrimiento en territorios
            concretos para orientar la recomendación comercial.
          </p>
        </div>

        <div className="reference-complete">
          <strong>{covered}/5</strong>
          <span>señales cubiertas</span>
        </div>
      </div>

      <div
        className="reference-card"
        style={{
          padding: 15,
          marginBottom: 14,
        }}
      >
        <div
          style={{
            display: "flex",
            gap: 14,
            flexWrap: "wrap",
            alignItems: "center",
            color: "#8c96a9",
            fontSize: 9,
          }}
        >
          <strong style={{ color: "#dde3ec" }}>
            Oportunidades construidas con
          </strong>
          <span>
            {brief.clientType || "Tipo de cliente no definido"}
          </span>
          {brief.brand && <span>Marca: {brief.brand}</span>}
          {brief.category && <span>Categoría: {brief.category}</span>}
        </div>

        <p
          style={{
            margin: "9px 0 0",
            color: "#68748b",
            fontSize: 8,
          }}
        >
          El sistema no reemplaza la información del cliente: la organiza para
          detectar dónde existe una oportunidad de activación.
        </p>
      </div>

      <div className="reference-op-grid">
        {opportunities.map((item) => (
          <button
            type="button"
            key={item.id}
            className="reference-op-kpi"
            onClick={() => onSelect(item.id)}
            style={{
              borderColor:
                selected === item.id ? "#ec0f82" : "#293248",
              cursor: "pointer",
              textAlign: "center",
            }}
          >
            <div
              className="format-icon"
              style={{ margin: "0 auto" }}
            >
              {item.kind === "pain" ? (
                <Target size={16} />
              ) : item.kind === "territory" ? (
                <MapPin size={16} />
              ) : item.kind === "audience" ? (
                <Users size={16} />
              ) : (
                <Layers3 size={16} />
              )}
            </div>

            <span className="big">{item.count}</span>
            <h3>{item.title}</h3>
            <small>{sourceLabels[item.kind]}</small>
          </button>
        ))}
      </div>

      <div className="reference-accordion">
        {opportunities.map((item) => (
          <div
            className="reference-accordion-row"
            key={item.id}
          >
            <button
              type="button"
              className="reference-accordion-button"
              onClick={() => onSelect(item.id)}
            >
              <div className="format-icon">
                {item.kind === "pain" ? (
                  <Target size={14} />
                ) : item.kind === "territory" ? (
                  <MapPin size={14} />
                ) : item.kind === "audience" ? (
                  <Users size={14} />
                ) : (
                  <Layers3 size={14} />
                )}
              </div>

              <span>{item.title}</span>
              <small style={{ color: "#6f7a91" }}>
                {item.count}
              </small>

              {selected === item.id ? (
                <ChevronDown size={14} />
              ) : (
                <ChevronRight size={14} />
              )}
            </button>

            {selected === item.id && (
              <div className="reference-accordion-body">
                <strong
                  style={{
                    color: "#eef2f7",
                    fontSize: 9,
                  }}
                >
                  ¿Por qué aparece esta oportunidad?
                </strong>

                <p
                  style={{
                    margin: "7px 0 0",
                    color: "#718098",
                    fontSize: 9,
                    lineHeight: 1.6,
                  }}
                >
                  {item.description}
                </p>

                <div
                  style={{
                    marginTop: 9,
                    paddingTop: 8,
                    borderTop: "1px solid #293248",
                    color: "#7f8aa0",
                    fontSize: 8,
                  }}
                >
                  Fuente: {sourceLabels[item.kind]}
                </div>
              </div>
            )}
          </div>
        ))}
      </div>

      {selectedOpportunityData && (
        <div
          className="reference-card"
          style={{
            marginTop: 13,
            padding: 15,
            borderColor: "#293248",
          }}
        >
          <span className="reference-eyebrow">
            OPORTUNIDAD SELECCIONADA
          </span>

          <h2
            style={{
              margin: "6px 0 0",
              fontSize: 14,
            }}
          >
            {selectedOpportunityData.title}
          </h2>

          <p
            style={{
              margin: "7px 0 0",
              color: "#718098",
              fontSize: 9,
              lineHeight: 1.6,
            }}
          >
            {selectedOpportunityData.description}
          </p>
        </div>
      )}

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginTop: 18,
          paddingTop: 14,
          borderTop: "1px solid #293248",
        }}
      >
        <button
          type="button"
          className="secondary-button"
          onClick={onBack}
        >
          <ArrowLeft size={13} />
          Insights
        </button>

        <button
          type="button"
          className="primary-button"
          onClick={onContinue}
        >
          Continuar a formatos
          <ArrowRight size={13} />
        </button>
      </div>
    </div>
  );
}

function FormatsScreen({
  formats,
  selected,
  setSelected,
  broadcasters,
  recommendedBroadcasters,
  selectedBroadcasterId,
  setSelectedBroadcasterId,
  franchises,
  recommendedFranchises,
  selectedFranchiseId,
  setSelectedFranchiseId,
  onBack,
  onContinue,
}: {
  formats: RecommendedFormat[];
  selected: string[];
  setSelected: React.Dispatch<React.SetStateAction<string[]>>;
  broadcasters: Broadcaster[];
  recommendedBroadcasters: RecommendedBroadcaster[];
  selectedBroadcasterId: string | null;
  setSelectedBroadcasterId: (id: string | null) => void;
  franchises: Franchise[];
  recommendedFranchises: RecommendedFranchise[];
  selectedFranchiseId: string | null;
  setSelectedFranchiseId: (id: string | null) => void;
  onBack: () => void;
  onContinue: () => void;
}) {
  const [tab, setTab] = useState("emitter");
  const [whyBroadcasterId, setWhyBroadcasterId] = useState<string | null>(null);
  const [whyFormatId, setWhyFormatId] = useState<string | null>(null);

  const selectedBroadcasterData =
    broadcasters.find((broadcaster) => broadcaster.id === selectedBroadcasterId) ?? null;

  const selectedFranchiseData =
    franchises.find((franchise) => franchise.id === selectedFranchiseId) ?? null;

  const toggleBroadcaster = (id: string) => {
    setSelectedBroadcasterId(
      selectedBroadcasterId === id ? null : id
    );
  };

  return (
    <div className="reference-content">
      <div className="reference-formats-top">
        <div className="reference-summary-pill">
          Emisora
          <strong>{selectedBroadcasterData?.name ?? "sin elegir"}</strong>
        </div>
        <div className="reference-summary-pill">
          Franquicia
          <strong>{selectedFranchiseData?.name ?? "sin elegir"}</strong>
        </div>
        <div className="reference-summary-pill">
          Formatos
          <strong>{selected.length} seleccionados</strong>
        </div>
      </div>

      <div className="reference-format-tabs">
        <button
          className={tab === "public" ? "reference-tab active" : "reference-tab"}
          onClick={() => setTab("public")}
        >
          QUÉ PÚBLICO
        </button>
        <button
          className={tab === "emitter" ? "reference-tab active" : "reference-tab"}
          onClick={() => setTab("emitter")}
        >
          DÓNDE · EMISORA
        </button>
        <button
          className={tab === "franchise" ? "reference-tab active" : "reference-tab"}
          onClick={() => setTab("franchise")}
        >
          CON QUÉ · FRANQUICIA
        </button>
        <button
          className={tab === "format" ? "reference-tab active" : "reference-tab"}
          onClick={() => setTab("format")}
        >
          CÓMO · FORMATO
        </button>
      </div>

      {tab === "emitter" && (
        <>
          <div
            className="reference-card"
            style={{ padding: 14, marginBottom: 12 }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 10,
              }}
            >
              <div style={{ width: "100%", textAlign: "center" }}>
                <span className="reference-eyebrow">
                  RECOMENDACIÓN AUTOMÁTICA
                </span>
                <h2 style={{ margin: "5px 0 0", fontSize: 14 }}>
                  Emisoras con mayor afinidad
                </h2>
                <span
                  style={{
                    display: "block",
                    marginTop: 5,
                    color: "#6f7b91",
                    fontSize: 8,
                  }}
                >
                  Basado en Brief + Discovery + catálogo
                </span>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
                gap: 9,
                marginTop: 12,
              }}
            >
              {recommendedBroadcasters.slice(0, 3).map((recommendation, index) => {
                const active = selectedBroadcasterId === recommendation.id;
                return (
                  <button
                    key={recommendation.id}
                    type="button"
                    onClick={() => toggleBroadcaster(recommendation.id)}
                    style={{
                      textAlign: "left",
                      border: `1px solid ${active ? "#ec0f82" : "#293248"}`,
                      background: active ? "#241325" : "#111827",
                      borderRadius: 10,
                      padding: 11,
                      color: "#eef2f7",
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        gap: 8,
                      }}
                    >
                      <strong style={{ fontSize: 9 }}>
                        {index + 1}. {recommendation.name}
                      </strong>
                      <strong
                        style={{
                          color: "#ff2b9d",
                          fontSize: 10,
                        }}
                      >
                        {recommendation.affinity}%
                      </strong>
                    </div>

                    <small
                      style={{
                        display: "block",
                        marginTop: 4,
                        color: "#718098",
                        fontSize: 7,
                      }}
                    >
                      Afinidad estimada
                    </small>

                    <span
                      style={{
                        display: "block",
                        marginTop: 7,
                        color: "#8f9aae",
                        fontSize: 7,
                        lineHeight: 1.45,
                      }}
                    >
                      Mayor afinidad porque:{" "}
                      {(recommendation.reasons.slice(0, 2).join(" · ")) ||
                        "coincidencia con la información disponible."}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          <p
            style={{
              color: "#93a0b5",
              fontSize: 9,
              marginBottom: 12,
            }}
          >
            Puedes aceptar la recomendación o elegir manualmente otra emisora.
          </p>

          <div className="reference-format-list">
            {broadcasters.map((broadcaster) => {
              const active = selectedBroadcasterId === broadcaster.id;
              const recommendation = recommendedBroadcasters.find(
                (item) => item.id === broadcaster.id
              );

              return (
                <div
                  className={
                    active
                      ? "reference-format-row selected"
                      : "reference-format-row"
                  }
                  key={broadcaster.id}
                  role="button"
                  tabIndex={0}
                  onClick={() => toggleBroadcaster(broadcaster.id)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      toggleBroadcaster(broadcaster.id);
                    }
                  }}
                  style={{ cursor: "pointer" }}
                >
                  <button
                    type="button"
                    className={
                      active
                        ? "reference-radio active"
                        : "reference-radio"
                    }
                    aria-label={`Seleccionar ${broadcaster.name}`}
                    onClick={(event) => {
                      event.stopPropagation();
                      toggleBroadcaster(broadcaster.id);
                    }}
                  />

                  <div className="reference-emitter-icon">◉</div>

                  <div>
                    <strong>{broadcaster.name}</strong>
                    <small>
                      {broadcaster.profile ||
                        broadcaster.coverage ||
                        "Perfil disponible en el catálogo PRISA"}
                    </small>
                  </div>

                  <div>
                    <div className="reference-progress-bar">
                      <div
                        style={{
                          width: `${recommendation?.affinity ?? 0}%`,
                        }}
                      />
                    </div>
                    <small>
                      {recommendation
                        ? `${recommendation.affinity}% afinidad`
                        : "Sin evaluación"}
                    </small>
                  </div>

                  <button
                    type="button"
                    className="reference-why"
                    onClick={(event) => {
                      event.stopPropagation();
                      setWhyBroadcasterId(
                        whyBroadcasterId === broadcaster.id
                          ? null
                          : broadcaster.id
                      );
                    }}
                  >
                    Por qué ›
                  </button>

                  {whyBroadcasterId === broadcaster.id &&
                    recommendation && (
                      <div
                        style={{
                          gridColumn: "2 / -1",
                          marginTop: 2,
                          padding: 10,
                          background: "#101624",
                          borderTop: "1px solid #293248",
                          borderRadius: 8,
                        }}
                      >
                        <strong
                          style={{
                            color: "#eef2f7",
                            fontSize: 9,
                          }}
                        >
                          ¿Por qué {recommendation.affinity}%?
                        </strong>

                        <div
                          style={{
                            display: "grid",
                            gridTemplateColumns: "repeat(5, 1fr)",
                            gap: 8,
                            marginTop: 9,
                          }}
                        >
                          {[
                            ["Objetivo", recommendation.breakdown.objective],
                            ["Audiencia", recommendation.breakdown.audience],
                            ["Geografía", recommendation.breakdown.geography],
                            ["Contexto", recommendation.breakdown.context],
                            ["Tipo cliente", recommendation.breakdown.clientType],
                          ].map(([label, value]) => (
                            <div key={String(label)}>
                              <small
                                style={{
                                  display: "block",
                                  color: "#6f7b91",
                                  fontSize: 7,
                                }}
                              >
                                {label}
                              </small>
                              <strong
                                style={{
                                  display: "block",
                                  marginTop: 3,
                                  color: "#ff2b9d",
                                  fontSize: 10,
                                }}
                              >
                                {value}%
                              </strong>
                            </div>
                          ))}
                        </div>

                        <div
                          style={{
                            marginTop: 9,
                            color: "#7d889d",
                            fontSize: 8,
                          }}
                        >
                          {recommendation.reasons.join(" · ")}
                        </div>
                      </div>
                    )}
                </div>
              );
            })}
          </div>
        </>
      )}

      {tab === "public" && (
        <div className="reference-panel">
          <h2>Público objetivo</h2>
          <p
            style={{
              marginTop: 8,
              color: "#718098",
              fontSize: 9,
            }}
          >
            {selectedBroadcasterData?.name
              ? `La selección actual es ${selectedBroadcasterData.name}. El perfil se toma directamente de la hoja Emisoras del catálogo.`
              : "La audiencia se determina a partir del brief, Discovery y el perfil disponible en el catálogo PRISA."}
          </p>
        </div>
      )}

      {tab === "franchise" && (
        <>
          <div
            className="reference-card"
            style={{ padding: 14, marginBottom: 12 }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 10,
              }}
            >
              <div style={{ width: "100%", textAlign: "center" }}>
                <span className="reference-eyebrow">
                  OPCIONAL · RECOMENDACIÓN AUTOMÁTICA
                </span>
                <h2 style={{ margin: "5px 0 0", fontSize: 14 }}>
                  Franquicias con mayor afinidad
                </h2>
                <span
                  style={{
                    display: "block",
                    marginTop: 5,
                    color: "#6f7b91",
                    fontSize: 8,
                  }}
                >
                  Basado en Brief + Discovery + catálogo
                </span>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
                gap: 9,
                marginTop: 12,
              }}
            >
              <button
                type="button"
                onClick={() => setSelectedFranchiseId(null)}
                style={{
                  textAlign: "left",
                  border: `1px solid ${!selectedFranchiseId ? "#ec0f82" : "#293248"}`,
                  background: !selectedFranchiseId ? "#241325" : "#111827",
                  borderRadius: 10,
                  padding: 11,
                  color: "#eef2f7",
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                  <strong style={{ fontSize: 9 }}>Sin franquicia</strong>
                  {!selectedFranchiseId && (
                    <Check size={13} color="#ff2b9d" />
                  )}
                </div>
                <small style={{ display: "block", marginTop: 4, color: "#718098", fontSize: 7 }}>
                  Opción válida para cualquier campaña sin franquicia específica.
                </small>
              </button>

              {recommendedFranchises.slice(0, 2).map((recommendation, index) => {
                const active = selectedFranchiseId === recommendation.id;
                return (
                  <button
                    key={recommendation.id}
                    type="button"
                    onClick={() => setSelectedFranchiseId(recommendation.id)}
                    style={{
                      textAlign: "left",
                      border: `1px solid ${active ? "#ec0f82" : "#293248"}`,
                      background: active ? "#241325" : "#111827",
                      borderRadius: 10,
                      padding: 11,
                      color: "#eef2f7",
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                      <strong style={{ fontSize: 9 }}>
                        {index + 1}. {recommendation.name}
                      </strong>
                      <strong style={{ color: "#ff2b9d", fontSize: 10 }}>
                        {recommendation.affinity}%
                      </strong>
                    </div>
                    <small style={{ display: "block", marginTop: 4, color: "#718098", fontSize: 7 }}>
                      {recommendation.category || "Franquicia PRISA"}
                    </small>
                    <span
                      style={{
                        display: "block",
                        marginTop: 7,
                        color: "#8f9aae",
                        fontSize: 7,
                        lineHeight: 1.45,
                      }}
                    >
                      Mayor afinidad porque:{" "}
                      {(recommendation.reasons.slice(0, 2).join(" · ")) ||
                        "coincidencia con la información disponible."}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          <p style={{ color: "#93a0b5", fontSize: 9, marginBottom: 12 }}>
            La franquicia es opcional. Puedes dejar "Sin franquicia" o seleccionar manualmente otra opción.
          </p>

          {franchises.length === 0 ? (
            <div className="reference-panel">
              <h2>No hay franquicias disponibles</h2>
              <p style={{ marginTop: 8, color: "#718098", fontSize: 9 }}>
                No se encontraron registros en la hoja "Franquicias" del catálogo PRISA.
              </p>
            </div>
          ) : (
            <div className="reference-format-list">
              {franchises.map((franchise) => {
                const recommendation = recommendedFranchises.find(
                  (item) => item.id === franchise.id
                );
                const active = selectedFranchiseId === franchise.id;
                const whyOpen = recommendation && selectedFranchiseId === franchise.id;

                return (
                  <div
                    className={active ? "reference-format-row selected" : "reference-format-row"}
                    key={franchise.id}
                    style={{ alignItems: "start", cursor: "pointer" }}
                    role="button"
                    tabIndex={0}
                    onClick={() => setSelectedFranchiseId(active ? null : franchise.id)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter" || event.key === " ") {
                        event.preventDefault();
                        setSelectedFranchiseId(active ? null : franchise.id);
                      }
                    }}
                  >
                    <button
                      type="button"
                      className={active ? "reference-radio active" : "reference-radio"}
                      aria-label={`Seleccionar ${franchise.name}`}
                      onClick={(event) => {
                        event.stopPropagation();
                        setSelectedFranchiseId(active ? null : franchise.id);
                      }}
                    />

                    <div className="reference-emitter-icon">✦</div>

                    <div>
                      <strong>{franchise.name}</strong>
                      <small>
                        {franchise.category || franchise.description || "Franquicia disponible en el catálogo PRISA"}
                      </small>
                    </div>

                    <div>
                      <div className="reference-progress-bar">
                        <div style={{ width: `${recommendation?.affinity ?? 0}%` }} />
                      </div>
                      <small>
                        {recommendation
                          ? `${recommendation.affinity}% afinidad`
                          : "Sin evaluación"}
                      </small>
                    </div>

                    <button
                      type="button"
                      className="reference-why"
                      onClick={(event) => {
                        event.stopPropagation();
                        setSelectedFranchiseId(active ? null : franchise.id);
                      }}
                    >
                      {active ? "Seleccionada" : "Elegir"}
                    </button>

                    {whyOpen && recommendation && (
                      <div
                        style={{
                          gridColumn: "2 / -1",
                          marginTop: 2,
                          padding: 10,
                          background: "#101624",
                          borderTop: "1px solid #293248",
                          borderRadius: 8,
                        }}
                      >
                        <strong style={{ color: "#eef2f7", fontSize: 9 }}>
                          ¿Por qué {recommendation.affinity}%?
                        </strong>

                        <div
                          style={{
                            display: "grid",
                            gridTemplateColumns: "repeat(5, 1fr)",
                            gap: 8,
                            marginTop: 9,
                          }}
                        >
                          {[
                            ["Objetivo", recommendation.breakdown.objective],
                            ["Audiencia", recommendation.breakdown.audience],
                            ["Contexto", recommendation.breakdown.context],
                            ["Categoría", recommendation.breakdown.category],
                            ["Territorio", recommendation.breakdown.geography],
                          ].map(([label, value]) => (
                            <div key={String(label)}>
                              <small style={{ display: "block", color: "#6f7b91", fontSize: 7 }}>
                                {label}
                              </small>
                              <strong style={{ display: "block", marginTop: 3, color: "#ff2b9d", fontSize: 10 }}>
                                {value}%
                              </strong>
                            </div>
                          ))}
                        </div>

                        <div style={{ marginTop: 9, color: "#7d889d", fontSize: 8 }}>
                          {recommendation.reasons.join(" · ")}
                        </div>

                        <div style={{ marginTop: 7, color: "#69758c", fontSize: 7 }}>
                          Fuente: hoja "Franquicias" del catálogo PRISA.
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </>
      )}

      {tab === "format" && (
        <>
          <div
            className="reference-card"
            style={{ padding: 14, marginBottom: 12 }}
          >
            <span className="reference-eyebrow">
              RECOMENDACIÓN AUTOMÁTICA
            </span>
            <p
              style={{
                margin: "7px 0 0",
                color: "#718098",
                fontSize: 8,
              }}
            >
              Los formatos están ordenados por afinidad estimada con el
              Brief y Discovery. La selección sigue siendo manual.
            </p>
          </div>

          <div className="formats-grid">
            {formats.map((format) => {
              const active = selected.includes(format.id);

              return (
                <div
                  key={format.id}
                  className={
                    active
                      ? "format-card selected"
                      : "format-card"
                  }
                  style={{ position: "relative" }}
                >
                  <button
                    type="button"
                    onClick={() =>
                      setSelected((prev) =>
                        prev.includes(format.id)
                          ? prev.filter((id) => id !== format.id)
                          : [...prev, format.id]
                      )
                    }
                    style={{
                      width: "100%",
                      background: "transparent",
                      border: 0,
                      color: "inherit",
                      textAlign: "left",
                      padding: 0,
                    }}
                  >
                    <div className="topline">
                      <div className="format-icon">
                        <Layers3 size={18} />
                      </div>

                      <div className="format-check">
                        {active && <Check size={13} />}
                      </div>
                    </div>

                    <div className="eyebrow">
                      {format.affinity}% AFINIDAD
                    </div>

                    <h3>{format.name}</h3>

                    <p>{format.description}</p>

                    <div className="format-card-footer">
                      <span>{format.reason}</span>
                    </div>
                  </button>

                  <button
                    type="button"
                    className="reference-why"
                    style={{
                      marginTop: 10,
                      display: "block",
                    }}
                    onClick={() =>
                      setWhyFormatId(
                        whyFormatId === format.id
                          ? null
                          : format.id
                      )
                    }
                  >
                    Por qué ›
                  </button>

                  {whyFormatId === format.id && (
                    <div
                      style={{
                        marginTop: 9,
                        paddingTop: 9,
                        borderTop: "1px solid #293248",
                      }}
                    >
                      <strong
                        style={{
                          display: "block",
                          color: "#eef2f7",
                          fontSize: 8,
                        }}
                      >
                        ¿Por qué {format.affinity}%?
                      </strong>

                      <div
                        style={{
                          display: "grid",
                          gridTemplateColumns:
                            "repeat(5, 1fr)",
                          gap: 6,
                          marginTop: 8,
                        }}
                      >
                        {[
                          ["Objetivo", format.breakdown.objective],
                          ["Audiencia", format.breakdown.audience],
                          ["Contexto", format.breakdown.context],
                          ["Categoría", format.breakdown.category],
                          ["Coincidencia", format.breakdown.keyword],
                        ].map(([label, value]) => (
                          <div key={String(label)}>
                            <small
                              style={{
                                display: "block",
                                color: "#69758c",
                                fontSize: 6,
                              }}
                            >
                              {label}
                            </small>
                            <strong
                              style={{
                                display: "block",
                                marginTop: 2,
                                color: "#ff2b9d",
                                fontSize: 8,
                              }}
                            >
                              {value}%
                            </strong>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginTop: 17,
          paddingTop: 14,
          borderTop: "1px solid #293248",
        }}
      >
        <button
          type="button"
          className="secondary-button"
          onClick={onBack}
        >
          <ArrowLeft size={13} /> Oportunidades
        </button>

        <button
          type="button"
          className="primary-button"
          onClick={onContinue}
        >
          Continuar a Propuesta
          <ArrowRight size={13} />
        </button>
      </div>
    </div>
  );
}

function ProposalScreen({
  brief,
  discovery,
  formats,
  selectedFormats,
  selectedBroadcasterId,
  selectedFranchiseId,
  broadcasters,
  franchises,
  onBack,
  onContinue,
  onEdit,
}: {
  brief: BriefData;
  discovery: DiscoveryData;
  formats: RecommendedFormat[];
  selectedFormats: string[];
  selectedBroadcasterId: string | null;
  selectedFranchiseId: string | null;
  broadcasters: Broadcaster[];
  franchises: Franchise[];
  onBack: () => void;
  onContinue: () => void;
  onEdit: (section: string) => void;
}) {
  const chosenFormats = formats.filter((format) =>
    selectedFormats.includes(format.id)
  );

  const selectedBroadcaster =
    broadcasters.find(
      (broadcaster) => broadcaster.id === selectedBroadcasterId
    ) ?? null;

  const selectedFranchise =
    franchises.find(
      (franchise) => franchise.id === selectedFranchiseId
    ) ?? null;

  const contextText =
    discovery.moreInformation.trim() ||
    brief.context ||
    "Sin contexto adicional capturado.";

  const proposalSections = [
    {
      title: "Marca",
      content: brief.brand || "Sin marca definida.",
      source: "Brief",
    },
    {
      title: "Contexto",
      content: brief.context || contextText,
      source: "Brief + Discovery",
    },
    {
      title: "Insight",
      content:
        discovery.challengeMore ||
        discovery.moreInformation ||
        "Se construirá a partir de las señales capturadas.",
      source: "Discovery",
    },
    {
      title: "Concepto de campaña",
      content:
        brief.product ||
        "Concepto pendiente de desarrollar.",
      source: "Brief",
    },
    {
      title: "Emisora / ecosistema",
      content: `Emisora: ${
        selectedBroadcaster?.name || "Sin emisora seleccionada."
      }`,
      source: "Selección + catálogo PRISA",
    },
  ];

  // La propuesta contiene únicamente los cinco rubros definidos.
  // Todos los demás elementos están integrados dentro de "Emisora / ecosistema".
  const completedSections = proposalSections.filter((section) => {
    const value = section.content.trim();
    return (
      value &&
      !value.toLowerCase().includes("pendiente") &&
      !value.toLowerCase().includes("sin marca") &&
      !value.toLowerCase().includes("sin formatos")
    );
  }).length;

  return (
    <div className="reference-content">
      <div className="reference-proposal-header">
        <div>
          <span className="reference-eyebrow">PROPUESTA</span>
          <h2 style={{ margin: "4px 0 0" }}>
            Borrador de propuesta
          </h2>
          <p
            style={{
              margin: "5px 0 0",
              color: "#778399",
              fontSize: 9,
            }}
          >
            Estructura comercial construida a partir del
            Brief, Discovery y las selecciones realizadas.
          </p>
        </div>

        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 12,
          }}
        >
          <div className="reference-complete">
            <strong>{completedSections}/5</strong>
            <span>secciones con información</span>
          </div>

          <button
            type="button"
            className="primary-button"
          >
            <Sparkles size={14} />
            Regenerar
          </button>
        </div>
      </div>

      <div
        className="reference-card"
        style={{
          padding: 14,
          marginTop: 12,
          marginBottom: 12,
        }}
      >
        <div
          style={{
            display: "grid",
            gridTemplateColumns:
              "repeat(3, minmax(0, 1fr))",
            gap: 10,
          }}
        >
          <div>
            <span className="reference-eyebrow">
              EMISORA
            </span>
            <strong
              style={{
                display: "block",
                marginTop: 5,
                fontSize: 10,
              }}
            >
              {selectedBroadcaster?.name ||
                "Sin emisora seleccionada"}
            </strong>
          </div>

          <div>
            <span className="reference-eyebrow">
              FRANQUICIA
            </span>
            <strong
              style={{
                display: "block",
                marginTop: 5,
                fontSize: 10,
              }}
            >
              {selectedFranchise?.name ||
                "Sin franquicia"}
            </strong>
          </div>

          <div>
            <span className="reference-eyebrow">
              FORMATOS
            </span>
            <strong
              style={{
                display: "block",
                marginTop: 5,
                fontSize: 10,
              }}
            >
              {chosenFormats.length} seleccionados
            </strong>
          </div>
        </div>
      </div>

      <div className="reference-proposal-list">
        {proposalSections.map((section) => (
          <div
            className="reference-proposal-section"
            key={section.title}
          >
            <button
              type="button"
              className="reference-edit"
              onClick={() => onEdit(section.title)}
            >
              Editar
            </button>

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 7,
              }}
            >
              <h3>{section.title}</h3>

              <span
                style={{
                  color: "#637087",
                  fontSize: 7,
                }}
              >
                · {section.source}
              </span>
            </div>

            <p>{section.content}</p>
          </div>
        ))}
      </div>

      <div
        className="reference-card"
        style={{
          padding: 15,
          marginTop: 12,
          borderColor: "#293248",
        }}
      >
        <span className="reference-eyebrow">
          ESTRUCTURA COMERCIAL
        </span>

        <p
          style={{
            marginTop: 7,
            color: "#8c97aa",
            fontSize: 9,
            lineHeight: 1.6,
          }}
        >
          La propuesta se concentra únicamente en cinco rubros:
          Marca, Contexto, Insight, Concepto de campaña y
          Emisora / ecosistema. Los demás elementos quedan
          integrados dentro de Emisora / ecosistema.
        </p>
      </div>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginTop: 17,
          paddingTop: 14,
          borderTop: "1px solid #293248",
        }}
      >
        <button
          type="button"
          className="secondary-button"
          onClick={onBack}
        >
          <ArrowLeft size={13} />
          Formatos
        </button>

        <button
          type="button"
          className="primary-button"
          onClick={onContinue}
        >
          Continuar a Simulación
          <ArrowRight size={13} />
        </button>
      </div>
    </div>
  );
}

function SimulationScreen({
  regions,
  selectedRegions = [],
  selectedBroadcaster,
  selectedFranchise,
  selectedFormats,
  tariffValues,
}: {
  regions: RegionInventory[];
  selectedRegions?: string[];
  selectedBroadcaster: Broadcaster | null;
  selectedFranchise: Franchise | null;
  selectedFormats: RecommendedFormat[];
  tariffValues: Array<{ name: string; value: number }>;
}) {
  const attributableRegions = regions.filter(
    (region) => normalize(region.name) !== normalize("Sin región asignada")
  );

  const selectedRegionSet = new Set(
    selectedRegions.map((region) => normalize(region))
  );

  const targetInventory = attributableRegions
    .filter((region) => selectedRegionSet.has(normalize(region.name)))
    .reduce((sum, region) => sum + region.value, 0);

  const total = attributableRegions.reduce(
    (sum, region) => sum + region.value,
    0
  );

  const totalAll = regions.reduce(
    (sum, region) => sum + region.value,
    0
  );

  const top = Math.max(
    ...regions.map((region) => region.value),
    1
  );

  const selectedFormatTariffs = selectedFormats.map((format) => {
    const tariff = tariffValues.find(
      (item) => normalize(item.name) === normalize(format.name)
    );
    return {
      format,
      tariff: tariff?.value ?? 0,
    };
  });

  const tariffTotal = selectedFormatTariffs.reduce(
    (sum, item) => sum + item.tariff,
    0
  );

  const hasTariffData = selectedFormatTariffs.some((item) => item.tariff > 0);

  return (
    <div className="reference-content simulation-reference-page">
      <div className="reference-title-row">
        <div>
          <span className="reference-eyebrow">SIMULACIÓN</span>
          <h1>Simulación de la combinación seleccionada</h1>
          <p>
            Visualiza la selección realizada y el inventario disponible en las
            regiones objetivo. Los valores provienen de las fuentes conectadas.
          </p>
        </div>
      </div>

      <div className="reference-info reference-card simulation-info-banner">
        <MapPin size={17} />
        <div>
          <strong>Esto es inventario disponible, no una proyección de campaña.</strong>{" "}
          Los datos de geofocalización corresponden al período disponible en el
          catálogo. La simulación no convierte automáticamente inventario en
          alcance contratado.
        </div>
      </div>

      <div className="reference-sim-stats simulation-kpis">
        <div className="reference-stat simulation-kpi">
          <span>Emisora seleccionada</span>
          <strong>
            {selectedBroadcaster?.name || "Sin emisora"}
          </strong>
          <span>
            {selectedBroadcaster
              ? "Selección actual"
              : "Pendiente de selección"}
          </span>
        </div>

        <div className="reference-stat simulation-kpi">
          <span>Franquicia</span>
          <strong>
            {selectedFranchise?.name || "Sin franquicia"}
          </strong>
          <span>
            {selectedFranchise
              ? "Selección actual"
              : "La franquicia es opcional"}
          </span>
        </div>

        <div className="reference-stat simulation-kpi">
          <span>Formatos</span>
          <strong>{selectedFormats.length}</strong>
          <span>
            {selectedFormats.length
              ? selectedFormats.map((format) => format.name).join(" · ")
              : "Sin formatos seleccionados"}
          </span>
        </div>
      </div>

      <div className="reference-sim-stats simulation-kpis" style={{ marginTop: 12 }}>
        <div className="reference-stat simulation-kpi">
          <span>Regiones objetivo</span>
          <strong>{selectedRegions.length}</strong>
          <span>
            {targetInventory.toLocaleString("es-CO")} impresiones de inventario
          </span>
        </div>

        <div className="reference-stat simulation-kpi">
          <span>Inventario total atribuible</span>
          <strong>{total.toLocaleString("es-CO")}</strong>
          <span>excluye «sin región asignada»</span>
        </div>

        <div className="reference-stat simulation-kpi">
          <span>Tarifas encontradas</span>
          <strong>
            {hasTariffData ? tariffTotal.toLocaleString("es-CO") : "N/D"}
          </strong>
          <span>
            {hasTariffData
              ? "suma de coincidencias encontradas en Tarifario 2026"
              : "No hay coincidencia directa con los formatos seleccionados"}
          </span>
        </div>
      </div>

      <div className="reference-card" style={{ padding: 16, marginTop: 14 }}>
        <span className="reference-eyebrow">SELECCIÓN ACTUAL</span>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
            gap: 12,
            marginTop: 10,
          }}
        >
          <div>
            <small style={{ color: "#6f7b91", fontSize: 7 }}>EMISORA</small>
            <strong style={{ display: "block", marginTop: 5, fontSize: 10 }}>
              {selectedBroadcaster?.name || "Sin seleccionar"}
            </strong>
          </div>

          <div>
            <small style={{ color: "#6f7b91", fontSize: 7 }}>FRANQUICIA</small>
            <strong style={{ display: "block", marginTop: 5, fontSize: 10 }}>
              {selectedFranchise?.name || "Sin franquicia"}
            </strong>
          </div>

          <div>
            <small style={{ color: "#6f7b91", fontSize: 7 }}>FORMATOS</small>
            <strong style={{ display: "block", marginTop: 5, fontSize: 10 }}>
              {selectedFormats.length
                ? selectedFormats.map((format) => format.name).join(" · ")
                : "Sin formatos"}
            </strong>
          </div>
        </div>
      </div>

      <div className="reference-sim-chart simulation-chart-card">
        <div className="simulation-chart-title">Impresiones por región</div>

        <div className="reference-chart simulation-chart">
          {regions.map((region) => {
            const isTarget = selectedRegionSet.has(normalize(region.name));
            return (
              <div
                className="reference-bar-row simulation-bar-row"
                key={region.name}
              >
                <span
                  className="reference-bar-label simulation-bar-label"
                  style={{
                    color: isTarget ? "#e9d4e5" : "#8f9aaf",
                    fontWeight: isTarget ? 700 : 400,
                  }}
                >
                  {isTarget ? "● " : ""}
                  {region.name}
                </span>

                <div className="reference-bar-track simulation-bar-track">
                  <div
                    className="reference-bar simulation-bar"
                    style={{
                      width: `${Math.max(
                        0,
                        Math.min(100, (region.value / top) * 100)
                      )}%`,
                      background: isTarget ? "#ec0f82" : "#2b3244",
                    }}
                    title={`${region.name}: ${region.value.toLocaleString("es-CO")}`}
                  />
                </div>

                <span className="simulation-bar-value">
                  {region.value.toLocaleString("es-CO")}
                </span>
              </div>
            );
          })}
        </div>

        <p className="simulation-footnote">
          ● = región marcada como objetivo en el brief. Inventario total del
          período: {totalAll.toLocaleString("es-CO")}.
        </p>
      </div>
    </div>
  );
}

function PresentationScreen({
  brief,
  discovery,
  formats,
  selectedFormats,
  selectedBroadcaster,
  selectedFranchise,
  onSave,
}: {
  brief: BriefData;
  discovery: DiscoveryData;
  formats: RecommendedFormat[];
  selectedFormats: string[];
  selectedBroadcaster: Broadcaster | null;
  selectedFranchise: Franchise | null;
  onSave: () => void;
}) {
  const chosen = formats.filter((format) => selectedFormats.includes(format.id));
  const [slide, setSlide] = useState(0);

  const broadcasterName = selectedBroadcaster?.name || "Emisora por seleccionar";
  const franchiseName = selectedFranchise?.name || "Sin franquicia";
  const audienceText = [
    brief.audience,
    brief.clientType === "B2B" ? brief.businessSector : "",
    brief.clientType === "B2B" ? brief.companySize : "",
    brief.clientType === "B2C" ? brief.gender : "",
    brief.clientType === "B2C" ? brief.ageRanges.join(", ") : "",
    brief.clientType === "B2C" ? brief.socioeconomic.join(", ") : "",
  ].filter(Boolean).join(" · ") || "Audiencia por definir";

  const challengeText = [
    ...discovery.challenge,
    discovery.challengeMore,
  ].filter(Boolean).join(" · ") || brief.context || "Reto pendiente de capturar con el cliente.";

  const resultText = [
    ...discovery.result,
    discovery.resultMore,
  ].filter(Boolean).join(" · ") || brief.objective || "Resultado por definir.";

  const reactionText = [
    ...discovery.reaction,
    discovery.reactionMore,
  ].filter(Boolean).join(" · ") || "Reacción por definir.";

  const additionalText = discovery.moreInformation.trim() || "Sin información adicional registrada.";

  const selectedRegions = brief.regions.length ? brief.regions.join(" · ") : "Cobertura nacional";
  const broadcasterDetails = [
    selectedBroadcaster?.profile,
    selectedBroadcaster?.genderProfile,
    selectedBroadcaster?.ages,
    selectedBroadcaster?.socioeconomic,
    selectedBroadcaster?.interests,
  ].filter(Boolean);

  const affinityTags = selectedBroadcaster?.interests
    ? selectedBroadcaster.interests.split(/[;,|]/).map((item) => item.trim()).filter(Boolean).slice(0, 4)
    : [];

  const prevSlide = () => setSlide((current) => Math.max(0, current - 1));
  const nextSlide = () => setSlide((current) => Math.min(5, current + 1));

  const slideTitles = [
    "Portada",
    "Por qué esta campaña",
    "Insight y concepto",
    "A quién y dónde",
    "Cómo se arma",
    "Recorrido y medición",
  ];

  return (
    <div className="reference-presentation-shell">
      <div className="reference-presentation-top presentation-top-centered">
        <div>
          <h2>Boceto de la presentación</h2>
          <p className="presentation-subtitle">
            Se arma solo con lo capturado en las fases anteriores. Es un boceto de estructura y argumento, no la pieza final de diseño.
          </p>
        </div>
        <div className="reference-presentation-meta">
          <button className="reference-icon" disabled={slide === 0} onClick={prevSlide}>
            <ChevronLeft size={17} />
          </button>
          <span className="reference-slide-counter">{slide + 1} / 6</span>
          <button className="reference-icon" disabled={slide === 5} onClick={nextSlide}>
            <ChevronRight size={17} />
          </button>
        </div>
      </div>

      <div className="reference-presentation-warning">
        ⚠ El borrador de propuesta puede estar incompleto; completa las etapas anteriores para construir la presentación con información real.
      </div>

      <div className={`reference-slide slide-${slide + 1}`}>
        {slide === 0 && (
          <div className="presentation-slide-content presentation-centered-content slide-cover-content">
            <span className="reference-slide-label">PROPUESTA DE INNOVACIÓN DIGITAL</span>
            <h1>{brief.brand || "Marca sin nombre"}</h1>
            <p className="slide-large-subtitle">{brief.category || "Categoría por definir"}</p>
            <small>Área de Innovación Digital · PRISA Media Colombia</small>
          </div>
        )}

        {slide === 1 && (
          <div className="presentation-slide-content presentation-centered-content">
            <span className="reference-slide-label">POR QUÉ ESTA CAMPAÑA</span>
            <h2>La historia detrás del brief</h2>
            <div className="story-list story-list-centered">
              <div className="story-item">
                <span className="story-number">1</span>
                <div><strong>Dónde está la marca hoy</strong><p>{brief.context || "Contexto pendiente de capturar con el cliente."}</p></div>
              </div>
              <div className="story-item">
                <span className="story-number story-red">2</span>
                <div><strong>Qué se le atravesó</strong><p>{challengeText}</p></div>
              </div>
              <div className="story-item">
                <span className="story-number story-green">3</span>
                <div><strong>Qué necesita que pase</strong><p>{resultText}</p></div>
              </div>
            </div>
            <p className="slide-centered-extra">{additionalText}</p>
          </div>
        )}

        {slide === 2 && (
          <div className="presentation-slide-content presentation-centered-content">
            <span className="reference-slide-label">EL INSIGHT Y EL CONCEPTO</span>
            <div className="insight-slide-body insight-slide-centered">
              <span>EL INSIGHT</span>
              <h2>{challengeText}</h2>
              <div className="concept-box concept-box-centered">
                <span>EL CONCEPTO</span>
                <strong>{brief.product || reactionText}</strong>
              </div>
              <p className="slide-centered-extra">Resultado buscado: {resultText}</p>
            </div>
          </div>
        )}

        {slide === 3 && (
          <div className="presentation-slide-content presentation-centered-content">
            <span className="reference-slide-label">A QUIÉN LE HABLAMOS Y DÓNDE</span>
            <h2>{broadcasterName}</h2>
            <p className="slide-large-subtitle">{audienceText}</p>
            <div className="audience-slide-centered-grid">
              <div className="audience-center-card">
                <span className="slide-caption">AUDIENCIA</span>
                <p>{audienceText}</p>
                <span className="slide-caption">REGIONES OBJETIVO</span>
                <p>{selectedRegions}</p>
              </div>
              <div className="audience-center-card">
                <span className="slide-caption">PERFIL DE LA EMISORA</span>
                <p>{broadcasterDetails.join(" · ") || "Perfil disponible en el catálogo PRISA."}</p>
                <span className="slide-caption">AFINIDADES</span>
                <div className="affinity-list affinity-list-centered">
                  {(affinityTags.length ? affinityTags : ["Sin afinidades registradas"]).map((value) => (
                    <span key={value}>{value}</span>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {slide === 4 && (
          <div className="presentation-slide-content presentation-centered-content">
            <span className="reference-slide-label">CÓMO SE ARMA LA SOLUCIÓN</span>
            <h2>Dónde entra cada pieza</h2>
            <div className="solution-flow">
              <div className="flow-card"><span className="flow-icon blue">◉</span><small>EMISORA</small><strong>{broadcasterName}</strong><p>Aporta audiencia, contexto editorial e inventario.</p></div>
              <ChevronRight className="flow-arrow" size={18} />
              <div className="flow-card"><span className="flow-icon green">◎</span><small>FRANQUICIA</small><strong>{franchiseName}</strong><p>{selectedFranchise ? selectedFranchise.description || "Integra el territorio editorial seleccionado." : "La franquicia es opcional para esta campaña."}</p></div>
              <ChevronRight className="flow-arrow" size={18} />
              <div className="flow-card"><span className="flow-icon pink">✦</span><small>FORMATOS</small><strong>{chosen.length ? chosen.map((format) => format.name).join(" · ") : "Por seleccionar"}</strong><p>Define la mecánica de interacción.</p></div>
              <ChevronRight className="flow-arrow" size={18} />
              <div className="flow-card"><span className="flow-icon yellow">⚡</span><small>ACCIÓN</small><strong>{reactionText}</strong><p>Lo que la audiencia debería hacer o sentir y que podemos medir.</p></div>
            </div>
          </div>
        )}

        {slide === 5 && (
          <div className="presentation-slide-content presentation-centered-content">
            <span className="reference-slide-label">CÓMO SE VIVE Y QUÉ MEDIMOS</span>
            <div className="journey-area journey-area-centered">
              <span className="slide-caption">CÓMO SE VIVE</span>
              <div className="journey journey-centered">
                {["Descubre", "Interactúa", "Explora", "Decide", "Convierte", "Comparte"].map((step, index) => (
                  <React.Fragment key={step}>
                    <span>{step}</span>
                    {index < 5 && <ChevronRight size={14} />}
                  </React.Fragment>
                ))}
              </div>
              <span className="slide-caption">QUÉ MEDIMOS</span>
              <p className="measure-text">{chosen.length ? `La propuesta medirá la experiencia asociada a ${chosen.map((format) => format.name).join(", ")}.` : "Los KPI se definirán a partir de la propuesta y los formatos seleccionados."}</p>
              <p className="slide-centered-extra">Resultado esperado: {resultText}</p>
            </div>
          </div>
        )}

        <span className="slide-index">0{slide + 1}</span>
      </div>

      <div className="reference-thumbnails">
        {slideTitles.map((title, index) => (
          <button className={index === slide ? "reference-thumb active" : "reference-thumb"} key={title} onClick={() => setSlide(index)}>
            <span>0{index + 1}</span>{title}
          </button>
        ))}
      </div>

      <div
        className="reference-presentation-footer"
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 12,
        }}
      >
        <button className="secondary-button" type="button" onClick={onSave}>
          <CheckCircle2 size={13} /> Guardar campaña
        </button>
        <button
          className="secondary-button"
          type="button"
          onClick={() => window.history.back()}
        >
          <ArrowLeft size={13} /> Propuesta
        </button>
      </div>
    </div>
  );
}

function ConfigurationScreen({ userName, onChangeName, catalogCount, tariffCount }: { userName: string; onChangeName: () => void; catalogCount: number; tariffCount: number }) {
  return <div className="reference-content"><div className="reference-title-row"><div><span className="reference-eyebrow">CONFIGURACIÓN</span><h1>Configuración</h1><p>Administra tu espacio y revisa las fuentes conectadas.</p></div></div><div className="reference-grid2"><div className="reference-panel"><User size={19}/><h2>Usuario</h2><div style={{ marginTop: 15, display: "flex", justifyContent: "space-between", alignItems: "center" }}><strong style={{ fontSize: 11 }}>{userName || "Sin nombre"}</strong><button className="secondary-button" onClick={onChangeName}>Cambiar</button></div></div><div className="reference-panel"><FileText size={19}/><h2>Fuentes conectadas</h2><div className="reference-list"><div className="reference-list-row"><div className="reference-list-icon"><CheckCircle2 size={15}/></div><div><strong>Catálogo PRISA</strong><span>Productos digitales, emisoras y geofocalización</span></div><div>{catalogCount}</div></div><div className="reference-list-row"><div className="reference-list-icon"><CheckCircle2 size={15}/></div><div><strong>Tarifario 2026</strong><span>Tarifas comerciales disponibles</span></div><div>{tariffCount}</div></div></div></div></div></div>;
}

export default function App() {
  const [currentPage, setCurrentPage] = useState<PageKey>("campaigns");
  const [userName, setUserName] = useState("");
  const [showNameModal, setShowNameModal] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [brief, setBrief] = useState<BriefData>(createEmptyBrief);
  const [discovery, setDiscovery] = useState<DiscoveryData>(createEmptyDiscovery);
  const [catalogRows, setCatalogRows] = useState<ExcelRow[]>([]);
  const [tariffRows, setTariffRows] = useState<ExcelRow[]>([]);
  const [projects, setProjects] = useState<Project[]>([]);
  const [currentProjectId, setCurrentProjectId] = useState<string | null>(null);
  const [selectedOpportunity, setSelectedOpportunity] = useState<string | null>(null);
  const [selectedBroadcasterId, setSelectedBroadcasterId] = useState<string | null>(
    () => readSelectionStorage(SELECTED_BROADCASTER_STORAGE_KEY)
  );
  const [selectedFranchiseId, setSelectedFranchiseId] = useState<string | null>(
    () => readSelectionStorage(SELECTED_FRANCHISE_STORAGE_KEY)
  );
  const [selectedFormats, setSelectedFormats] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);

  const updateSelectedBroadcaster = (id: string | null) => {
    setSelectedBroadcasterId(id);
    writeSelectionStorage(SELECTED_BROADCASTER_STORAGE_KEY, id);
  };

  const updateSelectedFranchise = (id: string | null) => {
    setSelectedFranchiseId(id);
    writeSelectionStorage(SELECTED_FRANCHISE_STORAGE_KEY, id);
  };
  const [error, setError] = useState("");

  useEffect(() => {
    const saved = localStorage.getItem("prisa_user_name");
    if (saved) setUserName(saved);
    else setShowNameModal(true);
  }, []);

  useEffect(() => {
    setProjects(loadSavedProjects());
  }, []);

  useEffect(() => {
    let mounted = true;
    Promise.all([
      readWorkbook(CATALOG_FILE),
      readWorkbook(TARIFF_FILE),
      readGeolocationWorkbook(CATALOG_FILE),
    ])
      .then(([catalog, tariff, geo]) => {
        if (!mounted) return;
        setCatalogRows([
          ...catalog,
          ...geo.map((row) => ({ ...row, __sheet: "Geofocalización" })),
        ]);
        setTariffRows(tariff);
      })
      .catch((err) => {
        if (mounted) setError(err instanceof Error ? err.message : "No se pudieron cargar los datos.");
      })
      .finally(() => mounted && setLoading(false));
    return () => { mounted = false; };
  }, []);

  const digitalProducts = useMemo<Product[]>(() => {
    return catalogRows
      .filter((row) => normalize(row.__sheet) === normalize("Productos digitales"))
      .map((row, index) => ({
        id: rowValue(row, ["ID"]) || `PD-${index + 1}`,
        name: rowValue(row, ["Nombre"]),
        category: rowValue(row, ["Categoría / Tipo"]),
        description: rowValue(row, ["¿Qué es?"]),
        howWorks: rowValue(row, ["¿Cómo funciona?"]),
        application: rowValue(row, ["¿Para qué aplica?"]),
        objectives: rowValue(row, ["Objetivos normalizados"]),
        audiences: rowValue(row, ["Audiencias normalizadas"]),
        pains: rowValue(row, ["Dolores normalizados"]),
        channels: rowValue(row, ["Canales normalizados"]),
        tags: rowValue(row, ["Tags"]),
        sheet: text(row.__sheet),
      }))
      .filter((row) => row.name);
  }, [catalogRows]);

  const broadcasters = useMemo<Broadcaster[]>(() => {
    return catalogRows
      .filter((row) => normalize(row.__sheet) === normalize("Emisoras"))
      .map((row, index) => ({
        id: rowValue(row, ["ID"]) || `EM-${index + 1}`,
        name: rowValue(row, ["Nombre"]),
        profile: rowValue(row, ["Cobertura"]),
        genderProfile: rowValue(row, ["Perfil demográfico"]),
        ages: rowValue(row, ["Edades"]),
        socioeconomic: rowValue(row, ["Nivel socioeconómico"]),
        interests: rowValue(row, ["Afinidades / temas de interés"]),
        coverage: rowValue(row, ["Cobertura"]),
      }))
      .filter((row) => row.name);
  }, [catalogRows]);

  const tariffValues = useMemo(() => {
    const rows = tariffRows.filter(
      (row) => normalize(row.__sheet) === normalize("Tarifas 2026")
    );

    return rows
      .map((row) => {
        const name = rowValue(row, [
          "FORMATOS",
          "formato",
          "concepto",
          "nombre",
        ]);
        const value = parseNumber(
          rowValue(row, [
            "Caracol",
            "CRC",
            "Valor CPM",
            "Valor",
            "tarifa",
          ])
        );
        return { name, value };
      })
      .filter((row) => row.name);
  }, [tariffRows]);

  const franchises = useMemo<Franchise[]>(() => {
    return catalogRows
      .filter((row) => normalize(row.__sheet) === normalize("Franquicias"))
      .map((row, index) => ({
        id: rowValue(row, ["ID", "id"]) || `FR-${index + 1}`,
        name: rowValue(row, ["Nombre", "Franquicia", "nombre"]),
        category: rowValue(row, ["Categoría", "Categoría / Tipo", "Tipo"]),
        description: rowValue(row, ["¿Qué es?", "Descripción", "Descripcion"]),
        application: rowValue(row, ["¿Para qué aplica?", "Aplicación", "Aplicacion"]),
        objectives: rowValue(row, ["Objetivos normalizados", "Objetivos", "Objetivo"]),
        audiences: rowValue(row, ["Audiencias normalizadas", "Audiencias", "Audiencia"]),
        tags: rowValue(row, ["Tags", "Etiquetas"]),
        territories: rowValue(row, ["Territorios", "Regiones", "Cobertura"]),
        sheet: text(row.__sheet),
      }))
      .filter((row) => row.name);
  }, [catalogRows]);

  const recommendedFranchises = useMemo<RecommendedFranchise[]>(() => {
    const objectiveText = [
      brief.objective,
      ...discovery.result,
      discovery.resultMore,
    ].filter(Boolean).join(" ");

    const audienceText = [
      brief.audience,
      brief.clientType,
      brief.businessSector,
      brief.companySize,
      brief.gender,
      ...brief.ageRanges,
      ...brief.socioeconomic,
      ...discovery.audience,
      discovery.audienceMore,
    ].filter(Boolean).join(" ");

    const contextText = [
      brief.brand,
      brief.category,
      brief.product,
      brief.context,
      ...discovery.challenge,
      discovery.challengeMore,
      ...discovery.reaction,
      discovery.reactionMore,
      discovery.moreInformation,
    ].filter(Boolean).join(" ");

    const regionText = brief.regions.join(" ");

    return franchises
      .map((franchise) => {
        const objective = overlapScore(
          objectiveText,
          [franchise.objectives, franchise.application, franchise.description].join(" ")
        );

        const audience = overlapScore(
          audienceText,
          [franchise.audiences, franchise.tags, franchise.description].join(" ")
        );

        const context = overlapScore(
          contextText,
          [franchise.description, franchise.application, franchise.tags, franchise.objectives].join(" ")
        );

        const category = brief.category
          ? overlapScore(
              brief.category,
              [franchise.category, franchise.application, franchise.description].join(" ")
            )
          : 55;

        const geography = regionText
          ? overlapScore(regionText, franchise.territories) || 50
          : 55;

        const affinity = capScore(
          objective * 0.30 +
          audience * 0.25 +
          context * 0.20 +
          category * 0.15 +
          geography * 0.10
        );

        const reasons = [
          objective >= 60 ? "objetivo compatible" : "",
          audience >= 60 ? "audiencia compatible" : "",
          context >= 60 ? "contexto compatible" : "",
          category >= 60 ? "categoría compatible" : "",
          geography >= 60 ? "territorio compatible" : "",
        ].filter(Boolean);

        return {
          id: franchise.id,
          name: franchise.name,
          affinity,
          description: franchise.description || franchise.application || "Franquicia disponible en el catálogo PRISA.",
          category: franchise.category,
          breakdown: {
            objective: Math.round(objective),
            audience: Math.round(audience),
            context: Math.round(context),
            category: Math.round(category),
            geography: Math.round(geography),
          },
          reasons: reasons.length
            ? reasons.map((item) => item.charAt(0).toUpperCase() + item.slice(1))
            : ["Coincidencia disponible dentro de la hoja Franquicias."],
        };
      })
      .sort((a, b) => b.affinity - a.affinity);
  }, [brief, discovery, franchises]);

  const regionInventory = useMemo<RegionInventory[]>(() => {
    const rows = catalogRows.filter(
      (row) => normalize(row.__sheet) === normalize("Geofocalización")
    );

    if (!rows.length) return [];

    const preferredNames = [
      "Bogotá / Centro",
      "Antioquia",
      "Costa Caribe",
      "Pacífico",
      "Santanderes",
      "Eje Cafetero",
      "Centro Oriente",
      "Llanos Orientales",
      "Sur de Colombia",
      "Sin región asignada",
      "Insular",
      "Amazonía",
    ];

    const available = preferredNames
      .filter((name) =>
        rows.some((row) =>
          Object.prototype.hasOwnProperty.call(row, name)
        )
      )
      .map((name) => ({
        name,
        value: rows.reduce(
          (sum, row) => sum + parseNumber(row[name]),
          0
        ),
      }));

    if (available.length) return available;

    const firstRow = rows[0];
    return Object.keys(firstRow)
      .filter(
        (key) =>
          ![
            "Formato publicitario",
            "Total",
            "__sheet",
          ].map(normalize).includes(normalize(key))
      )
      .map((key) => ({
        name: key,
        value: rows.reduce(
          (sum, row) => sum + parseNumber(row[key]),
          0
        ),
      }))
      .filter((item) => item.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [catalogRows]);

  const completion = useMemo(() => {
    const fields = [brief.brand, brief.website, brief.category, brief.product, brief.objective, brief.audience, brief.timingStart, brief.budget];
    return Math.round((fields.filter(Boolean).length / fields.length) * 100);
  }, [brief]);

  const recommendedFormats = useMemo<RecommendedFormat[]>(() => {
    const discoveryObjective = [
      brief.objective,
      ...discovery.result,
      discovery.resultMore,
    ].filter(Boolean).join(" ");

    const discoveryAudience = [
      brief.audience,
      brief.clientType,
      brief.businessSector,
      brief.companySize,
      brief.gender,
      ...brief.ageRanges,
      ...brief.socioeconomic,
      ...discovery.audience,
      discovery.audienceMore,
    ].filter(Boolean).join(" ");

    const discoveryContext = [
      brief.brand,
      brief.category,
      brief.product,
      brief.context,
      ...discovery.challenge,
      discovery.challengeMore,
      ...discovery.reaction,
      discovery.reactionMore,
      discovery.moreInformation,
    ].filter(Boolean).join(" ");

    return digitalProducts.map((product) => {
      const productObjectiveText = [
        product.objectives,
        product.channels,
        product.description,
      ].join(" ");

      const productAudienceText = [
        product.audiences,
        product.tags,
        product.description,
      ].join(" ");

      const productContextText = [
        product.pains,
        product.description,
        product.application,
        product.tags,
      ].join(" ");

      const objective = overlapScore(
        discoveryObjective,
        productObjectiveText
      );

      const audience = overlapScore(
        discoveryAudience,
        productAudienceText
      );

      const contextScore = overlapScore(
        discoveryContext,
        productContextText
      );

      const categoryScore = brief.category
        ? overlapScore(
            brief.category,
            [product.category, product.application, product.description].join(" ")
          )
        : 55;

      const keyword = overlapScore(
        [
          brief.objective,
          brief.category,
          brief.product,
          ...discovery.challenge,
          ...discovery.result,
        ].join(" "),
        [
          product.name,
          product.category,
          product.objectives,
          product.audiences,
          product.pains,
          product.tags,
        ].join(" ")
      );

      const score = capScore(
        objective * 0.30 +
        audience * 0.25 +
        contextScore * 0.20 +
        categoryScore * 0.15 +
        keyword * 0.10
      );

      const reasons = [
        objective >= 60 ? "objetivo compatible" : "",
        audience >= 60 ? "audiencia compatible" : "",
        contextScore >= 60 ? "contexto y dolores compatibles" : "",
        categoryScore >= 60 ? "categoría/aplicación compatible" : "",
      ].filter(Boolean);

      return {
        id: product.id,
        name: product.name,
        category: product.category,
        affinity: score,
        description: product.description,
        reason:
          reasons.length
            ? reasons.map((item) => item.charAt(0).toUpperCase() + item.slice(1)).join(" · ")
            : "Coincidencia disponible dentro del catálogo PRISA.",
        source: product.sheet,
        breakdown: {
          objective: Math.round(objective),
          audience: Math.round(audience),
          context: Math.round(contextScore),
          category: Math.round(categoryScore),
          keyword: Math.round(keyword),
        },
      };
    })
      .sort((a, b) => b.affinity - a.affinity)
      .slice(0, 8);
  }, [brief, discovery, digitalProducts]);

  const recommendedBroadcasters = useMemo<RecommendedBroadcaster[]>(() => {
    const objectiveText = [
      brief.objective,
      ...discovery.result,
      discovery.resultMore,
    ].filter(Boolean).join(" ");

    const audienceText = [
      brief.audience,
      brief.gender,
      ...brief.ageRanges,
      ...brief.socioeconomic,
      ...discovery.audience,
      discovery.audienceMore,
    ].filter(Boolean).join(" ");

    const contextText = [
      brief.brand,
      brief.category,
      brief.product,
      brief.context,
      ...discovery.challenge,
      discovery.challengeMore,
      ...discovery.reaction,
      discovery.reactionMore,
      discovery.moreInformation,
    ].filter(Boolean).join(" ");

    const regionText = brief.regions.join(" ");

    return broadcasters
      .map((broadcaster) => {
        const objective = overlapScore(
          objectiveText,
          [
            broadcaster.interests,
            broadcaster.profile,
          ].join(" ")
        );

        const audience = overlapScore(
          audienceText,
          [
            broadcaster.genderProfile,
            broadcaster.ages,
            broadcaster.socioeconomic,
            broadcaster.interests,
          ].join(" ")
        );

        const geography = regionText
          ? overlapScore(
              regionText,
              broadcaster.coverage
            )
          : 55;

        const contextScore = overlapScore(
          contextText,
          [
            broadcaster.profile,
            broadcaster.interests,
          ].join(" ")
        );

        const searchableClientType = normalize(
          [
            broadcaster.profile,
            broadcaster.interests,
          ].join(" ")
        );

        let clientType = 60;

        if (brief.clientType === "B2B") {
          const b2bTerms = [
            "empresa",
            "empresas",
            "negocio",
            "negocios",
            "profesional",
            "corporativo",
            "b2b",
          ];

          clientType = b2bTerms.some((term) =>
            searchableClientType.includes(term)
          )
            ? 100
            : 55;
        } else if (brief.clientType === "B2C") {
          clientType = 80;
        }

        const affinity = capScore(
          objective * 0.30 +
          audience * 0.25 +
          geography * 0.20 +
          contextScore * 0.15 +
          clientType * 0.10
        );

        const reasons = [
          objective >= 60 ? "objetivo alineado" : "",
          audience >= 60 ? "perfil de audiencia compatible" : "",
          geography >= 60 ? "cobertura geográfica compatible" : "",
          contextScore >= 60 ? "perfil/intereses compatibles" : "",
          brief.clientType === "B2B" && clientType >= 90
            ? "perfil empresarial compatible"
            : "",
        ].filter(Boolean);

        return {
          id: broadcaster.id,
          name: broadcaster.name,
          affinity,
          profile: broadcaster.profile,
          coverage: broadcaster.coverage,
          breakdown: {
            objective: Math.round(objective),
            audience: Math.round(audience),
            geography: Math.round(geography),
            context: Math.round(contextScore),
            clientType: Math.round(clientType),
          },
          reasons:
            reasons.length
              ? reasons.map((item) => item.charAt(0).toUpperCase() + item.slice(1))
              : ["Coincidencia disponible dentro de la hoja Emisoras."],
        };
      })
      .sort((a, b) => b.affinity - a.affinity);
  }, [brief, discovery, broadcasters]);

  const opportunities = useMemo<Opportunity[]>(() => [
    { id: "pain", title: "Dolores identificados", description: brief.context ? "Problemas o tensiones que la audiencia comparte." : "Ningún dolor capturado todavía.", count: brief.context ? 1 : 0, kind: "pain" },
    { id: "territory", title: "Territorios de oportunidad", description: brief.regions.length ? "Regiones marcadas como objetivo." : "Categorías de franquicia con mayor encaje.", count: brief.regions.length, kind: "territory" },
    { id: "audience", title: "Audiencias clave", description: brief.audience || "Segmentos prioritarios.", count: brief.audience ? 1 : 0, kind: "audience" },
    { id: "space", title: "Espacios estratégicos", description: brief.product || "Dónde podemos activar.", count: brief.product ? 1 : 0, kind: "space" },
  ], [brief]);

  const resetCampaign = () => {
    // Una campaña nueva debe empezar completamente limpia, pero sin tocar
    // las campañas que ya están guardadas en localStorage.
    setCurrentProjectId(null);
    setBrief(createEmptyBrief());
    setDiscovery(createEmptyDiscovery());
    setSelectedFormats([]);
    setSelectedOpportunity(null);
    updateSelectedBroadcaster(null);
    updateSelectedFranchise(null);
    setSearchTerm("");
    setCurrentPage("brief");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const saveCampaign = () => {
    const projectName =
      brief.campaignName.trim() ||
      brief.brand.trim() ||
      "Nueva campaña";

    const now = new Date();
    const formattedDate = now.toLocaleString("es-CO", {
      dateStyle: "short",
      timeStyle: "short",
    });

    const project: Project = {
      id: currentProjectId ?? `project-${Date.now()}`,
      name: projectName,
      status: "LIVE",
      updatedAt: formattedDate,
      brief,
      discovery,
      selectedOpportunity,
      selectedBroadcasterId,
      selectedFranchiseId,
      selectedFormats,
    };

    setProjects((previous) => {
      const exists = previous.some((item) => item.id === project.id);
      const next = exists
        ? previous.map((item) => (item.id === project.id ? project : item))
        : [project, ...previous];

      localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
      return next;
    });

    setCurrentProjectId(project.id);
  };

  const openProject = (project: Project) => {
    setCurrentProjectId(project.id);
    setBrief(project.brief);
    setDiscovery(project.discovery ?? createEmptyDiscovery());
    setSelectedOpportunity(project.selectedOpportunity ?? null);
    updateSelectedBroadcaster(project.selectedBroadcasterId ?? null);
    updateSelectedFranchise(project.selectedFranchiseId ?? null);
    setSelectedFormats(project.selectedFormats ?? []);
    setCurrentPage("summary");
  };

  const refreshProjects = () => {
    setProjects(loadSavedProjects());
  };

  const completionPage = () => {
    switch (currentPage) {
      case "campaigns":
        return (
          <CampaignsScreen
            projects={projects}
            onNewCampaign={resetCampaign}
            onOpenProject={openProject}
            onRefresh={refreshProjects}
            searchTerm={searchTerm}
          />
        );
      case "summary":
        return (
          <SummaryScreen
            brief={brief}
            formats={recommendedFormats}
            opportunities={opportunities}
            regions={regionInventory}
            discovery={discovery}
            selectedOpportunity={selectedOpportunity}
            selectedBroadcasterId={selectedBroadcasterId}
            selectedFranchiseId={selectedFranchiseId}
            selectedFormats={selectedFormats}
            onNavigate={setCurrentPage}
          />
        );
      case "brief":
        return (
          <BriefScreen
            brief={brief}
            completion={completion}
            setBrief={setBrief}
            onInsights={() => setCurrentPage("discovery")}
          />
        );
      case "discovery":
        return (
          <DiscoveryScreen
            discovery={discovery}
            setDiscovery={setDiscovery}
            onBack={() => setCurrentPage("brief")}
            onContinue={() => setCurrentPage("insights")}
          />
        );
      case "insights":
        return (
          <InsightsScreen
            brief={brief}
            discovery={discovery}
            onBack={() => setCurrentPage("discovery")}
            onContinue={() => setCurrentPage("opportunities")}
          />
        );
      case "opportunities": return <OpportunitiesScreen brief={brief} discovery={discovery} selected={selectedOpportunity} onSelect={setSelectedOpportunity} onBack={() => setCurrentPage("insights")} onContinue={() => setCurrentPage("formats")} />;
      case "formats":
        return (
          <FormatsScreen
            formats={recommendedFormats}
            selected={selectedFormats}
            setSelected={setSelectedFormats}
            broadcasters={broadcasters}
            recommendedBroadcasters={recommendedBroadcasters}
            selectedBroadcasterId={selectedBroadcasterId}
            setSelectedBroadcasterId={updateSelectedBroadcaster}
            franchises={franchises}
            recommendedFranchises={recommendedFranchises}
            selectedFranchiseId={selectedFranchiseId}
            setSelectedFranchiseId={updateSelectedFranchise}
            onBack={() => setCurrentPage("opportunities")}
            onContinue={() => {
              updateSelectedBroadcaster(selectedBroadcasterId);
              updateSelectedFranchise(selectedFranchiseId);
              setCurrentPage("proposal");
            }}
          />
        );
      case "proposal":
        return (
          <ProposalScreen
            brief={brief}
            discovery={discovery}
            formats={recommendedFormats}
            selectedFormats={selectedFormats}
            selectedBroadcasterId={selectedBroadcasterId}
            selectedFranchiseId={selectedFranchiseId}
            broadcasters={broadcasters}
            franchises={franchises}
            onBack={() => setCurrentPage("formats")}
            onContinue={() => setCurrentPage("simulation")}
            onEdit={(section) => {
              const formatSections = ["Emisora / ecosistema", "Franquicia", "Formatos recomendados", "Solución propuesta", "Mecánica", "Contenido", "Tecnología"];
              const discoverySections = ["Problema", "Insight", "Concepto de campaña", "Objetivo", "Audiencia", "KPI"];
              if (section === "Marca" || section === "Contexto") setCurrentPage("brief");
              else if (formatSections.includes(section)) setCurrentPage("formats");
              else if (discoverySections.includes(section)) setCurrentPage("discovery");
              else if (section === "Oportunidad") setCurrentPage("opportunities");
              else setCurrentPage("brief");
            }}
          />
        );
      case "simulation":
        return (
          <SimulationScreen
            regions={regionInventory}
            selectedRegions={brief.regions}
            selectedBroadcaster={
              broadcasters.find(
                (item) => item.id === selectedBroadcasterId
              ) ?? null
            }
            selectedFranchise={
              franchises.find(
                (item) => item.id === selectedFranchiseId
              ) ?? null
            }
            selectedFormats={recommendedFormats.filter((format) =>
              selectedFormats.includes(format.id)
            )}
            tariffValues={tariffValues}
          />
        );
      case "presentation":
        return (
          <PresentationScreen
            brief={brief}
            discovery={discovery}
            formats={recommendedFormats}
            selectedFormats={selectedFormats}
            selectedBroadcaster={broadcasters.find((item) => item.id === selectedBroadcasterId) ?? null}
            selectedFranchise={franchises.find((item) => item.id === selectedFranchiseId) ?? null}
            onSave={saveCampaign}
          />
        );
      case "configuration": return <ConfigurationScreen userName={userName} onChangeName={() => setShowNameModal(true)} catalogCount={digitalProducts.length} tariffCount={tariffRows.length} />;
      default: return null;
    }
  };

  if (loading) {
    return <><ReferenceStyles /><div className="full-loading"><Logo /><div className="loading-spinner" /><h2>Cargando información</h2><p>Conectando catálogo y tarifario PRISA...</p></div></>;
  }

  if (error) {
    return <><ReferenceStyles /><div className="full-loading"><X size={31}/><h2>No pudimos cargar los datos</h2><p>{error}</p><p>Verifica que los dos Excel estén dentro de <strong>public/data</strong>.</p><button className="primary-button" onClick={() => window.location.reload()}>Reintentar</button></div></>;
  }

  return (
    <>
      <ReferenceStyles />
      {showNameModal && <NameModal value={userName} onConfirm={(name) => { if (name) { setUserName(name); localStorage.setItem("prisa_user_name", name); } setShowNameModal(false); }} />}
      <div className="app-shell app-reference">
        <Sidebar currentPage={currentPage} setCurrentPage={setCurrentPage} onNewCampaign={resetCampaign} userName={userName} collapsed={collapsed} setCollapsed={setCollapsed} mobileMenuOpen={mobileMenuOpen} setMobileMenuOpen={setMobileMenuOpen} />
        <div className="app-main reference-main">
          <Header currentPage={currentPage} userName={userName} searchTerm={searchTerm} setSearchTerm={setSearchTerm} onOpenMobileMenu={() => setMobileMenuOpen(true)} />
          <main className="main-container" style={{ maxWidth: "none", padding: 0 }}>
            {completionPage()}
          </main>
        </div>
      </div>
    </>
  );
}

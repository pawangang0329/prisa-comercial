import * as XLSX from "xlsx";

/* ============================================================
   ARCHIVOS
   ============================================================ */

const CATALOG_FILE =
  "/data/Catalogo_PRISA_Normalizado_Actualizado.xlsx";

const TARIFF_FILE =
  "/data/Tarifario 2026 1.7 Comercial.xlsx";

/* ============================================================
   TIPOS GENERALES
   ============================================================ */

export type ExcelRow = Record<string, unknown>;

export type RawSheet = {
  name: string;
  rows: ExcelRow[];
};

export type PrisaCatalog = {
  products: ProductData[];
  franchises: FranchiseData[];
  broadcasters: BroadcasterData[];
  geolocation: RegionData[];
};

export type PrisaData = {
  catalog: PrisaCatalog;
  tariff: TariffData[];
};

/* ============================================================
   TIPOS DEL CATÁLOGO
   ============================================================ */

export type ProductData = {
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

export type FranchiseData = {
  id: string;
  name: string;
  description: string;
  category: string;
  broadcaster: string;
  audience: string;
  objective: string;
  components: string;
  inventory: string;
  interactions: string;
  reach: string;
  tags: string;
  sheet: string;
};

export type BroadcasterData = {
  id: string;
  name: string;
  profile: string;
  genderProfile: string;
  ages: string;
  socioeconomic: string;
  interests: string;
  coverage: string;
};

export type RegionData = {
  name: string;
  value: number;
  department?: string;
  region?: string;
  format?: string;
};

export type TariffData = {
  sheet: string;
  values: ExcelRow;
};

/* ============================================================
   UTILIDADES
   ============================================================ */

export function normalizeText(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function cleanText(value: unknown): string {
  return String(value ?? "").trim();
}

function toNumber(value: unknown): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const raw = String(value ?? "")
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(",", ".");

  const result = Number(raw);

  return Number.isFinite(result) ? result : 0;
}

/* ============================================================
   COLUMNAS
   ============================================================ */

function findColumn(
  row: ExcelRow,
  candidates: string[]
): string {
  const keys = Object.keys(row);

  /* Coincidencia exacta */
  for (const candidate of candidates) {
    const normalizedCandidate =
      normalizeText(candidate);

    const exact = keys.find(
      (key) =>
        normalizeText(key) === normalizedCandidate
    );

    if (exact) {
      return exact;
    }
  }

  /* Coincidencia parcial */
  for (const candidate of candidates) {
    const normalizedCandidate =
      normalizeText(candidate);

    const partial = keys.find((key) => {
      const normalizedKey = normalizeText(key);

      return (
        normalizedKey.includes(normalizedCandidate) ||
        normalizedCandidate.includes(normalizedKey)
      );
    });

    if (partial) {
      return partial;
    }
  }

  return "";
}

function getValue(
  row: ExcelRow,
  candidates: string[]
): string {
  const column = findColumn(row, candidates);

  if (!column) {
    return "";
  }

  return cleanText(row[column]);
}

/* ============================================================
   LEER UN LIBRO EXCEL
   ============================================================ */

async function fetchWorkbook(
  url: string
): Promise<XLSX.WorkBook> {
  const response = await fetch(url, {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(
      `No se pudo cargar el archivo: ${url}`
    );
  }

  const buffer = await response.arrayBuffer();

  return XLSX.read(buffer, {
    type: "array",
    cellDates: true,
  });
}

/* ============================================================
   CONVERTIR UNA HOJA A OBJETOS
   ============================================================ */

function sheetToRows(
  workbook: XLSX.WorkBook,
  sheetName: string
): ExcelRow[] {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    return [];
  }

  return XLSX.utils.sheet_to_json<ExcelRow>(sheet, {
    defval: "",
    raw: false,
  });
}

/* ============================================================
   PRODUCTOS DIGITALES
   ============================================================ */

function parseProducts(
  workbook: XLSX.WorkBook
): ProductData[] {
  const rows = sheetToRows(
    workbook,
    "Productos digitales"
  );

  return rows
    .map((row, index) => ({
      id:
        getValue(row, [
          "id",
          "codigo",
          "cod",
          "ID Producto",
        ]) || `PROD-${index + 1}`,

      name:
        getValue(row, [
          "nombre",
          "producto",
          "nombre producto",
          "formato",
          "nombre del producto",
        ]) || `Producto ${index + 1}`,

      category: getValue(row, [
        "categoria",
        "categoría",
        "tipo",
        "vertical",
      ]),

      description: getValue(row, [
        "descripcion",
        "descripción",
        "detalle",
        "descripcion comercial",
      ]),

      howWorks: getValue(row, [
        "como funciona",
        "cómo funciona",
        "funcionamiento",
        "mecanica",
        "mecánica",
      ]),

      application: getValue(row, [
        "aplicacion",
        "aplicación",
        "uso",
        "donde aplica",
      ]),

      objectives: getValue(row, [
        "objetivo",
        "objetivos",
        "objetivo comercial",
      ]),

      audiences: getValue(row, [
        "audiencia",
        "audiencias",
        "publico",
        "público",
      ]),

      pains: getValue(row, [
        "dolor",
        "dolores",
        "necesidad",
        "problema",
      ]),

      channels: getValue(row, [
        "canal",
        "canales",
        "medio",
      ]),

      tags: getValue(row, [
        "tags",
        "etiquetas",
        "keywords",
        "palabras clave",
      ]),

      sheet: "Productos digitales",
    }))
    .filter(
      (product) =>
        product.name.trim() !== ""
    );
}

/* ============================================================
   FRANQUICIAS
   ============================================================ */

function parseFranchises(
  workbook: XLSX.WorkBook
): FranchiseData[] {
  const rows = sheetToRows(
    workbook,
    "Franquicias"
  );

  return rows
    .map((row, index) => ({
      id:
        getValue(row, [
          "id",
          "codigo",
          "cod",
          "ID",
        ]) || `FR-${index + 1}`,

      name:
        getValue(row, [
          "nombre",
          "franquicia",
          "nombre franquicia",
          "nombre de franquicia",
        ]) || `Franquicia ${index + 1}`,

      description: getValue(row, [
        "descripcion",
        "descripción",
        "detalle",
      ]),

      category: getValue(row, [
        "categoria",
        "categoría",
        "tipo",
      ]),

      broadcaster: getValue(row, [
        "emisora",
        "emisor",
        "radio",
        "marca",
      ]),

      audience: getValue(row, [
        "audiencia",
        "audiencias",
        "publico",
        "público",
      ]),

      objective: getValue(row, [
        "objetivo",
        "objetivos",
      ]),

      components: getValue(row, [
        "componentes",
        "componente",
        "elementos",
      ]),

      inventory: getValue(row, [
        "inventario",
        "alcance",
        "disponibilidad",
      ]),

      interactions: getValue(row, [
        "interacciones",
        "interaccion",
        "interacción",
      ]),

      reach: getValue(row, [
        "reach",
        "cobertura",
        "reproducciones",
        "usuarios",
      ]),

      tags: getValue(row, [
        "tags",
        "etiquetas",
        "keywords",
      ]),

      sheet: "Franquicias",
    }))
    .filter(
      (franchise) =>
        franchise.name.trim() !== ""
    );
}

/* ============================================================
   EMISORAS
   ============================================================ */

function parseBroadcasters(
  workbook: XLSX.WorkBook
): BroadcasterData[] {
  const rows = sheetToRows(
    workbook,
    "Emisoras"
  );

  return rows
    .map((row, index) => ({
      id:
        getValue(row, [
          "id",
          "codigo",
          "cod",
          "ID Emisora",
        ]) || `EM-${String(index + 1).padStart(3, "0")}`,

      name:
        getValue(row, [
          "emisora",
          "nombre",
          "nombre emisora",
          "marca",
        ]) || `Emisora ${index + 1}`,

      profile: getValue(row, [
        "perfil",
        "perfil de audiencia",
        "descripcion",
        "descripción",
      ]),

      genderProfile: getValue(row, [
        "genero",
        "género",
        "perfil genero",
        "perfil de genero",
        "perfil de género",
      ]),

      ages: getValue(row, [
        "edades",
        "rango edad",
        "edad",
        "perfil edad",
      ]),

      socioeconomic: getValue(row, [
        "nivel socioeconomico",
        "nivel socioeconómico",
        "nse",
        "socioeconomico",
        "socioeconómico",
      ]),

      interests: getValue(row, [
        "intereses",
        "afinidades",
        "temas",
        "intereses y afinidades",
      ]),

      coverage: getValue(row, [
        "cobertura",
        "territorio",
        "alcance",
        "presencia",
      ]),
    }))
    .filter(
      (broadcaster) =>
        broadcaster.name.trim() !== ""
    );
}

/* ============================================================
   GEOLOCALIZACIÓN
   ============================================================ */

function parseGeolocation(
  workbook: XLSX.WorkBook
): RegionData[] {
  const sheet =
    workbook.Sheets["Geofocalización"];

  if (!sheet) {
    return [];
  }

  /*
   * Algunas hojas de geofocalización no tienen
   * el encabezado en la primera fila.
   *
   * Por eso buscamos la fila que contiene
   * "Formato Publicitario".
   */

  const matrix = XLSX.utils.sheet_to_json<
    unknown[]
  >(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });

  const headerIndex = matrix.findIndex(
    (row) =>
      Array.isArray(row) &&
      row.some(
        (cell) =>
          normalizeText(cell) ===
          "formato publicitario"
      )
  );

  if (headerIndex < 0) {
    return [];
  }

  const headers = (
    matrix[headerIndex] as unknown[]
  ).map((cell) => cleanText(cell));

  const result: RegionData[] = [];

  for (
    let i = headerIndex + 1;
    i < matrix.length;
    i += 1
  ) {
    const row = matrix[i] as unknown[];

    if (
      !row ||
      !row.some(
        (cell) => cleanText(cell) !== ""
      )
    ) {
      continue;
    }

    const object: ExcelRow = {};

    headers.forEach((header, columnIndex) => {
      if (!header) return;

      object[header] =
        row[columnIndex] ?? "";
    });

    const region =
      getValue(object, [
        "region",
        "región",
        "macroregion",
        "macroregión",
      ]) ||
      getValue(object, [
        "departamento",
      ]) ||
      "Sin región asignada";

    const department = getValue(
      object,
      ["departamento"]
    );

    const format = getValue(object, [
      "formato publicitario",
      "formato",
      "producto",
    ]);

    const value = toNumber(
      object[
        findColumn(object, [
          "impresiones",
          "inventario",
          "alcance",
          "cantidad",
          "valor",
        ])
      ]
    );

    result.push({
      name: region,
      value,
      department,
      region,
      format,
    });
  }

  return result;
}

/* ============================================================
   TARIFARIO
   ============================================================ */

function parseTariff(
  workbook: XLSX.WorkBook
): TariffData[] {
  const result: TariffData[] = [];

  workbook.SheetNames.forEach(
    (sheetName) => {
      const rows =
        sheetToRows(
          workbook,
          sheetName
        );

      rows.forEach((row) => {
        result.push({
          sheet: sheetName,
          values: row,
        });
      });
    }
  );

  return result;
}

/* ============================================================
   CARGAR TODO EL CATÁLOGO
   ============================================================ */

export async function loadCatalog(): Promise<PrisaCatalog> {
  const workbook =
    await fetchWorkbook(CATALOG_FILE);

  return {
    products: parseProducts(workbook),
    franchises: parseFranchises(workbook),
    broadcasters:
      parseBroadcasters(workbook),
    geolocation:
      parseGeolocation(workbook),
  };
}

/* ============================================================
   CARGAR TARIFARIO
   ============================================================ */

export async function loadTariff(): Promise<TariffData[]> {
  const workbook =
    await fetchWorkbook(TARIFF_FILE);

  return parseTariff(workbook);
}

/* ============================================================
   CARGAR TODOS LOS DATOS PRISA
   ============================================================ */

export async function loadPrisaData(): Promise<PrisaData> {
  const [catalog, tariff] =
    await Promise.all([
      loadCatalog(),
      loadTariff(),
    ]);

  return {
    catalog,
    tariff,
  };
}

/* ============================================================
   DIAGNÓSTICO
   ============================================================ */

export async function diagnosePrisaData() {
  const data =
    await loadPrisaData();

  console.group(
    "PRISA MEDIA · DATOS NORMALIZADOS"
  );

  console.log(
    "Productos:",
    data.catalog.products.length
  );

  console.log(
    "Franquicias:",
    data.catalog.franchises.length
  );

  console.log(
    "Emisoras:",
    data.catalog.broadcasters.length
  );

  console.log(
    "Geolocalización:",
    data.catalog.geolocation.length
  );

  console.log(
    "Tarifario:",
    data.tariff.length
  );

  console.log(
    "Emisoras:",
    data.catalog.broadcasters
  );

  console.log(
    "Franquicias:",
    data.catalog.franchises
  );

  console.log(
    "Productos:",
    data.catalog.products
  );

  console.groupEnd();

  return data;
}
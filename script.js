"use strict";

/* =========================================================
DASHBOARD CAPEX
========================================================= */

const PROJECTS_URL = "data/projects.json";
const SAVINGS_RATE = 60;

let projects = [];
let currentProject = null;
let currentWorkbook = null;

let financialData = {
budget: 0,
real: 0,
difference: 0,
rows: []
};

let purchaseData = [];

let cashflowData = {
periods: [],
rows: []
};

const charts = {};


/* =========================================================
INICIO
========================================================= */

document.addEventListener("DOMContentLoaded", async () => {
setupNavigation();
setupEvents();

try {
showLoading(true, "Cargando iniciativas...");
await loadProjects();
} catch (error) {
console.error("ERROR INICIAL:", error);
showMessage(
"No se pudo iniciar el dashboard: " + error.message,
"error"
);
} finally {
showLoading(false);
}
});


/* =========================================================
EVENTOS
========================================================= */

function setupEvents() {
const projectSelect = document.getElementById("projectSelect");

if (projectSelect) {
projectSelect.addEventListener("change", async function () {
if (this.value) {
await loadProject(this.value);
}
});
}

const refreshButton = document.getElementById("refreshButton");

if (refreshButton) {
refreshButton.addEventListener("click", async function () {
if (currentProject) {
await loadProject(currentProject.id);
}
});
}

const budgetSearch = document.getElementById("budgetSearch");

if (budgetSearch) {
budgetSearch.addEventListener("input", renderBudgetTable);
}

const purchaseSearch = document.getElementById("purchaseSearch");

if (purchaseSearch) {
purchaseSearch.addEventListener("input", renderPurchaseTable);
}

const purchaseStatusFilter =
document.getElementById("purchaseStatusFilter");

if (purchaseStatusFilter) {
purchaseStatusFilter.addEventListener(
"change",
renderPurchaseTable
);
}

const pendingSearch = document.getElementById("pendingSearch");

if (pendingSearch) {
pendingSearch.addEventListener("input", renderPendingTable);
}

const expandCashflow = document.getElementById("expandCashflow");

if (expandCashflow) {
expandCashflow.addEventListener("click", function () {
toggleAllCashflow(true);
});
}

const collapseCashflow =
document.getElementById("collapseCashflow");

if (collapseCashflow) {
collapseCashflow.addEventListener("click", function () {
toggleAllCashflow(false);
});
}
}


/* =========================================================
NAVEGACIÓN
========================================================= */

function setupNavigation() {
const buttons = document.querySelectorAll("#navTabs .nav-item");

buttons.forEach(function (button) {
button.addEventListener("click", function () {
buttons.forEach(function (item) {
item.classList.remove("active");
});

button.classList.add("active");

document.querySelectorAll(".page").forEach(function (page) {
page.classList.remove("active");
});

const target = document.getElementById(
"page-" + button.dataset.page
);

if (target) {
target.classList.add("active");
}
});
});
}


/* =========================================================
PROJECTS.JSON
========================================================= */

async function loadProjects() {
setText("projectLoadStatus", "Cargando iniciativas...");

const response = await fetch(
PROJECTS_URL + "?v=" + Date.now(),
{ cache: "no-store" }
);

if (!response.ok) {
throw new Error(
"No se pudo abrir projects.json. HTTP " + response.status
);
}

const data = await response.json();

if (Array.isArray(data)) {
projects = data;
} else if (data && Array.isArray(data.projects)) {
projects = data.projects;
} else {
throw new Error(
"projects.json no contiene una lista válida de iniciativas."
);
}

populateProjectSelect();

if (!projects.length) {
setText("projectLoadStatus", "Sin iniciativas");
showMessage(
"No hay iniciativas configuradas en projects.json.",
"warning"
);
return;
}

const firstProject = projects[0];

const select = document.getElementById("projectSelect");

if (select) {
select.value = String(firstProject.id);
}

await loadProject(firstProject.id);
}


function populateProjectSelect() {
const select = document.getElementById("projectSelect");

if (!select) {
return;
}

select.innerHTML = "";

projects.forEach(function (project) {
const option = document.createElement("option");

option.value = String(project.id);
option.textContent =
project.name || project.id || "Iniciativa";

select.appendChild(option);
});
}


/* =========================================================
CARGAR INICIATIVA
========================================================= */

async function loadProject(projectId) {
const project = projects.find(function (item) {
return String(item.id) === String(projectId);
});

if (!project) {
showMessage(
"La iniciativa seleccionada no existe.",
"error"
);
return;
}

currentProject = project;

resetDashboard();
updateProjectHeader(project);
hideMessage();

try {
showLoading(
true,
"Leyendo " + (project.name || project.id) + "..."
);

setText("projectLoadStatus", "Leyendo Excel...");

if (!project.file) {
throw new Error(
"La iniciativa no tiene archivo Excel definido en projects.json."
);
}

const separator = project.file.indexOf("?") >= 0 ? "&" : "?";

const url =
project.file +
separator +
"v=" +
Date.now();

console.log("Cargando Excel:", url);

const response = await fetch(url, {
cache: "no-store"
});

if (!response.ok) {
throw new Error(
"No se pudo abrir " +
project.file +
". HTTP " +
response.status
);
}

if (typeof XLSX === "undefined") {
throw new Error(
"No se cargó la librería XLSX."
);
}

const buffer = await response.arrayBuffer();

currentWorkbook = XLSX.read(buffer, {
type: "array",
cellDates: true
});

console.log(
"Hojas encontradas:",
currentWorkbook.SheetNames
);

parseWorkbook(currentWorkbook);

renderDashboard();

setText("projectLoadStatus", "Datos cargados");
} catch (error) {
console.error("ERROR CARGANDO PROYECTO:", error);

setText(
"projectLoadStatus",
"Error al leer Excel"
);

showMessage(
"No se pudo leer " +
(project.name || project.id) +
". " +
error.message,
"error"
);
} finally {
showLoading(false);
}
}


/* =========================================================
CABECERA
========================================================= */

function updateProjectHeader(project) {
setText(
"topProjectName",
project.name || project.id || "Dashboard CAPEX"
);

setText(
"topProjectSubtitle",
"Seguimiento integral de presupuesto, compras y flujo de caja."
);

const opening = parseISODate(project.openingDate);

if (!opening) {
setText("openingDate", "—");
setText("daysToOpening", "—");
setText("executiveOpeningDate", "—");
setText("executiveDays", "—");
return;
}

const formatted = formatDate(opening);

setText("openingDate", formatted);
setText("executiveOpeningDate", formatted);

const today = new Date();
today.setHours(0, 0, 0, 0);

opening.setHours(0, 0, 0, 0);

const difference =
opening.getTime() - today.getTime();

const days = Math.ceil(
difference / 86400000
);

setText("daysToOpening", days);
setText("executiveDays", days);
}


/* =========================================================
WORKBOOK
========================================================= */

function parseWorkbook(workbook) {
financialData = {
budget: 0,
real: 0,
difference: 0,
rows: []
};

purchaseData = [];

cashflowData = {
periods: [],
rows: []
};

const financialSheet = findSheet(
workbook,
"Presupuesto vs Real"
);

const purchaseSheet = findSheet(
workbook,
"BD Plan Detallado"
);

const cashflowSheet = findSheet(
workbook,
"Flujo de Caja"
);

const missing = [];

if (financialSheet) {
try {
financialData =
parseFinancialSheet(financialSheet);

console.log(
"Presupuesto vs Real:",
financialData
);
} catch (error) {
console.error(
"Error leyendo Presupuesto vs Real:",
error
);
}
} else {
missing.push("Presupuesto vs Real");
}

if (purchaseSheet) {
try {
purchaseData =
parsePurchaseSheet(purchaseSheet);

console.log(
"BD Plan Detallado:",
purchaseData
);
} catch (error) {
console.error(
"Error leyendo BD Plan Detallado:",
error
);
}
} else {
missing.push("BD Plan Detallado");
}

if (cashflowSheet) {
try {
cashflowData =
parseCashflowSheet(cashflowSheet);

console.log(
"Flujo de Caja:",
cashflowData
);
} catch (error) {
console.error(
"Error leyendo Flujo de Caja:",
error
);
}
} else {
missing.push("Flujo de Caja");
}

if (missing.length) {
showMessage(
"No se encontraron estas hojas: " +
missing.join(", "),
"warning"
);
}
}


function findSheet(workbook, target) {
const normalizedTarget =
normalizeText(target);

const sheetName =
workbook.SheetNames.find(function (name) {
return (
normalizeText(name) ===
normalizedTarget
);
});

if (!sheetName) {
return null;
}

return workbook.Sheets[sheetName];
}


function sheetToMatrix(sheet) {
return XLSX.utils.sheet_to_json(sheet, {
header: 1,
defval: null,
raw: true
});
}


/* =========================================================
PRESUPUESTO VS REAL
========================================================= */

function parseFinancialSheet(sheet) {
const matrix = sheetToMatrix(sheet);

const result = {
budget: 0,
real: 0,
difference: 0,
rows: []
};

if (!matrix.length) {
return result;
}

let totalRow = -1;

for (let r = 0; r < matrix.length; r++) {
const rowText = normalizeText(
matrix[r]
.filter(function (value) {
return value !== null;
})
.join(" ")
);

if (
rowText.indexOf(
"total general estimado"
) >= 0
) {
totalRow = r;
break;
}
}

const totals = findFinancialTotals(
matrix,
totalRow
);

if (totals) {
result.budget = totals.budget;
result.real = totals.real;
result.difference =
totals.difference;
}

const columns =
detectFinancialColumns(matrix);

let summaryRow = -1;

for (let r = 0; r < matrix.length; r++) {
const text = normalizeText(
matrix[r]
.filter(function (value) {
return value !== null;
})
.join(" ")
);

if (
text.indexOf(
"resumen de partidas"
) >= 0
) {
summaryRow = r;
break;
}
}

const start =
summaryRow >= 0
? summaryRow + 1
: 0;

for (
let r = start;
r < matrix.length;
r++
) {
const row = matrix[r] || [];

const name =
getFinancialLabel(row);

if (!name) {
continue;
}

const normalizedName =
normalizeText(name);

if (
normalizedName.indexOf(
"total general estimado"
) >= 0 ||
normalizedName.indexOf(
"resumen de partidas"
) >= 0
) {
continue;
}

const values =
getFinancialRowValues(
row,
columns
);

if (
values.budget === 0 &&
values.real === 0
) {
continue;
}

result.rows.push({
name: name,
budget: values.budget,
real: values.real,
difference:
values.difference
});
}

if (
result.budget === 0 &&
result.real === 0
) {
const majorRows =
result.rows.filter(function (row) {
return isMajorFinancialCategory(
row.name
);
});

if (majorRows.length) {
result.budget =
majorRows.reduce(function (
total,
row
) {
return total + row.budget;
}, 0);

result.real =
majorRows.reduce(function (
total,
row
) {
return total + row.real;
}, 0);

result.difference =
result.budget -
result.real;
}
}

if (
result.difference === 0 &&
result.budget !== result.real
) {
result.difference =
result.budget -
result.real;
}

return result;
}


function findFinancialTotals(
matrix,
totalRow
) {
if (totalRow < 0) {
return null;
}

const rowsToCheck = [
totalRow,
totalRow + 1,
totalRow - 1,
totalRow + 2
];

for (
let i = 0;
i < rowsToCheck.length;
i++
) {
const index = rowsToCheck[i];

if (
index < 0 ||
index >= matrix.length
) {
continue;
}

const numbers = [];

(matrix[index] || []).forEach(
function (value, column) {
const number = toNumber(value);

if (
Number.isFinite(number) &&
Math.abs(number) > 100
) {
numbers.push({
value: number,
column: column
});
}
}
);

if (numbers.length >= 3) {
const last =
numbers.slice(-3);

return {
budget: last[0].value,
real: last[1].value,
difference: last[2].value,
columns: [
last[0].column,
last[1].column,
last[2].column
]
};
}

if (numbers.length === 2) {
return {
budget: numbers[0].value,
real: numbers[1].value,
difference:
numbers[0].value -
numbers[1].value,
columns: [
numbers[0].column,
numbers[1].column,
null
]
};
}
}

return null;
}


function detectFinancialColumns(matrix) {
let budgetColumn = null;
let realColumn = null;
let differenceColumn = null;

const limit = Math.min(
matrix.length,
25
);

for (let r = 0; r < limit; r++) {
const row = matrix[r] || [];

for (
let c = 0;
c < row.length;
c++
) {
const value =
normalizeText(row[c]);

if (
value === "presupuestado" ||
value === "presupuesto"
) {
budgetColumn = c;
}

if (value === "real") {
realColumn = c;
}

if (value === "diferencia") {
differenceColumn = c;
}
}
}

return {
budgetColumn: budgetColumn,
realColumn: realColumn,
differenceColumn:
differenceColumn
};
}


function getFinancialLabel(row) {
const limit = Math.min(
row.length,
6
);

for (let c = 0; c < limit; c++) {
if (
typeof row[c] === "string" &&
cleanText(row[c])
) {
const value =
cleanText(row[c]);

const normalized =
normalizeText(value);

if (
normalized !== "usd" &&
normalized !== "usd$" &&
normalized !== "us$"
) {
return value;
}
}
}

return "";
}


function getFinancialRowValues(
row,
columns
) {
let budget = 0;
let real = 0;
let difference = 0;

if (
columns.budgetColumn !== null
) {
budget = toNumber(
row[columns.budgetColumn]
);
}

if (
columns.realColumn !== null
) {
real = toNumber(
row[columns.realColumn]
);
}

if (
columns.differenceColumn !== null
) {
difference = toNumber(
row[
columns.differenceColumn
]
);
}

if (
budget === 0 &&
real === 0
) {
const numbers = [];

row.forEach(function (value) {
const number = toNumber(value);

if (
Number.isFinite(number) &&
Math.abs(number) > 0
) {
numbers.push(number);
}
});

if (numbers.length >= 3) {
const values =
numbers.slice(-3);

budget = values[0];
real = values[1];
difference = values[2];
} else if (
numbers.length >= 2
) {
const values =
numbers.slice(-2);

budget = values[0];
real = values[1];
difference =
budget - real;
}
}

if (
difference === 0 &&
budget !== real
) {
difference =
budget - real;
}

return {
budget: budget,
real: real,
difference: difference
};
}


function isMajorFinancialCategory(name) {
const value = normalizeText(name);

return (
value === "edificacion" ||
value === "equipamiento" ||
value === "terreno"
);
}


/* =========================================================
BD PLAN DETALLADO
========================================================= */

function parsePurchaseSheet(sheet) {
const matrix = sheetToMatrix(sheet);

if (!matrix.length) {
return [];
}

const headerRow =
findPurchaseHeaderRow(matrix);

if (headerRow < 0) {
console.warn(
"No se encontró el encabezado de BD Plan Detallado."
);
return [];
}

const headers =
(matrix[headerRow] || []).map(
function (value) {
return normalizeText(value);
}
);

const columns = {
partida: findColumn(
headers,
["partidas", "partida"]
),

item: findColumn(
headers,
["item"]
),

order: findColumn(
headers,
["orden de compra"]
),

date: findColumn(
headers,
["fecha"]
),

quantity: findColumn(
headers,
["cant.", "cant", "cantidad"]
),

netDOP: findColumnWithWords(
headers,
["valor neto", "dop"]
),

budgetDOP: findColumnWithWords(
headers,
["orden interna", "dop"]
),

comment: findColumn(
headers,
[
"comentario",
"observacion",
"observaciones"
]
)
};

console.log(
"Columnas BD:",
columns
);

const result = [];

for (
let r = headerRow + 1;
r < matrix.length;
r++
) {
const row = matrix[r] || [];

const orderValue =
getCell(row, columns.order);

const status =
classifyPurchaseStatus(
orderValue
);

if (
status === "ignore" ||
status === "other"
) {
continue;
}

const partida =
cleanText(
getCell(
row,
columns.partida
)
);

const item =
cleanText(
getCell(
row,
columns.item
)
);

result.push({
partida: partida,

item: item,

order: cleanText(orderValue),

date: getCell(
row,
columns.date
),

quantity: toNumber(
getCell(
row,
columns.quantity
)
),

netDOP: toNumber(
getCell(
row,
columns.netDOP
)
),

budgetDOP: toNumber(
getCell(
row,
columns.budgetDOP
)
),

comment: cleanText(
getCell(
row,
columns.comment
)
),

status: status
});
}

return result;
}


function findPurchaseHeaderRow(matrix) {
const limit = Math.min(
matrix.length,
30
);

for (let r = 0; r < limit; r++) {
const values =
(matrix[r] || []).map(
function (value) {
return normalizeText(value);
}
);

const hasOrder =
values.some(function (value) {
return (
value.indexOf(
"orden de compra"
) >= 0
);
});

const hasPartida =
values.some(function (value) {
return (
value.indexOf("partida") >= 0
);
});

const hasItem =
values.some(function (value) {
return (
value.indexOf("item") >= 0
);
});

if (
hasOrder &&
(hasPartida || hasItem)
) {
return r;
}
}

return -1;
}


function classifyPurchaseStatus(value) {
const raw = cleanText(value);
const text = normalizeText(raw);

if (!text) {
return "other";
}

if (
text.indexOf(
"pendiente de compra"
) >= 0 ||
text === "pendiente"
) {
return "pending";
}

if (
text === "stock" ||
text.indexOf("stock") >= 0
) {
return "stock";
}

if (text === "plan") {
return "ignore";
}

/*
Comprobamos si la OC contiene al menos
6 dígitos sin usar expresiones regulares.
*/

let digits = "";

for (
let i = 0;
i < raw.length;
i++
) {
const character = raw[i];

if (
character >= "0" &&
character <= "9"
) {
digits += character;
}
}

if (digits.length >= 6) {
return "po";
}

return "other";
}


/* =========================================================
FLUJO DE CAJA
========================================================= */

function parseCashflowSheet(sheet) {
const matrix = sheetToMatrix(sheet);

const result = {
periods: [],
rows: []
};

if (!matrix.length) {
return result;
}

const periodInfo =
detectCashflowPeriods(matrix);

if (!periodInfo) {
console.warn(
"No se detectaron períodos en Flujo de Caja."
);
return result;
}

result.periods =
periodInfo.periods;

const firstPeriodColumn =
result.periods[0].column;

let currentParent = null;

for (
let r = periodInfo.rowIndex + 1;
r < matrix.length;
r++
) {
const row = matrix[r] || [];

const label =
getCashflowLabel(
row,
firstPeriodColumn
);

if (!label) {
continue;
}

const values =
result.periods.map(
function (period) {
return toNumber(
row[period.column]
);
}
);

const total =
values.reduce(function (
sum,
value
) {
return sum + value;
}, 0);

if (total === 0) {
continue;
}

const parent =
isCashflowParent(label);

if (parent) {
currentParent = label;
}

result.rows.push({
id: "cashflow-" + r,
label: label,
values: values,
total: total,
isParent: parent,
parent:
parent
? null
: currentParent
});
}

return result;
}


function detectCashflowPeriods(matrix) {
const limit = Math.min(
matrix.length,
30
);

for (let r = 0; r < limit; r++) {
const row = matrix[r] || [];

const candidates = [];

for (
let c = 0;
c < row.length;
c++
) {
const value =
toNumber(row[c]);

if (
Number.isInteger(value) &&
value >= 1 &&
value <= 60
) {
candidates.push({
period: value,
column: c
});
}
}

if (candidates.length < 3) {
continue;
}

let sequential = 1;

for (
let i = 1;
i < candidates.length;
i++
) {
if (
candidates[i].period ===
candidates[i - 1].period + 1
) {
sequential++;
}
}

if (sequential >= 3) {
return {
rowIndex: r,

periods:
candidates.map(
function (item) {
return {
period: item.period,
column: item.column,
label:
periodToMonth(
item.period
)
};
}
)
};
}
}

return null;
}


function getCashflowLabel(
row,
firstPeriodColumn
) {
const parts = [];

for (
let c = 0;
c < firstPeriodColumn;
c++
) {
const value =
cleanText(row[c]);

const normalized =
normalizeText(value);

if (
value &&
normalized !== "usd" &&
normalized !== "usd$" &&
normalized !== "rd$" &&
normalized !== "dop"
) {
parts.push(value);
}
}

return parts.join(" ").trim();
}


function isCashflowParent(label) {
const value = cleanText(label);

/*
Detecta 1.00, 2.00, 3.00...
*/

const firstSpace =
value.indexOf(" ");

const code =
firstSpace >= 0
? value.substring(0, firstSpace)
: value;

if (code.endsWith(".00")) {
return true;
}

/*
Detecta M1, M2, etc.
*/

if (
code.length >= 2 &&
code[0].toUpperCase() === "M"
) {
const rest =
code.substring(1);

if (
rest &&
!Number.isNaN(Number(rest))
) {
return true;
}
}

return false;
}


function periodToMonth(period) {
const monthNames = [
"Ene",
"Feb",
"Mar",
"Abr",
"May",
"Jun",
"Jul",
"Ago",
"Sep",
"Oct",
"Nov",
"Dic"
];

const index = period - 1;

const year =
2025 +
Math.floor(index / 12);

const month =
index % 12;

return (
monthNames[month] +
" " +
year
);
}


/* =========================================================
RENDER GENERAL
========================================================= */

function renderDashboard() {
renderFinancial();
renderPurchases();
renderCashflow();
renderTracking();
}


/* =========================================================
FINANCIERO
========================================================= */

function renderFinancial() {
const budget =
financialData.budget || 0;

const real =
financialData.real || 0;

const difference =
financialData.difference !== 0
? financialData.difference
: budget - real;

const execution =
budget !== 0
? real / budget
: 0;

const savingsDOP =
difference * SAVINGS_RATE;

setText(
"homeBudget",
formatUSD(budget)
);

setText(
"homeReal",
formatUSD(real)
);

setText(
"homeDifference",
formatUSD(difference)
);

setText(
"homeExecution",
formatPercent(execution)
);

setSavings(
"homeSavingsDOP",
savingsDOP
);


setText(
"executiveBudget",
formatUSD(budget)
);

setText(
"executiveReal",
formatUSD(real)
);

setText(
"executiveDifference",
formatUSD(difference)
);

setText(
"executiveExecution",
formatPercent(execution)
);

setText(
"executiveProgressText",
formatPercent(execution)
);

setSavings(
"executiveSavingsDOP",
savingsDOP
);


setText(
"budgetTotal",
formatUSD(budget)
);

setText(
"realTotal",
formatUSD(real)
);

setText(
"differenceTotal",
formatUSD(difference)
);

setText(
"budgetExecution",
formatPercent(execution)
);

setSavings(
"budgetSavingsDOP",
savingsDOP
);


const progress =
document.getElementById(
"executiveProgressBar"
);

if (progress) {
const percent =
Math.max(
0,
Math.min(
execution * 100,
100
)
);

progress.style.width =
percent + "%";
}

renderBudgetTable();
renderHomeCategoryTable();
renderFinancialCharts();
}


function setSavings(id, value) {
const element =
document.getElementById(id);

if (!element) {
return;
}

element.textContent =
formatDOP(Math.abs(value));

const card =
element.closest(".kpi-card");

if (!card) {
return;
}

const label =
card.querySelector(".kpi-label");

if (label) {
label.textContent =
value >= 0
? "Ahorro RD$"
: "Sobreejecución RD$";
}
}


/* =========================================================
TABLA PRESUPUESTO
========================================================= */

function renderBudgetTable() {
const tbody =
document.querySelector(
"#budgetTable tbody"
);

if (!tbody) {
return;
}

const input =
document.getElementById(
"budgetSearch"
);

const search =
normalizeText(
input ? input.value : ""
);

const rows =
financialData.rows.filter(
function (row) {
return (
!search ||
normalizeText(
row.name
).indexOf(search) >= 0
);
}
);

if (!rows.length) {
tbody.innerHTML =
emptyRow(5);
return;
}

tbody.innerHTML =
rows.map(function (row) {
const execution =
row.budget !== 0
? row.real / row.budget
: 0;

return (
"<tr>" +
"<td>" +
escapeHTML(row.name) +
"</td>" +
'<td class="number">' +
formatUSD(row.budget) +
"</td>" +
'<td class="number">' +
formatUSD(row.real) +
"</td>" +
'<td class="number">' +
formatUSD(row.difference) +
"</td>" +
'<td class="number">' +
formatPercent(execution) +
"</td>" +
"</tr>"
);
}).join("");
}


function renderHomeCategoryTable() {
const tbody =
document.querySelector(
"#homeCategoryTable tbody"
);

if (!tbody) {
return;
}

let rows =
financialData.rows.filter(
function (row) {
return isMajorFinancialCategory(
row.name
);
}
);

if (!rows.length) {
rows =
financialData.rows.slice(
0,
10
);
}

if (!rows.length) {
tbody.innerHTML =
emptyRow(5);
return;
}

tbody.innerHTML =
rows.map(function (row) {
const execution =
row.budget !== 0
? row.real / row.budget
: 0;

return (
"<tr>" +
"<td>" +
escapeHTML(row.name) +
"</td>" +
'<td class="number">' +
formatUSD(row.budget) +
"</td>" +
'<td class="number">' +
formatUSD(row.real) +
"</td>" +
'<td class="number">' +
formatUSD(row.difference) +
"</td>" +
'<td class="number">' +
formatPercent(execution) +
"</td>" +
"</tr>"
);
}).join("");
}


/* =========================================================
COMPRAS
========================================================= */

function renderPurchases() {
const po =
purchaseData.filter(
function (row) {
return row.status === "po";
}
);

const pending =
purchaseData.filter(
function (row) {
return row.status === "pending";
}
);

const stock =
purchaseData.filter(
function (row) {
return row.status === "stock";
}
);

const poValue =
po.reduce(function (
total,
row
) {
return total + row.netDOP;
}, 0);

const pendingValue =
pending.reduce(function (
total,
row
) {
return total + row.budgetDOP;
}, 0);

const pendingCategories =
new Set();

pending.forEach(function (row) {
if (row.partida) {
pendingCategories.add(
normalizeText(row.partida)
);
}
});


setText(
"homePOCount",
po.length
);

setText(
"homePendingCount",
pending.length
);

setText(
"executivePO",
po.length
);

setText(
"executivePending",
pending.length
);

setText(
"purchasePOCount",
po.length
);

setText(
"purchasePendingCount",
pending.length
);

setText(
"purchaseStockCount",
stock.length
);

setText(
"purchasePOValue",
formatDOP(poValue)
);

setText(
"pendingTotal",
pending.length
);

setText(
"pendingCategoryCount",
pendingCategories.size
);

setText(
"pendingValue",
formatDOP(pendingValue)
);

renderPurchaseTable();
renderPendingTable();
renderPurchaseCharts();
}


function renderPurchaseTable() {
const tbody =
document.querySelector(
"#purchaseTable tbody"
);

if (!tbody) {
return;
}

const searchInput =
document.getElementById(
"purchaseSearch"
);

const filterInput =
document.getElementById(
"purchaseStatusFilter"
);

const search =
normalizeText(
searchInput
? searchInput.value
: ""
);

const status =
filterInput
? filterInput.value
: "all";

const rows =
purchaseData.filter(
function (row) {
if (
status !== "all" &&
row.status !== status
) {
return false;
}

if (!search) {
return true;
}

const text =
normalizeText(
row.partida +
" " +
row.item +
" " +
row.order
);

return (
text.indexOf(search) >= 0
);
}
);

if (!rows.length) {
tbody.innerHTML =
emptyRow(7);
return;
}

tbody.innerHTML =
rows.map(function (row) {
return (
"<tr>" +
"<td>" +
escapeHTML(row.partida) +
"</td>" +
"<td>" +
escapeHTML(row.item) +
"</td>" +
"<td>" +
escapeHTML(row.order) +
"</td>" +
"<td>" +
formatExcelDate(row.date) +
"</td>" +
'<td class="number">' +
formatNumber(row.quantity) +
"</td>" +
'<td class="number">' +
formatDOP(row.netDOP) +
"</td>" +
"<td>" +
purchaseStatusBadge(
row.status
) +
"</td>" +
"</tr>"
);
}).join("");
}


function renderPendingTable() {
const tbody =
document.querySelector(
"#pendingTable tbody"
);

if (!tbody) {
return;
}

const input =
document.getElementById(
"pendingSearch"
);

const search =
normalizeText(
input ? input.value : ""
);

const rows =
purchaseData.filter(
function (row) {
if (
row.status !== "pending"
) {
return false;
}

if (!search) {
return true;
}

const text =
normalizeText(
row.partida +
" " +
row.item
);

return (
text.indexOf(search) >= 0
);
}
);

if (!rows.length) {
tbody.innerHTML =
emptyRow(5);
return;
}

tbody.innerHTML =
rows.map(function (row) {
return (
"<tr>" +
"<td>" +
escapeHTML(row.partida) +
"</td>" +
"<td>" +
escapeHTML(row.item) +
"</td>" +
'<td class="number">' +
formatNumber(row.quantity) +
"</td>" +
'<td class="number">' +
formatDOP(row.budgetDOP) +
"</td>" +
"<td>" +
escapeHTML(row.comment) +
"</td>" +
"</tr>"
);
}).join("");
}


function purchaseStatusBadge(status) {
if (status === "po") {
return (
'<span class="status-badge status-po">' +
"Con OC" +
"</span>"
);
}

if (status === "pending") {
return (
'<span class="status-badge status-pending">' +
"Pendiente de compra" +
"</span>"
);
}

if (status === "stock") {
return (
'<span class="status-badge status-stock">' +
"Stock" +
"</span>"
);
}

return "";
}


/* =========================================================
FLUJO DE CAJA
========================================================= */

function renderCashflow() {
const periods =
cashflowData.periods;

const rows =
cashflowData.rows;

const total =
calculateCashflowTotal(rows);

setText(
"cashflowTotal",
formatDOP(total)
);

setText(
"cashflowPeriodCount",
periods.length
);

const parents =
rows.filter(function (row) {
return row.isParent;
});

setText(
"cashflowCategoryCount",
parents.length ||
rows.length
);

renderCashflowTable();
renderCashflowChart();
}


function calculateCashflowTotal(rows) {
const parents =
rows.filter(function (row) {
return row.isParent;
});

const source =
parents.length
? parents
: rows;

return source.reduce(function (
total,
row
) {
return total + row.total;
}, 0);
}


function renderCashflowTable() {
const thead =
document.getElementById(
"cashflowTableHead"
);

const tbody =
document.getElementById(
"cashflowTableBody"
);

if (!thead || !tbody) {
return;
}

const periods =
cashflowData.periods;

let header =
"<tr><th>Partida</th>";

periods.forEach(function (period) {
header +=
'<th class="number">' +
escapeHTML(period.label) +
"</th>";
});

header +=
'<th class="number">Total</th></tr>';

thead.innerHTML = header;

if (!cashflowData.rows.length) {
tbody.innerHTML =
emptyRow(
periods.length + 2
);
return;
}

let html = "";

cashflowData.rows.forEach(
function (row) {
if (row.isParent) {
html +=
'<tr class="cashflow-group-row">' +
"<td>" +
'<button class="cashflow-toggle" type="button" data-target="' +
escapeHTML(row.label) +
'">' +
'<span class="toggle-symbol">−</span>' +
"<span>" +
escapeHTML(row.label) +
"</span>" +
"</button>" +
"</td>";

row.values.forEach(
function (value) {
html +=
'<td class="number">' +
formatDOPCompact(value) +
"</td>";
}
);

html +=
'<td class="number">' +
formatDOPCompact(
row.total
) +
"</td>" +
"</tr>";
} else {
html +=
'<tr class="cashflow-child-row" data-parent-label="' +
escapeHTML(
row.parent || ""
) +
'">' +
'<td class="cashflow-child-label">' +
escapeHTML(row.label) +
"</td>";

row.values.forEach(
function (value) {
html +=
'<td class="number">' +
formatDOPCompact(value) +
"</td>";
}
);

html +=
'<td class="number">' +
formatDOPCompact(
row.total
) +
"</td>" +
"</tr>";
}
}
);

tbody.innerHTML = html;

const buttons =
tbody.querySelectorAll(
".cashflow-toggle"
);

buttons.forEach(function (button) {
button.addEventListener(
"click",
function () {
toggleCashflowGroup(
button
);
}
);
});
}


function toggleCashflowGroup(button) {
const parent =
button.dataset.target;

const children =
document.querySelectorAll(
".cashflow-child-row"
);

let hide = false;

for (
let i = 0;
i < children.length;
i++
) {
if (
children[i].dataset
.parentLabel === parent
) {
hide =
children[i].style.display !==
"none";
break;
}
}

children.forEach(function (child) {
if (
child.dataset.parentLabel ===
parent
) {
child.style.display =
hide ? "none" : "";
}
});

const symbol =
button.querySelector(
".toggle-symbol"
);

if (symbol) {
symbol.textContent =
hide ? "+" : "−";
}
}


function toggleAllCashflow(expand) {
document
.querySelectorAll(
".cashflow-child-row"
)
.forEach(function (row) {
row.style.display =
expand ? "" : "none";
});

document
.querySelectorAll(
".cashflow-toggle .toggle-symbol"
)
.forEach(function (symbol) {
symbol.textContent =
expand ? "−" : "+";
});
}


/* =========================================================
SEGUIMIENTO
========================================================= */

function renderTracking() {
const tbody =
document.querySelector(
"#trackingTable tbody"
);

if (!tbody) {
return;
}

let rows =
financialData.rows.filter(
function (row) {
return isMajorFinancialCategory(
row.name
);
}
);

if (!rows.length) {
rows =
financialData.rows.slice(
0,
15
);
}

if (!rows.length) {
tbody.innerHTML =
emptyRow(7);
return;
}

tbody.innerHTML =
rows.map(function (row) {
const related =
purchaseData.filter(
function (purchase) {
return matchPartida(
purchase.partida,
row.name
);
}
);

const po =
related.filter(
function (item) {
return item.status === "po";
}
).length;

const pending =
related.filter(
function (item) {
return (
item.status === "pending"
);
}
).length;

const stock =
related.filter(
function (item) {
return (
item.status === "stock"
);
}
).length;

const execution =
row.budget !== 0
? row.real / row.budget
: 0;

return (
"<tr>" +
"<td>" +
escapeHTML(row.name) +
"</td>" +
'<td class="number">' +
formatUSD(row.budget) +
"</td>" +
'<td class="number">' +
formatUSD(row.real) +
"</td>" +
'<td class="number">' +
formatPercent(execution) +
"</td>" +
'<td class="number">' +
po +
"</td>" +
'<td class="number">' +
pending +
"</td>" +
'<td class="number">' +
stock +
"</td>" +
"</tr>"
);
}).join("");
}


function matchPartida(a, b) {
const left = normalizeText(a);
const right = normalizeText(b);

if (!left || !right) {
return false;
}

return (
left.indexOf(right) >= 0 ||
right.indexOf(left) >= 0
);
}


/* =========================================================
GRÁFICAS
========================================================= */

function renderFinancialCharts() {
createChart(
"homeBudgetChart",
{
type: "bar",

data: {
labels: [
"Presupuesto",
"Real"
],

datasets: [
{
label: "US$",
data: [
financialData.budget,
financialData.real
]
}
]
},

options:
standardChartOptions(
formatUSD
)
}
);


let rows =
financialData.rows.filter(
function (row) {
return isMajorFinancialCategory(
row.name
);
}
);

if (!rows.length) {
rows =
financialData.rows.slice(
0,
10
);
}

const labels =
rows.map(function (row) {
return row.name;
});

const budgets =
rows.map(function (row) {
return row.budget;
});

const reals =
rows.map(function (row) {
return row.real;
});

[
"executiveCategoryChart",
"budgetVsRealChart",
"trackingFinancialChart"
].forEach(function (id) {
createChart(id, {
type: "bar",

data: {
labels: labels,

datasets: [
{
label:
"Presupuesto US$",
data: budgets
},
{
label:
"Real US$",
data: reals
}
]
},

options:
standardChartOptions(
formatUSD
)
});
});
}


function renderPurchaseCharts() {
const po =
purchaseData.filter(
function (row) {
return row.status === "po";
}
).length;

const pending =
purchaseData.filter(
function (row) {
return (
row.status === "pending"
);
}
).length;

const stock =
purchaseData.filter(
function (row) {
return (
row.status === "stock"
);
}
).length;

[
"homePurchaseChart",
"trackingStatusChart"
].forEach(function (id) {
createChart(id, {
type: "doughnut",

data: {
labels: [
"Con OC",
"Pendiente",
"Stock"
],

datasets: [
{
data: [
po,
pending,
stock
]
}
]
},

options: {
responsive: true,
maintainAspectRatio: false,

plugins: {
legend: {
position: "bottom"
}
}
}
});
});
}


function renderCashflowChart() {
const periods =
cashflowData.periods;

const parentRows =
cashflowData.rows.filter(
function (row) {
return row.isParent;
}
);

const source =
parentRows.length
? parentRows
: cashflowData.rows;

const values =
periods.map(
function (period, index) {
return source.reduce(
function (total, row) {
return (
total +
(row.values[index] || 0)
);
},
0
);
}
);

createChart(
"cashflowChart",
{
type: "bar",

data: {
labels:
periods.map(
function (period) {
return period.label;
}
),

datasets: [
{
label: "Flujo RD$",
data: values
}
]
},

options:
standardChartOptions(
formatDOP
)
}
);
}


function createChart(id, config) {
const canvas =
document.getElementById(id);

if (
!canvas ||
typeof Chart === "undefined"
) {
return;
}

if (charts[id]) {
charts[id].destroy();
}

charts[id] =
new Chart(canvas, config);
}


function standardChartOptions(formatter) {
return {
responsive: true,
maintainAspectRatio: false,

plugins: {
legend: {
position: "bottom"
},

tooltip: {
callbacks: {
label: function (context) {
let value = 0;

if (
context.parsed &&
typeof context.parsed.y ===
"number"
) {
value =
context.parsed.y;
} else if (
typeof context.parsed ===
"number"
) {
value =
context.parsed;
}

return (
(context.dataset.label ||
"") +
": " +
formatter(value)
);
}
}
}
},

scales: {
y: {
beginAtZero: true,

ticks: {
callback: function (value) {
return compactNumber(value);
}
}
}
}
};
}


/* =========================================================
COLUMNAS
========================================================= */

function findColumn(
headers,
candidates
) {
for (
let i = 0;
i < headers.length;
i++
) {
const header = headers[i];

for (
let c = 0;
c < candidates.length;
c++
) {
const candidate =
normalizeText(
candidates[c]
);

if (
header === candidate ||
header.indexOf(candidate) >= 0
) {
return i;
}
}
}

return -1;
}


function findColumnWithWords(
headers,
words
) {
const normalizedWords =
words.map(function (word) {
return normalizeText(word);
});

for (
let i = 0;
i < headers.length;
i++
) {
const header = headers[i];

let valid = true;

for (
let w = 0;
w < normalizedWords.length;
w++
) {
if (
header.indexOf(
normalizedWords[w]
) < 0
) {
valid = false;
break;
}
}

if (valid) {
return i;
}
}

return -1;
}


function getCell(row, index) {
if (
index === null ||
index === undefined ||
index < 0
) {
return null;
}

return row[index];
}


/* =========================================================
TEXTO
========================================================= */

function cleanText(value) {
if (
value === null ||
value === undefined
) {
return "";
}

return String(value)
.trim()
.split(/\s+/)
.join(" ");
}


function normalizeText(value) {
return cleanText(value)
.normalize("NFD")
.replace(
/[\u0300-\u036f]/g,
""
)
.toLowerCase()
.trim();
}


/* =========================================================
NÚMEROS
========================================================= */

function toNumber(value) {
if (
value === null ||
value === undefined ||
value === ""
) {
return 0;
}

if (typeof value === "number") {
return Number.isFinite(value)
? value
: 0;
}

let text =
String(value).trim();

if (!text) {
return 0;
}

let negative = false;

if (
text.startsWith("(") &&
text.endsWith(")")
) {
negative = true;

text =
text.substring(
1,
text.length - 1
);
}

text =
text
.replaceAll("RD$", "")
.replaceAll("USD$", "")
.replaceAll("US$", "")
.replaceAll("DOP", "")
.replaceAll("USD", "")
.replaceAll("$", "")
.replaceAll(",", "")
.replaceAll(" ", "")
.trim();

let cleaned = "";

for (
let i = 0;
i < text.length;
i++
) {
const character = text[i];

if (
(character >= "0" &&
character <= "9") ||
character === "." ||
character === "-"
) {
cleaned += character;
}
}

if (
!cleaned ||
cleaned === "-" ||
cleaned === "."
) {
return 0;
}

const number =
Number(cleaned);

if (!Number.isFinite(number)) {
return 0;
}

return negative
? -Math.abs(number)
: number;
}


/* =========================================================
FORMATOS
========================================================= */

function formatUSD(value) {
return (
"US$ " +
Number(value || 0)
.toLocaleString(
"en-US",
{
minimumFractionDigits: 2,
maximumFractionDigits: 2
}
)
);
}


function formatDOP(value) {
return (
"RD$ " +
Number(value || 0)
.toLocaleString(
"en-US",
{
minimumFractionDigits: 2,
maximumFractionDigits: 2
}
)
);
}


function formatDOPCompact(value) {
const number =
Number(value || 0);

if (number === 0) {
return "—";
}

return (
"RD$ " +
number.toLocaleString(
"en-US",
{
minimumFractionDigits: 0,
maximumFractionDigits: 0
}
)
);
}


function formatPercent(value) {
return (
Number(
(value || 0) * 100
).toLocaleString(
"en-US",
{
minimumFractionDigits: 1,
maximumFractionDigits: 1
}
) +
"%"
);
}


function formatNumber(value) {
return Number(value || 0)
.toLocaleString(
"en-US",
{
maximumFractionDigits: 2
}
);
}


function compactNumber(value) {
const number =
Number(value || 0);

if (
Math.abs(number) >=
1000000
) {
return (
(number / 1000000)
.toFixed(1) +
"M"
);
}

if (
Math.abs(number) >= 1000
) {
return (
(number / 1000)
.toFixed(0) +
"K"
);
}

return number.toFixed(0);
}


/* =========================================================
FECHAS
========================================================= */

function parseISODate(value) {
if (!value) {
return null;
}

const parts =
String(value).split("-");

if (parts.length !== 3) {
return null;
}

const year =
Number(parts[0]);

const month =
Number(parts[1]);

const day =
Number(parts[2]);

if (
!year ||
!month ||
!day
) {
return null;
}

const date =
new Date(
year,
month - 1,
day
);

if (
Number.isNaN(
date.getTime()
)
) {
return null;
}

return date;
}


function formatDate(date) {
if (
!(date instanceof Date) ||
Number.isNaN(date.getTime())
) {
return "—";
}

return date.toLocaleDateString(
"es-DO",
{
day: "2-digit",
month: "short",
year: "numeric"
}
);
}


function formatExcelDate(value) {
if (!value) {
return "";
}

if (value instanceof Date) {
return formatDate(value);
}

if (
typeof value === "number" &&
typeof XLSX !== "undefined"
) {
const parsed =
XLSX.SSF.parse_date_code(
value
);

if (parsed) {
return formatDate(
new Date(
parsed.y,
parsed.m - 1,
parsed.d
)
);
}
}

const date = new Date(value);

if (
!Number.isNaN(
date.getTime()
)
) {
return formatDate(date);
}

return cleanText(value);
}


/* =========================================================
HTML
========================================================= */

function escapeHTML(value) {
let text = cleanText(value);

text = text.replaceAll(
"&",
"&amp;"
);

text = text.replaceAll(
"<",
"&lt;"
);

text = text.replaceAll(
">",
"&gt;"
);

text = text.replaceAll(
'"',
"&quot;"
);

text = text.replaceAll(
"'",
"&#039;"
);

return text;
}


function emptyRow(columns) {
return (
"<tr>" +
'<td colspan="' +
columns +
'" style="text-align:center;padding:28px;color:#7a8581;">' +
"No hay datos para mostrar." +
"</td>" +
"</tr>"
);
}


/* =========================================================
MENSAJES
========================================================= */

function showMessage(
message,
type
) {
const element =
document.getElementById(
"dashboardMessage"
);

if (!element) {
return;
}

element.textContent = message;

element.className =
"dashboard-message " +
(type || "info");
}


function hideMessage() {
const element =
document.getElementById(
"dashboardMessage"
);

if (!element) {
return;
}

element.textContent = "";

element.className =
"dashboard-message hidden";
}


/* =========================================================
LOADING
========================================================= */

function showLoading(
visible,
message
) {
const overlay =
document.getElementById(
"loadingOverlay"
);

const text =
document.getElementById(
"loadingText"
);

if (text) {
text.textContent =
message ||
"Leyendo archivo Excel";
}

if (!overlay) {
return;
}

if (visible) {
overlay.classList.remove(
"hidden"
);
} else {
overlay.classList.add(
"hidden"
);
}
}


/* =========================================================
SET TEXT
========================================================= */

function setText(id, value) {
const element =
document.getElementById(id);

if (!element) {
return;
}

element.textContent =
value === null ||
value === undefined
? ""
: String(value);
}


/* =========================================================
RESET
========================================================= */

function resetDashboard() {
financialData = {
budget: 0,
real: 0,
difference: 0,
rows: []
};

purchaseData = [];

cashflowData = {
periods: [],
rows: []
};

[
"homeBudget",
"homeReal",
"homeDifference",
"executiveBudget",
"executiveReal",
"executiveDifference",
"budgetTotal",
"realTotal",
"differenceTotal"
].forEach(function (id) {
setText(id, "US$ 0.00");
});

[
"homeSavingsDOP",
"executiveSavingsDOP",
"budgetSavingsDOP",
"purchasePOValue",
"pendingValue",
"cashflowTotal"
].forEach(function (id) {
setText(id, "RD$ 0.00");
});

[
"homePOCount",
"homePendingCount",
"executivePO",
"executivePending",
"purchasePOCount",
"purchasePendingCount",
"purchaseStockCount",
"pendingTotal",
"pendingCategoryCount",
"cashflowCategoryCount",
"cashflowPeriodCount"
].forEach(function (id) {
setText(id, "0");
});

[
"homeExecution",
"executiveExecution",
"executiveProgressText",
"budgetExecution"
].forEach(function (id) {
setText(id, "0%");
});

const progress =
document.getElementById(
"executiveProgressBar"
);

if (progress) {
progress.style.width = "0%";
}
}

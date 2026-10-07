/* ============================================================
DASHBOARD CAPEX
SCRIPT.JS
============================================================ */

"use strict";

/* ============================================================
CONFIGURACIÓN
============================================================ */

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


/* ============================================================
INICIO
============================================================ */

document.addEventListener("DOMContentLoaded", async () => {

try {

setupNavigation();
setupEvents();

showLoading(true, "Cargando iniciativas...");

await loadProjects();

} catch (error) {

console.error("ERROR INICIAL:", error);

showMessage(
`No se pudo iniciar el dashboard: ${error.message}`,
"error"
);

} finally {

showLoading(false);

}

});


/* ============================================================
EVENTOS
============================================================ */

function setupEvents() {

const projectSelect =
document.getElementById("projectSelect");

if (projectSelect) {

projectSelect.addEventListener(
"change",
async function () {

if (!this.value) return;

await loadProject(this.value);

}
);

}


const refreshButton =
document.getElementById("refreshButton");

if (refreshButton) {

refreshButton.addEventListener(
"click",
async () => {

if (!currentProject) return;

await loadProject(
currentProject.id,
true
);

}
);

}


const budgetSearch =
document.getElementById("budgetSearch");

if (budgetSearch) {

budgetSearch.addEventListener(
"input",
renderBudgetTable
);

}


const purchaseSearch =
document.getElementById("purchaseSearch");

if (purchaseSearch) {

purchaseSearch.addEventListener(
"input",
renderPurchaseTable
);

}


const purchaseStatusFilter =
document.getElementById(
"purchaseStatusFilter"
);

if (purchaseStatusFilter) {

purchaseStatusFilter.addEventListener(
"change",
renderPurchaseTable
);

}


const pendingSearch =
document.getElementById("pendingSearch");

if (pendingSearch) {

pendingSearch.addEventListener(
"input",
renderPendingTable
);

}


const expandCashflow =
document.getElementById(
"expandCashflow"
);

if (expandCashflow) {

expandCashflow.addEventListener(
"click",
() => toggleAllCashflow(true)
);

}


const collapseCashflow =
document.getElementById(
"collapseCashflow"
);

if (collapseCashflow) {

collapseCashflow.addEventListener(
"click",
() => toggleAllCashflow(false)
);

}

}


/* ============================================================
NAVEGACIÓN
============================================================ */

function setupNavigation() {

const buttons =
document.querySelectorAll(
"#navTabs .nav-item"
);

buttons.forEach(button => {

button.addEventListener(
"click",
() => {

buttons.forEach(item =>
item.classList.remove("active")
);

button.classList.add("active");

const pageName =
button.dataset.page;

document
.querySelectorAll(".page")
.forEach(page =>
page.classList.remove("active")
);

const page =
document.getElementById(
`page-${pageName}`
);

if (page) {

page.classList.add("active");

}

}
);

});

}


/* ============================================================
PROJECTS.JSON
============================================================ */

async function loadProjects() {

try {

setText(
"projectLoadStatus",
"Cargando iniciativas..."
);

const response = await fetch(
`${PROJECTS_URL}?v=${Date.now()}`,
{
cache: "no-store"
}
);


if (!response.ok) {

throw new Error(
`No se pudo abrir projects.json (${response.status})`
);

}


const data = await response.json();


if (Array.isArray(data)) {

projects = data;

} else if (
data &&
Array.isArray(data.projects)
) {

projects = data.projects;

} else {

throw new Error(
"projects.json no contiene una lista válida."
);

}


populateProjectSelect();


if (projects.length === 0) {

setText(
"projectLoadStatus",
"Sin iniciativas"
);

showMessage(
"No hay iniciativas configuradas en projects.json.",
"warning"
);

return;

}


const firstProject = projects[0];

const select =
document.getElementById(
"projectSelect"
);

if (select) {

select.value =
String(firstProject.id);

}


await loadProject(
firstProject.id
);


} catch (error) {

console.error(
"ERROR CARGANDO PROJECTS.JSON:",
error
);

setText(
"projectLoadStatus",
"Error cargando iniciativas"
);

showMessage(
`No se pudieron cargar las iniciativas: ${error.message}`,
"error"
);

}

}


function populateProjectSelect() {

const select =
document.getElementById(
"projectSelect"
);

if (!select) return;


select.innerHTML = "";


projects.forEach(project => {

const option =
document.createElement("option");

option.value =
String(project.id);

option.textContent =
project.name ||
project.id ||
"Iniciativa";

select.appendChild(option);

});

}


/* ============================================================
CARGAR PROYECTO
============================================================ */

async function loadProject(
projectId,
forceRefresh = false
) {

const project =
projects.find(
item =>
String(item.id) ===
String(projectId)
);


if (!project) {

showMessage(
"La iniciativa seleccionada no existe en projects.json.",
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
`Leyendo ${project.name || project.id}...`
);


setText(
"projectLoadStatus",
"Leyendo Excel..."
);


if (!project.file) {

throw new Error(
"La iniciativa no tiene un archivo Excel definido."
);

}


const separator =
project.file.includes("?")
? "&"
: "?";


const fileUrl =
forceRefresh
? `${project.file}${separator}v=${Date.now()}`
: `${project.file}${separator}v=${Date.now()}`;


const response =
await fetch(
fileUrl,
{
cache: "no-store"
}
);


if (!response.ok) {

throw new Error(
`No se pudo abrir ${project.file} (${response.status})`
);

}


const buffer =
await response.arrayBuffer();


if (
typeof XLSX === "undefined"
) {

throw new Error(
"La librería XLSX no está disponible."
);

}


currentWorkbook =
XLSX.read(
buffer,
{
type: "array",
cellDates: true
}
);


console.log(
"HOJAS ENCONTRADAS:",
currentWorkbook.SheetNames
);


parseWorkbook(
currentWorkbook
);


renderDashboard();


setText(
"projectLoadStatus",
"Datos cargados"
);


} catch (error) {

console.error(
"ERROR CARGANDO PROYECTO:",
error
);


setText(
"projectLoadStatus",
"Error al leer Excel"
);


showMessage(
`No se pudo leer ${project.name || project.id}. ${error.message}`,
"error"
);


} finally {

showLoading(false);

}

}


/* ============================================================
CABECERA
============================================================ */

function updateProjectHeader(project) {

if (!project) return;


setText(
"topProjectName",
project.name ||
project.id ||
"Dashboard CAPEX"
);


setText(
"topProjectSubtitle",
"Seguimiento integral de presupuesto, compras y flujo de caja."
);


const opening =
parseISODate(
project.openingDate
);


if (!opening) {

setText(
"openingDate",
"—"
);

setText(
"daysToOpening",
"—"
);

setText(
"executiveOpeningDate",
"—"
);

setText(
"executiveDays",
"—"
);

return;

}


const formatted =
formatDate(opening);


setText(
"openingDate",
formatted
);

setText(
"executiveOpeningDate",
formatted
);


const today = new Date();

today.setHours(0, 0, 0, 0);

opening.setHours(0, 0, 0, 0);


const days =
Math.ceil(
(
opening.getTime() -
today.getTime()
) /
86400000
);


setText(
"daysToOpening",
days
);

setText(
"executiveDays",
days
);

}


/*
Alias por compatibilidad.
Si alguna parte vieja del código llama renderHeader,
seguirá funcionando.
*/

function renderHeader(project = currentProject) {

updateProjectHeader(project);

}


/* ============================================================
LEER WORKBOOK
============================================================ */

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


const financialSheet =
findSheet(
workbook,
"Presupuesto vs Real"
);


const purchaseSheet =
findSheet(
workbook,
"BD Plan Detallado"
);


const cashflowSheet =
findSheet(
workbook,
"Flujo de Caja"
);


const missing = [];


if (financialSheet) {

try {

financialData =
parseFinancialSheet(
financialSheet
);

} catch (error) {

console.error(
"ERROR PRESUPUESTO VS REAL:",
error
);

}

} else {

missing.push(
"Presupuesto vs Real"
);

}


if (purchaseSheet) {

try {

purchaseData =
parsePurchaseSheet(
purchaseSheet
);

} catch (error) {

console.error(
"ERROR BD PLAN DETALLADO:",
error
);

}

} else {

missing.push(
"BD Plan Detallado"
);

}


if (cashflowSheet) {

try {

cashflowData =
parseCashflowSheet(
cashflowSheet
);

} catch (error) {

console.error(
"ERROR FLUJO DE CAJA:",
error
);

}

} else {

missing.push(
"Flujo de Caja"
);

}


if (missing.length > 0) {

showMessage(
`El Excel no contiene: ${missing.join(", ")}.`,
"warning"
);

}

}


/* ============================================================
BUSCAR HOJA
============================================================ */

function findSheet(
workbook,
targetName
) {

const target =
normalizeText(targetName);


const sheetName =
workbook.SheetNames.find(
name =>
normalizeText(name) ===
target
);


if (!sheetName) {

return null;

}


return workbook.Sheets[
sheetName
];

}


/* ============================================================
MATRIZ DE EXCEL
============================================================ */

function sheetToMatrix(sheet) {

return XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: null,
raw: true
}
);

}


/* ============================================================
PRESUPUESTO VS REAL
============================================================ */

function parseFinancialSheet(sheet) {

const matrix =
sheetToMatrix(sheet);


const result = {
budget: 0,
real: 0,
difference: 0,
rows: []
};


if (!matrix.length) {

return result;

}


/* ----------------------------------------------------------
TOTAL GENERAL
---------------------------------------------------------- */

let totalRowIndex = -1;


for (
let r = 0;
r < matrix.length;
r++
) {

const text =
normalizeText(
matrix[r]
.filter(value => value != null)
.join(" ")
);


if (
text.includes(
"total general estimado"
)
) {

totalRowIndex = r;
break;

}

}


let totalInfo = null;


if (totalRowIndex >= 0) {

totalInfo =
findFinancialAmountsNearRow(
matrix,
totalRowIndex
);

}


if (totalInfo) {

result.budget =
totalInfo.budget;

result.real =
totalInfo.real;

result.difference =
totalInfo.difference;

}


/* ----------------------------------------------------------
COLUMNAS PRESUPUESTO / REAL / DIFERENCIA
---------------------------------------------------------- */

const columns =
detectFinancialColumns(
matrix,
totalRowIndex
);


/* ----------------------------------------------------------
RESUMEN DE PARTIDAS
---------------------------------------------------------- */

let summaryIndex = -1;


for (
let r = 0;
r < matrix.length;
r++
) {

const rowText =
normalizeText(
matrix[r]
.filter(value => value != null)
.join(" ")
);


if (
rowText.includes(
"resumen de partidas"
)
) {

summaryIndex = r;
break;

}

}


const start =
summaryIndex >= 0
? summaryIndex + 1
: 0;


for (
let r = start;
r < matrix.length;
r++
) {

const row = matrix[r];


if (!row) continue;


const label =
getFinancialRowLabel(row);


if (!label) continue;


const normalized =
normalizeText(label);


if (
normalized.includes(
"total general estimado"
) ||
normalized.includes(
"resumen de partidas"
)
) {

continue;

}


const amounts =
extractFinancialRowAmounts(
row,
columns
);


if (
amounts.budget === 0 &&
amounts.real === 0
) {

continue;

}


result.rows.push({
name: label,
budget: amounts.budget,
real: amounts.real,
difference:
amounts.difference
});

}


/*
Si el total general no pudo detectarse,
usamos únicamente las categorías principales,
evitando sumar todas las subpartidas.
*/

if (
result.budget === 0 &&
result.real === 0
) {

const major =
result.rows.filter(
row =>
isMajorFinancialCategory(
row.name
)
);


if (major.length) {

result.budget =
major.reduce(
(sum, row) =>
sum + row.budget,
0
);

result.real =
major.reduce(
(sum, row) =>
sum + row.real,
0
);

result.difference =
result.budget -
result.real;

}

}


if (
!Number.isFinite(
result.difference
)
) {

result.difference =
result.budget -
result.real;

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


/* ============================================================
DETECTAR TOTALES FINANCIEROS
============================================================ */

function findFinancialAmountsNearRow(
matrix,
index
) {

const candidateRows = [
index,
index + 1,
index - 1,
index + 2
];


for (
const rowIndex of candidateRows
) {

if (
rowIndex < 0 ||
rowIndex >= matrix.length
) {

continue;

}


const row =
matrix[rowIndex];


const numbers =
row
.map(
(value, column) => ({
value:
toNumber(value),
column
})
)
.filter(
item =>
Number.isFinite(
item.value
) &&
Math.abs(
item.value
) > 100
);


if (numbers.length >= 3) {

const lastThree =
numbers.slice(-3);


return {
budget:
lastThree[0].value,

real:
lastThree[1].value,

difference:
lastThree[2].value,

columns:
lastThree.map(
item => item.column
)
};

}


if (numbers.length === 2) {

return {
budget:
numbers[0].value,

real:
numbers[1].value,

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


/* ============================================================
COLUMNAS FINANCIERAS
============================================================ */

function detectFinancialColumns(
matrix,
totalRowIndex
) {

let budgetColumn = null;
let realColumn = null;
let differenceColumn = null;


const maxRows =
Math.min(
matrix.length,
Math.max(
20,
totalRowIndex + 5
)
);


for (
let r = 0;
r < maxRows;
r++
) {

const row =
matrix[r] || [];


row.forEach(
(value, column) => {

const text =
normalizeText(value);


if (
text === "presupuestado" ||
text === "presupuesto"
) {

budgetColumn = column;

}


if (
text === "real"
) {

realColumn = column;

}


if (
text === "diferencia"
) {

differenceColumn =
column;

}

}
);

}


if (
budgetColumn == null ||
realColumn == null
) {

const info =
totalRowIndex >= 0
? findFinancialAmountsNearRow(
matrix,
totalRowIndex
)
: null;


if (
info &&
info.columns
) {

budgetColumn =
info.columns[0];

realColumn =
info.columns[1];

differenceColumn =
info.columns[2];

}

}


return {
budgetColumn,
realColumn,
differenceColumn
};

}


/* ============================================================
LABEL FINANCIERO
============================================================ */

function getFinancialRowLabel(row) {

for (
let c = 0;
c < Math.min(
row.length,
6
);
c++
) {

const value = row[c];


if (
typeof value === "string" &&
value.trim()
) {

const text =
value.trim();


if (
normalizeText(text) !==
"usd"
) {

return text;

}

}

}


return "";

}


/* ============================================================
VALORES FINANCIEROS POR FILA
============================================================ */

function extractFinancialRowAmounts(
row,
columns
) {

let budget = 0;
let real = 0;
let difference = 0;


if (
columns.budgetColumn != null
) {

budget =
toNumber(
row[
columns.budgetColumn
]
);

}


if (
columns.realColumn != null
) {

real =
toNumber(
row[
columns.realColumn
]
);

}


if (
columns.differenceColumn != null
) {

difference =
toNumber(
row[
columns.differenceColumn
]
);

}


/*
Algunos archivos tienen una columna
separada con "USD$".
Si las columnas detectadas no tienen números,
buscamos los últimos montos de la fila.
*/

if (
budget === 0 &&
real === 0
) {

const numbers =
row
.map(toNumber)
.filter(
value =>
Number.isFinite(value) &&
Math.abs(value) > 0
);


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
budget,
real,
difference
};

}


/* ============================================================
CATEGORÍA FINANCIERA PRINCIPAL
============================================================ */

function isMajorFinancialCategory(
name
) {

const text =
normalizeText(name);


return [
"edificacion",
"equipamiento",
"terreno"
].includes(text);

}


/* ============================================================
BD PLAN DETALLADO
============================================================ */

function parsePurchaseSheet(sheet) {

const matrix =
sheetToMatrix(sheet);


if (!matrix.length) {

return [];

}


const headerIndex =
findPurchaseHeaderRow(
matrix
);


if (headerIndex < 0) {

console.warn(
"No se encontró encabezado de BD Plan Detallado."
);

return [];

}


const headers =
matrix[headerIndex]
.map(normalizeText);


const columns = {

partida:
findColumn(
headers,
[
"partidas",
"partida"
]
),

item:
findColumn(
headers,
[
"item"
]
),

order:
findColumn(
headers,
[
"orden de compra"
]
),

date:
findColumn(
headers,
[
"fecha"
]
),

quantity:
findColumn(
headers,
[
"cant",
"cantidad"
]
),

netDOP:
findColumnIncludes(
headers,
[
"valor neto",
"dop"
]
),

budgetDOP:
findColumnIncludes(
headers,
[
"ppto",
"orden interna",
"dop"
]
),

comment:
findColumn(
headers,
[
"comentario",
"observacion",
"observaciones"
]
)

};


const result = [];


for (
let r = headerIndex + 1;
r < matrix.length;
r++
) {

const row =
matrix[r];


if (!row) continue;


const orderValue =
getCell(
row,
columns.order
);


const status =
classifyPurchaseStatus(
orderValue
);


/*
PLAN y valores desconocidos
no entran al general.
*/

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


if (
!partida &&
!item &&
!orderValue
) {

continue;

}


result.push({

partida,

item,

order:
cleanText(orderValue),

date:
getCell(
row,
columns.date
),

quantity:
toNumber(
getCell(
row,
columns.quantity
)
),

netDOP:
toNumber(
getCell(
row,
columns.netDOP
)
),

budgetDOP:
toNumber(
getCell(
row,
columns.budgetDOP
)
),

comment:
cleanText(
getCell(
row,
columns.comment
)
),

status

});

}


return result;

}


/* ============================================================
ENCONTRAR ENCABEZADO BD
============================================================ */

function findPurchaseHeaderRow(
matrix
) {

const limit =
Math.min(
matrix.length,
30
);


for (
let r = 0;
r < limit;
r++
) {

const values =
(matrix[r] || [])
.map(normalizeText);


const hasOrder =
values.some(
value =>
value.includes(
"orden de compra"
)
);


const hasItem =
values.some(
value =>
value === "item" ||
value.includes("item")
);


const hasPartida =
values.some(
value =>
value.includes(
"partida"
)
);


if (
hasOrder &&
(
hasItem ||
hasPartida
)
) {

return r;

}

}


return -1;

}


/* ============================================================
ESTADO DE COMPRA
============================================================ */

function classifyPurchaseStatus(
value
) {

const raw =
cleanText(value);


const text =
normalizeText(raw);


if (!text) {

return "other";

}


if (
text.includes(
"pendiente de compra"
) ||
text === "pendiente"
) {

return "pending";

}


if (
text === "stock" ||
text.includes("stock")
) {

return "stock";

}


if (
text === "plan" ||
text.includes("plan")
) {

return "ignore";

}


/*
Orden de compra real:
se aceptan números largos aunque Excel
los haya leído como número.
*/

const digits =
raw.replace(/\D/g, "");


if (digits.length >= 6) {

return "po";

}


return "other";

}


/* ============================================================
FLUJO DE CAJA
============================================================ */

function parseCashflowSheet(sheet) {

const matrix =
sheetToMatrix(sheet);


const result = {
periods: [],
rows: []
};


if (!matrix.length) {

return result;

}


const periodInfo =
detectCashflowPeriods(
matrix
);


if (!periodInfo) {

console.warn(
"No se detectaron períodos en Flujo de Caja."
);

return result;

}


result.periods =
periodInfo.periods;


const firstPeriodColumn =
periodInfo.periods[0].column;


let lastParent = null;


for (
let r = periodInfo.rowIndex + 1;
r < matrix.length;
r++
) {

const row =
matrix[r] || [];


const label =
getCashflowLabel(
row,
firstPeriodColumn
);


if (!label) continue;


const values =
periodInfo.periods.map(
period =>
toNumber(
row[period.column]
)
);


const total =
values.reduce(
(sum, value) =>
sum + value,
0
);


if (total === 0) {

continue;

}


const isParent =
isCashflowParent(label);


if (isParent) {

lastParent = label;

}


result.rows.push({

id:
`cf-${r}`,

label,

values,

total,

isParent,

parent:
isParent
? null
: lastParent

});

}


return result;

}


/* ============================================================
PERÍODOS FLUJO
============================================================ */

function detectCashflowPeriods(
matrix
) {

const limit =
Math.min(
matrix.length,
30
);


for (
let r = 0;
r < limit;
r++
) {

const row =
matrix[r] || [];


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


if (
candidates.length >= 3
) {

const sequential =
candidates.filter(
(item, index) => {

if (index === 0) {

return true;

}

return (
item.period ===
candidates[
index - 1
].period + 1
);

}
);


if (
sequential.length >= 3
) {

return {

rowIndex: r,

periods:
candidates.map(
item => ({
...item,
label:
periodToMonth(
item.period
)
})
)

};

}

}

}


return null;

}


/* ============================================================
LABEL FLUJO
============================================================ */

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


if (
value &&
!/^usd\\$?$/i.test(value) &&
!/^rd\\$?$/i.test(value)
) {

parts.push(value);

}

}


return parts.join(" ").trim();

}


/* ============================================================
PADRE / SUBPARTIDA FLUJO
============================================================ */

function isCashflowParent(label) {

const text =
cleanText(label);


if (
/^\\d+\\.00\\b/.test(text)
) {

return true;

}


if (
/^[A-Z]\\d+\\b/i.test(text)
) {

return true;

}


return false;

}


/* ============================================================
PERÍODO A MES
1 = ENERO 2025
============================================================ */

function periodToMonth(period) {

const baseYear = 2025;

const monthIndex =
period - 1;


const year =
baseYear +
Math.floor(
monthIndex / 12
);


const month =
monthIndex % 12;


const names = [
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


return `${names[month]} ${year}`;

}


/* ============================================================
RENDER GENERAL
============================================================ */

function renderDashboard() {

renderFinancial();
renderPurchases();
renderCashflow();
renderTracking();

}


/* ============================================================
RENDER FINANCIERO
============================================================ */

function renderFinancial() {

const budget =
financialData.budget || 0;

const real =
financialData.real || 0;

const difference =
financialData.difference ||
(
budget - real
);


const execution =
budget !== 0
? real / budget
: 0;


const savingsDOP =
difference *
SAVINGS_RATE;


/* HOME */

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


setSavingsValue(
"homeSavingsDOP",
savingsDOP
);


/* EXECUTIVE */

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


const progress =
document.getElementById(
"executiveProgressBar"
);


if (progress) {

progress.style.width =
`${Math.max(
0,
Math.min(
execution * 100,
100
)
)}%`;

}


setSavingsValue(
"executiveSavingsDOP",
savingsDOP
);


/* BUDGET PAGE */

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

setSavingsValue(
"budgetSavingsDOP",
savingsDOP
);


renderBudgetTable();
renderHomeCategoryTable();

renderFinancialCharts();

}


/* ============================================================
AHORRO / SOBREEJECUCIÓN
============================================================ */

function setSavingsValue(
elementId,
value
) {

const element =
document.getElementById(
elementId
);


if (!element) return;


element.textContent =
formatDOP(
Math.abs(value)
);


const card =
element.closest(
".kpi-card"
);


if (!card) return;


const label =
card.querySelector(
".kpi-label"
);


if (label) {

label.textContent =
value >= 0
? "Ahorro RD$"
: "Sobreejecución RD$";

}

}


/* ============================================================
TABLA PRESUPUESTO
============================================================ */

function renderBudgetTable() {

const tbody =
document.querySelector(
"#budgetTable tbody"
);


if (!tbody) return;


const search =
normalizeText(
document.getElementById(
"budgetSearch"
)?.value || ""
);


const rows =
financialData.rows.filter(
row =>
!search ||
normalizeText(
row.name
).includes(search)
);


tbody.innerHTML =
rows.length
? rows.map(row => {

const execution =
row.budget !== 0
? row.real /
row.budget
: 0;


return `
<tr>
<td>${escapeHTML(row.name)}</td>
<td class="number">${formatUSD(row.budget)}</td>
<td class="number">${formatUSD(row.real)}</td>
<td class="number">${formatUSD(row.difference)}</td>
<td class="number">${formatPercent(execution)}</td>
</tr>
`;

}).join("")
: emptyRow(5);

}


/* ============================================================
TABLA HOME
============================================================ */

function renderHomeCategoryTable() {

const tbody =
document.querySelector(
"#homeCategoryTable tbody"
);


if (!tbody) return;


let rows =
financialData.rows.filter(
row =>
isMajorFinancialCategory(
row.name
)
);


if (!rows.length) {

rows =
financialData.rows.slice(
0,
10
);

}


tbody.innerHTML =
rows.length
? rows.map(row => {

const execution =
row.budget !== 0
? row.real /
row.budget
: 0;


return `
<tr>
<td>${escapeHTML(row.name)}</td>
<td class="number">${formatUSD(row.budget)}</td>
<td class="number">${formatUSD(row.real)}</td>
<td class="number">${formatUSD(row.difference)}</td>
<td class="number">${formatPercent(execution)}</td>
</tr>
`;

}).join("")
: emptyRow(5);

}


/* ============================================================
RENDER COMPRAS
============================================================ */

function renderPurchases() {

const po =
purchaseData.filter(
row =>
row.status === "po"
);


const pending =
purchaseData.filter(
row =>
row.status ===
"pending"
);


const stock =
purchaseData.filter(
row =>
row.status ===
"stock"
);


const poValue =
po.reduce(
(sum, row) =>
sum + row.netDOP,
0
);


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


const affected =
new Set(
pending
.map(
row =>
normalizeText(
row.partida
)
)
.filter(Boolean)
);


setText(
"pendingCategoryCount",
affected.size
);


const pendingValue =
pending.reduce(
(sum, row) =>
sum +
row.budgetDOP,
0
);


setText(
"pendingValue",
formatDOP(
pendingValue
)
);


renderPurchaseTable();
renderPendingTable();
renderPurchaseCharts();

}


/* ============================================================
TABLA COMPRAS
============================================================ */

function renderPurchaseTable() {

const tbody =
document.querySelector(
"#purchaseTable tbody"
);


if (!tbody) return;


const search =
normalizeText(
document.getElementById(
"purchaseSearch"
)?.value || ""
);


const status =
document.getElementById(
"purchaseStatusFilter"
)?.value || "all";


const rows =
purchaseData.filter(row => {

if (
status !== "all" &&
row.status !== status
) {

return false;

}


if (!search) {

return true;

}


const haystack =
normalizeText(
[
row.partida,
row.item,
row.order
].join(" ")
);


return haystack.includes(
search
);

});


tbody.innerHTML =
rows.length
? rows.map(row => `
<tr>
<td>${escapeHTML(row.partida)}</td>
<td>${escapeHTML(row.item)}</td>
<td>${escapeHTML(row.order)}</td>
<td>${formatExcelDate(row.date)}</td>
<td class="number">${formatNumber(row.quantity)}</td>
<td class="number">${formatDOP(row.netDOP)}</td>
<td>${purchaseStatusBadge(row.status)}</td>
</tr>
`).join("")
: emptyRow(7);

}


/* ============================================================
TABLA PENDIENTES
============================================================ */

function renderPendingTable() {

const tbody =
document.querySelector(
"#pendingTable tbody"
);


if (!tbody) return;


const search =
normalizeText(
document.getElementById(
"pendingSearch"
)?.value || ""
);


const rows =
purchaseData.filter(
row => {

if (
row.status !==
"pending"
) {

return false;

}


if (!search) {

return true;

}


return normalizeText(
`${row.partida} ${row.item}`
).includes(search);

}
);


tbody.innerHTML =
rows.length
? rows.map(row => `
<tr>
<td>${escapeHTML(row.partida)}</td>
<td>${escapeHTML(row.item)}</td>
<td class="number">${formatNumber(row.quantity)}</td>
<td class="number">${formatDOP(row.budgetDOP)}</td>
<td>${escapeHTML(row.comment)}</td>
</tr>
`).join("")
: emptyRow(5);

}


/* ============================================================
BADGE ESTADO
============================================================ */

function purchaseStatusBadge(
status
) {

if (status === "po") {

return `
<span class="status-badge status-po">
Con OC
</span>
`;

}


if (status === "pending") {

return `
<span class="status-badge status-pending">
Pendiente de compra
</span>
`;

}


if (status === "stock") {

return `
<span class="status-badge status-stock">
Stock
</span>
`;

}


return "";

}


/* ============================================================
FLUJO DE CAJA
============================================================ */

function renderCashflow() {

const periods =
cashflowData.periods;

const rows =
cashflowData.rows;


const total =
calculateCashflowTotal(
rows
);


setText(
"cashflowTotal",
formatDOP(total)
);


setText(
"cashflowPeriodCount",
periods.length
);


const parentCount =
rows.filter(
row =>
row.isParent
).length;


setText(
"cashflowCategoryCount",
parentCount ||
rows.length
);


renderCashflowTable();
renderCashflowChart();

}


/* ============================================================
TOTAL FLUJO SIN DOBLE CONTEO
============================================================ */

function calculateCashflowTotal(
rows
) {

const parents =
rows.filter(
row =>
row.isParent
);


if (parents.length) {

return parents.reduce(
(sum, row) =>
sum + row.total,
0
);

}


return rows.reduce(
(sum, row) =>
sum + row.total,
0
);

}


/* ============================================================
TABLA FLUJO
============================================================ */

function renderCashflowTable() {

const thead =
document.getElementById(
"cashflowTableHead"
);


const tbody =
document.getElementById(
"cashflowTableBody"
);


if (
!thead ||
!tbody
) {

return;

}


const periods =
cashflowData.periods;


thead.innerHTML = `
<tr>
<th>Partida</th>

${periods
.map(
period =>
`<th class="number">${escapeHTML(period.label)}</th>`
)
.join("")}

<th class="number">Total</th>
</tr>
`;


if (
!cashflowData.rows.length
) {

tbody.innerHTML =
emptyRow(
periods.length + 2
);

return;

}


const html = [];


cashflowData.rows.forEach(
row => {

if (row.isParent) {

html.push(`
<tr
class="cashflow-group-row"
data-parent="${escapeHTML(row.id)}"
>

<td>

<button
class="cashflow-toggle"
type="button"
data-target="${escapeHTML(row.label)}"
>

<span class="toggle-symbol">
−
</span>

<span>
${escapeHTML(row.label)}
</span>

</button>

</td>

${row.values
.map(
value =>
`<td class="number">${formatDOPCompact(value)}</td>`
)
.join("")}

<td class="number">
${formatDOPCompact(row.total)}
</td>

</tr>
`);

} else {

html.push(`
<tr
class="cashflow-child-row"
data-parent-label="${escapeHTML(row.parent || "")}"
>

<td class="cashflow-child-label">
${escapeHTML(row.label)}
</td>

${row.values
.map(
value =>
`<td class="number">${formatDOPCompact(value)}</td>`
)
.join("")}

<td class="number">
${formatDOPCompact(row.total)}
</td>

</tr>
`);

}

}
);


tbody.innerHTML =
html.join("");


tbody
.querySelectorAll(
".cashflow-toggle"
)
.forEach(button => {

button.addEventListener(
"click",
() => {

const parent =
button.dataset.target;

const children =
tbody.querySelectorAll(
".cashflow-child-row"
);


let shouldHide = false;


for (
const child of children
) {

if (
child.dataset.parentLabel ===
parent
) {

shouldHide =
child.style.display !==
"none";

break;

}

}


children.forEach(
child => {

if (
child.dataset.parentLabel ===
parent
) {

child.style.display =
shouldHide
? "none"
: "";

}

}
);


const symbol =
button.querySelector(
".toggle-symbol"
);


if (symbol) {

symbol.textContent =
shouldHide
? "+"
: "−";

}

}
);

});

}


/* ============================================================
EXPANDIR / CONTRAER FLUJO
============================================================ */

function toggleAllCashflow(
expand
) {

document
.querySelectorAll(
".cashflow-child-row"
)
.forEach(row => {

row.style.display =
expand
? ""
: "none";

});


document
.querySelectorAll(
".cashflow-toggle .toggle-symbol"
)
.forEach(symbol => {

symbol.textContent =
expand
? "−"
: "+";

});

}


/* ============================================================
SEGUIMIENTO
============================================================ */

function renderTracking() {

const tbody =
document.querySelector(
"#trackingTable tbody"
);


if (!tbody) return;


let rows =
financialData.rows.filter(
row =>
isMajorFinancialCategory(
row.name
)
);


if (!rows.length) {

rows =
financialData.rows.slice(
0,
15
);

}


tbody.innerHTML =
rows.length
? rows.map(row => {

const purchases =
purchaseData.filter(
item =>
matchPartida(
item.partida,
row.name
)
);


const po =
purchases.filter(
item =>
item.status ===
"po"
).length;


const pending =
purchases.filter(
item =>
item.status ===
"pending"
).length;


const stock =
purchases.filter(
item =>
item.status ===
"stock"
).length;


const execution =
row.budget !== 0
? row.real /
row.budget
: 0;


return `
<tr>
<td>${escapeHTML(row.name)}</td>
<td class="number">${formatUSD(row.budget)}</td>
<td class="number">${formatUSD(row.real)}</td>
<td class="number">${formatPercent(execution)}</td>
<td class="number">${po}</td>
<td class="number">${pending}</td>
<td class="number">${stock}</td>
</tr>
`;

}).join("")
: emptyRow(7);

}


/* ============================================================
MATCH PARTIDAS
============================================================ */

function matchPartida(
a,
b
) {

const left =
normalizeText(a);

const right =
normalizeText(b);


if (
!left ||
!right
) {

return false;

}


return (
left.includes(right) ||
right.includes(left)
);

}


/* ============================================================
CHARTS FINANCIEROS
============================================================ */

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

datasets: [{
label: "US$",
data: [
financialData.budget,
financialData.real
]
}]

},

options:
standardChartOptions(
value =>
formatUSD(value)
)
}
);


let categories =
financialData.rows.filter(
row =>
isMajorFinancialCategory(
row.name
)
);


if (!categories.length) {

categories =
financialData.rows.slice(
0,
10
);

}


const categoryConfig = {

type: "bar",

data: {

labels:
categories.map(
row => row.name
),

datasets: [

{
label:
"Presupuesto US$",

data:
categories.map(
row =>
row.budget
)
},

{
label:
"Real US$",

data:
categories.map(
row =>
row.real
)
}

]

},

options:
standardChartOptions(
value =>
formatUSD(value)
)

};


createChart(
"executiveCategoryChart",
categoryConfig
);


createChart(
"budgetVsRealChart",
categoryConfig
);


createChart(
"trackingFinancialChart",
categoryConfig
);

}


/* ============================================================
CHARTS COMPRAS
============================================================ */

function renderPurchaseCharts() {

const po =
purchaseData.filter(
row =>
row.status === "po"
).length;


const pending =
purchaseData.filter(
row =>
row.status ===
"pending"
).length;


const stock =
purchaseData.filter(
row =>
row.status ===
"stock"
).length;


const config = {

type: "doughnut",

data: {

labels: [
"Con OC",
"Pendiente",
"Stock"
],

datasets: [{
data: [
po,
pending,
stock
]
}]

},

options: {

responsive: true,

maintainAspectRatio:
false,

plugins: {

legend: {
position: "bottom"
}

}

}

};


createChart(
"homePurchaseChart",
config
);


createChart(
"trackingStatusChart",
config
);

}


/* ============================================================
CHART FLUJO
============================================================ */

function renderCashflowChart() {

const periods =
cashflowData.periods;


const values =
periods.map(
(_, index) => {

const parents =
cashflowData.rows.filter(
row =>
row.isParent
);


const source =
parents.length
? parents
: cashflowData.rows;


return source.reduce(
(sum, row) =>
sum +
(
row.values[index] ||
0
),
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
period =>
period.label
),

datasets: [{
label: "Flujo RD$",
data: values
}]

},

options:
standardChartOptions(
value =>
formatDOP(value)
)

}
);

}


/* ============================================================
CREAR CHART
============================================================ */

function createChart(
canvasId,
config
) {

const canvas =
document.getElementById(
canvasId
);


if (
!canvas ||
typeof Chart ===
"undefined"
) {

return;

}


if (
charts[canvasId]
) {

charts[canvasId].destroy();

}


charts[canvasId] =
new Chart(
canvas,
config
);

}


/* ============================================================
OPCIONES CHART
============================================================ */

function standardChartOptions(
formatter
) {

return {

responsive: true,

maintainAspectRatio:
false,

plugins: {

legend: {
position: "bottom"
},

tooltip: {

callbacks: {

label(context) {

const value =
context.parsed.y ??
context.parsed ??
0;


return (
`${context.dataset.label || ""}: ` +
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

callback(value) {

return compactNumber(
value
);

}

}

}

}

};

}


/* ============================================================
RESET
============================================================ */

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


const usdIds = [
"homeBudget",
"homeReal",
"homeDifference",
"executiveBudget",
"executiveReal",
"executiveDifference",
"budgetTotal",
"realTotal",
"differenceTotal"
];


usdIds.forEach(
id =>
setText(
id,
"US$ 0.00"
)
);


const dopIds = [
"homeSavingsDOP",
"executiveSavingsDOP",
"budgetSavingsDOP",
"purchasePOValue",
"pendingValue",
"cashflowTotal"
];


dopIds.forEach(
id =>
setText(
id,
"RD$ 0.00"
)
);


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
].forEach(
id =>
setText(id, "0")
);


[
"homeExecution",
"executiveExecution",
"executiveProgressText",
"budgetExecution"
].forEach(
id =>
setText(id, "0%")
);


const progress =
document.getElementById(
"executiveProgressBar"
);


if (progress) {

progress.style.width =
"0%";

}

}


/* ============================================================
HELPERS COLUMNAS
============================================================ */

function findColumn(
headers,
candidates
) {

for (
let index = 0;
index < headers.length;
index++
) {

const header =
headers[index];


if (
candidates.some(
candidate =>
header ===
normalizeText(
candidate
) ||
header.includes(
normalizeText(
candidate
)
)
)
) {

return index;

}

}


return -1;

}


function findColumnIncludes(
headers,
requiredWords
) {

const words =
requiredWords.map(
normalizeText
);


for (
let index = 0;
index < headers.length;
index++
) {

const header =
headers[index];


if (
words.every(
word =>
header.includes(word)
)
) {

return index;

}

}


return -1;

}


function getCell(
row,
index
) {

if (
index == null ||
index < 0
) {

return null;

}


return row[index];

}


/* ============================================================
NÚMEROS
============================================================ */

function toNumber(value) {

if (
value == null ||
value === ""
) {

return 0;

}


if (
typeof value ===
"number"
) {

return Number.isFinite(value)
? value
: 0;

}


if (
typeof value ===
"boolean"
) {

return value
? 1
: 0;

}


let text =
String(value)
.trim();


if (!text) {

return 0;

}


const negative =
/^\\(.*\\)$/.test(text);


text =
text
.replace(/[()]/g, "")
.replace(/RD\\$/gi, "")
.replace(/USD\\$/gi, "")
.replace(/US\\$/gi, "")
.replace(/DOP/gi, "")
.replace(/USD/gi, "")
.replace(/\\$/g, "")
.replace(/,/g, "")
.replace(/\\s/g, "")
.replace(/[^0-9.\\-]/g, "");


if (
!text ||
text === "-" ||
text === "."
) {

return 0;

}


const number =
Number(text);


if (
!Number.isFinite(number)
) {

return 0;

}


return negative
? -Math.abs(number)
: number;

}


/* ============================================================
TEXTO
============================================================ */

function cleanText(value) {

if (
value == null
) {

return "";

}


return String(value)
.trim()
.replace(/\\s+/g, " ");

}


function normalizeText(value) {

return cleanText(value)
.normalize("NFD")
.replace(
/[\\u0300-\\u036f]/g,
""
)
.toLowerCase()
.replace(/\\s+/g, " ")
.trim();

}


/* ============================================================
FORMATOS
============================================================ */

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
(
number /
1000000
).toFixed(1) +
"M"
);

}


if (
Math.abs(number) >=
1000
) {

return (
(
number /
1000
).toFixed(0) +
"K"
);

}


return number.toFixed(0);

}


/* ============================================================
FECHAS
============================================================ */

function parseISODate(value) {

if (!value) {

return null;

}


const match =
String(value).match(
/^(\\d{4})-(\\d{2})-(\\d{2})$/
);


if (!match) {

return null;

}


const date =
new Date(
Number(match[1]),
Number(match[2]) - 1,
Number(match[3])
);


return Number.isNaN(
date.getTime()
)
? null
: date;

}


function formatDate(date) {

if (
!(date instanceof Date) ||
Number.isNaN(
date.getTime()
)
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


if (
value instanceof Date
) {

return formatDate(value);

}


if (
typeof value ===
"number"
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


const date =
new Date(value);


if (
!Number.isNaN(
date.getTime()
)
) {

return formatDate(date);

}


return cleanText(value);

}


/* ============================================================
HTML
============================================================ */

function escapeHTML(value) {

return cleanText(value)
.replace(
/&/g,
"&amp;"
)
.replace(
/</g,
"&lt;"
)
.replace(
/>/g,
"&gt;"
)
.replace(
/"/g,
"&quot;"
)
.replace(
/'/g,
"&#039;"
);

}


function emptyRow(columns) {

return `
<tr>
<td
colspan="${columns}"
style="
text-align:center;
padding:28px;
color:#7a8581;
"
>
No hay datos para mostrar.
</td>
</tr>
`;

}


/* ============================================================
MENSAJES
============================================================ */

function showMessage(
message,
type = "info"
) {

const element =
document.getElementById(
"dashboardMessage"
);


if (!element) return;


element.textContent =
message;


element.className =
`dashboard-message ${type}`;

}


function hideMessage() {

const element =
document.getElementById(
"dashboardMessage"
);


if (!element) return;


element.className =
"dashboard-message hidden";

element.textContent =
"";

}


/* ============================================================
LOADING
============================================================ */

function showLoading(
visible,
text = "Leyendo archivo Excel"
) {

const overlay =
document.getElementById(
"loadingOverlay"
);


const loadingText =
document.getElementById(
"loadingText"
);


if (loadingText) {

loadingText.textContent =
text;

}


if (!overlay) return;


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


/* ============================================================
SET TEXT
============================================================ */

function setText(
id,
value
) {

const element =
document.getElementById(id);


if (!element) {

return;

}


element.textContent =
value == null
? ""
: String(value);

}

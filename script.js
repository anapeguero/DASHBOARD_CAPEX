/* ============================================================
DASHBOARD CAPEX - GRUPO RAMOS
script.js

FUENTES:
1. Presupuesto vs Real
2. BD Plan Detallado
3. Flujo de Caja

REGLAS:
- Presupuesto vs Real se muestra en USD.
- Ahorro RD$ = Diferencia USD x 60.
- Flujo de Caja se muestra en RD$.
- "Pendiente de compra" = pendiente.
- OC real = NO pendiente.
- "Stock" = NO pendiente.
- "Plan" = se excluye de indicadores generales.
- No se utiliza información de proveedores.
============================================================ */


/* ============================================================
CONFIGURACIÓN
============================================================ */

const EXCHANGE_RATE = 60;

const CONFIG = {
projectsFile: "data/projects.json",
budgetSheet: "Presupuesto vs Real",
detailSheet: "BD Plan Detallado",
cashflowSheet: "Flujo de Caja"
};


/* ============================================================
ESTADO GLOBAL
============================================================ */

let projects = [];

let currentProject = null;

let dashboardData = null;

let charts = {};


/* ============================================================
INICIO
============================================================ */

document.addEventListener("DOMContentLoaded", async () => {

setupNavigation();

setupEvents();

await loadProjects();

});


/* ============================================================
NAVEGACIÓN
============================================================ */

function setupNavigation() {

const buttons = document.querySelectorAll(".nav-item");
const pages = document.querySelectorAll(".page");

buttons.forEach(button => {

button.addEventListener("click", () => {

const pageName = button.dataset.page;

buttons.forEach(btn => {
btn.classList.remove("active");
});

pages.forEach(page => {
page.classList.remove("active");
});

button.classList.add("active");

const targetPage =
document.getElementById(`page-${pageName}`);

if (targetPage) {
targetPage.classList.add("active");
}

resizeCharts();

});

});

}


/* ============================================================
EVENTOS
============================================================ */

function setupEvents() {

const projectSelect =
document.getElementById("projectSelect");

if (projectSelect) {

projectSelect.addEventListener(
"change",
async event => {

const projectId = event.target.value;

if (!projectId) return;

await loadProject(projectId);

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


const budgetSearch =
document.getElementById("budgetSearch");

if (budgetSearch) {

budgetSearch.addEventListener(
"input",
renderBudgetTable
);

}


const expandCashflow =
document.getElementById("expandCashflow");

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
CARGAR PROJECTS.JSON
============================================================ */

async function loadProjects() {

try {

showLoading(
true,
"Cargando iniciativas..."
);

const response = await fetch(
`${CONFIG.projectsFile}?v=${Date.now()}`
);

if (!response.ok) {

throw new Error(
`No se pudo cargar projects.json (${response.status})`
);

}

const json = await response.json();

projects = Array.isArray(json)
? json
: json.projects || [];

populateProjectSelector();

if (projects.length > 0) {

const firstProject = projects[0];

document.getElementById(
"projectSelect"
).value = firstProject.id;

await loadProject(firstProject.id);

} else {

showMessage(
"No hay iniciativas configuradas en projects.json.",
"warning"
);

}

} catch (error) {

console.error(error);

showMessage(
"No fue posible cargar data/projects.json.",
"error"
);

} finally {

showLoading(false);

}

}


/* ============================================================
SELECTOR DE PROYECTOS
============================================================ */

function populateProjectSelector() {

const select =
document.getElementById("projectSelect");

if (!select) return;

select.innerHTML = "";

projects.forEach(project => {

const option =
document.createElement("option");

option.value = project.id;

option.textContent =
project.name || project.id;

select.appendChild(option);

});

}


/* ============================================================
CARGAR EXCEL
============================================================ */

async function loadProject(
projectId,
forceRefresh = false
) {

const project = projects.find(
item => item.id === projectId
);

if (!project) return;


currentProject = project;


try {

showLoading(
true,
`Leyendo ${project.name}...`
);

hideMessage();

updateProjectHeader(project);


const separator =
project.file.includes("?")
? "&"
: "?";


const fileUrl =
forceRefresh
? `${project.file}${separator}v=${Date.now()}`
: `${project.file}${separator}v=${Date.now()}`;


console.log(
"Cargando Excel:",
fileUrl
);


const response =
await fetch(fileUrl);


if (!response.ok) {

throw new Error(
`No se pudo abrir ${project.file}. HTTP ${response.status}`
);

}


const arrayBuffer =
await response.arrayBuffer();


const workbook =
XLSX.read(
arrayBuffer,
{
type: "array",
cellDates: true,
cellFormula: true
}
);


console.log(
"Hojas encontradas:",
workbook.SheetNames
);


dashboardData =
parseWorkbook(workbook);


renderDashboard();


updateLoadStatus(
"Datos actualizados"
);


} catch (error) {

console.error(
"ERROR CARGANDO PROYECTO:",
error
);


updateLoadStatus(
"Error al cargar"
);


showMessage(
`No se pudo leer ${project.name}. ${error.message}`,
"error"
);


resetDashboard();

} finally {

showLoading(false);

}

}


/* ============================================================
PARSEAR WORKBOOK
============================================================ */

function parseWorkbook(workbook) {

const budgetSheet =
getSheetByName(
workbook,
CONFIG.budgetSheet
);


const detailSheet =
getSheetByName(
workbook,
CONFIG.detailSheet
);


const cashflowSheet =
getSheetByName(
workbook,
CONFIG.cashflowSheet
);


if (!budgetSheet) {

throw new Error(
'No encontré la hoja "Presupuesto vs Real".'
);

}


if (!detailSheet) {

throw new Error(
'No encontré la hoja "BD Plan Detallado".'
);

}


if (!cashflowSheet) {

throw new Error(
'No encontré la hoja "Flujo de Caja".'
);

}


const budget =
parseBudgetSheet(budgetSheet);


const purchases =
parseDetailSheet(detailSheet);


const cashflow =
parseCashflowSheet(cashflowSheet);


return {

budget,

purchases,

cashflow

};

}


/* ============================================================
BUSCAR HOJA SIN PROBLEMAS DE MAYÚSCULAS/ESPACIOS
============================================================ */

function getSheetByName(
workbook,
requestedName
) {

const normalizedRequested =
normalizeText(requestedName);


const foundName =
workbook.SheetNames.find(
sheetName =>
normalizeText(sheetName) ===
normalizedRequested
);


if (!foundName) return null;


return workbook.Sheets[foundName];

}


/* ============================================================
PRESUPUESTO VS REAL
============================================================ */

function parseBudgetSheet(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: null,
raw: true
}
);


let totalBudget = 0;

let totalReal = 0;

let totalDifference = 0;

let totalRowFound = false;


/* --------------------------------------------------------
TOTAL GENERAL ESTIMADO
--------------------------------------------------------- */

for (
let r = 0;
r < Math.min(rows.length, 40);
r++
) {

const row = rows[r] || [];

const joined =
normalizeText(
row
.map(value => value ?? "")
.join(" ")
);


if (
joined.includes(
"total general estimado"
)
) {

const numbers =
extractNumbers(row);


/*
En la hoja:
Presupuestado | Real | Diferencia

Tomamos los últimos tres valores
numéricos significativos de esa fila.
*/

if (numbers.length >= 3) {

const values =
numbers.slice(-3);


totalBudget =
values[0] || 0;


totalReal =
values[1] || 0;


totalDifference =
values[2] || 0;


totalRowFound = true;

}


break;

}

}


/* --------------------------------------------------------
BUSCAR ENCABEZADO DE PARTIDAS
--------------------------------------------------------- */

let headerRowIndex = -1;

let budgetColumn = -1;

let realColumn = -1;

let differenceColumn = -1;


for (
let r = 0;
r < Math.min(rows.length, 80);
r++
) {

const row = rows[r] || [];


const normalized =
row.map(normalizeText);


const hasBudget =
normalized.some(
value =>
value.includes(
"presupuestado"
)
);


const hasReal =
normalized.some(
value =>
value === "real" ||
value.includes(
"real"
)
);


const hasDifference =
normalized.some(
value =>
value.includes(
"diferencia"
)
);


if (
hasBudget &&
hasReal &&
hasDifference
) {

headerRowIndex = r;


budgetColumn =
normalized.findIndex(
value =>
value.includes(
"presupuestado"
)
);


realColumn =
normalized.findIndex(
value =>
value === "real" ||
value.includes(
"real"
)
);


differenceColumn =
normalized.findIndex(
value =>
value.includes(
"diferencia"
)
);


/*
Preferimos la última aparición,
porque la hoja puede tener etiquetas
repetidas.
*/

budgetColumn =
findLastColumnContaining(
normalized,
"presupuestado"
);


realColumn =
findLastColumnContaining(
normalized,
"real"
);


differenceColumn =
findLastColumnContaining(
normalized,
"diferencia"
);


break;

}

}


/* --------------------------------------------------------
PARTIDAS
--------------------------------------------------------- */

const categories = [];


if (headerRowIndex >= 0) {

for (
let r = headerRowIndex + 1;
r < rows.length;
r++
) {

const row = rows[r] || [];


const firstValues =
row
.slice(0, 4)
.map(value =>
cleanString(value)
);


let code = "";

let description = "";


/*
En la hoja normalmente:
A = código
B = descripción

Pero se busca de forma flexible.
*/

for (
let c = 0;
c < Math.min(4, row.length);
c++
) {

const value =
cleanString(row[c]);


if (!value) continue;


if (
!code &&
isPartCode(value)
) {

code = value;

continue;

}


if (
value &&
!isCurrencyLabel(value) &&
!isPartCode(value)
) {

if (!description) {

description = value;

}

}

}


if (
!code &&
!description
) {

continue;

}


const budget =
toNumber(
row[budgetColumn]
);


const real =
toNumber(
row[realColumn]
);


let difference =
toNumber(
row[differenceColumn]
);


if (
difference === 0 &&
(budget !== 0 || real !== 0)
) {

difference =
budget - real;

}


/*
Ignorar filas totalmente vacías
financieramente.
*/

if (
budget === 0 &&
real === 0 &&
difference === 0
) {

continue;

}


const name =
code && description
? `${code} ${description}`
: code || description;


categories.push({

code,

name,

budget,

real,

difference,

execution:
budget !== 0
? real / budget
: 0

});

}

}


/*
Si no logramos encontrar el total,
usamos los datos encontrados como
último recurso.
*/

if (!totalRowFound) {

const rootRows =
categories.filter(
category =>
isMainCategory(
category.code
)
);


const source =
rootRows.length
? rootRows
: categories;


totalBudget =
source.reduce(
(sum, item) =>
sum + item.budget,
0
);


totalReal =
source.reduce(
(sum, item) =>
sum + item.real,
0
);


totalDifference =
totalBudget - totalReal;

}


/*
La diferencia positiva es ahorro.
Convertimos únicamente el ahorro
a RD$ con tasa fija 60.
*/

const savingsUSD =
totalDifference > 0
? totalDifference
: 0;


const overrunUSD =
totalDifference < 0
? Math.abs(totalDifference)
: 0;


const savingsDOP =
savingsUSD * EXCHANGE_RATE;


const overrunDOP =
overrunUSD * EXCHANGE_RATE;


return {

totalBudget,

totalReal,

totalDifference,

execution:
totalBudget !== 0
? totalReal / totalBudget
: 0,

savingsUSD,

savingsDOP,

overrunUSD,

overrunDOP,

categories

};

}


/* ============================================================
BD PLAN DETALLADO
============================================================ */

function parseDetailSheet(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: null,
raw: true
}
);


let headerRowIndex = -1;


/* --------------------------------------------------------
ENCONTRAR ENCABEZADOS
--------------------------------------------------------- */

for (
let r = 0;
r < Math.min(rows.length, 30);
r++
) {

const normalized =
(rows[r] || [])
.map(normalizeText);


const hasPartida =
normalized.some(
value =>
value.includes(
"partida"
)
);


const hasOrder =
normalized.some(
value =>
value.includes(
"orden de compra"
)
);


const hasItem =
normalized.some(
value =>
value === "item" ||
value.includes(
"item"
)
);


if (
hasPartida &&
hasOrder &&
hasItem
) {

headerRowIndex = r;

break;

}

}


if (headerRowIndex < 0) {

console.warn(
"No se encontraron encabezados de BD Plan Detallado."
);


return createEmptyPurchaseData();

}


const headers =
makeUniqueHeaders(
rows[headerRowIndex]
);


const map =
createHeaderMap(headers);


console.log(
"Columnas BD Plan Detallado:",
map
);


const records = [];


for (
let r = headerRowIndex + 1;
r < rows.length;
r++
) {

const row = rows[r] || [];


if (
row.every(
value =>
value === null ||
value === ""
)
) {

continue;

}


const partida =
getMappedValue(
row,
map,
[
"partidas",
"partida"
]
);


const item =
getMappedValue(
row,
map,
[
"item"
]
);


const orderRaw =
getMappedValue(
row,
map,
[
"orden de compra",
"orden compra",
"oc"
]
);


const order =
cleanString(orderRaw);


const normalizedOrder =
normalizeText(order);


/*
REGLA PRINCIPAL
*/

let status = "ignored";


if (
normalizedOrder.includes(
"pendiente de compra"
)
) {

status = "pending";

}

else if (
normalizedOrder === "stock" ||
normalizedOrder.includes(
"stock"
)
) {

status = "stock";

}

else if (
normalizedOrder === "plan" ||
normalizedOrder.includes(
"plan"
)
) {

status = "ignored";

}

else if (
isValidPurchaseOrder(order)
) {

status = "po";

}

else {

status = "ignored";

}


/*
Ignoramos PLAN y cualquier
otro estado no utilizado.
*/

if (status === "ignored") {

continue;

}


const date =
getMappedValue(
row,
map,
[
"fecha"
]
);


const quantity =
toNumber(
getMappedValue(
row,
map,
[
"cant",
"cantidad"
]
)
);


const netValue =
toNumber(
getMappedValue(
row,
map,
[
"valor neto dop",
"valor neto [dop]",
"valor neto"
]
)
);


const budgetDOP =
toNumber(
getMappedValue(
row,
map,
[
"ppto orden interna dop",
"ppto. orden interna [dop]",
"ppto orden interna [dop]",
"presupuesto dop"
]
)
);


const comment =
cleanString(
getMappedValue(
row,
map,
[
"comentario",
"comentarios"
]
)
);


records.push({

partida:
cleanString(partida),

item:
cleanString(item),

order,

date,

quantity,

netValue,

budgetDOP,

comment,

status

});

}


const poRecords =
records.filter(
record =>
record.status === "po"
);


const pendingRecords =
records.filter(
record =>
record.status === "pending"
);


const stockRecords =
records.filter(
record =>
record.status === "stock"
);


const poValue =
poRecords.reduce(
(sum, record) =>
sum + record.netValue,
0
);


const pendingValue =
pendingRecords.reduce(
(sum, record) =>
sum + record.budgetDOP,
0
);


return {

records,

poRecords,

pendingRecords,

stockRecords,

poCount:
poRecords.length,

pendingCount:
pendingRecords.length,

stockCount:
stockRecords.length,

poValue,

pendingValue

};

}


/* ============================================================
FLUJO DE CAJA
============================================================ */

function parseCashflowSheet(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: null,
raw: true
}
);


/*
El Flujo de Caja no se convierte a USD.
Se mantiene en RD$.

Buscamos una fila con varios períodos.
Puede venir como:
P1 P2 P3...
o
1 2 3...
*/


let periodRowIndex = -1;

let periodColumns = [];


for (
let r = 0;
r < Math.min(rows.length, 40);
r++
) {

const row = rows[r] || [];

const candidates = [];


row.forEach(
(value, columnIndex) => {

const period =
parsePeriodNumber(value);


if (
period !== null
) {

candidates.push({

columnIndex,

period

});

}

}
);


/*
Evitar confundir códigos de partidas
con la fila de períodos.
*/

if (
candidates.length >= 4
) {

const periods =
candidates.map(
item => item.period
);


const unique =
new Set(periods);


if (
unique.size >= 4
) {

periodRowIndex = r;

periodColumns =
candidates;

break;

}

}

}


if (
periodRowIndex < 0 ||
periodColumns.length === 0
) {

console.warn(
"No se identificaron períodos en Flujo de Caja."
);


return {

periods: [],

rows: [],

groups: [],

totals: [],

grandTotal: 0

};

}


/*
Buscar nombres de meses cerca
de la fila de períodos.
*/

const periods =
periodColumns.map(
item => {

let label = "";


for (
let r =
Math.max(
0,
periodRowIndex - 3
);
r <=
Math.min(
rows.length - 1,
periodRowIndex + 3
);
r++
) {

if (
r === periodRowIndex
) {

continue;

}


const possible =
cleanString(
rows[r]?.[
item.columnIndex
]
);


if (
isMonthName(possible)
) {

label = possible;

break;

}

}


if (!label) {

label =
periodToMonthLabel(
item.period
);

}


return {

period:
item.period,

columnIndex:
item.columnIndex,

label

};

}
);


/*
Encontrar columna donde están
las partidas.
*/

let partidaColumn = 0;


for (
let c = 0;
c <
Math.min(
periodColumns[0].columnIndex,
8
);
c++
) {

let score = 0;


for (
let r = periodRowIndex + 1;
r <
Math.min(
rows.length,
periodRowIndex + 80
);
r++
) {

const value =
cleanString(
rows[r]?.[c]
);


if (
value &&
(
isPartCode(value) ||
value.length > 5
)
) {

score++;

}

}


if (score > 3) {

partidaColumn = c;

break;

}

}


const parsedRows = [];


for (
let r = periodRowIndex + 1;
r < rows.length;
r++
) {

const row = rows[r] || [];


let label =
cleanString(
row[partidaColumn]
);


/*
Si la descripción está en la
columna siguiente, unirla.
*/

const nextValue =
cleanString(
row[
partidaColumn + 1
]
);


if (
label &&
isPartCode(label) &&
nextValue &&
!isNumericLike(nextValue)
) {

label =
`${label} ${nextValue}`;

}


if (!label) {

continue;

}


const values =
periods.map(
period =>
toNumber(
row[
period.columnIndex
]
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


const code =
extractPartCode(label);


parsedRows.push({

label,

code,

values,

total,

rowIndex: r

});

}


/*
Crear jerarquía desplegable.
*/

const groups =
buildCashflowGroups(
parsedRows,
periods.length
);


const totals =
new Array(
periods.length
).fill(0);


parsedRows.forEach(
row => {

row.values.forEach(
(value, index) => {

totals[index] +=
value;

}
);

}
);


const grandTotal =
totals.reduce(
(sum, value) =>
sum + value,
0
);


return {

periods,

rows: parsedRows,

groups,

totals,

grandTotal

};

}


/* ============================================================
AGRUPAR FLUJO DE CAJA
============================================================ */

function buildCashflowGroups(
rows,
periodCount
) {

const groups = [];

let currentGroup = null;


rows.forEach(
row => {

if (
isMainCategory(
row.code
)
) {

currentGroup = {

id:
`cashflow-group-${groups.length}`,

label:
row.label,

code:
row.code,

values:
[...row.values],

total:
row.total,

children: []

};


groups.push(
currentGroup
);

}

else {

if (!currentGroup) {

currentGroup = {

id:
`cashflow-group-${groups.length}`,

label:
"Otras partidas",

code:
"",

values:
new Array(
periodCount
).fill(0),

total: 0,

children: []

};


groups.push(
currentGroup
);

}


currentGroup.children.push(
row
);

}

}
);


/*
Si una fila principal no contiene
realmente el subtotal, calcularlo
desde sus hijos.
*/

groups.forEach(
group => {

if (
group.children.length > 0 &&
group.total === 0
) {

group.values =
new Array(
periodCount
).fill(0);


group.children.forEach(
child => {

child.values.forEach(
(value, index) => {

group.values[index] +=
value;

}
);

}
);


group.total =
group.values.reduce(
(sum, value) =>
sum + value,
0
);

}

}
);


return groups;

}


/* ============================================================
RENDER GENERAL
============================================================ */

function renderDashboard() {

if (!dashboardData) return;


renderHeader();

renderHome();

renderExecutive();

renderBudget();

renderPurchases();

renderCashflow();

renderTracking();

renderPending();

}


/* ============================================================
HEADER
============================================================ */

function renderHeader() {

if (!currentProject) return;


setText(
"topProjectName",
currentProject.name
);


setText(
"topProjectSubtitle",
"Seguimiento integral de presupuesto, compras y flujo de caja."
);


const opening =
parseDateOnly(
currentProject.openingDate
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


const days =
calculateDaysToOpening(
opening
);


setText(
"openingDate",
formatted
);


setText(
"daysToOpening",
days
);


setText(
"executiveOpeningDate",
formatted
);


setText(
"executiveDays",
days
);

}


/* ============================================================
HOME
============================================================ */

function renderHome() {

const budget =
dashboardData.budget;


const purchases =
dashboardData.purchases;


setText(
"homeBudget",
formatUSD(
budget.totalBudget
)
);


setText(
"homeReal",
formatUSD(
budget.totalReal
)
);


setText(
"homeDifference",
formatUSD(
budget.totalDifference
)
);


setText(
"homeExecution",
formatPercent(
budget.execution
)
);


setText(
"homePOCount",
purchases.poCount
);


setText(
"homePendingCount",
purchases.pendingCount
);


/*
Esta tarjeta se agregará en el
index actualizado.
*/

if (
document.getElementById(
"homeSavingsDOP"
)
) {

setText(
"homeSavingsDOP",
budget.totalDifference >= 0
? formatDOP(
budget.savingsDOP
)
: formatDOP(
-budget.overrunDOP
)
);

}


renderBudgetRealChart(
"homeBudgetChart",
budget.totalBudget,
budget.totalReal
);


renderPurchaseStatusChart(
"homePurchaseChart",
purchases
);


renderCategoryTable(
"homeCategoryTable",
budget.categories
);

}


/* ============================================================
RESUMEN EJECUTIVO
============================================================ */

function renderExecutive() {

const budget =
dashboardData.budget;


const purchases =
dashboardData.purchases;


setText(
"executiveBudget",
formatUSD(
budget.totalBudget
)
);


setText(
"executiveReal",
formatUSD(
budget.totalReal
)
);


setText(
"executiveDifference",
formatUSD(
budget.totalDifference
)
);


setText(
"executiveExecution",
formatPercent(
budget.execution
)
);


setText(
"executivePO",
purchases.poCount
);


setText(
"executivePending",
purchases.pendingCount
);


if (
document.getElementById(
"executiveSavingsDOP"
)
) {

setText(
"executiveSavingsDOP",
budget.totalDifference >= 0
? formatDOP(
budget.savingsDOP
)
: formatDOP(
-budget.overrunDOP
)
);

}


const percentage =
Math.max(
0,
Math.min(
budget.execution * 100,
100
)
);


setText(
"executiveProgressText",
`${percentage.toFixed(1)}%`
);


const progressBar =
document.getElementById(
"executiveProgressBar"
);


if (progressBar) {

progressBar.style.width =
`${percentage}%`;

}


renderCategoryChart(
"executiveCategoryChart",
budget.categories
);

}


/* ============================================================
PRESUPUESTO VS REAL
============================================================ */

function renderBudget() {

const budget =
dashboardData.budget;


setText(
"budgetTotal",
formatUSD(
budget.totalBudget
)
);


setText(
"realTotal",
formatUSD(
budget.totalReal
)
);


setText(
"differenceTotal",
formatUSD(
budget.totalDifference
)
);


setText(
"budgetExecution",
formatPercent(
budget.execution
)
);


if (
document.getElementById(
"budgetSavingsDOP"
)
) {

setText(
"budgetSavingsDOP",
budget.totalDifference >= 0
? formatDOP(
budget.savingsDOP
)
: formatDOP(
-budget.overrunDOP
)
);

}


renderCategoryChart(
"budgetVsRealChart",
budget.categories
);


renderBudgetTable();

}


/* ============================================================
TABLA PRESUPUESTO
============================================================ */

function renderBudgetTable() {

if (!dashboardData) return;


const table =
document.getElementById(
"budgetTable"
);


if (!table) return;


const tbody =
table.querySelector("tbody");


if (!tbody) return;


const search =
normalizeText(
document.getElementById(
"budgetSearch"
)?.value || ""
);


const rows =
dashboardData.budget.categories
.filter(
item =>
!search ||
normalizeText(
item.name
).includes(search)
);


tbody.innerHTML = "";


rows.forEach(
item => {

const tr =
document.createElement("tr");


tr.innerHTML = `
<td>${escapeHTML(item.name)}</td>

<td class="number">
${formatUSD(item.budget)}
</td>

<td class="number">
${formatUSD(item.real)}
</td>

<td class="number">
${formatUSD(item.difference)}
</td>

<td class="number">
${formatPercent(item.execution)}
</td>
`;


tbody.appendChild(tr);

}
);

}


/* ============================================================
COMPRAS
============================================================ */

function renderPurchases() {

const data =
dashboardData.purchases;


setText(
"purchasePOCount",
data.poCount
);


setText(
"purchasePendingCount",
data.pendingCount
);


setText(
"purchaseStockCount",
data.stockCount
);


setText(
"purchasePOValue",
formatDOP(
data.poValue
)
);


renderPurchaseTable();

}


/* ============================================================
TABLA DE COMPRAS
============================================================ */

function renderPurchaseTable() {

if (!dashboardData) return;


const table =
document.getElementById(
"purchaseTable"
);


if (!table) return;


const tbody =
table.querySelector("tbody");


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


let records =
dashboardData.purchases.records;


if (status !== "all") {

records =
records.filter(
record =>
record.status === status
);

}


if (search) {

records =
records.filter(
record => {

const searchable =
normalizeText(
[
record.partida,
record.item,
record.order
].join(" ")
);


return searchable.includes(
search
);

}
);

}


tbody.innerHTML = "";


records.forEach(
record => {

const tr =
document.createElement("tr");


tr.innerHTML = `
<td>
${escapeHTML(record.partida)}
</td>

<td>
${escapeHTML(record.item)}
</td>

<td>
${escapeHTML(record.order)}
</td>

<td>
${formatExcelDate(record.date)}
</td>

<td class="number">
${formatNumber(record.quantity)}
</td>

<td class="number">
${formatDOP(record.netValue)}
</td>

<td>
${statusBadge(record.status)}
</td>
`;


tbody.appendChild(tr);

}
);

}


/* ============================================================
PENDIENTES
============================================================ */

function renderPending() {

const data =
dashboardData.purchases;


setText(
"pendingTotal",
data.pendingCount
);


const uniqueCategories =
new Set(
data.pendingRecords
.map(
item =>
normalizeText(
item.partida
)
)
.filter(Boolean)
);


setText(
"pendingCategoryCount",
uniqueCategories.size
);


setText(
"pendingValue",
formatDOP(
data.pendingValue
)
);


renderPendingTable();

}


/* ============================================================
TABLA PENDIENTES
============================================================ */

function renderPendingTable() {

if (!dashboardData) return;


const table =
document.getElementById(
"pendingTable"
);


if (!table) return;


const tbody =
table.querySelector("tbody");


if (!tbody) return;


const search =
normalizeText(
document.getElementById(
"pendingSearch"
)?.value || ""
);


let records =
dashboardData
.purchases
.pendingRecords;


if (search) {

records =
records.filter(
record => {

const searchable =
normalizeText(
[
record.partida,
record.item,
record.comment
].join(" ")
);


return searchable.includes(
search
);

}
);

}


tbody.innerHTML = "";


records.forEach(
record => {

const tr =
document.createElement("tr");


tr.innerHTML = `
<td>
${escapeHTML(record.partida)}
</td>

<td>
${escapeHTML(record.item)}
</td>

<td class="number">
${formatNumber(record.quantity)}
</td>

<td class="number">
${formatDOP(record.budgetDOP)}
</td>

<td>
${escapeHTML(record.comment)}
</td>
`;


tbody.appendChild(tr);

}
);

}


/* ============================================================
FLUJO DE CAJA
============================================================ */

function renderCashflow() {

const data =
dashboardData.cashflow;


setText(
"cashflowTotal",
formatDOP(
data.grandTotal
)
);


setText(
"cashflowCategoryCount",
data.groups.length
);


setText(
"cashflowPeriodCount",
data.periods.length
);


renderCashflowChart();

renderCashflowTable();

}


/* ============================================================
GRÁFICA FLUJO
============================================================ */

function renderCashflowChart() {

const data =
dashboardData.cashflow;


const labels =
data.periods.map(
period =>
period.label
);


createChart(
"cashflowChart",
{
type: "bar",

data: {

labels,

datasets: [

{
label:
"Flujo RD$",

data:
data.totals,

borderWidth: 0
}

]

},

options: {

responsive: true,

maintainAspectRatio: false,

plugins: {

legend: {
display: false
},

tooltip: {

callbacks: {

label: context =>
formatDOP(
context.raw
)

}

}

},

scales: {

y: {

beginAtZero: true,

ticks: {

callback: value =>
compactCurrency(
value,
"RD$"
)

}

}

}

}

}
);

}


/* ============================================================
TABLA FLUJO DESPLEGABLE
============================================================ */

function renderCashflowTable() {

const data =
dashboardData.cashflow;


const thead =
document.getElementById(
"cashflowTableHead"
);


const tbody =
document.getElementById(
"cashflowTableBody"
);


if (!thead || !tbody) return;


/* ENCABEZADO */

let headerHTML = `
<tr>
<th class="cashflow-partida-column">
Partida
</th>
`;


data.periods.forEach(
period => {

headerHTML += `
<th class="number">
${escapeHTML(period.label)}
</th>
`;

}
);


headerHTML += `
<th class="number">
Total
</th>
</tr>
`;


thead.innerHTML =
headerHTML;


/* CUERPO */

tbody.innerHTML = "";


data.groups.forEach(
(group, groupIndex) => {

const groupId =
`cashflow-${groupIndex}`;


const parent =
document.createElement("tr");


parent.className =
"cashflow-group-row";


let parentHTML = `
<td>
<button
type="button"
class="cashflow-toggle"
data-group="${groupId}"
aria-expanded="false"
>
<span class="toggle-symbol">
+
</span>

<strong>
${escapeHTML(group.label)}
</strong>
</button>
</td>
`;


group.values.forEach(
value => {

parentHTML += `
<td class="number">
${formatDOPCell(value)}
</td>
`;

}
);


parentHTML += `
<td class="number">
<strong>
${formatDOPCell(group.total)}
</strong>
</td>
`;


parent.innerHTML =
parentHTML;


tbody.appendChild(
parent
);


group.children.forEach(
child => {

const tr =
document.createElement("tr");


tr.className =
`cashflow-child-row ${groupId}`;


tr.style.display =
"none";


let html = `
<td class="cashflow-child-label">
${escapeHTML(child.label)}
</td>
`;


child.values.forEach(
value => {

html += `
<td class="number">
${formatDOPCell(value)}
</td>
`;

}
);


html += `
<td class="number">
${formatDOPCell(child.total)}
</td>
`;


tr.innerHTML =
html;


tbody.appendChild(
tr
);

}
);

}
);


/* TOTAL GENERAL */

if (data.periods.length) {

const totalRow =
document.createElement("tr");


totalRow.className =
"cashflow-total-row";


let totalHTML = `
<td>
<strong>
TOTAL
</strong>
</td>
`;


data.totals.forEach(
value => {

totalHTML += `
<td class="number">
<strong>
${formatDOPCell(value)}
</strong>
</td>
`;

}
);


totalHTML += `
<td class="number">
<strong>
${formatDOPCell(data.grandTotal)}
</strong>
</td>
`;


totalRow.innerHTML =
totalHTML;


tbody.appendChild(
totalRow
);

}


/*
Eventos desplegables
*/

document
.querySelectorAll(
".cashflow-toggle"
)
.forEach(
button => {

button.addEventListener(
"click",
() => {

toggleCashflowGroup(
button
);

}
);

}
);

}


/* ============================================================
ABRIR / CERRAR GRUPO FLUJO
============================================================ */

function toggleCashflowGroup(button) {

const group =
button.dataset.group;


const expanded =
button.getAttribute(
"aria-expanded"
) === "true";


const rows =
document.querySelectorAll(
`.${group}`
);


rows.forEach(
row => {

row.style.display =
expanded
? "none"
: "table-row";

}
);


button.setAttribute(
"aria-expanded",
String(!expanded)
);


const symbol =
button.querySelector(
".toggle-symbol"
);


if (symbol) {

symbol.textContent =
expanded
? "+"
: "−";

}

}


/* ============================================================
EXPANDIR / CONTRAER TODO
============================================================ */

function toggleAllCashflow(expand) {

document
.querySelectorAll(
".cashflow-toggle"
)
.forEach(
button => {

const group =
button.dataset.group;


document
.querySelectorAll(
`.${group}`
)
.forEach(
row => {

row.style.display =
expand
? "table-row"
: "none";

}
);


button.setAttribute(
"aria-expanded",
String(expand)
);


const symbol =
button.querySelector(
".toggle-symbol"
);


if (symbol) {

symbol.textContent =
expand
? "−"
: "+";

}

}
);

}


/* ============================================================
SEGUIMIENTO
============================================================ */

function renderTracking() {

const budget =
dashboardData.budget;


const purchases =
dashboardData.purchases;


renderPurchaseStatusChart(
"trackingStatusChart",
purchases
);


renderBudgetRealChart(
"trackingFinancialChart",
budget.totalBudget,
budget.totalReal
);


const table =
document.getElementById(
"trackingTable"
);


if (!table) return;


const tbody =
table.querySelector("tbody");


if (!tbody) return;


tbody.innerHTML = "";


budget.categories.forEach(
category => {

const normalizedCategory =
normalizeText(
category.name
);


const matching =
purchases.records.filter(
record => {

const partida =
normalizeText(
record.partida
);


if (
!partida ||
!normalizedCategory
) {

return false;

}


return (
normalizedCategory.includes(
partida
) ||
partida.includes(
normalizedCategory
)
);

}
);


const po =
matching.filter(
item =>
item.status === "po"
).length;


const pending =
matching.filter(
item =>
item.status === "pending"
).length;


const stock =
matching.filter(
item =>
item.status === "stock"
).length;


const tr =
document.createElement("tr");


tr.innerHTML = `
<td>
${escapeHTML(category.name)}
</td>

<td class="number">
${formatUSD(category.budget)}
</td>

<td class="number">
${formatUSD(category.real)}
</td>

<td class="number">
${formatPercent(category.execution)}
</td>

<td class="number">
${po}
</td>

<td class="number">
${pending}
</td>

<td class="number">
${stock}
</td>
`;


tbody.appendChild(
tr
);

}
);

}


/* ============================================================
TABLA GENÉRICA DE PARTIDAS
============================================================ */

function renderCategoryTable(
tableId,
categories
) {

const table =
document.getElementById(
tableId
);


if (!table) return;


const tbody =
table.querySelector("tbody");


if (!tbody) return;


tbody.innerHTML = "";


categories.forEach(
item => {

const tr =
document.createElement("tr");


tr.innerHTML = `
<td>
${escapeHTML(item.name)}
</td>

<td class="number">
${formatUSD(item.budget)}
</td>

<td class="number">
${formatUSD(item.real)}
</td>

<td class="number">
${formatUSD(item.difference)}
</td>

<td class="number">
${formatPercent(item.execution)}
</td>
`;


tbody.appendChild(
tr
);

}
);

}


/* ============================================================
GRÁFICAS
============================================================ */

function renderBudgetRealChart(
canvasId,
budget,
real
) {

createChart(
canvasId,
{

type: "bar",

data: {

labels: [
"Presupuesto",
"Real"
],

datasets: [

{
data: [
budget,
real
],

borderWidth: 0
}

]

},

options: {

responsive: true,

maintainAspectRatio: false,

plugins: {

legend: {
display: false
},

tooltip: {

callbacks: {

label: context =>
formatUSD(
context.raw
)

}

}

},

scales: {

y: {

beginAtZero: true,

ticks: {

callback: value =>
compactCurrency(
value,
"US$"
)

}

}

}

}

}
);

}


/* ============================================================
GRÁFICA ESTADO COMPRAS
============================================================ */

function renderPurchaseStatusChart(
canvasId,
purchases
) {

createChart(
canvasId,
{

type: "doughnut",

data: {

labels: [
"Con OC",
"Pendiente de compra",
"Stock"
],

datasets: [

{
data: [
purchases.poCount,
purchases.pendingCount,
purchases.stockCount
],

borderWidth: 2
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

}
);

}


/* ============================================================
GRÁFICA PARTIDAS
============================================================ */

function renderCategoryChart(
canvasId,
categories
) {

/*
Evitamos una gráfica interminable.
Mostramos las principales partidas
con presupuesto.
*/

const filtered =
categories
.filter(
item =>
item.budget !== 0 ||
item.real !== 0
)
.slice(0, 15);


createChart(
canvasId,
{

type: "bar",

data: {

labels:
filtered.map(
item =>
item.name
),

datasets: [

{
label:
"Presupuesto",

data:
filtered.map(
item =>
item.budget
)
},

{
label:
"Real",

data:
filtered.map(
item =>
item.real
)
}

]

},

options: {

responsive: true,

maintainAspectRatio: false,

indexAxis: "y",

plugins: {

legend: {

position: "bottom"

},

tooltip: {

callbacks: {

label: context =>
`${context.dataset.label}: ${formatUSD(context.raw)}`

}

}

},

scales: {

x: {

beginAtZero: true,

ticks: {

callback: value =>
compactCurrency(
value,
"US$"
)

}

}

}

}

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


if (!canvas) return;


if (
typeof Chart === "undefined"
) {

console.error(
"Chart.js no está cargado."
);

return;

}


if (charts[canvasId]) {

charts[canvasId].destroy();

}


charts[canvasId] =
new Chart(
canvas,
config
);

}


/* ============================================================
REDIMENSIONAR CHARTS
============================================================ */

function resizeCharts() {

setTimeout(
() => {

Object
.values(charts)
.forEach(
chart => {

try {

chart.resize();

} catch (_) {}

}
);

},
100
);

}


/* ============================================================
ESTADO OC
============================================================ */

function statusBadge(status) {

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
VALIDAR OC
============================================================ */

function isValidPurchaseOrder(value) {

if (
value === null ||
value === undefined
) {

return false;

}


const text =
cleanString(value);


if (!text) return false;


const normalized =
normalizeText(text);


if (
normalized === "plan" ||
normalized.includes(
"pendiente"
) ||
normalized.includes(
"stock"
)
) {

return false;

}


/*
Las OC reales observadas tienen
formato numérico largo.
*/

const digits =
text.replace(
/\D/g,
""
);


return digits.length >= 6;

}


/* ============================================================
HEADER MAP
============================================================ */

function createHeaderMap(headers) {

const map = {};


headers.forEach(
(header, index) => {

const normalized =
normalizeText(header);


if (!normalized) return;


map[normalized] =
index;

}
);


return map;

}


/* ============================================================
OBTENER VALOR POR POSIBLES ENCABEZADOS
============================================================ */

function getMappedValue(
row,
map,
possibleNames
) {

for (
const name of possibleNames
) {

const normalized =
normalizeText(name);


/*
Coincidencia exacta
*/

if (
Object.prototype
.hasOwnProperty
.call(
map,
normalized
)
) {

return row[
map[normalized]
];

}


/*
Coincidencia flexible
*/

const found =
Object.keys(map)
.find(
key =>
key.includes(
normalized
) ||
normalized.includes(
key
)
);


if (found) {

return row[
map[found]
];

}

}


return null;

}


/* ============================================================
HEADERS ÚNICOS
============================================================ */

function makeUniqueHeaders(row) {

const counts = {};


return row.map(
value => {

const base =
cleanString(value);


if (!base) return "";


const normalized =
normalizeText(base);


counts[normalized] =
(
counts[normalized] ||
0
) + 1;


if (
counts[normalized] === 1
) {

return base;

}


return `${base} ${counts[normalized]}`;

}
);

}


/* ============================================================
CASHFLOW: PERÍODO
============================================================ */

function parsePeriodNumber(value) {

if (
value === null ||
value === undefined ||
value === ""
) {

return null;

}


if (
typeof value === "number" &&
Number.isInteger(value) &&
value >= 1 &&
value <= 60
) {

return value;

}


const text =
normalizeText(value);


let match =
text.match(
/^p\s*(\d{1,2})$/
);


if (match) {

const number =
Number(match[1]);


if (
number >= 1 &&
number <= 60
) {

return number;

}

}


match =
text.match(
/^(\d{1,2})$/
);


if (match) {

const number =
Number(match[1]);


if (
number >= 1 &&
number <= 60
) {

return number;

}

}


return null;

}


/* ============================================================
PERÍODO 1 = ENERO 2025
============================================================ */

function periodToMonthLabel(period) {

if (
!Number.isFinite(period) ||
period < 1
) {

return `P${period}`;

}


const date =
new Date(
2025,
period - 1,
1
);


const month =
date.toLocaleDateString(
"es-DO",
{
month: "short"
}
);


const year =
date.getFullYear();


return (
month
.replace(".", "")
.replace(
/^./,
char =>
char.toUpperCase()
)
+
` ${year}`
);

}


/* ============================================================
MESES
============================================================ */

function isMonthName(value) {

const text =
normalizeText(value);


if (!text) return false;


const months = [

"enero",
"febrero",
"marzo",
"abril",
"mayo",
"junio",
"julio",
"agosto",
"septiembre",
"setiembre",
"octubre",
"noviembre",
"diciembre",

"ene",
"feb",
"mar",
"abr",
"may",
"jun",
"jul",
"ago",
"sep",
"oct",
"nov",
"dic"

];


return months.some(
month =>
text.includes(month)
);

}


/* ============================================================
CÓDIGOS DE PARTIDA
============================================================ */

function isPartCode(value) {

const text =
cleanString(value);


return (
/^\d+(\.\d+)+$/.test(text) ||
/^\d+\.\d+$/.test(text)
);

}


/* ============================================================
EXTRAER CÓDIGO
============================================================ */

function extractPartCode(value) {

const text =
cleanString(value);


const match =
text.match(
/^(\d+(?:\.\d+)+)/
);


return match
? match[1]
: "";

}


/* ============================================================
PARTIDA PRINCIPAL
============================================================ */

function isMainCategory(code) {

if (!code) return false;


const match =
String(code).match(
/^(\d+)\.(\d+)$/
);


if (!match) return false;


return Number(match[2]) === 0;

}


/* ============================================================
NÚMEROS
============================================================ */

function toNumber(value) {

if (
value === null ||
value === undefined ||
value === ""
) {

return 0;

}


if (
typeof value === "number"
) {

return Number.isFinite(value)
? value
: 0;

}


let text =
String(value)
.trim();


if (!text) return 0;


let negative = false;


/*
Excel puede mostrar negativos:
(5,160.86)
*/

if (
text.startsWith("(") &&
text.endsWith(")")
) {

negative = true;

}


text =
text
.replace(/[()]/g, "")
.replace(/RD\$/gi, "")
.replace(/US\$/gi, "")
.replace(/USD/gi, "")
.replace(/\$/g, "")
.replace(/\s/g, "")
.replace(/,/g, "");


const number =
Number(text);


if (
!Number.isFinite(number)
) {

return 0;

}


return negative
? -number
: number;

}


/* ============================================================
EXTRAER NÚMEROS DE UNA FILA
============================================================ */

function extractNumbers(row) {

return row
.map(
value => {

if (
typeof value === "number" &&
Number.isFinite(value)
) {

return value;

}


const number =
toNumber(value);


if (
number !== 0
) {

return number;

}


return null;

}
)
.filter(
value =>
value !== null
);

}


/* ============================================================
ÚLTIMA COLUMNA QUE CONTIENE TEXTO
============================================================ */

function findLastColumnContaining(
values,
search
) {

const target =
normalizeText(search);


for (
let i = values.length - 1;
i >= 0;
i--
) {

if (
values[i].includes(
target
)
) {

return i;

}

}


return -1;

}


/* ============================================================
FORMATOS MONEDA
============================================================ */

function formatUSD(value) {

const number =
Number(value) || 0;


return (
"US$ " +
number.toLocaleString(
"en-US",
{
minimumFractionDigits: 2,
maximumFractionDigits: 2
}
)
);

}


/* ============================================================
RD$
============================================================ */

function formatDOP(value) {

const number =
Number(value) || 0;


return (
"RD$ " +
number.toLocaleString(
"en-US",
{
minimumFractionDigits: 2,
maximumFractionDigits: 2
}
)
);

}


/* ============================================================
CELDA RD$
============================================================ */

function formatDOPCell(value) {

const number =
Number(value) || 0;


if (number === 0) {

return "—";

}


return (
"RD$ " +
number.toLocaleString(
"en-US",
{
minimumFractionDigits: 2,
maximumFractionDigits: 2
}
)
);

}


/* ============================================================
PORCENTAJE
============================================================ */

function formatPercent(value) {

const number =
Number(value) || 0;


return (
number * 100
).toLocaleString(
"es-DO",
{
minimumFractionDigits: 1,
maximumFractionDigits: 1
}
) + "%";

}


/* ============================================================
NÚMERO
============================================================ */

function formatNumber(value) {

const number =
Number(value) || 0;


return number.toLocaleString(
"en-US",
{
maximumFractionDigits: 2
}
);

}


/* ============================================================
MONEDA COMPACTA
============================================================ */

function compactCurrency(
value,
prefix
) {

const number =
Number(value) || 0;


const abs =
Math.abs(number);


if (abs >= 1000000) {

return (
prefix +
" " +
(
number /
1000000
).toFixed(1) +
"M"
);

}


if (abs >= 1000) {

return (
prefix +
" " +
(
number /
1000
).toFixed(0) +
"K"
);

}


return (
prefix +
" " +
number.toFixed(0)
);

}


/* ============================================================
FECHAS EXCEL
============================================================ */

function formatExcelDate(value) {

if (!value) return "—";


if (
value instanceof Date &&
!Number.isNaN(
value.getTime()
)
) {

return formatDate(value);

}


if (
typeof value === "number"
) {

const parsed =
XLSX.SSF.parse_date_code(
value
);


if (parsed) {

const date =
new Date(
parsed.y,
parsed.m - 1,
parsed.d
);


return formatDate(date);

}

}


const parsed =
new Date(value);


if (
!Number.isNaN(
parsed.getTime()
)
) {

return formatDate(parsed);

}


return cleanString(value);

}


/* ============================================================
DATE ONLY
============================================================ */

function parseDateOnly(value) {

if (!value) return null;


const match =
String(value).match(
/^(\d{4})-(\d{2})-(\d{2})$/
);


if (match) {

return new Date(
Number(match[1]),
Number(match[2]) - 1,
Number(match[3])
);

}


const date =
new Date(value);


if (
Number.isNaN(
date.getTime()
)
) {

return null;

}


return date;

}


/* ============================================================
FORMATO FECHA
============================================================ */

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


/* ============================================================
DÍAS PARA APERTURA
============================================================ */

function calculateDaysToOpening(
opening
) {

const today =
new Date();


const current =
new Date(
today.getFullYear(),
today.getMonth(),
today.getDate()
);


const target =
new Date(
opening.getFullYear(),
opening.getMonth(),
opening.getDate()
);


const difference =
target - current;


return Math.ceil(
difference /
86400000
);

}


/* ============================================================
NORMALIZAR TEXTO
============================================================ */

function normalizeText(value) {

return String(
value ?? ""
)
.trim()
.toLowerCase()
.normalize("NFD")
.replace(
/[\u0300-\u036f]/g,
""
)
.replace(
/\s+/g,
" "
);

}


/* ============================================================
LIMPIAR STRING
============================================================ */

function cleanString(value) {

if (
value === null ||
value === undefined
) {

return "";

}


return String(value).trim();

}


/* ============================================================
¿ES MONEDA?
============================================================ */

function isCurrencyLabel(value) {

const text =
normalizeText(value);


return (
text === "usd" ||
text === "us$" ||
text === "$" ||
text === "rd$" ||
text === "dop"
);

}


/* ============================================================
¿PARECE NÚMERO?
============================================================ */

function isNumericLike(value) {

const text =
cleanString(value);


if (!text) return false;


return (
/^[-+]?[\d,.()$ ]+$/.test(
text
)
);

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


if (element) {

element.textContent =
value;

}

}


/* ============================================================
ESCAPAR HTML
============================================================ */

function escapeHTML(value) {

return String(
value ?? ""
)
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


/* ============================================================
LOADING
============================================================ */

function showLoading(
show,
text = ""
) {

const overlay =
document.getElementById(
"loadingOverlay"
);


if (!overlay) return;


if (show) {

overlay.classList.remove(
"hidden"
);

} else {

overlay.classList.add(
"hidden"
);

}


if (text) {

setText(
"loadingText",
text
);

}

}


/* ============================================================
STATUS
============================================================ */

function updateLoadStatus(text) {

setText(
"projectLoadStatus",
text
);

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

}


/* ============================================================
DATA VACÍA
============================================================ */

function createEmptyPurchaseData() {

return {

records: [],

poRecords: [],

pendingRecords: [],

stockRecords: [],

poCount: 0,

pendingCount: 0,

stockCount: 0,

poValue: 0,

pendingValue: 0

};

}


/* ============================================================
RESET
============================================================ */

function resetDashboard() {

dashboardData = {

budget: {

totalBudget: 0,

totalReal: 0,

totalDifference: 0,

execution: 0,

savingsUSD: 0,

savingsDOP: 0,

overrunUSD: 0,

overrunDOP: 0,

categories: []

},

purchases:
createEmptyPurchaseData(),

cashflow: {

periods: [],

rows: [],

groups: [],

totals: [],

grandTotal: 0

}

};


renderDashboard();

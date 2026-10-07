/* =========================================================
DASHBOARD DE SEGUIMIENTO DE INICIATIVAS
GRUPO RAMOS
========================================================= */

const state = {
projects: [],
project: null,
workbook: null,

planOI: {
approved: 0,
real: 0,
difference: 0,
categories: []
},

budget: [],
cashflow: [],
issues: [],
considerations: [],

approvedBudget: 0,
currentBudget: 0,
orderedValue: 0,
pendingValue: 0,
savings: 0,

charts: {},

// Referencia general.
// Las filas de BD Plan Detallado pueden traer su propia tasa.
exchangeRate: 63
};

const $ = id => document.getElementById(id);

const COLORS = {
blue: "#064c8c",
blue2: "#0b6fb8",
cyan: "#28a9e0",
green: "#16a05d",
yellow: "#f4b323",
red: "#d94a4a",
gray: "#cbd7e2",
ink: "#18324b"
};

let countdownInterval = null;


/* =========================================================
INICIO
========================================================= */

document.addEventListener("DOMContentLoaded", init);

async function init() {

setupNavigation();
setupBudgetFilters();
setupLiquidation();
setupCashflowCurrency();

await loadProjects();

}


/* =========================================================
PROYECTOS
========================================================= */

async function loadProjects() {

try {

const response = await fetch(
"data/projects.json?v=" + Date.now()
);

if (!response.ok) {
throw new Error(
"No se pudo cargar data/projects.json"
);
}

const config = await response.json();

state.projects = Array.isArray(config.projects)
? config.projects
: [];

if (!state.projects.length) {
throw new Error(
"projects.json no contiene proyectos."
);
}

setupProjectSelector();

const saved =
localStorage.getItem("selectedProject");

const initial =
state.projects.find(
project => project.id === saved
) || state.projects[0];

if ($("projectSelect")) {
$("projectSelect").value = initial.id;
}

await loadProject(initial);

} catch (error) {

console.error(error);

showLoader(false);

alert(
"No se pudo cargar la lista de iniciativas. Revisa data/projects.json."
);

}

}


function setupProjectSelector() {

const select = $("projectSelect");

if (!select) return;

select.innerHTML = "";

state.projects.forEach(project => {

const option =
document.createElement("option");

option.value = project.id;

option.textContent =
project.name || project.id;

select.appendChild(option);

});

select.onchange = async event => {

const project =
state.projects.find(
item =>
item.id === event.target.value
);

if (!project) return;

localStorage.setItem(
"selectedProject",
project.id
);

await loadProject(project);

};

}


/* =========================================================
CARGAR PROYECTO
========================================================= */

async function loadProject(project) {

showLoader(true);

state.project = project;

updateProjectLabels(project);

try {

const response = await fetch(
project.file + "?v=" + Date.now()
);

if (!response.ok) {

throw new Error(
`No se encontró ${project.file}`
);

}

const buffer =
await response.arrayBuffer();

state.workbook =
XLSX.read(
buffer,
{
type: "array",
cellDates: true,
cellFormula: true,
cellNF: true
}
);

console.log(
"Hojas encontradas:",
state.workbook.SheetNames
);

parseWorkbook();

renderAll();

} catch (error) {

console.error(
"Error cargando proyecto:",
error
);

resetProjectData();

renderEmptyState();

}

showLoader(false);

}


function resetProjectData() {

state.planOI = {
approved: 0,
real: 0,
difference: 0,
categories: []
};

state.budget = [];
state.cashflow = [];
state.issues = [];
state.considerations = [];

state.approvedBudget = 0;
state.currentBudget = 0;
state.orderedValue = 0;
state.pendingValue = 0;
state.savings = 0;

}


/* =========================================================
ETIQUETAS DEL PROYECTO
========================================================= */

function updateProjectLabels(project) {

const name =
project.name ||
project.id ||
"Proyecto";

const brand =
project.brand ||
"Sirena";

setText(
"projectBrand",
brand
);

setText(
"projectName",
name
);

setText(
"homeProjectName",
name
);

setText(
"overviewTitle",
`${brand} ${name}`
);

setText(
"countdownProject",
name
);

setText(
"footerProject",
`Proyecto ${name}`
);


if (project.openingDate) {

const opening =
parseProjectDate(
project.openingDate
);

if (opening) {

setText(
"homeOpeningDate",
formatDate(opening)
);

setText(
"countdownDateLabel",
formatLongDate(opening)
);

startCountdown(
project.openingDate
);

}

} else {

setText(
"homeOpeningDate",
"Por definir"
);

setText(
"countdownDateLabel",
"Por definir"
);

setText("cdDays", "—");
setText("cdHours", "—");
setText("cdMinutes", "—");
setText("cdSeconds", "—");

setText(
"countdownStatus",
"Fecha de apertura no definida"
);

}

}


/* =========================================================
NAVEGACIÓN
========================================================= */

function setupNavigation() {

document
.querySelectorAll("#navTabs button")
.forEach(button => {

button.addEventListener(
"click",
() => {

openPage(
button.dataset.page
);

}
);

});


document
.querySelectorAll("[data-go-page]")
.forEach(button => {

button.addEventListener(
"click",
() => {

openPage(
button.dataset.goPage
);

}
);

});

}


function openPage(page) {

document
.querySelectorAll(".page")
.forEach(section => {

section.classList.remove(
"active"
);

});


document
.querySelectorAll(
"#navTabs button"
)
.forEach(button => {

button.classList.remove(
"active"
);

});


const target =
$(`page-${page}`);

if (target) {
target.classList.add("active");
}


const button =
document.querySelector(
`#navTabs button[data-page="${page}"]`
);

if (button) {
button.classList.add("active");
}


window.scrollTo({
top: 0,
behavior: "smooth"
});

}


/* =========================================================
LEER LIBRO COMPLETO
========================================================= */

function parseWorkbook() {

const wb = state.workbook;

resetProjectData();


/* =========================
PLAN OI
========================= */

const planSheet =
findExactSheet(
wb,
"Plan OI"
);

if (planSheet) {

state.planOI =
parsePlanOI(
planSheet
);

} else {

console.warn(
"No se encontró la hoja Plan OI"
);

}


/* =========================
BD PLAN DETALLADO
========================= */

const bdSheet =
findExactSheet(
wb,
"BD Plan Detallado"
);

if (bdSheet) {

state.budget =
parseBDPlanDetallado(
bdSheet
);

} else {

console.warn(
"No se encontró BD Plan Detallado"
);

}


/* =========================
FLUJO DE CAJA
========================= */

const flujoSheet =
findExactSheet(
wb,
"Flujo de Caja"
);

if (flujoSheet) {

state.cashflow =
parseFlujoCaja(
flujoSheet
);

} else {

console.warn(
"No se encontró Flujo de Caja"
);

}


/* =========================
HOJAS OPCIONALES
========================= */

const issueSheet =
findSheet(
wb,
[
"Puntos pendientes",
"Pendientes",
"Seguimiento"
]
);

if (issueSheet) {

state.issues =
genericSheetToObjects(
issueSheet
);

}


const considerationsSheet =
findSheet(
wb,
[
"Consideraciones",
"Notas"
]
);

if (considerationsSheet) {

state.considerations =
genericSheetToObjects(
considerationsSheet
);

}


calculateTotals();


console.log(
"PLAN OI FINAL:",
state.planOI
);

console.log(
"BD FINAL:",
state.budget
);

console.log(
"FLUJO FINAL:",
state.cashflow
);

console.log(
"TOTALES:",
{
approved:
state.approvedBudget,

real:
state.currentBudget,

savings:
state.savings,

ordered:
state.orderedValue,

pending:
state.pendingValue
}
);

}


/* =========================================================
PLAN OI
========================================================= */

function parsePlanOI(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: "",
raw: true
}
);


const result = {
approved: 0,
real: 0,
difference: 0,
categories: []
};


/*
Primero localizamos la fila:
TOTAL GENERAL ESTIMADO

En tu plantilla:
A/B = descripción
Presupuestado
Real
Diferencia
*/

for (
let r = 0;
r < rows.length;
r++
) {

const row = rows[r];

let totalColumn = -1;


for (
let c = 0;
c < row.length;
c++
) {

if (
normalizeText(
row[c]
).includes(
"totalgeneralestimado"
)
) {

totalColumn = c;
break;

}

}


if (totalColumn === -1) {
continue;
}


const numbers = [];


for (
let c =
totalColumn + 1;

c < row.length;

c++
) {

const value =
toNumber(
row[c]
);


if (
value !== 0
) {

numbers.push(value);

}

}


if (
numbers.length >= 2
) {

result.approved =
numbers[0];

result.real =
numbers[1];

result.difference =
numbers.length >= 3
? numbers[2]
: numbers[0] -
numbers[1];

}

break;

}


/*
Ahora buscamos encabezados:
Resumen de Partidas
Presupuestado
Real
Diferencia
*/

let headerRow = -1;

let descriptionCol = -1;
let approvedCol = -1;
let realCol = -1;
let differenceCol = -1;


for (
let r = 0;
r < Math.min(
rows.length,
40
);
r++
) {

const row = rows[r];


for (
let c = 0;
c < row.length;
c++
) {

const text =
normalizeText(
row[c]
);


if (
text.includes(
"resumendepartidas"
)
) {

descriptionCol = c;

}


if (
text ===
"presupuestado"
) {

approvedCol = c;

}


if (
text ===
"real"
) {

realCol = c;

}


if (
text.includes(
"diferencia"
)
) {

differenceCol = c;

}

}


if (
descriptionCol >= 0 &&
approvedCol >= 0 &&
realCol >= 0
) {

headerRow = r;
break;

}

}


if (
headerRow >= 0
) {

for (
let r =
headerRow + 1;

r < rows.length;

r++
) {

const row =
rows[r];


const description =
String(
row[
descriptionCol
] || ""
).trim();


if (!description) {
continue;
}


const normalized =
normalizeText(
description
);


/*
No volvemos a agregar
TOTAL GENERAL ESTIMADO
como categoría.
*/

if (
normalized.includes(
"totalgeneralestimado"
)
) {

continue;

}


const approved =
approvedCol >= 0
? toNumber(
row[
approvedCol
]
)
: 0;


const real =
realCol >= 0
? toNumber(
row[
realCol
]
)
: 0;


const difference =
differenceCol >= 0
? toNumber(
row[
differenceCol
]
)
: approved - real;


if (
approved !== 0 ||
real !== 0 ||
difference !== 0
) {

result.categories.push({

partida:
description,

approved,

current:
real,

savings:
difference

});

}

}

}


/*
Si Diferencia no fue encontrada
pero sí tenemos ambos totales.
*/

if (
!result.difference &&
(
result.approved ||
result.real
)
) {

result.difference =
result.approved -
result.real;

}


return result;

}


/* =========================================================
BD PLAN DETALLADO
========================================================= */

function parseBDPlanDetallado(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: "",
raw: true
}
);


if (!rows.length) {
return [];
}


/*
Detectamos la fila real de encabezados.
En tu archivo actualmente es fila 4,
pero lo dejamos robusto para otros
proyectos.
*/

let headerIndex = -1;


for (
let r = 0;
r < Math.min(
rows.length,
15
);
r++
) {

const normalized =
rows[r].map(
cell =>
normalizeText(
cell
)
);


const hasPartidas =
normalized.some(
x =>
x === "partidas" ||
x === "partida"
);


const hasItem =
normalized.some(
x =>
x === "item"
);


const hasPO =
normalized.some(
x =>
x.includes(
"ordendecompra"
)
);


if (
hasPartidas &&
hasItem &&
hasPO
) {

headerIndex = r;
break;

}

}


if (
headerIndex === -1
) {

console.warn(
"No se detectaron encabezados en BD Plan Detallado."
);

return [];

}


const headers =
makeUniqueHeaders(
rows[
headerIndex
]
);


const result = [];


for (
let r =
headerIndex + 1;

r < rows.length;

r++
) {

const row =
rows[r];


if (
row.every(
cell =>
String(
cell ?? ""
).trim() ===
""
)
) {

continue;

}


const obj = {};


headers.forEach(
(header, index) => {

obj[header] =
row[index] ?? "";

}
);


const partida =
getColumnValue(
obj,
[
"Partidas",
"Partida"
]
);


const capexType =
getColumnValue(
obj,
[
"OI / Capex Tipo Ppto",
"OI Capex Tipo Ppto",
"Capex Tipo Ppto"
]
);


const plan =
getColumnValue(
obj,
[
"Plan"
]
);


const item =
getColumnValue(
obj,
[
"Item",
"Ítem"
]
);


const originalPO =
getColumnValue(
obj,
[
"Orden de compra",
"Orden de Compra",
"OC"
]
);


const date =
getColumnValue(
obj,
[
"Fecha",
"Fecha OC"
]
);


const quantity =
toNumber(
getColumnValue(
obj,
[
"Cant.",
"Cant",
"Cantidad"
]
)
);


const currency =
String(
getColumnValue(
obj,
[
"Mon",
"Moneda"
]
) ||
"USD"
).trim();


const rate =
toNumber(
getColumnValue(
obj,
[
"Tasa"
]
)
);


const unitPrice =
toNumber(
getColumnValue(
obj,
[
"Precio Und",
"Precio Unitario",
"Precio"
]
)
);


const netValueDOP =
toNumber(
getColumnValue(
obj,
[
"Valor Neto [DOP]",
"Valor Neto DOP",
"Valor Neto"
]
)
);


const liquidation =
toNumber(
getColumnValue(
obj,
[
"Liquid. [%]",
"Liquid.",
"Liquidación",
"Liquidacion"
]
)
);


const budgetDOP =
toNumber(
getColumnValue(
obj,
[
"Ppto. Orden Interna [DOP]",
"Ppto Orden Interna DOP",
"Presupuesto Orden Interna DOP"
]
)
);


const budgetUSD =
toNumber(
getColumnValue(
obj,
[
"Ppto. Orden Interna [USD]",
"Ppto Orden Interna USD",
"Presupuesto Orden Interna USD"
]
)
);


const comment =
String(
getColumnValue(
obj,
[
"Comentario",
"Comentarios",
"Observación",
"Observacion"
]
) || ""
).trim();


const supplier =
String(
getColumnValue(
obj,
[
"Proveedor"
]
) || ""
).trim();


const totalCashflow =
toNumber(
getColumnValue(
obj,
[
"Total Flujo Caja",
"Total Flujo de Caja"
]
)
);


const observation =
String(
getColumnValue(
obj,
[
"Observación",
"Observacion"
]
) || ""
).trim();


const validPO =
isValidPurchaseOrder(
originalPO
);


const payments = [];


/*
Pago 1 a Pago 5 según la
estructura actual.
*/

for (
let p = 1;
p <= 5;
p++
) {

const period =
toNumber(
getColumnValue(
obj,
[
`Per.P${p}`,
`Per P${p}`
]
)
);


const advance =
toNumber(
getColumnValue(
obj,
[
`Av.P${p} [%]`,
`Av.P${p}`,
`Av P${p}`
]
)
);


const amount =
toNumber(
getColumnValue(
obj,
[
`Pago ${p}`,
`Pago${p}`
]
)
);


if (
period > 0 ||
amount !== 0
) {

payments.push({

number:
p,

period,

date:
periodToDate(
period
),

advance,

amount

});

}

}


/*
Ignorar filas realmente vacías.
*/

if (
!partida &&
!item &&
!netValueDOP &&
!budgetDOP &&
!budgetUSD &&
!originalPO
) {

continue;

}


result.push({

raw:
obj,

partida:
String(
partida ||
"Sin partida"
).trim(),

capexType:
String(
capexType ||
""
).trim(),

plan,

item:
String(
item ||
"Sin descripción"
).trim(),

originalPO,

po:
validPO
? String(
originalPO
).trim()
: "",

hasPO:
validPO,

date,

quantity,

currency,

rate,

unitPrice,

netValue:
netValueDOP,

netValueDOP,

liquidation,

internalBudgetDOP:
budgetDOP,

internalBudgetUSD:
budgetUSD,

supplier,

comment,

observation,

payments,

totalCashflow,

/*
Para las vistas de compras
trabajamos en DOP.
*/

current:
budgetDOP,

ordered:
validPO
? netValueDOP
: 0,

status:
validPO
? "Con OC"
: supplier
? "Pendiente de compra"
: "Sin proveedor"

});

}


return result;

}


/* =========================================================
VALIDAR ORDEN DE COMPRA

"Plan" NO es una OC.
"Pendiente de compra" NO es una OC.
========================================================= */

function isValidPurchaseOrder(value) {

if (
value === null ||
value === undefined
) {

return false;

}


const text =
String(value)
.trim();


if (!text) {
return false;
}


const normalized =
normalizeText(text);


const invalid = [
"plan",
"pendiente",
"pendientedecompra",
"sinoc",
"noaplica",
"na",
"-"
];


if (
invalid.includes(
normalized
)
) {

return false;

}


/*
Las órdenes reales que vimos
tienen varios dígitos.
*/

const digits =
text.replace(
/\D/g,
""
);


return (
digits.length >= 6
);

}


/* =========================================================
PERÍODOS

Per.P = 1 -> Enero 2025
Per.P = 2 -> Febrero 2025
...
Per.P = 12 -> Diciembre 2025
Per.P = 13 -> Enero 2026
========================================================= */

function periodToDate(period) {

const p =
Number(period);


if (
!Number.isFinite(p) ||
p <= 0
) {

return null;

}


return new Date(
2025,
p - 1,
1
);

}


/* =========================================================
FLUJO DE CAJA
========================================================= */

function parseFlujoCaja(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: "",
raw: true
}
);


if (!rows.length) {
return [];
}


/*
Buscamos fila P1, P2, P3...
*/

let periodRow = -1;


for (
let r = 0;
r < Math.min(
rows.length,
30
);
r++
) {

const matches =
rows[r].filter(
cell =>
/^p\d+$/i.test(
String(
cell || ""
).trim()
)
).length;


if (
matches >= 3
) {

periodRow = r;
break;

}

}


if (
periodRow === -1
) {

console.warn(
"No se encontró P1/P2/P3 en Flujo de Caja."
);

return [];

}


const periods = [];


rows[
periodRow
].forEach(
(cell, column) => {

const text =
String(
cell || ""
).trim();


const match =
text.match(
/^P(\d+)$/i
);


if (!match) {
return;
}


const period =
Number(
match[1]
);


periods.push({

column,

period,

date:
periodToDate(
period
)

});

}
);


if (!periods.length) {
return [];
}


const firstPeriodColumn =
Math.min(
...periods.map(
item =>
item.column
)
);


const result = [];


for (
let r =
periodRow + 1;

r < rows.length;

r++
) {

const row =
rows[r];


/*
Encontramos la descripción
más cercana hacia la izquierda
de P1.
*/

let description = "";


for (
let c =
firstPeriodColumn - 1;

c >= 0;

c--
) {

const candidate =
String(
row[c] || ""
).trim();


if (candidate) {

description =
candidate;

break;

}

}


periods.forEach(
periodInfo => {

const amount =
toNumber(
row[
periodInfo.column
]
);


if (
amount === 0
) {

return;

}


result.push({

partida:
description ||
"Sin partida",

period:
periodInfo.period,

date:
periodInfo.date,

amount,

/*
La hoja de flujo que vimos
trabaja con los valores
consolidados del proyecto.
*/

dop:
amount

});

}
);

}


return result;

}


/* =========================================================
TOTALES GENERALES
========================================================= */

function calculateTotals() {

/*
PLAN OI:
TOTAL GENERAL ESTIMADO

Presupuestado
Real
Diferencia

Estos KPI permanecen en USD
porque así está Plan OI.
*/

state.approvedBudget =
toNumber(
state.planOI.approved
);


state.currentBudget =
toNumber(
state.planOI.real
);


state.savings =
toNumber(
state.planOI.difference
);


if (
!state.savings &&
(
state.approvedBudget ||
state.currentBudget
)
) {

state.savings =
state.approvedBudget -
state.currentBudget;

}


/*
COMPRAS:
Valor Neto [DOP]
únicamente cuando hay OC válida.
*/

state.orderedValue =
sum(
state.budget
.filter(
row =>
row.hasPO
)
.map(
row =>
row.netValueDOP
)
);


/*
Pendiente de compra en DOP:
presupuesto de OI menos OC colocadas.
*/

const detailedBudgetDOP =
sum(
state.budget.map(
row =>
row.internalBudgetDOP
)
);


state.pendingValue =
Math.max(
detailedBudgetDOP -
state.orderedValue,
0
);

}


/* =========================================================
HOJAS
========================================================= */

function findExactSheet(
workbook,
requestedName
) {

const found =
workbook.SheetNames.find(
name =>
normalizeText(name) ===
normalizeText(
requestedName
)
);


return found
? workbook.Sheets[
found
]
: null;

}


function findSheet(
workbook,
candidates
) {

for (
const candidate
of candidates
) {

const exact =
workbook.SheetNames.find(
name =>
normalizeText(name) ===
normalizeText(
candidate
)
);


if (exact) {

return workbook.Sheets[
exact
];

}

}


for (
const candidate
of candidates
) {

const partial =
workbook.SheetNames.find(
name =>
normalizeText(name)
.includes(
normalizeText(
candidate
)
)
);


if (partial) {

return workbook.Sheets[
partial
];

}

}


return null;

}


/* =========================================================
LECTOR GENÉRICO
========================================================= */

function genericSheetToObjects(sheet) {

return XLSX.utils.sheet_to_json(
sheet,
{
defval: "",
raw: false
}
);

}


/* =========================================================
HEADERS ÚNICOS
========================================================= */

function makeUniqueHeaders(row) {

const used = {};


return row.map(
(value, index) => {

let name =
String(
value || ""
).trim();


if (!name) {

name =
`Columna_${index + 1}`;

}


const original =
name;


if (
used[original]
) {

used[original]++;

name =
`${original}_${used[original]}`;

} else {

used[original] = 1;

}


return name;

}
);

}


/* =========================================================
BUSCAR VALOR DE COLUMNA
========================================================= */

function getColumnValue(
object,
candidates
) {

const keys =
Object.keys(
object
);


for (
const candidate
of candidates
) {

const target =
normalizeText(
candidate
);


const exact =
keys.find(
key =>
normalizeText(
key
) ===
target
);


if (
exact !== undefined
) {

const value =
object[exact];


if (
value !== undefined &&
value !== null &&
String(value)
.trim() !==
""
) {

return value;

}

}

}


return "";

}


/* =========================================================
RENDER GENERAL
========================================================= */

function renderAll() {

renderOverview();
renderBudget();
renderCashflow();
renderTracking();
renderPending();
renderLiquidation();

}


/* =========================================================
RESUMEN EJECUTIVO
========================================================= */

function renderOverview() {

/*
PLAN OI está en USD.
*/

const approved =
state.approvedBudget;

const current =
state.currentBudget;

const saving =
state.savings;


const savingPct =
approved
? saving / approved
: 0;


/*
Compras detalladas están en DOP.
*/

const orderedDOP =
state.orderedValue;

const pendingDOP =
state.pendingValue;


/*
Para avance usamos el presupuesto
detallado DOP contra OC DOP.
*/

const detailedBudgetDOP =
sum(
state.budget.map(
row =>
row.internalBudgetDOP
)
);


const progress =
detailedBudgetDOP
? orderedDOP /
detailedBudgetDOP
: 0;


setText(
"kpiApprovedBudget",
usd(approved)
);


setText(
"kpiBudget",
usd(current)
);


setText(
"kpiOrdered",
money(orderedDOP)
);


setText(
"kpiPendingBuy",
money(pendingDOP)
);


setText(
"kpiWeightedProgress",
percent(progress)
);


setText(
"kpiOpenIssues",
getOpenIssuesCount()
);


setText(
"overviewDiscount",
usd(saving)
);


setText(
"overviewDiscountPct",
percent(
savingPct
)
);


renderHealth(
progress,
pendingDOP
);

renderBudgetMix();
renderBudgetEvolution();
renderSavingsByPart();
renderPurchaseStatus();
renderTopSuppliers();
renderDataQuality();
renderAttentionList();

}


/* =========================================================
HEALTH
========================================================= */

function renderHealth(
progress,
pending
) {

const container =
$("healthStrip");

if (!container) return;


const withPO =
state.budget.filter(
row =>
row.hasPO
).length;


const total =
state.budget.length;


const coverage =
total
? withPO / total
: 0;


const noSupplier =
state.budget.filter(
row =>
!row.supplier
).length;


const items = [

{
title:
"Avance de compras",

value:
percent(progress),

text:
"OC colocadas vs presupuesto detallado",

color:
progress >= .75
? "good"
: "warn"
},

{
title:
"Cobertura de OC",

value:
percent(coverage),

text:
`${withPO} de ${total} registros con OC`,

color:
coverage >= .75
? "good"
: "warn"
},

{
title:
"Pendiente de compra",

value:
money(pending),

text:
"Presupuesto DOP todavía no colocado",

color:
pending > 0
? "warn"
: "good"
},

{
title:
"Sin proveedor",

value:
String(
noSupplier
),

text:
"Registros sin proveedor",

color:
noSupplier
? "bad"
: "good"
}

];


container.innerHTML =
items.map(
item => `

<div class="health-item">

<div class="row">

<b>
${escapeHtml(
item.title
)}
</b>

<i class="dot ${item.color}">
</i>

</div>

<strong>
${escapeHtml(
item.value
)}
</strong>

<p>
${escapeHtml(
item.text
)}
</p>

</div>

`
).join("");

}


/* =========================================================
GRÁFICA PLAN OI
========================================================= */

function renderBudgetMix() {

const categories =
state.planOI.categories
.filter(
row =>
row.current !== 0
)
.slice(0, 15);


createChart(
"budgetMixChart",
"doughnut",

categories.map(
row =>
row.partida
),

categories.map(
row =>
row.current
),

{
plugins: {
legend: {
position:
"bottom"
}
}
}
);

}


/* =========================================================
PRESUPUESTADO VS REAL
========================================================= */

function renderBudgetEvolution() {

const categories =
state.planOI.categories
.filter(
row =>
row.approved !== 0 ||
row.current !== 0
)
.slice(0, 15);


createMultiChart(
"budgetEvolutionChart",
"bar",

categories.map(
row =>
row.partida
),

[

{
label:
"Presupuestado USD",

data:
categories.map(
row =>
row.approved
),

backgroundColor:
COLORS.blue
},

{
label:
"Real USD",

data:
categories.map(
row =>
row.current
),

backgroundColor:
COLORS.cyan
}

]

);

}


/* =========================================================
DIFERENCIA / AHORRO
========================================================= */

function renderSavingsByPart() {

const categories =
state.planOI.categories
.filter(
row =>
row.savings !== 0
)
.sort(
(a, b) =>
Math.abs(
b.savings
) -
Math.abs(
a.savings
)
)
.slice(0, 15);


createChart(
"savingsByPartChart",
"bar",

categories.map(
row =>
row.partida
),

categories.map(
row =>
row.savings
),

{
indexAxis:
"y",

plugins: {
legend: {
display:
false
}
}
}
);

}


/* =========================================================
ESTADO DE COMPRAS
========================================================= */

function renderPurchaseStatus() {

const withPO =
state.budget.filter(
row =>
row.hasPO
).length;


const pending =
state.budget.filter(
row =>
!row.hasPO
).length;


createChart(
"purchaseStatusChart",
"doughnut",

[
"Con OC",
"Pendiente"
],

[
withPO,
pending
]

);

}


/* =========================================================
TOP PROVEEDORES
========================================================= */

function renderTopSuppliers() {

const container =
$("topSuppliers");

if (!container) return;


const groups = {};


state.budget.forEach(
row => {

if (
!row.supplier ||
!row.netValueDOP
) {

return;

}


groups[
row.supplier
] =
(
groups[
row.supplier
] || 0
) +
row.netValueDOP;

}
);


const entries =
Object.entries(
groups
)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(
0,
6
);


container.innerHTML =
entries.length
? entries.map(
(
[supplier, value],
index
) => `

<div class="rank-row">

<span class="n">
${index + 1}
</span>

<b>
${escapeHtml(
supplier
)}
</b>

<span>
${money(
value
)}
</span>

</div>

`
).join("")
: emptyMessage(
"No hay proveedores disponibles."
);

}


/* =========================================================
CALIDAD DE DATOS
========================================================= */

function renderDataQuality() {

const container =
$("dataQuality");

if (!container) return;


const total =
state.budget.length ||
1;


const missingSupplier =
state.budget.filter(
row =>
!row.supplier
).length;


const missingPO =
state.budget.filter(
row =>
!row.hasPO
).length;


const missingDate =
state.budget.filter(
row =>
row.hasPO &&
!row.date
).length;


const items = [

{
label:
"Sin proveedor",
value:
missingSupplier
},

{
label:
"Sin OC",
value:
missingPO
},

{
label:
"OC sin fecha",
value:
missingDate
}

];


container.innerHTML =
items.map(
item => `

<div class="quality">

<div class="top">

<b>
${item.label}
</b>

<strong>
${item.value}
</strong>

</div>

<small>

${percent(
item.value /
total
)}

de los registros

</small>

</div>

`
).join("");

}


/* =========================================================
ATENCIÓN
========================================================= */

function renderAttentionList() {

const container =
$("attentionList");

if (!container) return;


const pending =
[...state.budget]
.filter(
row =>
!row.hasPO &&
row.internalBudgetDOP > 0
)
.sort(
(a, b) =>
b.internalBudgetDOP -
a.internalBudgetDOP
)
.slice(
0,
6
);


container.innerHTML =
pending.length
? pending.map(
row => `

<div class="attention">

<i class="dot warn">
</i>

<div>

<b>
${escapeHtml(
row.item
)}
</b>

<span>

${escapeHtml(
row.partida
)}

·

${money(
row.internalBudgetDOP
)}

</span>

</div>

</div>

`
).join("")
: emptyMessage(
"No se identifican compras pendientes."
);

}


/* =========================================================
FILTROS PRESUPUESTO
========================================================= */

function setupBudgetFilters() {

[
"budgetCategoryFilter",
"budgetStatusFilter",
"budgetSupplierFilter"
].forEach(id => {

const element =
$(id);

if (element) {

element.addEventListener(
"change",
renderBudgetTable
);

}

});


const reset =
$("budgetReset");


if (reset) {

reset.addEventListener(
"click",
() => {

if (
$("budgetCategoryFilter")
) {

$("budgetCategoryFilter").value =
"";

}


if (
$("budgetStatusFilter")
) {

$("budgetStatusFilter").value =
"";

}


if (
$("budgetSupplierFilter")
) {

$("budgetSupplierFilter").value =
"";

}


renderBudgetTable();

}
);

}

}


/* =========================================================
PRESUPUESTO Y COMPRAS
========================================================= */

function renderBudget() {

populateBudgetFilters();

renderBudgetTable();


const category =
groupSum(
state.budget,
"partida",
"internalBudgetDOP"
);


createChart(
"categoryChart",
"bar",

category.labels,
category.values,

{
indexAxis:
"y",

plugins: {
legend: {
display:
false
}
}
}
);


const supplier =
groupSupplierValue();


createChart(
"supplierChart",
"bar",

supplier.labels,
supplier.values,

{
plugins: {
legend: {
display:
false
}
}
}
);

}


function populateBudgetFilters() {

fillSelect(
"budgetCategoryFilter",

unique(
state.budget.map(
row =>
row.partida
)
)
);


fillSelect(
"budgetSupplierFilter",

unique(
state.budget
.map(
row =>
row.supplier
)
.filter(Boolean)
)
);

}


function fillSelect(
id,
values
) {

const select =
$(id);

if (!select) return;


const current =
select.value;


const first =
select.options[0]
? select.options[0]
.outerHTML
: '<option value="">Todos</option>';


select.innerHTML =
first;


values
.sort(
(a, b) =>
String(a)
.localeCompare(
String(b),
"es"
)
)
.forEach(
value => {

const option =
document.createElement(
"option"
);

option.value =
value;

option.textContent =
value;

select.appendChild(
option
);

}
);


if (
[...select.options]
.some(
option =>
option.value ===
current
)
) {

select.value =
current;

}

}


/* =========================================================
TABLA PRESUPUESTO
========================================================= */

function renderBudgetTable() {

const category =
$("budgetCategoryFilter")
?.value || "";


const status =
$("budgetStatusFilter")
?.value || "";


const supplier =
$("budgetSupplierFilter")
?.value || "";


let rows =
[...state.budget];


if (category) {

rows =
rows.filter(
row =>
row.partida ===
category
);

}


if (supplier) {

rows =
rows.filter(
row =>
row.supplier ===
supplier
);

}


if (status) {

if (
normalizeText(status) ===
"sinocdefinida"
) {

rows =
rows.filter(
row =>
!row.hasPO
);

} else {

rows =
rows.filter(
row =>
normalizeText(
row.status
) ===
normalizeText(
status
)
);

}

}


const budget =
sum(
rows.map(
row =>
row.internalBudgetDOP
)
);


const ordered =
sum(
rows
.filter(
row =>
row.hasPO
)
.map(
row =>
row.netValueDOP
)
);


setText(
"budgetFilteredTotal",
money(budget)
);


setText(
"budgetFilteredOrdered",
money(ordered)
);


setText(
"budgetFilteredPending",
money(
Math.max(
budget -
ordered,
0
)
)
);


setText(
"budgetFilteredItems",
rows.length
);


setText(
"budgetTableCount",
`${rows.length} registros`
);


const tbody =
$("budgetTableBody");

if (!tbody) return;


tbody.innerHTML =
rows.map(
row => `

<tr>

<td>
${escapeHtml(
row.partida
)}
</td>

<td>
${escapeHtml(
row.item
)}
</td>

<td>
${escapeHtml(
row.supplier ||
"—"
)}
</td>

<td>

<span class="status-pill ${
row.hasPO
? "good"
: row.supplier
? "warn"
: "bad"
}">

${escapeHtml(
row.hasPO
? row.po
: row.status
)}

</span>

</td>

<td>
${escapeHtml(
formatPossibleDate(
row.date
)
)}
</td>

<td class="num money">
${money(
row.internalBudgetDOP
)}
</td>

<td>
${escapeHtml(
row.currency
)}
</td>

</tr>

`
).join("");

}


/* =========================================================
FLUJO DE CAJA
========================================================= */

function renderCashflow() {

/*
Fuente principal:
hoja Flujo de Caja.

Si por alguna razón esa hoja no
trae datos, usamos los pagos de
BD Plan Detallado como respaldo.
*/

let payments =
[...state.cashflow];


if (!payments.length) {

payments =
extractPaymentsFromBD();

}


state.cashflowPayments =
payments;


renderCashflowKPIs(
payments
);

renderPaymentReminders(
payments
);

renderCashflowCharts(
payments
);

renderPaymentTable(
payments
);

renderOverviewPaymentReminders();

}


/* =========================================================
PAGOS DESDE BD
========================================================= */

function extractPaymentsFromBD() {

const result = [];


state.budget.forEach(
row => {

row.payments.forEach(
payment => {

if (
!payment.amount
) {

return;

}


/*
En BD los pagos que vimos
corresponden a valores
monetarios de la plantilla.

Convertimos a DOP cuando
la moneda es USD.
*/

const dop =
convertToDOP(
payment.amount,
row.currency,
row.rate
);


result.push({

partida:
row.partida,

item:
row.item,

supplier:
row.supplier,

period:
payment.period,

date:
payment.date,

amount:
payment.amount,

dop,

currency:
row.currency,

paymentNumber:
payment.number

});

}
);

}
);


return result;

}


/* =========================================================
FLUJO KPIs
========================================================= */

function renderCashflowKPIs(payments) {

const now =
startOfToday();


const total =
sum(
payments.map(
payment =>
getPaymentDOP(
payment
)
)
);


const past =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date <
now
)
.map(
payment =>
getPaymentDOP(
payment
)
)
);


const monthly =
groupPaymentsByMonth(
payments
);


const entries =
Object.entries(
monthly
).sort(
(a, b) =>
a[1].date -
b[1].date
);


const peak =
[...entries]
.sort(
(a, b) =>
b[1].amount -
a[1].amount
)[0];


const currentMonth =
new Date(
now.getFullYear(),
now.getMonth(),
1
);


const next =
entries.find(
([, value]) =>
value.date >=
currentMonth
);


const ninety =
new Date(now);

ninety.setDate(
ninety.getDate() +
90
);


const next90 =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date >=
now &&
payment.date <=
ninety
)
.map(
payment =>
getPaymentDOP(
payment
)
)
);


setText(
"cfTotal",
money(total)
);


setText(
"cfPaidToDate",
money(past)
);


setText(
"cfPeakMonth",
peak
? peak[0]
: "—"
);


setText(
"cfPeakAmount",
peak
? money(
peak[1].amount
)
: "—"
);


setText(
"cfNextMonth",
next
? next[0]
: "—"
);


setText(
"cfNextAmount",
next
? money(
next[1].amount
)
: "—"
);


setText(
"cf90Days",
money(next90)
);


renderUpcomingMonths(
entries
);

}


/* =========================================================
RECORDATORIOS
========================================================= */

function renderPaymentReminders(payments) {

const container =
$("paymentReminderList");

if (!container) return;


const now =
startOfToday();


const future =
payments
.filter(
payment =>
payment.date
)
.map(
payment => ({

...payment,

days:
Math.ceil(
(
payment.date -
now
) /
86400000
)

})
)
.filter(
payment =>
payment.days >=
-30
)
.sort(
(a, b) =>
a.days -
b.days
)
.slice(
0,
8
);


container.innerHTML =
future.length
? future.map(
payment => {

let cls =
"upcoming";

let label;


if (
payment.days < 0
) {

cls =
"overdue";

label =
`${Math.abs(
payment.days
)} días vencido`;

} else if (
payment.days === 0
) {

cls =
"current";

label =
"Este mes";

} else if (
payment.days <= 30
) {

cls =
"current";

label =
`En ${payment.days} días`;

} else {

label =
formatDate(
payment.date
);

}


return `

<div class="payment-reminder ${cls}">

<span class="label">
${escapeHtml(
label
)}
</span>

<h3>
${escapeHtml(
payment.item ||
payment.partida ||
"Flujo programado"
)}
</h3>

<b>
${money(
getPaymentDOP(
payment
)
)}
</b>

<p>

${escapeHtml(
payment.supplier ||
payment.partida ||
""
)}

</p>

</div>

`;

}
).join("")
: emptyMessage(
"No hay pagos próximos identificados."
);

}


/* =========================================================
RECORDATORIOS RESUMEN
========================================================= */

function renderOverviewPaymentReminders() {

const container =
$("overviewPaymentReminders");

if (!container) return;


const payments =
state.cashflowPayments ||
[];


const now =
startOfToday();


const items =
payments
.filter(
payment =>
payment.date &&
payment.date >=
now
)
.sort(
(a, b) =>
a.date -
b.date
)
.slice(
0,
4
);


container.innerHTML =
items.length
? items.map(
payment => {

const days =
Math.ceil(
(
payment.date -
now
) /
86400000
);


return `

<div class="reminder-mini ${
days <= 30
? "current"
: ""
}">

<span>
${
days === 0
? "Este mes"
: `En ${days} días`
}
</span>

<b>
${money(
getPaymentDOP(
payment
)
)}
</b>

<small>
${escapeHtml(
payment.item ||
payment.partida ||
"Pago programado"
)}
</small>

</div>

`;

}
).join("")
: emptyMessage(
"No hay próximos pagos identificados."
);

}


/* =========================================================
GRÁFICAS FLUJO
========================================================= */

function renderCashflowCharts(payments) {

const monthly =
groupPaymentsByMonth(
payments
);


const entries =
Object.entries(
monthly
).sort(
(a, b) =>
a[1].date -
b[1].date
);


const labels =
entries.map(
entry =>
entry[0]
);


const values =
entries.map(
entry =>
entry[1].amount
);


createChart(
"cashflowChart",
"bar",
labels,
values,
{
plugins: {
legend: {
display:
false
}
}
}
);


const now =
startOfToday();


const past =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date <
now
)
.map(
payment =>
getPaymentDOP(
payment
)
)
);


const future =
sum(
payments
.filter(
payment =>
!payment.date ||
payment.date >=
now
)
.map(
payment =>
getPaymentDOP(
payment
)
)
);


createChart(
"cashflowStatusChart",
"doughnut",

[
"Períodos anteriores",
"Por venir"
],

[
past,
future
]

);


/*
Acumulado
*/

let cumulative = 0;


const cumulativeValues =
values.map(
value => {

cumulative +=
value;

return cumulative;

}
);


createChart(
"cashflowCumulativeChart",
"line",
labels,
cumulativeValues,
{
plugins: {
legend: {
display:
false
}
}
}
);


/*
Por año
*/

const years = {};


payments.forEach(
payment => {

if (!payment.date) {
return;
}


const year =
String(
payment.date
.getFullYear()
);


years[year] =
(
years[year] ||
0
) +
getPaymentDOP(
payment
);

}
);


createChart(
"cashflowCurrencyChart",
"doughnut",
Object.keys(years),
Object.values(years)
);

}


/* =========================================================
AGRUPAR FLUJO POR MES
========================================================= */

function groupPaymentsByMonth(payments) {

const groups = {};


payments.forEach(
payment => {

if (!payment.date) {
return;
}


const date =
new Date(
payment.date
.getFullYear(),

payment.date
.getMonth(),

1
);


const key =
`${date.getFullYear()}-${String(
date.getMonth() + 1
).padStart(2, "0")}`;


if (!groups[key]) {

groups[key] = {

amount:
0,

date,

label:
date.toLocaleDateString(
"es-DO",
{
month:
"short",

year:
"numeric"
}
)

};

}


groups[key].amount +=
getPaymentDOP(
payment
);

}
);


const output = {};


Object.values(groups)
.sort(
(a, b) =>
a.date -
b.date
)
.forEach(
value => {

output[
value.label
] = value;

}
);


return output;

}


/* =========================================================
PRÓXIMOS MESES
========================================================= */

function renderUpcomingMonths(entries) {

const container =
$("cashflowUpcomingMonths");

if (!container) return;


const now =
new Date();


const currentMonth =
new Date(
now.getFullYear(),
now.getMonth(),
1
);


const future =
entries
.filter(
([, value]) =>
value.date >=
currentMonth
)
.slice(
0,
6
);


const max =
Math.max(
...future.map(
([, value]) =>
value.amount
),
0
);


container.innerHTML =
future.length
? future.map(
(
[label, value],
index
) => `

<div class="upcoming-month ${
index === 0
? "next"
: value.amount >=
max * .8
? "high"
: ""
}">

<div>

<span>
${
index === 0
? "Próximo compromiso"
: "Mes programado"
}
</span>

<b>
${escapeHtml(
label
)}
</b>

</div>

<div class="upcoming-amount">

<b>
${money(
value.amount
)}
</b>

<small>
Flujo programado
</small>

</div>

</div>

`
).join("")
: emptyMessage(
"No hay meses futuros con desembolsos."
);

}


/* =========================================================
TABLA DE PAGOS
========================================================= */

function renderPaymentTable(payments) {

const tbody =
$("paymentTable");

if (!tbody) return;


setText(
"paymentTableCount",
`${payments.length} movimientos`
);


const now =
startOfToday();


tbody.innerHTML =
payments.map(
payment => {

let status =
"Programado";

let cls =
"good";


if (
payment.date &&
payment.date <
now
) {

status =
"Período anterior";

cls =
"info";

}


return `

<tr>

<td>
${escapeHtml(
payment.date
? monthYear(
payment.date
)
: `P${
payment.period ||
""
}`
)}
</td>

<td>
${escapeHtml(
payment.partida ||
"—"
)}
</td>

<td>
${escapeHtml(
payment.item ||
payment.partida ||
"—"
)}
</td>

<td>
${escapeHtml(
payment.supplier ||
"—"
)}
</td>

<td>
${
payment.paymentNumber
? `Pago ${payment.paymentNumber}`
: `P${payment.period || ""}`
}
</td>

<td>
${escapeHtml(
payment.currency ||
"DOP"
)}
</td>

<td class="num">
${money(
getPaymentDOP(
payment
)
)}
</td>

<td class="num money">
${money(
getPaymentDOP(
payment
)
)}
</td>

<td>

<span class="status-pill ${cls}">
${status}
</span>

</td>

</tr>

`;

}
).join("");

}


/* =========================================================
OBTENER MONTO DOP
========================================================= */

function getPaymentDOP(payment) {

if (
Number.isFinite(
Number(
payment.dop
)
)
) {

return Number(
payment.dop
);

}


return Number(
payment.amount
) || 0;

}


/* =========================================================
CONVERSIÓN DOP
========================================================= */

function convertToDOP(
amount,
currency,
rowRate
) {

const c =
normalizeText(
currency
);


if (
c.includes("usd") ||
c.includes("dolar")
) {

const rate =
Number(rowRate) ||
state.exchangeRate;

return (
Number(amount) ||
0
) * rate;

}


return Number(
amount
) || 0;

}


/* =========================================================
SELECTOR DE MONEDA
========================================================= */

function setupCashflowCurrency() {

const select =
$("cashflowCurrencyView");

if (!select) return;


select.addEventListener(
"change",
renderCashflow
);

}


/* =========================================================
SEGUIMIENTO
========================================================= */

function renderTracking() {

const withPO =
state.budget.filter(
row =>
row.hasPO
);


const pending =
state.budget.filter(
row =>
!row.hasPO
);


const noSupplier =
state.budget.filter(
row =>
!row.supplier
);


const noDate =
state.budget.filter(
row =>
row.hasPO &&
!row.date
);


setText(
"trackWithPO",
withPO.length
);


setText(
"trackPending",
pending.length
);


setText(
"trackNoSupplier",
noSupplier.length
);


setText(
"trackNoDate",
noDate.length
);


renderPipeline(
withPO.length,
pending.length,
noSupplier.length,
noDate.length
);


renderBottlenecks(
pending
);


renderSupplierPO(
withPO
);


renderTrackingTable(
pending
);

}


/* =========================================================
PIPELINE
========================================================= */

function renderPipeline(
withPO,
pending,
noSupplier,
noDate
) {

const container =
$("pipeline");

if (!container) return;


const total =
state.budget.length;


const stages = [

{
label:
"Necesidades",

value:
total,

detail:
"Registros identificados"
},

{
label:
"Proveedor definido",

value:
total -
noSupplier,

detail:
"Con proveedor"
},

{
label:
"Orden colocada",

value:
withPO,

detail:
"OC válida"
},

{
label:
"Pendientes",

value:
pending,

detail:
`${noDate} OC sin fecha`
}

];


container.innerHTML =
stages.map(
stage => `

<div class="stage">

<span>
${stage.label}
</span>

<b>
${stage.value}
</b>

<small>
${stage.detail}
</small>

</div>

`
).join("");

}


/* =========================================================
CUELLOS DE BOTELLA
========================================================= */

function renderBottlenecks(pending) {

const top =
[...pending]
.filter(
row =>
row.internalBudgetDOP > 0
)
.sort(
(a, b) =>
b.internalBudgetDOP -
a.internalBudgetDOP
)
.slice(
0,
10
);


createChart(
"bottleneckChart",
"bar",

top.map(
row =>
row.item
),

top.map(
row =>
row.internalBudgetDOP
),

{
indexAxis:
"y",

plugins: {
legend: {
display:
false
}
}
}
);

}


/* =========================================================
OC POR PROVEEDOR
========================================================= */

function renderSupplierPO(withPO) {

const groups = {};


withPO.forEach(
row => {

const supplier =
row.supplier ||
"Sin proveedor";


groups[supplier] =
(
groups[supplier] ||
0
) + 1;

}
);


const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(
0,
10
);


createChart(
"supplierPOChart",
"bar",

entries.map(
item =>
item[0]
),

entries.map(
item =>
item[1]
),

{
plugins: {
legend: {
display:
false
}
}
}
);

}


/* =========================================================
TABLA SEGUIMIENTO
========================================================= */

function renderTrackingTable(pending) {

const tbody =
$("trackingTable");

if (!tbody) return;


setText(
"trackingCount",
`${pending.length} registros`
);


tbody.innerHTML =
[...pending]
.sort(
(a, b) =>
b.internalBudgetDOP -
a.internalBudgetDOP
)
.map(
row => `

<tr>

<td>
${escapeHtml(
row.item
)}
</td>

<td>
${escapeHtml(
row.partida
)}
</td>

<td>
${escapeHtml(
row.supplier ||
"—"
)}
</td>

<td>

<span class="status-pill ${
row.supplier
? "warn"
: "bad"
}">

${escapeHtml(
row.status
)}

</span>

</td>

<td>
${escapeHtml(
row.comment ||
row.observation ||
"—"
)}
</td>

<td class="num money">
${money(
row.internalBudgetDOP
)}
</td>

</tr>

`
).join("");

}


/* =========================================================
PENDIENTES
========================================================= */

function getOpenIssuesCount() {

if (
state.issues.length
) {

return state.issues.length;

}


return state.budget.filter(
row =>
!row.hasPO &&
(
row.comment ||
row.observation
)
).length;

}


function renderPending() {

let issues =
normalizeIssues(
state.issues
);


if (!issues.length) {

issues =
state.budget
.filter(
row =>
!row.hasPO &&
(
row.comment ||
row.observation
)
)
.map(
row => ({

title:
row.item,

owner:
row.supplier ||
"Sin responsable",

comment:
row.comment ||
row.observation,

status:
row.status

})
);

}


const owners =
unique(
issues.map(
issue =>
issue.owner
)
);


const withComments =
issues.filter(
issue =>
issue.comment
);


setText(
"pendingTotal",
issues.length
);


setText(
"pendingOwners",
owners.length
);


setText(
"pendingComments",
withComments.length
);


setText(
"pendingNoComments",
issues.length -
withComments.length
);


renderOwnerChart(
issues
);


renderConsiderations();


renderIssueBoard(
issues
);

}


function normalizeIssues(rows) {

return rows.map(
(row, index) => ({

title:
valueFrom(
row,
[
"Pendiente",
"Tema",
"Issue",
"Actividad",
"Descripción",
"Descripcion"
]
) ||
`Pendiente ${index + 1}`,

owner:
valueFrom(
row,
[
"Responsable",
"Owner",
"Líder",
"Lider"
]
) ||
"Sin responsable",

comment:
valueFrom(
row,
[
"Comentario",
"Comentarios",
"Observación",
"Observacion"
]
) || "",

status:
valueFrom(
row,
[
"Estado",
"Status"
]
) ||
"Abierto"

})
);

}


/* =========================================================
RESPONSABLES
========================================================= */

function renderOwnerChart(issues) {

const groups = {};


issues.forEach(
issue => {

groups[
issue.owner
] =
(
groups[
issue.owner
] || 0
) + 1;

}
);


const entries =
Object.entries(
groups
)
.sort(
(a, b) =>
b[1] - a[1]
);


createChart(
"ownerChart",
"bar",

entries.map(
item =>
item[0]
),

entries.map(
item =>
item[1]
),

{
plugins: {
legend: {
display:
false
}
}
}
);

}


/* =========================================================
CONSIDERACIONES
========================================================= */

function renderConsiderations() {

const container =
$("considerations");

if (!container) return;


let items =
state.considerations
.map(
row =>
valueFrom(
row,
[
"Consideración",
"Consideracion",
"Nota",
"Descripción",
"Descripcion"
]
)
)
.filter(Boolean);


if (!items.length) {

items = [

"Validar fechas de entrega de partidas críticas.",

"Confirmar proveedores de compras pendientes.",

"Actualizar comentarios de los registros abiertos.",

"Revisar el flujo de caja contra los períodos P1, P2, P3 y siguientes.",

"Mantener Plan OI y BD Plan Detallado actualizados."

];

}


container.innerHTML =
items
.slice(
0,
8
)
.map(
item => `

<div class="consideration">

${escapeHtml(
item
)}

</div>

`
).join("");

}


/* =========================================================
TABLERO DE PENDIENTES
========================================================= */

function renderIssueBoard(issues) {

const container =
$("issueBoard");

if (!container) return;


container.innerHTML =
issues.length
? issues.map(
issue => `

<article class="issue-card">

<b>
${escapeHtml(
issue.title
)}
</b>

<div class="owner">
${escapeHtml(
issue.owner
)}
</div>

<p>
${escapeHtml(
issue.comment ||
"Sin comentario registrado."
)}
</p>

</article>

`
).join("")
: `

<section class="card">
No hay pendientes registrados.
</section>

`;

}


/* =========================================================
PRE-LIQUIDACIÓN
========================================================= */

function setupLiquidation() {

[
"liqAmount",
"liqExw",
"liqFob",
"liqContainers",
"liqRate"
].forEach(
id => {

const input =
$(id);


if (input) {

input.addEventListener(
"input",
renderLiquidation
);

}

}
);

}


function renderLiquidation() {

const container =
$("liquidationScenarios");

if (!container) return;


const amount =
numberInput(
"liqAmount",
10000
);


const exw =
numberInput(
"liqExw",
0
);


const fob =
numberInput(
"liqFob",
1000
);


const containers =
numberInput(
"liqContainers",
1
);


const rate =
numberInput(
"liqRate",
state.exchangeRate
);


state.exchangeRate =
rate;


const scenarios = [

{
country:
"China",

freight:
4200,

duty:
.20,

insurance:
.01,

other:
1200
},

{
country:
"Estados Unidos",

freight:
2200,

duty:
.20,

insurance:
.01,

other:
900
},

{
country:
"Europa",

freight:
3500,

duty:
.20,

insurance:
.01,

other:
1100
},

{
country:
"México",

freight:
2600,

duty:
.20,

insurance:
.01,

other:
950
},

{
country:
"Puerto Rico",

freight:
1600,

duty:
.20,

insurance:
.01,

other:
800
}

];


container.innerHTML =
scenarios.map(
scenario => {

const freight =
scenario.freight *
containers;


const base =
amount +
exw +
fob;


const insurance =
base *
scenario.insurance;


const cif =
base +
freight +
insurance;


const duty =
cif *
scenario.duty;


const taxable =
cif +
duty +
scenario.other;


const itbis =
taxable *
.18;


const total =
taxable +
itbis;


return `

<article class="scenario">

<h3>
${scenario.country}
</h3>

<div class="metric">

<span>
Compra
</span>

<b>
${usd(amount)}
</b>

</div>

<div class="metric">

<span>
Flete estimado
</span>

<b>
${usd(freight)}
</b>

</div>

<div class="metric">

<span>
Seguro
</span>

<b>
${usd(insurance)}
</b>

</div>

<div class="metric">

<span>
CIF
</span>

<b>
${usd(cif)}
</b>

</div>

<div class="metric">

<span>
Gravamen
</span>

<b>
${usd(duty)}
</b>

</div>

<div class="metric">

<span>
Otros gastos
</span>

<b>
${usd(
scenario.other
)}
</b>

</div>

<div class="metric">

<span>
ITBIS
</span>

<b>
${usd(itbis)}
</b>

</div>

<div class="total">

<span>
Costo estimado puesto RD
</span>

<b>
${usd(total)}
</b>

<small>

≈
${money(
total *
rate
)}

</small>

</div>

</article>

`;

}
).join("");

}


/* =========================================================
AGRUPACIONES
========================================================= */

function groupSupplierValue() {

const groups = {};


state.budget.forEach(
row => {

if (
!row.supplier ||
!row.netValueDOP
) {

return;

}


groups[
row.supplier
] =
(
groups[
row.supplier
] || 0
) +
row.netValueDOP;

}
);


const entries =
Object.entries(
groups
)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(
0,
12
);


return {

labels:
entries.map(
item =>
item[0]
),

values:
entries.map(
item =>
item[1]
)

};

}


function groupSum(
rows,
labelKey,
valueKey
) {

const groups = {};


rows.forEach(
row => {

const label =
row[labelKey] ||
"Sin clasificar";


const value =
Number(
row[valueKey]
) || 0;


groups[label] =
(
groups[label] ||
0
) +
value;

}
);


const entries =
Object.entries(
groups
)
.filter(
([, value]) =>
value !== 0
)
.sort(
(a, b) =>
b[1] - a[1]
);


return {

labels:
entries.map(
item =>
item[0]
),

values:
entries.map(
item =>
item[1]
)

};

}


/* =========================================================
CHART.JS
========================================================= */

function createChart(
id,
type,
labels,
values,
options = {}
) {

const canvas =
$(id);

if (!canvas) return;


destroyChart(id);


const background =
type ===
"doughnut"

? [
COLORS.blue,
COLORS.cyan,
COLORS.green,
COLORS.yellow,
COLORS.red,
"#4388c7",
"#7bc7ed",
"#8b7dbb",
"#f6b65b"
]

: COLORS.cyan;


state.charts[id] =
new Chart(
canvas.getContext(
"2d"
),
{

type,

data: {

labels,

datasets: [
{

data:
values,

backgroundColor:
background,

borderColor:
type ===
"line"
? COLORS.blue
: undefined,

tension:
.3,

fill:
false

}
]

},

options:
mergeChartOptions(
options
)

}
);

}


function createMultiChart(
id,
type,
labels,
datasets,
options = {}
) {

const canvas =
$(id);

if (!canvas) return;


destroyChart(id);


state.charts[id] =
new Chart(
canvas.getContext(
"2d"
),
{

type,

data: {
labels,
datasets
},

options:
mergeChartOptions(
options
)

}
);

}


function mergeChartOptions(custom) {

const base = {

responsive:
true,

maintainAspectRatio:
false,

interaction: {

intersect:
false,

mode:
"index"

},

plugins: {

legend: {

labels: {

boxWidth:
10,

boxHeight:
10,

font: {
size: 10
}

}

}

},

scales: {

x: {

ticks: {

font: {
size: 9
},

maxRotation:
45,

minRotation:
0

},

grid: {
display:
false
}

},

y: {

ticks: {

font: {
size: 9
}

},

grid: {
color:
"#edf2f6"
}

}

}

};


/*
Doughnut no necesita escalas.
*/

if (
custom &&
custom.scales ===
false
) {

delete base.scales;

}


if (
custom &&
custom.plugins
) {

base.plugins = {
...base.plugins,
...custom.plugins
};

}


return deepMerge(
base,
custom
);

}


function deepMerge(
target,
source
) {

const output = {
...target
};


Object.keys(
source || {}
).forEach(
key => {

if (
source[key] &&
typeof source[key] ===
"object" &&
!Array.isArray(
source[key]
)
) {

output[key] =
deepMerge(
target[key] ||
{},
source[key]
);

} else {

output[key] =
source[key];

}

}
);


return output;

}


function destroyChart(id) {

if (
state.charts[id]
) {

state.charts[id]
.destroy();

delete state.charts[id];

}

}


/* =========================================================
COUNTDOWN
========================================================= */

function startCountdown(dateString) {

if (
countdownInterval
) {

clearInterval(
countdownInterval
);

}


const target =
parseProjectDate(
dateString
);


function update() {

if (!target) return;


const now =
new Date();


let diff =
target -
now;


const status =
$("countdownStatus");


if (
diff <= 0
) {

diff = 0;


if (status) {

status.textContent =
"Fecha de apertura alcanzada";

}

} else if (status) {

status.textContent =
"Hacia la apertura";

}


const days =
Math.floor(
diff /
86400000
);


const hours =
Math.floor(
(
diff %
86400000
) /
3600000
);


const minutes =
Math.floor(
(
diff %
3600000
) /
60000
);


const seconds =
Math.floor(
(
diff %
60000
) /
1000
);


setText(
"cdDays",
days
);


setText(
"cdHours",
String(hours)
.padStart(
2,
"0"
)
);


setText(
"cdMinutes",
String(minutes)
.padStart(
2,
"0"
)
);


setText(
"cdSeconds",
String(seconds)
.padStart(
2,
"0"
)
);

}


update();


countdownInterval =
setInterval(
update,
1000
);

}


/* =========================================================
FECHAS
========================================================= */

function parseProjectDate(value) {

if (!value) {
return null;
}


const date =
new Date(
`${value}T00:00:00`
);


return isNaN(date)
? null
: date;

}


function parseExcelDate(value) {

if (!value) {
return null;
}


if (
value instanceof Date &&
!isNaN(value)
) {

return value;

}


if (
typeof value ===
"number"
) {

const parsed =
XLSX.SSF
.parse_date_code(
value
);


if (parsed) {

return new Date(
parsed.y,
parsed.m - 1,
parsed.d
);

}

}


const text =
String(value)
.trim();


if (!text) {
return null;
}


const match =
text.match(
/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/
);


if (match) {

let year =
Number(
match[3]
);


if (
year < 100
) {

year += 2000;

}


return new Date(
year,
Number(
match[2]
) - 1,
Number(
match[1]
)
);

}


const direct =
new Date(text);


return isNaN(direct)
? null
: direct;

}


function startOfToday() {

const now =
new Date();


return new Date(
now.getFullYear(),
now.getMonth(),
now.getDate()
);

}


function formatDate(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
day:
"2-digit",

month:
"short",

year:
"numeric"
}
);

}


function formatLongDate(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
day:
"numeric",

month:
"long",

year:
"numeric"
}
);

}


function formatPossibleDate(value) {

const date =
parseExcelDate(
value
);


return date
? formatDate(date)
: String(
value ||
"—"
);

}


function monthYear(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
month:
"short",

year:
"numeric"
}
);

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


if (
typeof value ===
"number"
) {

return Number.isFinite(
value
)
? value
: 0;

}


let text =
String(value)
.trim();


/*
Paréntesis contables:
(1,900.58)
= -1900.58
*/

let negative = false;


if (
text.startsWith("(") &&
text.endsWith(")")
) {

negative = true;

text =
text.slice(
1,
-1
);

}


text =
text
.replace(
/\s/g,
""
)
.replace(
/RD\$/gi,
""
)
.replace(
/US\$/gi,
""
)
.replace(
/USD/gi,
""
)
.replace(
/\$/g,
""
)
.replace(
/%/g,
"");


if (
text.includes(",") &&
text.includes(".")
) {

if (
text.lastIndexOf(",") >
text.lastIndexOf(".")
) {

text =
text
.replace(
/\./g,
""
)
.replace(
",",
"."
);

} else {

text =
text.replace(
/,/g,
""
);

}

} else if (
text.includes(",")
) {

const parts =
text.split(",");


if (
parts.length === 2 &&
parts[1].length <= 2
) {

text =
parts[0]
.replace(
/\./g,
""
) +
"." +
parts[1];

} else {

text =
text.replace(
/,/g,
""
);

}

}


text =
text.replace(
/[^0-9.-]/g,
""
);


let number =
Number(text);


if (
!Number.isFinite(
number
)
) {

return 0;

}


if (negative) {

number =
-Math.abs(
number
);

}


return number;

}


/* =========================================================
TEXTO
========================================================= */

function normalizeText(value) {

return String(
value || ""
)
.normalize("NFD")
.replace(
/[\u0300-\u036f]/g,
""
)
.toLowerCase()
.replace(
/[^a-z0-9]/g,
""
);

}


function valueFrom(
row,
candidates
) {

return getColumnValue(
row,
candidates
);

}


/* =========================================================
FORMATOS
========================================================= */

function money(value) {

const number =
Number(value) ||
0;


if (
Math.abs(number) >=
1000000000
) {

return (
"RD$ " +
(
number /
1000000000
).toFixed(2) +
" B"
);

}


if (
Math.abs(number) >=
1000000
) {

return (
"RD$ " +
(
number /
1000000
).toFixed(2) +
" MM"
);

}


return new Intl
.NumberFormat(
"es-DO",
{
style:
"currency",

currency:
"DOP",

maximumFractionDigits:
0
}
)
.format(number);

}


function usd(value) {

return new Intl
.NumberFormat(
"en-US",
{
style:
"currency",

currency:
"USD",

maximumFractionDigits:
2
}
)
.format(
Number(value) ||
0
);

}


function percent(value) {

return new Intl
.NumberFormat(
"es-DO",
{
style:
"percent",

maximumFractionDigits:
1
}
)
.format(
Number(value) ||
0
);

}


/* =========================================================
UTILIDADES
========================================================= */

function sum(values) {

return values.reduce(
(
total,
value
) =>
total +
(
Number(value) ||
0
),
0
);

}


function unique(values) {

return [
...new Set(
values.filter(
value =>
value !== null &&
value !== undefined &&
String(value)
.trim() !==
""
)
)
];

}


function setText(
id,
value
) {

const element =
$(id);


if (element) {

element.textContent =
value;

}

}


function numberInput(
id,
fallback
) {

const element =
$(id);


if (!element) {
return fallback;
}


const value =
Number(
element.value
);


return Number.isFinite(
value
)
? value
: fallback;

}


function showLoader(show) {

const loader =
$("loader");


if (!loader) {
return;
}


loader.classList.toggle(
"hidden",
!show
);

}


function escapeHtml(value) {

return String(
value === null ||
value === undefined
? ""
: value
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


function emptyMessage(text) {

return `

<div class="muted">

${escapeHtml(
text
)}

</div>

`;

}


/* =========================================================
ESTADO VACÍO
========================================================= */

function renderEmptyState() {

setText(
"overviewSubtitle",
"No fue posible cargar la plantilla de esta iniciativa."
);


[
"kpiApprovedBudget",
"kpiBudget",
"kpiOrdered",
"kpiPendingBuy",
"kpiWeightedProgress",
"kpiOpenIssues",
"overviewDiscount",
"overviewDiscountPct"
].forEach(
id =>
setText(
id,
"—"
)
);

}

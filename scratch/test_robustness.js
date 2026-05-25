// Mock context and data to test index.js logic locally
const assert = require('assert');

// 1. Test Read Function logic
function testReadLogic(headers, row) {
    let obj = { rowNumber: 2 };
    headers.forEach((headerName, headerIndex) => {
        let finalHeader = headerName;
        if (headerName === "Column1" || headerName === "Numero" || headerName === "numero") {
            finalHeader = "numero";
        }
        obj[finalHeader] = row[headerIndex];
    });
    return obj;
}

// 2. Test Write Function logic
function testWriteLogic(headers, updates, rowValues) {
    for (const [key, value] of Object.entries(updates)) {
        let colIndex = headers.indexOf(key);
        
        // Si no encuentra la columna exacta, prueba mapeos comunes (compatibilidad con Excel viejo y nuevo)
        if (colIndex === -1 && (key === "numero" || key === "Column1" || key === "Numero")) {
            const fallbacks = ["numero", "Column1", "Numero"];
            for (const fallback of fallbacks) {
                colIndex = headers.indexOf(fallback);
                if (colIndex !== -1) break;
            }
        }
        
        if (colIndex !== -1) {
            rowValues[colIndex] = value;
        } else {
            console.log(`WARN: Columna ${key} no encontrada en Excel.`);
        }
    }
    return rowValues;
}

console.log("--- PROBANDO READ LOGIC ---");
const headers1 = ["Column1", "vencimiento", "deben"];
const row1 = ["57318...", 45432, "No"];
const result1 = testReadLogic(headers1, row1);
console.log("Headers con 'Column1':", result1);
assert.strictEqual(result1.numero, "57318...");

const headers2 = ["numero", "vencimiento", "deben"];
const row2 = ["57318...", 45432, "No"];
const result2 = testReadLogic(headers2, row2);
console.log("Headers con 'numero':", result2);
assert.strictEqual(result2.numero, "57318...");

const headers3 = ["Numero", "vencimiento", "deben"];
const row3 = ["57318...", 45432, "No"];
const result3 = testReadLogic(headers3, row3);
console.log("Headers con 'Numero':", result3);
assert.strictEqual(result3.numero, "57318...");


console.log("\n--- PROBANDO WRITE LOGIC ---");
// Caso A: Excel tiene "numero", enviamos "numero"
let rowValuesA = ["", "", ""];
const updatedA = testWriteLogic(["numero", "vencimiento", "deben"], { numero: "57318..." }, rowValuesA);
console.log("Excel tiene 'numero', enviamos 'numero':", updatedA);
assert.strictEqual(updatedA[0], "57318...");

// Caso B: Excel tiene "Column1", enviamos "numero"
let rowValuesB = ["", "", ""];
const updatedB = testWriteLogic(["Column1", "vencimiento", "deben"], { numero: "57318..." }, rowValuesB);
console.log("Excel tiene 'Column1', enviamos 'numero':", updatedB);
assert.strictEqual(updatedB[0], "57318...");

// Caso C: Excel tiene "Numero", enviamos "numero"
let rowValuesC = ["", "", ""];
const updatedC = testWriteLogic(["Numero", "vencimiento", "deben"], { numero: "57318..." }, rowValuesC);
console.log("Excel tiene 'Numero', enviamos 'numero':", updatedC);
assert.strictEqual(updatedC[0], "57318...");

// Caso D: Excel tiene "numero", enviamos "Column1"
let rowValuesD = ["", "", ""];
const updatedD = testWriteLogic(["numero", "vencimiento", "deben"], { Column1: "57318..." }, rowValuesD);
console.log("Excel tiene 'numero', enviamos 'Column1':", updatedD);
assert.strictEqual(updatedD[0], "57318...");

console.log("\n¡Todas las pruebas pasaron con éxito! El mapeo dinámico funciona perfectamente en todos los escenarios.");

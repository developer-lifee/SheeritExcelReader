require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');

// Helper to convert JS Object date back to Excel serial date if needed
function jsDateToExcelDate(dateString) {
    const jsDate = new Date(dateString);
    if (isNaN(jsDate.getTime())) return dateString; // Just return original if invalid
    // Excel date bug: 1900 leap year
    let returnDateTime = 25569.0 + ((jsDate.getTime() - (jsDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
    return returnDateTime;
}

module.exports = async function (context, req) {
    context.log('HTTP trigger function to Write Excel via MS Graph processed a request.');

    try {
        const clientId = process.env.AZURE_CLIENT_ID;
        const tenantId = 'common';
        const refreshToken = process.env.GRAPH_REFRESH_TOKEN;

        if (!clientId || !refreshToken) {
            context.res = { status: 400, body: "Faltan las variables de entorno AZURE_CLIENT_ID o GRAPH_REFRESH_TOKEN." };
            return;
        }

        const reqBody = req.body || {};
        const { rowNumber, updates } = reqBody;

        if (!rowNumber || !updates || typeof updates !== 'object') {
            context.res = { status: 400, body: "Debe enviar 'rowNumber' y un objeto 'updates'." };
            return;
        }

        // 1. Obtener Access Token
        const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', clientId);
        bodyParams.append('grant_type', 'refresh_token');
        bodyParams.append('refresh_token', refreshToken);
        bodyParams.append('scope', 'Files.ReadWrite.All offline_access');
        
        const tokenResponse = await fetch(tokenEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: bodyParams
        });

        const tokenData = await tokenResponse.json();
        if (!tokenResponse.ok) throw new Error("No se pudo obtener access_token usando el refresh token.");
        const accessToken = tokenData.access_token;

        const authProvider = { getAccessToken: async () => accessToken };
        const graphClient = Client.initWithMiddleware({ authProvider });

        const filePath = "Documentos/neflis_negro.xlsx"; 
        const worksheetName = "Hoja1";

        // 2. Obtener headers para saber el mapeo de columnas (A1:Z1)
        const headersEndpoint = `/me/drive/root:/${filePath}:/workbook/worksheets('${worksheetName}')/range(address='A1:Z1')`;
        const headersRes = await graphClient.api(headersEndpoint).get();
        const headers = headersRes.values[0] || [];

        // 3. Obtener la fila actual para no sobreescribir otros datos
        const rowEndpoint = `/me/drive/root:/${filePath}:/workbook/worksheets('${worksheetName}')/range(address='A${rowNumber}:Z${rowNumber}')`;
        const rowRes = await graphClient.api(rowEndpoint).get();
        let rowValues = rowRes.values[0] || new Array(26).fill("");

        // 4. Modificar los valores de la fila según los 'updates' mandados
        // updates = { "deben": "No", "observaciones": "Pagó" }
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
                // If it's the vencimiento date, maybe convert it back or just write a string 
                // Graph can accept a simple string 'yyyy-mm-dd' and usually parses it properly as a date if Excel is configured to
                rowValues[colIndex] = value;
            } else {
                context.log.warn(`Columna ${key} no encontrada en Excel.`);
            }
        }

        // 5. Enviar el PATCH para actualizar la fila
        const patchRes = await graphClient.api(rowEndpoint).patch({
            values: [rowValues]
        });

        context.res = {
            status: 200,
            headers: { 'Content-Type': 'application/json; charset=utf-8' },
            body: JSON.stringify({
                message: "Fila actualizada exitosamente.",
                rowNumber: rowNumber,
                updatedValues: rowValues
            }, null, 4)
        };
        
    } catch (error) {
        context.log.error("Error communicating with MS Graph:", error);
        context.res = {
            status: 500,
            body: { message: "Error updating Excel data.", error: error.message }
        };
    }
};

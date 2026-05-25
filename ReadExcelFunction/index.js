require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');

module.exports = async function (context, req) {
    context.log('HTTP trigger function to read Excel via MS Graph processed a request.');

    try {
        const clientId = process.env.AZURE_CLIENT_ID;
        // Para cuentas personales, el endpoint debe ser 'common' o 'consumers', no el Tenant ID específico de Entra.
        const tenantId = 'common';
        const refreshToken = process.env.GRAPH_REFRESH_TOKEN;

        if (!clientId || !refreshToken) {
            context.res = {
                status: 400,
                body: "Faltan las variables de entorno AZURE_CLIENT_ID o GRAPH_REFRESH_TOKEN."
            };
            return;
        }

        // 1. Obtener un nuevo Access Token usando el Refresh Token
        const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const bodyParams = new URLSearchParams();
        bodyParams.append('client_id', clientId);
        bodyParams.append('grant_type', 'refresh_token');
        bodyParams.append('refresh_token', refreshToken);
        bodyParams.append('scope', 'Files.Read.All offline_access');
        // Si el App Registration requiere client secret, se mandaría aquí, pero para OneDrive personal no solemos usarlo.
        
        const tokenResponse = await fetch(tokenEndpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: bodyParams
        });

        const tokenData = await tokenResponse.json();
        
        if (!tokenResponse.ok) {
             context.log.error("Fallo al renovar token:", tokenData);
             throw new Error("No se pudo obtener access_token usando el refresh token.");
        }

        const accessToken = tokenData.access_token;

        // 2. Inicializar cliente Graph con el nuevo token
        const authProvider = {
            getAccessToken: async () => {
                return accessToken;
            }
        };

        const graphClient = Client.initWithMiddleware({ authProvider });

        // 3. Obtener el archivo desde el root de OneDrive usando su nombre exácto
        // Archivo encontrado en Documentos
        const filePath = "Documentos/neflis_negro.xlsx"; 
        
        // El nombre exacto es Hoja1 (sin espacio)
        const worksheetName = "Hoja1";
        
        // Endpoint para leer los datos usados ('usedRange') de la hoja
        const endpoint = `/me/drive/root:/${filePath}:/workbook/worksheets('${worksheetName}')/usedRange`;
        
        const response = await graphClient.api(endpoint).get();
        
        // Mapear las columnas que el usuario pidió
        const rows = response.values || [];
        const headers = rows[0] || [];
        const dataRows = rows.slice(1);

        // Columnas solicitadas por el usuario:
        // Streaming, Nombre, apellido, whatsapp, Column1, correo, contraseña, vencimiento, Metodo de pago, customer mail, operador, pin perfil, deben, observaciones
        
        // Helper function for Excel dates
        function excelDateToJSDate(serial) {
            if (!serial || isNaN(serial)) return serial; // Devuelve como estaba si no es número
            // Excel dates start from 1900-01-01 (serial = 1)
            // Note: Excel has a leap year bug for 1900, so we subtract 1 day.
            let utc_days  = Math.floor(serial - 25569);
            let utc_value = utc_days * 86400;                                        
            let date_info = new Date(utc_value * 1000);
            
            // Format as YYYY-MM-DD
            let year = date_info.getUTCFullYear();
            let month = String(date_info.getUTCMonth() + 1).padStart(2, '0');
            let day = String(date_info.getUTCDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        const mappedData = dataRows.map((row, index) => {
            let obj = {
                rowNumber: index + 2 // Excel rows are 1-indexed, headers are row 1, data starts at row 2
            };
            headers.forEach((headerName, headerIndex) => {
                // Renombrar "Column1" o variantes a "numero"
                let finalHeader = headerName;
                if (headerName === "Column1" || headerName === "Numero" || headerName === "numero") {
                    finalHeader = "numero";
                }
                
                let val = row[headerIndex];
                
                // Formatear la fecha si la columna es "vencimiento"
                if (finalHeader === "vencimiento" && typeof val === 'number') {
                    val = excelDateToJSDate(val);
                }
                
                obj[finalHeader] = val;
            });
            return obj;
        });

        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'application/json; charset=utf-8'
            },
            body: JSON.stringify({
                message: "Datos obtenidos de OneDrive exitosamente.",
                totalRows: mappedData.length,
                data: mappedData
            }, null, 4)
        };
        
    } catch (error) {
        context.log.error("Error communicating with MS Graph:", error);
        context.res = {
            status: 500,
            body: {
                message: "Error retrieving Excel data.",
                error: error.message
            }
        };
    }
};

require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');

module.exports = async function (context, req) {
    context.log('HTTP trigger function to read Historico Excel via MS Graph processed a request.');

    try {
        const clientId = process.env.AZURE_CLIENT_ID;
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

        const filePath = "Documentos/neflis_negro.xlsx"; 
        
        // El nombre exacto de la hoja
        const worksheetName = "historico";
        
        // Endpoint para leer los datos usados ('usedRange') de la hoja
        const endpoint = `/me/drive/root:/${filePath}:/workbook/worksheets('${worksheetName}')/usedRange`;
        
        const response = await graphClient.api(endpoint).get();
        
        // Debido a que las columnas se pegan repetidas hacia la izquierda mes a mes,
        // no podemos hacer un mapeo limpio en un objeto con keys únicas.
        // Por eso, devolvemos las filas tal cual (una matriz 2D de valores).
        const rows = response.values || [];

        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'application/json; charset=utf-8'
            },
            body: JSON.stringify({
                message: "Datos de historial obtenidos exitosamente.",
                totalRows: rows.length,
                data: rows
            }, null, 4)
        };
        
    } catch (error) {
        context.log.error("Error communicating with MS Graph:", error);
        context.res = {
            status: 500,
            body: {
                message: "Error retrieving Historico data.",
                error: error.message
            }
        };
    }
};

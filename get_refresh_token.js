const msal = require('@azure/msal-node');
const readline = require('readline');

// ==========================================
// INSTRUCCIONES:
// 1. Reemplaza estos valores con los de tu App Registration en Azure
// ==========================================
const CLIENT_ID = "dd590625-bd57-487f-94c9-c8fb4c44ebfb"; 
const TENANT_ID = "common"; 

const config = {
    auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    }
};

const pca = new msal.PublicClientApplication(config);

const getTokenRequest = {
    // Permisos necesarios para leer y escribir el Excel en tu OneDrive
    scopes: ["Files.Read", "Files.Read.All", "Files.ReadWrite.All", "offline_access"],
    deviceCodeCallback: (response) => {
        console.log("\n===============================================");
        console.log(">>> PASO 1: ABRE TU NAVEGADOR EN ESTA URL <<<");
        console.log("https://microsoft.com/devicelogin");
        console.log("\n>>> PASO 2: INGRESA ESTE CÓDIGO EN LA WEB <<<");
        console.log(`         ${response.userCode}`); 
        console.log("===============================================\n");
        console.log("(Este NO es el código de tu Authenticator, es el código para vincular esta terminal)");
    }
};

pca.acquireTokenByDeviceCode(getTokenRequest).then((response) => {
    console.log("\n===============================================");
    console.log("¡Autenticación exitosa!");
    console.log("===============================================\n");
    // El objeto account.homeAccountId a menudo contiene información útil, pero lo que queremos 
    // es poder obtener tokens futuros silenciosamente.
    
    // MSAL Node maneja el caché de tokens para nosotros internamente en este objeto de configuración.
    // Sin embargo, para Azure Functions necesitamos el REFRESH TOKEN explícito para guardarlo.
    // Una forma sencilla es acceder al caché directamente:
    const tokenCache = pca.getTokenCache().serialize();
    const parsedCache = JSON.parse(tokenCache);
    
    let refreshToken = "";
    if (parsedCache.RefreshToken) {
        const keys = Object.keys(parsedCache.RefreshToken);
        if (keys.length > 0) {
            refreshToken = parsedCache.RefreshToken[keys[0]].secret;
        }
    }

    if (refreshToken) {
        console.log("ESTE ES TU REFRESH TOKEN. CÓPIALO Y GUÁRDALO:");
        console.log("-----------------------------------------------");
        console.log(refreshToken);
        console.log("-----------------------------------------------");
        console.log("\nPega este token en tu archivo local.settings.json como 'GRAPH_REFRESH_TOKEN'");
        console.log("Y también agrégalo como una variable de entorno en tu Azure Function App en el portal.");
    } else {
        console.log("No se pudo extraer el Refresh Token. Revisa los permisos (offline_access es requerido).");
    }

}).catch((error) => {
    console.error("Error durante la autenticación:", error);
});

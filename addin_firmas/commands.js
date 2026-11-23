/*
 * Complemento de Firmas Corporativas - Grupo Sandor
 * Conecta directamente a Entra ID (Microsoft Graph) via Nested App Authentication (NAA)
 * y actualiza la firma en vivo en OnNewMessageCompose y OnMessageFromChanged.
 */

const CLIENT_ID = "6f18f787-f35d-4bdb-91ea-770511d90e54";
const TENANT_ID = "c226adba-bbe6-461e-9764-3196c2c862e0";
const ASSETS_BASE = "https://liderticelite.github.io/Firmas-de-Correo";

let msalInstance = null;
let cachedProfile = null;

Office.onReady(() => {
    // Office listo
});

/**
 * Obtiene el perfil real del usuario desde Microsoft Graph (Entra ID)
 */
async function obtenerPerfilUsuario() {
    if (cachedProfile) {
        return cachedProfile;
    }

    const fallbackProfile = {
        displayName: Office.context.mailbox.userProfile.displayName || "",
        jobTitle: "",
        phone: ""
    };

    try {
        if (!msalInstance && typeof msal !== "undefined" && msal.createNestablePublicClientApplication) {
            msalInstance = await msal.createNestablePublicClientApplication({
                auth: {
                    clientId: CLIENT_ID,
                    authority: `https://login.microsoftonline.com/${TENANT_ID}`
                }
            });
        }

        if (msalInstance) {
            const tokenResponse = await msalInstance.acquireTokenSilent({
                scopes: ["User.Read"]
            });

            if (tokenResponse && tokenResponse.accessToken) {
                const res = await fetch(
                    "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,businessPhones,mail",
                    {
                        headers: {
                            Authorization: `Bearer ${tokenResponse.accessToken}`
                        }
                    }
                );

                if (res.ok) {
                    const data = await res.json();
                    const tel = data.mobilePhone || (data.businessPhones && data.businessPhones[0]) || "";
                    cachedProfile = {
                        displayName: data.displayName || fallbackProfile.displayName,
                        jobTitle: data.jobTitle || "",
                        phone: tel
                    };
                    return cachedProfile;
                }
            }
        }
    } catch (err) {
        console.warn("No se pudo consultar Graph en segundo plano, usando respaldo:", err);
    }

    return fallbackProfile;
}

/**
 * Genera el HTML corporativo exacto según el dominio del remitente
 */
function construirHtmlFirma(senderEmail, perfil) {
    const email = (senderEmail || "").trim().toLowerCase();
    const domain = email.split("@")[1] || "";

    // Dominios omitidos explícitamente (dejan el correo sin firma)
    if (domain === "toctocaseo.com" || domain === "fjsb.com.co") {
        return "";
    }

    const nombre = perfil.displayName || "";
    const cargo = perfil.jobTitle || "";

    // 1. ELITE FACILITY MANAGEMENT SAS
    if (domain === "elite.com.co") {
        const cdn = `${ASSETS_BASE}/assets_elite`;
        const tel = perfil.phone || "601 743 18 22";
        return `<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;border-collapse:collapse;background-color:#FFFFFF;"><tr><td width="229" height="131" style="vertical-align:middle;padding:0;width:229px;height:131px;"><a href="https://elitefms.com.co/" target="_blank" style="text-decoration:none;"><img src="${cdn}/logo_elite.png" alt="Elite Facility Management" width="229" height="131" style="display:block;border:0;width:229px;height:131px;"></a></td><td height="131" style="vertical-align:middle;padding-left:22px;padding-right:22px;height:131px;"><div style="font-size:14px;font-weight:600;color:#3A81B1;line-height:1.25;margin:0;white-space:nowrap;">${nombre}</div><div style="font-size:13px;font-weight:700;color:#3A81B1;line-height:1.25;margin-top:2px;margin-bottom:8px;white-space:nowrap;">${cargo}</div><table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding-bottom:5px;vertical-align:middle;width:20px;"><img src="${cdn}/icon_email.png" alt="Email" width="15" height="15" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:7px;vertical-align:middle;font-size:12.5px;color:#3A81B1;white-space:nowrap;"><a href="mailto:${email}" style="color:#3A81B1;text-decoration:none;">${email}</a></td></tr><tr><td style="padding-bottom:5px;vertical-align:middle;width:20px;"><img src="${cdn}/icon_phone.png" alt="Tel" width="15" height="15" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:7px;vertical-align:middle;font-size:12.5px;color:#3A81B1;white-space:nowrap;"><a href="tel:${tel}" style="color:#3A81B1;text-decoration:none;">${tel}</a></td></tr><tr><td style="vertical-align:middle;width:20px;"><img src="${cdn}/icon_location.png" alt="Ubicación" width="15" height="15" style="display:block;border:0;"></td><td style="padding-left:7px;vertical-align:middle;font-size:12.5px;color:#3A81B1;white-space:nowrap;">Calle 71 No. 72A – 41 Bogotá D. C</td></tr></table></td><td height="131" style="vertical-align:middle;padding:0;"><table cellpadding="0" cellspacing="0" border="0"><tr><td width="1" height="96" bgcolor="#3A81B1" style="width:1px;height:96px;background-color:#3A81B1;font-size:0;line-height:0;">&nbsp;</td></tr></table></td><td height="131" style="vertical-align:middle;padding-left:22px;padding-right:10px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right:4px;"><a href="https://www.instagram.com/elitefms.co/" target="_blank"><img src="${cdn}/icon_instagram.png" alt="Instagram" width="35" height="36" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.facebook.com/elitefms/" target="_blank"><img src="${cdn}/icon_facebook.png" alt="Facebook" width="35" height="36" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://co.linkedin.com/company/elitefms" target="_blank"><img src="${cdn}/icon_linkedin.png" alt="LinkedIn" width="35" height="36" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.tiktok.com/@elitefms.co" target="_blank"><img src="${cdn}/icon_tiktok.png" alt="TikTok" width="35" height="36" style="display:block;border:0;"></a></td><td><a href="https://www.youtube.com/@ELITEFACILITYMANAGEMENT" target="_blank"><img src="${cdn}/icon_youtube.png" alt="YouTube" width="35" height="36" style="display:block;border:0;"></a></td></tr></table></td></tr><tr><td align="center" style="padding-top:8px;"><a href="https://elitefms.com.co/" target="_blank"><img src="${cdn}/btn_web.png" alt="www.elitefms.com.co" width="143" height="22" style="display:block;border:0;margin:0 auto;"></a></td></tr></table></td></tr></table><table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;max-width:680px;margin-top:16px;"><tr><td style="font-size:11px;color:#3A81B1;line-height:1.35;"><strong>Aviso de confidencialidad:</strong> Este mensaje y sus anexos pueden contener información confidencial y están dirigidos exclusivamente a sus destinatarios. Si lo recibió por error, por favor informe al remitente y elimínelo. Las comunicaciones emitidas desde esta cuenta se realizan dentro del alcance de las funciones y facultades asignadas al remitente.</td></tr><tr><td style="padding-top:8px;font-size:11px;color:#3A81B1;font-weight:700;font-style:italic;">🌱 Antes de imprimir este correo, considere si es realmente necesario. Cuidar el medio ambiente es responsabilidad de todos.</td></tr></table>`;
    }

    // 2. SARUPETROL SAS
    if (domain === "sarupetrol.com") {
        const cdn = `${ASSETS_BASE}/assets_sarupetrol`;
        const tel = perfil.phone || "601 927 80 89";
        return `<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;border-collapse:collapse;background-color:#FFFFFF;"><tr><td style="vertical-align:middle;padding-right:18px;"><a href="https://www.sarupetrol.com/" target="_blank"><img src="${cdn}/logo_sarupetrol.png" alt="Sarupetrol" width="130" height="76" style="display:block;border:0;"></a></td><td width="1" height="96" bgcolor="#08833C" style="width:1px;height:96px;background-color:#08833C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;white-space:nowrap;"><div style="font-size:15px;font-weight:600;color:#08833C;margin:0;">${nombre}</div><div style="font-size:14px;font-weight:700;color:#08833C;margin-top:3px;">${cargo}</div></td><td width="1" height="96" bgcolor="#08833C" style="width:1px;height:96px;background-color:#08833C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_email.png" alt="Email" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#08833C;white-space:nowrap;"><a href="mailto:${email}" style="color:#08833C;text-decoration:none;font-weight:500;">${email}</a></td></tr><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_phone.png" alt="Tel" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#08833C;white-space:nowrap;"><a href="tel:${tel}" style="color:#08833C;text-decoration:none;">${tel}</a></td></tr><tr><td style="width:20px;vertical-align:middle;"><img src="${cdn}/icon_location.png" alt="Ubicación" width="16" height="16" style="display:block;border:0;"></td><td style="padding-left:8px;font-size:13px;color:#08833C;white-space:nowrap;">Calle 71 No. 72A – 41 Bogotá D. C</td></tr></table></td><td width="1" height="96" bgcolor="#08833C" style="width:1px;height:96px;background-color:#08833C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding-left:18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right:4px;"><a href="https://www.instagram.com/sarupetrol/" target="_blank"><img src="${cdn}/icon_instagram.png" alt="IG" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.facebook.com/reel/1517311290167042" target="_blank"><img src="${cdn}/icon_facebook.png" alt="FB" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.linkedin.com/company/sarupetrol-s-a-s" target="_blank"><img src="${cdn}/icon_linkedin.png" alt="IN" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.tiktok.com/@sarupetrol" target="_blank"><img src="${cdn}/icon_tiktok.png" alt="TT" width="33" height="34" style="display:block;border:0;"></a></td><td><a href="https://www.youtube.com/@ELITEFACILITYMANAGEMENT" target="_blank"><img src="${cdn}/icon_youtube.png" alt="YT" width="33" height="34" style="display:block;border:0;"></a></td></tr></table></td></tr><tr><td align="center" style="padding-top:7px;"><a href="https://www.sarupetrol.com/" target="_blank"><img src="${cdn}/btn_web.png" alt="www.sarupetrol.com" width="145" height="24" style="display:block;border:0;margin:0 auto;"></a></td></tr></table></td></tr></table><table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;max-width:680px;margin-top:16px;"><tr><td style="font-size:11px;color:#08833C;line-height:1.35;"><strong>Aviso de confidencialidad:</strong> Este mensaje y sus anexos pueden contener información confidencial y están dirigidos exclusivamente a sus destinatarios. Si lo recibió por error, por favor informe al remitente y elimínelo. Las comunicaciones emitidas desde esta cuenta se realizan dentro del alcance de las funciones y facultades asignadas al remitente.</td></tr><tr><td style="padding-top:8px;font-size:11px;color:#08833C;font-weight:700;font-style:italic;">🌱 Antes de imprimir este correo, considere si es realmente necesario. Cuidar el medio ambiente es responsabilidad de todos.</td></tr></table>`;
    }

    // 3. MASTER SERVICE GROUP
    if (domain === "masterservicegroup.com.co") {
        const cdn = `${ASSETS_BASE}/assets_masterservice`;
        const tel = perfil.phone || "601 743 18 22";
        return `<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;border-collapse:collapse;background-color:#FFFFFF;"><tr><td style="vertical-align:middle;padding-right:18px;"><a href="https://masterservicegroup.com.co/" target="_blank"><img src="${cdn}/logo_masterservice.png" alt="Master Service" width="160" height="68" style="display:block;border:0;"></a></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;white-space:nowrap;"><div style="font-size:15px;font-weight:600;color:#D9261C;margin:0;">${nombre}</div><div style="font-size:14px;font-weight:700;color:#D9261C;margin-top:3px;">${cargo}</div></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_email.png" alt="Email" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;"><a href="mailto:${email}" style="color:#D9261C;text-decoration:none;">${email}</a></td></tr><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_phone.png" alt="Tel" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;"><a href="tel:${tel}" style="color:#D9261C;text-decoration:none;">${tel}</a></td></tr><tr><td style="width:20px;vertical-align:middle;"><img src="${cdn}/icon_location.png" alt="Ubicación" width="16" height="16" style="display:block;border:0;"></td><td style="padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;">Calle 71 No. 72A – 41 Bogotá D. C</td></tr></table></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding-left:18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right:4px;"><a href="https://www.instagram.com/masterservicegroup.co/" target="_blank"><img src="${cdn}/icon_instagram.png" alt="IG" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.facebook.com/masterservicesgroup" target="_blank"><img src="${cdn}/icon_facebook.png" alt="FB" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://co.linkedin.com/company/master-service-group-sas" target="_blank"><img src="${cdn}/icon_linkedin.png" alt="IN" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.tiktok.com/@masterservicegroup.co" target="_blank"><img src="${cdn}/icon_tiktok.png" alt="TT" width="33" height="34" style="display:block;border:0;"></a></td><td><a href="https://www.youtube.com/@ELITEFACILITYMANAGEMENT" target="_blank"><img src="${cdn}/icon_youtube.png" alt="YT" width="33" height="34" style="display:block;border:0;"></a></td></tr></table></td></tr><tr><td align="center" style="padding-top:7px;"><a href="https://masterservicegroup.com.co/" target="_blank"><img src="${cdn}/btn_web.png" alt="www.masterservicegroup.com.co" width="175" height="24" style="display:block;border:0;margin:0 auto;"></a></td></tr></table></td></tr></table><table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;max-width:680px;margin-top:16px;"><tr><td style="font-size:11px;color:#D9261C;line-height:1.35;"><strong>Aviso de confidencialidad:</strong> Este mensaje y sus anexos pueden contener información confidencial y están dirigidos exclusivamente a sus destinatarios. Si lo recibió por error, por favor informe al remitente y elimínelo. Las comunicaciones emitidas desde esta cuenta se realizan dentro del alcance de las funciones y facultades asignadas al remitente.</td></tr><tr><td style="padding-top:8px;font-size:11px;color:#D9261C;font-weight:700;font-style:italic;">🌱 Antes de imprimir este correo, considere si es realmente necesario. Cuidar el medio ambiente es responsabilidad de todos.</td></tr></table>`;
    }

    // 4. SEGURIDAD Y VIGILANCIA CIEN POR CIENTO
    if (domain === "g.cienporciento.com.co" || domain === "cienporciento.com.co") {
        const cdn = `${ASSETS_BASE}/assets_cienporciento`;
        const tel = perfil.phone || "350 265 44 97";
        return `<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;border-collapse:collapse;background-color:#FFFFFF;"><tr><td style="vertical-align:middle;padding-right:18px;"><a href="https://cienporciento.com.co/" target="_blank"><img src="${cdn}/logo_cienporciento.png" alt="Cien Por Ciento" width="135" height="72" style="display:block;border:0;"></a></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;white-space:nowrap;"><div style="font-size:15px;font-weight:600;color:#D9261C;margin:0;">${nombre}</div><div style="font-size:14px;font-weight:700;color:#D9261C;margin-top:3px;">${cargo}</div></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding:0 18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_email.png" alt="Email" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;"><a href="mailto:${email}" style="color:#D9261C;text-decoration:none;">${email}</a></td></tr><tr><td style="padding-bottom:5px;width:20px;vertical-align:middle;"><img src="${cdn}/icon_phone.png" alt="Tel" width="16" height="16" style="display:block;border:0;"></td><td style="padding-bottom:5px;padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;"><a href="tel:${tel}" style="color:#D9261C;text-decoration:none;">${tel}</a></td></tr><tr><td style="width:20px;vertical-align:middle;"><img src="${cdn}/icon_location.png" alt="Ubicación" width="16" height="16" style="display:block;border:0;"></td><td style="padding-left:8px;font-size:13px;color:#D9261C;white-space:nowrap;">Bogotá D. C</td></tr></table></td><td width="1" height="96" bgcolor="#D9261C" style="width:1px;height:96px;background-color:#D9261C;font-size:0;line-height:0;">&nbsp;</td><td style="vertical-align:middle;padding-left:18px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right:4px;"><a href="https://www.instagram.com/seguridadcienporciento/" target="_blank"><img src="${cdn}/icon_instagram.png" alt="IG" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.facebook.com/SeguridadCienPorCiento" target="_blank"><img src="${cdn}/icon_facebook.png" alt="FB" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://co.linkedin.com/company/seguridad-y-vigilancia-cien-por-ciento-ltda" target="_blank"><img src="${cdn}/icon_linkedin.png" alt="IN" width="33" height="34" style="display:block;border:0;"></a></td><td style="padding-right:4px;"><a href="https://www.tiktok.com/@seguridad100ltda" target="_blank"><img src="${cdn}/icon_tiktok.png" alt="TT" width="33" height="34" style="display:block;border:0;"></a></td><td><a href="https://www.youtube.com/@ELITEFACILITYMANAGEMENT" target="_blank"><img src="${cdn}/icon_youtube.png" alt="YT" width="33" height="34" style="display:block;border:0;"></a></td></tr></table></td></tr><tr><td align="center" style="padding-top:7px;"><a href="https://cienporciento.com.co/" target="_blank"><img src="${cdn}/btn_web.png" alt="www.cienporciento.com.co" width="160" height="24" style="display:block;border:0;margin:0 auto;"></a></td></tr></table></td></tr></table><table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,sans-serif;max-width:680px;margin-top:16px;"><tr><td style="font-size:11px;color:#D9261C;line-height:1.35;"><strong>Aviso de confidencialidad:</strong> Este mensaje y sus anexos pueden contener información confidencial y están dirigidos exclusivamente a sus destinatarios. Si lo recibió por error, por favor informe al remitente y elimínelo. Las comunicaciones emitidas desde esta cuenta se realizan dentro del alcance de las funciones y facultades asignadas al remitente.</td></tr><tr><td style="padding-top:8px;font-size:11px;color:#D9261C;font-weight:700;font-style:italic;">🌱 Antes de imprimir este correo, considere si es realmente necesario. Cuidar el medio ambiente es responsabilidad de todos.</td></tr></table>`;
    }

    return "";
}

/**
 * Ejecuta la actualización de firma en la ventana de redacción
 */
async function aplicarFirmaActual(event) {
    try {
        const item = Office.context.mailbox.item;
        const perfil = await obtenerPerfilUsuario();

        item.from.getAsync((fromResult) => {
            let senderEmail = Office.context.mailbox.userProfile.emailAddress;
            if (fromResult.status === Office.AsyncResultStatus.Succeeded && fromResult.value && fromResult.value.emailAddress) {
                senderEmail = fromResult.value.emailAddress;
            }

            const htmlFirma = construirHtmlFirma(senderEmail, perfil);

            item.body.setSignatureAsync(
                htmlFirma,
                { coercionType: Office.CoercionType.Html },
                (setResult) => {
                    if (setResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Error setSignatureAsync:", setResult.error.message);
                    }
                    event.completed();
                }
            );
        });
    } catch (e) {
        console.error("Error general al aplicar firma:", e);
        event.completed();
    }
}

function onNewMessageComposeHandler(event) {
    aplicarFirmaActual(event);
}

function onMessageFromChangedHandler(event) {
    aplicarFirmaActual(event);
}

if (typeof global !== "undefined") {
    global.onNewMessageComposeHandler = onNewMessageComposeHandler;
    global.onMessageFromChangedHandler = onMessageFromChangedHandler;
}
self.onNewMessageComposeHandler = onNewMessageComposeHandler;
self.onMessageFromChangedHandler = onMessageFromChangedHandler;

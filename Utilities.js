// ===============================================================
// UTILIDADES GENERALES
// ===============================================================

/**
 * Obtiene el perfil del usuario activo (email, rol y foto de perfil).
 */
function getUserProfileLogic() {
    try {
        const email = Session.getActiveUser().getEmail();
        if (!email) throw new Error('No se pudo obtener el email del usuario activo');

        const role = USER_ROLES[email] || 'USER';

        // Generar foto de perfil usando Gravatar + fallback
        const photoUrl = generateProfileImageWithFallbacks(email);

        Logger.log(`👤 Perfil generado para: ${email}, Foto: ${photoUrl}`); // Debug

        return {
            email: email,
            role: role,
            photoUrl: photoUrl,
            status: "success"
        };
    } catch (e) {
        Logger.log(`Error crítico al obtener perfil de usuario: ${e.message}`);
        throw new Error(`No se pudo obtener el perfil de usuario: ${e.message}`);
    }
}

/**
 * Genera una foto de perfil con múltiples fallbacks.
 */
function generateProfileImageWithFallbacks(email) {
    try {
        // Limpiar y normalizar el email
        const cleanEmail = email.toLowerCase().trim();

        // Generar hash MD5 para Gravatar
        const emailHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cleanEmail)
            .map(byte => {
                const unsignedByte = byte < 0 ? byte + 256 : byte;
                return unsignedByte.toString(16).padStart(2, '0');
            })
            .join('');


        // ✅ SOLUCIÓN REAL Y DEFINITIVA: Fallback con Anonymous Animals para avatares 100% de animales.
        const fallbackUrl = `https://anonymous-animals.azurewebsites.net/avatar/${encodeURIComponent(cleanEmail)}`;

        // Esta línea USA la fallbackUrl. ¡Está correcta, no la cambies!
        const gravatarUrl = `https://www.gravatar.com/avatar/${emailHash}?s=80&d=${encodeURIComponent(fallbackUrl)}&r=pg`;

        Logger.log(`🔗 Gravatar URL con fallback a DiceBear: ${gravatarUrl}`);

        return gravatarUrl;

    } catch (e) {
        Logger.log(`Error generando imagen con fallbacks: ${e.message}`);

        // ✅ MEJORA: Fallback final consistente con DiceBear si todo lo demás falla.
        return `https://api.dicebear.com/8.x/micah/png?seed=${encodeURIComponent(email)}&radius=50`;
    }
}
/**
 * 🔧 FUNCIÓN DE SANITIZACIÓN ROBUSTA PARA EL BACKEND
 * Asegura que los datos de los tickets enviados al frontend tengan un formato consistente.
 */
function sanitizeTicketData(tickets) {
    return tickets.map(ticket => {
        const sanitized = {};
        const expectedFields = [
            'Timestamp', 'Creado Por', 'Usuario Cliente', 'N° Caso Yoizen',
            'Planilla (Cliente)', 'URL Adjunto', 'Herramienta',
            'Pedido de escalamiento (Asunto)', 'ID Derivacion FAN',
            'N° Ticket Escalamiento', 'Ticket Interno', 'Estado',
            'Observaciones Auditor', 'Prioridad', 'Tomado Por'
        ];
        expectedFields.forEach(field => {
            let value = ticket[field];
            if (field === 'Timestamp') {
                sanitized[field] = value ? new Date(value).toISOString() : new Date().toISOString();
            } else {
                sanitized[field] = (value === null || value === undefined) ? '' : String(value);
            }
        });
        return sanitized;
    });
}

// --- Funciones de Ejemplo / Prueba ---

function ejemploUsoCommentsSheet() {
    try {
        const commentsSheet = getOrCreateCommentsSheet();
        const validation = validateCommentsSheet(commentsSheet);
        if (!validation.valid) throw new Error(`Hoja inválida: ${validation.error}`);
        Logger.log('✅ Hoja de comentarios lista para usar');
        addCommentSafely('TK-TEST', 'Usuario Test', 'Este es un comentario de prueba');
    } catch (error) {
        Logger.log(`❌ Error en ejemploUsoCommentsSheet: ${error.message}`);
    }
}

function addCommentSafely(ticketInterno, autor, comentario) {
    try {
        const commentsSheet = getOrCreateCommentsSheet();
        const validation = validateCommentsSheet(commentsSheet);
        if (!validation.valid) throw new Error(`Hoja de comentarios inválida: ${validation.error}`);
        commentsSheet.appendRow([ticketInterno, new Date(), autor, comentario]);
        Logger.log(`✅ Comentario añadido para ticket ${ticketInterno}`);
        return true;
    } catch (error) {
        Logger.log(`❌ Error al añadir comentario: ${error.message}`);
        return false;
    }
}

function ejemploUsoHistorySheet() {
    try {
        const historySheet = getOrCreateHistorySheet();
        const validation = validateHistorySheet(historySheet);
        if (!validation.valid) throw new Error(`Hoja inválida: ${validation.error}`);
        Logger.log('✅ Hoja de historial lista para usar');
    } catch (error) {
        Logger.log(`❌ Error en ejemploUsoHistorySheet: ${error.message}`);
    }
}
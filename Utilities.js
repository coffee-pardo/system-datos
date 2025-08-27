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
                // Convertir byte con signo a sin signo
                const unsignedByte = byte < 0 ? byte + 256 : byte;
                return unsignedByte.toString(16).padStart(2, '0');
            })
            .join('');

        // Generar iniciales del nombre
        const name = cleanEmail.split('@')[0];
        const nameParts = name.split(/[._-]/); // Dividir por puntos, guiones bajos o guiones
        const initials = nameParts.map(part => part.charAt(0).toUpperCase()).join('').substring(0, 2);

        // Generar color consistente basado en el email
        const colorSeed = cleanEmail.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
        const colors = [
            '6366f1', // Indigo
            '8b5cf6', // Violet
            'ec4899', // Pink
            'ef4444', // Red
            'f97316', // Orange
            'eab308', // Yellow
            '22c55e', // Green
            '10b981', // Emerald
            '06b6d4', // Cyan
            '3b82f6'  // Blue
        ];
        const backgroundColor = colors[colorSeed % colors.length];

        // URL de fallback con UI Avatars (servicio gratuito y confiable)
        const fallbackUrl = `https://ui-avatars.com/api/?name=${encodeURIComponent(initials)}&background=${backgroundColor}&color=ffffff&size=80&rounded=true&bold=true&format=png`;

        // URL de Gravatar con fallback a UI Avatars
        const gravatarUrl = `https://www.gravatar.com/avatar/${emailHash}?s=80&d=${encodeURIComponent(fallbackUrl)}&r=pg`;

        Logger.log(`📷 URL generada - Email: ${cleanEmail}, Hash: ${emailHash}, Iniciales: ${initials}, Color: ${backgroundColor}`);
        Logger.log(`🔗 Gravatar URL: ${gravatarUrl}`);

        return gravatarUrl;

    } catch (e) {
        Logger.log(`Error generando imagen con fallbacks: ${e.message}`);
        // Fallback final si todo falla
        return 'https://ui-avatars.com/api/?name=U&background=6366f1&color=ffffff&size=80&rounded=true&bold=true&format=png';
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
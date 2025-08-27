// ===============================================================
// CONFIGURACIÓN CENTRALIZADA
// ===============================================================

/**
 * @const {string} El ID único de la Hoja de Cálculo de Google que actuará como base de datos.
 */
const SPREADSHEET_ID = "1gHM6mIH8n5LS6_qLQv42ytRuU8a3R1c0mN0pXxXzZ3M";

/**
 * @const {string} El ID de la carpeta de Google Drive donde se guardarán los archivos adjuntos.
 */
const SUPERUSER_DRIVE_FOLDER_ID = "16U50VOK8jSqY9qwQQDuXvXBgYSTiD2mB";

/**
 * @const {Object<string, string>} Un objeto que define roles de usuario para control de acceso.
 */
const USER_ROLES = {
    'alejandro.garciap@konecta.com': 'SUPERUSER',
    'facundo.alvarez@konecta.com': 'SUPERUSER',
    'lautaro.escalante@konecta.com': 'SUPERUSER',
    'lucas.maldonado@konecta.com': 'AUDITOR',
    'elias.stessens@konecta.com': 'AUDITOR'
};
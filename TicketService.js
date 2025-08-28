// ===============================================================
// SERVICIO DE TICKETS (LÓGICA DE NEGOCIO)
// ===============================================================

/**
 * Obtiene los datos para el dashboard, aplicando filtros y paginación.
 */
function getDashboardDataLogic(options = {}) {  // LÓGICA DE PERMISOS 
    try {
        const userProfile = getUserProfileLogic();
        const sheet = getOrCreateSheet();

        if (sheet.getLastRow() <= 1) {
            return { tickets: [], allTicketsForStats: [], totalTickets: 0, allPlanillas: [], allCreators: [], allMotivosEscalamiento: [], allSuperiores: [], status: "ok" };
        }

        const rawData = sheet.getDataRange().getValues();
        const headers = rawData.shift();

        const tickets = rawData.map(row => {
            const ticket = {};
            headers.forEach((header, index) => {
                let value = row[index];
                if (header === 'Timestamp' && value instanceof Date) value = value.toISOString();
                ticket[header] = value;
            });
            return ticket;
        });

        const getUniqueSortedValues = (key) => [...new Set(tickets.map(t => (t[key] || '').toString().trim()).filter(Boolean))].sort();
        const allPlanillas = getUniqueSortedValues('Planilla (Cliente)');
        const allCreators = getUniqueSortedValues('Creado Por');
        const allMotivosEscalamiento = getUniqueSortedValues('Pedido de escalamiento (Asunto)');
        const superioresResponse = getAllSuperiores();
        const allSuperiores = superioresResponse.status === 'success' ? superioresResponse.superiores : [];

        let filteredTickets = tickets;

        if (options.searchQuery) {
            const query = options.searchQuery.toLowerCase();
            filteredTickets = filteredTickets.filter(t =>
                Object.values(t).some(val => String(val).toLowerCase().includes(query))
            );
        }
        if (options.filterStatus) filteredTickets = filteredTickets.filter(t => t['Estado'] === options.filterStatus);
        if (options.filterTool) filteredTickets = filteredTickets.filter(t => t['Herramienta'] === options.filterTool);
        if (options.filterMotivoEscalamiento) filteredTickets = filteredTickets.filter(t => (t['Pedido de escalamiento (Asunto)'] || '') === options.filterMotivoEscalamiento);
        if (options.filterCreatedBy) filteredTickets = filteredTickets.filter(t => t['Creado Por'] === options.filterCreatedBy);
        if (options.filterPriority) filteredTickets = filteredTickets.filter(t => (t['Prioridad'] || 'Normal') === options.filterPriority);
        if (options.startDate) filteredTickets = filteredTickets.filter(t => new Date(t['Timestamp']) >= new Date(options.startDate));
        // EMPIEZO NUEVO: Filtro por equipo (superior)
        if (options.filterEquipo) {
            const emailsDelEquipo = getEmailsByEquipo(options.filterEquipo);
            if (emailsDelEquipo.length > 0) {
                filteredTickets = filteredTickets.filter(t => emailsDelEquipo.includes(t['Creado Por']));
            }
        }
        // TERMINA NUEVO: Filtro por equipo (superior)

        if (options.endDate) {
            const end = new Date(options.endDate);
            end.setHours(23, 59, 59, 999);
            filteredTickets = filteredTickets.filter(t => new Date(t['Timestamp']) <= end);
        }
        // NUEVO: Lógica para filtrar por "Mis Tickets" o mostrar todos
        let allTicketsForUser = filteredTickets;
        if (options.filterMyTickets) {
            allTicketsForUser = filteredTickets.filter(t => t['Creado Por'] === userProfile.email);
        }
        // FIN NUEVO: Lógica para filtrar por "Mis Tickets"

        // fix global condición TODOS (CONDICIONAL TICKETS GLOBAL USUARIOS FINALES)
        //let allTicketsForUser = (userProfile.role === 'USER') ? filteredTickets.filter(t => t['Creado Por'] === userProfile.email) : filteredTickets; (CONDICIONAL TICKETS GLOBAL USUARIOS FINALES)
        //let allTicketsForUser = filteredTickets; // Ahora todos los usuarios ven todos los tickets filtrados
        // TERMINA  fix global condición TODOS (CONDICIONAL TICKETS GLOBAL USUARIOS FINALES)
        allTicketsForUser.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));
        const totalFilteredTickets = allTicketsForUser.length;
        const pageSize = options.pageSize || 10;
        const page = options.page || 1;
        const startIndex = (page - 1) * pageSize;
        const paginatedTickets = allTicketsForUser.slice(startIndex, startIndex + pageSize);

        return {
            tickets: sanitizeTicketData(paginatedTickets),
            allTicketsForStats: sanitizeTicketData(allTicketsForUser),
            totalTickets: totalFilteredTickets,
            allPlanillas,
            allCreators,
            allMotivosEscalamiento,
            // EMPIEZA NUEVO: Incluir superiores en la respuesta
            allSuperiores, // NUEVO: Incluir superiores en la respuesta
            // TERMINA NUEVO: Incluir superiores en la respuesta
            status: "ok"
        };
    } catch (error) {
        Logger.log(`ERROR GRAVE al leer datos: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `⚠️ No se pudieron cargar los datos. Detalle: ${error.message}` };
    }
}


/**
 * Crea un nuevo ticket con los datos enviados desde el formulario del frontend.
 */
function createTicketLogic(submissionData) {
    try {
        const sheet = getOrCreateSheet();
        const requiredFields = {
            usuario: "Usuario (U)",
            casoYoizen: "Número de caso Yoizen",
            herramienta: "Herramienta",
            planilla: "Planilla Escalamiento",
            motivoEscalamiento: "Motivo Escalamiento",
            idDerivacion: "ID Derivación FAN"
        };
        for (const key in requiredFields) {
            if (!submissionData[key] || String(submissionData[key]).trim() === '') {
                throw new Error(`El campo "${requiredFields[key]}" es obligatorio.`);
            }
        }

        let imageUrl = "No se adjuntó imagen";
        if (submissionData && submissionData.fileContent) {
            const parts = submissionData.fileContent.split(",");
            const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), submissionData.fileType, submissionData.fileName);
            const folder = DriveApp.getFolderById(SUPERUSER_DRIVE_FOLDER_ID);
            const file = folder.createFile(blob);
            imageUrl = file.getUrl();
        }

        const timestamp = new Date();
        const currentUser = Session.getActiveUser().getEmail();
        const ticketInterno = `TK-${Math.floor(100000 + Math.random() * 900000)}`;
        const newRow = [
            timestamp, currentUser, submissionData.usuario.trim(), submissionData.casoYoizen.trim(),
            submissionData.planilla.trim(), imageUrl, submissionData.herramienta.trim(),
            submissionData.motivoEscalamiento.trim(), (submissionData.idDerivacion || '').trim(),
            (submissionData.numeroTicket || '').trim(), ticketInterno, 'Registrado', '', 'Normal', ''
        ];
        sheet.appendRow(newRow);
        logTicketChange(ticketInterno, `Ticket creado por ${currentUser}`);

        const notificationMessage = `Nuevo ticket #${ticketInterno} creado por ${currentUser} - Asunto: ${submissionData.motivoEscalamiento}`;
        addNotificationToSheet('SUPERUSER', notificationMessage, ticketInterno);
        addNotificationToSheet('AUDITOR', notificationMessage, ticketInterno);

        return { status: "success", message: "Escalamiento recibido.", ticket: ticketInterno };
    } catch (error) {
        Logger.log(`ERROR GRAVE al crear ticket: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `Error en el servidor al crear el ticket: ${error.message}` };
    }
}


/**
 * Actualiza el estado, observaciones, prioridad y/o 'Tomado Por' de un ticket.
 */
function updateTicketStatusLogic(params) {
    try {
        const { ticketId, newStatus, newObservations, newPriority, newTomadoPor } = params;
        const userProfile = getUserProfileLogic();
        if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
            throw new Error('Acceso denegado: No tienes permisos para modificar tickets.');
        }
        const sheet = getOrCreateSheet();
        if (sheet.getLastRow() <= 1) throw new Error("No hay tickets para actualizar.");

        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const ticketIdColIdx = headers.indexOf('Ticket Interno');
        const statusColIdx = headers.indexOf('Estado');
        const obsColIdx = headers.indexOf('Observaciones Auditor');
        const prioridadColIdx = headers.indexOf('Prioridad');
        const creadoPorColIdx = headers.indexOf('Creado Por');
        const tomadoPorColIdx = headers.indexOf('Tomado Por');

        if (ticketIdColIdx === -1) throw new Error('Columna "Ticket Interno" no encontrada.');

        const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
        if (rowIndex === -1) throw new Error(`Ticket con ID ${ticketId} no encontrado.`);

        const row = data[rowIndex];
        const sheetRowIndex = rowIndex + 2; // +1 for headers, +1 for 0-based index
        let changes = [];
        let ticketCreatorEmail = row[creadoPorColIdx];

        if (newStatus !== undefined && newStatus !== row[statusColIdx]) {
            sheet.getRange(sheetRowIndex, statusColIdx + 1).setValue(newStatus);
            logTicketChange(ticketId, `Estado cambiado de '${row[statusColIdx]}' a '${newStatus}'`);
            addNotificationToSheet(ticketCreatorEmail, `Tu ticket #${ticketId} ha cambiado de estado: ${row[statusColIdx]} -> ${newStatus}`, ticketId);
            changes.push('estado');
        }
        if (newObservations !== undefined && obsColIdx > -1 && newObservations.trim() !== (row[obsColIdx] || '').trim()) {
            sheet.getRange(sheetRowIndex, obsColIdx + 1).setValue(newObservations.trim());
            logTicketChange(ticketId, `Observaciones actualizadas.`);
            addNotificationToSheet(ticketCreatorEmail, `Nueva observación del auditor en tu ticket #${ticketId}`, ticketId);
            changes.push('observaciones');
        }
        if (newPriority !== undefined && prioridadColIdx > -1 && newPriority !== (row[prioridadColIdx] || 'Normal')) {
            sheet.getRange(sheetRowIndex, prioridadColIdx + 1).setValue(newPriority);
            logTicketChange(ticketId, `Prioridad cambiada de '${row[prioridadColIdx] || 'Normal'}' a '${newPriority}'`);
            addNotificationToSheet(ticketCreatorEmail, `La prioridad de tu ticket #${ticketId} ha cambiado a: ${newPriority}`, ticketId);
            changes.push('prioridad');
        }
        if (newTomadoPor !== undefined && tomadoPorColIdx > -1 && newTomadoPor.trim() !== (row[tomadoPorColIdx] || '').trim()) {
            sheet.getRange(sheetRowIndex, tomadoPorColIdx + 1).setValue(newTomadoPor.trim());
            logTicketChange(ticketId, `Tomado por: '${newTomadoPor.trim()}'`);
            addNotificationToSheet(ticketCreatorEmail, `Tu ticket #${ticketId} ha sido tomado por ${newTomadoPor.trim()}`, ticketId);
            changes.push("'Tomado Por'");
        }

        if (changes.length > 0) {
            return { status: "success", message: `Ticket actualizado correctamente (${changes.join(', ')}).` };
        }
        return { status: "info", message: "No se realizaron cambios en el ticket." };
    } catch (error) {
        Logger.log(`ERROR al actualizar ticket: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `Error en el servidor al actualizar: ${error.message}` };
    }
}

/**
 * Permite a un auditor/superuser "tomar" un ticket.
 */
function takeTicketLogic(params) {
    try {
        const { ticketId } = params;
        const userProfile = getUserProfileLogic();
        if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
            throw new Error('Acceso denegado: Solo auditores o superusuarios pueden tomar tickets.');
        }
        const sheet = getOrCreateSheet();
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const ticketIdColIdx = headers.indexOf('Ticket Interno');
        const tomadoPorColIdx = headers.indexOf('Tomado Por');
        const creadoPorColIdx = headers.indexOf('Creado Por');

        if (ticketIdColIdx === -1 || tomadoPorColIdx === -1) throw new Error('Columnas críticas no encontradas.');

        const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
        if (rowIndex === -1) throw new Error(`Ticket con ID ${ticketId} no encontrado.`);

        const row = data[rowIndex];
        const sheetRowIndex = rowIndex + 2;
        const currentTomadoPor = (row[tomadoPorColIdx] || '').trim();
        const currentUserShortName = userProfile.email.split('@')[0];

        if (currentTomadoPor === currentUserShortName) {
            return { status: "info", message: `Ya has tomado el ticket ${ticketId}.` };
        } else if (currentTomadoPor !== '') {
            return { status: "info", message: `El ticket ${ticketId} ya ha sido tomado por ${currentTomadoPor}.` };
        }

        sheet.getRange(sheetRowIndex, tomadoPorColIdx + 1).setValue(currentUserShortName);
        logTicketChange(ticketId, `Ticket tomado por ${currentUserShortName}`);

        const ticketCreatorEmail = row[creadoPorColIdx];
        if (ticketCreatorEmail && ticketCreatorEmail !== userProfile.email) {
            addNotificationToSheet(ticketCreatorEmail, `Tu ticket #${ticketId} ha sido tomado por ${currentUserShortName}.`, ticketId);
        }

        return { status: "success", message: `Ticket ${ticketId} tomado por ti.` };
    } catch (error) {
        Logger.log(`ERROR al tomar ticket: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `Error en el servidor al tomar el ticket: ${error.message}` };
    }
}

/**
 * Exporta tickets filtrados a un archivo Excel (.xlsx) en Google Drive.
 */
function exportTicketsToExcelLogic(exportParams = {}) {
    try {
        const userProfile = getUserProfileLogic();
        if (userProfile.role !== 'SUPERUSER' && userProfile.role !== 'AUDITOR') {
            return { status: "error", message: "Acceso denegado: No tienes permisos para exportar datos." };
        }
        const sheet = getOrCreateSheet();
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return { status: "error", message: "No hay datos de tickets para exportar." };

        const headers = data[0];
        let filteredData = data.slice(1);

        if (exportParams.startDate) {
            const start = new Date(exportParams.startDate);
            filteredData = filteredData.filter(row => new Date(row[headers.indexOf('Timestamp')]) >= start);
        }
        if (exportParams.endDate) {
            const end = new Date(exportParams.endDate);
            end.setHours(23, 59, 59, 999);
            filteredData = filteredData.filter(row => new Date(row[headers.indexOf('Timestamp')]) <= end);
        }

        const selectedColumns = exportParams.columns || headers;
        const selectedHeaderIndices = selectedColumns.map(colName => headers.indexOf(colName)).filter(idx => idx !== -1);

        const exportData = [selectedColumns];
        filteredData.forEach(row => {
            exportData.push(selectedHeaderIndices.map(idx => row[idx]));
        });

        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const workbook = SpreadsheetApp.create(`Tickets_Export_${timestamp}`);
        workbook.getSheets()[0].getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);

        const file = DriveApp.getFileById(workbook.getId());
        const folder = DriveApp.getFolderById(SUPERUSER_DRIVE_FOLDER_ID);
        folder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
        const fileUrl = file.getUrl();

        Logger.log(`Exportación completada por ${userProfile.email}. Archivo: ${fileUrl}`);
        return { status: "success", fileUrl: fileUrl, message: "Tickets exportados a Excel correctamente." };
    } catch (error) {
        Logger.log(`ERROR al exportar a Excel: ${error.message}`);
        return { status: "error", message: `Error en el servidor al exportar a Excel: ${error.message}` };
    }
}

/**
 * Actualiza la prioridad de los tickets a "Urgente" si hay más de 5 tickets activos para una misma planilla.
 */
function updateTicketPrioritiesByVolume() {
    const sheet = getOrCreateSheet();
    if (sheet.getLastRow() <= 1) return;
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const planillaColIdx = headers.indexOf('Planilla (Cliente)');
    const estadoColIdx = headers.indexOf('Estado');
    const prioridadColIdx = headers.indexOf('Prioridad');
    const ticketInternoColIdx = headers.indexOf('Ticket Interno');
    if ([planillaColIdx, estadoColIdx, prioridadColIdx, ticketInternoColIdx].includes(-1)) return;

    const planillaCounts = data.reduce((acc, row) => {
        const planilla = row[planillaColIdx];
        const estado = row[estadoColIdx];
        if (planilla && ['Registrado', 'En Progreso'].includes(estado)) {
            if (!acc[planilla]) acc[planilla] = [];
            acc[planilla].push({
                rowNum: acc.length + 2, // approximation, better to pass index
                id: row[ticketInternoColIdx],
                currentPriority: row[prioridadColIdx] || 'Normal'
            });
        }
        return acc;
    }, {});

    for (const planilla in planillaCounts) {
        if (planillaCounts[planilla].length > 5) {
            planillaCounts[planilla].forEach((ticketInfo, index) => {
                if (ticketInfo.currentPriority !== 'Urgente') {
                    const realRowNum = data.findIndex(r => r[ticketInternoColIdx] === ticketInfo.id) + 2;
                    sheet.getRange(realRowNum, prioridadColIdx + 1).setValue('Urgente');
                    logTicketChange(ticketInfo.id, `Prioridad cambiada a 'Urgente' automáticamente por volumen.`);
                }
            });
        }
    }
}
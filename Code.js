/**
 * MentorPal v3.3 - Con Mood Meter RULER y Email CRM
 * Sistema con confirmación por email y formato listo para CRM
 * © 2025 Karen A. Guzmán V. - MentorIA Tools
 */

// ============ CONFIGURACIÓN DE SEGURIDAD ============
const SECURITY_CONFIG = {
  ALLOWED_DOMAINS: ['tec.mx', 'itesm.mx'],
  ADMIN_EMAILS: ['kareng@tec.mx'], // Agregar más admins aquí
  BACKUP_ENABLED: true,
  BACKUP_FOLDER_ID: '', // ID de carpeta para backups (opcional)
  MAX_FILES: 5,
  MAX_FILE_SIZE: 25 * 1024 * 1024, // 25 MB
  MAX_SESSION_DURATION: 180, // minutos
  CACHE_DURATION: 600, // seg
  TIMEZONE: 'America/Monterrey'
};

// ============ CONFIGURACIÓN DEL SISTEMA ============
const SYSTEM_CONFIG = {
  SHEET_CATALOG: 'CATALOGO_ESTUDIANTES',
  SHEET_SESSIONS: 'PROCESSED_DATA',
  SHEET_MENTORS: 'MENTORES_CATALOG',
  SHEET_HOME: '🏠 INICIO',
  SHEET_AUDIT: 'AUDIT_LOG',
  SHEET_CONFIG: 'CONFIG',
  DRIVE_FOLDER_NAME: 'MENTORPAL_EVIDENCIAS',
  APP_NAME: 'MentorPal',
  APP_VERSION: '3.3',
  APP_TAGLINE: 'Tu apoyo inteligente en cada mentoría',
  APP_COLOR: '#3B82F6',
  APP_ICON: '🤖'
};

// ============ TEMAS DE COMUNIDADES ============
const COMMUNITY_THEMES = {
  'Talenta':   { name: 'Talenta',   hex: '#EC008C' },
  'Revo':      { name: 'Revo',      hex: '#C4829A' },
  'Kresko':    { name: 'Kresko',    hex: '#0DCCCC' },
  'Pasio':     { name: 'Pasio',     hex: '#CC0202' },
  'Energio':   { name: 'Energio',   hex: '#FD8204' },
  'Krei':      { name: 'Krei',      hex: '#79858B' },
  'Reflekto':  { name: 'Reflekto',  hex: '#FFDE17' },
  'Forta':     { name: 'Forta',     hex: '#87004A' },
  'Spirita':   { name: 'Spirita',   hex: '#5B0F8B' },
  'Ekvilibro': { name: 'Ekvilibro', hex: '#6FD34A' }
};

// ============ MOOD METER CONFIG ============
const MOOD_METER = {
  'red': {
    name: 'Alta Energía - Desagradable',
    emoji: '🔴',
    emotions: ['Enojado', 'Frustrado', 'Estresado', 'Ansioso'],
    riskLevel: 8
  },
  'yellow': {
    name: 'Alta Energía - Agradable',
    emoji: '🟡',
    emotions: ['Alegre', 'Emocionado', 'Optimista', 'Motivado'],
    riskLevel: 3
  },
  'blue': {
    name: 'Baja Energía - Desagradable',
    emoji: '🔵',
    emotions: ['Triste', 'Deprimido', 'Aburrido', 'Agotado'],
    riskLevel: 6
  },
  'green': {
    name: 'Baja Energía - Agradable',
    emoji: '🟢',
    emotions: ['Calmado', 'Relajado', 'Sereno', 'En paz'],
    riskLevel: 2
  }
};

// ============ CATEGORÍAS CRM ============
const CRM_CATEGORIES = {
  'PIN': '🆕 PIN (Primeros Ingresos)',
  'CAG': '🎓 CAG (Candidatos a Graduación)',
  'Transferencias': '🔄 Transferencias',
  'Condicionamiento': '⚠️ Condicionamiento Académico',
  'Comites': '👥 Comités',
  'Trayectoria': '📈 Trayectoria Estudiantil',
  'Otro': '📌 Otro'
};

// ============ MENÚ EN SHEETS ============
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();

  try {
    const accessLevel = validateUserAccess(userEmail);
    if (accessLevel === 'DENIED') {
      ui.alert('⛔ Acceso Restringido', `${SYSTEM_CONFIG.APP_NAME} solo está disponible para usuarios autorizados del dominio @tec.mx`, ui.ButtonSet.OK);
      return;
    }

    const menu = ui.createMenu(`${SYSTEM_CONFIG.APP_ICON} ${SYSTEM_CONFIG.APP_NAME}`);
    menu.addItem('📝 Abrir Panel', 'openMentorPal');

    if (accessLevel === 'ADMIN') {
      menu.addSeparator();
      
      // Herramientas de Administración
      menu.addItem('🔧 Panel de Administración', 'showAdminPanelV32');
      menu.addItem('📊 Dashboard Completo', 'showEnhancedDashboard');
      menu.addItem('📋 Ver Audit Log', 'showAuditLog');
      
      menu.addSeparator();
      
      // Reportes y Emails
      menu.addItem('📧 Configurar resumen diario (6pm)', 'setupDailyReportTrigger');
      menu.addItem('📤 Enviar resumen de hoy (test)', 'testDailySummary');
      menu.addItem('📊 Generar reporte semanal (test)', 'sendWeeklyMentorSummary'); // MODIFICADO
      
      menu.addSeparator();
      
      // Mantenimiento
      menu.addItem('💾 Backup Manual', 'manualBackup');
      menu.addItem('📂 Exportar a CSV', 'exportSessionsToCSV');
      menu.addItem('📥 Exportar a Excel', 'exportSessionsToExcel');
      menu.addItem('🗄️ Archivar sesiones antiguas', 'archiveOldSessions');
      menu.addItem('✅ Validar integridad de datos', 'validateDataIntegrity');
      
      menu.addSeparator();
      menu.addItem('🔓 Mostrar Hojas Ocultas', 'showAllSheets');
    }

    menu.addSeparator();
    menu.addItem('❓ Ayuda', 'showHelp');
    menu.addToUi();

    setupSheetsVisibility(accessLevel);
    ensureHomeSheet();
    logAction('OPEN_SPREADSHEET', { accessLevel });
  } catch (error) {
    console.error('Error en onOpen:', error);
    ui.alert('Error', 'No se pudo inicializar MentorPal: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ============ ACCESO ============
function validateUserAccess(email) {
  if (!email) return 'DENIED';
  const domain = email.split('@')[1];
  if (SECURITY_CONFIG.ADMIN_EMAILS.includes(email)) return 'ADMIN';
  if (SECURITY_CONFIG.ALLOWED_DOMAINS.includes(domain)) return 'USER';
  return 'DENIED';
}

function validateApiAccess() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const accessLevel = validateUserAccess(email);
  if (accessLevel === 'DENIED') throw new Error('Acceso no autorizado');
  return { email, accessLevel };
}

// ============ APIs CON VALIDACIÓN ============
function apiGetMentors() {
  try {
    validateApiAccess();
    const cache = CacheService.getDocumentCache();
    const cacheKey = 'mentors:list:v3';
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
    if (!sh) { initializeMentorsSheet(); return []; }
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const ixName = headers.indexOf('NombreCompleto');
    const ixEmail = headers.indexOf('Email');
    const ixActive = headers.indexOf('Activo');

    const list = [];
    for (let i = 1; i < data.length; i++) {
      const activo = data[i][ixActive] === true || data[i][ixActive] === 'TRUE' || data[i][ixActive] === 'Sí';
      if (!activo) continue;
      list.push({ name: data[i][ixName] || '', email: data[i][ixEmail] || '' });
    }
    list.sort((a, b) => a.name.localeCompare(b.name, 'es'));

    cache.put(cacheKey, JSON.stringify(list), SECURITY_CONFIG.CACHE_DURATION);
    return list;
  } catch (error) {
    console.error('Error en apiGetMentors:', error);
    throw new Error('No se pudieron cargar los mentores');
  }
}

function apiLookupMatricula(matricula) {
  try {
    const { email } = validateApiAccess();
    matricula = normalizeMatricula(matricula);

    if (!/^A0\d{7,8}$|^A00\d{6}$/.test(matricula)) {
      return { found: false, themeHex: SYSTEM_CONFIG.APP_COLOR, themeName: 'Default' };
    }

    const cache = CacheService.getUserCache();
    const cacheKey = `lookup:${matricula}:${email}`;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    logAction('LOOKUP_STUDENT', { matricula: matricula.substring(0, 3) + '***' });

    const ss = SpreadsheetApp.getActive();
    const cat = ss.getSheetByName(SYSTEM_CONFIG.SHEET_CATALOG);
    if (!cat) { initializeStudentsSheet(); return { found: false, themeHex: SYSTEM_CONFIG.APP_COLOR, themeName: 'Default' }; }

    const data = cat.getDataRange().getValues();
    if (data.length <= 1) return { found: false, themeHex: SYSTEM_CONFIG.APP_COLOR, themeName: 'Default' };

    const headers = data[0];
    const ixM = headers.indexOf('Matricula');
    const ixN = headers.indexOf('NombreCompleto');
    const ixC = headers.indexOf('Comunidad');
    const ixP = headers.indexOf('ProgramaAcademico');
    const ixMentorName = headers.indexOf('MentorAsignado');

    for (let i = 1; i < data.length; i++) {
      if (normalizeMatricula(data[i][ixM]) === matricula) {
        const mentorAsignado = (data[i][ixMentorName] || '').toString().trim();
        const mentorEmail = getMentorEmailByName(mentorAsignado);

        if (email && mentorEmail && email !== mentorEmail.toLowerCase()) {
          const accessLevel = validateUserAccess(email);
          if (accessLevel !== 'ADMIN') {
            return { found: false, themeHex: SYSTEM_CONFIG.APP_COLOR, themeName: 'Default' };
          }
        }

        const comunidad = data[i][ixC] || '';
        const result = {
          found: true,
          nombre: data[i][ixN] || '',
          comunidad,
          programa: data[i][ixP] || '',
          mentorAsignado,
          themeHex: COMMUNITY_THEMES[comunidad]?.hex || SYSTEM_CONFIG.APP_COLOR,
          themeName: COMMUNITY_THEMES[comunidad]?.name || 'Default'
        };

        cache.put(cacheKey, JSON.stringify(result), 300);
        return result;
      }
    }

    return { found: false, themeHex: SYSTEM_CONFIG.APP_COLOR, themeName: 'Default' };
  } catch (error) {
    console.error('Error en apiLookupMatricula:', error);
    throw new Error('Error al buscar estudiante');
  }
}

function apiSaveSession(payload) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(5000);

    const { email } = validateApiAccess();
    const [mentorName, mentorEmail] = (payload.mentor || '').split('|');
    if (!mentorName) throw new Error('Selecciona Mentor/a.');

    if (email && mentorEmail && email !== mentorEmail.toLowerCase()) {
      const accessLevel = validateUserAccess(email);
      if (accessLevel !== 'ADMIN') throw new Error('No puedes registrar sesiones para otra persona.');
    }

    payload.mentor = mentorName;
    payload.mentorEmail = mentorEmail; // Guardar email para envío

    const required = ['mentor', 'matricula', 'tipo', 'duracion', 'riesgo', 'fecha', 'resumen', 'moodState', 'categoria'];
    for (const k of required) { 
      if (!payload[k]) throw new Error(`Falta el campo obligatorio: ${k}`); 
    }

    payload.resumen = cleanText(payload.resumen);
    payload.notas = cleanText(payload.notas);
    if (payload.resumen.length < 50) throw new Error('El resumen debe tener al menos 50 caracteres.');

    const duracion = Number(payload.duracion) || 0;
    if (duracion > SECURITY_CONFIG.MAX_SESSION_DURATION) throw new Error(`La duración máxima es ${SECURITY_CONFIG.MAX_SESSION_DURATION} minutos.`);

    if (payload.seguimiento && payload.seguimiento.includes('Consejería') && Number(payload.riesgo) < 4) {
      throw new Error('Para derivación a Consejería, el nivel debe ser ≥ 4.');
    }

    payload.matricula = normalizeMatricula(payload.matricula);
    const est = _ensureStudent(payload);

    const sessionId = 'SES_' + Date.now() + '_' + Utilities.getUuid().substring(0, 8);
    const now = new Date();
    const dimensiones = (payload.dimensiones || []).join(', ');
    const evidencias = (payload.evidencias || []).join(' | ');

    const row = [
      sessionId, now, payload.mentor, est.id, payload.matricula,
      payload.nombre || est.nombre, _toDateLocal(payload.fecha), payload.tipo,
      _duracionToMinutes(payload.duracion), dimensiones, payload.resumen,
      payload.compEst || '', payload.compMentor || '', payload.seguimiento || '',
      Number(payload.riesgo) || '', _toDateLocal(payload.proxima),
      (payload.notas || '') + (evidencias ? `\nEvidencia: ${evidencias}` : ''),
      'Pendiente', // CRM_Documentado siempre empieza como Pendiente
      payload.moodState || '', // Guardar mood state
      payload.categoria || '', // Guardar categoría
      `${SYSTEM_CONFIG.APP_NAME} v${SYSTEM_CONFIG.APP_VERSION}`
    ];

    let sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    if (!sh) { initializeProcessedSheet(); sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS); }
    sh.appendRow(row);

    // Enviar email de confirmación con formato CRM
    try {
      sendCRMEmail(payload, sessionId);
    } catch (emailError) {
      console.error('Error enviando email:', emailError);
      // No fallar la operación si el email falla
    }

    logAction('SAVE_SESSION', { 
      sessionId, 
      mentor: mentorName, 
      riskLevel: payload.riesgo, 
      moodState: payload.moodState,
      categoria: payload.categoria,
      duration: duracion 
    });

    // Limpiar caches de historial (alcances 'mine' y 'all')
    const uc = CacheService.getUserCache();
    uc.remove('mentor:history');
    uc.remove('mentor:history:v3:mine');
    uc.remove('mentor:history:v3:all');

    if (Number(payload.riesgo) >= 7) {
      SpreadsheetApp.getActive().toast(
        `🔴 Caso crítico registrado - Nivel ${payload.riesgo}/10`,
        SYSTEM_CONFIG.APP_NAME,
        5
      );
      logAction('CRITICAL_CASE', { sessionId, level: payload.riesgo });
    }

    return { ok: true, sessionId };
  } catch (error) {
    console.error('[MentorPal Error]', error);
    logAction('ERROR_SAVE_SESSION', { error: error.toString() });
    throw new Error(error.message || 'Error al guardar. Intenta de nuevo.');
  } finally {
    lock.releaseLock();
  }
}
// ============ FUNCIÓN sendCRMEmail() COMPLETA MEJORADA v3.2.4 ============
// Versión con resumen completo e instrucciones claras

function sendCRMEmail(payload, sessionId) {
  const mentorEmail = payload.mentorEmail || getMentorEmailByName(payload.mentor);
  if (!mentorEmail) return;

  // Preparar información básica
  const categoria = CRM_CATEGORIES[payload.categoria] || payload.categoria || 'General';
  const moodInfo = MOOD_METER[payload.moodState] || { emoji: '⚪', name: 'No especificado' };
  const dimensionesTexto = (payload.dimensiones || []).join(', ') || 'No especificadas';
  
  // Asunto sugerido para CRM
  const asuntoSugerido = `${categoria} - ${payload.nombre || 'Estudiante'} - ${_fmtDateTZ(_toDateLocal(payload.fecha), 'dd/MM/yyyy')}`;
  
  // Descripción COMPLETA para CRM
  const descripcionCRM = `SESIÓN DE MENTORÍA - MentorPal

INFORMACIÓN BÁSICA:
- Fecha: ${_fmtDateTZ(_toDateLocal(payload.fecha), 'dd/MM/yyyy')}
- Tipo: ${payload.tipo}
- Duración: ${payload.duracion} minutos
- Categoría: ${categoria}

ESTUDIANTE:
- Nombre: ${payload.nombre || 'Por confirmar'}
- Matrícula: ${payload.matricula}
- Programa: ${payload.programa || 'Por confirmar'}
- Comunidad: ${payload.comunidad || 'Por confirmar'}

ESTADO EMOCIONAL (Mood Meter RULER):
- ${moodInfo.emoji} ${moodInfo.name}
- Nivel de atención: ${payload.riesgo}/10

DIMENSIONES DEL BIENESTAR ABORDADAS:
${dimensionesTexto}

RESUMEN DE LA SESIÓN:
${payload.resumen}

COMPROMISOS:
- Del estudiante: ${payload.compEst || 'Sin compromisos específicos'}
- Del mentor/a: ${payload.compMentor || 'Seguimiento regular'}

SEGUIMIENTO:
${payload.seguimiento || 'No requiere seguimiento especial'}

${payload.notas ? 'NOTAS ADICIONALES:\n' + payload.notas : ''}

${payload.proxima ? 'PRÓXIMA SESIÓN SUGERIDA: ' + _fmtDateTZ(_toDateLocal(payload.proxima), 'dd/MM/yyyy') : ''}

---
Registrado por: ${payload.mentor}
Sistema: MentorPal v${SYSTEM_CONFIG.APP_VERSION}
ID Sesión: ${sessionId}`;

  // Renderizar email desde plantilla con color de comunidad
  const comunidad = getMentorCommunityByName(payload.mentor) || getMentorCommunityByEmail(mentorEmail);
  const theme = getThemeForCommunity(comunidad);
  const t = HtmlService.createTemplateFromFile('sendCRMEmail');
  t.asuntoSugerido = asuntoSugerido;
  t.descripcionCRM = descripcionCRM.replace(/\n/g, '<br>');
  t.payload = payload;
  t.moodInfo = moodInfo;
  t.categoria = categoria;
  t.themeHex = theme.hex;
  t.appVersion = SYSTEM_CONFIG.APP_VERSION;
  const htmlBody = t.evaluate().getContent();

  // Enviar el email
  MailApp.sendEmail({
    to: mentorEmail,
    subject: `✅ Confirmación de sesión - ${payload.nombre || payload.matricula}`,
    htmlBody: htmlBody
  });

  logAction('CRM_EMAIL_SENT', { 
    sessionId: sessionId,
    mentor: mentorEmail,
    categoria: payload.categoria 
  });
}

/**
 * Enviar resumen diario a cada mentor (ejecutar a las 6pm)
 */
function sendDailyMentorSummary() {
  try {
    const todayKey = Utilities.formatDate(new Date(), SECURITY_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    
    if (!sh || sh.getLastRow() <= 1) {
      console.log('No hay datos para el resumen diario');
      return;
    }
    
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    
    // Índices de columnas
    const indices = {
      date: headers.indexOf('FechaSesion'),
      mentor: headers.indexOf('Mentor'),
      student: headers.indexOf('NombreEstudiante'),
      matricula: headers.indexOf('Matricula'),
      risk: headers.indexOf('NivelRiesgo'),
      mood: headers.indexOf('AI_Sentiment'),
      category: headers.indexOf('AI_RiskPrediction'),
      summary: headers.indexOf('ResumenTemas'),
      followUp: headers.indexOf('AccionSeguimiento'),
      crm: headers.indexOf('CRM_Documentado')
    };
    
    // Agrupar sesiones por mentor
    const mentorSessions = {};
    
    for (let i = 1; i < data.length; i++) {
      const sessionKey = Utilities.formatDate(new Date(data[i][indices.date]), SECURITY_CONFIG.TIMEZONE, 'yyyy-MM-dd');
      if (sessionKey === todayKey) {
        const mentorName = data[i][indices.mentor];
        if (!mentorName) continue;
        
        if (!mentorSessions[mentorName]) {
          mentorSessions[mentorName] = [];
        }
        
        mentorSessions[mentorName].push({
          student: data[i][indices.student],
          matricula: data[i][indices.matricula],
          risk: Number(data[i][indices.risk]) || 0,
          mood: data[i][indices.mood],
          category: data[i][indices.category],
          summary: (data[i][indices.summary] || '').substring(0, 100), // Primeros 100 caracteres
          followUp: data[i][indices.followUp],
          crm: data[i][indices.crm]
        });
      }
    }
    // Guard clause: no enviar si no hubo sesiones hoy
  if (Object.keys(mentorSessions).length === 0) {
  console.log('No hay sesiones hoy. No se enviará resumen diario.');
  logAction('DAILY_SUMMARY_SKIPPED', { reason: 'no_sessions_today' });
  return;
  }
    // Enviar email a cada mentor con sesiones del día
    let emailsSent = 0;
    
    for (const [mentorName, sessions] of Object.entries(mentorSessions)) {
      const mentorEmail = getMentorEmailByName(mentorName);
      if (!mentorEmail) continue;
      
      const today = new Date();
      const htmlContent = generateDailySummaryHTML(mentorName, sessions, today);
      
      MailApp.sendEmail({
        to: mentorEmail,
        subject: `📊 Tu resumen del día - ${sessions.length} ${sessions.length === 1 ? 'sesión' : 'sesiones'} registradas`,
        htmlBody: htmlContent
      });
      
      emailsSent++;
    }
    
    logAction('DAILY_SUMMARIES_SENT', { 
      mentorsWithSessions: Object.keys(mentorSessions).length,
      emailsSent: emailsSent,
      date: new Date().toLocaleDateString('es-MX')
    });
    
    console.log(`Resúmenes diarios enviados: ${emailsSent}`);
    
  } catch (error) {
    console.error('Error enviando resúmenes diarios:', error);
    logAction('ERROR_DAILY_SUMMARY', { error: error.toString() });
  }
}
/**
 * Generar HTML del resumen diario personal
 */
function generateDailySummaryHTML(mentorName, sessions, date) {
  const criticalCount = sessions.filter(s => s.risk >= 7).length;
  const pendingCRM = sessions.filter(s => s.crm === 'Pendiente' || !s.crm).length;
  const moodEmojis = { red: '🔴', yellow: '🟡', blue: '🔵', green: '🟢' };
  const formattedDate = _fmtDateTZ(date, "EEEE d 'de' MMMM");
  const community = getMentorCommunityByName(mentorName);
  const theme = getThemeForCommunity(community);
  const t = HtmlService.createTemplateFromFile('generateDailySummaryHTML');
  t.mentorName = mentorName;
  t.sessions = sessions;
  t.criticalCount = criticalCount;
  t.pendingCRM = pendingCRM;
  t.moodEmojis = moodEmojis;
  t.formattedDate = formattedDate;
  t.themeHex = theme.hex;
  t.appVersion = SYSTEM_CONFIG.APP_VERSION;
  return t.evaluate().getContent();
}

// ============ CONFIGURAR TRIGGER PARA REPORTE DIARIO ============

/**
 * Configurar el envío automático del resumen diario a las 6pm
 */
function setupDailyReportTrigger() {
  // Eliminar triggers existentes con el mismo nombre
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyMentorSummary') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger para las 6pm todos los días
  ScriptApp.newTrigger('sendDailyMentorSummary')
    .timeBased()
    .everyDays(1)
    .atHour(18) // 6 PM
    .create();
  
  SpreadsheetApp.getActive().toast(
    '✅ Resumen diario configurado para las 6:00 PM',
    'MentorPal',
    5
  );
  
  logAction('DAILY_TRIGGER_SETUP', { 
    time: '18:00',
    frequency: 'daily'
  });
}

/**
 * Ejecutar manualmente el resumen del día (para pruebas)
 */
function testDailySummary() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '📧 Enviar Resumen del Día',
    '¿Deseas enviar el resumen del día a todos los mentores con sesiones hoy?\n\n' +
    'Esto enviará un email a cada mentor con sus sesiones del día.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    sendDailyMentorSummary();
    ui.alert('✅ Resúmenes enviados', 'Los resúmenes diarios han sido enviados a los mentores con sesiones hoy.', ui.ButtonSet.OK);
  }
}

// ============ FUNCIÓN HELPER PARA OBTENER EMAIL POR NOMBRE ============
// (Esta función ya debería existir, pero la incluyo por si acaso)

function getMentorEmailByName(name) {
  if (!name) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return '';
  const headers = data[0];
  const ixName = headers.indexOf('NombreCompleto');
  const ixEmail = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ixName] || '').toString().trim() === name) {
      return (data[i][ixEmail] || '').toString().trim().toLowerCase();
    }
  }
  return '';
}

// ============ THEME HELPERS POR COMUNIDAD ============
function getMentorCommunityByEmail(email) {
  if (!email) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return '';
  const headers = data[0];
  const ixEmail = headers.indexOf('Email');
  const ixComu  = headers.indexOf('Comunidad');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ixEmail] || '').toString().trim().toLowerCase() === String(email).toLowerCase()) {
      return (data[i][ixComu] || '').toString().trim();
    }
  }
  return '';
}

function getMentorCommunityByName(name) {
  if (!name) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return '';
  const headers = data[0];
  const ixName = headers.indexOf('NombreCompleto');
  const ixComu = headers.indexOf('Comunidad');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ixName] || '').toString().trim() === name) {
      return (data[i][ixComu] || '').toString().trim();
    }
  }
  return '';
}

function getThemeForCommunity(comunidad) {
  const theme = COMMUNITY_THEMES[comunidad];
  return theme ? theme : { name: 'Default', hex: SYSTEM_CONFIG.APP_COLOR };
}

// ============ ACTUALIZACIÓN DEL MENÚ PARA ADMIN ============

/**
 * Agregar opciones al menú de administrador
 * (Agregar estas líneas en la función onOpen() existente, en la sección de ADMIN)
 */
function updateAdminMenu() {
  // En la función onOpen(), dentro del if (accessLevel === 'ADMIN'), agregar:
  // menu.addItem('📧 Configurar resumen diario', 'setupDailyReportTrigger');
  // menu.addItem('📤 Enviar resumen de hoy (test)', 'testDailySummary');
}

// ============ FIN DE FUNCIONES DE EMAIL MEJORADAS v3.2.1 ============

function apiGetMentorHistory(viewAll) {
  try {
    const { email, accessLevel } = validateApiAccess();

    const cache = CacheService.getUserCache();
    const scope = (viewAll === true && accessLevel === 'ADMIN') ? 'all' : 'mine';
    const cached = cache.get('mentor:history:v3:' + scope);
    if (cached) return JSON.parse(cached);

    const mentorName = getMentorNameByEmail(email);
    const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    if (!sh) return [];

    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const ixId   = headers.indexOf('Session_ID');
    const ixMent = headers.indexOf('Mentor');
    const ixFecha= headers.indexOf('FechaSesion');
    const ixNom  = headers.indexOf('NombreEstudiante');
    const ixMat  = headers.indexOf('Matricula');
    const ixRisk = headers.indexOf('NivelRiesgo');
    const ixSeg  = headers.indexOf('AccionSeguimiento');
    const ixCRM  = headers.indexOf('CRM_Documentado');
    const ixMood = headers.indexOf('AI_Sentiment'); // Reutilizamos esta columna para mood
    const ixCat  = headers.indexOf('AI_RiskPrediction'); // Reutilizamos esta columna para categoría

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const isMine = (data[i][ixMent] || '') === mentorName;
      const include = (scope === 'all') ? true : isMine;
      if (include) {
        rows.push({
          sessionId: data[i][ixId] || '',
          fecha: data[i][ixFecha],
          nombre: data[i][ixNom] || '',
          matricula: data[i][ixMat] || '',
          riesgo: data[i][ixRisk] || '',
          seguimiento: data[i][ixSeg] || '',
          crmSaved: data[i][ixCRM] || 'Pendiente',
          moodState: data[i][ixMood] || '',
          categoria: data[i][ixCat] || ''
        });
      }
    }

    const result = rows.slice(-15).reverse();
    cache.put('mentor:history:v3:' + scope, JSON.stringify(result), 120);
    return result;
  } catch (error) {
    console.error('[History Error]', error);
    return [];
  }
}

function apiGetMentorStats() {
  try {
    const { email } = validateApiAccess();
    const mentorName = getMentorNameByEmail(email);
    const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    if (!sh || sh.getLastRow() <= 1) {
      return { totalSessions: 0, weekSessions: 0, pendingCRM: 0, crmRate: 0, uniqueStudents: 0, weeklyAvg: 0, criticalCases: 0, avgRisk: 0 };
    }

    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const ixMentor = headers.indexOf('Mentor');
    const ixFecha = headers.indexOf('FechaSesion');
    const ixCRM = headers.indexOf('CRM_Documentado');
    const ixMatricula = headers.indexOf('Matricula');
    const ixRisk = headers.indexOf('NivelRiesgo');

    let totalSessions = 0, weekSessions = 0, pendingCRM = 0, documentedCRM = 0, criticalCases = 0, riskSum = 0;
    const uniqueStudents = new Set();
    const now = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    for (let i = 1; i < data.length; i++) {
      if (data[i][ixMentor] === mentorName) {
        totalSessions++;
        const d = new Date(data[i][ixFecha]);
        if (d >= weekAgo) weekSessions++;
        const risk = Number(data[i][ixRisk]) || 0;
        riskSum += risk;
        if (risk >= 7) criticalCases++;
        if ((data[i][ixCRM] || '') === 'Pendiente' || !data[i][ixCRM]) pendingCRM++;
        if ((data[i][ixCRM] || '') === 'Sí') documentedCRM++;
        if (data[i][ixMatricula]) uniqueStudents.add(data[i][ixMatricula]);
      }
    }

    const crmRate = totalSessions > 0 ? Math.round((documentedCRM / totalSessions) * 100) : 0;
    const weeklyAvg = totalSessions > 0 ? Math.round(totalSessions / 4) : 0;
    const avgRisk = totalSessions > 0 ? +(riskSum / totalSessions).toFixed(1) : 0;

    return {
      totalSessions,
      weekSessions,
      pendingCRM,
      crmRate,
      uniqueStudents: uniqueStudents.size,
      weeklyAvg,
      criticalCases,
      avgRisk
    };
  } catch (error) {
    console.error('[Stats Error]', error);
    return { totalSessions: 0, weekSessions: 0, pendingCRM: 0, crmRate: 0, uniqueStudents: 0, weeklyAvg: 0, criticalCases: 0, avgRisk: 0 };
  }
}

function apiSetCrmStatus(sessionId, crmSaved) {
  try {
    const { email } = validateApiAccess();
    const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const ixId = headers.indexOf('Session_ID');
    const ixCRM = headers.indexOf('CRM_Documentado');
    const ixMentor = headers.indexOf('Mentor');

    for (let i = 1; i < data.length; i++) {
      if (data[i][ixId] === sessionId) {
        const mentorName = getMentorNameByEmail(email);
        if (data[i][ixMentor] !== mentorName && validateUserAccess(email) !== 'ADMIN') {
          throw new Error('No puedes actualizar sesiones de otra persona');
        }
        sh.getRange(i + 1, ixCRM + 1).setValue(crmSaved ? 'Sí' : 'Pendiente');
        logAction('CRM_STATUS_UPDATE', { sessionId, status: crmSaved ? 'Documented' : 'Pending' });
        const uc = CacheService.getUserCache();
        uc.remove('mentor:history:v3:mine');
        uc.remove('mentor:history:v3:all');
        return { ok: true, message: crmSaved ? 'Marcado en CRM' : 'Pendiente en CRM' };
      }
    }
    throw new Error('Sesión no encontrada');
  } catch (error) {
    console.error('[CRM Update Error]', error);
    throw new Error(error.message || 'Error al actualizar estado CRM');
  }
}

// ============ HELPERS ============
function normalizeMatricula(m) {
  return (m || '').toUpperCase().replace(/[^A-Z0-9]/g, '').trim();
}

function cleanText(s) {
  return (s || '').replace(/[\u0000-\u001F\u007F]/g, ' ').trim().substring(0, 5000);
}

function _ensureStudent(payload) {
  const ss = SpreadsheetApp.getActive();
  let cat = ss.getSheetByName(SYSTEM_CONFIG.SHEET_CATALOG);
  if (!cat) { initializeStudentsSheet(); cat = ss.getSheetByName(SYSTEM_CONFIG.SHEET_CATALOG); }

  const values = cat.getDataRange().getValues();
  if (values.length <= 1) return createNewStudent(cat, payload);

  const headers = values[0];
  const ixId = headers.indexOf('Estudiante_ID');
  const ixM  = headers.indexOf('Matricula');
  const ixN  = headers.indexOf('NombreCompleto');

  for (let i = 1; i < values.length; i++) {
    if (normalizeMatricula(values[i][ixM]) === payload.matricula) {
      return { id: values[i][ixId], nombre: values[i][ixN] || payload.nombre || '' };
    }
  }
  return createNewStudent(cat, payload);
}

function createNewStudent(sheet, payload) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const id = 'EST_' + payload.matricula + '_' + Date.now();
  const now = new Date();
  const nueva = Array(headers.length).fill('');

  const ix = {
    id: headers.indexOf('Estudiante_ID'),
    mat: headers.indexOf('Matricula'),
    nom: headers.indexOf('NombreCompleto'),
    prog: headers.indexOf('ProgramaAcademico'),
    com: headers.indexOf('Comunidad'),
    ment: headers.indexOf('MentorAsignado'),
    fecha: headers.indexOf('FechaIngreso'),
    est: headers.indexOf('Estatus'),
    ses: headers.indexOf('SesionesTotales'),
    ult: headers.indexOf('UltimaSesion'),
    risk: headers.indexOf('PromedioRiesgo'),
    ia: headers.indexOf('IA_Profile')
  };

  nueva[ix.id] = id;
  nueva[ix.mat] = payload.matricula;
  nueva[ix.nom] = payload.nombre || '(pendiente)';
  nueva[ix.prog] = payload.programa || 'Por determinar';
  nueva[ix.com] = payload.comunidad || 'Por asignar';
  nueva[ix.ment] = payload.mentor || '';
  nueva[ix.fecha] = now;
  nueva[ix.est] = 'Nuevo';
  nueva[ix.ses] = 1;
  nueva[ix.ult] = now;
  nueva[ix.risk] = payload.riesgo || '';
  nueva[ix.ia] = '{}';

  sheet.appendRow(nueva);
  logAction('NEW_STUDENT', { id, matricula: payload.matricula.substring(0, 3) + '***' });
  return { id, nombre: nueva[ix.nom] };
}

function getMentorNameByEmail(email) {
  if (!email) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return '';
  const headers = data[0];
  const ixName = headers.indexOf('NombreCompleto');
  const ixEmail = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ixEmail] || '').toString().trim().toLowerCase() === email) {
      return (data[i][ixName] || '').toString().trim();
    }
  }
  return '';
}

function getMentorEmailByName(name) {
  if (!name) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return '';
  const headers = data[0];
  const ixName = headers.indexOf('NombreCompleto');
  const ixEmail = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ixName] || '').toString().trim() === name) {
      return (data[i][ixEmail] || '').toString().trim().toLowerCase();
    }
  }
  return '';
}

function _toDateLocal(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  // Acepta 'YYYY-MM-DD' desde el input <input type="date">
  var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); // local midnight
  var d = new Date(v);
  return isNaN(d.getTime()) ? '' : d;
}

/**
 * Formatea una fecha usando la zona horaria definida y patrón simple.
 * Soporta tokens: yyyy, MM, dd, HH, mm, EEE, EEEE, MMM, MMMM
 * Ejemplos:
 * _fmtDateTZ(new Date(), 'dd/MM/yyyy')
 * _fmtDateTZ(new Date(), "EEEE d 'de' MMMM yyyy")
 */
function _fmtDateTZ(date, pattern) {
  if (!date) return '';
  const tz = SECURITY_CONFIG.TIMEZONE || Session.getScriptTimeZone() || 'America/Mexico_City';
  const d = (date instanceof Date) ? date : _toDateLocal(date);
  const p = pattern || 'dd/MM/yyyy';

  // Partes numéricas con TZ (Utilities sigue el TZ tal cual)
  const parts = {
    'yyyy': Utilities.formatDate(d, tz, 'yyyy'),
    'MM'  : Utilities.formatDate(d, tz, 'MM'),
    'dd'  : Utilities.formatDate(d, tz, 'dd'),
    'HH'  : Utilities.formatDate(d, tz, 'HH'),
    'mm'  : Utilities.formatDate(d, tz, 'mm')
  };

  // Nombres en español con TZ (Intl respeta locale y TZ)
  const weekdayLong = new Intl.DateTimeFormat('es-MX', { timeZone: tz, weekday: 'long' }).format(d);
  const weekdayShort= new Intl.DateTimeFormat('es-MX', { timeZone: tz, weekday: 'short' }).format(d);
  const monthLong   = new Intl.DateTimeFormat('es-MX', { timeZone: tz, month: 'long'   }).format(d);
  const monthShort  = new Intl.DateTimeFormat('es-MX', { timeZone: tz, month: 'short'  }).format(d);

  const names = {
    'EEEE': weekdayLong,
    'EEE' : weekdayShort,
    'MMMM': monthLong,
    'MMM' : monthShort
  };

  // Reemplazar en orden de mayor a menor longitud para evitar solapes
  const tokens = ['EEEE','EEE','MMMM','MMM','yyyy','MM','dd','HH','mm'];
  let out = p;
  tokens.forEach(t => { out = out.replace(t, (names[t] || parts[t] || t)); });
  return out;
}


function _duracionToMinutes(key) {
  const map = { '15': 15, '16-30': 30, '31-45': 45, '46-60': 60, '60+': 90 };
  return map[key] || '';
}

function _ensureFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  const folder = DriveApp.createFolder(name);
  folder.setDescription(`Evidencias digitales - ${SYSTEM_CONFIG.APP_NAME} v${SYSTEM_CONFIG.APP_VERSION}`);
  const now = new Date();
  const year = now.getFullYear().toString();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  folder.createFolder(year).createFolder(month);
  return folder;
}

// ============ INICIALIZACIÓN DE HOJAS ============
function initializeMentorsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(SYSTEM_CONFIG.SHEET_MENTORS);
  const headers = ['Mentor_ID','NombreCompleto','Email','Comunidad','Activo','FechaIngreso','Telefono','TipoMentor','EstudiantesAsignados','SesionesTotales','UltimaSesion'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground(SYSTEM_CONFIG.APP_COLOR).setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!SECURITY_CONFIG.ADMIN_EMAILS.includes(userEmail)) sheet.hideSheet();
}

function initializeStudentsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(SYSTEM_CONFIG.SHEET_CATALOG);
  const headers = ['Estudiante_ID','Matricula','NombreCompleto','ProgramaAcademico','Comunidad','MentorAsignado','FechaIngreso','Estatus','SesionesTotales','UltimaSesion','PromedioRiesgo','EmailEstudiante','Telefono','Semestre','IA_Profile'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground(SYSTEM_CONFIG.APP_COLOR).setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!SECURITY_CONFIG.ADMIN_EMAILS.includes(userEmail)) sheet.hideSheet();
}

function initializeProcessedSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(SYSTEM_CONFIG.SHEET_SESSIONS);
  const headers = [
    'Session_ID','Timestamp','Mentor','Estudiante_ID','Matricula','NombreEstudiante','FechaSesion','TipoSesion',
    'Duracion','Dimensiones','ResumenTemas','CompromisosEst','CompromisosOCompania','AccionSeguimiento','NivelRiesgo',
    'ProximaSesion','NotasAdicionales','CRM_Documentado','AI_Sentiment','AI_RiskPrediction','ProcessedBy'
  ];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground(SYSTEM_CONFIG.APP_COLOR).setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!SECURITY_CONFIG.ADMIN_EMAILS.includes(userEmail)) sheet.hideSheet();
}

// ============ WEB APP ============
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('MentorPal - Sistema de Mentoría')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function validateWebAccess() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { authorized: false, message: 'No se pudo identificar usuario' };

    const domain = email.split('@')[1];
    const isAdmin = SECURITY_CONFIG.ADMIN_EMAILS.includes(email.toLowerCase());
    const isAllowedDomain = SECURITY_CONFIG.ALLOWED_DOMAINS.includes(domain);

    if (!isAllowedDomain && !isAdmin) return { authorized: false, message: 'Dominio no autorizado' };

    logAction('WEB_APP_ACCESS', { email, domain });

    return { authorized: true, email, role: isAdmin ? 'Administrador' : 'Mentor', domain };
  } catch (error) {
    console.error('Error validando acceso:', error);
    return { authorized: false, message: error.toString() };
  }
}

// Wrapper para WebApp
function webExportToExcel(viewAll) { 
  return exportMentorSessionsToExcel(viewAll === true); 
}

// Wrappers Web ↔️ Server
function webGetMentors()        { return apiGetMentors(); }
function webLookupStudent(m)    { return apiLookupMatricula(m); }
function webSaveSession(p)      { const acc = validateWebAccess(); if (!acc.authorized) throw new Error('Acceso no autorizado'); return apiSaveSession(p); }
function webGetHistory(v)       { return apiGetMentorHistory(v === true); }
function webGetStats()          { return apiGetMentorStats(); }
function webSetCrmStatus(id, b) { return apiSetCrmStatus(id, b); }
function webUploadEvidence(f)   { const acc = validateWebAccess(); if (!acc.authorized) throw new Error('Acceso no autorizado'); return apiUploadEvidence(f); }
function webGetAccessInfo()     { return validateWebAccess(); }
function webGetMyTheme() {
  const acc = validateWebAccess();
  if (!acc || !acc.authorized) return { hex: SYSTEM_CONFIG.APP_COLOR, name: 'Default' };
  const community = getMentorCommunityByEmail(acc.email);
  const theme = getThemeForCommunity(community);
  return { hex: theme.hex, name: community || theme.name };
}

// ============ FUNCIONES AUXILIARES DE ADMINISTRACIÓN ============
function showHelp() {
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const community = getMentorCommunityByEmail(userEmail);
  const theme = getThemeForCommunity(community);
  const themeHex = theme.hex || SYSTEM_CONFIG.APP_COLOR;
  const html = `
    <div style="font-family:'Segoe UI',Arial,sans-serif;padding:20px;">
      <h2 style="color:${themeHex};">${SYSTEM_CONFIG.APP_ICON} ${SYSTEM_CONFIG.APP_NAME}</h2>
      <p style="color:#666;font-style:italic;">${SYSTEM_CONFIG.APP_TAGLINE}</p>
      <h3>🚀 Inicio Rápido:</h3>
      <ol>
        <li>Menú ${SYSTEM_CONFIG.APP_NAME} → Abrir Panel</li>
        <li>Busca por matrícula</li>
        <li>Selecciona el estado emocional (Mood Meter)</li>
        <li>Completa y guarda</li>
      </ol>
      <h3>📧 Email de Confirmación:</h3>
      <p>Recibirás un email con:</p>
      <ul>
        <li>Asunto sugerido para CRM</li>
        <li>Descripción completa lista para copiar</li>
        <li>Instrucciones de documentación</li>
      </ul>
      <p style="text-align:center;margin-top:20px;">
        <button onclick="google.script.host.close()" style="padding:10px 20px;background:${themeHex};color:#fff;border:none;border-radius:5px;cursor:pointer;">Cerrar</button>
      </p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(550), `Ayuda - ${SYSTEM_CONFIG.APP_NAME}`);
}

function openMentorPal() {
  try {
    const { email, accessLevel } = validateApiAccess();
    logAction('OPEN_MODAL', { email });
    const html = HtmlService.createHtmlOutputFromFile('WebApp')
      .setWidth(980)
      .setHeight(750)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle(`${SYSTEM_CONFIG.APP_NAME} - ${SYSTEM_CONFIG.APP_TAGLINE}`);
    SpreadsheetApp.getUi().showModalDialog(html, SYSTEM_CONFIG.APP_NAME);
  } catch (error) {
    console.error('Error abriendo MentorPal:', error);
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function setupSheetsVisibility(accessLevel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sensitiveSheets = [
    SYSTEM_CONFIG.SHEET_CATALOG,
    SYSTEM_CONFIG.SHEET_SESSIONS,
    SYSTEM_CONFIG.SHEET_MENTORS,
    SYSTEM_CONFIG.SHEET_AUDIT,
    SYSTEM_CONFIG.SHEET_CONFIG
  ];

  if (accessLevel !== 'ADMIN') {
    sensitiveSheets.forEach(name => {
      try {
        const sh = ss.getSheetByName(name);
        if (sh && !sh.isSheetHidden()) sh.hideSheet();
      } catch (e) { console.log('No se pudo ocultar', name, e); }
    });
    const home = ss.getSheetByName(SYSTEM_CONFIG.SHEET_HOME);
    if (home && home.isSheetHidden()) home.showSheet();
    if (home) ss.setActiveSheet(home);
  }
}

function ensureHomeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_HOME);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_CONFIG.SHEET_HOME);
    sheet.getRange('B2').setValue(SYSTEM_CONFIG.APP_NAME).setFontSize(24).setFontWeight('bold').setFontColor(SYSTEM_CONFIG.APP_COLOR);
    sheet.getRange('B3').setValue(SYSTEM_CONFIG.APP_TAGLINE).setFontSize(14).setFontColor('#666');
    sheet.getRange('B5').setValue('Bienvenido/a al sistema de registro de mentorías').setFontSize(12);
    sheet.getRange('B7').setValue('📝 Para registrar una sesión:').setFontSize(12).setFontWeight('bold');
    sheet.getRange('B8').setValue('1. Menú "' + SYSTEM_CONFIG.APP_NAME + '" → Abrir Panel').setFontSize(11);
    sheet.getRange('B9').setValue('2. Completa el formulario').setFontSize(11);
    sheet.getRange('B12').setValue('✅ Nuevas características v3.2:').setFontSize(12).setFontWeight('bold');
    sheet.getRange('B13').setValue('• Mood Meter (RULER Approach)');
    sheet.getRange('B14').setValue('• Email con formato para CRM');
    sheet.getRange('B15').setValue('• Categorías de casos');
    sheet.getRange('B18').setValue('🔒 Seguridad:').setFontSize(12).setFontWeight('bold');
    sheet.getRange('B19').setValue('• Solo verás tus propias sesiones');
    sheet.getRange('B23').setValue('📞 Soporte:').setFontSize(12).setFontWeight('bold');
    sheet.getRange('B24').setValue('Karen A. Guzmán V. - kareng@tec.mx');
    sheet.getRange('B25').setValue('© 2025 MentorIA Tools').setFontSize(10).setFontColor('#999');
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 500);
    const protection = sheet.protect();
    protection.setDescription('Hoja de inicio - Solo lectura');
    protection.setWarningOnly(true);
  }
  ss.moveActiveSheet(1);
}

function logAction(action, details = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_AUDIT);
    if (!auditSheet) {
      auditSheet = ss.insertSheet(SYSTEM_CONFIG.SHEET_AUDIT);
      const headers = ['Timestamp', 'Email', 'Action', 'Details', 'Session_ID', 'IP_Address'];
      auditSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(SYSTEM_CONFIG.APP_COLOR).setFontColor('#FFF').setFontWeight('bold');
      auditSheet.setFrozenRows(1);
      const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
      if (!SECURITY_CONFIG.ADMIN_EMAILS.includes(userEmail)) auditSheet.hideSheet();
    }
    const row = [new Date(), Session.getActiveUser().getEmail(), action, JSON.stringify(details), Utilities.getUuid(), ''];
    auditSheet.appendRow(row);
    if (auditSheet.getLastRow() > 10000) auditSheet.deleteRows(2, 100);
  } catch (error) {
    console.error('Error en audit log:', error);
  }
}

function apiUploadEvidence(formObject) {
  try {
    validateApiAccess();

    const ALLOWED = /\.(jpe?g|png|pdf|docx?)$/i;
    const folder = _ensureFolder(SYSTEM_CONFIG.DRIVE_FOLDER_NAME);
    const out = [];
    const files = [].concat(formObject.evidencia || []);

    if (files.length > SECURITY_CONFIG.MAX_FILES) {
      throw new Error(`Máximo ${SECURITY_CONFIG.MAX_FILES} archivos permitidos`);
    }

    files.forEach(f => {
      if (!f || !f.getName) return;

      const fileName = f.getName();
      if (!ALLOWED.test(fileName)) { console.log('Archivo rechazado (extensión):', fileName); return; }

      const size = f.getBytes().length;
      if (size > SECURITY_CONFIG.MAX_FILE_SIZE) { console.log('Archivo rechazado (tamaño):', fileName, size); return; }

      const timestamp = Utilities.formatDate(new Date(), SECURITY_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
      const safeFileName = `${timestamp}_${fileName.replace(/[^a-zA-Z0-9._-]/g, '_')}`;
      const file = folder.createFile(f.setName(safeFileName));
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      out.push(file.getUrl());

      logAction('UPLOAD_EVIDENCE', { fileName: safeFileName, size });
    });

    return { ok: true, links: out };
  } catch (error) {
    console.error('[Upload Error]', error);
    logAction('ERROR_UPLOAD', { error: error.toString() });
    return { ok: false, links: [], error: error.message };
  }
}

/**
 * Exportar sesiones a Excel (.xlsx) y guardar en Drive
 */
function exportSessionsToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  if (!sh || sh.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No hay datos para exportar');
    return;
  }
  const fileName = `MentorPal_Export_${Utilities.formatDate(new Date(), SECURITY_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss')}.xlsx`;
  const xlsxBlob = ss.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  xlsxBlob.setName(fileName);

  // Guardar en Drive
  const file = DriveApp.createFile(xlsxBlob);
  const url = file.getUrl();

  SpreadsheetApp.getUi().alert(
    '✅ Exportación a Excel Completada',
    `El archivo Excel ha sido creado:\n\n${fileName}\n\nURL: ${url}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  logAction('DATA_EXPORT_XLSX', { fileName, url });
}
// ============ FUNCIONES ADMINISTRATIVAS ADICIONALES v3.2 ============
// Agregar estas funciones al final de Code.gs

/**
 * Dashboard mejorado con análisis de Mood Meter
 */
function showEnhancedDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  
  if (!sh || sh.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('📊 Dashboard', 'No hay datos suficientes', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const ixRisk = headers.indexOf('NivelRiesgo');
  const ixMentor = headers.indexOf('Mentor');
  const ixDate = headers.indexOf('FechaSesion');
  const ixMood = headers.indexOf('AI_Sentiment'); // Reutilizado para mood
  const ixCat = headers.indexOf('AI_RiskPrediction'); // Reutilizado para categoría
  const ixCRM = headers.indexOf('CRM_Documentado');
  
  let totalSessions = data.length - 1;
  let criticalCases = 0;
  let preventiveCases = 0;
  let normalCases = 0;
  let todaySessions = 0;
  let pendingCRM = 0;
  
  
  const mentorCount = {};
  const moodCount = { red: 0, yellow: 0, blue: 0, green: 0 };
  const categoryCount = {};
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    // Análisis de riesgo
    const risk = Number(data[i][ixRisk]) || 0;
    if (risk >= 7) criticalCases++;
    else if (risk >= 4) preventiveCases++;
    else normalCases++;
    
    // Contador de mentores
    const mentor = data[i][ixMentor];
    mentorCount[mentor] = (mentorCount[mentor] || 0) + 1;
    
    // Análisis de mood
    const mood = data[i][ixMood];
    if (mood && moodCount[mood] !== undefined) {
      moodCount[mood]++;
    }
    
    // Análisis de categorías
    const category = data[i][ixCat];
    if (category) {
      categoryCount[category] = (categoryCount[category] || 0) + 1;
    }
    
    // CRM pendientes
    if (data[i][ixCRM] === 'Pendiente' || !data[i][ixCRM]) {
      pendingCRM++;
    }
    
    // Sesiones de hoy
    const sessionDate = new Date(data[i][ixDate]);
    sessionDate.setHours(0, 0, 0, 0);
    if (sessionDate.getTime() === today.getTime()) {
      todaySessions++;
    }
  }
  
  // Top mentores
  const topMentors = Object.entries(mentorCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => `• ${name}: ${count} sesiones`)
    .join('\n');
  
  // Top categorías
  const topCategories = Object.entries(categoryCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([cat, count]) => `• ${CRM_CATEGORIES[cat] || cat}: ${count}`)
    .join('\n');
  
  // Análisis emocional
  const totalMood = moodCount.red + moodCount.yellow + moodCount.blue + moodCount.green;
  const moodPercentages = totalMood > 0 ? {
    red: Math.round((moodCount.red / totalMood) * 100),
    yellow: Math.round((moodCount.yellow / totalMood) * 100),
    blue: Math.round((moodCount.blue / totalMood) * 100),
    green: Math.round((moodCount.green / totalMood) * 100)
  } : { red: 0, yellow: 0, blue: 0, green: 0 };
  
  const message = `
📊 DASHBOARD MEJORADO ${SYSTEM_CONFIG.APP_NAME} v${SYSTEM_CONFIG.APP_VERSION}
${'━'.repeat(40)}

📝 RESUMEN GENERAL
• Sesiones totales: ${totalSessions}
• Sesiones hoy: ${todaySessions}
• Pendientes en CRM: ${pendingCRM} (${Math.round(pendingCRM/totalSessions*100)}%)

🚦 DISTRIBUCIÓN POR NIVEL DE ATENCIÓN
• 🔴 Críticos (7-10): ${criticalCases} (${Math.round(criticalCases/totalSessions*100)}%)
• 🟡 Preventivos (4-6): ${preventiveCases} (${Math.round(preventiveCases/totalSessions*100)}%)
• 🟢 Normales (1-3): ${normalCases} (${Math.round(normalCases/totalSessions*100)}%)

😊 ANÁLISIS EMOCIONAL (Mood Meter)
• 🔴 Alta energía - Desagradable: ${moodPercentages.red}%
• 🟡 Alta energía - Agradable: ${moodPercentages.yellow}%
• 🔵 Baja energía - Desagradable: ${moodPercentages.blue}%
• 🟢 Baja energía - Agradable: ${moodPercentages.green}%

🏷️ TOP 5 CATEGORÍAS
${topCategories || 'Sin datos de categorías'}

👥 TOP 5 MENTORES
${topMentors}

📈 PROMEDIO DIARIO
• ${Math.round(totalSessions / 30)} sesiones/día (últimos 30 días)

${'━'.repeat(40)}
Generado: ${new Date().toLocaleString('es-MX')}
  `;
  
  SpreadsheetApp.getUi().alert(`📊 Dashboard Mejorado ${SYSTEM_CONFIG.APP_NAME}`, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// =========================================================================================
// MODIFICACIÓN: La siguiente sección ha sido reescrita para enviar resúmenes semanales
//               individuales a cada mentor, en lugar de a un email de coordinación.
// =========================================================================================

/**
 * [MODIFICADO] Enviar resumen semanal a cada mentor.
 * Se ejecuta los viernes para resumir la actividad de la semana.
 */
// ============ REPORTE SEMANAL CON DISEÑO MEJORADO v3.3 ============
// Reemplaza las funciones sendWeeklyMentorSummary() y generateWeeklySummaryHTML() 
// con estas versiones mejoradas

/**
 * Enviar progreso semanal personalizado a cada mentor/a
 * Solo envía si hubo actividad durante la semana
 */
function sendWeeklyMentorSummary() {
  try {
    const mentors = apiGetMentors();
    let emailsSent = 0;
    let mentorsWithActivity = 0;
    let mentorsSkipped = 0;
    
    mentors.forEach(mentor => {
      if (!mentor.email) return;
      
      // Obtener estadísticas del mentor
      const stats = getStatsForMentor(mentor.name, mentor.email);
      
      // IMPORTANTE: Solo enviar si tiene actividad EN LA SEMANA
      if (stats.weekSessions === 0) {
        mentorsSkipped++;
        console.log(`Saltando ${mentor.name} - Sin actividad esta semana`);
        return;
      }
      
      mentorsWithActivity++;
      
      // Generar HTML (plantilla con colores por comunidad)
      const html = generateWeeklyProgressEmailV2(mentor.name, stats);
      
      MailApp.sendEmail({
        to: mentor.email,
        subject: `🌟 Tu resumen semanal MentorPal - ${stats.weekSessions} ${stats.weekSessions === 1 ? 'sesión completada' : 'sesiones completadas'}`,
        htmlBody: html
      });
      
      emailsSent++;
      
      logAction('WEEKLY_EMAIL_SENT', { 
        mentor: mentor.name, 
        sessions: stats.weekSessions,
        pending: stats.pendingCRM 
      });
    });
    
    // Log resumen completo
    logAction('WEEKLY_BATCH_COMPLETE', { 
      emailsSent,
      mentorsWithActivity,
      mentorsSkipped,
      totalMentors: mentors.length
    });
    
    // Notificación en el spreadsheet
    if (emailsSent > 0) {
      SpreadsheetApp.getActive().toast(
        `✅ Resúmenes semanales enviados a ${emailsSent} mentor${emailsSent === 1 ? '' : 'es'} con actividad`,
        'MentorPal',
        5
      );
    } else {
      SpreadsheetApp.getActive().toast(
        '📭 No hay mentores con actividad esta semana',
        'MentorPal',
        3
      );
    }
    
    console.log(`Resumen semanal: ${emailsSent} enviados, ${mentorsSkipped} saltados`);
    
  } catch (error) {
    console.error('[Weekly Email Error]', error);
    logAction('ERROR_WEEKLY_EMAIL', { error: error.toString() });
  }
}

/**
 * Obtener estadísticas detalladas para un mentor específico
 */
function getStatsForMentor(mentorName, mentorEmail) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  if (!sh) return { 
    totalSessions: 0, 
    weekSessions: 0, 
    pendingCRM: 0, 
    crmRate: 0, 
    uniqueStudents: 0,
    weeklyAvg: 0,
    criticalCases: 0,
    pendingSessions: []
  };
  
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const ixMentor = headers.indexOf('Mentor');
  const ixFecha = headers.indexOf('FechaSesion');
  const ixCRM = headers.indexOf('CRM_Documentado');
  const ixMatricula = headers.indexOf('Matricula');
  const ixRiesgo = headers.indexOf('NivelRiesgo');
  const ixNombre = headers.indexOf('NombreEstudiante');
  const ixMood = headers.indexOf('AI_Sentiment'); // Para mood meter
  const ixCategory = headers.indexOf('AI_RiskPrediction'); // Para categoría
  
  let totalSessions = 0;
  let weekSessions = 0;
  let pendingCRM = 0;
  let documentedCRM = 0;
  let criticalCases = 0;
  let weekCriticalCases = 0;
  const uniqueStudents = new Set();
  const weekStudents = new Set();
  const pendingSessions = [];
  
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  // Análisis de mood de la semana
  const weekMoodCount = { red: 0, yellow: 0, blue: 0, green: 0 };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][ixMentor] === mentorName) {
      totalSessions++;
      
      const sessionDate = new Date(data[i][ixFecha]);
      const riskLevel = Number(data[i][ixRiesgo]) || 0;
      
      // Sesiones de la semana
      if (sessionDate >= weekAgo) {
        weekSessions++;
        
        // Contar mood de la semana
        const mood = data[i][ixMood];
        if (mood && weekMoodCount[mood] !== undefined) {
          weekMoodCount[mood]++;
        }
        
        // Estudiantes únicos de la semana
        if (data[i][ixMatricula]) {
          weekStudents.add(data[i][ixMatricula]);
        }
        
        // Casos críticos de la semana
        if (riskLevel >= 7) {
          weekCriticalCases++;
        }
      }
      
      // Análisis general
      if (data[i][ixCRM] === 'Pendiente' || !data[i][ixCRM]) {
        pendingCRM++;
        // Guardar info de sesiones pendientes (máximo 5)
        if (pendingSessions.length < 5) {
          pendingSessions.push({
            date: sessionDate.toLocaleDateString('es-MX'),
            student: data[i][ixNombre] || data[i][ixMatricula] || 'Sin identificar',
            matricula: data[i][ixMatricula] || ''
          });
        }
      } else if (data[i][ixCRM] === 'Sí') {
        documentedCRM++;
      }
      
      if (riskLevel >= 7) {
        criticalCases++;
      }
      
      if (data[i][ixMatricula]) {
        uniqueStudents.add(data[i][ixMatricula]);
      }
    }
  }
  
  const crmRate = totalSessions > 0 ? Math.round((documentedCRM / totalSessions) * 100) : 0;
  const weeklyAvg = Math.round(totalSessions / 4);
  
  // Análisis de mood predominante
  let dominantMood = '';
  let maxMoodCount = 0;
  for (const [mood, count] of Object.entries(weekMoodCount)) {
    if (count > maxMoodCount) {
      maxMoodCount = count;
      dominantMood = mood;
    }
  }
  
  return {
    totalSessions,
    weekSessions,
    pendingCRM,
    crmRate,
    uniqueStudents: uniqueStudents.size,
    weekStudents: weekStudents.size,
    weeklyAvg,
    criticalCases,
    weekCriticalCases,
    pendingSessions,
    weekMoodCount,
    dominantMood
  };
}

/**
 * Generar HTML del email semanal con diseño mejorado y lenguaje inclusivo
 */
/**
 * Generar HTML del email semanal - VERSIÓN COMPATIBLE CON TODOS LOS CLIENTES
 */
function generateWeeklyProgressEmail(mentorName, stats) {
  const motivationalQuotes = [
    "Cada conversación que tienes marca una diferencia en la vida de alguien.",
    "Tu dedicación construye puentes hacia el éxito estudiantil.",
    "El impacto de tu mentoría trasciende más allá de las sesiones.",
    "Tu apoyo es la brújula que guía a quienes acompañas.",
    "Cada sesión es una semilla de transformación que plantas.",
    "Tu escucha activa es el regalo más valioso que ofreces.",
    "En cada encuentro, co-creas espacios de crecimiento y esperanza."
  ];
  
  const quote = motivationalQuotes[Math.floor(Math.random() * motivationalQuotes.length)];
  
  // Análisis de mood para mensaje personalizado
  let moodMessage = '';
  if (stats.dominantMood) {
    const moodInfo = MOOD_METER[stats.dominantMood];
    if (moodInfo) {
      moodMessage = `<tr><td style="padding:0 30px;text-align:center;"><p style="color:#666666;margin:10px 0 0 0;font-size:14px;font-family:Arial,sans-serif;">Esta semana, el estado emocional predominante fue ${moodInfo.emoji} ${moodInfo.name}</p></td></tr>`;
    }
  }
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Resumen Semanal MentorPal</title>
</head>
<body style="margin:0;padding:0;background-color:#f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f5f5f5;">
    <tr>
      <td align="center" style="padding:20px;">
        
        <!-- Contenedor principal -->
        <table width="600" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;">
          
          <!-- Header con color sólido -->
          <tr>
            <td style="background-color:#7c3aed;padding:40px 30px;text-align:center;">
              <h1 style="color:#ffffff;margin:0 0 10px 0;font-size:28px;font-family:Arial,sans-serif;font-weight:normal;">
                ¡Feliz viernes, ${mentorName}! 🎉
              </h1>
              <p style="color:#ffffff;margin:0;font-size:16px;font-family:Arial,sans-serif;opacity:0.95;">
                Tu resumen semanal de MentorPal
              </p>
            </td>
          </tr>
          
          <!-- Quote motivacional -->
          <tr>
            <td style="padding:30px;text-align:center;border-bottom:1px solid #eeeeee;">
              <p style="font-style:italic;color:#666666;font-size:18px;margin:0;font-family:Georgia,serif;">
                "${quote}"
              </p>
            </td>
          </tr>
          ${moodMessage}
          
          <!-- Título de estadísticas -->
          <tr>
            <td style="padding:30px 30px 20px 30px;">
              <h2 style="color:#333333;margin:0 0 20px 0;font-size:20px;font-family:Arial,sans-serif;">
                📊 Tu progreso esta semana
              </h2>
            </td>
          </tr>
          
          <!-- Estadísticas principales con tablas -->
          <tr>
            <td style="padding:0 30px 30px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td width="48%" style="background-color:#f0f4ff;padding:20px;text-align:center;border:1px solid #d0d9ff;">
                    <div style="font-size:36px;font-weight:bold;color:#7c3aed;font-family:Arial,sans-serif;">${stats.weekSessions}</div>
                    <div style="color:#666666;margin-top:5px;font-size:13px;font-family:Arial,sans-serif;">
                      ${stats.weekSessions === 1 ? 'Sesión' : 'Sesiones'} esta semana
                    </div>
                  </td>
                  <td width="4%">&nbsp;</td>
                  <td width="48%" style="background-color:${stats.pendingCRM > 0 ? '#fff0f0' : '#e8f5e9'};padding:20px;text-align:center;border:1px solid ${stats.pendingCRM > 0 ? '#ffcccc' : '#b3e5b3'};">
                    <div style="font-size:36px;font-weight:bold;color:${stats.pendingCRM > 0 ? '#dc2626' : '#16a34a'};font-family:Arial,sans-serif;">
                      ${stats.pendingCRM}
                    </div>
                    <div style="color:#666666;margin-top:5px;font-size:13px;font-family:Arial,sans-serif;">
                      Pendientes en CRM
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Esta semana específicamente -->
          ${stats.weekStudents > 0 ? `
          <tr>
            <td style="padding:0 30px 20px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#f0f9ff;padding:20px;border:1px solid #b3e0ff;">
                    <h3 style="color:#0369a1;margin:0 0 10px 0;font-size:16px;font-family:Arial,sans-serif;">
                      🌟 Esta semana acompañaste a:
                    </h3>
                    <p style="margin:5px 0;color:#0369a1;font-family:Arial,sans-serif;font-size:14px;">
                      • <strong>${stats.weekStudents}</strong> ${stats.weekStudents === 1 ? 'estudiante' : 'estudiantes diferentes'}
                    </p>
                    ${stats.weekCriticalCases > 0 ? `
                    <p style="margin:5px 0;color:#0369a1;font-family:Arial,sans-serif;font-size:14px;">
                      • Atendiste <strong>${stats.weekCriticalCases}</strong> ${stats.weekCriticalCases === 1 ? 'caso crítico' : 'casos críticos'}
                    </p>` : ''}
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}
          
          <!-- Métricas generales -->
          <tr>
            <td style="padding:0 30px 20px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#fafafa;padding:20px;border:1px solid #e5e5e5;">
                    <h3 style="color:#333333;margin:0 0 15px 0;font-size:16px;font-family:Arial,sans-serif;">
                      📈 Tu trayectoria acumulada
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,sans-serif;font-size:14px;color:#666666;">
                      <tr><td style="padding:4px 0;">• Has completado <strong style="color:#333333;">${stats.totalSessions}</strong> ${stats.totalSessions === 1 ? 'sesión' : 'sesiones'} en total</td></tr>
                      <tr><td style="padding:4px 0;">• Tu promedio semanal es de <strong style="color:#333333;">${stats.weeklyAvg}</strong> ${stats.weeklyAvg === 1 ? 'sesión' : 'sesiones'}</td></tr>
                      <tr><td style="padding:4px 0;">• Has acompañado a <strong style="color:#333333;">${stats.uniqueStudents}</strong> ${stats.uniqueStudents === 1 ? 'estudiante único' : 'estudiantes únicos'}</td></tr>
                      <tr><td style="padding:4px 0;">• Tasa de documentación en CRM: <strong style="color:#333333;">${stats.crmRate}%</strong></td></tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Call to action según estado -->
          ${stats.pendingCRM > 0 ? `
          <tr>
            <td style="padding:0 30px 20px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#fffbeb;padding:20px;border-left:4px solid #fbbf24;">
                    <h3 style="color:#333333;margin:0 0 10px 0;font-size:16px;font-family:Arial,sans-serif;">
                      ⏰ Acción recomendada para el cierre de semana
                    </h3>
                    <p style="color:#666666;margin:0 0 10px 0;font-family:Arial,sans-serif;font-size:14px;">
                      Tienes ${stats.pendingCRM} ${stats.pendingCRM === 1 ? 'sesión por documentar' : 'sesiones por documentar'} en el CRM.
                    </p>
                    ${stats.pendingSessions.length > 0 ? `
                    <p style="color:#999999;font-size:13px;margin:10px 0 5px 0;font-family:Arial,sans-serif;">
                      Sesiones pendientes más recientes:
                    </p>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                      ${stats.pendingSessions.map(s => `
                      <tr>
                        <td style="padding:2px 0;color:#999999;font-size:13px;font-family:Arial,sans-serif;">
                          • ${s.date} - ${s.student}
                        </td>
                      </tr>`).join('')}
                    </table>
                    ` : ''}
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:20px;">
                      <tr>
                        <td align="center">
                          <a href="#" style="display:inline-block;padding:12px 30px;background-color:#fbbf24;color:#333333;text-decoration:none;font-weight:bold;font-family:Arial,sans-serif;font-size:14px;">
                            Documentar en CRM →
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : `
          <tr>
            <td style="padding:0 30px 20px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#e8f5e9;padding:20px;border-left:4px solid #4ade80;">
                    <h3 style="color:#333333;margin:0 0 10px 0;font-size:16px;font-family:Arial,sans-serif;">
                      🌟 ¡Excelente trabajo!
                    </h3>
                    <p style="color:#666666;margin:0;font-family:Arial,sans-serif;font-size:14px;">
                      Todas tus sesiones están documentadas en el CRM. Tu constancia y dedicación marcan la diferencia.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          `}
          
          <!-- Mensaje personalizado -->
          <tr>
            <td style="padding:0 30px 20px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#f0f4ff;padding:20px;text-align:center;border:1px solid #d0d9ff;">
                    <p style="color:#7c3aed;margin:0;font-weight:bold;font-family:Arial,sans-serif;font-size:14px;">
                      💜 Tu progreso es único y valioso
                    </p>
                    <p style="color:#666666;margin:10px 0 0 0;font-size:13px;font-family:Arial,sans-serif;">
                      ${stats.weekSessions === 1 
                        ? 'Una sesión puede cambiar una vida. Gracias por estar presente.' 
                        : `${stats.weekSessions} conversaciones significativas esta semana. Cada una cuenta.`}
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Tip del fin de semana -->
          <tr>
            <td style="padding:0 30px 30px 30px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="background-color:#fffbeb;padding:15px;border:1px solid #fcd34d;">
                    <p style="color:#92400e;margin:0;font-size:13px;font-family:Arial,sans-serif;">
                      💡 <strong>Tip para el fin de semana:</strong> Tómate un momento para reflexionar sobre los momentos significativos de esta semana. Tu bienestar es tan importante como el de quienes acompañas.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="background-color:#fafafa;padding:20px;text-align:center;border-top:1px solid #eeeeee;">
              <p style="color:#999999;margin:0;font-size:12px;font-family:Arial,sans-serif;">
                MentorPal v${SYSTEM_CONFIG.APP_VERSION} - Tu apoyo inteligente en cada mentoría<br>
                © 2025 MentorIA Tools | Karen A. Guzmán V.
              </p>
              <p style="color:#cccccc;margin:5px 0 0 0;font-size:11px;font-family:Arial,sans-serif;">
                Este resumen se envía solo cuando has tenido actividad durante la semana
              </p>
            </td>
          </tr>
          
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
  `;
}

/**
 * Función de prueba para enviar resumen semanal manualmente
 */
function testWeeklySummary() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '📧 Enviar Resumen Semanal',
    '¿Deseas enviar el resumen semanal a todos los mentores con actividad?\n\n' +
    'Solo recibirán email quienes hayan tenido sesiones esta semana.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    sendWeeklyMentorSummary();
  }
}

// ============ FIN DE FUNCIONES DE REPORTE SEMANAL MEJORADO v3.3 ============
// =========================================================================================
// FIN DE LA MODIFICACIÓN
// =========================================================================================

// Nueva versión renderizada por plantilla con theme por comunidad
function generateWeeklyProgressEmailV2(mentorName, stats) {
  const motivationalQuotes = [
    "Cada conversación que tienes marca una diferencia en la vida de alguien.",
    "Tu dedicación construye puentes hacia el éxito estudiantil.",
    "El impacto de tu mentoría trasciende más allá de las sesiones.",
    "Tu apoyo es la brújula que guía a quienes acompañas.",
    "Cada sesión es una semilla de transformación que plantas.",
    "Tu escucha activa es el regalo más valioso que ofreces.",
    "En cada encuentro, co-creas espacios de crecimiento y esperanza."
  ];
  const quote = motivationalQuotes[Math.floor(Math.random() * motivationalQuotes.length)];
  let moodMessage = '';
  if (stats.dominantMood) {
    const moodInfo = MOOD_METER[stats.dominantMood];
    if (moodInfo) {
      moodMessage = `<tr><td style="padding:0 30px;text-align:center;"><p style="color:#666666;margin:10px 0 0 0;font-size:14px;font-family:Arial,sans-serif;">Esta semana, el estado emocional predominante fue ${moodInfo.emoji} ${moodInfo.name}</p></td></tr>`;
    }
  }
  const community = getMentorCommunityByName(mentorName);
  const theme = getThemeForCommunity(community);
  const t = HtmlService.createTemplateFromFile('generateWeeklyProgressEmail');
  t.mentorName = mentorName;
  t.stats = stats;
  t.quote = quote;
  t.moodMessage = moodMessage;
  t.themeHex = theme.hex;
  t.appVersion = SYSTEM_CONFIG.APP_VERSION;
  return t.evaluate().getContent();
}

/**
 * Exportar datos para análisis externo
 */
function exportSessionsToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  
  if (!sh || sh.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No hay datos para exportar');
    return;
  }
  
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const fileName = `MentorPal_Export_${Utilities.formatDate(new Date(), SECURITY_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss')}.csv`;
  
  // Crear CSV
  let csv = headers.join(',') + '\n';
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i].map(cell => {
      // Escapar comas y comillas en el contenido
      const value = String(cell || '');
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        return '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    });
    csv += row.join(',') + '\n';
  }
  
  // Crear archivo en Drive
  const blob = Utilities.newBlob(csv, 'text/csv', fileName);
  const file = DriveApp.createFile(blob);
  
  // Compartir enlace
  const url = file.getUrl();
  
  SpreadsheetApp.getUi().alert(
    '✅ Exportación Completada',
    `El archivo CSV ha sido creado:\n\n${fileName}\n\nURL: ${url}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  logAction('DATA_EXPORT', { fileName, records: data.length - 1 });
}

/**
 * Limpiar sesiones antiguas (archivar)
 */
function archiveOldSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  
  if (!sh || sh.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No hay sesiones para archivar');
    return;
  }
  
  // Crear o obtener hoja de archivo
  let archiveSheet = ss.getSheetByName('ARCHIVO_SESIONES');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('ARCHIVO_SESIONES');
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues();
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    archiveSheet.getRange(1, 1, 1, headers[0].length)
      .setBackground(SYSTEM_CONFIG.APP_COLOR)
      .setFontColor('#FFF')
      .setFontWeight('bold');
    archiveSheet.setFrozenRows(1);
  }
  
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const dateIndex = headers.indexOf('FechaSesion');
  
  const sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
  
  const toArchive = [];
  const toKeep = [headers]; // Mantener headers
  
  for (let i = 1; i < data.length; i++) {
    const sessionDate = new Date(data[i][dateIndex]);
    if (sessionDate < sixMonthsAgo) {
      toArchive.push(data[i]);
    } else {
      toKeep.push(data[i]);
    }
  }
  
  if (toArchive.length === 0) {
    SpreadsheetApp.getUi().alert('No hay sesiones de más de 6 meses para archivar');
    return;
  }
  
  // Agregar al archivo
  if (toArchive.length > 0) {
    archiveSheet.getRange(
      archiveSheet.getLastRow() + 1,
      1,
      toArchive.length,
      toArchive[0].length
    ).setValues(toArchive);
  }
  
  // Actualizar hoja principal
  sh.clear();
  sh.getRange(1, 1, toKeep.length, toKeep[0].length).setValues(toKeep);
  
  // Restaurar formato de headers
  sh.getRange(1, 1, 1, headers.length)
    .setBackground(SYSTEM_CONFIG.APP_COLOR)
    .setFontColor('#FFF')
    .setFontWeight('bold');
  sh.setFrozenRows(1);
  
  SpreadsheetApp.getUi().alert(
    '✅ Archivo Completado',
    `Se archivaron ${toArchive.length} sesiones de más de 6 meses.\n\nSesiones activas: ${toKeep.length - 1}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  logAction('ARCHIVE_SESSIONS', {
    archived: toArchive.length,
    remaining: toKeep.length - 1
  });
}

/**
 * Validar integridad de datos
 */
function validateDataIntegrity() {
  const issues = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Verificar hoja de sesiones
  const sessionsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
  if (!sessionsSheet) {
    issues.push('❌ Falta hoja de sesiones (PROCESSED_DATA)');
  } else {
    const data = sessionsSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const requiredColumns = [
        'Session_ID', 'Mentor', 'Matricula', 'FechaSesion',
        'NivelRiesgo', 'CRM_Documentado'
      ];
      
      requiredColumns.forEach(col => {
        if (!headers.includes(col)) {
          issues.push(`❌ Falta columna: ${col}`);
        }
      });
      
      // Verificar datos
      let sessionsWithoutID = 0;
      let sessionsWithoutMentor = 0;
      let sessionsWithoutMatricula = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (!data[i][headers.indexOf('Session_ID')]) sessionsWithoutID++;
        if (!data[i][headers.indexOf('Mentor')]) sessionsWithoutMentor++;
        if (!data[i][headers.indexOf('Matricula')]) sessionsWithoutMatricula++;
      }
      
      if (sessionsWithoutID > 0) issues.push(`⚠️ ${sessionsWithoutID} sesiones sin ID`);
      if (sessionsWithoutMentor > 0) issues.push(`⚠️ ${sessionsWithoutMentor} sesiones sin mentor`);
      if (sessionsWithoutMatricula > 0) issues.push(`⚠️ ${sessionsWithoutMatricula} sesiones sin matrícula`);
    }
  }
  
  // Verificar hoja de mentores
  const mentorsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_MENTORS);
  if (!mentorsSheet) {
    issues.push('❌ Falta hoja de mentores (MENTORES_CATALOG)');
  } else {
    const mentorData = mentorsSheet.getDataRange().getValues();
    if (mentorData.length <= 1) {
      issues.push('⚠️ No hay mentores registrados');
    }
  }
  
  // Verificar hoja de estudiantes
  const studentsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_CATALOG);
  if (!studentsSheet) {
    issues.push('❌ Falta hoja de estudiantes (CATALOGO_ESTUDIANTES)');
  }
  
  // Mostrar resultados
  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Validación Exitosa',
      'No se encontraron problemas de integridad en los datos.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '⚠️ Problemas de Integridad Encontrados',
      issues.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  logAction('DATA_INTEGRITY_CHECK', { issues: issues.length });
  return issues.length === 0;
}

/**
 * Menú de administrador mejorado
 */
function showAdminPanelV32() {
  const userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!SECURITY_CONFIG.ADMIN_EMAILS.includes(userEmail)) {
    SpreadsheetApp.getUi().alert('Acceso denegado','No tienes permisos de administrador', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const community = getMentorCommunityByEmail(userEmail);
  const theme = getThemeForCommunity(community);
  const themeHex = theme.hex || SYSTEM_CONFIG.APP_COLOR;
  
  const html = `
    <div style="font-family: 'Segoe UI', Arial; padding: 20px;">
      <h2 style="color:${themeHex};">${SYSTEM_CONFIG.APP_ICON} Panel de Administración - ${SYSTEM_CONFIG.APP_NAME} v3.3</h2>
      
      <h3 style="color:${themeHex};">🆕 Nuevas Funcionalidades v3.3</h3>
      <ul>
        <li>✅ Mood Meter (RULER Approach)</li>
        <li>✅ Categorías de casos (PIN, CAG, etc.)</li>
        <li>✅ Email con formato CRM</li>
        <li>✅ Dashboard mejorado con análisis emocional</li>
      </ul>
      
      <h3>📊 Herramientas Administrativas</h3>
      <p>Ejecuta estas funciones desde el editor de Apps Script:</p>
      <ul>
        <li><code>showEnhancedDashboard()</code> - Dashboard con análisis de mood</li>
        <li><code>sendWeeklyMentorSummary()</code> - Reporte semanal individual</li>
        <li><code>exportSessionsToCSV()</code> - Exportar datos</li>
        <li><code>exportSessionsToExcel()</code> - Exportar a Excel</li>
        <li><code>archiveOldSessions()</code> - Archivar sesiones antiguas</li>
        <li><code>validateDataIntegrity()</code> - Validar integridad</li>
      </ul>
      
      <h3>📧 Configuración de Emails</h3>
      <p>Los emails de confirmación incluyen ahora:</p>
      <ul>
        <li>Asunto sugerido para CRM</li>
        <li>Descripción completa formateada</li>
        <li>Estado emocional del estudiante</li>
        <li>Categoría del caso</li>
      </ul>
      
      <h3>🔒 Configuración</h3>
      <ul>
        <li>Dominios permitidos: ${SECURITY_CONFIG.ALLOWED_DOMAINS.join(', ')}</li>
        <li>Administradores: ${SECURITY_CONFIG.ADMIN_EMAILS.length}</li>
        <li>Versión: ${SYSTEM_CONFIG.APP_VERSION}</li>
      </ul>
      
      <p style="color:#666;font-size:12px;margin-top:20px;">
        ${SYSTEM_CONFIG.APP_TAGLINE}<br>
        © 2025 Karen A. Guzmán V. - MentorIA Tools
      </p>
    </div>`;
    
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(600).setHeight(700),
    'Admin Panel v3.3 - ' + SYSTEM_CONFIG.APP_NAME
  );
}
/**
 * Exportar sesiones del mentor a Excel
 * @return {Object} URL del archivo Excel generado
 */
function exportMentorSessionsToExcel(viewAll) {
  try {
    const { email, accessLevel } = validateApiAccess();
    const mentorName = getMentorNameByEmail(email);
    
    const sh = SpreadsheetApp.getActive().getSheetByName(SYSTEM_CONFIG.SHEET_SESSIONS);
    if (!sh || sh.getLastRow() <= 1) {
      return { ok: false, message: 'No hay datos para exportar' };
    }
    
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const ixMentor = headers.indexOf('Mentor');
    const ixDate = headers.indexOf('FechaSesion');
    const ixCRM = headers.indexOf('CRM_Documentado');
    
    // Filtrar según alcance permitido
    const exportData = [headers];
    let crmPending = 0;
    const scopeAll = (viewAll === true && accessLevel === 'ADMIN');
    
    for (let i = 1; i < data.length; i++) {
      if (scopeAll || data[i][ixMentor] === mentorName) {
        exportData.push(data[i]);
        if (data[i][ixCRM] === 'Pendiente' || !data[i][ixCRM]) {
          crmPending++;
        }
      }
    }
    
    if (exportData.length <= 1) {
      return { ok: false, message: 'No tienes sesiones registradas' };
    }
    
    // Crear nuevo Spreadsheet temporal
    const who = scopeAll ? 'TODOS' : mentorName;
    const fileName = `MentorPal_Export_${who}_${Utilities.formatDate(new Date(), SECURITY_CONFIG.TIMEZONE, 'yyyy-MM-dd')}`;
    const tempSS = SpreadsheetApp.create(fileName);
    const tempSheet = tempSS.getActiveSheet();
    
    // Establecer datos
    tempSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    
    // Formato de encabezados
    tempSheet.getRange(1, 1, 1, exportData[0].length)
      .setBackground('#3B82F6')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    // Ajustar columnas
    tempSheet.autoResizeColumns(1, exportData[0].length);
    
    // Agregar hoja de resumen
    const summarySheet = tempSS.insertSheet('RESUMEN');
    summarySheet.getRange('A1:B8').setValues([
      ['RESUMEN DE EXPORTACIÓN', ''],
      ['Fecha:', new Date().toLocaleString('es-MX')],
      ['Alcance:', scopeAll ? 'Todos los mentores (admin)' : 'Solo mis sesiones'],
      ['Mentor:', mentorName],
      ['Total Sesiones:', exportData.length - 1],
      ['Pendientes CRM:', crmPending],
      ['Documentadas:', (exportData.length - 1) - crmPending],
      ['Tasa CRM:', Math.round(((exportData.length - 1 - crmPending) / (exportData.length - 1)) * 100) + '%']
    ]);
    
    summarySheet.getRange('A1:B1')
      .merge()
      .setBackground('#3B82F6')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Obtener URL de descarga
    const fileId = tempSS.getId();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
    const viewUrl = tempSS.getUrl();
    
    logAction('EXPORT_TO_EXCEL', { mentor: mentorName, scope: scopeAll ? 'ALL' : 'MINE', sessions: exportData.length - 1, fileId: fileId });
    
    return { 
      ok: true, 
      downloadUrl: downloadUrl,
      viewUrl: viewUrl,
      fileName: fileName + '.xlsx',
      sessions: exportData.length - 1,
      message: 'Excel generado exitosamente'
    };
    
  } catch (error) {
    console.error('Error exportando a Excel:', error);
    logAction('ERROR_EXPORT_EXCEL', { error: error.toString() });
    return { ok: false, message: 'Error al generar Excel: ' + error.toString() };
  }
}
/**
 * Borra solo los triggers que gestionamos aquí (diario y semanal).
 */
function resetSummaryTriggers() {
  var names = [
    'sendDailyMentorSummary',
    'sendWeeklyMentorSummary' // MODIFICADO
  ];
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var h = triggers[i].getHandlerFunction();
    if (names.indexOf(h) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Instala los triggers:
 * - Diario: Lunes a Jueves a las 18:00 (4 triggers, uno por dia)
 * - Semanal: Viernes a las 18:00
 */
function setupSummaryTriggers() {
  resetSummaryTriggers();

  // Diario L-J 18:00
  ScriptApp.newTrigger('sendDailyMentorSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(18).create();

  ScriptApp.newTrigger('sendDailyMentorSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(18).create();

  ScriptApp.newTrigger('sendDailyMentorSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(18).create();

  ScriptApp.newTrigger('sendDailyMentorSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(18).create();

  // Semanal Viernes 18:00
  // MODIFICADO: Llama a la nueva función para resúmenes individuales.
  ScriptApp.newTrigger('sendWeeklyMentorSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(18).create();
}
/**
 * Configurar envío semanal automático (viernes 6pm)
 */
function setupWeeklyTrigger() {
  // Eliminar triggers anteriores
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendWeeklyMentorSummary') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger para viernes 6pm
  ScriptApp.newTrigger('sendWeeklyMentorSummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(18)
    .create();
  
  SpreadsheetApp.getActive().toast(
    '✅ Resumen semanal configurado para viernes 6:00 PM',
    'MentorPal',
    5
  );
  
  logAction('WEEKLY_TRIGGER_SETUP', { 
    day: 'Friday',
    time: '18:00'
  });
}

/**
 * Utilidad opcional para revisar en logs los triggers instalados.
 */
function listSummaryTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    Logger.log(triggers[i].getHandlerFunction() + ' @ ' + JSON.stringify(triggers[i].getTriggerSource()));
  }
}

// ============ FIN DE FUNCIONES ADMINISTRATIVAS v3.2 ============
// ============ FIN DEL SISTEMA MENTORPAL v3.2 ============

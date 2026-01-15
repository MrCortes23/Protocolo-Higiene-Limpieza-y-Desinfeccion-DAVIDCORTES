// ==================== CONFIGURACIÓN ====================
const SPREADSHEET_ID = '1nR7gF7TNYtdb0Eq3ffB62pEBT7vT25CCsys9xzPrx0U'; 

// Nombres de las hojas
const SHEETS = {
  USUARIOS: 'USUARIOS',
  MAQUINAS: 'MAQUINAS',
  ELEMENTOS: 'ELEMENTOS',
  PLANEACION: 'PLANEACION',
  REGISTROS_LIMPIEZA: 'REGISTROS_LIMPIEZA',
  PROCESOS: 'PROCESOS' 
};

const CORREOS_POR_TURNO = {
  'MAÑANA': 'pragestionhumana@pastascomarrico.com',
  'TARDE': 'pragestionhumana@pastascomarrico.com',
  'NOCHE': 'pragestionhumana@pastascomarrico.com'
};

function obtenerCorreoPorTurno(turno) {
  turno = turno.toUpperCase();
  return CORREOS_POR_TURNO[turno] || 'pragestionhumana@pastascomarrico.com';
}

function obtenerTurnoContrario(turno) {
  const turnos = {
    'MAÑANA': 'TARDE',
    'TARDE': 'NOCHE',
    'NOCHE': 'MAÑANA'
  };
  
  return turnos[turno.toUpperCase()] || 'TARDE';
}


// ==================== FUNCIONES WEB APP ====================
function doGet(e) {
  const title = 'PHLYD';
  const faviconUrl = 'https://alimentosdoria.com/wp-content/uploads/2023/01/logo-doria.png';
  
  return HtmlService.createTemplateFromFile('frontend/index')
    .evaluate()
    .setTitle(title)
    .setFaviconUrl(faviconUrl)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
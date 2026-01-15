// ==================== FUNCIONES DE USUARIO ====================
function autenticarUsuario(cedula) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usuariosSheet = ss.getSheetByName(SHEETS.USUARIOS)
    
    if (!usuariosSheet) {
      return { success: false, message: 'No se encontró la hoja de usuarios' };
    }
    
    const data = usuariosSheet.getDataRange().getValues();
    const headers = data[0];
    
    const cedulaCol = headers.indexOf('CEDULA');
    const nombreCol = headers.indexOf('NOMBRE');
    const rolCol = headers.indexOf('ROL');
    const procesoCol = headers.indexOf('PROCESO');
    
    if (cedulaCol === -1 || nombreCol === -1) {
      return { success: false, message: 'Estructura de hoja de usuarios incorrecta' };
    }
    
    // Buscar usuario
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[cedulaCol] && row[cedulaCol].toString().trim() === cedula.trim()) {
        return {
          success: true,
          usuario: {
            cedula: row[cedulaCol],
            nombre: row[nombreCol],
            rol: rolCol !== -1 ? (row[rolCol] || 'operario') : 'operario',
            proceso: procesoCol !== -1 ? (row[procesoCol] || 'GENERAL') : 'GENERAL'
          }
        };
      }
    }
    
    return { success: false, message: 'Usuario no encontrado' };
    
  } catch (error) {
    console.error('Error en autenticación:', error);
    return { success: false, message: 'Error del sistema: ' + error.message };
  }
}
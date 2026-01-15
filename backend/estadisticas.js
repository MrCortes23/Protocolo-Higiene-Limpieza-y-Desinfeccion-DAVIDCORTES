// ==================== FUNCIONES DE ESTADÍSTICAS ====================
function obtenerEstadisticasMaquinas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetRegistros = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    
    if (!sheetMaquinas) {
      return { success: true, estadisticas: [] };
    }
    
    const maquinasData = sheetMaquinas.getDataRange().getValues();
    const registrosData = sheetRegistros ? sheetRegistros.getDataRange().getValues() : [];
    
    const estadisticas = [];
    
    for (let i = 1; i < maquinasData.length; i++) {
      if (maquinasData[i][0]) {
        const maquinaId = maquinasData[i][0].toString();
        let totalRegistros = 0;
        let completados = 0;
        
        for (let j = 1; j < registrosData.length; j++) {
          if (registrosData[j][2] && registrosData[j][2].toString() === maquinaId) {
            totalRegistros++;
            if (registrosData[j][7] === 'COMPLETADO') {
              completados++;
            }
          }
        }
        
        const porcentaje = totalRegistros > 0 ? Math.round((completados / totalRegistros) * 100) : 0;
        
        estadisticas.push({
          maquinaId: maquinaId,
          maquinaNombre: maquinasData[i][1] || '',
          totalRegistros: totalRegistros,
          completados: completados,
          porcentaje: porcentaje
        });
      }
    }
    
    return { success: true, estadisticas: estadisticas };
  } catch (error) {
    return { success: false, message: 'Error al obtener estadísticas: ' + error.message, estadisticas: [] };
  }
}
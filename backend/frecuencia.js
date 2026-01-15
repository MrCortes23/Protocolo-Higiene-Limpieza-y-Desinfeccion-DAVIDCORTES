function obtenerInfoValidacionMaquina(maquinaId) {
  try {
    console.log('🔍 Obteniendo información de máquina:', maquinaId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja de registros no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('ℹ️ Hoja vacía o solo con headers');
      return { 
        success: true, 
        datos: {
          total: 0,
          completados: 0,
          pendientes: 0,
          yaValidados: 0,
          porcentaje: 0,
          puedeValidar: false,
          message: 'No hay registros en la hoja'
        }
      };
    }
    
    const headers = data[0];
    
    // Buscar columnas con nombres exactos
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const estadoCol = headers.indexOf('ESTADO');
    const validadoPorCol = headers.indexOf('VALIDADO_POR');
    
    console.log('🔍 Índices de columnas:');
    console.log(`- MAQUINA_ID: ${maquinaIdCol} (columna ${maquinaIdCol + 1})`);
    console.log(`- ESTADO: ${estadoCol} (columna ${estadoCol + 1})`);
    console.log(`- VALIDADO_POR: ${validadoPorCol} (columna ${validadoPorCol + 1})`);
    
    if (maquinaIdCol === -1) {
      console.error('❌ MAQUINA_ID no encontrada');
      return { success: false, message: 'No se encontró la columna MAQUINA_ID' };
    }
    
    let total = 0;
    let completados = 0;
    let pendientes = 0;
    let yaValidados = 0;
    
    console.log(`🔍 Analizando registros para máquina: ${maquinaId}`);
    console.log(`📊 Total filas: ${data.length - 1}`);
    
    // Contar registros de esta máquina
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Saltar filas vacías
      if (!row[maquinaIdCol] || row[maquinaIdCol].toString().trim() === '') {
        continue;
      }
      
      const rowMaquinaId = row[maquinaIdCol].toString().trim();
      
      // DEBUG: Verificar coincidencias
      if (i < 10) { // Solo mostrar primeros 10 para debug
        console.log(`Fila ${i + 1}: ID="${rowMaquinaId}", Buscando: "${maquinaId}"`);
      }
      
      if (rowMaquinaId === maquinaId.toString().trim()) {
        total++;
        
        const estado = row[estadoCol] ? row[estadoCol].toString().trim() : 'PENDIENTE';
        
        console.log(`   ✅ Coincidencia fila ${i + 1}: Estado="${estado}"`);
        
        if (estado === 'COMPLETADO') {
          completados++;
          
          // Verificar si ya está validado
          if (validadoPorCol !== -1 && row[validadoPorCol]) {
            const validador = row[validadoPorCol].toString().trim();
            if (validador !== '') {
              yaValidados++;
              console.log(`       Ya validado por: ${validador}`);
            }
          }
        } else {
          pendientes++;
        }
      }
    }
    
    const puedeValidar = completados === total && total > 0 && yaValidados === 0;
    
    console.log('📊 RESULTADOS ANÁLISIS:');
    console.log(`- Total registros: ${total}`);
    console.log(`- Completados: ${completados}`);
    console.log(`- Pendientes: ${pendientes}`);
    console.log(`- Ya validados: ${yaValidados}`);
    console.log(`- Porcentaje completados: ${total > 0 ? Math.round((completados / total) * 100) : 0}%`);
    console.log(`- Puede validar: ${puedeValidar} (${completados}/${total} completados, ${yaValidados} ya validados)`);
    
    return {
      success: true,
      datos: {
        total: total,
        completados: completados,
        pendientes: pendientes,
        yaValidados: yaValidados,
        porcentaje: total > 0 ? Math.round((completados / total) * 100) : 0,
        puedeValidar: puedeValidar,
        message: `Análisis completado. ${total} registros encontrados.`
      }
    };
    
  } catch (error) {
    console.error('💥 Error en obtenerInfoValidacionMaquina:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message,
      datos: {
        total: 0,
        completados: 0,
        pendientes: 0,
        yaValidados: 0,
        porcentaje: 0,
        puedeValidar: false,
        message: 'Error al obtener información'
      }
    };
  }
}

// Agrega esta función en tu código de Apps Script
function obtenerFrecuenciaPorElemento(maquinaId, elementoId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    
    if (!sheetPlaneacion) {
      return 'No planificado';
    }
    
    const data = sheetPlaneacion.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === maquinaId.toString().trim()) {
        try {
          const configStr = data[i][7] || '[]';
          const elementosConfig = JSON.parse(configStr);
          
          if (Array.isArray(elementosConfig)) {
            for (const componente of elementosConfig) {
              if (componente.elementos && Array.isArray(componente.elementos)) {
                for (const elemento of componente.elementos) {
                  if (elemento.elementoId && elemento.elementoId.toString() === elementoId.toString()) {
                    return data[i][3] || 'Mensual'; // Retorna la frecuencia de esta planeación
                  }
                }
              }
            }
          }
        } catch (e) {
          console.warn('Error parseando elementosConfig:', e);
        }
      }
    }
    
    return 'No planificado';
  } catch (error) {
    console.error('Error obteniendo frecuencia:', error);
    return 'No planificado';
  }
}

function cambiarFrecuenciaPlaneacion(maquinaId, nuevaFrecuencia) {
  try {
    console.log('🔄 Cambiando frecuencia para máquina:', maquinaId, '->', nuevaFrecuencia);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    
    if (!sheetPlaneacion) {
      return { success: false, message: 'Hoja PLANEACION no encontrada' };
    }
    
    const data = sheetPlaneacion.getDataRange().getValues();
    const headers = data[0];
    
    // Buscar índices de columnas
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const frecuenciaCol = headers.indexOf('FRECUENCIA');
    
    if (maquinaIdCol === -1 || frecuenciaCol === -1) {
      return { success: false, message: 'Estructura de hoja incorrecta' };
    }
    
    let actualizadas = 0;
    
    // Buscar todas las planeaciones de esta máquina
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[maquinaIdCol] && row[maquinaIdCol].toString().trim() === maquinaId.toString().trim()) {
        // Actualizar frecuencia
        sheetPlaneacion.getRange(i + 1, frecuenciaCol + 1).setValue(nuevaFrecuencia);
        actualizadas++;
      }
    }
    
    if (actualizadas > 0) {
      return {
        success: true,
        message: `Frecuencia actualizada a "${nuevaFrecuencia}"`,
        actualizadas: actualizadas,
        maquinaId: maquinaId
      };
    } else {
      return { 
        success: false, 
        message: 'No se encontraron planeaciones para esta máquina' 
      };
    }
    
  } catch (error) {
    console.error('💥 Error cambiando frecuencia:', error);
    return { 
      success: false, 
      message: 'Error al cambiar frecuencia: ' + error.message 
    };
  }
}

function obtenerFrecuenciaActual(maquinaId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    
    if (!sheetPlaneacion) {
      return { success: false, message: 'Hoja PLANEACION no encontrada' };
    }
    
    const data = sheetPlaneacion.getDataRange().getValues();
    const headers = data[0];
    
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const frecuenciaCol = headers.indexOf('FRECUENCIA');
    
    // Buscar la primera planeación de esta máquina
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[maquinaIdCol] && row[maquinaIdCol].toString().trim() === maquinaId.toString().trim()) {
        return {
          success: true,
          frecuencia: row[frecuenciaCol] || 'Mensual',
          maquinaId: maquinaId
        };
      }
    }
    
    return { 
      success: false, 
      message: 'No se encontró la máquina',
      frecuencia: 'Mensual' 
    };
    
  } catch (error) {
    console.error('Error obteniendo frecuencia:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message,
      frecuencia: 'Mensual' 
    };
  }
}
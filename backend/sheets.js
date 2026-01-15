// ==================== INICIALIZACIÓN DE HOJAS ====================
function inicializarHojas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // USUARIOS: CEDULA | NOMBRE | ROL | AREA
  let sheet = ss.getSheetByName(SHEETS.USUARIOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.USUARIOS);
    sheet.appendRow(['CEDULA', 'NOMBRE', 'ROL', 'AREA']);
  }
  
  // MAQUINAS: ID | NOMBRE | DESCRIPCION | ACTIVA
  sheet = ss.getSheetByName(SHEETS.MAQUINAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MAQUINAS);
    sheet.appendRow(['ID', 'NOMBRE', 'DESCRIPCION', 'ACTIVA']);
  }
  
  // ELEMENTOS: ID | MAQUINA_ID | NOMBRE | DESCRIPCION
  sheet = ss.getSheetByName(SHEETS.ELEMENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.ELEMENTOS);
    sheet.appendRow(['ID', 'MAQUINA_ID', 'NOMBRE', 'DESCRIPCION']);
  }
  
  // PLANEACION
  sheet = ss.getSheetByName(SHEETS.PLANEACION);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PLANEACION);
    sheet.appendRow(['ID', 'MAQUINA_ID', 'MAQUINA_NOMBRE', 'FRECUENCIA', 'LIMPIEZA_SECO', 'LIMPIEZA_HUMEDO', 'DESINFECCION', 'ELEMENTOS_CONFIG', 'FECHA_CREACION', 'USUARIO_CREADOR', 'ESTADO']);
  }
  
  // REGISTROS_LIMPIEZA
  sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.REGISTROS_LIMPIEZA);
    sheet.appendRow(['ID', 'PLANEACION_ID', 'MAQUINA_ID', 'MAQUINA_NOMBRE', 'ELEMENTO_ID', 'ELEMENTO_NOMBRE', 'TIPO_LIMPIEZA', 'ESTADO', 'RESPONSABLE', 'FECHA_REALIZACION', 'OBSERVACIONES', 'FECHA_CREACION', 'FECHA_FINALIZACION']);
  }

  sheet = ss.getSheetByName(SHEETS.PROCESOS);
if (!sheet) {
  sheet = ss.insertSheet(SHEETS.PROCESOS);
  sheet.appendRow(['ID', 'NOMBRE', 'DESCRIPCION', 'ACTIVO']);
  // Agregar procesos básicos
  sheet.appendRow(['PROD', 'Producción', 'Proceso de producción principal', 'SI']);
  sheet.appendRow(['CAL', 'Calidad', 'Control de calidad', 'SI']);
  sheet.appendRow(['MANT', 'Mantenimiento', 'Mantenimiento de equipos', 'SI']);
  sheet.appendRow(['LIMP', 'Limpieza', 'Proceso de limpieza', 'SI']);
  sheet.appendRow(['GENERAL', 'General', 'Proceso general (todos)', 'SI']);
}
  
  return 'Hojas inicializadas correctamente';
}

// Función para agregar datos de prueba
function agregarDatosPrueba() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Usuarios de prueba
  const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
  if (sheetUsuarios) {
    sheetUsuarios.appendRow(['123456', 'Juan Pérez', 'JEFE', 'Producción']);
    sheetUsuarios.appendRow(['789012', 'María García', 'OPERARIO', 'Limpieza']);
  }
  
  // Máquinas de prueba
  const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
  if (sheetMaquinas) {
    for (let i = 17; i <= 33; i++) {
      sheetMaquinas.appendRow([i.toString(), 'Maquina ' + i, 'Máquina de producción ' + i, 'SI']);
    }
  }
  
  // Elementos de prueba
  const sheetElementos = ss.getSheetByName(SHEETS.ELEMENTOS);
  if (sheetElementos) {
    for (let i = 17; i <= 33; i++) {
      for (let j = 1; j <= 3; j++) {
        sheetElementos.appendRow([i + '_' + j, i.toString(), 'Elemento ' + j, 'Área ' + j + ' de la máquina ' + i]);
      }
    }
  }
  
  return 'Datos de prueba agregados';
}

function diagnosticarRegistros() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja REGISTROS_LIMPIEZA no existe' };
    }
    
    const data = sheet.getDataRange().getValues();
    const resultados = {
      totalRegistros: data.length - 1,
      registros: []
    };
    
    console.log('=== DIAGNÓSTICO REGISTROS_LIMPIEZA ===');
    console.log('Total filas:', data.length);
    
    for (let i = 1; i < Math.min(data.length, 10); i++) { // Mostrar solo primeros 10
      if (data[i][0]) {
        resultados.registros.push({
          fila: i + 1,
          id: data[i][0],
          maquinaId: data[i][2],
          maquinaNombre: data[i][3],
          elementoId: data[i][4],
          elementoNombre: data[i][5],
          tipoLimpieza: data[i][6],
          estado: data[i][7]
        });
        
        console.log(`Fila ${i + 1}:`, {
          maquinaId: data[i][2],
          elementoId: data[i][4],
          tipo: data[i][6],
          estado: data[i][7]
        });
      }
    }
    
    return { success: true, diagnostico: resultados };
    
  } catch (error) {
    return { success: false, message: 'Error en diagnóstico: ' + error.message };
  }
}

function inicializarColumnasValidacion() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      console.error('❌ Hoja REGISTROS_LIMPIEZA no existe');
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    // Obtener todos los datos
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    console.log('📊 Estado actual de la hoja:');
    console.log('Filas:', data.length);
    console.log('Columnas:', data.length > 0 ? data[0].length : 0);
    
    if (data.length === 0) {
      // Hoja completamente vacía - crear headers completos
      const headersCompletos = [
        'ID', 'PLANEACION_ID', 'MAQUINA_ID', 'MAQUINA_NOMBRE', 
        'ELEMENTO_ID', 'ELEMENTO_NOMBRE', 'TIPO_LIMPIEZA', 'ESTADO',
        'RESPONSABLE', 'FECHA_REALIZACION', 'OBSERVACIONES', 
        'FECHA_CREACION', 'FECHA_FINALIZACION', 'COMPONENTE',
        'VALIDADO_POR', 'FECHA_VALIDACION'
      ];
      
      sheet.appendRow(headersCompletos);
      console.log('✅ Headers completos creados');
      return { success: true, message: 'Headers creados' };
    }
    
    const headers = data[0];
    console.log('Headers actuales:', headers);
    
    // Verificar columnas requeridas
    const columnasRequeridas = ['COMPONENTE', 'VALIDADO_POR', 'FECHA_VALIDACION'];
    let columnasFaltantes = [];
    
    columnasRequeridas.forEach(columna => {
      if (headers.indexOf(columna) === -1) {
        columnasFaltantes.push(columna);
      }
    });
    
    if (columnasFaltantes.length > 0) {
      console.log('📋 Columnas faltantes:', columnasFaltantes);
      
      // Agregar columnas faltantes
      columnasFaltantes.forEach(columna => {
        const newColIndex = headers.length + 1;
        sheet.getRange(1, newColIndex).setValue(columna);
        headers.push(columna);
        console.log(`✅ Columna ${columna} agregada en posición ${newColIndex}`);
      });
      
      return { 
        success: true, 
        message: `Columnas agregadas: ${columnasFaltantes.join(', ')}` 
      };
    } else {
      console.log('✅ Todas las columnas requeridas existen');
      return { success: true, message: 'Columnas OK' };
    }
    
  } catch (error) {
    console.error('💥 Error en inicializarColumnasValidacion:', error);
    return { success: false, message: 'Error: ' + error.message };
  }
}
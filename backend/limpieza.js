// ==================== FUNCIONES DE REGISTROS DE LIMPIEZA ====================
function obtenerRegistrosLimpiezaPorProceso(procesoUsuario, maquinaId, elementoId) {
  try {
    console.log('🔍 Buscando registros para proceso:', procesoUsuario);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetRegistros = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    
    if (!sheetRegistros) {
      return { success: true, registros: [] };
    }
    
    // Obtener máquinas permitidas para este proceso
    let maquinasPermitidas = [];
    if (sheetMaquinas) {
      const maquinasData = sheetMaquinas.getDataRange().getValues();
      const headers = maquinasData[0];
      const procesoCol = headers.indexOf('PROCESO');
      
      for (let i = 1; i < maquinasData.length; i++) {
        if (maquinasData[i][0]) {
          const maquinaId = maquinasData[i][0].toString();
          const procesoMaquina = procesoCol !== -1 ? 
            (maquinasData[i][procesoCol] || 'GENERAL').toString() : 'GENERAL';
          
          if (procesoMaquina === 'GENERAL' || procesoMaquina === procesoUsuario) {
            maquinasPermitidas.push(maquinaId);
          }
        }
      }
    }
    
    const data = sheetRegistros.getDataRange().getValues();
    const headers = data[0];
    
    // Buscar índice de columna RESPONSABLE_ASIGNADO
    const responsableAsignadoCol = headers.indexOf('RESPONSABLE_ASIGNADO');
    console.log('🔍 Columna RESPONSABLE_ASIGNADO en PorProceso:', responsableAsignadoCol);
    
    const registros = [];
    
    const maquinaIdBusqueda = maquinaId ? maquinaId.toString().trim() : null;
    const elementoIdBusqueda = elementoId ? elementoId.toString().trim() : null;
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === '') continue;
      
      const regMaquinaId = data[i][2] ? data[i][2].toString().trim() : '';
      
      // Primero filtrar por proceso
      if (maquinasPermitidas.length > 0 && !maquinasPermitidas.includes(regMaquinaId)) {
        continue; // Máquina no permitida para este proceso
      }
      
      const regElementoId = data[i][4] ? data[i][4].toString().trim() : '';
      
      const matchMaquina = !maquinaIdBusqueda || regMaquinaId === maquinaIdBusqueda;
      const matchElemento = !elementoIdBusqueda || regElementoId === elementoIdBusqueda;
      
      if (matchMaquina && matchElemento) {
        // Formatear fechas
        const fechaCreacion = data[i][11] ? formatearFechaCompleta(data[i][11]) : '';
        const fechaFinalizacion = data[i][12] ? formatearFechaCompleta(data[i][12]) : '';
        const fechaRealizacion = data[i][9] ? formatearFechaCompleta(data[i][9]) : '';
        const fechaValidacion = data[i][14] ? formatearFechaCompleta(data[i][14]) : '';
        
        // Obtener responsableAsignado
        const responsableAsignado = responsableAsignadoCol !== -1 && data[i][responsableAsignadoCol] ? 
                                  data[i][responsableAsignadoCol].toString().trim() : 'OPERARIO';
        
        const registro = {
          id: data[i][0].toString(),
          planeacionId: data[i][1] ? data[i][1].toString() : '',
          maquinaId: regMaquinaId,
          maquinaNombre: data[i][3] || '',
          elementoId: regElementoId,
          elementoNombre: data[i][5] || '',
          tipoLimpieza: data[i][6] || '',
          estado: data[i][7] || 'PENDIENTE',
          responsable: data[i][8] || '',
          fechaRealizacion: fechaRealizacion,
          observaciones: data[i][10] || '',
          fechaCreacion: fechaCreacion,
          fechaFinalizacion: fechaFinalizacion,
          componente: data[i][13] || '',
          validadoPor: data[i][14] || '',
          fechaValidacion: fechaValidacion,
          // NUEVO: Incluir responsableAsignado
          responsableAsignado: responsableAsignado
        };
        
        console.log('✅ REGISTRO COINCIDE:', {
          id: registro.id,
          elementoNombre: registro.elementoNombre,
          responsableAsignado: registro.responsableAsignado
        });
        
        registros.push(registro);
      }
    }
    
    console.log('🎯 Total registros encontrados:', registros.length);
    
    return { 
      success: true, 
      registros: registros,
      proceso: procesoUsuario,
      message: 'Registros filtrados por proceso'
    };
    
  } catch (error) {
    console.error('💥 Error en obtenerRegistrosLimpiezaPorProceso:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message, 
      registros: [] 
    };
  }
}

function obtenerRegistrosLimpieza(maquinaId, elementoId) {
  try {
    console.log('🔍 === INICIANDO BÚSQUEDA DE REGISTROS ===');
    console.log('Parámetros recibidos:', { maquinaId, elementoId });
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      console.log('❌ Hoja REGISTROS_LIMPIEZA no encontrada');
      return { success: true, registros: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    console.log('📊 Total filas en REGISTROS_LIMPIEZA:', data.length);
    
    if (data.length <= 1) {
      console.log('ℹ️ Solo hay encabezados en registros');
      return { success: true, registros: [] };
    }
    
    // Mostrar headers para referencia
    const headers = data[0];
    console.log('📝 Headers:', headers);
    
    // Buscar índice de columna RESPONSABLE_ASIGNADO
    const responsableAsignadoCol = headers.indexOf('RESPONSABLE_ASIGNADO');
    console.log('🔍 Columna RESPONSABLE_ASIGNADO:', responsableAsignadoCol);
    
    const registros = [];
    const maquinaIdBusqueda = maquinaId ? maquinaId.toString().trim() : null;
    const elementoIdBusqueda = elementoId ? elementoId.toString().trim() : null;
    
    console.log('🔍 Búsqueda con:', { 
      maquinaIdBusqueda, 
      elementoIdBusqueda 
    });
    
    let coincidencias = 0;
    
    for (let i = 1; i < data.length; i++) {
      const fila = i + 1;
      
      if (!data[i][0] || data[i][0].toString().trim() === '') {
        continue;
      }
      
      const regMaquinaId = data[i][2] ? data[i][2].toString().trim() : '';
      const regElementoId = data[i][4] ? data[i][4].toString().trim() : '';
      
      const matchMaquina = !maquinaIdBusqueda || regMaquinaId === maquinaIdBusqueda;
      const matchElemento = !elementoIdBusqueda || regElementoId === elementoIdBusqueda;
      
      if (matchMaquina && matchElemento) {
        // Formatear fechas
        const fechaCreacion = data[i][11] ? formatearFechaCompleta(data[i][11]) : '';
        const fechaFinalizacion = data[i][12] ? formatearFechaCompleta(data[i][12]) : '';
        const fechaRealizacion = data[i][9] ? formatearFechaCompleta(data[i][9]) : '';
        const fechaValidacion = data[i][14] ? formatearFechaCompleta(data[i][14]) : '';
        
        // Obtener responsableAsignado (columna 16 o la que corresponda)
        const responsableAsignado = responsableAsignadoCol !== -1 && data[i][responsableAsignadoCol] ? 
                                  data[i][responsableAsignadoCol].toString().trim() : 'OPERARIO';
        
        const registro = {
          id: data[i][0].toString(),
          planeacionId: data[i][1] ? data[i][1].toString() : '',
          maquinaId: regMaquinaId,
          maquinaNombre: data[i][3] || '',
          elementoId: regElementoId,
          elementoNombre: data[i][5] || '',
          tipoLimpieza: data[i][6] || '',
          estado: data[i][7] || 'PENDIENTE',
          responsable: data[i][8] || '',
          fechaRealizacion: fechaRealizacion,
          observaciones: data[i][10] || '',
          fechaCreacion: fechaCreacion,
          fechaFinalizacion: fechaFinalizacion,
          componente: data[i][13] || '',
          validadoPor: data[i][14] || '',
          fechaValidacion: fechaValidacion,
          // NUEVO: Incluir responsableAsignado en la respuesta
          responsableAsignado: responsableAsignado
        };
        
        console.log(`✅ REGISTRO COINCIDE ${++coincidencias}:`, {
          id: registro.id,
          elementoNombre: registro.elementoNombre,
          tipoLimpieza: registro.tipoLimpieza,
          estado: registro.estado,
          responsableAsignado: registro.responsableAsignado
        });
        
        registros.push(registro);
      }
    }
    
    console.log('🎯 RESULTADO BÚSQUEDA:');
    console.log('- Total registros encontrados:', registros.length);
    
    return { 
      success: true, 
      registros: registros,
      message: `${registros.length} registros encontrados`
    };
    
  } catch (error) {
    console.error('💥 Error crítico en obtenerRegistrosLimpieza:', error);
    return { 
      success: false, 
      message: 'Error al obtener registros: ' + error.message, 
      registros: [] 
    };
  }
}

// Función específica para debug de un elemento
function debugRegistrosElemento(maquinaId, elementoId) {
  try {
    console.log('🐛 DEBUG ESPECÍFICO para elemento:', { maquinaId, elementoId });
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const resultados = [];
    
    const maquinaIdBusqueda = maquinaId.toString().trim();
    const elementoIdBusqueda = elementoId.toString().trim();
    
    console.log('🔍 Búsqueda con IDs:', { maquinaIdBusqueda, elementoIdBusqueda });
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      const regMaquinaId = data[i][2] ? data[i][2].toString().trim() : '';
      const regElementoId = data[i][4] ? data[i][4].toString().trim() : '';
      
      console.log(`Fila ${i + 1}:`, {
        regMaquinaId: regMaquinaId,
        regElementoId: regElementoId,
        matchMaquina: regMaquinaId === maquinaIdBusqueda,
        matchElemento: regElementoId === elementoIdBusqueda
      });
      
      if (regMaquinaId === maquinaIdBusqueda && regElementoId === elementoIdBusqueda) {
        resultados.push({
          fila: i + 1,
          id: data[i][0],
          maquinaId: regMaquinaId,
          elementoId: regElementoId,
          elementoNombre: data[i][5],
          tipoLimpieza: data[i][6],
          estado: data[i][7]
        });
      }
    }
    
    return { 
      success: true, 
      resultados: resultados,
      total: resultados.length,
      busqueda: { maquinaIdBusqueda, elementoIdBusqueda }
    };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function formatearFechaCompleta(fecha) {
  if (!fecha) return '';
  
  try {
    let fechaObj;
    
    // Convertir a objeto Date
    if (typeof fecha === 'string') {
      fechaObj = new Date(fecha);
    } else if (fecha instanceof Date) {
      fechaObj = fecha;
    } else {
      // Intentar parsear
      fechaObj = new Date(fecha);
    }
    
    // Verificar si es fecha válida
    if (isNaN(fechaObj.getTime())) {
      console.warn('⚠️ Fecha inválida en formatearFechaCompleta:', fecha);
      return '';
    }
    
    // Formatear a yyyy-MM-dd para inputs type="date"
    const año = fechaObj.getFullYear();
    const mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
    const día = String(fechaObj.getDate()).padStart(2, '0');
    
    return `${año}-${mes}-${día}`;
    
  } catch (e) {
    console.warn('⚠️ Error formateando fecha:', e, 'Valor:', fecha);
    return '';
  }
}
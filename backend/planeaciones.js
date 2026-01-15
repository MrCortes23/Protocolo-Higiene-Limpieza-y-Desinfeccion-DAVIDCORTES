// ==================== FUNCIONES DE PLANEACIÓN ====================
function guardarPlaneacion(datos) {
  try {
    console.log('💾 Guardando planeación con responsable:', datos.responsable);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PLANEACION);
    
    if (!sheet) {
      return { success: false, message: 'Hoja PLANEACION no encontrada' };
    }
    
    const fechaCreacion = new Date();
    const id = Utilities.getUuid();
    
    // Verificar última columna para ver si ya existe RESPONSABLE
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Si no existe columna RESPONSABLE, agregarla
    if (headers.indexOf('RESPONSABLE') === -1) {
      sheet.getRange(1, lastColumn + 1).setValue('RESPONSABLE');
    }
    
    const responsableCol = headers.indexOf('RESPONSABLE') !== -1 ? 
                         headers.indexOf('RESPONSABLE') + 1 : lastColumn + 1;
    
    // Preparar fila
    const nuevaFila = [
      id,
      datos.maquinaId || '',
      datos.maquinaNombre || '',
      datos.frecuencia || 'Mensual',
      datos.limpiezaSeco ? 'SI' : 'NO',
      datos.limpiezaHumedo ? 'SI' : 'NO',
      datos.desinfeccion ? 'SI' : 'NO',
      JSON.stringify(datos.elementosConfig || []),
      fechaCreacion,
      datos.usuarioCreador || 'Sistema',
      'ACTIVA'
    ];
    
    // Agregar fila
    sheet.appendRow(nuevaFila);
    
    // Si se agregó columna nueva, llenar responsable en la fila recién agregada
    if (responsableCol > nuevaFila.length) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, responsableCol).setValue(datos.responsable || 'OPERARIO');
    }
    
    console.log('✅ Planeación guardada con ID:', id);
    
    // Crear registros de limpieza pendientes - PASANDO EL RESPONSABLE
    const registrosCreados = crearRegistrosPendientesJerarquicos(datos, id);
    
    if (registrosCreados > 0) {
      return { 
        success: true, 
        message: 'Planeación guardada correctamente', 
        id: id,
        registrosCreados: registrosCreados
      };
    } else {
      return { 
        success: false, 
        message: 'Planeación guardada pero no se crearon registros de limpieza',
        id: id
      };
    }
    
  } catch (error) {
    console.error('💥 Error guardando planeación:', error);
    return { success: false, message: 'Error al guardar planeación: ' + error.message };
  }
}

function crearRegistrosPendientesJerarquicos(datos, planeacionId) {
  try {
    console.log('📝 === INICIANDO CREAR REGISTROS MEJORADO ===');
    console.log('Planeacion ID:', planeacionId);
    console.log('Datos completos:', datos);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      console.error('❌ CRÍTICO: Hoja REGISTROS_LIMPIEZA no encontrada');
      return 0;
    }
    
    console.log('✅ Hoja REGISTROS_LIMPIEZA encontrada');
    
    // Verificar si la hoja tiene datos
    const lastRow = sheet.getLastRow();
    console.log('Última fila en hoja:', lastRow);
    
    // Inicializar columnas
    const initResult = inicializarColumnasValidacion();
    console.log('Resultado inicialización columnas:', initResult);
    
    const fechaCreacion = new Date();
    const elementosConfig = datos.elementosConfig || [];
    
    console.log('📊 ELEMENTOS CONFIG RECIBIDOS:', elementosConfig);
    console.log('Tipo de elementosConfig:', typeof elementosConfig);
    console.log('Es array?:', Array.isArray(elementosConfig));
    console.log('Longitud:', elementosConfig.length);
    
    if (!Array.isArray(elementosConfig)) {
      console.error('❌ elementosConfig no es un array:', elementosConfig);
      return 0;
    }
    
    if (elementosConfig.length === 0) {
      console.error('❌ elementosConfig está vacío');
      return 0;
    }
    
    let registrosCreados = 0;
    
    // PROCESAR CADA ITEM
    elementosConfig.forEach((item, index) => {
      console.log(`\n--- PROCESANDO ITEM ${index + 1} ---`);
      console.log('Item completo:', item);
      console.log('Tipo del item:', typeof item);
      
      let elementos = [];
      let componenteNombre = 'Componente PRINCIPAL';
      
      // ANÁLISIS DETALLADO DE LA ESTRUCTURA
      if (item && typeof item === 'object') {
        console.log('✅ Item es un objeto');
        
        if (item.elementos && Array.isArray(item.elementos)) {
          console.log('✅ Tiene propiedad "elementos" como array');
          elementos = item.elementos;
          componenteNombre = item.componenteNombre || 'Componente PRINCIPAL';
          console.log(`🔧 Componente: "${componenteNombre}"`);
          console.log(`🔧 Número de elementos: ${elementos.length}`);
        } else if (Array.isArray(item)) {
          console.log('✅ Item es directamente un array de elementos');
          elementos = item;
          componenteNombre = 'COMPONENTES';
        } else {
          console.log('❌ Estructura no reconocida en item:', Object.keys(item));
        }
      } else {
        console.log('❌ Item no es un objeto válido:', item);
      }
      
      // PROCESAR ELEMENTOS
      console.log(`🔄 Procesando ${elementos.length} elementos...`);
      
      elementos.forEach((elemento, elemIndex) => {
        console.log(`\n    📋 ELEMENTO ${elemIndex + 1}:`, elemento);
        console.log('    Tipo del elemento:', typeof elemento);
        
        if (elemento && typeof elemento === 'object') {
          // EXTRAER DATOS CON MÚLTIPLAS OPCIONES
          const elementoId = elemento.elementoId || elemento.id || '';
          const elementoNombre = elemento.elementoNombre || elemento.nombre || '';
          
          console.log(`    🔍 ID: "${elementoId}", Nombre: "${elementoNombre}"`);
          
          if (!elementoId || !elementoNombre) {
            console.log('    ❌ Elemento sin ID o Nombre válido');
            console.log('    Claves disponibles:', Object.keys(elemento));
            return;
          }
          
          // DETERMINAR TIPOS DE LIMPIEZA CON MÚLTIPLAS VERIFICACIONES
          const tiposLimpieza = [];
          
          // Verificar seco
          if (elemento.seco === true || elemento.seco === 'true' || elemento.seco === 'SI' || elemento.seco === 1) {
            tiposLimpieza.push('SECO');
            console.log('    ✅ Limpieza SECO activada');
          }
          
          // Verificar humedo
          if (elemento.humedo === true || elemento.humedo === 'true' || elemento.humedo === 'SI' || elemento.humedo === 1) {
            tiposLimpieza.push('HUMEDO');
            console.log('    ✅ Limpieza HUMEDO activada');
          }
          
          // Verificar desinfeccion
          if (elemento.desinfeccion === true || elemento.desinfeccion === 'true' || elemento.desinfeccion === 'SI' || elemento.desinfeccion === 1) {
            tiposLimpieza.push('DESINFECCION');
            console.log('    ✅ Desinfección activada');
          }
          
          console.log(`    🧹 Tipos finales: ${tiposLimpieza.join(', ')}`);
          
          if (tiposLimpieza.length === 0) {
            console.log('    ⚠️  Elemento sin tipos de limpieza activados');
            console.log('    Valores:', {
              seco: elemento.seco,
              humedo: elemento.humedo,
              desinfeccion: elemento.desinfeccion
            });
          }
          
          // CREAR REGISTROS
          tiposLimpieza.forEach(tipo => {
            const registroId = Utilities.getUuid();
            
            const nuevaFila = [
              registroId,                    // ID
              planeacionId,                  // PLANEACION_ID
              datos.maquinaId || '',         // MAQUINA_ID
              datos.maquinaNombre || '',     // MAQUINA_NOMBRE
              elementoId,                    // ELEMENTO_ID
              elementoNombre,                // ELEMENTO_NOMBRE
              tipo,                          // TIPO_LIMPIEZA
              'PENDIENTE',                   // ESTADO
              '',                            // RESPONSABLE (se llena al hacer limpieza)
              '',                            // FECHA_REALIZACION
              '',                            // OBSERVACIONES
              fechaCreacion,                 // FECHA_CREACION
              '',                            // FECHA_FINALIZACION
              componenteNombre,              // COMPONENTE
              '',                            // VALIDADO_POR
              '',                            // FECHA_VALIDACION
              elemento.responsable || datos.responsable || 'OPERARIO' // ← NUEVO: RESPONSABLE_ASIGNADO
            ];

            
            console.log(`    ➕ CREANDO REGISTRO: ${elementoNombre} - ${tipo}`);
            console.log('    Fila a agregar:', nuevaFila);
            
            try {
              sheet.appendRow(nuevaFila);
              registrosCreados++;
              console.log(`    ✅ Registro ${registrosCreados} creado exitosamente`);
            } catch (appendError) {
              console.error('    ❌ Error al agregar fila:', appendError);
            }
          });
        } else {
          console.log('    ❌ Elemento no es un objeto válido:', elemento);
        }
      });
    });
    
    console.log('\n=== RESUMEN FINAL ===');
    console.log(`📈 Registros totales creados: ${registrosCreados}`);
    
    if (registrosCreados === 0) {
      console.log('❌ FALLO CRÍTICO: No se crearon registros');
      console.log('📋 Datos originales recibidos:', {
        planeacionId: planeacionId,
        maquinaId: datos.maquinaId,
        maquinaNombre: datos.maquinaNombre,
        elementosConfig: datos.elementosConfig
      });
    }
    
    return registrosCreados;
    
  } catch (error) {
    console.error('💥 ERROR CRÍTICO EN crearRegistrosPendientesMejorado:', error);
    console.error('Stack trace:', error.stack);
    return 0;
  }
}

function obtenerPlaneacionesPorProceso(procesoUsuario = 'GENERAL') {
  try {
    console.log('🔍 Iniciando obtención de planeaciones para proceso:', procesoUsuario);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    
    if (!sheetPlaneacion) {
      console.log('❌ Hoja PLANEACION no encontrada');
      return { success: true, planeaciones: [] };
    }
    
    // Obtener todas las máquinas primero para filtrar por proceso
    let maquinasPermitidas = [];
    if (sheetMaquinas) {
      const maquinasData = sheetMaquinas.getDataRange().getValues();
      const headers = maquinasData[0];
      const procesoCol = headers.indexOf('PROCESO');
      const maquinaIdCol = headers.indexOf('ID');
      
      console.log('📊 Procesando máquinas. Total:', maquinasData.length - 1);
      
      // Si no existe columna PROCESO, mostrar todas las máquinas
      if (procesoCol === -1) {
        console.log('⚠️ No hay columna PROCESO en MAQUINAS, mostrando todas');
        // Obtener todos los IDs de máquinas
        for (let i = 1; i < maquinasData.length; i++) {
          if (maquinasData[i][maquinaIdCol]) {
            maquinasPermitidas.push(maquinasData[i][maquinaIdCol].toString().trim());
          }
        }
      } else {
        // Filtrar máquinas por proceso
        for (let i = 1; i < maquinasData.length; i++) {
          if (maquinasData[i][maquinaIdCol]) {
            const maquinaId = maquinasData[i][maquinaIdCol].toString().trim();
            const procesoMaquina = maquinasData[i][procesoCol] ? 
              maquinasData[i][procesoCol].toString().trim() : 'GENERAL';
            
            // Incluir máquinas con proceso GEN (General) o que coincidan con el usuario
            if (procesoMaquina === 'GENERAL' || procesoMaquina === procesoUsuario) {
              maquinasPermitidas.push(maquinaId);
              console.log(`✅ Máquina ${maquinaId} permitida (proceso: ${procesoMaquina})`);
            } else {
              console.log(`❌ Máquina ${maquinaId} excluida (proceso: ${procesoMaquina} ≠ ${procesoUsuario})`);
            }
          }
        }
      }
    } else {
      console.log('⚠️ Hoja MAQUINAS no encontrada, no se puede filtrar por proceso');
    }
    
    const data = sheetPlaneacion.getDataRange().getValues();
    console.log('📊 Datos crudos de planeación:', data.length, 'filas');
    
    // Si solo hay encabezados
    if (data.length <= 1) {
      console.log('ℹ️ Solo hay encabezados en planeación');
      return { 
        success: true, 
        planeaciones: [],
        message: 'No hay planeaciones registradas',
        procesoUsuario: procesoUsuario
      };
    }
    
    const planeaciones = [];
    let totalPlaneaciones = 0;
    let planeacionesFiltradas = 0;
    
    for (let i = 1; i < data.length; i++) {
      totalPlaneaciones++;
      
      // Verificar que hay datos válidos en la fila (al menos ID)
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        const maquinaId = data[i][1] ? data[i][1].toString().trim() : '';
        
        // Filtrar por proceso si hay máquinas filtradas
        if (maquinasPermitidas.length > 0 && !maquinasPermitidas.includes(maquinaId)) {
          console.log(`⏭️ Planeación ${data[i][0]} omitida (máquina ${maquinaId} no permitida)`);
          continue; // Saltar planeación de máquina no permitida
        }
        
        let elementosConfig = [];
        try {
          const configStr = data[i][7] || '[]';
          if (typeof configStr === 'string' && configStr.trim() !== '') {
            elementosConfig = JSON.parse(configStr);
          }
        } catch (e) {
          console.warn('⚠️ Error parseando elementosConfig fila', i + 1, e);
          elementosConfig = [];
        }
        
        const planeacion = {
          id: data[i][0] ? data[i][0].toString().trim() : 'ID_' + i,
          maquinaId: maquinaId,
          maquinaNombre: data[i][2] ? data[i][2].toString().trim() : 'Sin nombre',
          frecuencia: data[i][3] ? data[i][3].toString().trim() : 'Mensual',
          limpiezaSeco: data[i][4] === 'SI',
          limpiezaHumedo: data[i][5] === 'SI',
          desinfeccion: data[i][6] === 'SI',
          elementosConfig: elementosConfig,
          fechaCreacion: data[i][8] ? new Date(data[i][8]).toISOString() : new Date().toISOString(),
          usuarioCreador: data[i][9] ? data[i][9].toString().trim() : 'Sistema',
          estado: data[i][10] ? data[i][10].toString().trim() : 'ACTIVA',
          // Información adicional útil
          procesoAsignado: obtenerProcesoMaquina(maquinaId) || 'GENERAL'
        };
        
        console.log('✅ Planeación incluida:', {
          nombre: planeacion.maquinaNombre,
          id: planeacion.id,
          proceso: planeacion.procesoAsignado
        });
        
        planeaciones.push(planeacion);
        planeacionesFiltradas++;
        
      } else {
        console.log('❌ Fila', i + 1, 'sin ID válido, omitiendo');
      }
    }
    
    console.log('🎯 Resumen planeaciones:');
    console.log('- Total en sistema:', totalPlaneaciones);
    console.log('- Filtradas por proceso:', planeacionesFiltradas);
    console.log('- Proceso usuario:', procesoUsuario);
    console.log('- Máquinas permitidas:', maquinasPermitidas.length);
    
    return { 
      success: true, 
      planeaciones: planeaciones,
      message: `Planeaciones cargadas: ${planeacionesFiltradas} de ${totalPlaneaciones}`,
      procesoUsuario: procesoUsuario,
      estadisticas: {
        total: totalPlaneaciones,
        filtradas: planeacionesFiltradas,
        maquinasPermitidas: maquinasPermitidas.length
      }
    };
    
  } catch (error) {
    console.error('💥 Error crítico en obtenerPlaneaciones:', error);
    return { 
      success: false, 
      message: 'Error al obtener planeaciones: ' + error.message, 
      planeaciones: [],
      procesoUsuario: procesoUsuario || 'GENERAL'
    };
  }
}

// Función auxiliar para obtener el proceso de una máquina
function obtenerProcesoMaquina(maquinaId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    
    if (!sheetMaquinas) return 'GENERAL';
    
    const data = sheetMaquinas.getDataRange().getValues();
    const headers = data[0];
    
    const maquinaIdCol = headers.indexOf('ID');
    const procesoCol = headers.indexOf('PROCESO');
    
    if (maquinaIdCol === -1 || procesoCol === -1) return 'GENERAL';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][maquinaIdCol] && data[i][maquinaIdCol].toString().trim() === maquinaId.toString().trim()) {
        return data[i][procesoCol] ? data[i][procesoCol].toString().trim() : 'GENERAL';
      }
    }
    
    return 'GENERAL';
  } catch (error) {
    console.warn('⚠️ Error obteniendo proceso de máquina:', error);
    return 'GENERAL';
  }
}

function eliminarPlaneacionesMaquina(maquinaId) {
  try {
    console.log('🗑️ Eliminando planeaciones de máquina:', maquinaId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    const sheetRegistros = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheetPlaneacion) {
      return { success: false, message: 'Hoja PLANEACION no encontrada' };
    }
    
    const planeacionesData = sheetPlaneacion.getDataRange().getValues();
    let eliminadas = 0;
    let planeacionesAEliminar = [];
    
    // Encontrar planeaciones a eliminar
    for (let i = 1; i < planeacionesData.length; i++) {
      if (planeacionesData[i][1] && planeacionesData[i][1].toString().trim() === maquinaId.toString().trim()) {
        planeacionesAEliminar.push({
          fila: i + 1,
          id: planeacionesData[i][0],
          maquinaNombre: planeacionesData[i][2]
        });
      }
    }
    
    // Eliminar en orden inverso para no afectar índices
    planeacionesAEliminar.reverse().forEach(planeacion => {
      sheetPlaneacion.deleteRow(planeacion.fila);
      eliminadas++;
      console.log(`✅ Planeación eliminada: ${planeacion.id} (${planeacion.maquinaNombre})`);
    });
    
    // También eliminar registros de limpieza asociados
    let registrosEliminados = 0;
    if (sheetRegistros && eliminadas > 0) {
      const registrosData = sheetRegistros.getDataRange().getValues();
      let registrosAEliminar = [];
      
      for (let i = 1; i < registrosData.length; i++) {
        if (registrosData[i][2] && registrosData[i][2].toString().trim() === maquinaId.toString().trim()) {
          registrosAEliminar.push(i + 1);
        }
      }
      
      // Eliminar en orden inverso
      registrosAEliminar.reverse().forEach(fila => {
        sheetRegistros.deleteRow(fila);
        registrosEliminados++;
      });
      
      console.log(`🗑️ Eliminados ${registrosEliminados} registros de limpieza`);
    }
    
    return { 
      success: true, 
      message: `Eliminadas ${eliminadas} planeación(es) y ${registrosEliminados} registro(s) de limpieza`,
      eliminadas: eliminadas,
      registrosEliminados: registrosEliminados
    };
    
  } catch (error) {
    console.error('💥 Error eliminando planeaciones:', error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

function obtenerTodasLasPlaneaciones() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPlaneacion = ss.getSheetByName(SHEETS.PLANEACION);
    
    if (!sheetPlaneacion) {
      return { success: true, planeaciones: [] };
    }
    
    const data = sheetPlaneacion.getDataRange().getValues();
    
    // Si solo hay encabezados
    if (data.length <= 1) {
      return { 
        success: true, 
        planeaciones: [],
        message: 'No hay planeaciones registradas'
      };
    }
    
    const planeaciones = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        const maquinaId = data[i][1] ? data[i][1].toString().trim() : '';
        
        let elementosConfig = [];
        try {
          const configStr = data[i][7] || '[]';
          if (typeof configStr === 'string' && configStr.trim() !== '') {
            elementosConfig = JSON.parse(configStr);
          }
        } catch (e) {
          elementosConfig = [];
        }
        
        // Obtener proceso de la máquina
        const procesoAsignado = obtenerProcesoMaquina(maquinaId) || 'GENERAL';
        
        const planeacion = {
          id: data[i][0].toString(),
          maquinaId: maquinaId,
          maquinaNombre: data[i][2] || 'Sin nombre',
          frecuencia: data[i][3] || 'Mensual',
          limpiezaSeco: data[i][4] === 'SI',
          limpiezaHumedo: data[i][5] === 'SI',
          desinfeccion: data[i][6] === 'SI',
          elementosConfig: elementosConfig,
          fechaCreacion: data[i][8] ? new Date(data[i][8]).toISOString() : new Date().toISOString(),
          usuarioCreador: data[i][9] || 'Sistema',
          estado: data[i][10] || 'ACTIVA',
          procesoAsignado: procesoAsignado
        };
        
        planeaciones.push(planeacion);
      }
    }
    
    return { 
      success: true, 
      planeaciones: planeaciones,
      message: `Total planeaciones: ${planeaciones.length}`
    };
    
  } catch (error) {
    console.error('💥 Error en obtenerTodasLasPlaneaciones:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message, 
      planeaciones: [] 
    };
  }
}

function obtenerTodosRegistrosLimpieza() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: true, registros: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: true, registros: [] };
    }
    
    const registros = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === '') continue;
      
      // Formatear fechas
      const fechaCreacion = data[i][11] ? formatearFechaCompleta(data[i][11]) : '';
      const fechaFinalizacion = data[i][12] ? formatearFechaCompleta(data[i][12]) : '';
      const fechaRealizacion = data[i][9] ? formatearFechaCompleta(data[i][9]) : '';
      const fechaValidacion = data[i][15] ? formatearFechaCompleta(data[i][15]) : '';
      
      const registro = {
        id: data[i][0].toString(),
        planeacionId: data[i][1] || '',
        maquinaId: data[i][2] ? data[i][2].toString().trim() : '',
        maquinaNombre: data[i][3] || '',
        elementoId: data[i][4] ? data[i][4].toString().trim() : '',
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
        fechaValidacion: fechaValidacion
      };
      
      registros.push(registro);
    }
    
    return { 
      success: true, 
      registros: registros,
      message: `Total registros: ${registros.length}`
    };
    
  } catch (error) {
    console.error('💥 Error en obtenerTodosRegistrosLimpieza:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message, 
      registros: [] 
    };
  }
}
function validarLimpiezaCompleta(maquinaId, elementoId, validadorNombre) {
  try {
    console.log('👑 Iniciando validación de limpieza completa...', {
      maquinaId: maquinaId,
      elementoId: elementoId,
      validadorNombre: validadorNombre
    });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja REGISTROS_LIMPIEZA no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const idCol = headers.indexOf('ID');
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const elementoIdCol = headers.indexOf('ELEMENTO_ID');
    const estadoCol = headers.indexOf('ESTADO');
    const validadoPorCol = headers.indexOf('VALIDADO_POR');
    const fechaValidacionCol = headers.indexOf('FECHA_VALIDACION');
    
    // Si no existen las columnas de validación, las agregamos
    let needsHeaderUpdate = false;
    if (validadoPorCol === -1) {
      sheet.getRange(1, headers.length + 1).setValue('VALIDADO_POR');
      needsHeaderUpdate = true;
    }
    if (fechaValidacionCol === -1) {
      sheet.getRange(1, headers.length + (needsHeaderUpdate ? 2 : 1)).setValue('FECHA_VALIDACION');
      needsHeaderUpdate = true;
    }
    
    // Si actualizamos headers, refrescamos los datos
    if (needsHeaderUpdate) {
      const newData = sheet.getDataRange().getValues();
      const newHeaders = newData[0];
      
      // Reasignar índices con las nuevas columnas
      const newValidadoPorCol = newHeaders.indexOf('VALIDADO_POR');
      const newFechaValidacionCol = newHeaders.indexOf('FECHA_VALIDACION');
      
      var finalValidadoPorCol = newValidadoPorCol;
      var finalFechaValidacionCol = newFechaValidacionCol;
    } else {
      var finalValidadoPorCol = validadoPorCol;
      var finalFechaValidacionCol = fechaValidacionCol;
    }
    
    let registrosActualizados = 0;
    let registrosPendientes = 0;
    const fechaActual = new Date();
    
    // Buscar y actualizar registros
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar que tenemos los datos básicos
      if (!row[maquinaIdCol] || !row[elementoIdCol]) continue;
      
      const rowMaquinaId = row[maquinaIdCol].toString().trim();
      const rowElementoId = row[elementoIdCol].toString().trim();
      
      // Verificar coincidencia
      if (rowMaquinaId === maquinaId.toString().trim() && 
          rowElementoId === elementoId.toString().trim()) {
        
        // Verificar que esté completado
        if (row[estadoCol] === 'COMPLETADO') {
          // Actualizar validación
          if (finalValidadoPorCol !== -1) {
            sheet.getRange(i + 1, finalValidadoPorCol + 1).setValue(validadorNombre);
          }
          if (finalFechaValidacionCol !== -1) {
            sheet.getRange(i + 1, finalFechaValidacionCol + 1).setValue(fechaActual);
          }
          registrosActualizados++;
          console.log(`✅ Registro ${row[idCol]} validado`);
        } else {
          registrosPendientes++;
          console.log(`❌ Registro ${row[idCol]} no está COMPLETADO, estado: ${row[estadoCol]}`);
        }
      }
    }
    
    if (registrosActualizados > 0) {
      return { 
        success: true, 
        message: `Limpieza validada correctamente por ${validadorNombre}. ${registrosActualizados} registros actualizados.`,
        registrosActualizados: registrosActualizados,
        registrosPendientes: registrosPendientes
      };
    } else if (registrosPendientes > 0) {
      return { 
        success: false, 
        message: `No se puede validar. ${registrosPendientes} registro(s) no están completados.` 
      };
    } else {
      return { 
        success: false, 
        message: 'No se encontraron registros para validar' 
      };
    }
    
  } catch (error) {
    console.error('💥 Error validando limpieza:', error);
    return { 
      success: false, 
      message: 'Error del sistema al validar: ' + error.message 
    };
  }
}

function validarMaquinaCompletaPorJefe(maquinaId, validador) {
  try {
    console.log('👑 JEFE: Validando máquina completa...', {
      maquinaId: maquinaId,
      validador: validador
    });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja REGISTROS_LIMPIEZA no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: 'No hay registros en la hoja' };
    }
    
    const headers = data[0];
    console.log('📋 Headers encontrados:', headers);
    
    // Buscar índices de columnas con nombres exactos
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const estadoCol = headers.indexOf('ESTADO');
    const validadoPorCol = headers.indexOf('VALIDADO_POR');
    const fechaValidacionCol = headers.indexOf('FECHA_VALIDACION');
    
    console.log('🔍 Índices de columnas:');
    console.log(`- MAQUINA_ID: ${maquinaIdCol}`);
    console.log(`- ESTADO: ${estadoCol}`);
    console.log(`- VALIDADO_POR: ${validadoPorCol}`);
    console.log(`- FECHA_VALIDACION: ${fechaValidacionCol}`);
    
    // Verificar columnas críticas
    if (maquinaIdCol === -1) {
      return { success: false, message: 'No se encontró la columna MAQUINA_ID' };
    }
    
    if (estadoCol === -1) {
      return { success: false, message: 'No se encontró la columna ESTADO' };
    }
    
    let registrosActualizados = 0;
    let registrosPendientes = 0;
    let registrosYaValidados = 0;
    let registrosNoCompletados = 0;
    const fechaActual = new Date();
    
    console.log(`🔍 Buscando registros de máquina ID: ${maquinaId}`);
    console.log(`📊 Total filas a revisar: ${data.length - 1}`);
    
    // Buscar y actualizar registros de esta máquina
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar que es de esta máquina
      if (!row[maquinaIdCol]) {
        continue; // Saltar filas sin MAQUINA_ID
      }
      
      const rowMaquinaId = row[maquinaIdCol].toString().trim();
      
      if (rowMaquinaId === maquinaId.toString().trim()) {
        console.log(`📝 Fila ${i + 1}: Máquina ${rowMaquinaId}`);
        
        // Obtener estado
        const estado = row[estadoCol] ? row[estadoCol].toString().trim() : '';
        console.log(`   Estado: "${estado}"`);
        
        if (estado === 'COMPLETADO') {
          
          // Verificar si ya está validado
          let yaValidado = false;
          if (validadoPorCol !== -1 && row[validadoPorCol]) {
            const validadorActual = row[validadoPorCol].toString().trim();
            yaValidado = validadorActual !== '';
            console.log(`   Validado por: "${validadorActual}" (Ya validado: ${yaValidado})`);
          }
          
          if (!yaValidado) {
            // Actualizar validación si no está ya validado
            if (validadoPorCol !== -1) {
              sheet.getRange(i + 1, validadoPorCol + 1).setValue(validador);
              console.log(`   ✅ Asignado validador: ${validador}`);
            }
            
            if (fechaValidacionCol !== -1) {
              sheet.getRange(i + 1, fechaValidacionCol + 1).setValue(fechaActual);
              console.log(`   📅 Asignada fecha: ${fechaActual}`);
            }
            
            registrosActualizados++;
          } else {
            registrosYaValidados++;
            console.log(`   ℹ️ Ya validado, no se modifica`);
          }
        } else if (estado === 'PENDIENTE' || estado === 'EN PROCESO' || estado === '') {
          registrosNoCompletados++;
          console.log(`   ❌ No completado (${estado})`);
        } else {
          console.log(`   ⚠️ Estado desconocido: ${estado}`);
        }
      }
    }
    
    console.log('📈 RESULTADOS FINALES:');
    console.log(`- Registros actualizados: ${registrosActualizados}`);
    console.log(`- Registros ya validados: ${registrosYaValidados}`);
    console.log(`- Registros no completados: ${registrosNoCompletados}`);
    
    // Determinar mensaje final
    if (registrosActualizados > 0) {
      return { 
        success: true, 
        message: `✅ Máquina validada correctamente por ${validador}. 
                  ${registrosActualizados} registros actualizados.
                  ${registrosYaValidados} ya estaban validados.
                  ${registrosNoCompletados} no estaban completados.`,
        registrosActualizados: registrosActualizados,
        registrosYaValidados: registrosYaValidados,
        registrosNoCompletados: registrosNoCompletados
      };
    } else if (registrosYaValidados > 0 && registrosNoCompletados === 0) {
      return { 
        success: false, 
        message: `ℹ️ Todos los registros (${registrosYaValidados}) ya están validados. No hay nada nuevo que validar.`
      };
    } else if (registrosNoCompletados > 0) {
      return { 
        success: false, 
        message: `❌ No se puede validar. ${registrosNoCompletados} registro(s) no están completados.`
      };
    } else {
      return { 
        success: false, 
        message: 'ℹ️ No se encontraron registros de esta máquina para validar.'
      };
    }
    
  } catch (error) {
    console.error('💥 Error validando máquina completa:', error);
    return { 
      success: false, 
      message: 'Error del sistema: ' + error.message 
    };
  }
}

function actualizarRegistroLimpiezaCompleto(registroId, datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTROS_LIMPIEZA);
    
    if (!sheet) {
      return { success: false, message: 'Hoja REGISTROS_LIMPIEZA no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('ID');
    const estadoCol = headers.indexOf('ESTADO');
    const responsableCol = headers.indexOf('RESPONSABLE');
    const fechaRealizacionCol = headers.indexOf('FECHA_REALIZACION');
    const observacionesCol = headers.indexOf('OBSERVACIONES');
    const fechaFinalizacionCol = headers.indexOf('FECHA_FINALIZACION');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] && data[i][idCol].toString() === registroId.toString()) {
        // Actualizar campos básicos
        sheet.getRange(i + 1, estadoCol + 1).setValue(datos.estado || 'PENDIENTE');
        sheet.getRange(i + 1, responsableCol + 1).setValue(datos.responsable || '');
        sheet.getRange(i + 1, fechaRealizacionCol + 1).setValue(datos.fechaRealizacion || '');
        sheet.getRange(i + 1, observacionesCol + 1).setValue(datos.observaciones || '');
        
        // Si se marca como COMPLETADO, registrar fecha de finalización
        if (datos.estado === 'COMPLETADO') {
          sheet.getRange(i + 1, fechaFinalizacionCol + 1).setValue(new Date());
        }
        
        return { success: true, message: 'Registro actualizado correctamente' };
      }
    }
    
    return { success: false, message: 'Registro no encontrado' };
  } catch (error) {
    return { success: false, message: 'Error al actualizar registro: ' + error.message };
  }
}
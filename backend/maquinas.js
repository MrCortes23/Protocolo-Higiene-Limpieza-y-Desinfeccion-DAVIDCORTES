// ==================== FUNCIONES DE MÁQUINAS ====================
function obtenerMaquinas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.MAQUINAS);
    if (!sheet) {
      return { success: false, message: 'Hoja MAQUINAS no encontrada', maquinas: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    
    const maquinas = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        maquinas.push({
          id: data[i][0].toString(),
          nombre: data[i][1] || '',
          descripcion: data[i][2] || '',
          activa: data[i][3] || 'SI'
        });
      }
    }
    
    return { success: true, maquinas: maquinas };
  } catch (error) {
    return { success: false, message: 'Error al obtener máquinas: ' + error.message, maquinas: [] };
  }
}

function obtenerElementosPorMaquina(maquinaId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ELEMENTOS);
    if (!sheet) {
      return { success: false, message: 'Hoja ELEMENTOS no encontrada', elementos: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    
    const elementos = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === maquinaId.toString()) {
        elementos.push({
          id: data[i][0].toString(),
          maquinaId: data[i][1].toString(),
          nombre: data[i][2] || '',
          descripcion: data[i][3] || ''
        });
      }
    }
    
    return { success: true, elementos: elementos };
  } catch (error) {
    return { success: false, message: 'Error al obtener elementos: ' + error.message, elementos: [] };
  }
}

function obtenerMaquinasConElementosPorProceso(procesoUsuario) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    const sheetElementos = ss.getSheetByName(SHEETS.ELEMENTOS);
    
    if (!sheetMaquinas) {
      return { success: false, message: 'Hoja MAQUINAS no encontrada', maquinas: [] };
    }
    
    const maquinasData = sheetMaquinas.getDataRange().getValues();
    const elementosData = sheetElementos ? sheetElementos.getDataRange().getValues() : [];
    
    // Obtener índices de columnas
    const headers = maquinasData[0];
    const procesoCol = headers.indexOf('PROCESO');
    
    // Si no existe columna PROCESO, mostrar todas
    const mostrarTodas = procesoCol === -1;
    
    // Crear mapa de elementos por máquina y componente
    const elementosPorMaquina = {};
    for (let i = 1; i < elementosData.length; i++) {
      if (elementosData[i][0] && elementosData[i][1]) {
        const maquinaId = elementosData[i][1].toString();
        const componenteId = elementosData[i][2] ? elementosData[i][2].toString() : 'principal';
        
        if (!elementosPorMaquina[maquinaId]) {
          elementosPorMaquina[maquinaId] = {};
        }
        if (!elementosPorMaquina[maquinaId][componenteId]) {
          elementosPorMaquina[maquinaId][componenteId] = [];
        }
        
        elementosPorMaquina[maquinaId][componenteId].push({
          id: elementosData[i][0].toString(),
          maquinaId: maquinaId,
          componenteId: componenteId,
          nombre: elementosData[i][3] || '',
          descripcion: elementosData[i][4] || ''
        });
      }
    }
    
    // Crear lista de máquinas filtradas por proceso
    const maquinasConElementos = [];
    for (let i = 1; i < maquinasData.length; i++) {
      if (maquinasData[i][0]) {
        const maquinaId = maquinasData[i][0].toString();
        
        // Filtrar por proceso si existe la columna
        if (!mostrarTodas) {
          const procesoMaquina = maquinasData[i][procesoCol] ? maquinasData[i][procesoCol].toString().trim() : 'GENERAL';
          
          // Mostrar máquinas con proceso GEN (General) o que coincidan con el usuario
          if (procesoMaquina !== 'GENERAL' && procesoMaquina !== procesoUsuario) {
            continue; // Saltar esta máquina
          }
        }
        
        const componentes = [];
        
        // Si hay elementos para esta máquina, organizarlos por componente
        if (elementosPorMaquina[maquinaId]) {
          Object.keys(elementosPorMaquina[maquinaId]).forEach(componenteId => {
            const elementos = elementosPorMaquina[maquinaId][componenteId];
            componentes.push({
              id: componenteId,
              nombre: componenteId === 'principal' ? 'COMPONENTES' : componenteId,
              elementos: elementos
            });
          });
        }
        
        maquinasConElementos.push({
          id: maquinaId,
          nombre: maquinasData[i][1] || '',
          descripcion: maquinasData[i][2] || '',
          activa: maquinasData[i][3] || 'SI',
          proceso: mostrarTodas ? 'GENERAL' : (maquinasData[i][procesoCol] || 'GENERAL'),
          componentes: componentes
        });
      }
    }
    
    return { 
      success: true, 
      maquinas: maquinasConElementos,
      procesoUsuario: procesoUsuario,
      totalFiltradas: maquinasConElementos.length
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message, maquinas: [] };
  }
}
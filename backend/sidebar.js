// ==================== FUNCIONES PARA OBTENER ESTRUCTURA JERÁRQUICA ====================

function obtenerComponentesPorMaquina(maquinaId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetElementos = ss.getSheetByName(SHEETS.ELEMENTOS);
    
    if (!sheetElementos) {
      console.log('⚠️ Hoja ELEMENTOS no encontrada');
      return [];
    }
    
    const data = sheetElementos.getDataRange().getValues();
    const headers = data[0];
    
    // Buscar índices de columnas
    const maquinaIdCol = headers.indexOf('MAQUINA_ID');
    const componenteCol = headers.indexOf('COMPONENTE');
    const elementoIdCol = headers.indexOf('ID');
    const elementoNombreCol = headers.indexOf('NOMBRE');
    
    // Si no existe columna COMPONENTE, usar estructura simple
    const hasComponentes = componenteCol !== -1;
    
    const componentesMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar que el elemento pertenezca a esta máquina
      if (row[maquinaIdCol] && row[maquinaIdCol].toString().trim() === maquinaId.toString().trim()) {
        const componenteNombre = hasComponentes && row[componenteCol] ? 
          row[componenteCol].toString().trim() : 'COMPONENTES';
        
        // Si no existe el componente en el mapa, crearlo
        if (!componentesMap.has(componenteNombre)) {
          componentesMap.set(componenteNombre, {
            id: componenteNombre.toLowerCase().replace(/\s+/g, '-'),
            nombre: componenteNombre,
            elementos: []
          });
        }
        
        // Agregar elemento al componente
        const elemento = {
          id: row[elementoIdCol] ? row[elementoIdCol].toString().trim() : '',
          nombre: row[elementoNombreCol] || 'Sin nombre',
          descripcion: row[headers.indexOf('DESCRIPCION')] || ''
        };
        
        componentesMap.get(componenteNombre).elementos.push(elemento);
      }
    }
    
    // Convertir mapa a array
    const componentes = Array.from(componentesMap.values());
    
    // Si no hay componentes, crear uno genérico
    if (componentes.length === 0) {
      componentes.push({
        id: 'principal',
        nombre: 'COMPONENTES',
        elementos: []
      });
    }
    
    console.log(`✅ Componentes para máquina ${maquinaId}: ${componentes.length}`);
    
    return componentes;
    
  } catch (error) {
    console.error('💥 Error en obtenerComponentesPorMaquina:', error);
    return [];
  }
}

function obtenerTodasLasMaquinasConElementos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetMaquinas = ss.getSheetByName(SHEETS.MAQUINAS);
    
    if (!sheetMaquinas) {
      return { success: false, message: 'Hoja MAQUINAS no encontrada', maquinas: [] };
    }
    
    const maquinasData = sheetMaquinas.getDataRange().getValues();
    const headers = maquinasData[0];
    
    const maquinas = [];
    
    // Obtener índices de columnas
    const idCol = headers.indexOf('ID');
    const nombreCol = headers.indexOf('NOMBRE');
    const procesoCol = headers.indexOf('PROCESO');
    
    for (let i = 1; i < maquinasData.length; i++) {
      const row = maquinasData[i];
      
      if (row[idCol]) {
        const maquinaId = row[idCol].toString().trim();
        const maquina = {
          id: maquinaId,
          nombre: row[nombreCol] || 'Sin nombre',
          procesoAsignado: procesoCol !== -1 ? (row[procesoCol] || 'GENERAL').toString().trim() : 'GENERAL',
          descripcion: row[headers.indexOf('DESCRIPCION')] || '',
          activa: row[headers.indexOf('ACTIVA')] || 'SI',
          componentes: obtenerComponentesPorMaquina(maquinaId) // Usar la nueva función
        };
        
        maquinas.push(maquina);
      }
    }
    
    console.log(`✅ Todas las máquinas obtenidas: ${maquinas.length}`);
    
    return { 
      success: true, 
      maquinas: maquinas,
      message: `Total máquinas: ${maquinas.length}`
    };
    
  } catch (error) {
    console.error('💥 Error en obtenerTodasLasMaquinasConElementos:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message, 
      maquinas: [] 
    };
  }
}

function obtenerMaquinasConElementos(procesoUsuario = 'GENERAL') {
  try {
    // Si el usuario es GEN, obtener todas las máquinas
    if (procesoUsuario === 'GENERAL') {
      return obtenerTodasLasMaquinasConElementos();
    }
    
    // Si no es GEN, usar la función existente
    return obtenerMaquinasConElementosPorProceso(procesoUsuario);
    
  } catch (error) {
    console.error('💥 Error en obtenerMaquinasConElementos:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.message, 
      maquinas: [] 
    };
  }
}
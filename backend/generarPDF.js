function generarReporteMaquinaPDF(maquinaId, maquinaNombre, jefeOrigenNombre, jefeOrigenCedula, turnoOrigen, turnoDestino, proceso, emailDestino) {
  try {
    console.log('📊 Iniciando generación de reporte PDF...');
    
    // Validar parámetros
    if (!emailDestino || emailDestino.trim() === '') {
      throw new Error('Correo destino no especificado');
    }
    
    // 1. Obtener información de la máquina
    const infoMaquina = obtenerInfoValidacionMaquina(maquinaId);
    
    if (!infoMaquina.success) {
      throw new Error('Error al obtener información de la máquina: ' + infoMaquina.message);
    }
    
    const datosMaquina = infoMaquina.datos;
    
    // 2. Verificar si puede generar reporte
    if (!datosMaquina.puedeValidar && datosMaquina.yaValidados === 0) {
      return { 
        success: false, 
        message: `No se puede generar reporte. La máquina no está completamente validada.` 
      };
    }
    
    // 3. Generar HTML del reporte SIMPLE
    const htmlContent = generarHTMLReporteAvances(
      maquinaId,
      maquinaNombre,
      datosMaquina,
      jefeOrigenNombre,
      turnoOrigen,
      turnoDestino,
      proceso,
      emailDestino
    );
    
    // 4. Crear PDF
    const pdfBlob = crearPDF(htmlContent);
    
    // 5. Enviar email
    enviarEmailReporteAvances(
      emailDestino, 
      pdfBlob, 
      maquinaId,
      maquinaNombre,
      datosMaquina,
      jefeOrigenNombre,
      turnoOrigen,
      turnoDestino
    );
    
    // 6. Registrar el envío
    registrarEnvioReporte(
      maquinaId,
      maquinaNombre,
      jefeOrigenNombre,
      jefeOrigenCedula,
      turnoOrigen,
      turnoDestino,
      emailDestino,
      datosMaquina
    );
    
    console.log('✅ Reporte PDF simple generado y enviado');
    
    return { 
      success: true, 
      message: `✅ Reporte de avances enviado a ${emailDestino}`,
      emailDestino: emailDestino
    };
    
  } catch (error) {
    console.error('💥 Error generando reporte:', error);
    return { 
      success: false, 
      message: 'Error al generar reporte: ' + error.message 
    };
  }
}

function generarHTMLReporteAvances(maquinaId, maquinaNombre, datosMaquina, jefeOrigen, turnoOrigen, turnoDestino, proceso, emailDestino) {
  const fecha = new Date();
  const fechaFormateada = fecha.toLocaleDateString('es-ES', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Reporte Avances - ${maquinaNombre}</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
          font-size: 12px;
          color: #333;
        }
        
        .container {
          max-width: 600px;
          margin: 0 auto;
          background: white;
          border: 1px solid #ddd;
          padding: 20px;
        }
        
        .header {
          text-align: center;
          margin-bottom: 20px;
          padding-bottom: 15px;
          border-bottom: 2px solid #1e40af;
        }
        
        .header h1 {
          color: #1e40af;
          margin: 0 0 5px 0;
          font-size: 18px;
        }
        
        .header h2 {
          color: #374151;
          margin: 0;
          font-size: 14px;
          font-weight: normal;
        }
        
        .info-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 10px;
          margin: 15px 0;
          font-size: 11px;
        }
        
        .info-item {
          padding: 8px;
          background: #f8f9fa;
          border-radius: 4px;
        }
        
        .info-label {
          font-weight: bold;
          color: #555;
          margin-bottom: 3px;
        }
        
        .stats-box {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 8px;
          padding: 15px;
          margin: 20px 0;
        }
        
        .stats-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 15px;
        }
        
        .stat-item {
          text-align: center;
          padding: 10px;
          background: white;
          border-radius: 6px;
          border: 1px solid #e5e7eb;
        }
        
        .stat-number {
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 5px;
        }
        
        .stat-total { color: #1e40af; }
        .stat-completed { color: #10b981; }
        .stat-pending { color: #f59e0b; }
        .stat-validated { color: #8b5cf6; }
        
        .stat-label {
          font-size: 11px;
          color: #64748b;
          text-transform: uppercase;
        }
        
        .progress-section {
          margin: 20px 0;
        }
        
        .progress-bar {
          height: 20px;
          background: #e5e7eb;
          border-radius: 10px;
          margin: 10px 0;
          overflow: hidden;
        }
        
        .progress-fill {
          height: 100%;
          background: linear-gradient(90deg, #3b82f6, #1d4ed8);
          border-radius: 10px;
          text-align: center;
          color: white;
          font-size: 12px;
          line-height: 20px;
          font-weight: bold;
        }
        
        .summary {
          background: #f0f9ff;
          border-left: 4px solid #0ea5e9;
          padding: 15px;
          margin: 20px 0;
          border-radius: 0 8px 8px 0;
        }
        
        .summary h3 {
          margin: 0 0 10px 0;
          color: #0369a1;
          font-size: 13px;
        }
        
        .summary ul {
          margin: 0;
          padding-left: 20px;
          color: #0c4a6e;
        }
        
        .summary li {
          margin-bottom: 5px;
          font-size: 11px;
        }
        
        .footer {
          text-align: center;
          margin-top: 25px;
          padding-top: 15px;
          border-top: 1px solid #e5e7eb;
          color: #6b7280;
          font-size: 10px;
        }
        
        .footer strong {
          color: #1e40af;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <!-- Header -->
        <div class="header">
          <h1>📋 REPORTE DE AVANCES</h1>
          <h2>${maquinaNombre} (ID: ${maquinaId})</h2>
          <div style="font-size: 11px; color: #6b7280; margin-top: 5px;">
            Generado: ${fechaFormateada}
          </div>
        </div>
        
        <!-- Información básica -->
        <div class="info-grid">
          <div class="info-item">
            <div class="info-label">Proceso</div>
            <div>${proceso || 'GENERAL'}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Jefe Responsable</div>
            <div>${jefeOrigen}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Turno Origen</div>
            <div>${turnoOrigen}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Turno Destino</div>
            <div>${turnoDestino} (${emailDestino})</div>
          </div>
        </div>
        
        <!-- Estadísticas principales -->
        <div class="stats-box">
          <div class="stats-grid">
            <div class="stat-item">
              <div class="stat-number stat-total">${datosMaquina.total}</div>
              <div class="stat-label">Total Elementos</div>
            </div>
            <div class="stat-item">
              <div class="stat-number stat-completed">${datosMaquina.completados}</div>
              <div class="stat-label">Completados</div>
            </div>
            <div class="stat-item">
              <div class="stat-number stat-pending">${datosMaquina.pendientes}</div>
              <div class="stat-label">Pendientes</div>
            </div>
            <div class="stat-item">
              <div class="stat-number stat-validated">${datosMaquina.yaValidados}</div>
              <div class="stat-label">Validados</div>
            </div>
          </div>
        </div>
        
        <!-- Barra de progreso -->
        <div class="progress-section">
          <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
            <span style="font-weight: bold; color: #374151;">Progreso General</span>
            <span style="font-weight: bold; color: #1e40af;">${datosMaquina.porcentaje}%</span>
          </div>
          <div class="progress-bar">
            <div class="progress-fill" style="width: ${datosMaquina.porcentaje}%">
              ${datosMaquina.porcentaje}%
            </div>
          </div>
          <div style="text-align: center; font-size: 11px; color: #6b7280; margin-top: 5px;">
            ${datosMaquina.completados} de ${datosMaquina.total} elementos
          </div>
        </div>
        
        <!-- Resumen ejecutivo -->
        <div class="summary">
          <h3>📊 RESUMEN EJECUTIVO</h3>
          <ul>
            <li><strong>Estado actual:</strong> ${datosMaquina.porcentaje >= 100 ? '✅ COMPLETADO' : 
                                                   datosMaquina.porcentaje >= 70 ? '⚡ EN PROCESO AVANZADO' : 
                                                   datosMaquina.porcentaje >= 30 ? '🔄 EN PROCESO' : '⏳ PENDIENTE'}</li>
            <li><strong>Avance:</strong> ${datosMaquina.completados}/${datosMaquina.total} elementos completados</li>
            <li><strong>Validación:</strong> ${datosMaquina.yaValidados} elementos validados por jefe</li>
            <li><strong>Transferencia:</strong> ${turnoOrigen} → ${turnoDestino}</li>
            <li><strong>Responsable:</strong> ${jefeOrigen}</li>
          </ul>
        </div>
        
        <!-- Footer -->
        <div class="footer">
          <div><strong>PHLYD</strong> - Sistema de Gestión de Limpieza</div>
          <div>Reporte generado automáticamente</div>
          <div style="margin-top: 5px; font-size: 9px;">
            Este documento es una transferencia oficial entre turnos
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
}

function crearPDF(htmlContent) {
  try {
    console.log('📄 Creando PDF...');
    
    // Crear blob con el HTML
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'reporte.html');
    
    // Convertir a PDF
    const pdf = blob.getAs('application/pdf');
    
    // Nombre del archivo
    const fecha = new Date();
    const nombreArchivo = `Reporte_Avances_${fecha.getFullYear()}${(fecha.getMonth()+1).toString().padStart(2,'0')}${fecha.getDate().toString().padStart(2,'0')}_${fecha.getHours()}${fecha.getMinutes()}.pdf`;
    
    pdf.setName(nombreArchivo);
    console.log('✅ PDF creado:', nombreArchivo);
    
    return pdf;
  } catch (error) {
    console.error('💥 Error creando PDF:', error);
    throw error;
  }
}

function enviarEmailReporteAvances(emailDestino, pdfBlob, maquinaId, maquinaNombre, datosMaquina, jefeOrigen, turnoOrigen, turnoDestino) {
  try {
    console.log('📧 Enviando email a:', emailDestino);
    
    const fecha = new Date().toLocaleDateString('es-ES');
    const subject = `📋 Reporte de Avances - ${maquinaNombre} (${turnoOrigen} → ${turnoDestino})`;
    
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #1e40af; border-bottom: 2px solid #1e40af; padding-bottom: 10px;">
          📋 Reporte de Avances - Sistema PHLYD
        </h2>
        
        <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin: 20px 0;">
          <h3 style="color: #374151; margin-top: 0;">
            ${maquinaNombre}
          </h3>
          
          <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 15px 0;">
            <div>
              <strong>ID Máquina:</strong><br>
              ${maquinaId}
            </div>
            <div>
              <strong>Progreso:</strong><br>
              ${datosMaquina.porcentaje}% (${datosMaquina.completados}/${datosMaquina.total})
            </div>
            <div>
              <strong>Jefe Origen:</strong><br>
              ${jefeOrigen}
            </div>
            <div>
              <strong>Turno:</strong><br>
              ${turnoOrigen} → ${turnoDestino}
            </div>
          </div>
        </div>
        
        <div style="background: #dbeafe; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <h4 style="color: #1e40af; margin-top: 0;">📊 Resumen de Estado</h4>
          <ul style="margin: 10px 0; padding-left: 20px;">
            <li>Total elementos: ${datosMaquina.total}</li>
            <li>Completados: ${datosMaquina.completados}</li>
            <li>Pendientes: ${datosMaquina.pendientes}</li>
            <li>Ya validados: ${datosMaquina.yaValidados}</li>
            <li>Progreso general: ${datosMaquina.porcentaje}%</li>
          </ul>
        </div>
        
        <p>Se adjunta reporte PDF detallado con la información completa de la máquina.</p>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e5e7eb; color: #6b7280; font-size: 12px;">
          <p><em>Este es un mensaje automático generado por el Sistema PHLYD.</em></p>
          <p><em>Fecha de envío: ${fecha}</em></p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailDestino,
      subject: subject,
      htmlBody: body,
      attachments: [pdfBlob],
      name: 'Sistema PHLYD - Reportes'
    });
    
    console.log('✅ Email enviado exitosamente');
    return true;
    
  } catch (error) {
    console.error('💥 Error enviando email:', error);
    throw error;
  }
}

function registrarReporteMaquina(maquinaId, maquinaNombre, jefeOrigen, jefeOrigenCedula, turnoOrigen, turnoDestino, emailDestino) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetReportes = ss.getSheetByName('REPORTES_MAQUINAS');
    
    if (!sheetReportes) {
      sheetReportes = ss.insertSheet('REPORTES_MAQUINAS');
      sheetReportes.appendRow(['FECHA', 'MAQUINA_ID', 'MAQUINA_NOMBRE', 'JEFE_ORIGEN', 'TURNO_ORIGEN', 'TURNO_DESTINO', 'EMAIL_DESTINO', 'ESTADO']);
    }
    
    sheetReportes.appendRow([
      new Date().toISOString(),
      maquinaId,
      maquinaNombre,
      jefeOrigen,
      turnoOrigen,
      turnoDestino,
      emailDestino,
      'ENVIADO'
    ]);
    
  } catch (error) {
    console.error('Error registrando reporte:', error);
  }
}

function calcularEstadisticasPorTipo(registros) {
  const tipos = {};
  
  registros.forEach(registro => {
    const tipo = registro.tipoLimpieza || 'NO ESPECIFICADO';
    
    if (!tipos[tipo]) {
      tipos[tipo] = {
        total: 0,
        completados: 0,
        porcentaje: 0
      };
    }
    
    tipos[tipo].total++;
    
    if (registro.estado === 'COMPLETADO' || registro.estado === 'VALIDADO') {
      tipos[tipo].completados++;
    }
  });
  
  // Calcular porcentajes
  Object.keys(tipos).forEach(tipo => {
    const stats = tipos[tipo];
    stats.porcentaje = stats.total > 0 ? 
      Math.round((stats.completados / stats.total) * 100) : 0;
  });
  
  return tipos;
}

function registrarEnvioReporte(maquinaId, maquinaNombre, jefeOrigen, jefeCedula, turnoOrigen, turnoDestino, emailDestino, datosMaquina) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Crear o obtener hoja de reportes
    let sheet = ss.getSheetByName('ENVIOS_REPORTES');
    if (!sheet) {
      sheet = ss.insertSheet('ENVIOS_REPORTES');
      sheet.appendRow([
        'FECHA_ENVIO',
        'MAQUINA_ID',
        'MAQUINA_NOMBRE',
        'JEFE_NOMBRE',
        'JEFE_CEDULA',
        'TURNO_ORIGEN',
        'TURNO_DESTINO',
        'EMAIL_DESTINO',
        'TOTAL_ELEMENTOS',
        'COMPLETADOS',
        'PENDIENTES',
        'YA_VALIDADOS',
        'PORCENTAJE',
        'ESTADO_ENVIO'
      ]);
      
      // Formatear headers
      const range = sheet.getRange(1, 1, 1, 14);
      range.setBackground('#1e40af')
           .setFontColor('white')
           .setFontWeight('bold');
    }
    
    // Agregar registro
    sheet.appendRow([
      new Date().toISOString(),
      maquinaId,
      maquinaNombre,
      jefeOrigen,
      jefeCedula,
      turnoOrigen,
      turnoDestino,
      emailDestino,
      datosMaquina.total,
      datosMaquina.completados,
      datosMaquina.pendientes,
      datosMaquina.yaValidados,
      datosMaquina.porcentaje,
      'ENVIADO'
    ]);
    
    console.log('📝 Registro de envío guardado');
    return true;
    
  } catch (error) {
    console.error('⚠️ Error guardando registro de envío:', error);
    // No lanzamos error para no afectar el flujo principal
    return false;
  }
}
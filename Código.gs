/**
 * IACI · Auditoría & Control de Deuda - Backend Engine
 * VERSIÓN ACTUALIZADA: Con Seguridad por Whitelist Dinámica
 * Built by BondiApps. 2026
 */

const ID_PLANILLA_MADRE = "1oFE4TnnZkpJFDQIOQHTHB9tk5RGX0U7Q4oM9FMm6NzI"; 
const SS_LOCAL = SpreadsheetApp.getActiveSpreadsheet();

function doGet() {
  // --- SEGURIDAD: VERIFICACIÓN DE WHITELIST ---
  var userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  if (!userEmail) userEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();

  var sheetWhitelist = SS_LOCAL.getSheetByName('Whitelist');
  var accesoPermitido = false;

  if (sheetWhitelist) {
    var dataWhitelist = sheetWhitelist.getRange(1, 1, sheetWhitelist.getLastRow(), 1).getValues();
    var listaEmails = dataWhitelist.map(function(fila) {
      return fila[0].toString().toLowerCase().trim();
    });
    accesoPermitido = listaEmails.indexOf(userEmail) !== -1;
  }

  // Si no está en la lista, bloqueamos el acceso inmediatamente
  if (!accesoPermitido) {
    return HtmlService.createHtmlOutput(renderPantallaDenegada(userEmail))
        .setTitle('Acceso Denegado')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // --------------------------------------------

  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('IACI · Auditoría')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTML para el bloqueo de acceso
 */
function renderPantallaDenegada(emailDetectado) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Inter', sans-serif; display: flex; height: 100vh; justify-content: center; align-items: center; background: #f8f9fa; margin: 0; color: #333; }
          .card { background: white; padding: 40px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align: center; max-width: 420px; width: 90%; border: 1px solid #eee; }
          h1 { color: #dc3545; margin: 15px 0 10px 0; font-size: 22px; font-weight: 600; }
          p { color: #6c757d; line-height: 1.6; margin-bottom: 20px; }
          .icon { font-size: 50px; display: block; margin-bottom: 10px; }
          .contact { font-size: 0.85rem; background: #f1f3f5; padding: 12px; border-radius: 8px; color: #495057; border: 1px solid #e9ecef; }
          .detected { font-size: 0.7rem; color: #ced4da; margin-top: 20px; border-top: 1px solid #f1f3f5; padding-top: 10px; }
        </style>
      </head>
      <body>
        <div class="card">
          <span class="icon">⛔</span>
          <h1>Acceso denegado</h1>
          <p><strong>Easy IACI</strong> tiene acceso restringido por razones de seguridad.</p>
          <div class="contact">
            Si es un error, por favor comunícate con el proveedor:<br>
            <strong>bondiapps.com</strong>
          </div>
          <div class="detected">ID Detectado: ${emailDetectado || "No identificado"}</div>
        </div>
      </body>
    </html>
  `;
}

/**
 * Función auxiliar para normalizar texto
 */
function normalizar(txt) {
  if (!txt) return "";
  return txt.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * MOTOR DE AUDITORÍA: Cruza Alumnos (EXTERNOS) vs Movimientos y Precios
 */
function obtenerMatrizAuditoria() {
  const ssMadre = SpreadsheetApp.openById(ID_PLANILLA_MADRE);
  const hojaAlu = ssMadre.getSheetByName("DB_ALUMNOS");
  const hojaMov = ssMadre.getSheetByName("DB_MOVIMIENTOS");
  
  let hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  
  if (!hojaAlu || !hojaMov) {
    throw new Error("No se encontró DB_ALUMNOS o DB_MOVIMIENTOS en la Planilla Madre.");
  }
  
  if (!hojaPrecios) {
    crearPestañaConfigPrecios();
    hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  }

  const alumnosRaw = hojaAlu.getDataRange().getValues();
  const movsRaw = hojaMov.getDataRange().getValues();
  const datosPrecios = hojaPrecios.getDataRange().getValues();
  
  let mapaPrecios = {};
  const cabeceraPrecios = datosPrecios[0];
  datosPrecios.slice(1).forEach(fila => {
    let cursoNom = normalizar(fila[0]);
    mapaPrecios[cursoNom] = {};
    for (let i = 1; i < cabeceraPrecios.length; i++) {
      mapaPrecios[cursoNom][normalizar(cabeceraPrecios[i])] = parseFloat(fila[i]) || 0;
    }
  });

  const alumnos = alumnosRaw.slice(1);
  const movs = movsRaw.slice(1);    
  
  const MESES_CONTROL = ["MATRICULA", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  let pagosPorAlumno = {}; 
  
  movs.forEach(m => {
    let monto = parseFloat(m[5]) || 0;
    let mesPagoStr = normalizar(m[7]);
    let estado = normalizar(m[8]);
    let idAlu = normalizar(m[9]);
    
    if ((estado.includes("COMPLETADO") || estado === "OK") && idAlu !== "") {
      if (!pagosPorAlumno[idAlu]) pagosPorAlumno[idAlu] = {};
      let conceptos = mesPagoStr.split(",").map(s => s.trim());
      let montoDividido = monto / (conceptos.length || 1);
      
      conceptos.forEach(con => {
        let conNorm = normalizar(con);
        pagosPorAlumno[idAlu][conNorm] = (pagosPorAlumno[idAlu][conNorm] || 0) + montoDividido;
      });
    }
  });

  let matriz = alumnos
    .filter(al => al[0] && al[0].toString().trim() !== "") 
    .map(al => {
      let id = normalizar(al[0]); 
      let curso = normalizar(al[3] || "SIN CURSO");
      
      let cuotaFamiliar = parseFloat(al[16]) || 0;
      let cuotaRef = parseFloat(al[4]) || 0; 
      
      let misPagosRealizados = pagosPorAlumno[id] || {};
      let montosPorMes = {};
      let deudaCalculada = 0;

      MESES_CONTROL.forEach(mes => {
        let mesNorm = normalizar(mes);
        let pagadoReal = misPagosRealizados[mesNorm] || 0;
        montosPorMes[mes] = pagadoReal;

        let precioOficialCurso = (mapaPrecios[curso] && mapaPrecios[curso][mesNorm] > 0) 
                          ? mapaPrecios[curso][mesNorm] : 0;
        
        let precioAplicable = (cuotaFamiliar > 0) ? cuotaFamiliar : (precioOficialCurso > 0 ? precioOficialCurso : cuotaRef);
        
        if (pagadoReal < (precioAplicable - 15)) {
          deudaCalculada += (precioAplicable - pagadoReal);
        }
      });

      return {
        id: al[0].toString().trim(), 
        nombre: `${al[1]}, ${al[2]}`, 
        curso: al[3],
        cuota: cuotaFamiliar > 0 ? cuotaFamiliar : cuotaRef, 
        esFamiliar: cuotaFamiliar > 0,
        pagos: montosPorMes,
        totalDeuda: Math.round(deudaCalculada),
        tieneDeuda: deudaCalculada > 15
      };
    });

  return { meses: MESES_CONTROL, datos: matriz, timestamp: Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm") };
}

/**
 * GESTIÓN DE CURSOS: DASHBOARD ESTRUCTURAL Y REVENUE
 */
function obtenerDashboardCursos() {
  const ssMadre = SpreadsheetApp.openById(ID_PLANILLA_MADRE);
  const hojaAlu = ssMadre.getSheetByName("DB_ALUMNOS");
  const hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS") || crearPestañaConfigPrecios();
  
  const datosAlu = hojaAlu.getDataRange().getValues().slice(1);
  const datosPrecios = hojaPrecios.getDataRange().getValues();
  
  let conteoAlumnos = {};
  datosAlu.forEach(fila => {
    let c = fila[3] || "SIN CURSO";
    conteoAlumnos[c] = (conteoAlumnos[c] || 0) + 1;
  });

  let dashboard = datosPrecios.slice(1).map(fila => {
    let nombreCurso = fila[0];
    let cant = conteoAlumnos[nombreCurso] || 0;
    let arancelRef = parseFloat(fila[2]) || parseFloat(fila[1]) || 0; 
    let revenueEstimado = cant * arancelRef;

    return {
      nombre: nombreCurso,
      matricula: fila[1],
      cuotaMarzo: fila[2],
      alumnos: cant,
      revenueMensual: revenueEstimado
    };
  });

  return dashboard;
}

/**
 * CREAR NUEVO CURSO DESDE LA APP
 */
function crearNuevoCurso(datos) {
  let hoja = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  if (!hoja) hoja = crearPestañaConfigPrecios();
  
  const nombreNuevo = datos.nombre.trim();
  const valores = hoja.getDataRange().getValues();
  
  for (let i = 1; i < valores.length; i++) {
    if (normalizar(valores[i][0]) === normalizar(nombreNuevo)) {
      return "Error: El curso ya existe.";
    }
  }

  let nuevaFila = [nombreNuevo, parseFloat(datos.matricula) || 0];
  for (let i = 0; i < 10; i++) {
    nuevaFila.push(parseFloat(datos.cuotaMensual) || 0);
  }
  
  hoja.appendRow(nuevaFila);
  sincronizarColumnaE();
  return "✅ Curso '" + nombreNuevo + "' creado exitosamente.";
}

/**
 * FUNCIONES DEL EDITOR CON SINCRONIZACIÓN DINÁMICA
 */
function actualizarMontoRango(datos) {
  let hoja = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  if (!hoja) return "Error: No se encontró la hoja.";
  const valores = hoja.getDataRange().getValues();
  const cabecera = valores[0];
  const idxInicio = cabecera.indexOf(datos.mesInicio);
  const idxFin = cabecera.indexOf(datos.mesFin);
  let filaIndex = -1;
  for (let i = 1; i < valores.length; i++) {
    if (valores[i][0] === datos.curso) { filaIndex = i + 1; break; }
  }
  if (filaIndex === -1 || idxInicio === -1 || idxFin === -1) return "Error de ubicación.";
  const numColumnas = idxFin - idxInicio + 1;
  hoja.getRange(filaIndex, idxInicio + 1, 1, numColumnas).setValues([new Array(numColumnas).fill(parseFloat(datos.monto))]);
  
  sincronizarColumnaE();
  return `✅ Actualizado e Impactado en Planilla Madre: ${datos.curso}`;
}

function sincronizarColumnaE() {
  const ssMadre = SpreadsheetApp.openById(ID_PLANILLA_MADRE);
  const hojaAlu = ssMadre.getSheetByName("DB_ALUMNOS");
  const hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  
  const datosPrecios = hojaPrecios.getDataRange().getValues();
  const cabeceraPrecios = datosPrecios[0];
  
  const hoy = new Date();
  const mesNro = hoy.getMonth(); 
  let mesNombre = "";
  
  if (mesNro === 0 || mesNro === 1) {
    mesNombre = "MATRICULA";
  } else {
    const nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    mesNombre = nombres[mesNro];
  }

  const colIdx = cabeceraPrecios.indexOf(mesNombre);
  if (colIdx === -1) return; 

  let preciosHoy = {};
  datosPrecios.slice(1).forEach(f => {
    preciosHoy[normalizar(f[0])] = f[colIdx];
  });

  const aluData = hojaAlu.getDataRange().getValues();
  for (let i = 1; i < aluData.length; i++) {
    let cursoAlu = normalizar(aluData[i][3]);
    let tieneFamiliar = parseFloat(aluData[i][16]) || 0; 
    
    if (tieneFamiliar === 0 && preciosHoy[cursoAlu] !== undefined) {
      hojaAlu.getRange(i + 1, 5).setValue(preciosHoy[cursoAlu]);
    }
  }
}

function obtenerCursosYMeses() {
  const ssMadre = SpreadsheetApp.openById(ID_PLANILLA_MADRE);
  const hojaAlu = ssMadre.getSheetByName("DB_ALUMNOS");
  let hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  if (!hojaPrecios) { crearPestañaConfigPrecios(); hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS"); }
  const datosAlu = hojaAlu.getDataRange().getValues();
  const cursosUnicos = [...new Set(datosAlu.slice(1).map(fila => fila[3]))].filter(c => c && c.toString().trim() !== "").sort();
  const meses = hojaPrecios.getDataRange().getValues()[0].slice(1);
  return { cursos: cursosUnicos, meses: meses };
}

function obtenerGrillaPrecios() {
  let hoja = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  if (!hoja) { crearPestañaConfigPrecios(); hoja = SS_LOCAL.getSheetByName("CONFIG_PRECIOS"); }
  const datos = hoja.getDataRange().getValues();
  return { cabecera: datos[0], valores: datos.slice(1) };
}

function guardarGrillaPrecios(matrizCompleta) {
  let hoja = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  hoja.clearContents();
  hoja.getRange(1, 1, matrizCompleta.length, matrizCompleta[0].length).setValues(matrizCompleta);
  sincronizarColumnaE(); 
  return "Precios actualizados localmente e impactados en base.";
}

function crearPestañaConfigPrecios() {
  let hojaPrecios = SS_LOCAL.getSheetByName("CONFIG_PRECIOS");
  if (hojaPrecios) return hojaPrecios;
  
  const ssMadre = SpreadsheetApp.openById(ID_PLANILLA_MADRE);
  const hojaAlu = ssMadre.getSheetByName("DB_ALUMNOS");
  hojaPrecios = SS_LOCAL.insertSheet("CONFIG_PRECIOS");
  const meses = ["CURSO", "MATRICULA", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  hojaPrecios.appendRow(meses);
  hojaPrecios.getRange(1, 1, 1, meses.length).setBackground("#1a1a1a").setFontColor("white").setFontWeight("bold");
  if (hojaAlu) {
    const alumnos = hojaAlu.getDataRange().getValues();
    const cursos = [...new Set(alumnos.slice(1).map(a => a[3]))].filter(c => c).sort();
    cursos.forEach(c => {
      let fila = [c];
      for(let i=1; i<meses.length; i++) fila.push(0);
      hojaPrecios.appendRow(fila);
    });
  }
  return hojaPrecios;
}
/**
 * Función extra para que el HTML pueda consultar 
 * el estado de acceso al iniciar.
 */
function validarUsuario() {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  if (!userEmail) userEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();

  var sheetWhitelist = SS_LOCAL.getSheetByName('Whitelist');
  if (!sheetWhitelist) return false;

  var dataWhitelist = sheetWhitelist.getRange(1, 1, sheetWhitelist.getLastRow(), 1).getValues();
  var listaEmails = dataWhitelist.map(function(fila) {
    return fila[0].toString().toLowerCase().trim();
  });

  return listaEmails.indexOf(userEmail) !== -1;
}

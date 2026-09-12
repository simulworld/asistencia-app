// ════════════════════════════════════════════════════════════════
// SISTEMA DE ASISTENCIA — Backend v5.0
// Soporta múltiples "edificios" en las mismas coordenadas
// (mismo punto físico, distintos clientes/pisos)
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = "1w3g2SCyI6AWsGIyY5wL2HXnRimnWht8ohZHugIEEBsQ";

const CORREOS_REPORTE = [
  "marcosledo97@gmail.com",
  // "segundo@correo.com",
  // "tercero@correo.com"
].join(",");

const TURNOS = {
  "Matutino":   { entrada: { h: 7,  m: 0 }, salida: { h: 15, m: 0 }, tolerancia_min: 10 },
  "Mixto":      { entrada: { h: 7,  m: 0 }, salida: { h: 17, m: 0 }, tolerancia_min: 10 },
  "Vespertino": { entrada: { h: 13, m: 0 }, salida: { h: 21, m: 0 }, tolerancia_min: 10 }
};

// ════════════════════════════════════════════════════════════════
// doGet — maneja catálogos y registros
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  var accion = e.parameter.accion;
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (accion === "empleados") return doGet_Empleados(ss);
  if (accion === "edificios") return doGet_Edificios(ss);
  if (accion === "registro")  return doGet_Registro(ss, e.parameter);

  return respuestaJSON({ status: "ERROR", mensaje: "Acción no reconocida." });
}

function doGet_Empleados(ss) {
  var datos     = ss.getSheetByName("Empleados").getDataRange().getValues();
  var empleados = [];
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][2] === true) {
      empleados.push({ id: String(datos[i][0]), nombre: String(datos[i][1]) });
    }
  }
  return respuestaJSON(empleados);
}

function doGet_Edificios(ss) {
  var datos     = ss.getSheetByName("Edificios").getDataRange().getValues();
  var edificios = [];
  for (var i = 1; i < datos.length; i++) {
    edificios.push({
      id:     String(datos[i][0]),
      nombre: String(datos[i][1]),
      lat:    parseFloat(datos[i][2]),
      lon:    parseFloat(datos[i][3]),
      radio:  parseFloat(datos[i][4]) || 100
    });
  }
  return respuestaJSON(edificios);
}

// ════════════════════════════════════════════════════════════════
// REGISTRO
// Ahora el empleado (vía la app) indica EXACTAMENTE qué edificio
// eligió — el servidor solo valida que esa elección sea geográfi-
// camente válida. Esto resuelve el caso de dos "edificios" (dos
// clientes/pisos) en el mismo punto GPS, donde ya no es posible
// adivinar cuál quiso el empleado solo con lat/lon.
// ════════════════════════════════════════════════════════════════
function doGet_Registro(ss, p) {
  // ── Candado de concurrencia ─────────────────────────────────
  // Evita que dos registros simultáneos (ej. hora pico de entrada)
  // lean y escriban Sesiones_Activas al mismo tiempo y generen
  // sesiones duplicadas o datos corruptos. Solo UNA ejecución a la
  // vez puede estar dentro de este bloque; las demás esperan en fila.
  var lock = LockService.getScriptLock();
  var candadoObtenido = false;

  try {
    candadoObtenido = lock.tryLock(10000); // espera hasta 10s por el candado
    if (!candadoObtenido) {
      return respuestaJSON({
        status:  "ERROR",
        mensaje: "El sistema está procesando otros registros en este momento. Intenta de nuevo en unos segundos."
      });
    }

    // 1. Validar empleado activo
    var empleado = buscarEmpleado(ss, p.id_empleado);
    if (!empleado) {
      return respuestaJSON({ status: "ERROR", mensaje: "Empleado no encontrado o inactivo." });
    }

    // 2. Validar turno y evento ANTES de tocar cualquier hoja —
    //    si vienen vacíos, mal escritos o corruptos, se rechazan
    //    aquí sin dejar rastros a medias en Log ni en sesiones.
    var turno  = String(p.turno  || "").trim();
    var evento = String(p.evento || "").trim();

    if (Object.keys(TURNOS).indexOf(turno) === -1) {
      return respuestaJSON({ status: "ERROR", mensaje: "Turno no válido: '" + turno + "'." });
    }
    var EVENTOS_VALIDOS = ["Entrada", "Inicio Comida", "Fin Comida", "Salida"];
    if (EVENTOS_VALIDOS.indexOf(evento) === -1) {
      return respuestaJSON({ status: "ERROR", mensaje: "Evento no válido: '" + evento + "'." });
    }

    // 3. Validar que el edificio elegido sea geográficamente válido
    var latitud  = parseFloat(p.latitud);
    var longitud = parseFloat(p.longitud);
    var edificio = validarEdificioSeleccionado(ss, p.id_edificio, latitud, longitud);
    if (!edificio.valido) {
      return respuestaJSON({ status: "ERROR", mensaje: edificio.mensaje });
    }

    var ahora = new Date();
    var fecha = Utilities.formatDate(ahora, "America/Mexico_City", "yyyyMMdd");

    // 4. Resolver a qué sesión pertenece este evento
    var sesionInfo = resolverSesion(ss, empleado.id, turno, edificio, fecha, evento);

    // 5. Guardar SIEMPRE en el Log crudo, con id_sesion incluido
    var estatusLog = (evento === "Entrada" || evento === "Salida")
      ? calcularEstatus(turno, evento, ahora)
      : "—";

    ss.getSheetByName("Log").appendRow([
      sesionInfo.idSesion,
      ahora, p.id_empleado, empleado.nombre,
      edificio.id_edificio, edificio.nombre,
      turno, evento, estatusLog, latitud, longitud
    ]);

    if (sesionInfo.error) {
      return respuestaJSON({ status: "ERROR", mensaje: sesionInfo.error });
    }

    if (evento === "Entrada") {
      return manejarEntrada(ss, empleado, edificio, turno, fecha, ahora, sesionInfo.idSesion);
    }
    if (evento === "Inicio Comida" || evento === "Fin Comida") {
      return manejarComida(ss, sesionInfo.sesion, evento, ahora);
    }
    if (evento === "Salida") {
      return manejarSalida(ss, empleado, edificio, turno, fecha, ahora, sesionInfo.sesion);
    }

    return respuestaJSON({ status: "ERROR", mensaje: "Evento no reconocido." });

  } catch(err) {
    return respuestaJSON({ status: "ERROR", mensaje: "Error interno: " + err.toString() });

  } finally {
    // El candado SIEMPRE se libera, incluso si hubo error o return
    // anticipado — de lo contrario, todas las peticiones siguientes
    // se quedarían esperando un candado que nadie va a soltar.
    if (candadoObtenido) lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// Valida que el id_edificio que mandó la app sea real Y que las
// coordenadas GPS recibidas caigan dentro del radio de ESE edificio
// en particular (no de cualquier edificio cercano).
// ════════════════════════════════════════════════════════════════
function validarEdificioSeleccionado(ss, idEdificio, lat, lon) {
  if (!idEdificio) {
    return { valido: false, mensaje: "No se especificó un edificio. Selecciona tu ubicación en la app." };
  }

  var datos = ss.getSheetByName("Edificios").getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(idEdificio)) {
      var latEd  = parseFloat(datos[i][2]);
      var lonEd  = parseFloat(datos[i][3]);
      var radio  = parseFloat(datos[i][4]) || 100;
      var dist   = haversine(lat, lon, latEd, lonEd);

      if (dist <= radio) {
        return {
          valido:      true,
          id_edificio: String(datos[i][0]),
          nombre:      String(datos[i][1])
        };
      }
      return {
        valido:  false,
        mensaje: "Estás fuera del área de ese punto de registro (a " + Math.round(dist) + "m). Verifica tu ubicación."
      };
    }
  }
  return { valido: false, mensaje: "El edificio seleccionado no existe en el catálogo." };
}

// ════════════════════════════════════════════════════════════════
// Resuelve el id_sesion correcto para cualquier evento entrante
// ════════════════════════════════════════════════════════════════
function resolverSesion(ss, idEmpleado, turno, edificio, fecha, evento) {
  var abierta = buscarSesionAbierta(ss, idEmpleado, turno, edificio.id_edificio, fecha);

  if (evento === "Entrada") {
    if (abierta) {
      return {
        idSesion: abierta.datos[0],
        sesion:   abierta,
        error:    "Ya tienes una entrada abierta para este turno y edificio hoy. Marca 'Salida' antes de volver a entrar."
      };
    }
    var sufijo   = contarSesionesCerradasHoy(ss, idEmpleado, turno, edificio.nombre, fecha);
    var idSesion = idEmpleado + "_" + fecha + "_" + turno + "_" + edificio.id_edificio +
                   (sufijo > 0 ? "_" + (sufijo + 1) : "");
    return { idSesion: idSesion, sesion: null, error: null };
  }

  if (!abierta) {
    return {
      idSesion: "",
      sesion:   null,
      error:    "No tienes una entrada activa en este turno/edificio. Registra tu Entrada primero."
    };
  }
  return { idSesion: abierta.datos[0], sesion: abierta, error: null };
}

// ── ENTRADA ─────────────────────────────────────────────────
function manejarEntrada(ss, empleado, edificio, turno, fecha, ahora, idSesion) {
  ss.getSheetByName("Sesiones_Activas").appendRow([
    idSesion, empleado.id, empleado.nombre, turno,
    edificio.id_edificio, edificio.nombre, fecha,
    ahora, "", ""
  ]);
  var estatus = calcularEstatus(turno, "Entrada", ahora);
  return respuestaJSON({ status: "OK", mensaje: "Entrada registrada. " + estatus });
}

// ── INICIO/FIN COMIDA ───────────────────────────────────────
function manejarComida(ss, sesion, evento, ahora) {
  var hoja = ss.getSheetByName("Sesiones_Activas");
  var col  = (evento === "Inicio Comida") ? 9 : 10;
  hoja.getRange(sesion.fila, col).setValue(ahora);
  return respuestaJSON({ status: "OK", mensaje: evento + " registrado." });
}

// ── SALIDA ──────────────────────────────────────────────────
function manejarSalida(ss, empleado, edificio, turno, fecha, ahora, sesion) {
  var horaEntrada = new Date(sesion.datos[7]);
  var horaIniCom  = sesion.datos[8] ? new Date(sesion.datos[8]) : null;
  var horaFinCom  = sesion.datos[9] ? new Date(sesion.datos[9]) : null;

  var msTotales = ahora - horaEntrada;
  var msComida  = (horaIniCom && horaFinCom) ? (horaFinCom - horaIniCom) : 0;
  var msNetos   = msTotales - msComida;

  var horasTrabajadas = Math.round((msNetos / 3600000) * 100) / 100;
  var horasComida     = Math.round((msComida / 3600000) * 100) / 100;

  var estatusEntrada = calcularEstatus(turno, "Entrada", horaEntrada);
  var estatusSalida  = calcularEstatus(turno, "Salida", ahora);

  ss.getSheetByName("Resumen_Turnos").appendRow([
    sesion.datos[0], empleado.id, empleado.nombre, fecha,
    edificio.nombre, turno, horaEntrada, ahora,
    horasComida, horasTrabajadas, estatusEntrada, estatusSalida
  ]);

  ss.getSheetByName("Sesiones_Activas").deleteRow(sesion.fila);

  return respuestaJSON({
    status: "OK",
    mensaje: "Salida registrada. " + estatusSalida + " — " + horasTrabajadas + " hrs trabajadas."
  });
}

// ════════════════════════════════════════════════════════════════
// Helpers de sesión
// ════════════════════════════════════════════════════════════════
function buscarSesionAbierta(ss, idEmpleado, turno, idEdificio, fecha) {
  var hoja  = ss.getSheetByName("Sesiones_Activas");
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]) === String(idEmpleado) &&
        datos[i][3] === turno &&
        String(datos[i][4]) === String(idEdificio) &&
        String(datos[i][6]) === fecha) {
      return { fila: i + 1, datos: datos[i] };
    }
  }
  return null;
}

function contarSesionesCerradasHoy(ss, idEmpleado, turno, nombreEdificio, fecha) {
  var hoja  = ss.getSheetByName("Resumen_Turnos");
  var datos = hoja.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]) === String(idEmpleado) &&
        datos[i][5] === turno &&
        datos[i][4] === nombreEdificio &&
        String(datos[i][3]) === fecha) {
      count++;
    }
  }
  return count;
}

// ════════════════════════════════════════════════════════════════
// Helpers generales
// ════════════════════════════════════════════════════════════════
function buscarEmpleado(ss, id) {
  var datos = ss.getSheetByName("Empleados").getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id) && datos[i][2] === true) {
      return { id: datos[i][0], nombre: datos[i][1] };
    }
  }
  return null;
}

function haversine(lat1, lon1, lat2, lon2) {
  var R  = 6371000;
  var p1 = lat1 * Math.PI / 180, p2 = lat2 * Math.PI / 180;
  var dp = (lat2 - lat1) * Math.PI / 180;
  var dl = (lon2 - lon1) * Math.PI / 180;
  var a  = Math.sin(dp/2) * Math.sin(dp/2)
           + Math.cos(p1) * Math.cos(p2) * Math.sin(dl/2) * Math.sin(dl/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function calcularEstatus(turno, evento, momento) {
  var regla = TURNOS[turno];
  if (!regla) return "Turno desconocido";

  var limiteEntrada = new Date(momento);
  limiteEntrada.setHours(regla.entrada.h, regla.entrada.m + regla.tolerancia_min, 0, 0);

  var limiteSalida = new Date(momento);
  limiteSalida.setHours(regla.salida.h, regla.salida.m, 0, 0);

  if (evento === "Entrada") {
    if (momento <= limiteEntrada) return "A tiempo";
    var minutos = Math.round((momento - limiteEntrada) / 60000);
    return "Retardo " + minutos + " min";
  }
  if (evento === "Salida") {
    if (momento >= limiteSalida) return "A tiempo";
    var minAntes = Math.round((limiteSalida - momento) / 60000);
    return "Salida anticipada " + minAntes + " min";
  }
  return "—";
}

function respuestaJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// ALERTA DE SESIONES ABIERTAS SIN CERRAR
// ════════════════════════════════════════════════════════════════
function revisarSesionesAbandonadas() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var datos = ss.getSheetByName("Sesiones_Activas").getDataRange().getValues();
  var ahora = new Date();
  var abandonadas = [];

  for (var i = 1; i < datos.length; i++) {
    var horaEntrada  = new Date(datos[i][7]);
    var horasAbierta = (ahora - horaEntrada) / 3600000;
    if (horasAbierta > 16) {
      abandonadas.push(datos[i][2] + " — " + datos[i][5] + " — " + datos[i][3] +
                        " (abierta desde " + Utilities.formatDate(horaEntrada, "America/Mexico_City", "dd/MM HH:mm") + ")");
    }
  }

  if (abandonadas.length > 0) {
    GmailApp.sendEmail(
      CORREOS_REPORTE,
      "⚠ Sesiones sin cerrar — Sistema de Asistencia",
      "Los siguientes registros llevan más de 16 horas sin marcar Salida:\n\n" +
      abandonadas.join("\n") +
      "\n\nRevisa con el empleado y corrige manualmente si es necesario."
    );
  }
}

// ════════════════════════════════════════════════════════════════
// REPORTE SEMANAL
// ════════════════════════════════════════════════════════════════
function enviarReporteSemanal() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var datos   = ss.getSheetByName("Resumen_Turnos").getDataRange().getValues();
  var hoy     = new Date();
  var hace7d  = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);

  var filas = datos.slice(1).filter(function(f) {
    var fechaSesion = new Date(f[6]);
    return fechaSesion >= hace7d && fechaSesion <= hoy;
  });

  var totalEmpleados = new Set(filas.map(function(f) { return f[1]; })).size;
  var totalHoras      = filas.reduce(function(sum, f) { return sum + (parseFloat(f[9]) || 0); }, 0);
  var retardos         = filas.filter(function(f) { return String(f[10]).startsWith("Retardo"); }).length;
  var anticipadas      = filas.filter(function(f) { return String(f[11]).startsWith("Salida anticipada"); }).length;

  var resumen =
    "Reporte de asistencia — últimos 7 días\n" +
    "─────────────────────────────────────\n" +
    "Turnos completados:      " + filas.length         + "\n" +
    "Empleados únicos:        " + totalEmpleados       + "\n" +
    "Horas totales trabajadas:" + totalHoras.toFixed(1) + "\n" +
    "Retardos detectados:     " + retardos             + "\n" +
    "Salidas anticipadas:     " + anticipadas          + "\n\n" +
    "Se adjunta el detalle completo en CSV.\n";

  var enc = ["ID Sesión","ID Empleado","Nombre","Fecha","Edificio","Turno",
             "Hora Entrada","Hora Salida","Horas Comida","Horas Trabajadas",
             "Estatus Entrada","Estatus Salida"];
  var csv = enc.join(",") + "\n";
  filas.forEach(function(f) {
    csv += f.map(function(c) {
      return '"' + String(c).replace(/"/g, '""') + '"';
    }).join(",") + "\n";
  });

  var fechaStr = Utilities.formatDate(hoy, "America/Mexico_City", "dd-MMM-yyyy");
  var blob     = Utilities.newBlob(csv, "text/csv", "asistencia_" + fechaStr + ".csv");

  if (filas.length === 0) {
    GmailApp.sendEmail(CORREOS_REPORTE,
      "Reporte Semanal Asistencia — sin registros (" + fechaStr + ")",
      "No hubo turnos completados en los últimos 7 días.");
    return;
  }

  GmailApp.sendEmail(CORREOS_REPORTE,
    "Reporte Semanal Asistencia — " + fechaStr,
    resumen,
    { attachments: [blob] }
  );
}

// ════════════════════════════════════════════════════════════════
// FUNCIONES DE PRUEBA
// ════════════════════════════════════════════════════════════════
function probarFlujoCompleto() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Logger.log("── 1. ENTRADA ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", evento: "Entrada",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Utilities.sleep(1000);
  Logger.log("── 2. INICIO COMIDA ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", evento: "Inicio Comida",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Utilities.sleep(1000);
  Logger.log("── 3. FIN COMIDA ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", evento: "Fin Comida",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Utilities.sleep(1000);
  Logger.log("── 4. SALIDA ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", evento: "Salida",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Logger.log("Revisa Resumen_Turnos y Log — ambos deben tener el mismo id_sesion.");
}

// Prueba específica para el caso de dos edificios en las mismas coordenadas
function probarDosEdificiosMismoPunto() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Ajusta ED01 y ED02 a los IDs reales de tus dos registros en el mismo punto
  Logger.log("── Registro eligiendo ED01 ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", evento: "Entrada",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Logger.log("── Registro eligiendo ED02 (mismas coordenadas, cliente distinto) ──");
  Logger.log(doGet_Registro(ss, {
    id_empleado: "E002", id_edificio: "ED02", turno: "Matutino", evento: "Entrada",
    latitud: "19.44337740", longitud: "-99.18361210"
  }).getContent());

  Logger.log("Ambos deben responder OK y quedar en Sesiones_Activas con su propio id_edificio.");
}

function diagnosticarEmpleado() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var datos = ss.getSheetByName("Empleados").getDataRange().getValues();
  Logger.log("Total de filas: " + datos.length);
  for (var i = 1; i < datos.length; i++) {
    Logger.log("Fila " + (i+1) + ": ID='" + datos[i][0] + "' Nombre='" + datos[i][1] + "' Activo=" + datos[i][2]);
  }
}

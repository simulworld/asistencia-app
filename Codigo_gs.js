// ════════════════════════════════════════════════════════════════
// SISTEMA DE ASISTENCIA — Backend v9.0
// - Múltiples "edificios" en las mismas coordenadas (mismo punto
//   físico, distintos clientes/pisos)
// - Candado de concurrencia (LockService) contra registros simultáneos
// - Validación temprana de turno/evento
// - Turnos: Matutino, Mixto, Vespertino
// - Turno "Comodín" solo para empleados con base = COMODIN
// - Resumen_Turnos con horas de comida junto a hora_entrada/salida
// - Comida incompleta: descuento automático de 2 horas
// - Sesiones abiertas de días anteriores: cierre como jornada abandonada
// - Soporte offline con hora de evento + hora de sincronización
// - id_evento para evitar duplicados por reintentos
// - Validación de precisión GPS en servidor
// - radio_metros tomado de cada fila de Edificios
// - Respaldo automático diario a Drive con rotación de 30 días
// - Alertas de error por correo + bitácora en hoja Errores_Sistema
// - Verificación de estructura que AGREGA columnas faltantes sin borrar/reordenar
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
  "Vespertino": { entrada: { h: 13, m: 0 }, salida: { h: 21, m: 0 }, tolerancia_min: 10 },
  // Sin horario fijo: para empleados con base "COMODIN" que atienden
  // más de un servicio/edificio al día. Nunca genera retardos ni
  // salidas anticipadas — solo registra la hora, sin evaluarla.
  "Comodín":    { sinHorario: true }
};

// Nombre EXACTO (mayúsculas) que debe tener la columna "base" en la
// hoja Empleados para que un empleado pueda usar el turno Comodín.
const VALOR_BASE_COMODIN = "COMODIN";

// Parámetros operativos de robustez
// radio_metros NO es un parámetro global: se toma de cada fila de Edificios.
const LIMITE_PRECISION_GPS_METROS = 150;
const MAX_HORAS_OFFLINE = 48;
const MAX_ADELANTO_RELOJ_OFFLINE_MIN = 5;
const DESCUENTO_COMIDA_INCOMPLETA_HORAS = 2;
const NOMBRE_HOJA_ABANDONADAS = "Sesiones_Abandonadas";
const VERSION_BACKEND = "9.0";

// ════════════════════════════════════════════════════════════════
// doGet — maneja catálogos y registros
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  var accion = String((e && e.parameter && e.parameter.accion) || "").trim();
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (accion === "empleados") return doGet_Empleados(ss);
  if (accion === "edificios") return doGet_Edificios(ss);
  if (accion === "registro")  return doGet_Registro(ss, e.parameter || {});

  return respuestaJSON({ status: "ERROR", mensaje: "Acción no reconocida." });
}

function doGet_Empleados(ss) {
  var hoja    = ss.getSheetByName("Empleados");
  if (!hoja) return respuestaJSON([]);
  var datos   = hoja.getDataRange().getValues();
  var headers = datos[0] || [];
  var idxId   = idxColumnaPorNombre(headers, "id_empleado");
  var idxNom  = idxColumnaPorNombre(headers, "nombre_completo");
  var idxAct  = idxColumnaPorNombre(headers, "activo");
  var idxBase = idxColumnaPorNombre(headers, "base");

  var empleados = [];
  for (var i = 1; i < datos.length; i++) {
    var activo = idxAct !== -1 ? normalizarBooleano(datos[i][idxAct]) : false;
    if (!activo) continue;
    var valorBase = idxBase !== -1 ? String(datos[i][idxBase] || "").trim().toUpperCase() : "";
    empleados.push({
      id:      String(idxId !== -1 ? datos[i][idxId] : datos[i][0]),
      nombre:  String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1]),
      comodin: valorBase === VALOR_BASE_COMODIN
    });
  }
  return respuestaJSON(empleados);
}

function idxColumnaPorNombre(headers, nombreBuscado) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === nombreBuscado.toLowerCase()) {
      return i;
    }
  }
  return -1;
}

function normalizarBooleano(valor) {
  if (valor === true) return true;
  return String(valor).trim().toUpperCase() === "TRUE";
}

function doGet_Edificios(ss) {
  var hoja = ss.getSheetByName("Edificios");
  if (!hoja) return respuestaJSON([]);
  var datos = hoja.getDataRange().getValues();
  var headers = datos[0] || [];
  var idxId  = idxColumnaPorNombre(headers, "id_edificio");
  var idxNom = idxColumnaPorNombre(headers, "nombre");
  var idxLat = idxColumnaPorNombre(headers, "latitud");
  var idxLon = idxColumnaPorNombre(headers, "longitud");
  var idxRad = idxColumnaPorNombre(headers, "radio_metros");
  var idxAct = idxColumnaPorNombre(headers, "activo");

  var edificios = [];
  for (var i = 1; i < datos.length; i++) {
    var activo = (idxAct === -1) ? true : normalizarBooleano(datos[i][idxAct]);
    if (!activo) continue;

    var radioValor = parseFloat(idxRad !== -1 ? datos[i][idxRad] : datos[i][4]);
    if (!isFinite(radioValor) || radioValor <= 0) continue;

    edificios.push({
      id:     String(idxId !== -1 ? datos[i][idxId] : datos[i][0]),
      nombre: String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1]),
      lat:    parseFloat(idxLat !== -1 ? datos[i][idxLat] : datos[i][2]),
      lon:    parseFloat(idxLon !== -1 ? datos[i][idxLon] : datos[i][3]),
      radio:  parseFloat(idxRad !== -1 ? datos[i][idxRad] : datos[i][4])
    });
  }
  return respuestaJSON(edificios);
}

// ════════════════════════════════════════════════════════════════
// REGISTRO
// ════════════════════════════════════════════════════════════════
function doGet_Registro(ss, p) {
  var lock = LockService.getScriptLock();
  var candadoObtenido = false;

  try {
    candadoObtenido = lock.tryLock(10000);
    if (!candadoObtenido) {
      return respuestaJSON({
        status: "ERROR",
        codigo: "LOCK_TIMEOUT",
        reintentar: true,
        mensaje: "El sistema está procesando otros registros. Intenta de nuevo en unos segundos."
      });
    }

    var ahoraServidor = new Date();
    var empleado = buscarEmpleado(ss, p.id_empleado);
    if (!empleado) {
      return respuestaJSON({ status: "ERROR", codigo: "EMPLEADO_INVALIDO", reintentar: false, mensaje: "Empleado no encontrado o inactivo." });
    }

    var turno  = String(p.turno || "").trim();
    var evento = String(p.evento || "").trim();
    var idEvento = String(p.id_evento || "").trim();
    var origenOriginal = String(p.origen || "ONLINE").trim().toUpperCase();

    if (Object.keys(TURNOS).indexOf(turno) === -1) {
      return respuestaJSON({ status: "ERROR", codigo: "TURNO_INVALIDO", reintentar: false, mensaje: "Turno no válido." });
    }

    var EVENTOS_VALIDOS = ["Entrada", "Inicio Comida", "Fin Comida", "Salida"];
    if (EVENTOS_VALIDOS.indexOf(evento) === -1) {
      return respuestaJSON({ status: "ERROR", codigo: "EVENTO_INVALIDO", reintentar: false, mensaje: "Evento no válido." });
    }

    if (turno === "Comodín" && !empleado.comodin) {
      return respuestaJSON({ status: "ERROR", codigo: "COMODIN_NO_AUTORIZADO", reintentar: false, mensaje: "El turno Comodín no está disponible para tu perfil." });
    }

    // Antes de resolver la sesión, cerramos automáticamente sesiones de días anteriores.
    cerrarSesionesAbandonadasPorCambioDeDia(ss, Utilities.formatDate(ahoraServidor, "America/Mexico_City", "yyyyMMdd"), ahoraServidor);

    if (idEvento && eventoYaProcesado(ss, idEvento)) {
      return respuestaJSON({
        status: "OK",
        codigo: "DUPLICADO",
        reintentar: false,
        mensaje: "El registro ya había sido procesado anteriormente."
      });
    }

    var momentoInfo = resolverMomentoEvento(p, ahoraServidor);
    if (!momentoInfo.valido) {
      registrarLogRegistro(ss, {
        idSesion: "",
        momentoEvento: momentoInfo.momento || ahoraServidor,
        momentoServidor: ahoraServidor,
        idEmpleado: empleado.id,
        nombre: empleado.nombre,
        edificioId: String(p.id_edificio || ""),
        edificioNombre: "",
        turno: turno,
        evento: evento,
        estatus: "—",
        latitud: p.latitud,
        longitud: p.longitud,
        precision: p.precision,
        origen: origenOriginal,
        idEvento: idEvento,
        resultado: "REVISAR_OFFLINE"
      });
      return respuestaJSON({ status: "ERROR", codigo: "OFFLINE_RELOJ_ANOMALO", reintentar: false, mensaje: momentoInfo.mensaje });
    }

    var latitud  = parseFloat(p.latitud);
    var longitud = parseFloat(p.longitud);
    var precision = parseFloat(p.precision);
    if (!isFinite(latitud) || !isFinite(longitud)) {
      return respuestaJSON({ status: "ERROR", codigo: "GPS_INVALIDO", reintentar: false, mensaje: "No se pudo validar tu ubicación. Activa el GPS y vuelve a intentarlo." });
    }
    if (!isFinite(precision) || precision <= 0 || precision > LIMITE_PRECISION_GPS_METROS) {
      return respuestaJSON({ status: "ERROR", codigo: "GPS_PRECISION", reintentar: false, mensaje: "La precisión GPS no es suficiente para registrar. Espera unos segundos o sal a un área con mejor señal." });
    }

    var edificio = validarEdificioSeleccionado(ss, p.id_edificio, latitud, longitud);
    if (!edificio.valido) {
      registrarLogRegistro(ss, {
        idSesion: "",
        momentoEvento: momentoInfo.momento,
        momentoServidor: ahoraServidor,
        idEmpleado: empleado.id,
        nombre: empleado.nombre,
        edificioId: String(p.id_edificio || ""),
        edificioNombre: edificio.nombre || "",
        turno: turno,
        evento: evento,
        estatus: "—",
        latitud: latitud,
        longitud: longitud,
        precision: precision,
        origen: origenOriginal,
        idEvento: idEvento,
        resultado: "RECHAZADO_GPS"
      });
      return respuestaJSON({ status: "ERROR", codigo: "GPS_FUERA", reintentar: false, mensaje: edificio.mensaje });
    }

    var fechaEvento = Utilities.formatDate(momentoInfo.momento, "America/Mexico_City", "yyyyMMdd");
    var sesionInfo = resolverSesion(ss, empleado.id, turno, edificio, fechaEvento, evento, momentoInfo.esOffline);

    var estatusLog = (evento === "Entrada" || evento === "Salida")
      ? calcularEstatus(turno, evento, momentoInfo.momento)
      : "—";

    if (sesionInfo.error) {
      registrarLogRegistro(ss, {
        idSesion: "",
        momentoEvento: momentoInfo.momento,
        momentoServidor: ahoraServidor,
        idEmpleado: empleado.id,
        nombre: empleado.nombre,
        edificioId: edificio.id_edificio,
        edificioNombre: edificio.nombre,
        turno: turno,
        evento: evento,
        estatus: estatusLog,
        latitud: latitud,
        longitud: longitud,
        precision: precision,
        origen: origenOriginal,
        idEvento: idEvento,
        resultado: "RECHAZADO_SESION"
      });
      return respuestaJSON({ status: "ERROR", codigo: "SESION_INVALIDA", reintentar: false, mensaje: sesionInfo.error });
    }

    var resultado = "OK";
    var respuesta;

    if (evento === "Entrada") {
      respuesta = manejarEntrada(ss, empleado, edificio, turno, fechaEvento, momentoInfo.momento, sesionInfo.idSesion);
    } else if (evento === "Inicio Comida" || evento === "Fin Comida") {
      respuesta = manejarComida(ss, sesionInfo.sesion, evento, momentoInfo.momento);
    } else if (evento === "Salida") {
      respuesta = manejarSalida(ss, empleado, edificio, turno, fechaEvento, momentoInfo.momento, sesionInfo.sesion, momentoInfo.esOffline);
    }

    var respuestaData = JSON.parse(respuesta.getContent());
    resultado = respuestaData.status === "OK" ? (respuestaData.codigo || "OK") : (respuestaData.codigo || "RECHAZADO");

    registrarLogRegistro(ss, {
      idSesion: sesionInfo.idSesion,
      momentoEvento: momentoInfo.momento,
      momentoServidor: ahoraServidor,
      idEmpleado: empleado.id,
      nombre: empleado.nombre,
      edificioId: edificio.id_edificio,
      edificioNombre: edificio.nombre,
      turno: turno,
      evento: evento,
      estatus: estatusLog,
      latitud: latitud,
      longitud: longitud,
      precision: precision,
      origen: origenOriginal,
      idEvento: idEvento,
      resultado: resultado
    });

    return respuestaJSON(Object.assign(respuestaData, { id_evento: idEvento || null }));

  } catch (err) {
    try { registrarError("doGet_Registro", err); } catch (ignore) {}
    return respuestaJSON({ status: "ERROR", codigo: "ERROR_INTERNO", reintentar: true, mensaje: "No fue posible completar el registro. Intenta nuevamente." });
  } finally {
    if (candadoObtenido) lock.releaseLock();
  }
}

function resolverMomentoEvento(p, ahoraServidor) {
  var origen = String(p.origen || "ONLINE").trim().toUpperCase();
  var esOffline = origen !== "ONLINE";
  if (!esOffline) return { valido: true, momento: ahoraServidor, esOffline: false };

  var ts = String(p.ts_cliente || "").trim();
  if (!ts) return { valido: false, momento: ahoraServidor, esOffline: true, mensaje: "No se recibió la hora original del registro offline." };

  var momento = new Date(ts);
  if (isNaN(momento.getTime())) {
    return { valido: false, momento: ahoraServidor, esOffline: true, mensaje: "No se pudo interpretar la hora del registro offline." };
  }

  var diferenciaHoras = (ahoraServidor - momento) / 3600000;
  var adelantoMin = (momento - ahoraServidor) / 60000;
  if (adelantoMin > MAX_ADELANTO_RELOJ_OFFLINE_MIN || diferenciaHoras > MAX_HORAS_OFFLINE) {
    return {
      valido: false,
      momento: momento,
      esOffline: true,
      mensaje: "El registro sin conexión tiene una hora fuera del rango permitido. Se conservó como incidencia para revisión."
    };
  }

  return { valido: true, momento: momento, esOffline: true };
}

function registrarLogRegistro(ss, r) {
  var hoja = ss.getSheetByName("Log");
  if (!hoja) return;
  hoja.appendRow([
    r.idSesion || "",
    r.momentoEvento || "",
    r.idEmpleado || "",
    r.nombre || "",
    r.edificioId || "",
    r.edificioNombre || "",
    r.turno || "",
    r.evento || "",
    r.estatus || "—",
    r.latitud || "",
    r.longitud || "",
    r.precision || "",
    r.origen || "ONLINE",
    r.momentoServidor || "",
    r.idEvento || "",
    r.resultado || ""
  ]);
}

function eventoYaProcesado(ss, idEvento) {
  if (!idEvento) return false;
  var hoja = ss.getSheetByName("Log");
  if (!hoja || hoja.getLastRow() < 2) return false;
  var lastCol = hoja.getLastColumn();
  var headers = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = idxColumnaPorNombre(headers, "id_evento");
  if (idx === -1) return false;
  if (hoja.getLastRow() < 2) return false;
  return hoja.getRange(2, idx + 1, hoja.getLastRow() - 1, 1)
    .createTextFinder(String(idEvento))
    .matchEntireCell(true)
    .findNext() !== null;
}

// ════════════════════════════════════════════════════════════════
// Valida que el id_edificio que mandó la app sea real Y que las
// coordenadas GPS recibidas caigan dentro del radio de ESE edificio.
// ════════════════════════════════════════════════════════════════
function validarEdificioSeleccionado(ss, idEdificio, lat, lon) {
  if (!idEdificio) {
    return { valido: false, mensaje: "No se especificó un edificio. Selecciona tu ubicación en la app." };
  }

  var hoja = ss.getSheetByName("Edificios");
  if (!hoja) return { valido: false, mensaje: "No está disponible el catálogo de edificios." };
  var datos = hoja.getDataRange().getValues();
  var headers = datos[0] || [];
  var idxId = idxColumnaPorNombre(headers, "id_edificio");
  var idxNom = idxColumnaPorNombre(headers, "nombre");
  var idxLat = idxColumnaPorNombre(headers, "latitud");
  var idxLon = idxColumnaPorNombre(headers, "longitud");
  var idxRad = idxColumnaPorNombre(headers, "radio_metros");
  var idxAct = idxColumnaPorNombre(headers, "activo");

  for (var i = 1; i < datos.length; i++) {
    var id = String(idxId !== -1 ? datos[i][idxId] : datos[i][0]);
    if (id !== String(idEdificio)) continue;
    if (idxAct !== -1 && !normalizarBooleano(datos[i][idxAct])) {
      return { valido: false, mensaje: "Ese punto de registro está inactivo." };
    }

    var latEd = parseFloat(idxLat !== -1 ? datos[i][idxLat] : datos[i][2]);
    var lonEd = parseFloat(idxLon !== -1 ? datos[i][idxLon] : datos[i][3]);
    var radio = parseFloat(idxRad !== -1 ? datos[i][idxRad] : datos[i][4]);
    if (!isFinite(radio) || radio <= 0) {
      return { valido: false, nombre: String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1]), mensaje: "El edificio seleccionado no tiene un radio_metros válido configurado." };
    }

    var dist = haversine(lat, lon, latEd, lonEd);

    if (dist <= radio) {
      return {
        valido: true,
        id_edificio: id,
        nombre: String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1])
      };
    }
    return {
      valido: false,
      nombre: String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1]),
      mensaje: "Estás fuera del área de ese punto de registro (a " + Math.round(dist) + "m). Verifica tu ubicación."
    };
  }
  return { valido: false, mensaje: "El edificio seleccionado no existe en el catálogo." };
}

// ════════════════════════════════════════════════════════════════
// Resuelve el id_sesion correcto para cualquier evento entrante
// ════════════════════════════════════════════════════════════════
function resolverSesion(ss, idEmpleado, turno, edificio, fecha, evento, esOffline) {
  var abierta = buscarSesionAbierta(ss, idEmpleado, turno, edificio.id_edificio, fecha);

  if (evento === "Entrada") {
    if (abierta) {
      return { idSesion: abierta.datos[0], sesion: abierta, error: "Ya tienes una entrada abierta para este turno y edificio hoy. Marca 'Salida' antes de volver a entrar." };
    }

    // Si una entrada offline de un día anterior llega después de medianoche,
    // recuperamos la sesión que habíamos marcado como abandonada automáticamente.
    if (esOffline) {
      var abandonadaEntrada = buscarSesionAbandonada(ss, idEmpleado, turno, edificio.id_edificio, fecha);
      if (abandonadaEntrada) {
        var reactivadaEntrada = reactivarSesionAbandonada(ss, abandonadaEntrada);
        return { idSesion: reactivadaEntrada.idSesion, sesion: reactivadaEntrada.sesion, error: null, recuperadaOffline: true };
      }
    }

    var sufijo = contarSesionesCerradasHoy(ss, idEmpleado, turno, edificio.nombre, fecha);
    var idSesion = idEmpleado + "_" + fecha + "_" + turno + "_" + edificio.id_edificio +
                   (sufijo > 0 ? "_" + (sufijo + 1) : "");
    return { idSesion: idSesion, sesion: null, error: null };
  }

  if (!abierta && esOffline) {
    var abandonada = buscarSesionAbandonada(ss, idEmpleado, turno, edificio.id_edificio, fecha);
    if (abandonada) {
      var reactivada = reactivarSesionAbandonada(ss, abandonada);
      abierta = reactivada.sesion;
    }
  }

  if (!abierta) {
    return { idSesion: "", sesion: null, error: "No tienes una entrada activa en este turno/edificio. Registra tu Entrada primero." };
  }

  if (evento === "Inicio Comida" && abierta.datos[8]) {
    return { idSesion: abierta.datos[0], sesion: abierta, error: "Ya tienes un Inicio Comida registrado para esta sesión." };
  }
  if (evento === "Inicio Comida" && abierta.datos[9]) {
    return { idSesion: abierta.datos[0], sesion: abierta, error: "La comida ya tiene un Fin Comida registrado; no puedes iniciar otra comida en esta sesión." };
  }
  if (evento === "Fin Comida" && !abierta.datos[8]) {
    return { idSesion: abierta.datos[0], sesion: abierta, error: "No puedes registrar Fin Comida sin haber registrado primero Inicio Comida." };
  }
  if (evento === "Fin Comida" && abierta.datos[9]) {
    return { idSesion: abierta.datos[0], sesion: abierta, error: "Ya tienes un Fin Comida registrado para esta sesión." };
  }

  return { idSesion: abierta.datos[0], sesion: abierta, error: null };
}

// ── ENTRADA ─────────────────────────────────────────────────
function manejarEntrada(ss, empleado, edificio, turno, fecha, momento, idSesion) {
  ss.getSheetByName("Sesiones_Activas").appendRow([
    idSesion, empleado.id, empleado.nombre, turno,
    edificio.id_edificio, edificio.nombre, fecha,
    momento, "", ""
  ]);
  var estatus = calcularEstatus(turno, "Entrada", momento);
  return respuestaJSON({ status: "OK", codigo: "OK", mensaje: "Entrada registrada. " + estatus });
}

// ── INICIO/FIN COMIDA ───────────────────────────────────────
function manejarComida(ss, sesion, evento, momento) {
  var hoja = ss.getSheetByName("Sesiones_Activas");
  var col  = (evento === "Inicio Comida") ? 9 : 10;
  hoja.getRange(sesion.fila, col).setValue(momento);
  return respuestaJSON({ status: "OK", codigo: "OK", mensaje: evento + " registrado." });
}

// ── SALIDA ──────────────────────────────────────────────────
function manejarSalida(ss, empleado, edificio, turno, fecha, momento, sesion, esOffline) {
  var horaEntrada = new Date(sesion.datos[7]);
  var horaIniCom  = sesion.datos[8] ? new Date(sesion.datos[8]) : null;
  var horaFinCom  = sesion.datos[9] ? new Date(sesion.datos[9]) : null;

  var msTotales = momento - horaEntrada;
  if (msTotales < 0) {
    return respuestaJSON({ status: "ERROR", codigo: "TIEMPO_INVALIDO", reintentar: false, mensaje: "La hora de salida no puede ser anterior a la entrada." });
  }

  var comidaCompleta = !!(horaIniCom && horaFinCom);
  var msComidaReal = comidaCompleta ? (horaFinCom - horaIniCom) : 0;
  if (comidaCompleta && (msComidaReal < 0 || horaIniCom < horaEntrada || horaFinCom > momento)) {
    return respuestaJSON({ status: "ERROR", codigo: "COMIDA_INVALIDA", reintentar: false, mensaje: "Los registros de comida no tienen un orden válido." });
  }

  var descuentoComidaHoras = comidaCompleta ? (msComidaReal / 3600000) : DESCUENTO_COMIDA_INCOMPLETA_HORAS;
  var msDescuento = comidaCompleta ? msComidaReal : DESCUENTO_COMIDA_INCOMPLETA_HORAS * 3600000;
  var msNetos = msTotales - msDescuento;
  if (msNetos < 0) msNetos = 0;

  var horasTrabajadas = Math.round((msNetos / 3600000) * 100) / 100;
  var horasComida     = comidaCompleta ? Math.round((msComidaReal / 3600000) * 100) / 100 : 0;

  var estatusEntrada = calcularEstatus(turno, "Entrada", horaEntrada);
  var estatusSalida  = calcularEstatus(turno, "Salida", momento);
  var estatusComida   = comidaCompleta ? "Completa" : "Incompleta — descuento automático 2 h";
  var observaciones  = comidaCompleta ? "" : "Faltó Inicio Comida o Fin Comida; se aplicó descuento automático de 2 h.";
  if (esOffline) observaciones += (observaciones ? " " : "") + " Salida recibida por sincronización offline.";

  ss.getSheetByName("Resumen_Turnos").appendRow([
    sesion.datos[0], empleado.id, empleado.nombre, fecha,
    edificio.nombre, turno, horaEntrada, momento,
    horaIniCom || "", horaFinCom || "",
    horasComida, horasTrabajadas, estatusEntrada, estatusSalida,
    estatusComida, Math.round(descuentoComidaHoras * 100) / 100, observaciones
  ]);

  ss.getSheetByName("Sesiones_Activas").deleteRow(sesion.fila);

  var extra = comidaCompleta ? "" : " Se aplicó un descuento automático de 2 horas por comida incompleta.";
  if (esOffline) extra += " Registro sincronizado desde modo sin conexión.";

  return respuestaJSON({
    status: "OK",
    codigo: "OK",
    mensaje: "Salida registrada. " + estatusSalida + " — " + horasTrabajadas + " hrs trabajadas." + extra
  });
}

// ════════════════════════════════════════════════════════════════
// Helpers de sesión
// ════════════════════════════════════════════════════════════════
function buscarSesionAbierta(ss, idEmpleado, turno, idEdificio, fecha) {
  var hoja  = ss.getSheetByName("Sesiones_Activas");
  if (!hoja || hoja.getLastRow() < 2) return null;
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]) === String(idEmpleado) &&
        String(datos[i][3]) === String(turno) &&
        String(datos[i][4]) === String(idEdificio) &&
        String(datos[i][6]) === String(fecha)) {
      return { fila: i + 1, datos: datos[i] };
    }
  }
  return null;
}

function buscarSesionAbandonada(ss, idEmpleado, turno, idEdificio, fecha) {
  var hoja = ss.getSheetByName(NOMBRE_HOJA_ABANDONADAS);
  if (!hoja || hoja.getLastRow() < 2) return null;
  var datos = hoja.getDataRange().getValues();
  for (var i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][1]) === String(idEmpleado) &&
        String(datos[i][3]) === String(turno) &&
        String(datos[i][4]) === String(idEdificio) &&
        String(datos[i][6]) === String(fecha) &&
        String(datos[i][13] || "") === "ABANDONADA") {
      return { fila: i + 1, datos: datos[i] };
    }
  }
  return null;
}

function reactivarSesionAbandonada(ss, abandonada) {
  var hojaAb = ss.getSheetByName(NOMBRE_HOJA_ABANDONADAS);
  var hojaAc = ss.getSheetByName("Sesiones_Activas");
  var d = abandonada.datos;

  var nueva = [
    d[0], d[1], d[2], d[3], d[4], d[5], d[6],
    d[7], d[8] || "", d[9] || ""
  ];
  hojaAc.appendRow(nueva);
  hojaAb.getRange(abandonada.fila, 14).setValue("RECUPERADA_OFFLINE");
  hojaAb.getRange(abandonada.fila, 15).setValue("Se recibió posteriormente un evento offline correspondiente al mismo día de la jornada.");

  var ultimaFila = hojaAc.getLastRow();
  return { fila: ultimaFila, datos: hojaAc.getRange(ultimaFila, 1, 1, hojaAc.getLastColumn()).getValues()[0], idSesion: String(d[0]) };
}

function cerrarSesionesAbandonadasPorCambioDeDia(ss, fechaActual, ahoraServidor) {
  var hoja = ss.getSheetByName("Sesiones_Activas");
  if (!hoja || hoja.getLastRow() < 2) return;

  var hojaAb = obtenerOCrearHojaAbandonadas(ss);
  var datos = hoja.getDataRange().getValues();
  var filasABorrar = [];

  for (var i = 1; i < datos.length; i++) {
    var fechaSesion = String(datos[i][6] || "");
    if (!/^\d{8}$/.test(fechaSesion) || fechaSesion >= fechaActual) continue;

    hojaAb.appendRow([
      datos[i][0], datos[i][1], datos[i][2], datos[i][3], datos[i][4], datos[i][5],
      datos[i][6], datos[i][7], datos[i][8] || "", datos[i][9] || "",
      ahoraServidor,
      "No se registró Salida antes del cambio de día.",
      0,
      "ABANDONADA",
      "La sesión fue cerrada automáticamente al iniciar un nuevo día."
    ]);
    filasABorrar.push(i + 1);
  }

  for (var j = filasABorrar.length - 1; j >= 0; j--) {
    hoja.deleteRow(filasABorrar[j]);
  }
}

function obtenerOCrearHojaAbandonadas(ss) {
  var hoja = ss.getSheetByName(NOMBRE_HOJA_ABANDONADAS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_ABANDONADAS);
  }

  asegurarEncabezadosHoja(hoja, [
    "id_sesion", "id_empleado", "nombre", "turno", "id_edificio", "nombre_edificio",
    "fecha_jornada", "hora_entrada", "hora_inicio_comida", "hora_fin_comida",
    "fecha_hora_abandono", "motivo", "horas_trabajadas", "estatus", "observaciones"
  ]);
  return hoja;
}

function contarSesionesCerradasHoy(ss, idEmpleado, turno, nombreEdificio, fecha) {
  var hoja  = ss.getSheetByName("Resumen_Turnos");
  if (!hoja || hoja.getLastRow() < 2) return 0;
  var datos = hoja.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]) === String(idEmpleado) &&
        String(datos[i][5]) === String(turno) &&
        String(datos[i][4]) === String(nombreEdificio) &&
        String(datos[i][3]) === String(fecha)) {
      count++;
    }
  }
  return count;
}

// ════════════════════════════════════════════════════════════════
// Helpers generales
// ════════════════════════════════════════════════════════════════
function buscarEmpleado(ss, id) {
  var hoja = ss.getSheetByName("Empleados");
  if (!hoja) return null;
  var datos = hoja.getDataRange().getValues();
  var headers = datos[0] || [];
  var idxId = idxColumnaPorNombre(headers, "id_empleado");
  var idxNom = idxColumnaPorNombre(headers, "nombre_completo");
  var idxAct = idxColumnaPorNombre(headers, "activo");
  var idxBase = idxColumnaPorNombre(headers, "base");

  for (var i = 1; i < datos.length; i++) {
    var idFila = String(idxId !== -1 ? datos[i][idxId] : datos[i][0]);
    var activo = idxAct !== -1 ? normalizarBooleano(datos[i][idxAct]) : false;
    if (idFila === String(id) && activo) {
      var valorBase = idxBase !== -1 ? String(datos[i][idxBase] || "").trim().toUpperCase() : "";
      return {
        id: idFila,
        nombre: String(idxNom !== -1 ? datos[i][idxNom] : datos[i][1]),
        comodin: valorBase === VALOR_BASE_COMODIN
      };
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

  // Turnos sin horario fijo (Comodín) — nunca se evalúa puntualidad,
  // solo se registra la hora tal cual.
  if (regla.sinHorario) return "—";

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
// CIERRE/ALERTA DE SESIONES ABIERTAS SIN CERRAR
// ════════════════════════════════════════════════════════════════
function revisarSesionesAbandonadas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var lock = LockService.getScriptLock();
  var ok = lock.tryLock(10000);
  if (!ok) throw new Error("No fue posible adquirir el candado para revisar sesiones abandonadas.");

  try {
    var ahora = new Date();
    var fechaActual = Utilities.formatDate(ahora, "America/Mexico_City", "yyyyMMdd");
    cerrarSesionesAbandonadasPorCambioDeDia(ss, fechaActual, ahora);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// REPORTE SEMANAL
// ════════════════════════════════════════════════════════════════
function enviarReporteSemanal() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var datos   = ss.getSheetByName("Resumen_Turnos").getDataRange().getValues();
  var ahora   = new Date();
  var hace7d  = new Date(ahora.getTime() - 7 * 24 * 60 * 60 * 1000);

  var filas = datos.slice(1).filter(function(f) {
    var fechaSesion = new Date(f[6]);
    return fechaSesion >= hace7d && fechaSesion <= ahora;
  });

  var totalEmpleados = new Set(filas.map(function(f) { return f[1]; })).size;
  var totalHoras      = filas.reduce(function(sum, f) { return sum + (parseFloat(f[11]) || 0); }, 0);
  var retardos         = filas.filter(function(f) { return String(f[12]).startsWith("Retardo"); }).length;
  var anticipadas      = filas.filter(function(f) { return String(f[13]).startsWith("Salida anticipada"); }).length;
  var comidasIncomp    = filas.filter(function(f) { return String(f[14]).startsWith("Incompleta"); }).length;
  var abandonadas      = obtenerConteoAbandonadasUltimos7Dias(ss, hace7d, ahora);

  var resumen =
    "Reporte de asistencia — últimos 7 días\n" +
    "─────────────────────────────────────\n" +
    "Turnos completados:       " + filas.length          + "\n" +
    "Empleados únicos:         " + totalEmpleados        + "\n" +
    "Horas totales trabajadas: " + totalHoras.toFixed(1) + "\n" +
    "Retardos detectados:      " + retardos              + "\n" +
    "Salidas anticipadas:      " + anticipadas          + "\n" +
    "Comidas incompletas:      " + comidasIncomp         + "\n" +
    "Jornadas abandonadas:     " + abandonadas           + "\n\n" +
    "Se adjunta el detalle completo en CSV.\n";

  var enc = ["ID Sesión","ID Empleado","Nombre","Fecha","Edificio","Turno",
             "Hora Entrada","Hora Salida","Hora Inicio Comida","Hora Fin Comida",
             "Horas Comida","Horas Trabajadas","Estatus Entrada","Estatus Salida",
             "Estatus Comida","Descuento Comida (h)","Observaciones"];
  var csv = enc.join(",") + "\n";
  filas.forEach(function(f) {
    csv += f.map(function(c) {
      return '"' + String(c).replace(/"/g, '""') + '"';
    }).join(",") + "\n";
  });

  var fechaStr = Utilities.formatDate(ahora, "America/Mexico_City", "dd-MMM-yyyy");
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

function obtenerConteoAbandonadasUltimos7Dias(ss, desde, hasta) {
  var hoja = ss.getSheetByName(NOMBRE_HOJA_ABANDONADAS);
  if (!hoja || hoja.getLastRow() < 2) return 0;
  var datos = hoja.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < datos.length; i++) {
    var fecha = datos[i][10] ? new Date(datos[i][10]) : null;
    if (fecha && fecha >= desde && fecha <= hasta && String(datos[i][13]) === "ABANDONADA") count++;
  }
  return count;
}

// ════════════════════════════════════════════════════════════════
// FUNCIONES DE PRUEBA
// ════════════════════════════════════════════════════════════════
function probarFlujoCompleto() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var base = { id_empleado: "E001", id_edificio: "ED01", turno: "Matutino", latitud: "19.44337740", longitud: "-99.18361210", precision: "10", origen: "ONLINE" };

  ["Entrada", "Inicio Comida", "Fin Comida", "Salida"].forEach(function(evento, i) {
    var p = Object.assign({}, base, { evento: evento, id_evento: "PRUEBA-" + new Date().getTime() + "-" + i });
    Logger.log(evento + ": " + doGet_Registro(ss, p).getContent());
    Utilities.sleep(1000);
  });

  Logger.log("Revisa Resumen_Turnos y Log — ambos deben reflejar el mismo id_sesion.");
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

function normalizarEncabezado(texto) {
  return String(texto || "").trim().toLowerCase();
}

function asegurarEncabezadosHoja(hoja, encabezadosRequeridos) {
  if (!hoja) throw new Error("No se recibió una hoja válida.");

  var lastCol = hoja.getLastColumn();
  var existentes = lastCol > 0
    ? hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizarEncabezado)
    : [];

  var siguientes = lastCol;
  encabezadosRequeridos.forEach(function(encabezado) {
    var buscado = normalizarEncabezado(encabezado);
    if (!buscado) return;

    if (existentes.indexOf(buscado) === -1) {
      siguientes++;
      hoja.getRange(1, siguientes).setValue(encabezado);
      existentes.push(buscado);
    }
  });
}

function asegurarEstructuraOperativa(ss) {
  var definiciones = {
    "Log": [
      "id_sesion", "fecha_hora_evento", "id_empleado", "nombre",
      "id_edificio", "nombre_edificio", "turno", "evento", "estatus",
      "latitud", "longitud", "precision_gps", "origen",
      "fecha_hora_servidor", "id_evento", "resultado"
    ],
    "Resumen_Turnos": [
      "id_sesion", "id_empleado", "nombre", "fecha", "edificio", "turno",
      "hora_entrada", "hora_salida", "hora_inicio_comida", "hora_fin_comida",
      "horas_comida", "horas_trabajadas", "estatus_entrada", "estatus_salida",
      "estatus_comida", "descuento_comida_horas", "observaciones"
    ],
    "Errores_Sistema": [
      "fecha", "funcion", "mensaje_error", "detalle_tecnico",
      "tipo_error", "id_evento", "id_sesion", "accion"
    ]
  };

  Object.keys(definiciones).forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) throw new Error("Falta la hoja requerida: " + nombre);
    asegurarEncabezadosHoja(hoja, definiciones[nombre]);
  });

  asegurarEncabezadosHoja(obtenerOCrearHojaAbandonadas(ss), [
    "id_sesion", "id_empleado", "nombre", "turno", "id_edificio", "nombre_edificio",
    "fecha_jornada", "hora_entrada", "hora_inicio_comida", "hora_fin_comida",
    "fecha_hora_abandono", "motivo", "horas_trabajadas", "estatus", "observaciones"
  ]);
}

function verificarEstructuraHojas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var requeridas = ["Empleados", "Edificios", "Log", "Sesiones_Activas", "Resumen_Turnos", "Errores_Sistema"];
  requeridas.forEach(function(n) {
    if (!ss.getSheetByName(n)) throw new Error("Falta la hoja requerida: " + n);
  });

  asegurarEstructuraOperativa(ss);

  Logger.log("Estructura verificada/actualizada. Backend " + VERSION_BACKEND +
             ". No se eliminó ni reordenó ninguna columna existente.");
}

function diagnosticarEmpleado() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var datos = ss.getSheetByName("Empleados").getDataRange().getValues();
  Logger.log("Total de filas: " + datos.length);
  for (var i = 1; i < datos.length; i++) {
    Logger.log("Fila " + (i+1) + ": ID='" + datos[i][0] + "' Nombre='" + datos[i][1] + "' Activo=" + datos[i][2]);
  }
}

// ════════════════════════════════════════════════════════════════
// RESPALDO AUTOMÁTICO DIARIO
// Crea una copia completa del archivo cada noche en una carpeta
// de Drive, y elimina automáticamente las copias de más de 30 días
// para que el espacio no crezca indefinidamente.
// ════════════════════════════════════════════════════════════════
const NOMBRE_CARPETA_RESPALDOS = "Respaldos_Sistema_Asistencia";
const DIAS_A_CONSERVAR_RESPALDOS = 30;

function crearRespaldoDiario() {
  var archivoOriginal = DriveApp.getFileById(SPREADSHEET_ID);
  var carpeta = obtenerOCrearCarpetaRespaldos(NOMBRE_CARPETA_RESPALDOS);

  var fechaStr = Utilities.formatDate(new Date(), "America/Mexico_City", "yyyy-MM-dd");
  archivoOriginal.makeCopy("Respaldo_Asistencia_" + fechaStr, carpeta);

  limpiarRespaldosViejos(carpeta, DIAS_A_CONSERVAR_RESPALDOS);
}

function obtenerOCrearCarpetaRespaldos(nombre) {
  var carpetas = DriveApp.getFoldersByName(nombre);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombre);
}

function limpiarRespaldosViejos(carpeta, diasAConservar) {
  var limite   = new Date(Date.now() - diasAConservar * 24 * 60 * 60 * 1000);
  var archivos = carpeta.getFiles();
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    if (archivo.getDateCreated() < limite) {
      archivo.setTrashed(true); // va a la papelera de Drive, no se borra permanente de inmediato
    }
  }
}

// ════════════════════════════════════════════════════════════════
// ARCHIVO ANUAL
// Crea una copia de corte del archivo al cierre de cada 31 de diciembre.
// No modifica ni borra la hoja operativa.
// ════════════════════════════════════════════════════════════════
const NOMBRE_CARPETA_ARCHIVOS_ANUALES = "Archivos_Anuales_Sistema_Asistencia";

function crearArchivoAnual() {
  var ahora = new Date();
  var mes = parseInt(Utilities.formatDate(ahora, "America/Mexico_City", "MM"), 10);
  var dia = parseInt(Utilities.formatDate(ahora, "America/Mexico_City", "dd"), 10);
  if (mes !== 12 || dia !== 31) return;

  var anio = Utilities.formatDate(ahora, "America/Mexico_City", "yyyy");
  var nombre = "Asistencia_" + anio + "_ARCHIVO";
  var carpeta = obtenerOCrearCarpetaRespaldos(NOMBRE_CARPETA_ARCHIVOS_ANUALES);
  var archivos = carpeta.getFilesByName(nombre);
  if (archivos.hasNext()) return;

  DriveApp.getFileById(SPREADSHEET_ID).makeCopy(nombre, carpeta);
}

// ════════════════════════════════════════════════════════════════
// SISTEMA DE ALERTAS DE ERROR
// Envuelve cualquier función automática: si truena, avisa por correo
// Y deja un registro permanente en la hoja "Errores_Sistema" — así
// un correo perdido entre otros mensajes no es la única forma de
// enterarte de que algo falló.
// ════════════════════════════════════════════════════════════════
function ejecutarConAlerta(nombreFuncion, funcion) {
  try {
    funcion();
  } catch(err) {
    registrarError(nombreFuncion, err);
    enviarAlertaError(nombreFuncion, err);
  }
}

function registrarError(nombreFuncion, err) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hoja = ss.getSheetByName("Errores_Sistema");
    if (hoja) {
      hoja.appendRow([
        new Date(), nombreFuncion, err.toString(), err.stack || "",
        "ERROR", "", "", ""
      ]);
    }
  } catch(e2) {
    // Si ni siquiera esto funciona, no hay más remedio que dejarlo pasar —
    // el correo de abajo sigue siendo el respaldo final.
  }
}

function enviarAlertaError(nombreFuncion, err) {
  try {
    GmailApp.sendEmail(
      CORREOS_REPORTE,
      "⚠️ Error en el Sistema de Asistencia — " + nombreFuncion,
      "La función '" + nombreFuncion + "' falló con el siguiente error:\n\n" +
      err.toString() + "\n\n" +
      "Hora: " + new Date().toLocaleString("es-MX") + "\n\n" +
      "Revisa la hoja 'Errores_Sistema' para el detalle completo, o el editor de Apps Script para depurar."
    );
  } catch(e2) {
    // Último recurso — si hasta el correo falla, no hay más que hacer desde aquí.
  }
}

// ════════════════════════════════════════════════════════════════
// FUNCIONES DE ACTIVADOR (TRIGGERS)
// Estas son las que debes seleccionar en "Activadores" — nunca
// selecciones directamente enviarReporteSemanal, revisarSesiones-
// Abandonadas ni crearRespaldoDiario, porque entonces un error ahí
// no generaría ninguna alerta.
// ════════════════════════════════════════════════════════════════
function trigger_ReporteSemanal() {
  ejecutarConAlerta("enviarReporteSemanal", enviarReporteSemanal);
}

function trigger_SesionesAbandonadas() {
  ejecutarConAlerta("revisarSesionesAbandonadas", revisarSesionesAbandonadas);
}

function trigger_RespaldoDiario() {
  ejecutarConAlerta("crearRespaldoDiario", crearRespaldoDiario);
}

function trigger_ArchivoAnual() {
  ejecutarConAlerta("crearArchivoAnual", crearArchivoAnual);
}

// ════════════════════════════════════════════════════════════════
// PRUEBAS DE LAS NUEVAS FUNCIONES
// ════════════════════════════════════════════════════════════════
function probarRespaldoManual() {
  crearRespaldoDiario();
  Logger.log("Respaldo creado. Revisa tu Drive — carpeta '" + NOMBRE_CARPETA_RESPALDOS + "'.");
}

function probarAlertaError() {
  // Simula una función que truena a propósito, para confirmar que
  // el correo de alerta y el registro en Errores_Sistema funcionan.
  ejecutarConAlerta("funcionDePrueba", function() {
    throw new Error("Este es un error de prueba — todo funciona si recibiste el correo.");
  });
  Logger.log("Revisa tu correo y la hoja Errores_Sistema.");
}

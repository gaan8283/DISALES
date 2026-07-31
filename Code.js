const SHEET_ID = "1qu1tQlfrAU5FralclQCu64SPfoW7wgIsTfoKYVlwJhI";

const CONFIG = {
  HOJAS: {
    CLIENTES: "CLIENTES",
    PRODUCTOS: "PRODUCTOS",
    UNIDADES: "UNIDADES",
    FACTURAS: "FACTURAS",
    DETALLE: "DETALLEFACTURAS",
    ABONOS: "ABONOS",
    CONFIG: "CONFIG"
  }
};

function getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(nombre);
  if (!sh) throw new Error('La hoja "' + nombre + '" no existe en el Google Sheet.');
  return sh;
}

// ==================== MANEJO DE FECHAS ====================
// IMPORTANTE: Google Sheets convierte automaticamente los strings tipo
// "29/07/2026" a objetos Date reales al escribirlos con appendRow/setValue.
// Esto significa que una misma columna puede contener, dependiendo de cuando
// se guardo el registro, un string "dd/MM/yyyy" O un objeto Date. Comparar
// fechas con === sobre strings (como hacia el codigo original) se rompe en
// cuanto la celda es un Date real, por eso los reportes no encontraban
// facturas. Estas funciones normalizan CUALQUIERA de los dos formatos a un
// objeto Date puro (sin hora) para poder comparar de forma confiable.

function parseFechaFlexible(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }
  const str = String(valor || "").trim();
  if (!str) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const p = str.split("-"); // yyyy-MM-dd
    return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    const p = str.split("/"); // dd/MM/yyyy
    return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  }
  return null;
}

function esMismaFecha(a, b) {
  if (!a || !b) return false;
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

// Devuelve siempre "dd/MM/yyyy" para mostrar en la UI, sin importar si la
// celda de origen es un Date real o un string en cualquiera de los 2 formatos.
function formatFechaDisplay(valor) {
  const d = parseFechaFlexible(valor);
  if (!d) return String(valor || "");
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function getNextFacturaNum() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet(CONFIG.HOJAS.CONFIG);
    const data = sh.getDataRange().getValues();
    let rowIndex = -1;
    let lastNum = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "LAST_FACTURA_NUM") {
        rowIndex = i + 1;
        lastNum = parseInt(data[i][1]) || 0;
        break;
      }
    }
    const nextNum = lastNum + 1;
    if (rowIndex > 0) {
      sh.getRange(rowIndex, 2).setValue(nextNum);
    } else {
      sh.appendRow(["LAST_FACTURA_NUM", nextNum, "Contador interno"]);
    }
    return nextNum;
  } finally {
    lock.releaseLock();
  }
}

function getNextClienteId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet(CONFIG.HOJAS.CLIENTES);
    const data = sh.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0] || "");
      const match = id.match(/^CLI-(\d+)$/);
      if (match) {
        const num = parseInt(match[1]);
        if (num > maxNum) maxNum = num;
      }
    }
    return "CLI-" + String(maxNum + 1).padStart(3, '0');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("DISALES - Mobile")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getClientes() {
  try {
    const sh = getSheet(CONFIG.HOJAS.CLIENTES);
    const data = sh.getDataRange().getValues();
    const clientes = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === "SI") {
        clientes.push({
          id: data[i][0],
          nombre: data[i][1],
          telefono: data[i][2] || "",
          direccion: data[i][3] || "",
          observaciones: data[i][4] || ""
        });
      }
    }
    return { success: true, data: clientes };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getProductos() {
  try {
    const sh = getSheet(CONFIG.HOJAS.PRODUCTOS);
    const data = sh.getDataRange().getValues();
    const productos = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === "SI") {
        productos.push({
          id: data[i][0],
          nombre: data[i][1],
          categoria: data[i][2] || "",
          precio: parseFloat(data[i][3]) || 0
        });
      }
    }
    return { success: true, data: productos };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getUnidades() {
  try {
    const sh = getSheet(CONFIG.HOJAS.UNIDADES);
    const data = sh.getDataRange().getValues();
    const unidades = [];
    for (let i = 1; i < data.length; i++) {
      unidades.push({ id: data[i][0], nombre: data[i][1] });
    }
    return { success: true, data: unidades };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function guardarCliente(data) {
  try {
    if (!data.nombre || data.nombre.trim() === "") {
      return { success: false, error: "El nombre del cliente es obligatorio." };
    }
    const clienteId = getNextClienteId();
    const sh = getSheet(CONFIG.HOJAS.CLIENTES);
    const hoy = new Date();
    const fechaHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    sh.appendRow([
      clienteId,
      data.nombre.trim(),
      data.telefono || "",
      data.direccion || "",
      data.observaciones || "",
      "SI",
      fechaHoy
    ]);
    return { success: true, clienteId: clienteId, mensaje: "Cliente registrado correctamente." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function validateVenta(data) {
  const errors = [];
  if (!data.clienteId) errors.push("Cliente requerido");
  if (!data.fecha || !/^\d{4}-\d{2}-\d{2}$/.test(data.fecha)) errors.push("Fecha invalida");
  if (!Array.isArray(data.detalle) || data.detalle.length === 0) errors.push("Agrega al menos un producto");
  if (!["CONTADO", "CREDITO"].includes(data.tipoVenta)) errors.push("Tipo de venta invalido");
  const sumaDetalle = data.detalle.reduce((sum, item) => sum + (parseFloat(item.subtotal) || 0), 0);
  const total = parseFloat(data.totalFactura) || 0;
  if (Math.abs(sumaDetalle - total) > 0.01) errors.push("El total no coincide con la suma del detalle");
  const abonoInicial = parseFloat(data.abonoInicial) || 0;
  if (data.tipoVenta === "CREDITO" && abonoInicial > total) errors.push("El abono no puede ser mayor al total");
  data.detalle.forEach((item, idx) => {
    if (parseFloat(item.cantidad) <= 0) errors.push("Cantidad invalida en linea " + (idx + 1));
    if (parseFloat(item.precio) < 0) errors.push("Precio invalido en linea " + (idx + 1));
  });
  return errors;
}

function guardarVenta(data) {
  const lock = LockService.getScriptLock();
  try {
    const errores = validateVenta(data);
    if (errores.length > 0) {
      return { success: false, error: errores.join(", ") };
    }
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shFacturas = ss.getSheetByName(CONFIG.HOJAS.FACTURAS);
    const shDetalle = ss.getSheetByName(CONFIG.HOJAS.DETALLE);
    if (!shFacturas) throw new Error("No se encontro la hoja FACTURAS");
    if (!shDetalle) throw new Error("No se encontro la hoja DETALLEFACTURAS");

    const facturaId = Utilities.getUuid().substring(0, 8);
    const numeroFactura = getNextFacturaNum();
    const abonoInicial = parseFloat(data.abonoInicial) || 0;
    const totalFactura = parseFloat(data.totalFactura) || 0;
    const saldoPendiente = data.tipoVenta === "CREDITO" ? totalFactura - abonoInicial : 0;
    const estado = saldoPendiente <= 0.01 ? "PAGADA" : (data.tipoVenta === "CREDITO" ? "PENDIENTE" : "PAGADA");

    const fechaVenta = parseFechaFlexible(data.fecha);
    if (!fechaVenta) throw new Error("Fecha invalida");
    const fechaStr = Utilities.formatDate(fechaVenta, Session.getScriptTimeZone(), "dd/MM/yyyy");

    const facturaRow = [
      facturaId,
      numeroFactura,
      fechaVenta,
      data.clienteId,
      data.tipoVenta,
      totalFactura,
      abonoInicial,
      saldoPendiente,
      estado,
      data.observaciones || ""
    ];

    const detalleRows = data.detalle.map(item => [
      Utilities.getUuid().substring(0, 8),
      facturaId,
      item.productoId,
      item.unidadId,
      parseFloat(item.cantidad),
      parseFloat(item.precio),
      parseFloat(item.subtotal)
    ]);

    shFacturas.appendRow(facturaRow);

    if (detalleRows.length > 0) {
      const startRow = shDetalle.getLastRow() + 1;
      const range = shDetalle.getRange(startRow, 1, detalleRows.length, detalleRows[0].length);
      range.setValues(detalleRows);
    }

    if (abonoInicial > 0) {
      const shAbonos = ss.getSheetByName(CONFIG.HOJAS.ABONOS);
      if (shAbonos) {
        shAbonos.appendRow([
          Utilities.getUuid().substring(0, 8),
          facturaId,
          data.clienteId,
          fechaVenta,
          abonoInicial,
          "Abono inicial"
        ]);
      }
    }

    return {
      success: true,
      facturaId: facturaId,
      numeroFactura: numeroFactura,
      mensaje: "Factura #" + numeroFactura + " guardada correctamente"
    };
  } catch (e) {
    console.error("Error en guardarVenta:", e);
    return { success: false, error: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getFacturasPendientes(clienteId) {
  try {
    if (!clienteId) return { success: false, error: "Cliente requerido" };
    const shFacturas = getSheet(CONFIG.HOJAS.FACTURAS);
    const data = shFacturas.getDataRange().getValues();
    const facturas = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === clienteId && data[i][8] === "PENDIENTE") {
        facturas.push({
          facturaId: data[i][0],
          numeroFactura: data[i][1],
          fecha: formatFechaDisplay(data[i][2]),
          totalFactura: parseFloat(data[i][5]) || 0,
          totalAbonos: parseFloat(data[i][6]) || 0,
          saldoPendiente: parseFloat(data[i][7]) || 0
        });
      }
    }
    return { success: true, data: facturas };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function guardarAbono(data) {
  const lock = LockService.getScriptLock();
  try {
    if (!data.clienteId) return { success: false, error: "Cliente requerido" };
    if (!data.facturaId) return { success: false, error: "Factura requerida" };
    if (!data.monto || parseFloat(data.monto) <= 0) return { success: false, error: "Monto invalido" };

    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shAbonos = ss.getSheetByName(CONFIG.HOJAS.ABONOS);
    const shFacturas = ss.getSheetByName(CONFIG.HOJAS.FACTURAS);
    if (!shAbonos) throw new Error("No se encontro la hoja ABONOS");
    if (!shFacturas) throw new Error("No se encontro la hoja FACTURAS");

    const facturasData = shFacturas.getDataRange().getValues();
    let facturaRowIndex = -1;
    let facturaData = null;
    for (let i = 1; i < facturasData.length; i++) {
      if (facturasData[i][0] === data.facturaId) {
        facturaRowIndex = i + 1;
        facturaData = facturasData[i];
        break;
      }
    }
    if (!facturaData) return { success: false, error: "Factura no encontrada" };

    const saldoActual = parseFloat(facturaData[7]) || 0;
    const montoAbono = parseFloat(data.monto);
    if (montoAbono > saldoActual) {
      return { success: false, error: "El abono (C$ " + montoAbono + ") no puede ser mayor al saldo (C$ " + saldoActual + ")" };
    }

    const abonoId = Utilities.getUuid().substring(0, 8);
    const hoy = new Date();
    const fechaHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

    shAbonos.appendRow([
      abonoId,
      data.facturaId,
      data.clienteId,
      fechaHoy,
      montoAbono,
      data.observaciones || ""
    ]);

    const nuevoAbonoTotal = (parseFloat(facturaData[6]) || 0) + montoAbono;
    const nuevoSaldo = (parseFloat(facturaData[5]) || 0) - nuevoAbonoTotal;
    const nuevoEstado = nuevoSaldo <= 0.01 ? "PAGADA" : "PENDIENTE";

    shFacturas.getRange(facturaRowIndex, 7).setValue(nuevoAbonoTotal);
    shFacturas.getRange(facturaRowIndex, 8).setValue(nuevoSaldo);
    shFacturas.getRange(facturaRowIndex, 9).setValue(nuevoEstado);

    return {
      success: true,
      abonoId: abonoId,
      mensaje: "Abono de C$ " + montoAbono.toFixed(2) + " registrado. Saldo restante: C$ " + nuevoSaldo.toFixed(2)
    };
  } catch (e) {
    console.error("Error en guardarAbono:", e);
    return { success: false, error: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getEstadoCuenta(clienteId) {
  try {
    if (!clienteId) return { success: false, error: "Cliente requerido" };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shFacturas = ss.getSheetByName(CONFIG.HOJAS.FACTURAS);
    const shAbonos = ss.getSheetByName(CONFIG.HOJAS.ABONOS);
    const shClientes = ss.getSheetByName(CONFIG.HOJAS.CLIENTES);

    const clientesData = shClientes.getDataRange().getValues();
    let clienteNombre = clienteId;
    for (let i = 1; i < clientesData.length; i++) {
      if (clientesData[i][0] === clienteId) {
        clienteNombre = clientesData[i][1];
        break;
      }
    }

    const facturasData = shFacturas.getDataRange().getValues();
    const facturas = [];
    for (let i = 1; i < facturasData.length; i++) {
      if (facturasData[i][3] === clienteId) {
        facturas.push({
          facturaId: facturasData[i][0],
          numeroFactura: facturasData[i][1],
          fecha: formatFechaDisplay(facturasData[i][2]),
          tipoVenta: facturasData[i][4],
          total: parseFloat(facturasData[i][5]) || 0,
          abonos: parseFloat(facturasData[i][6]) || 0,
          saldo: parseFloat(facturasData[i][7]) || 0,
          estado: facturasData[i][8]
        });
      }
    }

    const abonosData = shAbonos.getDataRange().getValues();
    const abonos = [];
    for (let i = 1; i < abonosData.length; i++) {
      if (abonosData[i][2] === clienteId) {
        abonos.push({
          abonoId: abonosData[i][0],
          facturaId: abonosData[i][1],
          fecha: formatFechaDisplay(abonosData[i][3]),
          monto: parseFloat(abonosData[i][4]) || 0,
          observaciones: abonosData[i][5] || ""
        });
      }
    }

    const totalDeuda = facturas.reduce((sum, f) => sum + f.saldo, 0);
    return {
      success: true,
      clienteNombre: clienteNombre,
      facturas: facturas,
      abonos: abonos,
      totalDeuda: totalDeuda
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getReporteDia(fechaStr) {
  try {
    const fechaObjetivo = parseFechaFlexible(fechaStr);
    if (!fechaObjetivo) return { success: false, error: "Fecha invalida" };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shFacturas = ss.getSheetByName(CONFIG.HOJAS.FACTURAS);
    const shAbonos = ss.getSheetByName(CONFIG.HOJAS.ABONOS);

    const facturasData = shFacturas.getDataRange().getValues();
    let totalVentas = 0, totalContado = 0, totalCredito = 0, numFacturas = 0;
    const facturasHoy = [];

    for (let i = 1; i < facturasData.length; i++) {
      const fechaCelda = parseFechaFlexible(facturasData[i][2]);
      if (esMismaFecha(fechaCelda, fechaObjetivo)) {
        const total = parseFloat(facturasData[i][5]) || 0;
        totalVentas += total; numFacturas++;
        if (facturasData[i][4] === "CONTADO") totalContado += total;
        else totalCredito += total;
        facturasHoy.push({ numero: facturasData[i][1], clienteId: facturasData[i][3], tipo: facturasData[i][4], total: total });
      }
    }

    const abonosData = shAbonos.getDataRange().getValues();
    let totalAbonos = 0;
    for (let i = 1; i < abonosData.length; i++) {
      const fechaCelda = parseFechaFlexible(abonosData[i][3]);
      if (esMismaFecha(fechaCelda, fechaObjetivo)) {
        totalAbonos += parseFloat(abonosData[i][4]) || 0;
      }
    }

    return {
      success: true,
      fecha: formatFechaDisplay(fechaObjetivo),
      totalVentas: totalVentas,
      totalContado: totalContado,
      totalCredito: totalCredito,
      totalAbonos: totalAbonos,
      numFacturas: numFacturas,
      facturas: facturasHoy
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

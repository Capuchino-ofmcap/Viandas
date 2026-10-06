/**
 * Viandas Fiesta Familiar — backend (Google Apps Script)
 *   - Productos: Producto | Precio | Categoría | Estado
 *   - Usuarios:  Nombre | FechaRegistro
 *   - Pedidos:   Fecha | Usuario | FormaPago | Producto | Cantidad | Precio | Subtotal | TotalPedido
 * Esquema final de Pedidos: una fila por producto, nunca una fila con el detalle completo en una sola celda.

 *
 * Cómo instalar:
 *   1. Abre tu proyecto de Apps Script (el que ya tienes desplegado).
 *   2. Reemplaza el contenido de tu archivo .gs por este código completo.
 *   3. Guarda y vuelve a implementar (Implementar > Administrar implementaciones
 *      > Editar > Nueva versión > Implementar). La URL /exec se mantiene igual.
 */

const VERSION_BACKEND = '2026-10-02-sheets-v10-async-stock';
const CABECERA_PEDIDOS = ['Fecha', 'Usuario', 'FormaPago', 'Producto', 'Cantidad', 'Precio', 'Subtotal', 'TotalPedido', 'TipoMovimiento', 'OperacionId'];
const CABECERA_PRODUCTOS = ['Producto', 'Precio', 'Categoría', 'Stock'];
const CABECERA_COLA_STOCK = ['OperacionId', 'Estado', 'Plan', 'Actualizado'];
const MAX_PRODUCTOS_MENU = 50;

function doGet(e) {
  const accion = (e.parameter.accion || e.parameter.action || 'menu');
  try {
    if (accion === 'salud') return salidaJSON(obtenerEstadoBackend());
    if (accion === 'stock') return salidaJSON({ stock: obtenerStockRapido() });
    if (accion === 'menu') return salidaJSON({
      menu: obtenerProductosRapidoDesdeHoja(),
      backend: VERSION_BACKEND
    });
    if (accion === 'usuarios') return salidaJSON({ usuarios: obtenerUsuarios() });
    if (accion === 'resumen') return salidaJSON(obtenerResumen(e.parameter.usuario || ''));
    if (accion === 'historial') return salidaJSON(obtenerHistorial(e.parameter.usuario || ''));
    return salidaJSON({ error: 'Acción no reconocida: ' + accion });
  } catch (err) {
    return salidaJSON({ error: err.message });
  }
}

function obtenerEstadoBackend() {
  const libro = obtenerLibro();
  const hojas = libro.getSheets().map(hoja => hoja.getName());
  return {
    ok: true,
    backend: VERSION_BACKEND,
    fecha: new Date().toISOString(),
    hojas: hojas,
    tieneProductos: hojas.indexOf('Productos') >= 0,
    tienePedidos: hojas.indexOf('Pedidos') >= 0
  };
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    let cuerpo = {};

    try {
      const texto = String(raw || '').trim();
      if (!texto) {
        cuerpo = {};
      } else if (texto.startsWith('{') || texto.startsWith('[')) {
        cuerpo = JSON.parse(texto);
      } else if (texto.includes('=')) {
        const params = {};
        texto.split('&').forEach((parte) => {
          if (!parte) return;
          const idx = parte.indexOf('=');
          const clave = idx >= 0 ? decodeURIComponent(parte.slice(0, idx)) : decodeURIComponent(parte);
          const valor = idx >= 0 ? decodeURIComponent(parte.slice(idx + 1)) : '';
          params[clave] = valor;
        });
        cuerpo = params;
      } else {
        cuerpo = JSON.parse(texto);
      }
    } catch (parseErr) {
      return salidaJSON({ ok: false, error: 'POST inválido: el cuerpo no es JSON válido.' });
    }

    const accion = cuerpo.accion || cuerpo.action || cuerpo.tipo;
    const datos = cuerpo.payload || cuerpo.data || {};

    if (accion === 'pedido') {
      const resultado = registrarPedido(datos);
      if (resultado && resultado.ok) {
        let estadoStock = { ok: true, colaVacia: false };
        for (let lote = 0; lote < 10; lote++) {
          estadoStock = procesarColaStock();
          if (!estadoStock || estadoStock.ok !== true || estadoStock.colaVacia === true) break;
        }
        resultado.stockPendiente = !estadoStock || estadoStock.ok !== true || estadoStock.colaVacia !== true;
        if (!estadoStock || estadoStock.ok !== true) {
          resultado.stockError = estadoStock && estadoStock.error
            ? estadoStock.error
            : 'No se pudo confirmar la actualización del stock.';
        }
      }
      return salidaJSON(resultado);
    }
    if (accion === 'usuario') return salidaJSON(registrarUsuario(datos));
    if (accion === 'producto') return salidaJSON(guardarProducto(datos));
    if (accion === 'ajuste') return salidaJSON(registrarAjuste(datos));
    if (accion === 'procesarStockPendiente') return salidaJSON(procesarColaStock());
    return salidaJSON({ ok: false, error: 'Acción no reconocida: ' + accion });
  } catch (err) {
    return salidaJSON({ ok: false, error: 'POST inválido: ' + err.message });
  }
}

function salidaJSON(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizarClaveTexto(valor) {
  return String(valor || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function generarOperacionId(prefijo) {
  return (prefijo || 'OP') + '-' + Date.now() + '-' + Math.floor(Math.random() * 100000);
}

function obtenerLibro() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const libro = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!libro) {
    throw new Error('No se encontró la hoja vinculada. Configura SPREADSHEET_ID o vuelve a vincular el proyecto a la hoja correcta.');
  }
  return libro;
}

function obtenerHoja(nombre, encabezados) {
  const libro = obtenerLibro();
  let hoja = libro.getSheetByName(nombre);

  if (!hoja) {
    hoja = libro.insertSheet(nombre);
    hoja.appendRow(encabezados);
  }

  return hoja;
}

function buscarFilaPorValor(hoja, columna, valor) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return 0;
  const resultado = hoja.getRange(2, columna, ultimaFila - 1, 1)
    .createTextFinder(String(valor))
    .matchEntireCell(true)
    .findNext();
  return resultado ? resultado.getRow() : 0;
}

function obtenerColaStock() {
  return obtenerHoja('ColaStock', CABECERA_COLA_STOCK);
}

function encolarStockPedido(operacionId) {
  const hojaCola = obtenerColaStock();
  const filaExistente = buscarFilaPorValor(hojaCola, 1, operacionId);
  if (filaExistente) return filaExistente;
  const filaNueva = hojaCola.getLastRow() + 1;
  hojaCola.getRange(filaNueva, 1, 1, CABECERA_COLA_STOCK.length)
    .setValues([[operacionId, 'pendiente', '', new Date()]]);
  return filaNueva;
}

// ---------- Productos ----------

function normalizarStock(stock) {
  const valor = Number(stock);
  if (!Number.isFinite(valor)) return 0;
  return Math.max(0, Math.floor(valor));
}

function obtenerProductosRapidoDesdeHoja() {
  const libro = obtenerLibro();
  const hoja = libro.getSheetByName('Productos');
  if (!hoja) return [];

  const filas = hoja.getRange(2, 1, MAX_PRODUCTOS_MENU, 4).getValues();
  const productos = filas
    .filter(fila => String(fila[0] || '').trim())
    .map(fila => ({
      Producto: String(fila[0]).trim(),
      Precio: Number(fila[1]) || 0,
      Categoría: String(fila[2] || '').trim(),
      Stock: normalizarStock(fila[3])
    }));
  return productos;
}

function obtenerStockRapido() {
  const stock = obtenerProductosRapidoDesdeHoja().map(producto => ({
    Producto: producto.Producto,
    Stock: producto.Stock
  }));
  return stock;
}

function guardarProducto(datos) {
  const hoja = obtenerHoja('Productos', CABECERA_PRODUCTOS);
  const modo = datos.modo;
  const bloqueo = LockService.getScriptLock();
  const lockOk = bloqueo.tryLock(10000);

  if (!lockOk) {
    return { ok: false, error: 'La edición de productos está ocupada. Intenta otra vez en unos segundos.' };
  }

  try {
    if (modo === 'agregar') {
      const p = datos.producto || {};
      const nombre = String(p.Producto || '').trim();
      if (!nombre) return { error: 'Falta el nombre del producto' };
      const stock = normalizarStock(p.Stock);
      hoja.appendRow([nombre, Number(p.Precio) || 0, p.Categoría || 'General', stock]);
      return { ok: true };
    }

    if (modo === 'editar') {
      const original = String(datos.original || '').trim();
      const p = datos.producto || {};
      const nombreNuevo = String(p.Producto || '').trim();
      const valores = hoja.getDataRange().getValues();
      const claveOriginal = normalizarClaveTexto(original);

      for (let i = 1; i < valores.length; i++) {
        const nombreActual = String(valores[i][0] || '').trim();
        if (normalizarClaveTexto(nombreActual) === claveOriginal) {
          const stock = normalizarStock(p.Stock);
          hoja.getRange(i + 1, 1, 1, 4).setValues([[
            nombreNuevo || nombreActual,
            Number(p.Precio) || 0,
            p.Categoría || 'General',
            stock
          ]]);
          return { ok: true };
        }
      }
      return { error: 'No se encontró el producto: ' + original };
    }

    return { error: 'Modo no reconocido: ' + modo };
  } finally {
    bloqueo.releaseLock();
  }
}

// ---------- Usuarios ----------

function normalizarNombre(nombre) {
  return String(nombre || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function obtenerUsuarios() {
  const hoja = obtenerHoja('Usuarios', ['Nombre', 'FechaRegistro']);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  const nombres = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues()
    .map(fila => String(fila[0] || '').trim())
    .filter(Boolean);

  const validos = [];
  const vistos = new Set();
  nombres.forEach(nombre => {
    const key = normalizarNombre(nombre);
    if (!vistos.has(key)) {
      vistos.add(key);
      validos.push(nombre);
    }
  });

  return validos;
}

function registrarUsuario(datos) {
  const modo = String(datos.modo || 'crear').trim().toLowerCase();
  const nombre = String(datos.nombre || '').trim();
  const nombreActual = String(datos.nombreActual || '').trim();

  if (!nombre) return { error: 'El nombre está vacío' };

  const hojaUsuarios = obtenerHoja('Usuarios', ['Nombre', 'FechaRegistro']);
  const hojaPedidos = obtenerHoja('Pedidos', CABECERA_PEDIDOS);

  if (modo === 'editar') {
    if (!nombreActual) return { error: 'Falta el nombre actual para editar' };

    const registros = hojaUsuarios.getDataRange().getValues().slice(1);
    const nombreActualNormalizado = normalizarNombre(nombreActual);
    let encontrado = false;

    for (let i = 0; i < registros.length; i++) {
      const valor = String(registros[i][0] || '').trim();
      if (valor && normalizarNombre(valor) === nombreActualNormalizado) {
        hojaUsuarios.getRange(i + 2, 1).setValue(nombre);
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el nombre actual para reemplazar: ' + nombreActual };
    }

    const filasPedidos = hojaPedidos.getDataRange().getValues();
    let pedidosActualizados = 0;
    const filasActualizadas = filasPedidos.map((fila, index) => {
      const filaNueva = fila.slice();
      while (filaNueva.length < CABECERA_PEDIDOS.length) filaNueva.push('');
      if (index === 0) return filaNueva;
      const usuarioPedido = String(fila[1] || '').trim();
      if (normalizarNombre(usuarioPedido) === nombreActualNormalizado) {
        pedidosActualizados += 1;
        filaNueva[1] = nombre;
        return filaNueva;
      }
      return filaNueva;
    });

    if (pedidosActualizados > 0) {
      const ultimaFila = filasActualizadas.length;
      const ultimaColumna = Math.max(1, filasActualizadas[0] ? filasActualizadas[0].length : CABECERA_PEDIDOS.length);
      hojaPedidos.getRange(1, 1, ultimaFila, ultimaColumna).setValues(filasActualizadas);
    }

    return {
      ok: true,
      modo: 'editar',
      actualizado: true,
      pedidosActualizados: pedidosActualizados
    };
  }

  const existentes = obtenerUsuarios();
  const yaExiste = existentes.some(n => normalizarNombre(n) === normalizarNombre(nombre));

  if (!yaExiste) {
    hojaUsuarios.appendRow([nombre, new Date()]);
  }

  return { ok: true, existia: yaExiste, modo: 'crear' };
}

// ---------- Pedidos ----------

function registrarPedido(datos) {
  const items = Array.isArray(datos.items) ? datos.items : [];
  if (items.length === 0) return { ok: false, error: 'El pedido no tiene productos' };

  const bloqueo = LockService.getScriptLock();
  const lockOk = bloqueo.tryLock(8000);
  if (!lockOk) {
    return { ok: false, error: 'La app está ocupada. Intenta confirmar el pedido en unos segundos.' };
  }

  try {
    const hoja = obtenerHoja('Pedidos', CABECERA_PEDIDOS);
    const operacionId = String(datos.requestId || '').trim() || generarOperacionId('VTA');
    const filaExistente = buscarFilaPorValor(hoja, 10, operacionId);
    if (filaExistente) {
      const totalPedidoExistente = Number(hoja.getRange(filaExistente, 8).getValue()) || 0;
      encolarStockPedido(operacionId);
      return {
        ok: true,
        duplicado: true,
        operacionId: operacionId,
        totalPedido: totalPedidoExistente,
        stockPendiente: true
      };
    }

    for (const item of items) {
      const nombre = String(item.producto || '').trim();
      const cantidad = Number(item.cantidad) || 0;
      if (!nombre) return { ok: false, error: 'Hay un producto sin nombre' };
      if (cantidad <= 0) return { ok: false, error: 'La cantidad de ' + nombre + ' debe ser mayor que cero' };
    }

    const totalPedido = items.reduce((total, item) => {
      return total + ((Number(item.cantidad) || 0) * (Number(item.precio) || 0));
    }, 0);
    const filas = items.map(item => {
      const cantidad = Number(item.cantidad) || 0;
      const precio = Number(item.precio) || 0;
      return [
        new Date(),
        String(datos.usuario || '').trim() || 'Sin nombre',
        String(datos.formaPago || '').trim() || 'efectivo',
        String(item.producto || '').trim(),
        cantidad,
        precio,
        cantidad * precio,
        totalPedido,
        'venta',
        operacionId
      ];
    });

    const filaInicial = hoja.getLastRow() === 0 ? 1 : hoja.getLastRow() + 1;
    hoja.getRange(filaInicial, 1, filas.length, CABECERA_PEDIDOS.length).setValues(filas);
    encolarStockPedido(operacionId);

    return {
      ok: true,
      filasAgregadas: filas.length,
      totalPedido: totalPedido,
      operacionId: operacionId,
      stockPendiente: true
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo guardar el pedido: ' + err.message };
  } finally {
    bloqueo.releaseLock();
  }
}

function procesarColaStock() {
  const bloqueo = LockService.getScriptLock();
  const lockOk = bloqueo.tryLock(8000);
  if (!lockOk) return { ok: false, error: 'La app está ocupada. El stock queda pendiente.' };

  try {
    const hojaCola = obtenerColaStock();
    const ultimaFilaCola = hojaCola.getLastRow();
    if (ultimaFilaCola < 2) return { ok: true, procesados: 0, colaVacia: true };

    const hojaPedidos = obtenerHoja('Pedidos', CABECERA_PEDIDOS);
    const hojaProductos = obtenerHoja('Productos', CABECERA_PRODUCTOS);
    const cola = hojaCola.getRange(2, 1, ultimaFilaCola - 1, CABECERA_COLA_STOCK.length).getValues();
    let procesados = 0;

    for (let indice = 0; indice < cola.length; indice++) {
      const operacionId = String(cola[indice][0] || '').trim();
      const estado = String(cola[indice][1] || '').trim();
      if (!operacionId || estado === 'aplicado') continue;

      const filaCola = indice + 2;
      let plan;
      let productosActuales;
      if (estado === 'procesando' && cola[indice][2]) {
        plan = JSON.parse(String(cola[indice][2]));
        productosActuales = hojaProductos.getDataRange().getValues();
      } else {
        const ultimaFilaPedidos = hojaPedidos.getLastRow();
        if (ultimaFilaPedidos < 2) throw new Error('No se encontró el pedido ' + operacionId);
        const coincidencias = hojaPedidos.getRange(2, 10, ultimaFilaPedidos - 1, 1)
          .createTextFinder(operacionId)
          .matchEntireCell(true)
          .findAll();
        if (coincidencias.length === 0) throw new Error('No se encontró el pedido ' + operacionId);

        const primeraFila = coincidencias[0].getRow();
        const ultimaFila = coincidencias[coincidencias.length - 1].getRow();
        const filasPedido = hojaPedidos.getRange(
          primeraFila,
          1,
          ultimaFila - primeraFila + 1,
          CABECERA_PEDIDOS.length
        ).getValues().filter(fila => String(fila[9] || '').trim() === operacionId);
        const cantidadesPorProducto = {};
        filasPedido.forEach(fila => {
          const clave = normalizarClaveTexto(fila[3]);
          cantidadesPorProducto[clave] = (cantidadesPorProducto[clave] || 0) + (Number(fila[4]) || 0);
        });

        productosActuales = hojaProductos.getDataRange().getValues();
        plan = productosActuales.slice(1).reduce((cambios, fila, productoIndice) => {
          const nombre = String(fila[0] || '').trim();
          const clave = normalizarClaveTexto(nombre);
          const cantidad = cantidadesPorProducto[clave] || 0;
          if (nombre && cantidad > 0) {
            const stockAntes = normalizarStock(fila[3]);
            cambios.push({
              fila: productoIndice + 2,
              producto: nombre,
              stockAntes: stockAntes,
              stockDespues: Math.max(0, stockAntes - cantidad)
            });
            delete cantidadesPorProducto[clave];
          }
          return cambios;
        }, []);

        const productosNoEncontrados = Object.keys(cantidadesPorProducto);
        if (productosNoEncontrados.length > 0) {
          throw new Error('No se encontró el producto: ' + productosNoEncontrados.join(', '));
        }
        hojaCola.getRange(filaCola, 2, 1, 3).setValues([['procesando', JSON.stringify(plan), new Date()]]);
      }

      let hayCambios = false;
      plan.forEach(cambio => {
        const indiceProducto = cambio.fila - 1;
        if (!productosActuales[indiceProducto]) throw new Error('No se encontró la fila de ' + cambio.producto);
        const stockActual = normalizarStock(productosActuales[indiceProducto][3]);
        if (stockActual === cambio.stockDespues) return;
        if (stockActual !== cambio.stockAntes) {
          throw new Error('El stock de ' + cambio.producto + ' cambió durante la actualización.');
        }
        productosActuales[indiceProducto][3] = cambio.stockDespues;
        hayCambios = true;
      });

      if (hayCambios && productosActuales.length > 1) {
        hojaProductos.getRange(2, 4, productosActuales.length - 1, 1)
          .setValues(productosActuales.slice(1).map(fila => [normalizarStock(fila[3])]));
      }
      hojaCola.getRange(filaCola, 2, 1, 3).setValues([['aplicado', JSON.stringify(plan), new Date()]]);
      procesados++;
      const quedanPendientes = cola.slice(indice + 1).some(fila =>
        String(fila[0] || '').trim() && String(fila[1] || '').trim() !== 'aplicado'
      );
      return { ok: true, procesados: procesados, colaVacia: !quedanPendientes };
    }

    return { ok: true, procesados: procesados, colaVacia: true };
  } catch (err) {
    return { ok: false, error: 'No se pudo actualizar el stock: ' + err.message };
  } finally {
    bloqueo.releaseLock();
  }
}

function obtenerResumen(usuarioFiltro) {
  const libro = obtenerLibro();
  const hoja = libro.getSheetByName('Pedidos');
  let filas = [];
  if (hoja) {
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila > 1) {
      filas = hoja.getRange(2, 1, ultimaFila - 1, CABECERA_PEDIDOS.length).getValues();
    }
  }
  const filtro = normalizarNombre(usuarioFiltro);
  const productos = {};
  const formasPago = {
    efectivo: { totalRecaudado: 0, totalUnidades: 0 },
    yape: { totalRecaudado: 0, totalUnidades: 0 }
  };
  let totalRecaudado = 0;
  let totalUnidades = 0;
  let totalOfrendas = 0;

  filas.forEach(fila => {
    const usuario = String(fila[1] || '').trim();
    if (filtro && normalizarNombre(usuario) !== filtro) return;

    const tipoMovimiento = normalizarClaveTexto(fila[8] || 'venta');
    const nombre = String(fila[3] || '').trim();
    const cantidad = Number(fila[4]) || 0;
    const precio = Number(fila[5]) || 0;
    const subtotal = Number(fila[6]) || cantidad * precio;

    const formaPagoTexto = normalizarClaveTexto(fila[2]);
    const formaPago = formaPagoTexto.includes('yape') || formaPagoTexto.includes('plin') ? 'yape' : 'efectivo';

    if (tipoMovimiento === 'ofrenda') {
      formasPago[formaPago].totalRecaudado += subtotal;
      totalRecaudado += subtotal;
      totalOfrendas += subtotal;
      return;
    }

    if (!nombre || cantidad === 0) return;

    const claveProducto = normalizarClaveTexto(nombre);
    if (!productos[claveProducto]) {
      productos[claveProducto] = {
        producto: nombre,
        totalUnidades: 0,
        totalRecaudado: 0,
        precios: {}
      };
    }

    const resumenProducto = productos[claveProducto];
    const clavePrecio = String(precio);
    if (!resumenProducto.precios[clavePrecio]) {
      resumenProducto.precios[clavePrecio] = {
        precio: precio,
        unidades: 0,
        subtotal: 0
      };
    }

    resumenProducto.precios[clavePrecio].unidades += cantidad;
    resumenProducto.precios[clavePrecio].subtotal += subtotal;
    resumenProducto.totalUnidades += cantidad;
    resumenProducto.totalRecaudado += subtotal;
    formasPago[formaPago].totalUnidades += cantidad;
    formasPago[formaPago].totalRecaudado += subtotal;
    totalUnidades += cantidad;
    totalRecaudado += subtotal;
  });

  const listaProductos = Object.keys(productos).filter(clave => {
    const producto = productos[clave];
    return producto.totalUnidades !== 0 || producto.totalRecaudado !== 0;
  }).map(clave => {
    const producto = productos[clave];
    producto.precios = Object.keys(producto.precios)
      .map(precio => producto.precios[precio])
      .sort((a, b) => b.precio - a.precio);
    return producto;
  }).sort((a, b) => a.producto.localeCompare(b.producto, 'es', { sensitivity: 'base' }));

  return {
    ok: true,
    usuario: usuarioFiltro || '',
    totalRecaudado: totalRecaudado,
    totalOfrendas: totalOfrendas,
    totalUnidades: totalUnidades,
    formasPago: formasPago,
    productos: listaProductos,
    usuarios: obtenerUsuarios()
  };
}

function registrarAjuste(datos) {
  const devueltos = Array.isArray(datos.devueltos) ? datos.devueltos : [];
  const nuevos = Array.isArray(datos.nuevos) ? datos.nuevos : [];
  if (devueltos.length === 0 && nuevos.length === 0) {
    return { ok: false, error: 'Agrega al menos un producto devuelto o nuevo.' };
  }

  const bloqueo = LockService.getScriptLock();
  const lockOk = bloqueo.tryLock(8000);
  if (!lockOk) {
    return { ok: false, error: 'La operación está ocupada. Intenta nuevamente en unos segundos.' };
  }

  try {
    const hojaProductos = obtenerHoja('Productos', CABECERA_PRODUCTOS);
    const productosActuales = hojaProductos.getDataRange().getValues();
    const productosPorClave = {};
    productosActuales.slice(1).forEach((fila, indice) => {
      const nombre = String(fila[0] || '').trim();
      if (nombre) {
        productosPorClave[normalizarClaveTexto(nombre)] = {
          fila: indice + 2,
          nombre: nombre,
          stock: normalizarStock(fila[3])
        };
      }
    });

    const formaPago = String(datos.formaPago || 'efectivo').trim() || 'efectivo';
    const usuario = String(datos.usuario || '').trim() || 'Sin nombre';
    const operacionId = generarOperacionId('AJU');
    const tipoMovimiento = nuevos.length > 0 ? 'cambio' : 'devolucion';
    const movimientosStock = {};
    const filas = [];
    let totalOperacion = 0;

    devueltos.forEach(item => {
      const nombre = String(item.producto || '').trim();
      const clave = normalizarClaveTexto(nombre);
      const producto = productosPorClave[clave];
      const cantidad = Math.abs(Number(item.cantidad) || 0);
      const precio = Math.max(0, Number(item.precio) || 0);
      if (!producto) throw new Error('No se encontró el producto devuelto: ' + nombre);
      if (cantidad <= 0) throw new Error('La cantidad devuelta de ' + nombre + ' debe ser mayor que cero.');

      const subtotal = -(cantidad * precio);
      filas.push([
        new Date(), usuario, formaPago, producto.nombre, -cantidad, precio,
        subtotal, totalOperacion, tipoMovimiento, operacionId
      ]);
      totalOperacion += subtotal;
      if (item.reponerStock) {
        movimientosStock[clave] = (movimientosStock[clave] || 0) + cantidad;
      }
    });

    nuevos.forEach(item => {
      const nombre = String(item.producto || '').trim();
      const clave = normalizarClaveTexto(nombre);
      const producto = productosPorClave[clave];
      const cantidad = Math.abs(Number(item.cantidad) || 0);
      const precio = Math.max(0, Number(item.precio) || 0);
      if (!producto) throw new Error('No se encontró el producto nuevo: ' + nombre);
      if (cantidad <= 0) throw new Error('La cantidad nueva de ' + nombre + ' debe ser mayor que cero.');

      const subtotal = cantidad * precio;
      filas.push([
        new Date(), usuario, formaPago, producto.nombre, cantidad, precio,
        subtotal, totalOperacion, tipoMovimiento, operacionId
      ]);
      totalOperacion += subtotal;
      movimientosStock[clave] = (movimientosStock[clave] || 0) - cantidad;
    });

    const tipoDiferencia = String(datos.diferenciaTipo || '').trim().toLowerCase();
    if (tipoDiferencia === 'ofrenda' && totalOperacion < 0) {
      const ofrenda = -totalOperacion;
      filas.push([
        new Date(), usuario, formaPago, 'Ofrenda', 1, ofrenda,
        ofrenda, totalOperacion + ofrenda, 'ofrenda', operacionId
      ]);
      totalOperacion += ofrenda;
    }

    Object.keys(movimientosStock).forEach(clave => {
      const producto = productosPorClave[clave];
      const stockFinal = producto.stock + movimientosStock[clave];
      if (stockFinal < 0) {
        throw new Error('No hay stock suficiente de ' + producto.nombre + ' para completar el cambio.');
      }
    });

    const hoja = obtenerHoja('Pedidos', CABECERA_PEDIDOS);
    const filaInicial = hoja.getLastRow() === 0 ? 1 : hoja.getLastRow() + 1;
    filas.forEach(fila => { fila[7] = totalOperacion; });
    hoja.getRange(filaInicial, 1, filas.length, CABECERA_PEDIDOS.length).setValues(filas);

    Object.keys(movimientosStock).forEach(clave => {
      const producto = productosPorClave[clave];
      hojaProductos.getRange(producto.fila, 4).setValue(producto.stock + movimientosStock[clave]);
    });

    return {
      ok: true,
      operacionId: operacionId,
      totalOperacion: totalOperacion,
      diferencia: totalOperacion,
      tipoMovimiento: tipoMovimiento
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo registrar la operación: ' + err.message };
  } finally {
    if (lockOk) bloqueo.releaseLock();
  }
}

function obtenerHistorial(usuarioFiltro) {
  const libro = obtenerLibro();
  const hoja = libro.getSheetByName('Pedidos');
  const filtro = normalizarNombre(usuarioFiltro);
  const limite = 15;
  if (!filtro) return { ok: true, operaciones: [] };
  if (!hoja || hoja.getLastRow() <= 1) {
    return { ok: true, operaciones: [] };
  }

  const filaUltima = hoja.getLastRow();
  const filasPorBloque = 500;
  let operacionesIncluidas = 0;
  let filaFinal = filaUltima;
  let operacionActual = null;
  const operaciones = [];

  function finalizarOperacion() {
    if (!operacionActual) return false;
    operacionActual.items.reverse();
    if (operacionesIncluidas < limite) {
      operaciones.push(operacionActual);
      operacionesIncluidas++;
    } else {
      return true;
    }
    operacionActual = null;
    return false;
  }

  while (filaFinal >= 2 && operacionesIncluidas < limite) {
    const filaInicial = Math.max(2, filaFinal - filasPorBloque + 1);
    const filas = hoja.getRange(
      filaInicial,
      1,
      filaFinal - filaInicial + 1,
      CABECERA_PEDIDOS.length
    ).getValues();

    for (let indice = filas.length - 1; indice >= 0; indice--) {
      const fila = filas[indice];
      const numeroFila = filaInicial + indice;
      const usuario = String(fila[1] || '').trim();
      if (filtro && normalizarNombre(usuario) !== filtro) continue;
      const operacionId = String(fila[9] || '').trim() || 'LEG-' + numeroFila;
      if (!operacionActual || operacionActual.operacionId !== operacionId) {
        if (finalizarOperacion()) break;
        operacionActual = {
          operacionId: operacionId,
          fecha: fila[0],
          usuario: usuario || 'Sin nombre',
          formaPago: String(fila[2] || '').trim() || 'efectivo',
          tipoMovimiento: String(fila[8] || 'venta').trim() || 'venta',
          total: 0,
          items: []
        };
      }

      const subtotal = Number(fila[6]) || 0;
      operacionActual.total += subtotal;
      operacionActual.items.push({
        producto: String(fila[3] || '').trim(),
        cantidad: Number(fila[4]) || 0,
        precio: Number(fila[5]) || 0,
        subtotal: subtotal
      });
    }

    filaFinal = filaInicial - 1;
  }

  if (operacionesIncluidas < limite) finalizarOperacion();

  return { ok: true, operaciones: operaciones };
}

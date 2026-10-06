// ===== Viandas Fiesta Familiar — lógica compartida =====

// 🛑 URL de tu Web App de Apps Script (la misma para todas las páginas).
// Si cambias de deployment, puedes sobreescribirla desde la consola del navegador:
// window.__VIANDAS_SCRIPT_URL__ = 'https://script.google.com/macros/s/.../exec';
const SCRIPT_URL = (typeof window !== 'undefined' && window.__VIANDAS_SCRIPT_URL__) ||
  'https://script.google.com/macros/s/AKfycbw8KhCsnI2-abm73K9XuTr_s5jG0vylgDU13tvy3yY-thXVdKpOpZxDluiruukuBp0nmA/exec';

// Claves de almacenamiento local (por dispositivo/navegador)
const CLAVE_NOMBRE = 'viandasFF_nombre';
const CLAVE_ULTIMO_NOMBRE = 'viandasFF_ultimo_nombre';
const CLAVE_PAGO = 'viandasFF_pago';
const CLAVE_CACHE_MENU = 'viandasFF_menu_cache_v1';
const TTL_CACHE_MENU = 600000;

function leerCacheJSON(clave, ttlMs = 300000) {
  try {
    const raw = localStorage.getItem(clave);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    if (typeof parsed.expiresAt !== 'number' || Date.now() > parsed.expiresAt) {
      localStorage.removeItem(clave);
      return null;
    }
    return parsed.data;
  } catch (_) {
    return null;
  }
}

function guardarCacheJSON(clave, data, ttlMs = 300000) {
  try {
    localStorage.setItem(clave, JSON.stringify({ data, expiresAt: Date.now() + ttlMs }));
  } catch (_) {
    // Si el navegador tiene almacenamiento lleno, ignoramos la caché.
  }
}

function conTimeout(promise, ms, mensaje) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(mensaje)), ms);
    promise.then(
      value => {
        clearTimeout(timer);
        resolve(value);
      },
      error => {
        clearTimeout(timer);
        reject(error);
      }
    );
  });
}

const Sesion = {
  getNombre() {
    return (localStorage.getItem(CLAVE_NOMBRE) || '').trim();
  },
  setNombre(nombre) {
    const valor = (nombre || '').trim();
    if (!valor) return;
    localStorage.setItem(CLAVE_NOMBRE, valor);
    localStorage.setItem(CLAVE_ULTIMO_NOMBRE, valor);
  },
  getUltimoNombre() {
    return (localStorage.getItem(CLAVE_ULTIMO_NOMBRE) || '').trim();
  },
  getPago() {
    return localStorage.getItem(CLAVE_PAGO) || 'efectivo';
  },
  setPago(pago) {
    localStorage.setItem(CLAVE_PAGO, pago);
  },
  cerrar() {
    const ultimo = this.getNombre();
    if (ultimo) {
      localStorage.setItem(CLAVE_ULTIMO_NOMBRE, ultimo);
    }
    localStorage.removeItem(CLAVE_NOMBRE);
    localStorage.removeItem(CLAVE_PAGO);
  }
};

// Si no hay nombre guardado, vuelve a la pantalla de inicio.
function exigirSesion() {
  if (!Sesion.getNombre()) {
    window.location.href = 'index.html';
  }
}

// Lecturas (GET) — sin preflight, funcionan directo con fetch.
async function apiGet(accion, params = {}, reintentos = 1, cacheKey = null, ttlMs = 300000) {
  const url = new URL(SCRIPT_URL);
  url.searchParams.set('accion', accion);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));

  const cacheFallback = cacheKey ? leerCacheJSON(cacheKey, ttlMs) : null;

  for (let intento = 0; intento <= reintentos; intento++) {
    try {
      const res = await conTimeout(
        fetch(url.toString(), { cache: 'no-store' }),
        accion === 'resumen' || accion === 'historial' || accion === 'menu' || accion === 'stock' ? 30000 : 10000,
        'La respuesta de la lista tardó demasiado.'
      );

      if (!res.ok) throw new Error('No se pudo conectar (HTTP ' + res.status + ')');
      const texto = await res.text();
      if (/^\s*<!doctype html|^\s*<html/i.test(texto)) {
        throw new Error('La Web App está solicitando acceso de Google en lugar de devolver JSON. Publícala como “Cualquiera” y ejecuta como tú.');
      }
      let datos;
      try {
        datos = JSON.parse(texto);
      } catch (_) {
        throw new Error('La Web App devolvió una respuesta inválida en lugar de JSON.');
      }
      if (datos.error) throw new Error(datos.error);
      if (cacheKey) guardarCacheJSON(cacheKey, datos, ttlMs);
      return datos;
    } catch (err) {
      if (cacheFallback && intento === reintentos) return cacheFallback;
      if (intento === reintentos) throw err;
      await new Promise(resolve => setTimeout(resolve, 500 + intento * 1000));
    }
  }
}

// Escrituras (POST) — Content-Type text/plain evita el preflight CORS
// y permite leer la respuesta real de Apps Script.
async function apiPost(accion, payload = {}, reintentos = 1) {
  const body = JSON.stringify({ accion, payload });

  for (let intento = 0; intento <= reintentos; intento++) {
    try {
      const res = await conTimeout(
        fetch(SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body,
          cache: 'no-store'
        }),
        accion === 'pedido' || accion === 'ajuste' ? 60000 : 10000,
        accion === 'pedido'
          ? 'La respuesta del pedido tardó demasiado. El backend puede estar ocupado o la hoja está lenta.'
          : 'La respuesta de la operación tardó demasiado. El backend puede estar ocupado o la hoja está lenta.'
      );

      const texto = await res.text();
      let datos = {};

      if (texto) {
        const textoLimpio = texto.trim();
        if (textoLimpio) {
          try {
            datos = JSON.parse(textoLimpio);
          } catch (_) {
            datos = { error: textoLimpio };
          }
        }
      }

      if (!res.ok) {
        throw new Error('No se pudo conectar (HTTP ' + res.status + '). ' + (datos && datos.error ? datos.error : ''));
      }
      if (datos && datos.error) throw new Error(datos.error);

      return datos;
    } catch (err) {
      if (intento === reintentos) {
        const respuestaServidor = String(err && err.message ? err.message : '');
        const mensaje = respuestaServidor.includes('Failed to fetch') || respuestaServidor.includes('HTTP 403') || respuestaServidor.includes('HTTP 401')
          ? 'No se pudo conectar con la Web App de Apps Script. Revisa que esté publicada y accesible para “Anyone” (o “Anyone with a Google account”), y que la URL actual en app-common.js sea la correcta.'
          : respuestaServidor || 'No se pudo completar la operación.';
        throw new Error(mensaje);
      }
        await new Promise(resolve => setTimeout(resolve, 500 + intento * 1000));
    }
  }
}

// Cajón de navegación: cada página incluye el mismo bloque de HTML del
// drawer; esta función lo conecta (abrir/cerrar + nombre del usuario).
function textoPago(pago) {
  return pago === 'yape' ? 'Pago por Yape / Plin' : 'Pago en efectivo';
}

function iniciarCajonNavegacion() {
  const boton = document.getElementById('menuToggle');
  const cajon = document.getElementById('drawer');
  const fondo = document.getElementById('drawerOverlay');
  if (!boton || !cajon || !fondo) return;

  const abrir = () => { cajon.classList.add('open'); fondo.classList.add('open'); };
  const cerrar = () => { cajon.classList.remove('open'); fondo.classList.remove('open'); };

  boton.addEventListener('click', abrir);
  fondo.addEventListener('click', cerrar);

  const nombre = Sesion.getNombre();
  const pago = Sesion.getPago();
  const elNombre = document.getElementById('drawerNombre');
  const elPago = document.getElementById('drawerPago');
  const elInicial = document.getElementById('drawerInicial');

  if (elNombre) elNombre.textContent = nombre || 'Invitado';
  if (elPago) elPago.textContent = textoPago(pago);
  if (elInicial) elInicial.textContent = (nombre ? nombre[0] : '?').toUpperCase();
}

function formatoSoles(numero) {
  return 'S/. ' + Number(numero || 0).toFixed(2);
}

if ('serviceWorker' in navigator && window.isSecureContext) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .catch(error => console.warn('No se pudo registrar la app instalable.', error));
  });
}

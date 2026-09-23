/* ═══════════════════════════════════════════════════════════════════════════
   GEP201 · capa común de sorteo_gep201.html e index.html
   Todo lo que habla con el backend (Apps Script GEP201_API) vive aquí:
   clave de acceso, lecturas, cola de envío con confirmación real, nombres y
   fecha local. La presentación queda en cada HTML.

   Contrato con el backend (ver scripts/GEP201_API_actualizado.gs):
   - Toda llamada lleva `token` (clave guardada en este dispositivo). Si el
     backend tiene API_TOKEN configurado y no coincide → {status:'unauthorized'}.
   - Cada POST lleva `_id` único; el backend lo anota y responde
     {status:'duplicate'} si ya lo procesó → reintentar nunca duplica filas.
   - Los POST van como text/plain (sin preflight CORS), así la respuesta del
     backend sí se puede leer y la cola solo borra lo que el Sheet confirmó.
   ═══════════════════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';

  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzSur7nSfzlTjhTF286vnu34wrY5-U1sYkgztpGRk0i8o9Mh9yGNCrjgqH5_p9qqOwMWA/exec';
  const TOKEN_KEY = 'gep201_token';
  const PENDIENTES_KEY = 'gep201_pendientes_registro';   // misma llave que la versión anterior: no se pierde lo ya encolado
  const TIMEOUT_MS = 30000;   // Apps Script puede tardar en "despertar"; reintentar ya es seguro (idempotente)

  /* ── almacenamiento local (tolerante a modo privado / bloqueos) ── */
  const store = {
    get(k, def) { try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch (e) { return def; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
    raw(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
    setRaw(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} },
  };

  /* ── clave de acceso ── */
  const getToken = () => store.raw(TOKEN_KEY);
  const setToken = (t) => store.setRaw(TOKEN_KEY, String(t || '').trim());
  const clearToken = () => store.del(TOKEN_KEY);

  class AuthError extends Error { constructor() { super('unauthorized'); this.name = 'AuthError'; } }

  async function withTimeout(url, opts) {
    const ctrl = new AbortController();
    const to = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
    try { return await fetch(url, { ...opts, signal: ctrl.signal }); }
    finally { clearTimeout(to); }
  }

  /* ── lecturas ── */
  // Un reintento automático: justo después de un despliegue (o si Apps Script
  // "duerme"), la primera respuesta a veces llega como página de error de Google.
  async function get(action, intentos = 3) {
    const url = SCRIPT_URL + '?action=' + encodeURIComponent(action) + '&token=' + encodeURIComponent(getToken());
    let data;
    try { data = await (await withTimeout(url, {})).json(); }
    catch (e) {
      if (intentos > 1) { await new Promise(r => setTimeout(r, intentos === 3 ? 1500 : 3000)); return get(action, intentos - 1); }
      throw e;
    }
    if (data.status === 'unauthorized') { clearToken(); throw new AuthError(); }
    if (data.status !== 'ok') throw new Error(data.msg || 'respuesta inválida');
    return data;
  }

  async function post(payload) {
    const res = await withTimeout(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...payload, token: getToken() }),
    });
    return res.json();
  }

  /* ── nómina normalizada: {code, name, nombres, apellidos, group} ── */
  async function cargarNomina() {
    const data = await get('nomina');
    return (data.alumnos || []).map(a => ({
      code: String(a.codigo), name: a.nombre, nombres: a.nombres || '', apellidos: a.apellidos || '',
      group: String(a.grupo).trim(),
    }));
  }
  async function cargarConteos() { return (await get('sorteo_conteos')).conteos || {}; }
  async function sorteosRecientes() { return (await get('sorteo_estado')).sorteos || []; }

  /* ── nombres: regla fija "primer nombre + primer apellido" ── */
  function nombre(st) {
    let given = st.nombres || '', family = st.apellidos || '';
    if (!given || !family) {
      const p = String(st.name || '').split(' ');
      given = given || p[0] || ''; family = family || p.slice(1).join(' ');
    }
    const first = given.split(' ')[0];
    const fam1 = family.split(' ')[0] || given.split(' ')[1] || '';
    return { given, family, first, fam1, short: (first + ' ' + fam1).trim(), full: (given + ' ' + family).trim() };
  }

  /* ── fecha local (no UTC: en Lima, después de las 19:00 UTC ya es "mañana") ── */
  function hoyLocal() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  const nuevoId = () => Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 10);

  /* ═══════════ COLA DE ENVÍO ═══════════
     El registro se guarda en el equipo de inmediato y se envía en segundo
     plano. Solo sale de la cola cuando el backend responde ok/duplicate.
     Sobrevive a mala señal, pestaña cerrada o batería agotada. */
  let pendientes = store.get(PENDIENTES_KEY, []);
  if (!Array.isArray(pendientes)) pendientes = [];
  let enviando = false;
  const estado = { ultimaLatenciaMs: null, error: null, sinClave: false };
  const oyentes = new Set();

  const guardar = () => store.set(PENDIENTES_KEY, pendientes);
  const avisar = () => { const s = syncState(); oyentes.forEach(fn => { try { fn(s); } catch (e) {} }); };

  function syncState() {
    return {
      pendientes: pendientes.length,
      esperandoMs: pendientes.length ? Date.now() - (pendientes[0]._queuedAt || Date.now()) : 0,
      enviando,
      ...estado,
    };
  }

  function encolar(payload) {
    const item = { ...payload, _id: payload._id || nuevoId(), _queuedAt: Date.now() };
    pendientes.push(item);
    guardar(); avisar(); procesar();
    return item._id;
  }

  async function procesar() {
    if (enviando || !pendientes.length) return;
    enviando = true; avisar();
    while (pendientes.length) {
      const item = pendientes[0];
      let data;
      try { data = await post(item); }
      catch (e) { estado.error = null; break; }            // sin red / timeout: se reintenta luego
      if (data.status === 'unauthorized') {
        clearToken(); estado.sinClave = true; break;
      }
      if (data.status !== 'ok' && data.status !== 'duplicate') {
        estado.error = data.msg || 'error del Sheet';      // no se borra: mejor atascado y visible que perdido
        break;
      }
      estado.error = null; estado.sinClave = false;
      estado.ultimaLatenciaMs = Date.now() - item._queuedAt;
      console.log('[GEP201] ' + data.status + ' tras ' + (estado.ultimaLatenciaMs / 1000).toFixed(1) + 's · ' + (item.alumno || '') + ' · ' + (item.action || ''));
      pendientes.shift(); guardar(); avisar();
    }
    enviando = false; avisar();
  }

  function onSync(fn) { oyentes.add(fn); fn(syncState()); return () => oyentes.delete(fn); }

  setInterval(procesar, 15000);
  global.addEventListener('online', procesar);
  global.addEventListener('pageshow', procesar);

  /* ═══════════ PANTALLA DE CLAVE ═══════════
     Solo aparece si el backend rechaza la llamada. Hereda colores de la
     página (--bg, --accent) y su tipografía. Campo tipo password: la clave
     no queda a la vista si la pantalla está proyectada. */
  function pedirClave(mensaje) {
    return new Promise(resolve => {
      let ov = document.getElementById('gep-auth');
      if (ov) ov.remove();
      ov = document.createElement('div');
      ov.id = 'gep-auth';
      ov.setAttribute('role', 'dialog');
      ov.setAttribute('aria-modal', 'true');
      ov.innerHTML =
        '<form class="gep-auth-box">' +
          '<div class="gep-auth-t">Clave del curso</div>' +
          '<div class="gep-auth-m">' + (mensaje || 'Se pide una sola vez en este dispositivo.') + '</div>' +
          '<input type="password" autocomplete="current-password" autocapitalize="off" autocorrect="off" spellcheck="false" required aria-label="Clave">' +
          '<button type="submit">Entrar</button>' +
        '</form>';
      const css = document.createElement('style');
      css.textContent =
        '#gep-auth{position:fixed;inset:0;z-index:1000;display:flex;align-items:center;justify-content:center;padding:24px;background:var(--bg)}' +
        '.gep-auth-box{width:100%;max-width:340px;display:flex;flex-direction:column;gap:12px}' +
        '.gep-auth-t{font-size:20px;font-weight:600}' +
        '.gep-auth-m{font-size:14px;opacity:.75;line-height:1.4}' +
        '.gep-auth-box input{font:inherit;font-size:16px;color:inherit;background:transparent;border:1px solid currentColor;border-radius:10px;padding:0 14px;min-height:48px;opacity:.9}' +
        '.gep-auth-box button{font:inherit;font-size:16px;font-weight:600;min-height:48px;border:none;border-radius:10px;cursor:pointer;background:var(--accent);color:var(--on-accent,var(--accent-ink,#fff));touch-action:manipulation}';
      ov.appendChild(css);
      document.body.appendChild(ov);
      const input = ov.querySelector('input');
      ov.querySelector('form').addEventListener('submit', e => {
        e.preventDefault();
        setToken(input.value);
        ov.remove();
        estado.sinClave = false;
        procesar();
        resolve();
      });
      setTimeout(() => input.focus(), 50);
    });
  }

  /* Ejecuta fn(); si el backend pide clave, la solicita y reintenta. */
  async function conClave(fn) {
    let msg;
    for (;;) {
      try { return await fn(); }
      catch (e) {
        if (!(e instanceof AuthError)) throw e;
        await pedirClave(msg);
        msg = 'Clave incorrecta. Intenta de nuevo.';
      }
    }
  }

  function describirError(e) {
    if (!e) return 'desconocido';
    if (e.name === 'AbortError') return 'el servidor tardó demasiado (señal lenta)';
    if (e instanceof SyntaxError) return 'Google respondió con una página de error en vez de datos';
    if (e instanceof TypeError) return 'sin conexión con el servidor';
    return e.message || String(e);
  }

  global.GEP = {
    SCRIPT_URL, store, get, post, cargarNomina, cargarConteos, sorteosRecientes,
    nombre, hoyLocal, nuevoId, encolar, procesar, onSync, syncState,
    pedirClave, conClave, AuthError, getToken, clearToken, describirError,
  };
})(window);

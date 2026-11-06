/* DISALES · capa Firebase
   Hace que la misma app (app/disales.html) funcione fuera de Claude:
   - datos compartidos en Firestore (en tiempo real y con trabajo sin conexión)
   - entrada con cuenta de Google, solo correos autorizados
   - numeración de documentos con transacción del servidor (nunca se repite)
   - descarga de PDF y respaldo como archivo normal
   La app pide sus servicios con claude.use('db' | 'user' | 'downloads'); aquí se responden con Firebase.
   window.DISALES_FB (config, admin, correos) lo define el script de construcción. */
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, collection, setDoc, updateDoc, getDoc,
  getDocs, deleteDoc, onSnapshot, runTransaction, serverTimestamp, FieldPath, deleteField }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import { crearPaquetes, enPaquetes } from './paquetes.js';

const CFG = window.DISALES_FB;
const app = initializeApp(CFG.config);
const auth = getAuth(app);
auth.languageCode = 'es';
let fs;
try {
  fs = initializeFirestore(app, { ignoreUndefinedProperties: true,
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
} catch (e) { fs = initializeFirestore(app, { ignoreUndefinedProperties: true }); }

/* errores de Firebase -> códigos que la app ya sabe explicar */
const MAP = { 'permission-denied': 'invalid_argument', 'unauthenticated': 'revoked', 'unavailable': 'unavailable',
  'resource-exhausted': 'resource_exhausted', 'deadline-exceeded': 'unavailable', 'aborted': 'busy', 'failed-precondition': 'unavailable' };
const wrap = p => p.catch(e => { throw { code: MAP[e && e.code] || (e && e.code) || 'error', message: (e && e.message) || String(e) }; });
const snapDoc = s => ({ exists: s.exists(), data: () => s.data() });

const PQ = crearPaquetes({ doc, collection, setDoc, updateDoc, getDoc, getDocs, onSnapshot, FieldPath, deleteField }, fs);
const errSnap = err => e => err && err({ code: MAP[e.code] || e.code, message: e.message });

const db = {
  doc(path) {
    const [c, id] = path.split('/');
    if (enPaquetes(c) && id) return {   // ventas, cobros, productos… viven en paquetes (ver paquetes.js)
      set: d => wrap(PQ.set(c, id, d)),
      update: d => wrap(PQ.update(c, id, d)),
      get: () => wrap(PQ.get(c, id)),
      delete: () => wrap(PQ.delete(c, id)),
      acquire: async () => ({ acquired: true }),
      onSnapshot: () => () => {},
    };
    const r = doc(fs, path);
    return {
      set: d => wrap(setDoc(r, d)),
      update: d => wrap(updateDoc(r, d)),
      get: () => wrap(getDoc(r)).then(snapDoc),
      delete: () => wrap(deleteDoc(r)),
      acquire: async () => ({ acquired: true }),   // la numeración real usa nextSeq (transacción)
      onSnapshot: (cb, err) => onSnapshot(r, s => cb(snapDoc(s)), e => err && err({ code: MAP[e.code] || e.code, message: e.message })),
    };
  },
  collection(c) {
    if (enPaquetes(c)) return { onSnapshot: (cb, err) => PQ.escuchar(c, cb, errSnap(err)) };
    return {
      onSnapshot: (cb, err) => onSnapshot(collection(fs, c), { includeMetadataChanges: true },
        s => cb({ docs: s.docs.map(d => ({ id: d.id, data: () => d.data() })), metadata: { hasPendingWrites: s.metadata.hasPendingWrites } }),
        e => err && err({ code: MAP[e.code] || e.code, message: e.message })),
    };
  },
  /* siguiente número de factura / cotización / recibo, sin repetirse aunque facturen los dos a la vez */
  nextSeq(kind) {
    const r = doc(fs, 'seq/' + kind);
    return wrap(runTransaction(fs, async t => {
      const s = await t.get(r);
      const n = ((s.exists() && Number(s.data().n)) || 0) + 1;
      t.set(r, { n });
      return n;
    }));
  },
};

/* nombres y fotos de los usuarios, para mostrar quién hizo cada cosa */
const perfiles = {};
let me = null;
const user = {
  id: async () => me.uid,
  me: async () => ({ id: me.uid, name: me.displayName || me.email }),
  isOwner: async () => (me.email || '').toLowerCase() === CFG.admin,
  canEdit: async () => (me.email || '').toLowerCase() === CFG.admin,
  can: async () => true,
  profiles: async ids => Object.fromEntries(ids.map(i => [i, perfiles[i] || { name: '', avatarUrl: '' }])),
};

const downloads = {
  save: async ({ filename, data }) => {
    const tipo = /\.pdf$/i.test(filename) ? 'application/pdf' : /\.xlsx$/i.test(filename) ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/json';
    const blob = data instanceof Blob ? data : new Blob([data], { type: tipo });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = filename; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
  },
};

/* ===== pantalla de entrada ===== */
function pantalla(html) {
  let el = document.getElementById('fbLogin');
  if (!el) {
    el = document.createElement('div'); el.id = 'fbLogin';
    el.style.cssText = 'position:fixed;inset:0;z-index:100;background:var(--bg,#F4F2EF);display:flex;align-items:center;justify-content:center;padding:16px';
    document.body.appendChild(el);
  }
  el.innerHTML = `<div style="max-width:380px;width:100%;background:#fff;border:1px solid #DDD6CC;border-radius:16px;padding:24px;display:flex;flex-direction:column;gap:14px;text-align:center;font-family:Figtree,system-ui,sans-serif">
    <img src="icono-192.png" alt="DISALES" style="width:96px;height:96px;margin:0 auto;border-radius:20px">
    <h2 style="margin:0;font-size:1.3rem">DISALES</h2>${html}</div>`;
}
const quitarPantalla = () => { const el = document.getElementById('fbLogin'); if (el) el.remove(); };
window.DISALES_FB.salir = () => signOut(auth).then(() => location.reload());

async function entrar(btn) {
  btn.disabled = true; btn.textContent = 'Abriendo Google…';
  const prov = new GoogleAuthProvider(); prov.setCustomParameters({ prompt: 'select_account' });
  try { await signInWithPopup(auth, prov); }
  catch (e) {
    if (e && (e.code === 'auth/popup-blocked' || e.code === 'auth/operation-not-supported-in-this-environment')) return signInWithRedirect(auth, prov);
    btn.disabled = false; btn.textContent = 'Entrar con Google';
    if (e && e.code !== 'auth/popup-closed-by-user' && e.code !== 'auth/cancelled-popup-request')
      document.getElementById('fbErr').textContent = 'No se pudo entrar: ' + (e.message || e.code);
  }
}
function pedirEntrada(msg) {
  pantalla(`<div style="color:#6A635B">Ventas, cobros e inventario.<br>Entrá con tu cuenta de Google autorizada.</div>
    <button id="fbBtn" style="appearance:none;border:0;border-radius:12px;padding:14px;font:600 1rem system-ui;background:#D2601C;color:#fff;cursor:pointer">Entrar con Google</button>
    <div id="fbErr" style="color:#B3261E;font-size:.9rem">${msg || ''}</div>`);
  document.getElementById('fbBtn').onclick = e => entrar(e.currentTarget);
}

getRedirectResult(auth).catch(() => {});
let listo = false;
onAuthStateChanged(auth, async u => {
  if (!u) { pedirEntrada(); return; }
  const email = (u.email || '').toLowerCase();
  if (!CFG.correos.includes(email)) {
    await signOut(auth);
    pedirEntrada(`El correo ${email} no está autorizado. Pedile al administrador que lo agregue.`);
    return;
  }
  me = u;
  // primera vez con paquetes: pasa los datos de la versión anterior (no borra nada; se puede repetir)
  try {
    const m = await getDoc(doc(fs, 'meta/paquetes'));
    if (!m.exists()) {
      pantalla(`<div>Actualizando la base de datos…</div><div id="fbAv" style="color:#6A635B;font-size:.95rem">Esto pasa una sola vez.</div>`);
      const n = await PQ.migrar(k => { const el = document.getElementById('fbAv'); if (el) el.textContent = `${k} registros pasados…`; });
      await setDoc(doc(fs, 'meta/paquetes'), { v: 1, registros: n, fecha: serverTimestamp(), por: email });
    }
  } catch (e) {
    pantalla(`<div style="color:#B3261E">No se pudo actualizar la base de datos: ${(e && (e.message || e.code)) || e}</div>
      <button onclick="location.reload()" style="appearance:none;border:0;border-radius:12px;padding:14px;font:600 1rem system-ui;background:#D2601C;color:#fff;cursor:pointer">Reintentar</button>`);
    return;
  }
  quitarPantalla();
  // guarda el perfil propio y escucha los de los demás
  setDoc(doc(fs, 'usuarios/' + u.uid), { nombre: u.displayName || email, email, foto: u.photoURL || '', visto: serverTimestamp() }, { merge: true }).catch(() => {});
  onSnapshot(collection(fs, 'usuarios'), s => {
    s.docs.forEach(d => { const x = d.data(); perfiles[d.id] = { name: x.nombre || x.email || '', avatarUrl: x.foto || '' }; });
    window.dispatchEvent(new Event('disales-perfiles'));
  }, () => {});
  if (!listo) { listo = true; window.__fbResolve({ use: k => ({ db, user, downloads })[k] || null }); }
});

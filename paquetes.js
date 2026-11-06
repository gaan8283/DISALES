/* DISALES · paquetes
   Firestore cobra (y limita en el plan gratis) por DOCUMENTO leído. Si cada venta fuera un documento,
   abrir la app leería miles. Aquí cada colección se guarda en un número fijo de "paquetes":
   el registro <id> vive en el campo r.<id> del paquete paq_<colección>/b<n>, con n = hash(id) mod K.
   Abrir la app lee solo los paquetes (unas 180 lecturas en total), sin importar cuántas ventas haya.
   - set    -> reemplaza solo ese registro (mergeFields)
   - update -> cambia solo esos campos del registro
   - delete -> quita solo ese registro
   Dos personas guardando a la vez en el mismo paquete no se pisan: cada escritura toca solo su campo.
   Capacidad: 1 MiB por paquete; con ventas en 64 paquetes caben decenas de miles de facturas. */

export const K = { ventas: 64, pagos: 32, entregas: 16, devoluciones: 8, movs: 16, compras: 8, traslados: 8,
  cierres: 8, productos: 4, clientes: 4, proveedores: 2 };
export const enPaquetes = c => Object.prototype.hasOwnProperty.call(K, c);

/* hash estable (FNV-1a de 32 bits) para repartir los registros */
export function cubeta(c, id) {
  let h = 0x811c9dc5;
  for (let i = 0; i < id.length; i++) { h ^= id.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
  return 'b' + (h % K[c]);
}

/* api: funciones de Firestore (doc, collection, setDoc, updateDoc, getDoc, getDocs, onSnapshot, FieldPath, deleteField)
   fs: la base. Devuelve operaciones con la misma forma que usa la app. */
export function crearPaquetes(api, fs) {
  const ref = (c, id) => api.doc(fs, 'paq_' + c, cubeta(c, id));
  const campo = (id, ...k) => new api.FieldPath('r', id, ...k);
  return {
    set: (c, id, data) => api.setDoc(ref(c, id), { r: { [id]: data } }, { mergeFields: [campo(id)] }),
    update: (c, id, data) => {
      const args = [];
      for (const [k, v] of Object.entries(data)) if (v !== undefined) args.push(campo(id, k), v);
      if (!args.length) return Promise.resolve();
      return api.updateDoc(ref(c, id), ...args);
    },
    delete: (c, id) => api.updateDoc(ref(c, id), campo(id), api.deleteField()).catch(e => { if (e && e.code === 'not-found') return; throw e; }),
    get: async (c, id) => {
      const s = await api.getDoc(ref(c, id));
      const r = s.exists() ? (s.data().r || {}) : {};
      const x = r[id];
      return { exists: x != null, data: () => x };
    },
    /* escucha todos los paquetes de la colección y los entrega como una lista plana de registros */
    escuchar: (c, cb, err) => api.onSnapshot(api.collection(fs, 'paq_' + c), { includeMetadataChanges: true }, s => {
      const docs = [];
      s.docs.forEach(d => { const r = (d.data() || {}).r || {}; for (const id in r) if (r[id] != null) docs.push({ id, data: () => r[id] }); });
      cb({ docs, metadata: { hasPendingWrites: s.metadata.hasPendingWrites } });
    }, err),
    /* pasa los datos guardados como documentos sueltos (versión anterior) a paquetes; se puede repetir sin duplicar */
    migrar: async (avance) => {
      let n = 0;
      for (const c of Object.keys(K)) {
        const s = await api.getDocs(api.collection(fs, c));
        if (s.empty) continue;
        const grupos = {};
        s.docs.forEach(d => { const b = cubeta(c, d.id); (grupos[b] = grupos[b] || {})[d.id] = d.data(); });
        for (const [b, recs] of Object.entries(grupos)) {
          await api.setDoc(api.doc(fs, 'paq_' + c, b), { r: recs }, { mergeFields: Object.keys(recs).map(id => campo(id)) });
          n += Object.keys(recs).length; avance && avance(n);
        }
      }
      return n;
    },
  };
}

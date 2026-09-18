// api/sync.js — se despliega solo al conectar el repo a Vercel, sin pasos extra.
// Se dispara visitando /api/sync (o con el botón "Sincronizar desde Excel" de index.html).
//
// IMPORTANTE sobre qué pisa y qué no:
// - tiempos_base y "cant_base" en modelos: SIEMPRE se actualizan desde el Excel — son datos
//   de referencia, nadie los edita a mano en la página.
// - cantidad / fecha_limite / prioridad en modelos, y todo Capacidad (operarios/turnos/%extra):
//   son las palancas que la gente edita EN VIVO desde la página. Este script nunca las pisa en
//   un modelo que ya existe — solo les pone un valor inicial la primera vez que aparece ese modelo.

const { createClient } = require('@supabase/supabase-js');

// Google Sheets (gviz) devuelve las fechas como texto tipo "Date(2026,9,15)" —
// hay que parsearlo a mano, un Date() común no lo entiende.
function gvizFechaAISO(v) {
  if (v == null) return null;
  if (typeof v === 'string') {
    const m = v.match(/Date\((\d+),(\d+),(\d+)/);
    if (m) {
      const anio = Number(m[1]), mes = Number(m[2]), dia = Number(m[3]);
      const mm = String(mes + 1).padStart(2, '0'); // el mes de gviz viene 0-indexado
      const dd = String(dia).padStart(2, '0');
      return `${anio}-${mm}-${dd}`;
    }
  }
  const d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

module.exports = async (req, res) => {
  try {
    const supabase = createClient(
      process.env.SUPABASE_URL,
      process.env.SUPABASE_SERVICE_ROLE_KEY
    );
    const sheetId = process.env.SHEET_ID;

    async function leerHoja(hoja, rango) {
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(hoja)}&range=${rango}`;
      const r = await fetch(url);
      const text = await r.text();
      const match = text.match(/(?<=\().*(?=\);)/s);
      if (!match) throw new Error(`No se pudo leer la hoja "${hoja}" — revisá que el Sheet esté compartido como lector público.`);
      return JSON.parse(match[0]).table.rows;
    }

    // ---- Plan_Comercial: cant_base siempre se actualiza; el resto solo si el modelo es nuevo ----
    // (va PRIMERO: tiempos_base tiene una foreign key a modelos, así que los modelos
    // tienen que existir antes de poder insertar sus horas por etapa)
    const filasPC = await leerHoja('Plan_Comercial', 'B6:F11');
    const { data: existentes } = await supabase.from('modelos').select('nombre');
    const nombresExistentes = new Set((existentes || []).map(m => m.nombre));

    const nuevos = [];
    const actualizarCantBase = [];
    filasPC.filter(r => r.c[0]?.v).forEach(r => {
      const nombre = r.c[0].v;
      const cant_base = r.c[1]?.v || 1;
      if (nombresExistentes.has(nombre)) {
        actualizarCantBase.push({ nombre, cant_base });
      } else {
        nuevos.push({
          nombre,
          cant_base,
          cantidad: r.c[2]?.v || 0,
          fecha_limite: gvizFechaAISO(r.c[3]?.v) || new Date().toISOString().slice(0, 10),
          prioridad: r.c[4]?.v || 99,
        });
      }
    });
    if (nuevos.length) {
      const { error } = await supabase.from('modelos').insert(nuevos);
      if (error) throw error;
    }
    for (const m of actualizarCantBase) {
      const { error } = await supabase.from('modelos').update({ cant_base: m.cant_base }).eq('nombre', m.nombre);
      if (error) throw error;
    }

    // ---- Tiempos_Base: siempre se pisa entero, no es editable en la página ----
    // (va DESPUÉS de modelos, por la foreign key)
    const filasTB = await leerHoja('Tiempos_Base', 'B3:E44');
    const tiemposBase = filasTB
      .filter(r => r.c[0]?.v)
      .map(r => ({
        modelo: r.c[0].v,
        orden_etapa: r.c[1].v,
        etapa: r.c[2].v,
        horas_base: r.c[3]?.v || 0,
      }));
    const { error: errTB } = await supabase.from('tiempos_base').upsert(tiemposBase, { onConflict: 'modelo,orden_etapa' });
    if (errTB) throw errTB;

    // ---- Capacidad: se siembra una sola vez (si la tabla está vacía); después es 100% editable en la página ----
    const { data: capExistente } = await supabase.from('capacidad').select('orden_etapa');
    if (!capExistente || capExistente.length === 0) {
      const filasCap = await leerHoja('Capacidad', 'B3:D9');
      const capacidad = filasCap.filter(r => r.c[0]?.v != null).map(r => ({
        orden_etapa: r.c[0].v,
        etapa: r.c[1].v,
        operarios: r.c[2]?.v || 1,
        turnos: 1,
        horas_extra_pct: 0,
      }));
      const { error } = await supabase.from('capacidad').insert(capacidad);
      if (error) throw error;
    }

    res.status(200).json({ ok: true, synced: new Date().toISOString(), modelos_nuevos: nuevos.length, modelos_actualizados: actualizarCantBase.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || String(err) });
  }
};

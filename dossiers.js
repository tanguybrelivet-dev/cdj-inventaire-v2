import { supabase } from "./supabaseClient.js";

const TABLE = "dossiers";

export async function listDossiers() {
  if (!supabase) return { data: [], error: new Error("Supabase non configuré") };
  const { data, error } = await supabase
    .from(TABLE)
    .select("id, created_at, updated_at, name, address, status, total, lot_count")
    .order("updated_at", { ascending: false });
  return { data: data || [], error };
}

export async function loadDossier(id) {
  if (!supabase) return { data: null, error: new Error("Supabase non configuré") };
  const { data, error } = await supabase.from(TABLE).select("*").eq("id", id).single();
  return { data, error };
}

export async function saveDossier({ id, caseInfo, lots, zones, categories }) {
  if (!supabase) return { data: null, error: new Error("Supabase non configuré") };
  const payload = {
    name: caseInfo.name || "Dossier sans nom",
    address: caseInfo.address || "",
    status: "en_cours",
    total: (lots || []).reduce((sum, l) => sum + (l.value || 0), 0),
    lot_count: (lots || []).length,
    payload: { caseInfo, lots, zones, categories },
    updated_at: new Date().toISOString(),
  };
  if (id) {
    const { data, error } = await supabase.from(TABLE).update(payload).eq("id", id).select().single();
    return { data, error };
  }
  const { data, error } = await supabase
    .from(TABLE)
    .insert({ ...payload, created_at: new Date().toISOString() })
    .select()
    .single();
  return { data, error };
}

export async function deleteDossier(id) {
  if (!supabase) return { error: new Error("Supabase non configuré") };
  const { error } = await supabase.from(TABLE).delete().eq("id", id);
  return { error };
}

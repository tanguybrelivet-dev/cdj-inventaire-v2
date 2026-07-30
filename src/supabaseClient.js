import { createClient } from "@supabase/supabase-js";

// ---------------------------------------------------------------------------
// À COMPLÉTER : collez ici l'URL du projet et la clé "anon public" trouvées
// dans Supabase → Project Settings → API. Ces deux valeurs ne sont PAS
// secrètes (elles sont conçues pour être exposées côté client), à condition
// d'avoir activé les règles de sécurité (RLS) fournies dans README.md.
// ---------------------------------------------------------------------------
const SUPABASE_URL = "https://VOTRE-PROJET.supabase.co";
const SUPABASE_ANON_KEY = "VOTRE-CLE-ANON-PUBLIC";

export const supabaseConfigured =
  !SUPABASE_URL.includes("VOTRE-PROJET") && !SUPABASE_ANON_KEY.includes("VOTRE-CLE");

export const supabase = supabaseConfigured ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

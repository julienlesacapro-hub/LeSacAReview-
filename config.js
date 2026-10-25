// Configuration des comptes utilisateurs (Supabase). Voir README, section « Comptes utilisateurs ».
// Ces deux valeurs sont PUBLIQUES par conception (la protection repose sur les règles de supabase/setup.sql).
// Ne JAMAIS y mettre la clé « service_role ».
// Laissées vides : l'appli fonctionne en mode « code d'accès partagé » avec data.enc.
export const CONFIG = {
  supabaseUrl: '', // ex. https://abcdefghijkl.supabase.co
  supabaseAnonKey: '', // clé « anon » / « publishable »
};

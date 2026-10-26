// Configuration des comptes utilisateurs (Supabase). Voir README, section « Comptes utilisateurs ».
// Ces deux valeurs sont PUBLIQUES par conception (la protection repose sur les règles de supabase/setup.sql).
// Ne JAMAIS y mettre la clé « service_role ».
// Laissées vides : l'appli fonctionne en mode « code d'accès partagé » avec data.enc.
export const CONFIG = {
  supabaseUrl: 'https://hifjdjetaxrxumtyvyov.supabase.co/', // ex. https://abcdefghijkl.supabase.co
  supabaseAnonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhpZmpkamV0YXhyeHVtdHl2eW92Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI5MzMwMjAsImV4cCI6MjEwODUwOTAyMH0.MAC3nFJetx9wy1EjKw8uVHmSE15iMNP7VAJL3oa5e1o', // clé « anon » / « publishable »
};

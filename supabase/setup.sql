-- À exécuter UNE FOIS dans Supabase : SQL Editor → New query → coller → Run.
-- Met en place : bucket privé des données, droits de lecture/écriture, et changement
-- obligatoire du mot de passe initial (contrôlé côté serveur, pas seulement dans l'appli).

-- 1. Bucket privé contenant le fichier de données (aucun accès public)
insert into storage.buckets (id, name, public)
values ('data', 'data', false)
on conflict (id) do update set public = false;

-- 2. Dès qu'un mot de passe est changé (par l'utilisateur ou via « mot de passe oublié »),
--    le compte est marqué « pwd_changed » dans app_metadata (non modifiable par l'utilisateur).
create or replace function public.mark_password_changed()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  if new.encrypted_password is distinct from old.encrypted_password then
    new.raw_app_meta_data :=
      coalesce(new.raw_app_meta_data, '{}'::jsonb) || '{"pwd_changed": true}'::jsonb;
  end if;
  return new;
end;
$$;

drop trigger if exists on_password_changed on auth.users;
create trigger on_password_changed
  before update of encrypted_password on auth.users
  for each row execute function public.mark_password_changed();

-- 3. Droits sur le fichier de données
drop policy if exists "lecture des donnees" on storage.objects;
create policy "lecture des donnees" on storage.objects
  for select to authenticated
  using (
    bucket_id = 'data'
    and (auth.jwt() -> 'app_metadata' ->> 'pwd_changed') = 'true'
  );

drop policy if exists "ecriture des donnees (admin)" on storage.objects;
create policy "ecriture des donnees (admin)" on storage.objects
  for all to authenticated
  using (bucket_id = 'data' and (auth.jwt() -> 'app_metadata' ->> 'role') = 'admin')
  with check (bucket_id = 'data' and (auth.jwt() -> 'app_metadata' ->> 'role') = 'admin');

-- 4. VOTRE compte administrateur (remplacer l'e-mail, après avoir créé le compte dans
--    Authentication → Users). Il pourra publier les données depuis l'appli.
--    À lancer séparément, en remplaçant l'adresse :
--
-- update auth.users
--    set raw_app_meta_data = coalesce(raw_app_meta_data, '{}'::jsonb)
--                            || '{"role": "admin", "pwd_changed": true}'::jsonb
--  where email = 'votre.adresse@exemple.fr';
--
-- Pour FORCER de nouveau le changement de mot de passe d'un utilisateur :
--
-- update auth.users
--    set raw_app_meta_data = raw_app_meta_data - 'pwd_changed'
--  where email = 'personne@exemple.fr';

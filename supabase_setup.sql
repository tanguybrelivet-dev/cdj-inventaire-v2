-- À exécuter dans Supabase → SQL Editor → New query, puis "Run".

create table if not exists dossiers (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  name text,
  address text,
  status text default 'en_cours',
  total numeric default 0,
  lot_count integer default 0,
  payload jsonb
);

-- Active la sécurité au niveau des lignes (obligatoire dès que la clé "anon"
-- est utilisée côté client, ce qui est le cas ici).
alter table dossiers enable row level security;

-- Politique simple pour démarrer : tout le monde disposant du lien de l'app
-- peut lire/écrire les dossiers. Suffisant pour un usage interne à l'étude
-- où l'app n'est pas rendue publique. Pour restreindre davantage plus tard
-- (comptes utilisateurs, étude par étude...), remplacez ces politiques par
-- des règles basées sur `auth.uid()`.
create policy "Lecture publique" on dossiers
  for select using (true);

create policy "Ecriture publique" on dossiers
  for insert with check (true);

create policy "Mise a jour publique" on dossiers
  for update using (true);

create policy "Suppression publique" on dossiers
  for delete using (true);

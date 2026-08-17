-- Run this once in your Supabase project's SQL Editor.
-- It creates four tables (one per kind of data the app saves) and locks
-- every row to the user who owns it, using Supabase's built-in auth.uid().

create table if not exists custom_bushou (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  char text not null,
  pinyin text,
  meaning text,
  sv text,
  strokes int,
  created_at timestamptz default now(),
  unique (user_id, char)
);

create table if not exists custom_characters (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  char text not null,
  pinyin text,
  meaning text,
  sv text,
  components jsonb default '[]'::jsonb,
  lists jsonb default '[]'::jsonb,
  created_at timestamptz default now(),
  unique (user_id, char)
);

create table if not exists deleted_characters (
  user_id uuid not null references auth.users(id) on delete cascade,
  char text not null,
  primary key (user_id, char)
);

create table if not exists needs_review (
  user_id uuid not null references auth.users(id) on delete cascade,
  char text not null,
  primary key (user_id, char)
);

-- Turn on Row Level Security for all four tables.
alter table custom_bushou enable row level security;
alter table custom_characters enable row level security;
alter table deleted_characters enable row level security;
alter table needs_review enable row level security;

-- Each policy says: you may only touch rows where user_id = your own auth id.
create policy "own rows only" on custom_bushou
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

create policy "own rows only" on custom_characters
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

create policy "own rows only" on deleted_characters
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

create policy "own rows only" on needs_review
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

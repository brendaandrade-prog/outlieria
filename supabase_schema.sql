create extension if not exists pgcrypto;

create table if not exists public.analyses (
  id uuid primary key default gen_random_uuid(),
  client_name text not null,
  farm_name text not null,
  analysis_year integer not null default 2026,
  radius_m numeric(12,2) not null default 300,
  inserted_at timestamptz not null default now(),
  area_2026_ha numeric(14,2),
  point_count integer,
  outlier_count integer not null default 0,
  attention_count integer not null default 0,
  source_hash text not null unique,
  drive_folder_id text,
  status text not null default 'processed',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.analysis_files (
  id uuid primary key default gen_random_uuid(),
  analysis_id uuid not null references public.analyses(id) on delete cascade,
  file_kind text not null check (file_kind in ('excel_original', 'shp', 'reanálise_2026', 'médias_consolidadas')),
  file_name text not null,
  drive_file_id text not null,
  drive_web_url text,
  created_at timestamptz not null default now(),
  unique (analysis_id, file_kind)
);

alter table public.analyses add column if not exists radius_m numeric(12,2) not null default 300;

create index if not exists analyses_client_farm_idx on public.analyses (client_name, farm_name);
create index if not exists analysis_files_analysis_idx on public.analysis_files (analysis_id);

alter table public.analyses enable row level security;
alter table public.analysis_files enable row level security;

-- A aplicação usa a chave de serviço no backend/Streamlit.
-- Não publique essa chave no GitHub nem no navegador.

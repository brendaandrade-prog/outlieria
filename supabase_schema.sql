create extension if not exists pgcrypto;

create table if not exists public.analyses (
  id uuid primary key default gen_random_uuid(),
  client_name text not null,
  farm_name text not null,
  analysis_year integer not null default 2026,
  radius_m numeric(12,2) not null default 300,
  lof_min_neighbors integer not null default 3,
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
alter table public.analyses add column if not exists lof_min_neighbors integer not null default 3;
alter table public.analyses alter column lof_min_neighbors set default 3;

create index if not exists analyses_client_farm_idx on public.analyses (client_name, farm_name);
create index if not exists analysis_files_analysis_idx on public.analysis_files (analysis_id);

create table if not exists public.analysis_points (
  id uuid primary key default gen_random_uuid(),
  analysis_id uuid not null references public.analyses(id) on delete cascade,
  point_key text,
  point_id text,
  sample_index integer,
  source_row integer,
  year integer,
  crop text,
  depth text,
  longitude numeric(12,8),
  latitude numeric(12,8),
  municipality text,
  state text,
  polygon_index integer,
  created_at timestamptz not null default now()
);

create table if not exists public.analysis_results (
  id uuid primary key default gen_random_uuid(),
  analysis_id uuid not null references public.analyses(id) on delete cascade,
  point_id_ref uuid references public.analysis_points(id) on delete cascade,
  element text not null,
  value numeric(18,6),
  historical_mean numeric(18,6),
  historical_median numeric(18,6),
  q1 numeric(18,6),
  q3 numeric(18,6),
  iqr numeric(18,6),
  lower_fence numeric(18,6),
  upper_fence numeric(18,6),
  statistical_status text,
  direction text,
  ml_status text,
  ml_score numeric(18,8),
  ml_neighbor_count integer,
  ml_k_effective integer,
  ml_radius_m numeric(12,2),
  ml_rule text,
  ml_mean_distance_m numeric(18,4),
  ml_regional_median numeric(18,6),
  ml_reference_scope text,
  ml_historical_percentile numeric(18,8),
  ml_historical_distribution_shift numeric(18,8),
  ml_current_cohort_shift numeric(18,8),
  reg_status text,
  reg_prediction numeric(18,6),
  reg_residual numeric(18,6),
  reg_score numeric(18,8),
  reg_cycles integer,
  reg_rule text,
  ph_smp_status text,
  display_status text,
  sample_index integer,
  final_status text,
  created_at timestamptz not null default now(),
  unique (analysis_id, point_id_ref, element)
);

-- Migrações aditivas para bases que já existiam antes da persistência detalhada.
alter table public.analysis_points add column if not exists point_key text;
alter table public.analysis_points add column if not exists sample_index integer;
alter table public.analysis_results add column if not exists historical_median numeric(18,6);
alter table public.analysis_results add column if not exists ml_regional_median numeric(18,6);
alter table public.analysis_results add column if not exists ml_reference_scope text;
alter table public.analysis_results add column if not exists ml_historical_percentile numeric(18,8);
alter table public.analysis_results add column if not exists ml_historical_distribution_shift numeric(18,8);
alter table public.analysis_results add column if not exists ml_current_cohort_shift numeric(18,8);
alter table public.analysis_results add column if not exists reg_status text;
alter table public.analysis_results add column if not exists reg_prediction numeric(18,6);
alter table public.analysis_results add column if not exists reg_residual numeric(18,6);
alter table public.analysis_results add column if not exists reg_score numeric(18,8);
alter table public.analysis_results add column if not exists reg_cycles integer;
alter table public.analysis_results add column if not exists reg_rule text;
alter table public.analysis_results add column if not exists ph_smp_status text;
alter table public.analysis_results add column if not exists display_status text;
alter table public.analysis_results add column if not exists sample_index integer;

create index if not exists analysis_points_analysis_idx on public.analysis_points (analysis_id);
create index if not exists analysis_points_location_idx on public.analysis_points (municipality, state);
create index if not exists analysis_points_key_idx on public.analysis_points (analysis_id, point_key);
create index if not exists analysis_results_analysis_idx on public.analysis_results (analysis_id);
create index if not exists analysis_results_element_idx on public.analysis_results (element, final_status);

alter table public.analyses enable row level security;
alter table public.analysis_files enable row level security;
alter table public.analysis_points enable row level security;
alter table public.analysis_results enable row level security;

-- A aplicação usa a chave de serviço no backend/Streamlit.
-- Não publique essa chave no GitHub nem no navegador.

create extension if not exists pgcrypto;

create table if not exists public.utilizacion_il (
  id uuid not null default gen_random_uuid(),
  session_id uuid not null,
  usuario_login text null,
  nombre_profesional text null,
  programa text null,
  login_at timestamptz not null,
  app_closed_at timestamptz null,
  created_at timestamptz not null default now(),
  constraint utilizacion_il_pkey primary key (id),
  constraint utilizacion_il_session_id_key unique (session_id)
);

create index if not exists idx_utilizacion_il_usuario_login
  on public.utilizacion_il (usuario_login);
create index if not exists idx_utilizacion_il_login_at
  on public.utilizacion_il (login_at desc);

create table if not exists public.utilizacion_il_eventos (
  id uuid not null default gen_random_uuid(),
  event_id uuid not null,
  session_id uuid not null,
  usuario_login text null,
  form_id text not null,
  form_name text null,
  opened_at timestamptz not null,
  finished_at timestamptz null,
  created_at timestamptz not null default now(),
  constraint utilizacion_il_eventos_pkey primary key (id),
  constraint utilizacion_il_eventos_event_id_key unique (event_id),
  constraint utilizacion_il_eventos_session_fkey
    foreign key (session_id)
    references public.utilizacion_il (session_id)
    on delete cascade
);

create index if not exists idx_utilizacion_il_eventos_session
  on public.utilizacion_il_eventos (session_id);
create index if not exists idx_utilizacion_il_eventos_form_id
  on public.utilizacion_il_eventos (form_id);
create index if not exists idx_utilizacion_il_eventos_opened_at
  on public.utilizacion_il_eventos (opened_at desc);

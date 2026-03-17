-- Tabla de auditoria de formularios finalizados (RECA Inclusion Laboral)
-- Guarda: usuario que finaliza, hora Colombia (UTC-5), formulario y empresa.

create table if not exists public.formatos_finalizados_il (
  registro_id uuid primary key default gen_random_uuid(),
  session_id uuid null,
  usuario_login text not null,
  nombre_usuario text not null,
  nombre_formato text not null,
  nombre_empresa text not null,
  source_item_key text null,
  payload_schema_version integer not null default 1,
  payload_source text not null default 'form_cache',
  payload_raw jsonb null,
  payload_normalized jsonb null,
  payload_generated_at timestamptz null,
  drive_file_id text null,
  finalizado_at_colombia timestamp without time zone not null default timezone('America/Bogota', now()),
  finalizado_at_iso text null,
  created_at timestamp with time zone not null default now(),
  path_formato text null,
  revisado boolean null,
  upload_status text null,
  upload_error text null,
  upload_attempted_at timestamptz null,
  uploaded_at timestamptz null
);

alter table public.formatos_finalizados_il
  add column if not exists drive_file_id text null;

alter table public.formatos_finalizados_il
  add column if not exists source_item_key text null;

alter table public.formatos_finalizados_il
  add column if not exists payload_schema_version integer not null default 1;

alter table public.formatos_finalizados_il
  add column if not exists payload_source text not null default 'form_cache';

alter table public.formatos_finalizados_il
  add column if not exists payload_raw jsonb null;

alter table public.formatos_finalizados_il
  add column if not exists payload_normalized jsonb null;

alter table public.formatos_finalizados_il
  add column if not exists payload_generated_at timestamptz null;

alter table public.formatos_finalizados_il
  add column if not exists path_formato text null;

alter table public.formatos_finalizados_il
  alter column payload_schema_version set default 1;

alter table public.formatos_finalizados_il
  alter column payload_source set default 'form_cache';

alter table public.formatos_finalizados_il
  add column if not exists revisado boolean null;

alter table public.formatos_finalizados_il
  add column if not exists upload_status text null;

alter table public.formatos_finalizados_il
  add column if not exists upload_error text null;

alter table public.formatos_finalizados_il
  add column if not exists upload_attempted_at timestamptz null;

alter table public.formatos_finalizados_il
  add column if not exists uploaded_at timestamptz null;

do $$
begin
  if exists (
    select 1
    from information_schema.columns
    where table_schema = 'public'
      and table_name = 'formatos_finalizados_il'
      and column_name = 'upload_attempted_at'
      and data_type <> 'timestamp with time zone'
  ) then
    alter table public.formatos_finalizados_il
      alter column upload_attempted_at type timestamptz
      using nullif(upload_attempted_at::text, '')::timestamptz;
  end if;
end
$$;

do $$
begin
  if exists (
    select 1
    from information_schema.columns
    where table_schema = 'public'
      and table_name = 'formatos_finalizados_il'
      and column_name = 'uploaded_at'
      and data_type <> 'timestamp with time zone'
  ) then
    alter table public.formatos_finalizados_il
      alter column uploaded_at type timestamptz
      using nullif(uploaded_at::text, '')::timestamptz;
  end if;
end
$$;

create index if not exists idx_formatos_finalizados_il_created_at
  on public.formatos_finalizados_il (created_at desc);

create index if not exists idx_formatos_finalizados_il_usuario
  on public.formatos_finalizados_il (usuario_login);

create index if not exists idx_formatos_finalizados_il_empresa
  on public.formatos_finalizados_il (nombre_empresa);

create unique index if not exists uq_formatos_finalizados_il_source_item_key
  on public.formatos_finalizados_il (source_item_key)
  where source_item_key is not null;

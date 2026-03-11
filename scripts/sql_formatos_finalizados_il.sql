-- Tabla de auditoria de formularios finalizados (RECA Inclusion Laboral)
-- Guarda: usuario que finaliza, hora Colombia (UTC-5), formulario y empresa.

create table if not exists public.formatos_finalizados_il (
  registro_id uuid primary key default gen_random_uuid(),
  session_id uuid null,
  usuario_login text not null,
  nombre_usuario text not null,
  nombre_formato text not null,
  nombre_empresa text not null,
  drive_file_id text null,
  finalizado_at_colombia timestamp without time zone not null default timezone('America/Bogota', now()),
  finalizado_at_iso text null,
  created_at timestamp with time zone not null default now(),
  path_formato text null,
  revisado boolean null
);

alter table public.formatos_finalizados_il
  add column if not exists drive_file_id text null;

alter table public.formatos_finalizados_il
  add column if not exists path_formato text null;

alter table public.formatos_finalizados_il
  add column if not exists revisado boolean null;

create index if not exists idx_formatos_finalizados_il_created_at
  on public.formatos_finalizados_il (created_at desc);

create index if not exists idx_formatos_finalizados_il_usuario
  on public.formatos_finalizados_il (usuario_login);

create index if not exists idx_formatos_finalizados_il_empresa
  on public.formatos_finalizados_il (nombre_empresa);

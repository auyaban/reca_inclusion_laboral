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

create unique index if not exists uq_formatos_finalizados_il_source_item_key
  on public.formatos_finalizados_il (source_item_key)
  where source_item_key is not null;

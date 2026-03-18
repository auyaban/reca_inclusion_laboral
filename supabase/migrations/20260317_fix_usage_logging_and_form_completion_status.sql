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

update public.formatos_finalizados_il
set upload_status = case
      when coalesce(trim(drive_file_id), '') <> '' or coalesce(trim(path_formato), '') <> '' then 'synced'
      else 'pending'
    end
where coalesce(trim(upload_status), '') = '';

update public.formatos_finalizados_il
set upload_attempted_at = coalesce(
      upload_attempted_at,
      finalizado_at_colombia at time zone 'America/Bogota'
    )
where coalesce(trim(upload_status), '') in ('synced', 'failed');

update public.formatos_finalizados_il
set uploaded_at = coalesce(
      uploaded_at,
      finalizado_at_colombia at time zone 'America/Bogota'
    )
where coalesce(trim(upload_status), '') = 'synced';

update public.utilizacion_il_eventos e
set finished_at = greatest(e.opened_at, u.app_closed_at)
from public.utilizacion_il u
where e.session_id = u.session_id
  and e.finished_at is null
  and u.app_closed_at is not null;

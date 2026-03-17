alter table public.usuarios_reca
  add column if not exists empresa_nit text,
  add column if not exists empresa_nombre text;

comment on column public.usuarios_reca.empresa_nit is
  'Ultimo NIT de empresa asociado a la cedula desde el proceso de contratacion.';

comment on column public.usuarios_reca.empresa_nombre is
  'Ultimo nombre de empresa asociado a la cedula desde el proceso de contratacion.';

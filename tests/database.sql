-- Run in an isolated PostgreSQL database after the migration.
do $$
begin
  if (select rolsuper from pg_roles where rolname = 'clearsend_migrator') then
    raise exception 'Migration regression must use a non-superuser';
  end if;
  if (select pg_get_userbyid(proowner) from pg_proc
      where oid = 'public.increment_usage_counter(text)'::regprocedure) <> 'clearsend_counter_writer' then
    raise exception 'Counter function must have a restricted owner';
  end if;
  if has_schema_privilege('clearsend_counter_writer', 'public', 'CREATE') or
     pg_has_role('clearsend_migrator', 'clearsend_counter_writer', 'SET') or
     pg_has_role('clearsend_migrator', 'clearsend_counter_writer', 'USAGE') then
    raise exception 'Temporary ownership-transfer privileges must be revoked';
  end if;
  if not (select relrowsecurity from pg_class where oid = 'public.usage_counters'::regclass) then
    raise exception 'Row level security must remain enabled';
  end if;
  if has_table_privilege('anon', 'public.usage_counters', 'SELECT') or
     has_table_privilege('authenticated', 'public.usage_counters', 'INSERT') or
     has_function_privilege('anon', 'public.increment_usage_counter(text)', 'EXECUTE') or
     has_function_privilege('authenticated', 'public.increment_usage_counter(text)', 'EXECUTE') then
    raise exception 'Public access must be denied';
  end if;
end $$;
insert into public.usage_counters values (current_date - 100, 'process_click', 10);
set role service_role;
select public.increment_usage_counter('process_click');
select public.increment_usage_counter('process_click');
reset role;
do $$
begin
  if (select count from public.usage_counters where event='process_click' and day=(now() at time zone 'UTC')::date) <> 2 then
    raise exception 'Counter lost events';
  end if;
  if exists(select 1 from public.usage_counters where day < current_date - 89) then raise exception 'Retention failed'; end if;
  begin
    perform public.increment_usage_counter('private@example.com');
    raise exception 'Invalid event accepted';
  exception when raise_exception then
    if sqlerrm <> 'Invalid event' then raise; end if;
  end;
end $$;

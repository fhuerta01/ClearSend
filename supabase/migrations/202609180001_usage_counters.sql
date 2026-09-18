begin;
-- Aggregate-only storage: no event rows, recipient fields or visitor identifiers.
create table public.usage_counters (
  day date not null,
  event text not null check (event in (
    'pane_open', 'process_click', 'process_success', 'process_blocked', 'process_error',
    'quick_clean_click', 'quick_clean_success', 'quick_clean_blocked', 'quick_clean_error',
    'undo_click', 'export_click', 'export_invalid_click', 'remove_click', 'copy_click',
    'refresh_click', 'settings_click'
  )),
  count bigint not null default 0 check (count >= 0),
  primary key (day, event)
);
alter table public.usage_counters enable row level security;
revoke all on public.usage_counters from public, anon, authenticated;
-- A narrowly scoped owner for the SECURITY DEFINER function.
create role clearsend_counter_writer nologin noinherit;
grant usage on schema public to clearsend_counter_writer;
grant select, insert, update, delete on public.usage_counters to clearsend_counter_writer;
create policy counter_writer on public.usage_counters for all to clearsend_counter_writer
  using (true) with check (true);
create function public.increment_usage_counter(event_name text)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  today date := (now() at time zone 'UTC')::date;
begin
  if event_name is null or event_name not in (
    'pane_open', 'process_click', 'process_success', 'process_blocked', 'process_error',
    'quick_clean_click', 'quick_clean_success', 'quick_clean_blocked', 'quick_clean_error',
    'undo_click', 'export_click', 'export_invalid_click', 'remove_click', 'copy_click',
    'refresh_click', 'settings_click'
  ) then raise exception 'Invalid event'; end if;
  insert into public.usage_counters(day, event, count) values (today, event_name, 1)
    on conflict (day, event) do update set count = public.usage_counters.count + 1;
  -- Keep at most 90 UTC calendar days when events are received.
  delete from public.usage_counters where day < today - 89;
end;
$$;
-- Apply function permissions while the migration administrator still owns it.
revoke all on function public.increment_usage_counter(text) from public, anon, authenticated;
grant execute on function public.increment_usage_counter(text) to service_role;
-- PostgreSQL 16+ grants role creators ADMIN but not SET automatically.
-- Temporarily allow the ownership transfer, including in Supabase's SQL editor.
grant clearsend_counter_writer to current_user with inherit false, set true;
grant create on schema public to clearsend_counter_writer;
alter function public.increment_usage_counter(text) owner to clearsend_counter_writer;
revoke create on schema public from clearsend_counter_writer;
revoke set option for clearsend_counter_writer from current_user;
commit;

-- Disposable test databases only. Supabase already supplies its API roles.
create role anon;
create role authenticated;
create role service_role;
-- Match the relevant Supabase restriction: its SQL editor administrator is not
-- a PostgreSQL superuser. A superuser would hide ownership-transfer failures.
create role clearsend_migrator login createrole inherit nosuperuser password 'local-test';
grant usage, create on schema public to clearsend_migrator with grant option;

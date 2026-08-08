-- Run this once in your Supabase project's SQL editor (Project > SQL Editor > New query).

create table if not exists transactions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  type text not null check (type in ('expense', 'income', 'saving')),
  amount numeric not null check (amount > 0),
  category text not null,
  note text,
  date timestamptz not null default now(),
  created_at timestamptz not null default now()
);

alter table transactions enable row level security;

create policy "Users can view their own transactions"
  on transactions for select
  using (auth.uid() = user_id);

create policy "Users can insert their own transactions"
  on transactions for insert
  with check (auth.uid() = user_id);

create policy "Users can delete their own transactions"
  on transactions for delete
  using (auth.uid() = user_id);

create index if not exists transactions_user_id_idx on transactions (user_id);

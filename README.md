# Spendly

A free money manager with real accounts — sign in and your transactions follow you
to any device. Backed by Supabase (free tier: database + auth, no server to run).

## 1. Create your free Supabase project

1. Go to https://supabase.com, sign up (free), and create a new project.
   Pick any name/region/password — the password there is for the database itself,
   not something you'll use day-to-day.
2. Once it's created, go to **Project Settings → API**. You'll need two values:
   - **Project URL**
   - **anon public** key

## 2. Create the transactions table

1. In your Supabase project, open **SQL Editor → New query**.
2. Paste in the contents of `sql/schema.sql` (included in this project) and run it.
   This creates the table and locks it down so each person can only ever see
   their own transactions.

## 3. Connect the app to your project

1. Copy `.env.example` to a new file named `.env`.
2. Fill in the two values from step 1:
   ```
   VITE_SUPABASE_URL=https://your-project-ref.supabase.co
   VITE_SUPABASE_ANON_KEY=your-anon-public-key
   ```
   (This file is gitignored — don't worry about it leaking if you push to GitHub.)

## 4. Turn off "confirm email" if you want instant sign-up (optional)

By default Supabase makes new users click a confirmation link in an email before
they can sign in. For a personal project that's usually unnecessary friction:
**Authentication → Providers → Email → turn off "Confirm email"**.

## 5. Run it locally

You'll need Node.js (https://nodejs.org, LTS version). Then:

```
npm install
npm run dev
```

Open the URL it prints, create an account, and start adding transactions.

## 6. Deploy it for free (Vercel)

1. Push this project to a GitHub repo.
2. Go to https://vercel.com → sign in with GitHub → "New Project" → pick the repo.
3. Before deploying, add your two env vars under **Environment Variables**:
   `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` (same values as your `.env`).
4. Deploy. You'll get a live URL — sign in from your phone and laptop and your
   transactions will be the same on both.

## Notes

- Every transaction is tagged with your account and Supabase's row-level security
  makes sure no one can see another user's data, even if they guessed an ID.
- Supabase's free tier is generous for personal use (500MB database, no time limit,
  pauses after a week of inactivity but wakes back up when you visit).
- PDF/photo import still runs entirely in your browser — those files are never
  uploaded anywhere.

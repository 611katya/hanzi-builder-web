# Deploying Hanzi Builder as a real website

You'll do five things: create a Supabase project, run one SQL script, get your
project onto GitHub, deploy it on Vercel, and add your two secret keys. About
20–30 minutes total, no prior experience needed.

## 1. Create your Supabase project (database + login system)

1. Go to https://supabase.com and sign up (free).
2. Click **New project**. Pick any name, set a database password (save it
   somewhere), pick a region close to you, click **Create new project**.
   Wait ~2 minutes while it provisions.

## 2. Set up the database tables

1. In your new project, open the **SQL Editor** in the left sidebar.
2. Click **New query**.
3. Open `schema.sql` (included in this project folder), copy its entire
   contents, paste into the SQL editor.
4. Click **Run**. You should see "Success. No rows returned."

This created four tables (`custom_bushou`, `custom_characters`,
`deleted_characters`, `needs_review`) and locked every row so a user can only
ever see or edit their own — enforced by the database itself, not just the
app's code.

## 3. Turn on email login

1. In the sidebar, go to **Authentication → Providers**.
2. **Email** is on by default — that's all you need for sign-up/log-in with
   email + password.
3. Optional: **Authentication → URL Configuration** → set your future site
   URL once you have it (step 5) so confirmation emails link back correctly.
   You can skip this for now and come back to it.

## 4. Get your project's API keys

1. Go to **Settings → API**.
2. Copy the **Project URL** (looks like `https://xxxxx.supabase.co`).
3. Copy the **anon public** key (a long string). Do NOT use the
   `service_role` key — that one must never go in a frontend app.

Keep these two values handy for step 6.

## 5. Put the project on GitHub

1. Go to https://github.com, sign up if needed, click **New repository**,
   name it `hanzi-builder-web`, keep it private or public, **Create
   repository**.
2. On your computer, unzip the project folder I gave you, then in a terminal
   inside that folder run:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/hanzi-builder-web.git
   git push -u origin main
   ```
   (No `git`/terminal experience? GitHub Desktop app does the same thing with
   buttons instead of commands — download it from desktop.github.com.)

## 6. Deploy on Vercel

1. Go to https://vercel.com, sign up using your GitHub account (one click).
2. Click **Add New → Project**, pick the `hanzi-builder-web` repo you just
   pushed.
3. Vercel auto-detects it's a Vite project — leave the build settings as
   default.
4. Before clicking Deploy, expand **Environment Variables** and add:
   - `VITE_SUPABASE_URL` → paste your Project URL from step 4
   - `VITE_SUPABASE_ANON_KEY` → paste your anon public key from step 4
5. Click **Deploy**. Wait ~1 minute.

You'll get a live URL like `https://hanzi-builder-web.vercel.app`. That's
your real, public website.

## 7. Try it

1. Open the URL, click "Sign up", enter an email + password.
2. Check that email for a confirmation link (check spam folder too), click
   it.
3. Log in. Add a few characters. Log out.
4. Open the same URL on your phone, log in with the same email/password —
   your data is there.

## Updating the site later

Any time you want to change the app: edit the code, then
```
git add .
git commit -m "describe your change"
git push
```
Vercel automatically redeploys within about a minute of every push.

## If something goes wrong

- **"Missing VITE_SUPABASE_URL" in the browser console** → the environment
  variables weren't set, or were set after the last deploy. Fix them in
  Vercel → Project → Settings → Environment Variables, then redeploy.
- **Confirmation email never arrives** → check Supabase → Authentication →
  Users to see if the account was created anyway; Supabase's default email
  sender is sometimes slow or lands in spam on the free tier.
- **"row-level security policy" error when saving** → double-check step 2
  ran successfully (all four tables should show a lock icon under
  Database → Tables in Supabase, meaning RLS is on).

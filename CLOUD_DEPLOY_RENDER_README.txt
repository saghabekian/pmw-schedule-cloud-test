# PMW Cloud Test v2 Startup Fix - Render Free Test

This package is built from the working PMW v25 Auto Save / Mobile version.

## What this version does

- Runs on Render as a public web app.
- Works from iPhone, PC, cellular, home, and job sites.
- Keeps login system and permissions.
- Keeps mobile layout and auto-save.
- Uses the same local SQLite database style for the first free test.

## Important limitation for the free test

This first cloud test is meant to prove access and workflow with other users.

If you run it without a cloud PostgreSQL database, saved data may reset if the free cloud server restarts or redeploys. That is normal for a quick free test.

For real production use, the next step is PostgreSQL.

## Render setup

1. Make a free GitHub account if you do not have one.
2. Create a new repository, for example:
   pmw-schedule-cloud-test
3. Upload all files from this folder into that GitHub repository.
4. Go to Render.com.
5. Create a new Web Service.
6. Connect the GitHub repository.
7. Use:
   Build Command:
   pip install -r requirements.txt

   Start Command:
   gunicorn app:app

8. Deploy.
9. Render will give you a public URL like:
   https://pmw-schedule-cloud-test.onrender.com

## First login

admin / admin123
shop / shop123
viewer / view123

## Recommended test

- Open the Render URL on your PC.
- Open the same URL on your iPhone using cellular.
- Login from both.
- Edit a few cells.
- Refresh the other device to confirm the changes show.

## Production next step

After the cloud test works, upgrade the database to PostgreSQL and add automatic backups.


## v2 Fix

This version fixes Render/Gunicorn startup by running the database setup when the app is imported, not only when started locally. This fixes errors like:

sqlite3.OperationalError: no such column: bg_color

To update Render:
1. Upload/replace all files from this v2 folder in GitHub.
2. Commit changes.
3. Render should redeploy automatically.
4. If not, click Manual Deploy > Deploy latest commit.

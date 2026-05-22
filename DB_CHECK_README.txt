PMW Cloud Test v4 DB Check

This version shows a Database Mode banner at the top of the app.

What you want to see:
Database Mode: PostgreSQL

If you see:
Database Mode: SQLite TEMP

Then Render is NOT using your Postgres database yet.

Most common causes:
1. DATABASE_URL was added to the Postgres database service, not the web service.
2. DATABASE_URL key is misspelled.
3. You pasted the wrong value.
4. The web service has not been redeployed after saving the environment variable.

Exact check:
1. Open your Render Web Service: pmw-schedule-cloud-test
2. Click Environment
3. Confirm there is a variable:
   DATABASE_URL
4. Value should start with:
   postgres://
   or
   postgresql://
5. Save Changes
6. Manual Deploy > Deploy latest commit
7. Open the app.
8. Confirm banner says:
   Database Mode: PostgreSQL

You can also open:
/db_check
after logging in as admin.

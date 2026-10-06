# Outlook Directory Agent

Web app with a browser agent (Playwright/Chromium) that signs in to
<https://outlook.office.com/people/> interactively and returns a JSON payload for a chosen user.

## Run
```
npm install
npm start            # http://localhost:3000
```
Set `CHROMIUM_PATH` if Chromium isn't auto-detected (`PLAYWRIGHT_BROWSERS_PATH` is honoured).
`HEADFUL=1` shows a real window when running locally.

## Use
1. Click **Start browser**; the server-side browser is streamed into the page. Click/type to complete
   sign-in (password, MFA, SSO). The badge turns green once you reach Outlook. The session is
   saved to `.session/state.json` (git-ignored) so you only log in once.
2. Search by name/email, click a result to get its JSON.

## API
- `GET /api/search?q=jane` → normalised list of matches
- `GET /api/user?email=jane@corp.com` (or `id=` / `name=`) → JSON for one user; `&raw=1` adds the source objects
- `GET /api/status`, `POST /api/logout`

## How it works / caveats
The agent types into the People search box and captures the JSON responses the Outlook web app itself
fetches, then normalises persona-like objects (`displayName` + email/id). No private APIs are called
directly, so the payload shape follows what Microsoft's UI returns and may change. Selectors for the
search box are heuristic; adjust in `runSearch` if Microsoft changes the page.
`.session/state.json` contains live auth cookies: protect it, and don't expose this server publicly
(it has no auth and gives full control of a signed-in browser). Use only with accounts you're authorised to query.

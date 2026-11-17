# Drishti front-end source

- `App.jsx`: the whole app (React, one file). Edit this.
- `gh-entry.jsx`: starts the app (and not inside the Microsoft sign-in popup).
- `ms-guard.js`, `orig/msal-bridge.js`, `orig/msal-lib.js`: Microsoft sign-in helpers that the build puts inside index.html.
- The page itself is `index.html`, one folder up, with `config.js` and `embed.js` next to it.

Build, in this folder:

    npm ci
    ./build.sh          # writes the new app into ../index.html

`build.sh` replaces only the app and the Microsoft sign-in blocks inside `../index.html`. The rest of the page stays as it
is, including the line that loads `embed.js` (Brahmastra's embedded view). Commit `index.html` together with the change
to `App.jsx`. Building this source unchanged gives back the published `index.html` byte for byte.

Not in this folder: the Google Apps Script backend and the test files of the developer kit.
Never put passwords, keys or real patient or claim data in the repo.

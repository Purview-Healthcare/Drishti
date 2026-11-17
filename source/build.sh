set -e
cd "$(dirname "$0")"
npx esbuild App.jsx --bundle --outfile=/dev/null --external:react --external:react-dom --external:xlsx --loader:.jsx=jsx --log-level=error
npx esbuild gh-entry.jsx --bundle --minify --format=iife --target=es2019 --outfile=gh.js --loader:.jsx=jsx --define:process.env.NODE_ENV='"production"' --legal-comments=none --log-level=error
python3 - <<'PY'
import re
D = ""
html = open("../index.html", encoding="utf-8").read()
js = open(D + "gh.js", encoding="utf-8").read()
assert "</script" not in js.lower()
start = html.index("<script>\n") + len("<script>\n"); end = html.rindex("\n</script>")
html = html[:start] + js + html[end:]
# Microsoft sign-in: the same msal-browser build as before, the pop-up guard, then the library itself
bridge = open(D + "orig/msal-bridge.js", encoding="utf-8").read().strip()
lib = open(D + "orig/msal-lib.js", encoding="utf-8").read().strip()
guard = open(D + "ms-guard.js", encoding="utf-8").read().strip()
for part in (bridge, lib, guard): assert "</script" not in part.lower()
block = ('<script id="msal-bridge">\n' + bridge + '\n</script>\n<script id="ms-guard">\n' + guard + '\n</script>\n'
         '<script id="msal-lib">\n' + lib + '\n</script>\n')
html = re.sub(r'<script id="msal-bridge">.*?</script>\n<script id="ms-guard">.*?</script>\n<script id="msal-lib">.*?</script>\n', '', html, flags=re.S)
assert html.count('<script src="config.js"></script>') == 1
html = html.replace('<script src="config.js"></script>', block + '<script src="config.js"></script>')
open("../index.html", "w", encoding="utf-8").write(html)
print("index.html %.2f MB" % (len(html.encode()) / 1e6))
PY

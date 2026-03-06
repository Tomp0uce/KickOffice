import os
import re

backend_routes_dir = r"f:\KickAI\KickOffice\backend\src\routes"
server_file = r"f:\KickAI\KickOffice\backend\src\server.js"

# 1. Update server.js to add reqId to res.locals
with open(server_file, 'r', encoding='utf-8') as f:
    server_code = f.read()

reqid_middleware = """
// Add request ID
app.use((req, res, next) => {
  res.locals.reqId = crypto.randomUUID()
  next()
})
"""

if "res.locals.reqId" not in server_code:
    server_code = server_code.replace("app.use(express.json({ limit: '4mb' }))", reqid_middleware + "\napp.use(express.json({ limit: '4mb' }))")
    with open(server_file, 'w', encoding='utf-8') as f:
        f.write(server_code)
    print("Updated server.js")

# 2. Add systemLog to chat.js and image.js, and convert console.error
for filename in os.listdir(backend_routes_dir):
    if not filename.endswith('.js'): continue
    filepath = os.path.join(backend_routes_dir, filename)
    with open(filepath, 'r', encoding='utf-8') as f:
        code = f.read()
    
    modified = False
    
    if "console.error" in code:
        # If it's upload.js, and systemLog is already used, just delete console.error
        if filename == "upload.js":
            code = re.sub(r"^\s*console\.error\([^)]+\)\n", "", code, flags=re.MULTILINE)
            modified = True
        else:
            # Need to make sure systemLog is imported
            if "systemLog" not in code:
                code = code.replace("import { logAndRespond } from '../utils/http.js'", "import { logAndRespond } from '../utils/http.js'\nimport { systemLog } from '../utils/logger.js'")
                modified = True
            
            # replace console.error('...', err) with systemLog('ERROR', '...', err)
            # This is a bit of regex parsing. We'll find `console.error(MSG, ERR)`
            code = re.sub(r"console\.error\(([^,]+),\s*([^)]+)\)", r"systemLog('ERROR', \1, \2)", code)
            code = re.sub(r"console\.error\(([^)]+)\)", r"systemLog('ERROR', \1)", code) # For single arg
            modified = True
            
    if modified:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(code)
        print(f"Updated {filename}")

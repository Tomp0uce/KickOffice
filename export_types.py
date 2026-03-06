import os
import re
import shutil

TYPES_FILE = r"f:\KickAI\KickOffice\frontend\src\types\index.d.ts"
NEW_TYPES_FILE = r"f:\KickAI\KickOffice\frontend\src\types\index.ts"

types = [
    'IStringKeyMap', 'insertTypes', 'ModelTier', 'ModelInfo', 
    'ToolInputSchema', 'ToolProperty', 'ToolCategory', 'ToolDefinition', 
    'WordToolDefinition', 'ExcelToolDefinition', 'PowerPointToolDefinition', 
    'OutlookToolDefinition', 'OfficeHostType'
]

# Rename and export in index.ts
with open(TYPES_FILE, 'r', encoding='utf-8') as f:
    content = f.read()

content = re.sub(r'^(type|interface)\s+', r'export \1 ', content, flags=re.MULTILINE)

with open(NEW_TYPES_FILE, 'w', encoding='utf-8') as f:
    f.write(content)

os.remove(TYPES_FILE)

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Find which types are used in this file
    # Ensure they are standalone words
    found_types = []
    for t in types:
        if re.search(rf'\b{t}\b', content) and not re.search(rf"import .*\b{t}\b.* from '@/types'", content):
            found_types.append(t)
            
    if not found_types:
        return
        
    print(f"Modifying {filepath} with {found_types}")
    
    import_stmt = f"import type {{ {', '.join(found_types)} }} from '@/types'\n"
    
    if filepath.endswith('.vue'):
        content = re.sub(r'(<script[^>]*>\n)', r'\1' + import_stmt, content, count=1)
    elif filepath.endswith('.ts'):
        # insert after first block of comments or at top
        content = import_stmt + content
            
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

src_dir = r"f:\KickAI\KickOffice\frontend\src"
for root, _, files in os.walk(src_dir):
    for file in files:
        if file.endswith('.vue') or file.endswith('.ts'):
            if file == 'index.ts': continue
            filepath = os.path.join(root, file)
            process_file(filepath)

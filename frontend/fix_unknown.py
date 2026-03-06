import os
import re

def fix_unknown(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # replace args: Record<string, unknown> with args: Record<string, any>
    new_content = re.sub(r'args:\s*Record<string,\s*unknown>', r'args: Record<string, any>', content)
    # replace args?: Record<string, unknown> with args?: Record<string, any>
    new_content = re.sub(r'args\?:\s*Record<string,\s*unknown>', r'args?: Record<string, any>', content)
    
    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Fixed unknown to any in {filepath}")

base_dir = r"f:\KickAI\KickOffice\frontend\src\utils"
for file in ['wordTools.ts', 'excelTools.ts', 'powerpointTools.ts', 'outlookTools.ts']:
    fix_unknown(os.path.join(base_dir, file))

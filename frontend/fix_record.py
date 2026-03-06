import os
import re

def fix_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    new_content = content.replace("Record<string, unknown>", "Record<string, any>")
    # also rename insertTypes to InsertType, but be careful of whole words
    new_content = re.sub(r'\binsertTypes\b', 'InsertType', new_content)
    
    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Fixed {os.path.basename(filepath)}")

base_dir = r"f:\KickAI\KickOffice\frontend\src"

for root, _, files in os.walk(base_dir):
    for file in files:
        if file.endswith('.ts') or file.endswith('.vue'):
            fix_file(os.path.join(root, file))

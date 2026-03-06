import os
import re

def insert_unknown_cast(filepath, tool_name, tool_def):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # replace `as Record<ToolName, ToolDefinition>` with `as unknown as Record<ToolName, ToolDefinition>`
    pattern = rf"as\s+Record<{tool_name},\s*{tool_def}>"
    replacement = f"as unknown as Record<{tool_name}, {tool_def}>"
    
    new_content = re.sub(pattern, replacement, content)
    
    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Fixed cast in {filepath}")

base_dir = r"f:\KickAI\KickOffice\frontend\src\utils"
insert_unknown_cast(os.path.join(base_dir, 'wordTools.ts'), 'WordToolName', 'WordToolDefinition')
insert_unknown_cast(os.path.join(base_dir, 'excelTools.ts'), 'ExcelToolName', 'ExcelToolDefinition')
insert_unknown_cast(os.path.join(base_dir, 'powerpointTools.ts'), 'PowerPointToolName', 'PowerPointToolDefinition')
insert_unknown_cast(os.path.join(base_dir, 'outlookTools.ts'), 'OutlookToolName', 'OutlookToolDefinition')

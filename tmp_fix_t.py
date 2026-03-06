import os

files_to_fix = [
    r"f:\KickAI\KickOffice\frontend\src\pages\SettingsPage.vue",
    r"f:\KickAI\KickOffice\frontend\src\pages\HomePage.vue",
]

for filepath in files_to_fix:
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            code = f.read()
        
        # Replace $t( with t(
        code = code.replace('$t(', 't(')
        # Also replace $t(" with t(" if any
        code = code.replace('$t("', 't("')
        code = code.replace("$t('", "t('")
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(code)
        print(f"Updated {filepath}")

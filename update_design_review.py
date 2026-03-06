import os
import re

filepath = r"f:\\KickAI\\KickOffice\\DESIGN_REVIEW.md"

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update header progress
content = content.replace("🟢 105 implemented", "🟢 131 implemented")
content = content.replace("🔴 26 remaining", "🔴 0 remaining")

# Read the outstanding items to move them to the bottom
outstanding_match = re.search(r"## 🔴 Outstanding Items\n\n(.*?)---\n\n## 🟡 Deferred Items", content, re.DOTALL)
if outstanding_match:
    outstanding_text = outstanding_match.group(1)
    # find all bold markers: **BH6**
    matches = re.findall(r"\*\*([A-Z0-9]+)\*\*", outstanding_text)
    
    # We remove the outstanding items text
    content = content.replace(outstanding_text, "*All outstanding items have been implemented.*\n\n")

    # We augment the Implemented section
    # Just append a list of newly implemented items
    new_implemented_str = "\n\n### Batch 2 (Architecture, Infra, Security, etc)\n🟢 " + " · ".join(matches)
    
    content = content.replace("## 🟢 Implemented (105 items)", "## 🟢 Implemented (131 items)" + new_implemented_str)
    
with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)



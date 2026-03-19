import re

path = "C:/Users/aaron/Desktop/RECA_INCLUSION_LABORAL/app.py"
with open(path, encoding='utf-8-sig') as f:
    content = f.read()

original = content

# Pattern: remove the askyesno block + the "if not resume" block in all _maybe_resume_form methods.
# The pattern is always:
#   resume = messagebox.askyesno(\n        "Reanudar",\n        "...",\n    )\n
#   if not resume:\n        {module}.clear_cache_file()\n        {module}.clear_form_cache()\n        return False\n
# The lines after "load_cache_from_file()" are kept as-is.

pattern = re.compile(
    r'        resume = messagebox\.askyesno\(\n'
    r'            "Reanudar",\n'
    r'            "[^"]+",\n'
    r'        \)\n'
    r'        if not resume:\n'
    r'            \S+\.clear_cache_file\(\)\n'
    r'            \S+\.clear_form_cache\(\)\n'
    r'            return False\n'
)

matches = pattern.findall(content)
print(f"Found {len(matches)} dialog blocks to remove")

content = pattern.sub('', content)

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"Done. Removed {len(original) - len(content)} chars")

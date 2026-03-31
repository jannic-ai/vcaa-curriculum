"""Execute all file renames via git mv."""
import os, sys, subprocess

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Read the rename commands from the table file
renames = []
with open('_rename_table.txt', 'r', encoding='utf-8') as f:
    in_commands = False
    for line in f:
        line = line.strip()
        if line == '=== RENAME COMMANDS ===':
            in_commands = True
            continue
        if in_commands and line.startswith('git mv '):
            parts = line.split('"')
            if len(parts) >= 4:
                old_path = parts[1].replace("\\", "/")
                new_path = parts[3].replace("\\", "/")
                renames.append((old_path, new_path))

print(f"Found {len(renames)} renames to execute")

success = 0
errors = 0
for old_path, new_path in renames:
    if not os.path.exists(old_path):
        print(f"SKIP (not found): {old_path}")
        errors += 1
        continue

    result = subprocess.run(
        ['git', 'mv', old_path, new_path],
        capture_output=True, text=True, encoding='utf-8'
    )
    if result.returncode == 0:
        success += 1
    else:
        print(f"ERROR: {old_path} -> {new_path}: {result.stderr.strip()}")
        errors += 1

print(f"\nDone: {success} renamed, {errors} errors")

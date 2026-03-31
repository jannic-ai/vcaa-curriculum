"""Move all VCE CSVs into CSV/ folder per subject."""
import os, sys, subprocess

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

total_moved = 0
total_errors = 0

for root, dirs, files in sorted(os.walk(".")):
    for f in sorted(files):
        if not f.endswith(".csv"):
            continue

        rel_dir = os.path.relpath(root, ".").replace("\\", "/")

        # Find subject folder (parent of Curriculum/Assessment/Glossary/Rubric)
        parts = rel_dir.split("/")
        subject_folder = None
        for i, p in enumerate(parts):
            if p in ("Curriculum", "Assessment", "Glossary", "Rubric"):
                subject_folder = "/".join(parts[:i])
                break

        if not subject_folder:
            subject_folder = rel_dir

        # Skip if already in CSV/
        if "/CSV" in rel_dir or rel_dir.endswith("/CSV"):
            continue

        csv_dir = os.path.join(subject_folder, "CSV")
        os.makedirs(csv_dir, exist_ok=True)

        old_path = os.path.join(rel_dir, f)
        new_path = os.path.join(csv_dir, f)

        result = subprocess.run(
            ['git', 'mv', old_path, new_path],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode == 0:
            total_moved += 1
        else:
            print(f"ERROR: {old_path} -> {new_path}: {result.stderr.strip()}")
            total_errors += 1

print(f"Done: {total_moved} moved, {total_errors} errors")

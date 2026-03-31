"""Move all CSVs from sub-folders into a single CSV/ folder per subject."""
import os, sys, subprocess

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Subject folders (relative to F-10 Curriculum/)
SUBJECT_FOLDERS = [
    "Design and Technologies/Design Tech",
    "Design and Technologies/Digital Tech",
    "English/EAL",
    "English/English",
    "HASS/Civics and Citizenship",
    "HASS/Economics and Business",
    "HASS/Geography",
    "HASS/History",
    "HPE",
    "Maths",
    "Science",
    "The Arts/Dance",
    "The Arts/Drama",
    "The Arts/Media Arts",
    "The Arts/Music",
    "The Arts/Visual Arts",
    "The Arts/Visual Communication Design",
]

# Source subfolders to consolidate from
SOURCE_SUBFOLDERS = [
    "Curriculum",
    "Achievement Standards",
    "Curriculum Comparison",
    "Glossary",
]

total_moved = 0
total_errors = 0

for subject in SUBJECT_FOLDERS:
    csv_dir = os.path.join(subject, "CSV")
    os.makedirs(csv_dir, exist_ok=True)

    moved_this_subject = 0

    for subfolder in SOURCE_SUBFOLDERS:
        source_dir = os.path.join(subject, subfolder)
        if not os.path.isdir(source_dir):
            continue

        for f in sorted(os.listdir(source_dir)):
            if not f.endswith(".csv"):
                continue

            old_path = os.path.join(source_dir, f)
            new_path = os.path.join(csv_dir, f)

            result = subprocess.run(
                ['git', 'mv', old_path, new_path],
                capture_output=True, text=True, encoding='utf-8'
            )
            if result.returncode == 0:
                total_moved += 1
                moved_this_subject += 1
            else:
                print(f"ERROR: {old_path} -> {new_path}: {result.stderr.strip()}")
                total_errors += 1

    print(f"{subject}: {moved_this_subject} CSVs moved to CSV/")

print(f"\nTotal: {total_moved} moved, {total_errors} errors")

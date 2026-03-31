"""Plan CSV consolidation for VCE — move all CSVs into CSV/ per subject."""
import os, sys, csv as csvmod

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

GITHUB_BASE = "https://github.com/jannic-ai/vcaa-curriculum/blob/main/VCE"

# Find all subject folders (those containing CSVs)
subject_folders = set()
for root, dirs, files in os.walk("."):
    for f in files:
        if f.endswith(".csv"):
            rel = os.path.relpath(root, ".").replace("\\", "/")
            # Walk up to find the subject folder (parent of Curriculum/, Assessment/, Glossary/, etc.)
            parts = rel.split("/")
            # Subject folder is the deepest folder before Curriculum/Assessment/Glossary/Rubric
            for i, p in enumerate(parts):
                if p in ("Curriculum", "Assessment", "Glossary", "Rubric"):
                    subject_folders.add("/".join(parts[:i]))
                    break

rows = []
for root, dirs, files in sorted(os.walk(".")):
    for f in sorted(files):
        if not f.endswith(".csv"):
            continue
        rel_dir = os.path.relpath(root, ".").replace("\\", "/")
        rel_path = f"{rel_dir}/{f}"

        # Find subject folder
        parts = rel_dir.split("/")
        subject_folder = None
        subfolder = None
        for i, p in enumerate(parts):
            if p in ("Curriculum", "Assessment", "Glossary", "Rubric"):
                subject_folder = "/".join(parts[:i])
                subfolder = "/".join(parts[i:])
                break

        if not subject_folder:
            subject_folder = rel_dir
            subfolder = ""

        new_dir = f"{subject_folder}/CSV"
        new_path = f"{new_dir}/{f}"
        github_current = f"{GITHUB_BASE}/{rel_path}".replace(" ", "%20")
        github_new = f"{GITHUB_BASE}/{new_path}".replace(" ", "%20")

        rows.append((subject_folder, subfolder, f, rel_path, new_path, github_current))

# Print table
print(f"| {'Subject Folder':<55} | {'Current Subfolder':<40} | {'Filename':<70} | {'New Location':<80} |")
print(f"|{'-'*56}|{'-'*42}|{'-'*72}|{'-'*82}|")
for subject, subfolder, filename, old, new, gh in rows:
    print(f"| {subject:<55} | {subfolder:<40} | {filename:<70} | {new:<80} |")

print(f"\nTotal: {len(rows)} CSVs to move")
subjects_with_csvs = set(r[0] for r in rows)
print(f"Subject folders: {len(subjects_with_csvs)}")
for s in sorted(subjects_with_csvs):
    count = sum(1 for r in rows if r[0] == s)
    print(f"  {s}: {count} CSVs")

# Write CSV for Excel
outpath = r"C:\Users\parky\OneDrive\Jannic PA\Education PA\Curriculum\vce-csv-consolidation.csv"
with open(outpath, "w", encoding="utf-8-sig", newline="") as out:
    writer = csvmod.writer(out)
    writer.writerow(["Subject Folder", "Current Subfolder", "Filename", "Current Path", "New Path", "GitHub URL (current)"])
    for row in rows:
        writer.writerow(row)
print(f"\nCSV written to: {outpath}")

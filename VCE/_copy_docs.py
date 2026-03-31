"""Copy documentation files from local drive to GitHub repo clone."""
import os, sys, shutil
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

LOCAL_ROOT = Path(r"C:\Users\parky\OneDrive\Jannic PA\Education PA\Curriculum\Australian\Victoria\Victoria - VCAA - Senior Years\VCE")
GITHUB_ROOT = Path(".")  # VCE/ in the repo

# Mapping: local folder name → GitHub folder name (for mismatches)
FOLDER_MAP = {
    "English/English and EALDS": "English/English and EALD",
}

copied = 0
skipped = 0
created_dirs = set()

for local_subj_dir in sorted(LOCAL_ROOT.iterdir()):
    if not local_subj_dir.is_dir():
        continue
    # Walk into subject area → subject folders
    for subj_dir in sorted(local_subj_dir.iterdir()):
        if not subj_dir.is_dir():
            continue
        doc_dir = subj_dir / "documentation"
        if not doc_dir.is_dir():
            continue

        # Build relative path
        rel = f"{local_subj_dir.name}/{subj_dir.name}"

        # Apply folder name mapping
        github_rel = FOLDER_MAP.get(rel, rel)
        github_doc_dir = GITHUB_ROOT / github_rel / "documentation"

        # Collect files
        files = [f for f in doc_dir.iterdir()
                 if f.is_file()
                 and (f.suffix.lower() in ('.docx', '.pdf'))
                 and not f.name.startswith('~$')]

        if not files:
            continue

        # Create documentation folder if needed
        if not github_doc_dir.exists():
            github_doc_dir.mkdir(parents=True, exist_ok=True)
            created_dirs.add(str(github_doc_dir))

        for f in sorted(files):
            target = github_doc_dir / f.name
            if target.exists():
                skipped += 1
                continue
            shutil.copy2(str(f), str(target))
            copied += 1

    # Also handle direct subjects (HPE, Maths, Science — no parent area folder nesting)
    doc_dir = local_subj_dir / "documentation"
    if doc_dir.is_dir():
        rel = local_subj_dir.name
        github_rel = FOLDER_MAP.get(rel, rel)
        github_doc_dir = GITHUB_ROOT / github_rel / "documentation"

        files = [f for f in doc_dir.iterdir()
                 if f.is_file()
                 and (f.suffix.lower() in ('.docx', '.pdf'))
                 and not f.name.startswith('~$')]

        if files:
            if not github_doc_dir.exists():
                github_doc_dir.mkdir(parents=True, exist_ok=True)
                created_dirs.add(str(github_doc_dir))

            for f in sorted(files):
                target = github_doc_dir / f.name
                if target.exists():
                    skipped += 1
                    continue
                shutil.copy2(str(f), str(target))
                copied += 1

print(f"Copied: {copied}")
print(f"Skipped (already exists): {skipped}")
print(f"New directories created: {len(created_dirs)}")
for d in sorted(created_dirs):
    print(f"  {d}")

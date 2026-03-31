"""Generate rename mapping for all F-10 Curriculum files."""
import re, os, sys

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

GITHUB_BASE = "https://github.com/jannic-ai/vcaa-curriculum/blob/main/F-10 Curriculum"

# CSV subject prefixes in original filenames
CSV_PREFIXES = {
    "Vic Design Tech V2": "design-tech",
    "Vic Digital Tech V2": "digital-tech",
    "Vic English V2": "english",
    "Vic EAL V2": "eal",
    "Vic CandC V2": "civics",
    "Vic EandB V2": "economics-business",
    "Vic Geography V2": "geography",
    "Vic History V2": "history",
    "Vic HPE V2": "hpe",
    "Vic Maths V2": "maths",
    "Vic Science V2": "science",
    "Vic Dance V2": "dance",
    "Vic Drama V2": "drama",
    "Vic Media Arts V2": "media-arts",
    "Vic Music V2": "music",
    "Vic Visual Arts V2": "visual-arts",
    "Vic VCD V2": "vcd",
    "Vic VCDV2": "vcd",
}

# Subject slug from folder path — use path fragments that uniquely identify each subject
# Order: longest/most-specific first to avoid partial matches
FOLDER_SLUGS = [
    ("Design and Technologies/Design Tech", "design-tech"),
    ("Design and Technologies/Digital Tech", "digital-tech"),
    ("English/English", "english"),
    ("English/EAL", "eal"),
    ("HASS/Civics and Citizenship", "civics"),
    ("HASS/Economics and Business", "economics-business"),
    ("HASS/Geography", "geography"),
    ("HASS/History", "history"),
    ("The Arts/Dance", "dance"),
    ("The Arts/Drama", "drama"),
    ("The Arts/Media Arts", "media-arts"),
    ("The Arts/Music", "music"),
    ("The Arts/Visual Arts", "visual-arts"),
    ("The Arts/Visual Communication Design", "vcd"),
    ("HPE", "hpe"),
    ("Maths", "maths"),
    ("Science", "science"),
    ("Documentation", "vcd"),  # root Documentation folder = VCD
]


def get_slug_from_path(rel_path):
    """Match subject slug from relative path, handling both / and \\ separators."""
    normalized = rel_path.replace("\\", "/")
    for folder_key, slug in FOLDER_SLUGS:
        if folder_key in normalized:
            return slug
    return None


def normalize_band(band):
    """Convert band text to short form."""
    if not band:
        return ""
    band = band.strip().rstrip(".")

    # EAL pathways (standalone like A1, BL, C4)
    if re.match(r'^[ABC][L1234]$', band, re.I):
        return band.lower()
    if re.match(r'^-?\s*AL to C4$', band, re.I):
        return "al-c4"

    # Foundation A to Year 10A
    if re.match(r'^Foundation A to Year 10A$', band, re.I):
        return "fa-y10a"
    # Foundation A to Year 10
    if re.match(r'^Foundation A to Year 10$', band, re.I):
        return "fa-y10"
    # FA to Year 10
    if re.match(r'^FA to Year 10$', band, re.I):
        return "fa-y10"
    # Foundation variants
    if band == "Foundation":
        return "f"
    m = re.match(r'^Foundation ([A-D])$', band)
    if m:
        return "f" + m.group(1).lower()
    # F to Year X
    m = re.match(r'^F to Year (\d+)$', band, re.I)
    if m:
        return f"f-y{m.group(1)}"
    # Foundation to Year X
    m = re.match(r'^Foundation to Year (\d+)$', band, re.I)
    if m:
        return f"f-y{m.group(1)}"
    # Year X
    m = re.match(r'^Year (\d+[aA]?)$', band)
    if m:
        return "y" + m.group(1).lower()
    # Years X and Y
    m = re.match(r'^Years? (\d+) and (\d+)$', band)
    if m:
        return f"y{m.group(1)}-y{m.group(2)}"
    # Year X to Year Y
    m = re.match(r'^Year (\d+) to Year (\d+)$', band)
    if m:
        return f"y{m.group(1)}-y{m.group(2)}"
    # Year X and Y (without "Years")
    m = re.match(r'^Year (\d+) and (\d+)$', band)
    if m:
        return f"y{m.group(1)}-y{m.group(2)}"
    return band.lower().replace(" ", "-")


def normalize_csv_type(raw):
    """Normalize CSV document type to slug form."""
    raw = raw.strip()
    rl = raw.lower()
    # Achievement standards
    if rl in ("as and comparison", "achievement standards and comparison"):
        return "as-comparison"
    if rl in ("as components", "achievement standards components"):
        return "as-components"
    # Curriculum comparison variants
    if rl in ("curriculum comparison", "ccomparison", "c compare", "cc", "ccompare"):
        return "curriculum-comparison"
    if rl == "curriculum":
        return "curriculum"
    if rl == "glossary":
        return "glossary"
    return re.sub(r'\s+', '-', rl)


def rename_csv(filename):
    """Generate new CSV filename from original."""
    name = filename[:-4]  # strip .csv
    slug = None
    remainder = name
    for prefix, s in sorted(CSV_PREFIXES.items(), key=lambda x: -len(x[0])):
        if name.startswith(prefix):
            slug = s
            remainder = name[len(prefix):].strip(" -")
            break
    if not slug:
        return None

    # Remove GraphRAG POC / GraphRAG
    remainder = re.sub(r'\s*-?\s*GraphRAG\s*POC\s*-?\s*', ' ', remainder).strip(" -")
    remainder = re.sub(r'\s*-?\s*GraphRAG\s*-?\s*', ' ', remainder).strip(" -")

    # EAL special: "Curriculum A1" (no dash separator)
    eal_match = re.match(r'^Curriculum\s+([ABC][L1234])$', remainder, re.I)
    if eal_match:
        return f"vcaa-f10-{slug}-v2-curriculum-{eal_match.group(1).lower()}.csv"

    # EAL comparison: "Curriculum Comparison -AL to C4"
    eal_comp = re.match(r'^Curriculum Comparison\s*-?\s*(.+)$', remainder, re.I)
    if eal_comp:
        band = normalize_band(eal_comp.group(1).strip())
        return f"vcaa-f10-{slug}-v2-curriculum-comparison-{band}.csv"

    # AS Comparison with band (no standard separator): "AS Comparison - AL to C4"
    as_comp = re.match(r'^AS Comparison\s*-?\s*(.+)$', remainder, re.I)
    if as_comp:
        band = normalize_band(as_comp.group(1).strip())
        return f"vcaa-f10-{slug}-v2-as-comparison-{band}.csv"

    # AS Components (no band)
    if remainder.strip().lower() == "as components":
        return f"vcaa-f10-{slug}-v2-as-components.csv"

    # Standard: "Type - Band" (split on " - ")
    parts = [p.strip() for p in remainder.split(" - ") if p.strip()]
    doc_type = normalize_csv_type(parts[0]) if parts else "unknown"
    band = ""

    if len(parts) >= 2:
        band = normalize_band(parts[1])
    elif len(parts) == 1:
        # Try to extract band from end of type string (e.g. "AS and Comparison Foundation to Year 10")
        m = re.match(r'^(.+?)\s+(Foundation.*|F to.*|Year.*|Years?.*)$', parts[0], re.I)
        if m:
            doc_type = normalize_csv_type(m.group(1))
            band = normalize_band(m.group(2))

    new_name = f"vcaa-f10-{slug}-v2-{doc_type}"
    if band:
        new_name += f"-{band}"
    return new_name + ".csv"


def rename_doc(filename, slug):
    """Generate new DOCX filename from original."""
    name = filename[:-5]  # strip .docx
    nl = name.lower()

    if 'glossary' in nl:
        return f"vcaa-f10-{slug}-v2-glossary.docx"
    if 'comparison' in nl:
        # EAL pathway-specific comparison
        m = re.search(r'Pathway ([A-C])', name, re.I)
        if m:
            return f"vcaa-f10-{slug}-v2-comparison-pathway-{m.group(1).lower()}.docx"
        # Foundation-specific comparison
        if 'f a to d' in nl or 'levels f' in nl or 'foundation' in nl.split('comparison')[0] if 'comparison' in nl else False:
            return f"vcaa-f10-{slug}-v2-comparison-fa-fd.docx"
        # F to 10 comparison
        if 'f to 10' in nl:
            return f"vcaa-f10-{slug}-v2-comparison-f-y10.docx"
        return f"vcaa-f10-{slug}-v2-comparison.docx"
    if 'curriculum' in nl or 'additional language' in nl:
        return f"vcaa-f10-{slug}-v2-curriculum.docx"
    if 'achievment' in nl or 'achievement' in nl:
        if 'foundation' in nl or 'level 6' in nl:
            return f"vcaa-f10-{slug}-v2-achievements-f-y6.docx"
        if 'level' in nl and '10' in nl:
            return f"vcaa-f10-{slug}-v2-achievements-y7-y10a.docx"
        return f"vcaa-f10-{slug}-v2-achievements.docx"
    if 'leader' in nl or 'guide' in nl:
        return f"vcaa-f10-{slug}-v2-leader-guide.docx"
    if 'transitional' in nl:
        return f"vcaa-f10-{slug}-v2-transitional-advice.docx"
    if 'introducing' in nl:
        return f"vcaa-f10-{slug}-v2-introduction.docx"
    return f"vcaa-f10-{slug}-v2-doc.docx"


# Generate table
rows = []
for root, dirs, files in sorted(os.walk(".")):
    for f in sorted(files):
        if f.startswith("~$") or f.startswith("_"):
            continue
        if not (f.endswith(".csv") or f.endswith(".docx")):
            continue

        filepath = os.path.join(root, f)
        rel = filepath[2:]  # strip ./
        folder = os.path.dirname(rel).replace("\\", "/")
        github_path = f"{GITHUB_BASE}/{folder}/{f}".replace(" ", "%20")

        if f.endswith(".csv"):
            new_name = rename_csv(f)
            ftype = "CSV"
        else:
            slug = get_slug_from_path(rel)
            new_name = rename_doc(f, slug) if slug else None
            ftype = "DOCX"

        rows.append((folder, ftype, f, new_name or "*** UNMAPPED ***", github_path))

# Print table
print(f"| {'Folder':<60} | {'Type':<4} | {'Original':<95} | {'New Name':<60} | GitHub Path")
print(f"|{'-'*61}|{'-'*6}|{'-'*97}|{'-'*62}|{'-'*40}")
for folder, ftype, orig, new, gh in rows:
    flag = " **UNMAPPED**" if "UNMAPPED" in new else ""
    print(f"| {folder:<60} | {ftype:<4} | {orig:<95} | {new:<60} | {gh}{flag}")

print(f"\nTotal: {len(rows)} files")
csv_count = sum(1 for r in rows if r[1] == "CSV")
doc_count = sum(1 for r in rows if r[1] == "DOCX")
unmapped = sum(1 for r in rows if "UNMAPPED" in r[3])
print(f"CSVs: {csv_count}, DOCXs: {doc_count}, Unmapped: {unmapped}")

# Also output as simple old→new for review
print("\n\n=== RENAME COMMANDS ===\n")
for folder, ftype, orig, new, gh in rows:
    if "UNMAPPED" not in new:
        old_path = f"{folder}/{orig}".replace("/", "\\")
        new_path = f"{folder}/{new}".replace("/", "\\")
        print(f'git mv "{old_path}" "{new_path}"')
    else:
        print(f'# UNMAPPED: {folder}/{orig}')

"""
Victorian Senior Years Exam Information Parser
==============================================

Parses VCE English exam-related documents (Assessment Criteria, Exam Reports)
into CSV format for Neo4j database loading.

Required Input Files:
- english-exam-assessment_critera-descriptors-24-27.docx (Assessment Criteria document)
- english-exam-report-2024.docx (Chief Assessor's Exam Report)

CSV Outputs:
Cross-section:
- vcaa_vce_sd_english_exam_report_overview.csv (replaces _section.csv)

Section A detail (6 files):
- vcaa_vce_sd_english_exam_report_topics_a.csv
- vcaa_vce_sd_english_exam_report_invitations_a.csv
- vcaa_vce_sd_english_exam_report_verbs_a.csv
- vcaa_vce_sd_english_exam_report_implications_a.csv
- vcaa_vce_sd_english_exam_report_interrelated_skills_a.csv
- vcaa_vce_sd_english_exam_report_strategies_a.csv

Section B detail (6 files):
- vcaa_vce_sd_english_exam_report_header_b.csv
- vcaa_vce_sd_english_exam_report_assessment_b.csv
- vcaa_vce_sd_english_exam_report_strategies_b.csv
- vcaa_vce_sd_english_exam_report_responses_b.csv
- vcaa_vce_sd_english_exam_report_examples_b.csv
- vcaa_vce_sd_english_exam_report_annotations_b.csv

Section C detail (5 files):
- vcaa_vce_sd_english_exam_report_header_c.csv
- vcaa_vce_sd_english_exam_report_context_c.csv
- vcaa_vce_sd_english_exam_report_argument_c.csv
- vcaa_vce_sd_english_exam_report_language_c.csv
- vcaa_vce_sd_english_exam_report_strategies_c.csv

Plus (from pre-exam parser):
- vcaa_vce_sd_english_final-exam-expected_qualities.csv (Grade descriptors)

==============================================================================
VERSION CONTROL
==============================================================================
Version | Date           | Changes
--------|----------------|-----------------------------------------------------
2.4     | Feb 05, 2026   | SECTION C COMPLETE:
        |                | - Added 5 Section C detail files: header, context,
        |                |   argument, language, strategies.
        |                | - BulletLevel column on header for skill hierarchy.
        |                | - Argument steps with StepNumber column.
        |                | - Language/Visual distinction via LanguageType column.
        |                | - Limited strategies split across tables 20+21.
        |                | - Added ExampleText column to annotations_b.
2.3     | Feb 05, 2026   | FILENAME SHORTENING:
        |                | - Removed '_section' from CSV filenames to fix
        |                |   filename length issues (e.g. _section_a → _a,
        |                |   _section_b → _b).
2.2     | Feb 05, 2026   | OVERVIEW + SECTION RENAME:
        |                | - Replaced _section.csv with _overview.csv covering
        |                |   Sections A, B, and C as table-of-contents.
        |                | - Fixed UnitASCode per section: A=U3AS1+U4AS1,
        |                |   B=U3AS2, C=U4AS2 (was all sections same).
        |                | - Fixed Section C ExamReportCode: VCEEEARC (was VCEEERC).
        |                | - Renamed all Section A detail files with _a
        |                |   suffix to support section-specific file sets.
        |                | - Overview uses H/RC SectionCode pattern with
        |                |   SectionHeaderDescription and ReportCoverage columns.
2.1     | Feb 05, 2026   | BAND + RENAME:
        |                | - Changed Band from '12' to 'Year 12' for
        |                |   consistency with curriculum CSVs.
        |                | - Renamed parser from exam-information to
        |                |   exam-report to clarify purpose.
2.0     | Feb 04, 2026   | NORMALIZED FILE STRUCTURE:
        |                | - Restructured from 4 combined files to 7 normalized
        |                |   files per section for cleaner Neo4j loading
        |                | - Added UnitASCode column for curriculum linkage
        |                |   (e.g., "VCEEU3AS1,VCEEU4AS1" for Section A)
        |                | - CRITICAL: All text extraction is VERBATIM from
        |                |   source documents - no paraphrasing allowed
        |                | - Preserved line breaks in multi-line content
        |                | - Separate files: section, topics, invitations,
        |                |   verbs, implications, interrelated_skills, strategies
        |                | - Code format: VCEEEARA (Section A), with suffixes
        |                |   for row types (H, T, IN, V, IM, EQ, ER, LR)
1.3     | Jan 19, 2026   | ASSESSMENT INFORMATION TRACKING:
        |                | - Added AssessmentInformationDetails field
        |                | - Added AssessmentYears field
        |                | - Expected Qualities: "Pre Exam Criteria", "2024-2027"
        |                | - Section Framework: "Post Exam Report", "2024"
        |                | - Topic Information: "Post Exam Report", "2024"
        |                | - Strategies: "Post Exam Report", "2024"
1.2     | Jan 19, 2026   | ASSESSMENT TYPE DIFFERENTIATION:
        |                | - Added AssessmentType field to all CSV files
        |                | - Value: "Final Exam" for this parser
1.1     | Jan 19, 2026   | SECTION A COMPLETE:
        |                | - Added Section Framework, Topic Information, Strategies
        |                | - Added Grade column to Expected Qualities
1.0     | Jan 19, 2026   | Initial release - Expected Qualities extraction
==============================================================================

Current Version: 2.4
"""

import re
import csv
from collections import defaultdict
from docx import Document
from pathlib import Path
from typing import List, Dict, Tuple, Optional


# ============================================================================
# CONFIGURATION
# ============================================================================

CONFIG = {
    'subject': 'English',
    'subject_area': 'English',
    'subject_code': 'E',
    'band': 'Year 12',  # VCE Units 3-4 = Year 12
    'assessment_type': 'Final Exam',
    
    # Assessment information details
    'assessment_info_pre_exam': {
        'details': 'Pre Exam Criteria',
        'years': '2024-2027'
    },
    'assessment_info_post_exam': {
        'details': 'Post Exam Report',
        'years': '2024'
    },
    
    # Input files
    'criteria_doc': '/mnt/user-data/uploads/english-exam-assessment_critera-descriptors-24-27.docx',
    'report_doc': '/mnt/user-data/uploads/english-exam-report-2024.docx',
    
    # Output
    'output_dir': '/home/claude/vce_exam_information_output',
    
    # Section configuration with curriculum linkage
    'sections': {
        'A': {
            'name': 'Analytical response to a text',
            'exam_report_code': 'VCEEEARA',
            'unit_as_code': 'VCEEU3AS1,VCEEU4AS1',  # Units 3&4 AoS 1: Reading and responding
            'enhanced_table_index': 5,  # Table index in Word doc
            'limited_table_index': 6,
            'topics_table_index': 1,
            'invitations_table_index': 2,
            'verbs_table_index': 3,
            'implications_table_index': 4
        },
        'B': {
            'name': 'Creating a text',
            'exam_report_code': 'VCEEEARB',
            'unit_as_code': 'VCEEU3AS2',  # Unit 3 AoS 2: Creating texts
            # Paragraph indices in source doc (for non-table content)
            'header_para_indices': [46, 47, 48],  # Additional paras after overview para (45)
            'assessment_intro_para': 50,  # "Scripts were assessed holistically..."
            'assessment_skill_paras': [51, 52, 53],  # Numbered skills
            'enhanced_bullet_paras': [55, 56, 57, 58, 59],
            'limited_bullet_paras': [62, 63, 64, 65, 66],
            'warning_para': 67,  # Appropriation warning
            'responses_paras': [69, 70, 71],  # 3 body paragraphs under Responses heading
            # Framework → Example → Table mapping
            'frameworks': [
                {
                    'number': 1,
                    'name': 'Country',
                    'examples': [
                        {'example_number': 1, 'desc_paras': [74, 75, 76], 'table_index': 8}
                    ]
                },
                {
                    'number': 2,
                    'name': 'Protest',
                    'examples': [
                        {'example_number': 2, 'desc_paras': [81, 82, 83], 'table_index': 9},
                        {'example_number': 3, 'desc_paras': [85, 86, 87], 'table_index': 10},
                        {'example_number': 4, 'desc_paras': [90, 91, 92], 'table_index': 11}
                    ]
                },
                {
                    'number': 3,
                    'name': 'Personal Journeys',
                    'examples': [
                        {'example_number': 5, 'desc_paras': [95, 96, 97], 'table_index': 12},
                        {'example_number': 6, 'desc_paras': [101, 102, 103], 'table_index': 13}
                    ]
                },
                {
                    'number': 4,
                    'name': 'Play',
                    'examples': [
                        {'example_number': 7, 'desc_paras': [106, 107, 108], 'table_index': 14},
                        {'example_number': 8, 'desc_paras': [110, 111, 112], 'table_index': 15},
                        {'example_number': 9, 'desc_paras': [114, 115, 116], 'table_index': 16}
                    ]
                }
            ]
        },
        'C': {
            'name': 'Analysis of argument and language',
            'exam_report_code': 'VCEEEARC',
            'unit_as_code': 'VCEEU4AS2',  # Unit 4 AoS 2: Analysing argument
            # Interrelated skills (header file)
            'assessment_intro_para': 119,  # "Students were invited to 'write an analysis..."
            'skill_bullet_paras': [120, 121, 122, 123, 124, 125],
            # Context/background paragraphs
            'context_paras': [126, 127, 128, 129, 130, 131],
            # Line of argument
            'argument_intro_para': 132,  # "One way of representing the line of argument..."
            'argument_step_paras': [133, 134, 135, 136, 137, 138, 139],
            'argument_transition_para': 140,  # "Students were not expected to identify..."
            # Language choices
            'language_intro_para': 141,  # "Some strategic language choices included..."
            'language_bullet_paras': [142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152],
            'visual_intro_para': 153,  # "Some strategic visual cues..."
            'visual_bullet_paras': [154, 155, 156, 157, 158],
            'language_closing_paras': [159, 160],
            # Strategies
            'enhanced_table_index': 19,  # 7 rows x 3 cols (Skill, Strategy, Example)
            'limited_table_indices': [20, 21],  # Split across 2 tables
            'strategies_closing_para': 167,  # "In Section C, as in other sections..."
        }
    },
    
    # Section code suffixes
    'section_codes': {
        'header': 'H',
        'topics': 'T',
        'invitations': 'IN',
        'verbs': 'V',
        'implications': 'IM',
        'expected_qualities': 'EQ',
        'enhanced': 'ER',
        'limited': 'LR'
    },
    
    # Grade mapping (mark to grade)
    'grade_mapping': {
        0: 'E', 1: 'E', 2: 'E+', 3: 'D', 4: 'D+',
        5: 'C', 6: 'C+', 7: 'B', 8: 'B+', 9: 'A', 10: 'A+'
    },
    
    # LLM-generated verb definitions (capitalised)
    'verb_definitions': {
        'Attempts': 'To make an effort to achieve, demonstrate, or convey something; to try to accomplish a particular interpretation, argument, or effect within the text',
        'Challenges': 'To question, dispute, or test the validity of an idea, value, assumption, or established norm presented in or through the text',
        'Defends': 'To argue in support of or justify a position, value, character, or idea; to present evidence that upholds or protects something from criticism',
        'Demonstrates': 'To show clearly through evidence, examples, or argument; to make evident or prove a point about the text\'s ideas or values',
        'Determines': 'To be the decisive factor in shaping, influencing, or controlling an outcome, character\'s fate, or thematic development',
        'Explores': 'To examine, investigate, or consider in depth; to look at multiple aspects, perspectives, or implications of an idea, theme, or value',
        'Interplays': 'To show how two or more elements interact with, influence, and affect each other reciprocally; to examine the dynamic relationship between concepts',
        'Portrays': 'To represent, depict, or characterise; to present a particular image, interpretation, or understanding of a character, setting, or idea',
        'Re-enforces': 'To strengthen, support, or give additional weight to an idea, value, or theme; to confirm or intensify through repetition or emphasis',
        'Reveals': 'To make known, uncover, or disclose something previously hidden or not immediately apparent about characters, themes, or values',
        'Shapes': 'To influence, form, or determine the nature, development, or outcome of something; to mould understanding or perception',
        'Shows': 'To present, display, or make visible; to demonstrate through the text\'s content, structure, or language how something is represented'
    },
    
    # New verbs identified in 2024 exam (for VerbsIsNew flag)
    'new_verbs_2024': ['Interplays']
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def validate_input_files(config: Dict) -> Tuple[bool, str]:
    """Validate that required input files exist."""
    report_path = Path(config['report_doc'])
    
    missing = []
    if not report_path.exists():
        missing.append(f"Exam Report document: {config['report_doc']}")
    
    if missing:
        msg = "ERROR: Required files are missing.\n\n"
        msg += "Missing files:\n"
        for m in missing:
            msg += f"  - {m}\n"
        return False, msg
    
    return True, "All required files present."


def get_metadata_row(config: Dict, section_key: str) -> List[str]:
    """Return common metadata columns for a section."""
    section = config['sections'][section_key]
    return [
        config['subject_area'],
        config['subject'],
        config['band'],
        config['assessment_type'],
        config['assessment_info_post_exam']['details'],
        config['assessment_info_post_exam']['years'],
        section['unit_as_code'],
        section['exam_report_code']
    ]


def make_section_code(exam_report_code: str, suffix: str) -> str:
    """Generate section code from exam report code and suffix."""
    return f"{exam_report_code}{suffix}"


def capitalise_first(text: str) -> str:
    """Capitalise the first letter of text, preserving the rest."""
    if not text:
        return text
    return text[0].upper() + text[1:]


def split_sentences(text: str) -> List[str]:
    """Split text into sentences, handling common abbreviations.
    
    Splits on '. ' (period + space) and '.' at end of text.
    Preserves abbreviations like 'e.g.', 'i.e.', 'EAL' etc.
    """
    if not text:
        return []
    
    # Simple sentence split: split on period followed by space and uppercase
    # or period at end of text
    sentences = []
    current = []
    
    # Use regex to split on sentence boundaries
    # Match period followed by space and uppercase letter, or period at end
    parts = re.split(r'(?<=[.!?])\s+(?=[A-Z])', text)
    
    for part in parts:
        part = part.strip()
        if part:
            sentences.append(part)
    
    return sentences


# ============================================================================
# SECTION A EXTRACTION FUNCTIONS
# ============================================================================

def extract_overview(doc: Document, config: Dict) -> List[List[str]]:
    """Extract overview (table-of-contents) for all exam report sections.
    
    Produces rows with H suffix (header/description) and RC suffix (report coverage items).
    Each section gets: 1 header row with SectionHeaderDescription, then RC rows for coverage items.
    """
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'SectionHeader', 'SectionHeaderDescription', 'SectionHeaderContent',
        'ReportCoverage', 'Sequence'
    ])
    
    # ── Section A ──────────────────────────────────────────────────────
    meta_a = get_metadata_row(config, 'A')
    code_a = config['sections']['A']['exam_report_code']
    h_code_a = make_section_code(code_a, 'H')
    rc_code_a = make_section_code(code_a, 'RC')
    
    # Section A description (verbatim - single paragraph)
    desc_a = ("In Units 3 and 4, students explored two of the 20 texts on the 2024 VCE "
              "English and English as an Additional Language (EAL) text list. The exam "
              "invited students to write an essay of analysis on one of these two texts. "
              "For each text there was a choice of two topics. These topics invited students "
              "to consider the ideas and/or values that the author presented in relation to "
              "the topic.")
    
    rows.append(meta_a + [h_code_a, f"Section A: {config['sections']['A']['name']}", desc_a, '', '', '1'])
    
    # Section A coverage items
    coverage_a = [
        'Topics',
        'Verbs used in topics',
        'Implications',
        'Expected Qualities',
        'Strategies that enhanced responses',
        'Strategies that limited responses'
    ]
    for seq, item in enumerate(coverage_a, start=2):
        rows.append(meta_a + [rc_code_a, '', '', '', item, str(seq)])
    
    # ── Section B ──────────────────────────────────────────────────────
    meta_b = get_metadata_row(config, 'B')
    code_b = config['sections']['B']['exam_report_code']
    h_code_b = make_section_code(code_b, 'H')
    rc_code_b = make_section_code(code_b, 'RC')
    
    # Section B description (verbatim - first paragraph only)
    desc_b = ("In Units 3 and 4 students explored one of the four Frameworks of Ideas "
              "and associated mentor texts on the 2024 VCE English and English as an "
              "Additional Language (EAL) text list. The exam invited students to create "
              "a text that responded to a nominated title, at least one of the three "
              "pieces of stimulus material provided and explored meaningful ideas "
              "associated with the selected Framework of Ideas.")
    
    rows.append(meta_b + [h_code_b, f"Section B: {config['sections']['B']['name']}", desc_b, '', '', '1'])
    
    # Section B coverage items
    coverage_b = [
        'Assessment',
        'Strategies that enhanced responses',
        'Strategies that limited responses',
        'Responses',
        'Framework of Ideas 1: Country',
        'Framework of Ideas 2: Protest',
        'Framework of Ideas 3: Personal Journeys',
        'Framework of Ideas 4: Play'
    ]
    for seq, item in enumerate(coverage_b, start=2):
        rows.append(meta_b + [rc_code_b, '', '', '', item, str(seq)])
    
    # ── Section C ──────────────────────────────────────────────────────
    meta_c = get_metadata_row(config, 'C')
    code_c = config['sections']['C']['exam_report_code']
    h_code_c = make_section_code(code_c, 'H')
    rc_code_c = make_section_code(code_c, 'RC')
    
    # Section C description (verbatim - first paragraph only)
    desc_c = ("Students were invited to \u2018write an analysis of the ways in which "
              "arguments, written and spoken language and visuals were used \u2026 to "
              "persuade the intended audience to share a point of view.\u2019 As with "
              "the other sections of the exam, this section was assessed holistically "
              "using the assessment criteria and expected qualities. Benchmarks were "
              "used to establish the standard required in each of three interrelated "
              "skills.")
    
    rows.append(meta_c + [h_code_c, f"Section C: {config['sections']['C']['name']}", desc_c, '', '', '1'])
    
    # Section C coverage items
    coverage_c = [
        'Strategies that enhanced responses',
        'Strategies that limited responses'
    ]
    for seq, item in enumerate(coverage_c, start=2):
        rows.append(meta_c + [rc_code_c, '', '', '', item, str(seq)])
    
    return rows


def extract_a_topics(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A topic types from Word document table (VERBATIM)."""
    section = config['sections']['A']
    section_code = make_section_code(section['exam_report_code'], config['section_codes']['topics'])
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'TopicHeader', 'TopicType', 'TopicExplanation', 'TopicExamples', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    # Header row
    rows.append(meta + [section_code, 'Topics', '', '', '', '1'])
    
    # Extract from table (index 1) - VERBATIM
    table = doc.tables[config['sections']['A']['topics_table_index']]
    seq = 2
    for row in table.rows[1:]:  # Skip header row
        topic_type = row.cells[0].text.strip()
        explanation = row.cells[1].text.strip()
        example = row.cells[2].text.strip() if len(row.cells) > 2 else ''
        
        rows.append(meta + [section_code, '', topic_type, explanation, example, str(seq)])
        seq += 1
    
    return rows


def extract_a_invitations(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A topic invitations from Word document table (VERBATIM)."""
    section = config['sections']['A']
    section_code = make_section_code(section['exam_report_code'], config['section_codes']['invitations'])
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'TopicInvitationHeader', 'TopicInvitationDescription', 'TopicInvitationDefinition',
        'TopicInvitationExplanation', 'TopicInvitationExample', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    # Header row
    header_desc = "Another way of cataloguing these topics was to consider the invitation they offered students. This may include, but is not restricted to, topics that:"
    rows.append(meta + [section_code, 'Topic Invitation', header_desc, '', '', '', '1'])
    
    # Extract from table (index 2) - VERBATIM
    table = doc.tables[config['sections']['A']['invitations_table_index']]
    seq = 2
    for row in table.rows[1:]:  # Skip header row
        definition = row.cells[0].text.strip()
        explanation = row.cells[1].text.strip()
        example = row.cells[2].text.strip() if len(row.cells) > 2 else ''
        
        rows.append(meta + [section_code, '', '', definition, explanation, example, str(seq)])
        seq += 1
    
    return rows


def extract_a_verbs(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A verbs with LLM-generated definitions."""
    section = config['sections']['A']
    section_code = make_section_code(section['exam_report_code'], config['section_codes']['verbs'])
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'VerbsHeader', 'VerbsDescription', 'VerbsTerm', 'VerbsDefinition', 'VerbsClosing',
        'VerbsDefinitionSource', 'VerbsIsNew', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    # Header row
    header_desc = """Please note there is no definitive list of verbs used by the examination panel to explore the way in which meaning is conveyed by an author and new terms are regularly introduced.
Students would do well to develop their understanding of such terms. Terms are often nominalised in topics and students should be prepared for this"""
    rows.append(meta + [section_code, 'Verbs used in topics', header_desc, '', '', '', '', '', '1'])
    
    # Verb rows
    seq = 2
    for verb, definition in config['verb_definitions'].items():
        is_new = 'Y' if verb in config['new_verbs_2024'] else 'N'
        rows.append(meta + [section_code, '', '', verb, definition, '', 'LLM-generated', is_new, str(seq)])
        seq += 1
    
    # Closing row
    closing = "Or the use of a linking verb to imply there is a direct connection by use of is/are or is not/are not"
    rows.append(meta + [section_code, '', '', '', '', closing, '', '', str(seq)])
    
    return rows


def extract_a_implications(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A implications from Word document table (VERBATIM)."""
    section = config['sections']['A']
    section_code = make_section_code(section['exam_report_code'], config['section_codes']['implications'])
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'ImplicationsHeader', 'ImplicationsDescription', 'ImplicationsStrategy', 'ImplicationsExample', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    # Header row
    rows.append(meta + [section_code, 'Implications', '', '', '', '1'])
    
    # Description row
    desc = """Students who can demonstrate an understanding of the implications of the topics will typically develop responses that achieve higher scores.
There are often multiple implications in a topic. The following table outlines some strategies used in essay topics in 2024."""
    rows.append(meta + [section_code, '', desc, '', '', '2'])
    
    # Extract from table (index 4) - VERBATIM
    table = doc.tables[config['sections']['A']['implications_table_index']]
    seq = 3
    for row in table.rows[1:]:  # Skip header row
        strategy = row.cells[0].text.strip()
        example = row.cells[1].text.strip() if len(row.cells) > 1 else ''
        
        rows.append(meta + [section_code, '', '', strategy, example, str(seq)])
        seq += 1
    
    return rows


def extract_a_interrelated_skills(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A interrelated skills."""
    section = config['sections']['A']
    section_code = make_section_code(section['exam_report_code'], config['section_codes']['expected_qualities'])
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'EQHeader', 'EQDescription', 'EQSkills', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    rows.append(meta + [section_code, 'Expected Qualities', '', '', '1'])
    rows.append(meta + [section_code, '', 
        'Scripts were assessed holistically using the published EQs. Benchmarks were used to establish the standard required in each of three interrelated skills:',
        '', '2'])
    rows.append(meta + [section_code, '', '', "The capacity to create a reading of the text's ideas/values", '3'])
    rows.append(meta + [section_code, '', '', "The capacity to sequence and substantiate ideas relevant to the topic", '4'])
    rows.append(meta + [section_code, '', '', "The capacity to communicate these ideas", '5'])
    
    return rows


def extract_a_strategies(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section A enhanced and limited strategies (VERBATIM)."""
    section = config['sections']['A']
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'ExamResponsesSkillType', 'ExamResponsesSkillHeader', 'ExamResponsesSkill',
        'ExamResponsesExplanation', 'ExamResponsesExample', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'A')
    
    # === ENHANCED STRATEGIES ===
    enhanced_code = make_section_code(section['exam_report_code'], config['section_codes']['enhanced'])
    rows.append(meta + [enhanced_code, 'EnhancedStrategies', 'Strategies that enhanced responses', '', '', '', '1'])
    
    # Extract from table (index 5) - VERBATIM
    enhanced_table = doc.tables[config['sections']['A']['enhanced_table_index']]
    seq = 2
    for row in enhanced_table.rows[1:]:  # Skip header row
        skill = row.cells[0].text.strip()
        explanation = row.cells[1].text.strip()
        example = row.cells[2].text.strip() if len(row.cells) > 2 else ''
        
        rows.append(meta + [enhanced_code, 'EnhancedStrategies', '', skill, explanation, example, str(seq)])
        seq += 1
    
    # === LIMITED STRATEGIES ===
    limited_code = make_section_code(section['exam_report_code'], config['section_codes']['limited'])
    rows.append(meta + [limited_code, 'LimitedStrategies', 'Strategies that limited responses', '', '', '', '1'])
    
    # Extract from table (index 6) - VERBATIM with row-per-bullet structure
    limited_table = doc.tables[config['sections']['A']['limited_table_index']]
    seq = 2
    for row in limited_table.rows[1:]:  # Skip header row
        skill = row.cells[0].text.strip()
        full_explanation = row.cells[1].text.strip()
        
        # Split into intro and bullet points
        lines = full_explanation.split('\n')
        
        # First line is intro
        intro_line = lines[0].strip() if lines else ''
        rows.append(meta + [limited_code, 'LimitedStrategies', '', skill, intro_line, '', str(seq)])
        seq += 1
        
        # Remaining lines are bullet points - each as separate row
        for line in lines[1:]:
            line = line.strip()
            if line:
                rows.append(meta + [limited_code, 'LimitedStrategies', '', skill, line, '', str(seq)])
                seq += 1
    
    return rows


# ============================================================================
# SECTION B EXTRACTION FUNCTIONS
# ============================================================================

def extract_b_header(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B additional header paragraphs (beyond overview).
    
    The first paragraph (para 45) is captured in the overview CSV.
    This file captures the remaining descriptive paragraphs (paras 46-48).
    Each paragraph is split: first sentence as HeaderText, remaining
    sentences as individual HeaderContent rows (bullet-point style).
    """
    section = config['sections']['B']
    section_code = make_section_code(section['exam_report_code'], 'H')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'HeaderText', 'HeaderContent', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    seq = 1
    for para_idx in section['header_para_indices']:
        text = doc.paragraphs[para_idx].text.strip()
        if not text:
            continue
        
        # Split into sentences
        sentences = split_sentences(text)
        if not sentences:
            continue
        
        # First sentence as header
        rows.append(meta + [section_code, capitalise_first(sentences[0]), '', str(seq)])
        seq += 1
        
        # Remaining sentences as content bullets
        for sentence in sentences[1:]:
            rows.append(meta + [section_code, '', capitalise_first(sentence), str(seq)])
            seq += 1
    
    return rows


def extract_b_assessment(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B assessment criteria (interrelated skills)."""
    section = config['sections']['B']
    section_code = make_section_code(section['exam_report_code'], 'EQ')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'AssessmentHeader', 'AssessmentDescription', 'AssessmentSkill', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    # Header row with intro sentence
    intro = doc.paragraphs[section['assessment_intro_para']].text.strip()
    rows.append(meta + [section_code, 'Assessment', intro, '', '1'])
    
    # Numbered skills
    seq = 2
    for para_idx in section['assessment_skill_paras']:
        skill = doc.paragraphs[para_idx].text.strip()
        if skill:
            rows.append(meta + [section_code, '', '', capitalise_first(skill), str(seq)])
            seq += 1
    
    return rows


def extract_b_strategies(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B enhanced/limited strategies and warning paragraph."""
    section = config['sections']['B']
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'StrategyType', 'StrategyHeader', 'Strategy', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    # === ENHANCED STRATEGIES ===
    enhanced_code = make_section_code(section['exam_report_code'], 'ER')
    rows.append(meta + [enhanced_code, 'EnhancedStrategies', 'Strategies that enhanced responses', '', '1'])
    
    seq = 2
    for para_idx in section['enhanced_bullet_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [enhanced_code, 'EnhancedStrategies', '', text, str(seq)])
            seq += 1
    
    # === LIMITED STRATEGIES ===
    limited_code = make_section_code(section['exam_report_code'], 'LR')
    rows.append(meta + [limited_code, 'LimitedStrategies', 'Strategies that limited responses', '', '1'])
    
    seq = 2
    for para_idx in section['limited_bullet_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [limited_code, 'LimitedStrategies', '', text, str(seq)])
            seq += 1
    
    # === WARNING ===
    warning_code = make_section_code(section['exam_report_code'], 'LW')
    warning_text = doc.paragraphs[section['warning_para']].text.strip()
    
    # Split warning: first sentence as header, remaining as content bullets
    warning_sentences = split_sentences(warning_text)
    if warning_sentences:
        rows.append(meta + [warning_code, 'Warning', capitalise_first(warning_sentences[0]), '', '1'])
        w_seq = 2
        for sentence in warning_sentences[1:]:
            rows.append(meta + [warning_code, 'Warning', '', capitalise_first(sentence), str(w_seq)])
            w_seq += 1
    
    return rows


def extract_b_responses(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B responses intro paragraphs."""
    section = config['sections']['B']
    section_code = make_section_code(section['exam_report_code'], 'RSP')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'ResponsesHeader', 'ResponsesContent', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    # Header row
    rows.append(meta + [section_code, 'Responses', '', '1'])
    
    # Body paragraphs
    seq = 2
    for para_idx in section['responses_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [section_code, '', text, str(seq)])
            seq += 1
    
    return rows


def extract_b_examples(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B framework headers and example descriptions."""
    section = config['sections']['B']
    section_code = make_section_code(section['exam_report_code'], 'RSPEX')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'FrameworkNumber', 'FrameworkName', 'ExampleNumber', 'ExampleDescription', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    seq = 1
    for fw in section['frameworks']:
        fw_name = f"Framework of Ideas {fw['number']}: {fw['name']}"
        
        # Framework header row
        rows.append(meta + [section_code,
            str(fw['number']),
            fw_name,
            '', '', str(seq)])
        seq += 1
        
        # Example description rows
        for ex in fw['examples']:
            for para_idx in ex['desc_paras']:
                text = doc.paragraphs[para_idx].text.strip()
                if text:
                    rows.append(meta + [section_code,
                        str(fw['number']),
                        fw_name,
                        str(ex['example_number']),
                        text,
                        str(seq)])
                    seq += 1
    
    return rows


def extract_b_annotations(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section B annotated examples: one row per student text paragraph.
    
    Uses paragraph index alignment within Word table cells to map examiner
    remarks to the student text paragraphs they annotate. Empty paragraphs
    act as vertical spacers — a remark at index N maps to the closest text
    paragraph at or before index N.
    
    Multiple remarks mapping to the same text paragraph are combined with
    newline separators. Text paragraphs with no remark have an empty
    ExaminersRemark field.
    
    The first paragraph of each example is treated as the piece's title
    and placed in the TextHeader column.
    """
    section = config['sections']['B']
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'FrameworkNumber', 'FrameworkName', 'ExampleNumber', 'ExampleText',
        'TextHeader', 'TextSegment', 'ExaminersRemark', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'B')
    
    for fw in section['frameworks']:
        for ex in fw['examples']:
            # Per-example SectionCode: VCEEEARBAN01 through VCEEEARBAN09
            section_code = f"{section['exam_report_code']}AN{ex['example_number']:02d}"
            
            table = doc.tables[ex['table_index']]
            text_cell = table.rows[0].cells[0]
            remark_cell = table.rows[0].cells[1]
            
            # Build text paragraphs index (para_index → text)
            text_paras = {}
            for i, para in enumerate(text_cell.paragraphs):
                t = para.text.strip()
                if t:
                    text_paras[i] = t
            
            text_indices = sorted(text_paras.keys())
            if not text_indices:
                continue
            
            # Build remark paragraphs index
            remark_paras = {}
            for i, para in enumerate(remark_cell.paragraphs):
                t = para.text.strip()
                if t:
                    remark_paras[i] = t
            
            # Map each remark to closest preceding text paragraph
            text_to_remarks = defaultdict(list)
            for r_idx in sorted(remark_paras.keys()):
                candidates = [t for t in text_indices if t <= r_idx]
                if candidates:
                    mapped = candidates[-1]
                else:
                    mapped = text_indices[0]
                text_to_remarks[mapped].append(remark_paras[r_idx])
            
            # Build rows: one per text paragraph
            seq = 1
            title_idx = text_indices[0]  # First paragraph is always the title
            
            fw_name = f"Framework of Ideas {fw['number']}: {fw['name']}"
            
            # Build ExampleText from description paragraphs
            example_text_parts = []
            for para_idx in ex['desc_paras']:
                t = doc.paragraphs[para_idx].text.strip()
                if t:
                    example_text_parts.append(t)
            example_text = '\n'.join(example_text_parts)
            
            for t_idx in text_indices:
                text = text_paras[t_idx]
                remarks = text_to_remarks.get(t_idx, [])
                combined_remarks = '\n'.join(remarks) if remarks else ''
                
                # ExampleText only on first row (seq 1)
                ex_text = example_text if seq == 1 else ''
                
                if t_idx == title_idx:
                    # Title row: TextHeader populated, TextSegment empty
                    rows.append(meta + [section_code,
                        str(fw['number']),
                        fw_name,
                        str(ex['example_number']),
                        ex_text,
                        capitalise_first(text),  # Title
                        '',
                        combined_remarks,
                        str(seq)])
                else:
                    # Body row: TextSegment populated, TextHeader empty
                    rows.append(meta + [section_code,
                        str(fw['number']),
                        fw_name,
                        str(ex['example_number']),
                        ex_text,
                        '',
                        text,
                        combined_remarks,
                        str(seq)])
                seq += 1
    
    return rows



# ============================================================================
# SECTION C EXTRACTION FUNCTIONS
# ============================================================================

def extract_c_header(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section C interrelated skills (assessment criteria).
    
    Mirrors Section B assessment structure: intro paragraph then
    skill bullets. Bullet hierarchy preserved via BulletLevel column.
    """
    section = config['sections']['C']
    section_code = make_section_code(section['exam_report_code'], 'EQ')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'AssessmentHeader', 'AssessmentDescription', 'AssessmentSkill',
        'BulletLevel', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'C')
    
    # Intro paragraph - split: first sentence as header, rest as description
    intro = doc.paragraphs[section['assessment_intro_para']].text.strip()
    sentences = split_sentences(intro)
    if sentences:
        rows.append(meta + [section_code, 'Assessment', capitalise_first(sentences[0]), '', '', '1'])
        seq = 2
        for sentence in sentences[1:]:
            rows.append(meta + [section_code, '', capitalise_first(sentence), '', '', str(seq)])
            seq += 1
    else:
        seq = 2
    
    # Skill bullets with level detection
    for para_idx in section['skill_bullet_paras']:
        para = doc.paragraphs[para_idx]
        text = para.text.strip()
        if text:
            # Detect bullet level from style name
            level = '2' if 'level 2' in (para.style.name or '').lower() else '1'
            rows.append(meta + [section_code, '', '', capitalise_first(text), level, str(seq)])
            seq += 1
    
    return rows


def extract_c_context(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section C context/background paragraphs.
    
    These paragraphs describe the task context: background information,
    the text presented, and the task design. Each paragraph becomes a row
    with first-sentence splitting for scannability.
    """
    section = config['sections']['C']
    section_code = make_section_code(section['exam_report_code'], 'CTX')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'ContextHeader', 'ContextContent', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'C')
    
    # Header row
    rows.append(meta + [section_code, 'Context', '', '1'])
    
    seq = 2
    for para_idx in section['context_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if not text:
            continue
        
        sentences = split_sentences(text)
        if not sentences:
            continue
        
        # First sentence as header
        rows.append(meta + [section_code, capitalise_first(sentences[0]), '', str(seq)])
        seq += 1
        
        # Remaining sentences as content bullets
        for sentence in sentences[1:]:
            rows.append(meta + [section_code, '', capitalise_first(sentence), str(seq)])
            seq += 1
    
    return rows


def extract_c_argument(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section C line of argument and transition paragraphs.
    
    Captures the numbered argument steps and the framing paragraphs
    before and after them.
    """
    section = config['sections']['C']
    section_code = make_section_code(section['exam_report_code'], 'ARG')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'ArgumentHeader', 'ArgumentStep', 'StepNumber', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'C')
    
    # Intro paragraph as header
    intro = doc.paragraphs[section['argument_intro_para']].text.strip()
    rows.append(meta + [section_code, capitalise_first(intro), '', '', '1'])
    
    # Numbered argument steps
    seq = 2
    step_num = 1
    for para_idx in section['argument_step_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [section_code, '', capitalise_first(text), str(step_num), str(seq)])
            seq += 1
            step_num += 1
    
    # Transition paragraph
    transition = doc.paragraphs[section['argument_transition_para']].text.strip()
    if transition:
        rows.append(meta + [section_code, capitalise_first(transition), '', '', str(seq)])
    
    return rows


def extract_c_language(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section C language choices and visual cues.
    
    Captures language strategy bullets and visual cue bullets, each
    with their intro paragraph. Uses LanguageType to distinguish
    'Language' from 'Visual' entries.
    """
    section = config['sections']['C']
    section_code = make_section_code(section['exam_report_code'], 'LNG')
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'LanguageType', 'LanguageHeader', 'LanguageContent', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'C')
    
    # === LANGUAGE CHOICES ===
    lang_intro = doc.paragraphs[section['language_intro_para']].text.strip()
    rows.append(meta + [section_code, 'Language', capitalise_first(lang_intro), '', '1'])
    
    seq = 2
    for para_idx in section['language_bullet_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [section_code, 'Language', '', capitalise_first(text), str(seq)])
            seq += 1
    
    # === VISUAL CUES ===
    visual_intro = doc.paragraphs[section['visual_intro_para']].text.strip()
    rows.append(meta + [section_code, 'Visual', capitalise_first(visual_intro), '', str(seq)])
    seq += 1
    
    for para_idx in section['visual_bullet_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [section_code, 'Visual', '', capitalise_first(text), str(seq)])
            seq += 1
    
    # === CLOSING PARAGRAPHS ===
    for para_idx in section['language_closing_paras']:
        text = doc.paragraphs[para_idx].text.strip()
        if text:
            rows.append(meta + [section_code, 'Closing', capitalise_first(text), '', str(seq)])
            seq += 1
    
    return rows


def extract_c_strategies(doc: Document, config: Dict) -> List[List[str]]:
    """Extract Section C enhanced and limited strategies.
    
    Enhanced: Table 19 (3 cols: Skill, Successful strategy, Example).
    Limited: Tables 20+21 (2 cols: Strategy, Explanation) — table 21
    is a continuation of 20 with no header row.
    
    Limited explanations use BulletLevel to distinguish intro lines (1)
    from sub-bullets (2).
    """
    section = config['sections']['C']
    
    rows = []
    rows.append([
        'SubjectArea', 'Subject', 'Band', 'AssessmentType', 'AssessmentInformationDetails',
        'AssessmentYears', 'UnitASCode', 'ExamReportCode', 'SectionCode',
        'StrategyType', 'StrategyHeader', 'StrategySkill', 'StrategyExplanation',
        'StrategyExample', 'BulletLevel', 'Sequence'
    ])
    
    meta = get_metadata_row(config, 'C')
    
    # === ENHANCED STRATEGIES ===
    enhanced_code = make_section_code(section['exam_report_code'], 'ER')
    rows.append(meta + [enhanced_code, 'EnhancedStrategies', 'Strategies that enhanced responses', '', '', '', '', '1'])
    
    enhanced_table = doc.tables[section['enhanced_table_index']]
    seq = 2
    for row in enhanced_table.rows[1:]:  # Skip header row
        skill = row.cells[0].text.strip()
        strategy = row.cells[1].text.strip()
        example = row.cells[2].text.strip() if len(row.cells) > 2 else ''
        
        rows.append(meta + [enhanced_code, 'EnhancedStrategies', '',
            capitalise_first(skill), capitalise_first(strategy),
            example, '', str(seq)])
        seq += 1
    
    # === LIMITED STRATEGIES ===
    limited_code = make_section_code(section['exam_report_code'], 'LR')
    rows.append(meta + [limited_code, 'LimitedStrategies', 'Strategies that limited responses', '', '', '', '', '1'])
    
    seq = 2
    for table_idx in section['limited_table_indices']:
        table = doc.tables[table_idx]
        
        # Table 20 has a header row, table 21 does not
        start_row = 1 if table_idx == section['limited_table_indices'][0] else 0
        
        for row in table.rows[start_row:]:
            strategy_name = row.cells[0].text.strip()
            full_explanation = row.cells[1].text.strip()
            
            # Split on newlines — first line is intro (level 1), rest are sub-bullets (level 2)
            lines = full_explanation.split('\n')
            intro_line = lines[0].strip() if lines else ''
            
            rows.append(meta + [limited_code, 'LimitedStrategies', '',
                capitalise_first(strategy_name), capitalise_first(intro_line),
                '', '1', str(seq)])
            seq += 1
            
            for line in lines[1:]:
                line = line.strip()
                if line:
                    rows.append(meta + [limited_code, 'LimitedStrategies', '',
                        capitalise_first(strategy_name), capitalise_first(line),
                        '', '2', str(seq)])
                    seq += 1
    
    # === CLOSING ===
    closing_code = make_section_code(section['exam_report_code'], 'LC')
    closing = doc.paragraphs[section['strategies_closing_para']].text.strip()
    if closing:
        rows.append(meta + [closing_code, 'Closing', capitalise_first(closing), '', '', '', '', str(seq)])
    
    return rows


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def write_csv(rows: List[List[str]], filepath: str):
    """Write rows to CSV file."""
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)
    print(f"  Created: {filepath} ({len(rows)-1} data rows)")


def main():
    """Main entry point."""
    print("=" * 70)
    print("VCE English Exam Report Parser v2.4")
    print("=" * 70)
    
    # Validate input files
    valid, msg = validate_input_files(CONFIG)
    if not valid:
        print(msg)
        return
    
    print(f"\n{msg}")
    
    # Create output directory
    output_dir = Path(CONFIG['output_dir'])
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"\nOutput directory: {output_dir}")
    
    # Load documents
    print("\nLoading documents...")
    report_doc = Document(CONFIG['report_doc'])
    print(f"  Loaded: {CONFIG['report_doc']}")
    print(f"  Tables found: {len(report_doc.tables)}")
    
    # === OVERVIEW (All Sections) ===
    print("\n" + "-" * 50)
    print("Processing Overview: All Sections")
    print("-" * 50)
    
    overview_rows = extract_overview(report_doc, CONFIG)
    write_csv(overview_rows, output_dir / 'vcaa_vce_sd_english_exam_report_overview.csv')
    
    # === SECTION A DETAIL ===
    print("\n" + "-" * 50)
    print("Processing Section A: Analytical response to a text")
    print("-" * 50)
    
    # Extract and write each file
    topics_rows = extract_a_topics(report_doc, CONFIG)
    write_csv(topics_rows, output_dir / 'vcaa_vce_sd_english_exam_report_topics_a.csv')
    
    invitations_rows = extract_a_invitations(report_doc, CONFIG)
    write_csv(invitations_rows, output_dir / 'vcaa_vce_sd_english_exam_report_invitations_a.csv')
    
    verbs_rows = extract_a_verbs(report_doc, CONFIG)
    write_csv(verbs_rows, output_dir / 'vcaa_vce_sd_english_exam_report_verbs_a.csv')
    
    implications_rows = extract_a_implications(report_doc, CONFIG)
    write_csv(implications_rows, output_dir / 'vcaa_vce_sd_english_exam_report_implications_a.csv')
    
    interrelated_rows = extract_a_interrelated_skills(report_doc, CONFIG)
    write_csv(interrelated_rows, output_dir / 'vcaa_vce_sd_english_exam_report_interrelated_skills_a.csv')
    
    strategies_rows = extract_a_strategies(report_doc, CONFIG)
    write_csv(strategies_rows, output_dir / 'vcaa_vce_sd_english_exam_report_strategies_a.csv')
    
    # === SECTION B DETAIL ===
    print("\n" + "-" * 50)
    print("Processing Section B: Creating a text")
    print("-" * 50)
    
    header_b_rows = extract_b_header(report_doc, CONFIG)
    write_csv(header_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_header_b.csv')
    
    assessment_b_rows = extract_b_assessment(report_doc, CONFIG)
    write_csv(assessment_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_assessment_b.csv')
    
    strategies_b_rows = extract_b_strategies(report_doc, CONFIG)
    write_csv(strategies_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_strategies_b.csv')
    
    responses_b_rows = extract_b_responses(report_doc, CONFIG)
    write_csv(responses_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_responses_b.csv')
    
    examples_b_rows = extract_b_examples(report_doc, CONFIG)
    write_csv(examples_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_examples_b.csv')
    
    annotations_b_rows = extract_b_annotations(report_doc, CONFIG)
    write_csv(annotations_b_rows, output_dir / 'vcaa_vce_sd_english_exam_report_annotations_b.csv')
    
    # === SECTION C DETAIL ===
    print("\n" + "-" * 50)
    print("Processing Section C: Analysis of argument and language")
    print("-" * 50)
    
    header_c_rows = extract_c_header(report_doc, CONFIG)
    write_csv(header_c_rows, output_dir / 'vcaa_vce_sd_english_exam_report_header_c.csv')
    
    context_c_rows = extract_c_context(report_doc, CONFIG)
    write_csv(context_c_rows, output_dir / 'vcaa_vce_sd_english_exam_report_context_c.csv')
    
    argument_c_rows = extract_c_argument(report_doc, CONFIG)
    write_csv(argument_c_rows, output_dir / 'vcaa_vce_sd_english_exam_report_argument_c.csv')
    
    language_c_rows = extract_c_language(report_doc, CONFIG)
    write_csv(language_c_rows, output_dir / 'vcaa_vce_sd_english_exam_report_language_c.csv')
    
    strategies_c_rows = extract_c_strategies(report_doc, CONFIG)
    write_csv(strategies_c_rows, output_dir / 'vcaa_vce_sd_english_exam_report_strategies_c.csv')
    
    print("\n" + "=" * 70)
    print("COMPLETE: Overview + Section A + Section B + Section C Detail")
    print("=" * 70)
    print("\n  1 overview CSV (all sections)")
    print("  6 Section A detail CSVs")
    print("  6 Section B detail CSVs")
    print("  5 Section C detail CSVs")


if __name__ == '__main__':
    main()

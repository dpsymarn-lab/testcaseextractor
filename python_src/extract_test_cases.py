#!/usr/bin/env python3
"""
Extract CBAM E2E test cases from .docx/.txt/.md files and output a Jira
Zephyr Scale-compatible CSV.

This version handles both known Word structures:
  A) DLM-style: one test-case table with metadata rows and 5-column step rows
     at the top level: N | Actor(s) Actions | Data | N | System Response
  B) DRMC-style: a 2-column metadata table where the "Steps" value cell contains
     a nested 4-column table: No. | Appl./User | Action | Expectations
"""

import argparse
import csv
import os
import re
import sys


FIELDNAMES = [
    'Name',
    'ID',
    'Source Document',
    'Objective',
    'Precondition',
    'Labels',
    'Priority',
    'Status',
    'Step',
    'Step Action',
    'Step Data',
    'Step Expected Result',
    'Postcondition',
    'Pass Criteria',
]


def clean_text(value):
    """Normalize Word cell/paragraph text for CSV output."""
    if value is None:
        return ''
    value = value.replace('\xa0', ' ')
    value = value.replace('\u200b', '')
    value = value.replace('\r', '\n')
    # Keep bullets readable but make the CSV compact.
    value = re.sub(r'[ \t]*\n[ \t]*', ' ', value)
    value = re.sub(r'\s+', ' ', value)
    return value.strip()


def norm_key(value):
    """Normalize labels so small naming differences do not break extraction."""
    return re.sub(r'[^a-z0-9]+', '', clean_text(value).lower())


def dedupe_consecutive(values):
    """Word merged cells appear as repeated cell values; collapse repeats."""
    result = []
    for value in values:
        text = clean_text(value)
        if not result or text != result[-1]:
            result.append(text)
    return result


def row_cells(row):
    return dedupe_consecutive([cell.text for cell in row.cells])


def is_step_number(value):
    return bool(re.match(r'^\d+(?:\.\d+)*\.?$', clean_text(value)))


def first_cbam_tc_id(text):
    m = re.search(r'CBAM-[A-Z0-9-]+-TC-\d+', text or '', flags=re.IGNORECASE)
    return m.group(0) if m else ''


def parse_dlm_steps(table):
    """Parse DLM-style top-level 5-column steps from one test-case table."""
    steps = []
    in_steps = False

    for row in table.rows:
        cells = row_cells(row)
        keys = [norm_key(c) for c in cells]

        if any(k == 'postconditions' or k == 'postcondition' for k in keys):
            if in_steps:
                break

        # Step header usually dedupes to: Actor(s) Actions | Data | System Response
        if cells and norm_key(cells[0]) in {'actorsactions', 'actoractions'}:
            in_steps = True
            continue

        if not in_steps:
            continue

        # Expected step row after dedupe:
        # [N, action, data, N, expected]
        if len(cells) >= 5 and is_step_number(cells[0]):
            steps.append({
                'action': cells[1],
                'data': cells[2],
                'expected': cells[4],
            })

    return steps


def parse_nested_4col_steps(nested_table):
    """Parse DRMC-style nested table: No. | Appl./User | Action | Expectations."""
    steps = []
    for row in nested_table.rows:
        cells = row_cells(row)
        if len(cells) < 4:
            continue

        # Skip header and merged introduction rows.
        if not is_step_number(cells[0]):
            continue

        steps.append({
            'action': cells[2],
            'data': cells[1],
            'expected': cells[3],
        })
    return steps


def parse_nested_steps_from_metadata_table(table):
    """Find a row named Steps and parse a nested 4-column table inside its value cell."""
    for row in table.rows:
        if len(row.cells) < 2:
            continue
        if norm_key(row.cells[0].text) != 'steps':
            continue

        # The nested table is normally in the second logical cell. Because python-docx
        # still exposes physical cells, check all cells except the label cell.
        for cell in row.cells[1:]:
            for nested in cell.tables:
                parsed = parse_nested_4col_steps(nested)
                if parsed:
                    return parsed
    return []


def parse_testcase_table(table, source_doc, current_heading=''):
    """Return one test-case dict if this table is a supported test-case table."""
    rows = [row_cells(row) for row in table.rows]
    if not rows:
        return None

    tc = {
        'Name': '',
        'ID': '',
        'Source Document': source_doc,
        'Purpose': '',
        'Actors': '',
        'Pre-Conditions': '',
        'Post-Conditions': '',
        'Pass Criteria': '',
        'Steps': [],
    }

    # DLM-style title row is a single merged title cell.
    if len(rows[0]) == 1 and not rows[0][0].lower().startswith('table '):
        tc['Name'] = rows[0][0]

    for cells in rows:
        if len(cells) < 2:
            continue
        key = norm_key(cells[0])
        value = cells[1]

        if key in {'id', 'testcaseid'}:
            tc['ID'] = first_cbam_tc_id(value) or value
        elif key in {'purpose', 'testdescription', 'description'}:
            tc['Purpose'] = value
        elif key in {'actors', 'actor', 'appluser', 'applicationuser'}:
            tc['Actors'] = value
        elif key in {'preconditions', 'precondition'}:
            tc['Pre-Conditions'] = value
        elif key in {'postconditions', 'postcondition'}:
            tc['Post-Conditions'] = value
        elif key in {'passcriteria', 'passcriterion'}:
            tc['Pass Criteria'] = value

    if not tc['ID']:
        # Some tables might not have an ID row, but the heading/title may contain it.
        candidates = ' '.join([' '.join(r) for r in rows[:3]] + [current_heading])
        tc['ID'] = first_cbam_tc_id(candidates)

    if not tc['ID']:
        return None

    # Avoid false positives from traceability/list tables that contain test case IDs
    # but are not actual test-case definition tables.
    if not tc['Purpose']:
        return None

    if not tc['Name']:
        heading = clean_text(current_heading)
        if heading and first_cbam_tc_id(heading):
            tc['Name'] = re.sub(r'^.*?CBAM-[A-Z0-9-]+-TC-\d+\s*:?\s*', '', heading, flags=re.IGNORECASE).strip()
        if not tc['Name']:
            tc['Name'] = tc['Purpose'] or tc['ID']

    # Prefer DLM top-level steps. If none, try DRMC nested steps.
    tc['Steps'] = parse_dlm_steps(table)
    if not tc['Steps']:
        tc['Steps'] = parse_nested_steps_from_metadata_table(table)

    return tc


def iter_docx_blocks(doc):
    """Yield paragraphs and tables from the document body in document order."""
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit('}', 1)[-1]
        if tag == 'p':
            yield 'paragraph', Paragraph(child, doc)
        elif tag == 'tbl':
            yield 'table', Table(child, doc)


def extract_test_cases_from_docx(path, source_doc):
    try:
        from docx import Document
    except ImportError:
        sys.exit('ERROR: python-docx is required for .docx input. Install it with: pip install python-docx')

    doc = Document(path)
    cases = []
    current_heading = ''

    for kind, block in iter_docx_blocks(doc):
        if kind == 'paragraph':
            text = clean_text(block.text)
            if not text:
                continue
            style = block.style.name if block.style else ''
            if style.lower().startswith('heading') or first_cbam_tc_id(text):
                current_heading = text
        else:
            tc = parse_testcase_table(block, source_doc, current_heading)
            if tc:
                cases.append(tc)

    return cases


# ---------------------------------------------------------------------------
# Text/markdown fallback from the original script, with a more tolerant parser
# ---------------------------------------------------------------------------

def extract_field(tc_lines, key_pattern):
    for line in tc_lines:
        stripped = line.strip()
        m = re.match(rf'^\|\s*{key_pattern}\s*\|\s*(.*?)\s*\|?\s*$', stripped, re.IGNORECASE)
        if m:
            return m.group(1).strip().rstrip('|').strip()
    return ''


def split_md_row(line):
    return [cell.strip() for cell in line.strip().strip('|').split('|')]


def parse_steps(tc_lines):
    steps = []
    in_steps = False
    step_format = None

    for line in tc_lines:
        stripped = line.strip()
        if not stripped.startswith('|'):
            continue

        cells = split_md_row(stripped)
        normalized = [norm_key(c) for c in cells]

        if cells and normalized[0] in {'actorsactions', 'actoractions'}:
            in_steps = True
            step_format = 'cbam_5_col'
            continue

        if 'appluser' in normalized or 'applicationuser' in normalized:
            in_steps = True
            step_format = 'four_col'
            continue

        if in_steps:
            if cells and normalized[0] in {'postconditions', 'postcondition'}:
                break
            if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                continue

            if step_format == 'cbam_5_col' and len(cells) >= 5 and is_step_number(cells[0]):
                steps.append({'action': cells[1], 'data': cells[2], 'expected': cells[4]})
            elif step_format == 'four_col' and len(cells) >= 4 and is_step_number(cells[0]):
                steps.append({'action': cells[2], 'data': cells[1], 'expected': cells[3]})

    return steps


def extract_test_cases_from_text(text, source_doc):
    lines = text.split('\n')
    tc_starts = [
        i for i, line in enumerate(lines)
        if re.match(r'^#{1,6}\s+CBAM-.*TC-\d+', line.strip(), re.IGNORECASE)
    ]

    test_cases = []
    for idx, start in enumerate(tc_starts):
        end = tc_starts[idx + 1] if idx + 1 < len(tc_starts) else len(lines)
        tc_lines = lines[start:end]
        heading = tc_lines[0].strip().lstrip('#').strip()
        parts = heading.split(':', 1)
        tc_id = parts[0].strip()
        tc_name = parts[1].strip() if len(parts) == 2 else heading

        table_id = extract_field(tc_lines, r'(?:ID|Test Case ID)')
        if first_cbam_tc_id(table_id):
            tc_id = first_cbam_tc_id(table_id)

        test_cases.append({
            'Name': tc_name,
            'ID': tc_id,
            'Source Document': source_doc,
            'Purpose': extract_field(tc_lines, r'(?:Purpose|Test Description)'),
            'Actors': extract_field(tc_lines, r'Actor\(s\)'),
            'Pre-Conditions': extract_field(tc_lines, r'(?:Pre-Conditions|Preconditions)'),
            'Post-Conditions': extract_field(tc_lines, r'(?:Post-Conditions|Postcondition)'),
            'Pass Criteria': extract_field(tc_lines, r'Pass Criteria'),
            'Steps': parse_steps(tc_lines),
        })

    return test_cases


def load_file(path):
    source_name = os.path.splitext(os.path.basename(path))[0]
    ext = os.path.splitext(path)[1].lower()

    if ext == '.docx':
        return extract_test_cases_from_docx(path, source_name)

    with open(path, 'r', encoding='utf-8') as f:
        return extract_test_cases_from_text(f.read(), source_name)


def write_zephyr_csv(test_cases, output_path):
    rows = []
    for tc in test_cases:
        base = {
            'Name': tc['Name'],
            'ID': tc['ID'],
            'Source Document': tc['Source Document'],
            'Objective': tc['Purpose'],
            'Precondition': tc['Pre-Conditions'],
            'Labels': 'E2E',
            'Priority': 'Medium',
            'Status': 'Draft',
            'Postcondition': tc['Post-Conditions'],
            'Pass Criteria': tc['Pass Criteria'],
        }

        steps = tc.get('Steps') or []
        if steps:
            for i, step in enumerate(steps):
                row = base.copy() if i == 0 else {k: '' for k in FIELDNAMES}
                row['Step'] = str(i + 1)
                row['Step Action'] = step.get('action', '')
                row['Step Data'] = step.get('data', '')
                row['Step Expected Result'] = step.get('expected', '')
                rows.append(row)
        else:
            row = base.copy()
            row.update({'Step': '', 'Step Action': '', 'Step Data': '', 'Step Expected Result': ''})
            rows.append(row)

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description='Extract CBAM E2E test cases from .docx/.txt/.md to Jira Zephyr CSV.')
    parser.add_argument('files', nargs='+', metavar='FILE')
    parser.add_argument('-o', '--output', default='test_cases_zephyr.csv', metavar='OUTPUT')
    args = parser.parse_args()

    all_cases = []
    for path in args.files:
        if not os.path.isfile(path):
            print(f'WARNING: File not found, skipping: {path}', file=sys.stderr)
            continue
        print(f'Processing: {path}')
        cases = load_file(path)
        print(f'  -> {len(cases)} test cases found, {sum(len(c["Steps"]) for c in cases)} steps found')
        for tc in cases:
            if not tc['Steps']:
                print(f'WARNING: No steps found for {tc["ID"]}', file=sys.stderr)
        all_cases.extend(cases)

    if not all_cases:
        sys.exit('No test cases found.')

    write_zephyr_csv(all_cases, args.output)
    print(f'\nTotal test cases : {len(all_cases)}')
    print(f'Total steps      : {sum(len(tc["Steps"]) for tc in all_cases)}')
    print(f'CSV written to   : {args.output}')


if __name__ == '__main__':
    main()

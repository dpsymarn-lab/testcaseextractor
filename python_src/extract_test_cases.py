#!/usr/bin/env python3
"""
Extract E2E test cases from CBAM .docx or pre-converted .txt/.md files
and output a Jira Zephyr Scale-compatible CSV.

Usage:
    # One or more .docx files:
    python extract_test_cases.py doc1.docx doc2.docx

    # Pre-converted text/markdown files:
    python extract_test_cases.py doc1.txt doc2.md

    # Mix of both:
    python extract_test_cases.py doc1.docx notes.txt

    # Custom output path (default: test_cases_zephyr.csv):
    python extract_test_cases.py doc1.docx -o my_output.csv

Requirements:
    pip install python-docx          # only needed for .docx input
"""

import re
import csv
import sys
import os
import argparse
import tempfile


# ---------------------------------------------------------------------------
# DOCX -> plain text conversion (markdown-ish tables)
# ---------------------------------------------------------------------------

def docx_to_text(filepath):
    """Convert a .docx file to a plain-text string using python-docx.
    Tables are rendered as pipe-delimited markdown rows so the rest of
    the parser can handle them identically to pre-converted text files.
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        sys.exit(
            "ERROR: python-docx is required to parse .docx files.\n"
            "Install it with:  pip install python-docx"
        )

    doc = Document(filepath)
    lines = []

    # Iterate the document body in order (paragraphs and tables interleaved)
    body = doc.element.body
    for child in body:
        tag = child.tag.split('}')[-1]  # strip namespace

        if tag == 'p':
            # Paragraph
            from docx.text.paragraph import Paragraph
            para = Paragraph(child, doc)
            style = para.style.name if para.style else ''
            text = para.text.strip()
            if not text:
                continue
            # Map heading styles to markdown headings
            m = re.match(r'Heading (\d)', style)
            if m:
                level = int(m.group(1))
                lines.append('#' * level + ' ' + text)
            else:
                lines.append(text)

        elif tag == 'tbl':
            # Table
            from docx.table import Table
            tbl = Table(child, doc)
            for row in tbl.rows:
                cells = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                lines.append('| ' + ' | '.join(cells) + ' |')
            lines.append('')  # blank line after table

    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

def extract_field(tc_lines, key_pattern):
    """Return the value from a markdown table row matching | key | value |."""
    for line in tc_lines:
        stripped = line.strip()
        m = re.match(rf'^\|\s*{key_pattern}\s*\|\s*(.*?)\s*\|?\s*$', stripped)
        if m:
            return m.group(1).strip().rstrip('|').strip()
    return ''


def parse_steps(tc_lines):
    """Parse step rows, returning list of {action, data, expected} dicts."""
    steps = []
    in_steps = False

    for line in tc_lines:
        stripped = line.strip()

        if re.match(r'^\|\s*Actor\(s\)\s*Actions\s*\|', stripped, re.IGNORECASE):
            in_steps = True
            continue

        if in_steps:
            if re.match(r'^\|\s*Post-Conditions\s*\|', stripped, re.IGNORECASE):
                break
            # Row format: | N | action | data | N | expected |
            m = re.match(
                r'^\|\s*(\d+)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*\d*\s*\|\s*(.*?)\s*\|?\s*$',
                stripped
            )
            if m:
                steps.append({
                    'action':   m.group(2).strip(),
                    'data':     m.group(3).strip(),
                    'expected': m.group(4).strip().rstrip('|').strip(),
                })

    return steps


# ---------------------------------------------------------------------------
# Core extraction
# ---------------------------------------------------------------------------

def extract_test_cases(text, source_doc):
    """Extract test cases from a plain-text (markdown-table) string."""
    lines = text.split('\n')

    # Locate headings that contain a test case ID pattern
    tc_starts = [
        i for i, line in enumerate(lines)
        if re.match(r'^#{1,6}\s+CBAM-.*TC-\d+', line.strip())
    ]

    test_cases = []
    for idx, start in enumerate(tc_starts):
        end = tc_starts[idx + 1] if idx + 1 < len(tc_starts) else len(lines)
        tc_lines = lines[start:end]

        heading = tc_lines[0].strip().lstrip('#').strip()
        parts = heading.split(':', 1)
        tc_id   = parts[0].strip()
        tc_name = parts[1].strip() if len(parts) == 2 else heading

        # Prefer ID from the table body if available
        table_id = extract_field(tc_lines, 'ID')
        if table_id.startswith('CBAM'):
            tc_id = table_id

        test_cases.append({
            'Name':            tc_name,
            'ID':              tc_id,
            'Source Document': source_doc,
            'Purpose':         extract_field(tc_lines, 'Purpose'),
            'Actors':          extract_field(tc_lines, r'Actor\(s\)'),
            'Pre-Conditions':  extract_field(tc_lines, 'Pre-Conditions'),
            'Post-Conditions': extract_field(tc_lines, 'Post-Conditions'),
            'Pass Criteria':   extract_field(tc_lines, 'Pass Criteria'),
            'Steps':           parse_steps(tc_lines),
        })

    return test_cases


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

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


def write_zephyr_csv(test_cases, output_path):
    """Write Jira Zephyr Scale import CSV (one row per step)."""
    rows = []
    for tc in test_cases:
        base = {
            'Name':            tc['Name'],
            'ID':              tc['ID'],
            'Source Document': tc['Source Document'],
            'Objective':       tc['Purpose'],
            'Precondition':    tc['Pre-Conditions'],
            'Labels':          'E2E',
            'Priority':        'Medium',
            'Status':          'Draft',
            'Postcondition':   tc['Post-Conditions'],
            'Pass Criteria':   tc['Pass Criteria'],
        }

        steps = tc['Steps']
        if steps:
            for i, step in enumerate(steps):
                row = base.copy() if i == 0 else {k: '' for k in FIELDNAMES}
                row['Step'] = str(i + 1)
                row['Step Action']          = step['action']
                row['Step Data']            = step['data']
                row['Step Expected Result'] = step['expected']
                rows.append(row)
        else:
            base.update({'Step': '', 'Step Action': '', 'Step Data': '', 'Step Expected Result': ''})
            rows.append(base)

    # UTF-8 BOM so Excel opens it cleanly without encoding issues
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def load_file(path):
    """Return (text_content, source_name) for a .docx or text file."""
    ext = os.path.splitext(path)[1].lower()
    source_name = os.path.splitext(os.path.basename(path))[0]

    if ext == '.docx':
        text = docx_to_text(path)
    else:
        with open(path, 'r', encoding='utf-8') as f:
            text = f.read()

    return text, source_name


def main():
    parser = argparse.ArgumentParser(
        description='Extract CBAM E2E test cases from .docx/.txt files to Jira Zephyr CSV.'
    )
    parser.add_argument(
        'files',
        nargs='+',
        metavar='FILE',
        help='One or more .docx or .txt/.md files containing test cases'
    )
    parser.add_argument(
        '-o', '--output',
        default='test_cases_zephyr.csv',
        metavar='OUTPUT',
        help='Output CSV file path (default: test_cases_zephyr.csv)'
    )
    args = parser.parse_args()

    all_cases = []
    for path in args.files:
        if not os.path.isfile(path):
            print(f"WARNING: File not found, skipping: {path}", file=sys.stderr)
            continue
        print(f"Processing: {path}")
        text, source_name = load_file(path)
        cases = extract_test_cases(text, source_name)
        print(f"  -> {len(cases)} test cases found")
        all_cases.extend(cases)

    if not all_cases:
        sys.exit("No test cases found. Check that the files contain CBAM-...-TC-NNN headings.")

    total_steps = sum(len(tc['Steps']) for tc in all_cases)
    print(f"\nTotal test cases : {len(all_cases)}")
    print(f"Total steps      : {total_steps}")

    write_zephyr_csv(all_cases, args.output)
    print(f"CSV written to   : {args.output}")


if __name__ == '__main__':
    main()

import re
import csv

def parse_steps(tc_lines):
    """Parse step rows from test case lines, returning list of dicts."""
    steps = []
    in_steps = False
    
    for line in tc_lines:
        stripped = line.strip()
        if re.match(r'^\|\s*Actor\(s\)\s*Actions\s*\|', stripped, re.IGNORECASE):
            in_steps = True
            continue
        if in_steps:
            # Detect end of steps section
            if re.match(r'^\|\s*Post-Conditions\s*\|', stripped, re.IGNORECASE):
                break
            # Match step row: | N | action | data | N | response |
            m = re.match(r'^\|\s*(\d+)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*\d*\s*\|\s*(.*?)\s*\|?\s*$', stripped)
            if m:
                action = m.group(2).strip()
                data = m.group(3).strip()
                response = m.group(4).strip().rstrip('|').strip()
                steps.append({'action': action, 'data': data, 'expected': response})
    return steps

def extract_field(tc_lines, key_pattern):
    for line in tc_lines:
        stripped = line.strip()
        m = re.match(rf'^\|\s*{key_pattern}\s*\|\s*(.*?)\s*\|?\s*$', stripped)
        if m:
            return m.group(1).strip().rstrip('|').strip()
    return ''

def extract_test_cases(filepath, source_doc):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Find test case headings at any level (##, ###, ####)
    tc_starts = []
    for i, line in enumerate(lines):
        if re.match(r'^#{1,6}\s+CBAM-.*TC-\d+', line.strip()):
            tc_starts.append(i)
    
    test_cases = []
    for idx, start in enumerate(tc_starts):
        end = tc_starts[idx + 1] if idx + 1 < len(tc_starts) else len(lines)
        tc_lines = [l.rstrip('\n') for l in lines[start:end]]
        
        heading = tc_lines[0].strip().lstrip('#').strip()
        heading_parts = heading.split(':', 1)
        tc_id = heading_parts[0].strip()
        tc_name = heading_parts[1].strip() if len(heading_parts) == 2 else heading
        
        # Prefer ID from table if exists
        table_id = extract_field(tc_lines, 'ID')
        if table_id.startswith('CBAM'):
            tc_id = table_id

        tc = {
            'Name': tc_name,
            'ID': tc_id,
            'Source Document': source_doc,
            'Purpose': extract_field(tc_lines, 'Purpose'),
            'Actors': extract_field(tc_lines, r'Actor\(s\)'),
            'Pre-Conditions': extract_field(tc_lines, 'Pre-Conditions'),
            'Post-Conditions': extract_field(tc_lines, 'Post-Conditions'),
            'Pass Criteria': extract_field(tc_lines, 'Pass Criteria'),
            'Steps': parse_steps(tc_lines),
        }
        test_cases.append(tc)
    
    return test_cases


def write_zephyr_csv(test_cases, output_path):
    """
    Jira Zephyr Scale CSV format.
    One row per step; metadata repeated on first step row.
    """
    fieldnames = [
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
    
    rows = []
    for tc in test_cases:
        steps = tc['Steps']
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
        
        if steps:
            for i, step in enumerate(steps):
                row = base.copy() if i == 0 else {k: '' for k in fieldnames}
                if i > 0:
                    row['Name'] = tc['Name']  # keep name for Zephyr grouping
                    row['ID'] = tc['ID']
                row['Step'] = str(i + 1)
                row['Step Action'] = step['action']
                row['Step Data'] = step['data']
                row['Step Expected Result'] = step['expected']
                rows.append(row)
        else:
            base.update({'Step': '', 'Step Action': '', 'Step Data': '', 'Step Expected Result': ''})
            rows.append(base)
    
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


# Run extraction
dlm_cases = extract_test_cases('C:\Projects\TES-TAXUD\Development\helpfiles\output.md', 'SDEV-CBAM-DLM-TDS-E2ETC-v2_00')
# igm_cases = extract_test_cases('/home/claude/igm_content.txt', 'SDEV-CBAM-IGM-TDS-E2ETC-v3_00')
all_cases = dlm_cases # + igm_cases

print(f"DLM test cases: {len(dlm_cases)}")
#print(f"IGM test cases: {len(igm_cases)}")
print(f"Total test cases: {len(all_cases)}")
total_steps = sum(len(tc['Steps']) for tc in all_cases)
print(f"Total steps: {total_steps}")

write_zephyr_csv(all_cases, 'C:\Projects\TES-TAXUD\Development\helpfiles\CBAM_E2E_TestCases_Zephyr.csv')
print("CSV written to outputs.")

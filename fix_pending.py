"""Fix PENDING_REMARK_MAP and get_next_action in both generate_report.py files."""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# New simple REMARK map - only 3 categories
NEW_REMARK_MAP = """\
PENDING_REMARK_MAP = {
    '306': 'Deliver',
    '309': 'Deliver',
    '311': 'Deliver',
    '210': 'Check',
    '300': 'Check',
    '302': 'Check',
    '310': 'Check',
    '500': 'Return',
}"""

# New get_next_action - removed, just return empty
NEW_NEXT_ACTION = """\
def get_next_action(status_code):
    sc = str(status_code).strip()
    if sc in ('306', '309', '311'):
        return 'Deliver'
    elif sc in ('300', '302', '210', '310'):
        return 'Check'
    elif sc in ('500',):
        return 'Return'
    return ''"""

for rel_path in ['generate_report.py', 'push_bot/generate_report.py']:
    path = os.path.join(HERE, rel_path)
    if not os.path.exists(path):
        print(f"SKIP: {rel_path}")
        continue

    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace PENDING_REMARK_MAP
    start = content.find('PENDING_REMARK_MAP = {')
    if start == -1:
        print(f"REMARK MAP not found in {rel_path}")
        continue
    end = content.find('}', start) + 1
    content = content[:start] + NEW_REMARK_MAP + content[end:]

    # Replace get_next_action function (root file only)
    fn_start = content.find('def get_next_action(')
    if fn_start != -1:
        # Find end of function (next def or double newline)
        fn_end = content.find('\n\n', fn_start)
        if fn_end == -1:
            fn_end = content.find('\ndef ', fn_start + 1)
        content = content[:fn_start] + NEW_NEXT_ACTION + content[fn_end:]

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"DONE: {rel_path}")

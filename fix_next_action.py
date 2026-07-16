"""Fix _pending_next_action in push_bot/generate_report.py"""
import os, re

path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'push_bot', 'generate_report.py')
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find the _pending_next_action function and replace it
pattern = r"(            def _pending_next_action\(row\):.*?return '')"
replacement = """            def _pending_next_action(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                if sc in ('306', '309', '311'):
                    return 'Deliver'
                elif sc in ('300', '302', '210', '310'):
                    return 'Check'
                elif sc in ('500',):
                    return 'Return'
                return ''"""

content_new = re.sub(pattern, replacement, content, flags=re.DOTALL)

if content_new != content:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content_new)
    print("DONE")
else:
    print("NO MATCH")

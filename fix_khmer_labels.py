import os, re

HERE = os.path.dirname(os.path.abspath(__file__))

# Khmer labels: Deliver, Return, Check
KM_DELIVER = '\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793 (Deliver)'
KM_RETURN = '\u1795\u17d2\u1789\u17be\u178f\u17d2\u179a\u17a1\u1794\u17cb (Return)'
KM_CHECK = '\u178f\u17d2\u179a\u17bd\u178f\u1796\u17b7\u1793\u17b7\u178f\u17d2\u1799 (Check)'

NEW_MAP = f"""PENDING_REMARK_MAP = {{
    '306': '{KM_DELIVER}',
    '309': '{KM_DELIVER}',
    '311': '{KM_DELIVER}',
    '210': '{KM_CHECK}',
    '300': '{KM_CHECK}',
    '302': '{KM_CHECK}',
    '310': '{KM_CHECK}',
    '500': '{KM_RETURN}',
}}"""

for rel in ['generate_report.py', 'push_bot/generate_report.py']:
    path = os.path.join(HERE, rel)
    if not os.path.exists(path):
        continue
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace PENDING_REMARK_MAP
    start = content.find('PENDING_REMARK_MAP = {')
    if start == -1:
        print(f'SKIP {rel}')
        continue
    end = content.find('}', start) + 1
    content = content[:start] + NEW_MAP + content[end:]

    # Replace _pending_next_action inline function
    pattern = r"(            def _pending_next_action\(row\):.*?return '')"
    repl = f"""            def _pending_next_action(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                if sc in ('306', '309', '311'):
                    return '{KM_DELIVER}'
                elif sc in ('300', '302', '210', '310'):
                    return '{KM_CHECK}'
                elif sc in ('500',):
                    return '{KM_RETURN}'
                return ''"""
    content = re.sub(pattern, repl, content, flags=re.DOTALL)

    # Also fix get_next_action if it exists
    pattern2 = r"(def get_next_action\(status_code\):.*?return '')"
    repl2 = f"""def get_next_action(status_code):
    sc = str(status_code).strip()
    if sc in ('306', '309', '311'):
        return '{KM_DELIVER}'
    elif sc in ('300', '302', '210', '310'):
        return '{KM_CHECK}'
    elif sc in ('500',):
        return '{KM_RETURN}'
    return ''"""
    content = re.sub(pattern2, repl2, content, flags=re.DOTALL)

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'DONE: {rel}')

"""Add ACTION and NEXT_STEP columns to the Delivery report."""
import os, re

HERE = os.path.dirname(os.path.abspath(__file__))

# Delivery action mapping
# Code -> (action, next_step)
DELIVERY_ACTION_MAP_CODE = """
DELIVERY_ACTION_MAP = {
    '400': ('\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793', '\u1794\u17c2\u1784\u1785\u17c2\u1780\u17a2\u17d2\u1793\u1780\u178a\u17b9\u1780'),
    '401': ('\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793', '\u178a\u17b9\u1780\u1787\u17bc\u1793'),
    '402': ('\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793', '\u178a\u17b9\u1780\u1787\u17bc\u1793\u17a1\u17be\u1784\u179c\u17b7\u1789'),
    '420': ('\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793', '\u1787\u17bc\u1793\u178a\u17c6\u178e\u17b9\u1784\u1797\u17d2\u1789\u17c0\u179c'),
    '430': ('\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793', '\u1791\u17b6\u1780\u17cb\u1791\u1784\u1797\u17d2\u1789\u17c0\u179c'),
    '460': ('\u178f\u17d2\u179a\u17a1\u1794\u17cb', '\u1795\u17d2\u1789\u17be\u178f\u17d2\u179a\u17a1\u1794\u17cb'),
    '470': ('\u178f\u17d2\u179a\u17a1\u1794\u17cb', '\u1794\u1793\u17d2\u178f\u178f\u17d2\u179a\u17a1\u1794\u17cb'),
    '471': ('\u1796\u17b7\u1793\u17b7\u178f\u17d2\u1799', '\u1796\u17b7\u1793\u17b7\u178f\u17d2\u1799'),
    '472': ('\u1796\u17b7\u1793\u17b7\u178f\u17d2\u1799', '\u178a\u17c4\u17c7\u179f\u17d2\u179a\u17b6\u1799'),
    '480': ('\u1796\u17b7\u1793\u17b7\u178f\u17d2\u1799', '\u1780\u17c2\u17a2\u17b6\u179f\u1799\u178a\u17d2\u178b\u17b6\u1793'),
}
"""

for rel in ['generate_report.py', 'push_bot/generate_report.py']:
    path = os.path.join(HERE, rel)
    if not os.path.exists(path):
        continue
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Add DELIVERY_ACTION_MAP after PENDING_REMARK_MAP
    if 'DELIVERY_ACTION_MAP' not in content:
        insert_pos = content.find('PENDING_REMARK_MAP = {')
        # Find end of PENDING_REMARK_MAP
        end_pos = content.find('}', insert_pos) + 1
        content = content[:end_pos] + '\n' + DELIVERY_ACTION_MAP_CODE + content[end_pos:]

    # 2. Update Delivery columns to include ACTION and NEXT_STEP
    old_delivery_cols = "'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER']"
    new_delivery_cols = "'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER', 'ACTION', 'NEXT_STEP']"
    content = content.replace(old_delivery_cols, new_delivery_cols)

    # 3. Add ACTION and NEXT_STEP population for Delivery type
    # Find where Pending remark is added and add similar for Delivery before it
    marker = "# Add REMARK column for Pending"
    if marker in content and "# Add ACTION columns for Delivery" not in content:
        delivery_code = """        # Add ACTION columns for Delivery
        if rn == 'Delivery' and 'STATUS_CODE' in df_t.columns:
            def _delivery_action(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return DELIVERY_ACTION_MAP.get(sc, ('', ''))[0]
            def _delivery_next_step(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return DELIVERY_ACTION_MAP.get(sc, ('', ''))[1]
            df_t['ACTION'] = df_t.apply(_delivery_action, axis=1)
            df_t['NEXT_STEP'] = df_t.apply(_delivery_next_step, axis=1)
"""
        content = content.replace(marker, delivery_code + "        " + marker)

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'DONE: {rel}')

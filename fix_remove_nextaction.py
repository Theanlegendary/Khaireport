import os

for rel in ['generate_report.py', 'push_bot/generate_report.py']:
    path = os.path.join(r'c:\Users\DELL\Desktop\daily_push', rel)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Remove NEXT_ACTION from Pending columns list
    content = content.replace(
        "'Pending':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'NEXT_ACTION', 'REMARK']",
        "'Pending':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK']"
    )
    content = content.replace(
        "'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION']",
        "'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK']"
    )
    content = content.replace(
        "'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'NEXT_ACTION', 'REMARK']",
        "'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK']"
    )

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'DONE: {rel}')

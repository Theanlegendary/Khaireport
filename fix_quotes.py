import os

path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'generate_report.py')
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace smart quotes with normal quotes
content = content.replace('\u2018', "'")
content = content.replace('\u2019', "'")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Fixed smart quotes')

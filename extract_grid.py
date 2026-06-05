import json, sys

with open(r'c:\Users\gagal\Desktop\Faculdade\IA\NNmodel.ipynb', encoding='utf-8') as f:
    nb = json.load(f)

for i, cell in enumerate(nb['cells']):
    if cell['cell_type'] != 'code': continue
    src = ''.join(cell['source'])
    if 'grid_results' not in src.lower() and 'grid_df' not in src.lower():
        continue
    outputs = cell.get('outputs', [])
    if not outputs: continue
    print(f'=== Cell {i} ===')
    for out in outputs:
        txt = ''.join(out.get('text', out.get('data', {}).get('text/plain', [])))
        sys.stdout.buffer.write(txt.encode('utf-8'))
    print()

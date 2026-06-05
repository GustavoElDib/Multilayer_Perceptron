import json, sys

with open(r'c:\Users\gagal\Desktop\Faculdade\IA\NNmodel.ipynb', encoding='utf-8') as f:
    nb = json.load(f)

keywords = ['grid_results','tempo_s','Fold','val_acc','confus','Acur','fold_df','test_acc','Acurácia']

for i, cell in enumerate(nb['cells']):
    if cell['cell_type'] != 'code': continue
    src = ''.join(cell['source'])
    if any(k.lower() in src.lower() for k in keywords):
        outputs = cell.get('outputs', [])
        if not outputs: continue
        print(f'=== Cell {i} ===', flush=True)
        for out in outputs:
            txt = ''.join(out.get('text', out.get('data', {}).get('text/plain', [])))
            # skip long training logs, keep summary lines
            lines = txt.split('\n')
            for line in lines:
                if any(k in line for k in ['Tempo','Tabela','Melhores','Acur','val_acc',
                                            'Fold','fold','Acur', 'test_acc', 'Média','Desvio',
                                            'Parada','hidden','tempo']):
                    sys.stdout.buffer.write((line + '\n').encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

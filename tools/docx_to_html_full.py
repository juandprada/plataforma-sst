import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

paths = [
    r'a:\seagate\sst\1.PLANEAR\12.MATRIZ DE COMUNICACION.docx',
    r'a:\seagate\sst\1.PLANEAR\13.MATRIZ LEGAL.docx',
    r'a:\seagate\sst\1.PLANEAR\14.FORMATO DE SELECCION DE PROVEEDOR.docx',
    r'a:\seagate\sst\1.PLANEAR\15.FORMATO DE VERIFICACION DE CONTRATISTAS.docx',
    r'a:\seagate\sst\2.HACER\5.REGISTRO DE GESTION DE CAMBIO.docx',
    r'a:\seagate\sst\2.HACER\6.FORMATO DE INVESTIGACION INTERNO.docx',
    r'a:\seagate\sst\3.VERIFICAR\1.INDICADORES DE GESTION.docx',
    r'a:\seagate\sst\3.VERIFICAR\2.ACTA DE REVISION.docx',
    r'a:\seagate\sst\3.VERIFICAR\3.INFORME AUDITORIA INTERNA.docx',
    r'a:\seagate\sst\4.ACTUAR\1.ACCIONES CORRECTIVAS Y PREVENTIVAS.docx',
    r'a:\seagate\sst\4.ACTUAR\2.PROGRAMA DE MEDICINA PREVENTIVA.docx'
]

output_dir = r'a:\seagate\sst\plataforma-sst\tools\html_output'
os.makedirs(output_dir, exist_ok=True)

for path in paths:
    filename = os.path.basename(path).replace('.docx', '.html')
    out_path = os.path.join(output_dir, filename)
    print(f"Processing {path} -> {out_path}")
    
    html = []
    try:
        doc = docx.Document(path)
        for i, table in enumerate(doc.tables):
            html.append('<table class="doc-tabla">')
            for r_idx, row in enumerate(table.rows):
                if r_idx == 0:
                    html.append('  <thead><tr>')
                    tag = 'th'
                else:
                    if r_idx == 1:
                        html.append('  <tbody>')
                    html.append('  <tr>')
                    tag = 'td'
                    
                seen_cells = set()
                for cell in row.cells:
                    if cell in seen_cells: continue
                    seen_cells.add(cell)
                    text = cell.text.strip().replace('\n', '<br>')
                    html.append(f'    <{tag}>{text}</{tag}>')
                    
                if r_idx == 0:
                    html.append('  </tr></thead>')
                else:
                    html.append('  </tr>')
            if len(table.rows) > 0:
                html.append('  </tbody>')
            html.append('</table><br>')
            
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(html))
    except Exception as e:
        print(f"Error reading {path}: {e}")

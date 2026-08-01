import docx
import sys
import os

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

for path in paths:
    print(f"\n--- {path} ---")
    try:
        doc = docx.Document(path)
        for i, table in enumerate(doc.tables):
            if len(table.rows) > 0:
                header = [cell.text.strip().replace('\n', ' ') for cell in table.rows[0].cells]
                # Removing duplicates due to merged cells
                clean_header = []
                for h in header:
                    if h not in clean_header or h == '':
                        clean_header.append(h)
                print(f"Table {i} header: {clean_header}")
    except Exception as e:
        print(f"Error reading {path}: {e}")

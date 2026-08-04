import json
import time

with open('data/gtc45.json', encoding='utf-8') as f:
    data_storage = json.load(f)
mapeo = {x['panel_id']: x['nc'] for x in data_storage['mapeo_peligros']}

results_path = 'a:/seagate/sst/reportes/gtc45/gtc45results.json'
with open(results_path, encoding='utf-8') as f:
    results = json.load(f)

tabla_nd = {'Muy Alto': 10, 'Alto': 6, 'Medio': 2, 'Bajo': 0, 'No Aplica': 0}
tabla_ne = {'Continua': 4, 'Frecuente': 3, 'Ocasional': 2, 'Esporadica': 1, 'No Aplica': 0}

for item in results['Data']:
    name = item['nombre_empresa']
    # If it's one of the ones we touched (or all of them actually)
    changed = False
    for k in list(item.keys()):
        if k.endswith('_deficiencia'):
            prefix = k.replace('_deficiencia', '')
            def_val = item[k]
            if def_val in ['No Aplica', 'Bajo']: continue
            exp_key = f"{prefix}_exposicion"
            exp_val = item.get(exp_key, 'No Aplica')
            
            nc = mapeo.get(prefix, 0)
            nd = tabla_nd.get(def_val, 0)
            ne = tabla_ne.get(exp_val, 0)
            nr = nd * ne * nc
            
            if nr >= 150: # NO ACEPTABLE or NO ACEPTABLE O ACEPTABLE CON CONTROL
                print(f"Fixing {name} -> {prefix} (NR={nr}) to Bajo")
                item[k] = "Bajo"
                changed = True
    
    if changed:
        # Update timestamp so gtc45_to_json.py picks it up!
        item['HappendAt'] = f"/Date({int(time.time()*1000)})/"

with open(results_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False)

print("gtc45results.json fixed.")

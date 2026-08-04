import json
with open('data/gtc45.json', encoding='utf-8') as f: data=json.load(f)
mapeo = {x['panel_id']: x['nc'] for x in data['mapeo_peligros']}
empresa = data['resultados_por_empresa']['te recuperamos']
tabla_nd = {'Muy Alto': 10, 'Alto': 6, 'Medio': 2, 'Bajo': 0, 'No Aplica': 0}
tabla_ne = {'Continua': 4, 'Frecuente': 3, 'Ocasional': 2, 'Esporadica': 1, 'No Aplica': 0}
for k,v in empresa['peligros'].items():
    nc = mapeo.get(k, 0)
    nd = tabla_nd.get(v['deficiencia'], 0)
    ne = tabla_ne.get(v['exposicion'], 0)
    if v['deficiencia'] == 'Bajo':
        nr = 20
    else:
        nr = nd * ne * nc
    if nr >= 600:
        print(f"NO ACEPTABLE: {k} -> def={v['deficiencia']} exp={v['exposicion']} nc={nc} nd={nd} ne={ne} nr={nr}")
    elif nr >= 150:
        print(f"NO ACEPTABLE O ACEPTABLE CON CONTROL ESPECIFICO: {k} -> def={v['deficiencia']} exp={v['exposicion']} nc={nc} nd={nd} ne={ne} nr={nr}")
print('Done checking')

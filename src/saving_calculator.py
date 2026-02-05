import pandas as pd
from math import ceil, radians, cos, sin, sqrt, atan2
from datetime import timedelta
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time
import ssl
import unicodedata
import json
import os

# Ignorar verificação SSL
ssl._create_default_https_context = ssl._create_unverified_context

# Valores por transportadora
valores_transportadoras = {
    'TRANSPORTADORA_A': {'valor_por_plt': 55.00, 'transbordo': 2400.00, 'diaria': 0.00},
    'TRANSPORTADORA_B': {'valor_por_plt': 33.00, 'transbordo': 1900.00, 'diaria': 0.00},
    'TRANSPORTADORA_C': {'valor_por_plt': 32.00, 'transbordo': 0.00, 'diaria': 0.00},
    'TRANSPORTADORA_D': {'valor_por_plt': 0.00, 'transbordo': 0.00, 'diaria': 900.00},
    'TRANSPORTADORA_E': {'valor_por_plt': 0.00, 'transbordo': 0.00, 'diaria': 750.00},
    'TRANSPORTADORA_F': {'valor_por_plt': 80.00, 'transbordo': 0.00, 'diaria': 0.00}
}

# Coordenadas fixas para cidades problemáticas
coordenadas_fixas = {
    "SAO PAULO, SP": (-23.5505, -46.6333),
    "CAMPINAS, SP": (-22.9056, -47.0608),
    "CURITIBA, PR": (-25.4284, -49.2733),
    "RIO DE JANEIRO, RJ": (-22.9068, -43.1729),
    # Adicione outras cidades conforme necessário
}

# Função para normalizar nomes
def normalize_city(city):
    if not isinstance(city, str):
        return ""
    return ''.join(c for c in unicodedata.normalize('NFD', city) if unicodedata.category(c) != 'Mn').upper().strip()

# Geolocalizador
geolocator = Nominatim(user_agent="saving_calculator")

# Cache persistente
cache_file = 'coord_cache.json'
if os.path.exists(cache_file):
    with open(cache_file, 'r') as f:
        coord_cache = json.load(f)
else:
    coord_cache = {}

def save_cache():
    with open(cache_file, 'w') as f:
        json.dump(coord_cache, f)

def get_coordinates(city_name, uf=None, retries=3):
    city_name = normalize_city(city_name)
    uf = normalize_city(uf) if uf else ""
    query = f"{city_name}, {uf}" if uf else city_name

    # Verifica coordenadas fixas
    if f"{city_name}, {uf}" in coordenadas_fixas:
        return coordenadas_fixas[f"{city_name}, {uf}"]

    # Verifica cache
    if query in coord_cache:
        return coord_cache[query]

    # Tenta geocodificação
    for _ in range(retries):
        try:
            location = geolocator.geocode(f"{city_name}, {uf}, Brazil")
            if location:
                coord = (location.latitude, location.longitude)
                coord_cache[query] = coord
                save_cache()
                time.sleep(1)
                return coord
        except GeocoderTimedOut:
            time.sleep(1)

    # Fallback: tenta só com cidade
    try:
        location = geolocator.geocode(f"{city_name}, Brazil")
        if location:
            coord = (location.latitude, location.longitude)
            coord_cache[query] = coord
            save_cache()
            return coord
    except GeocoderTimedOut:
        pass

    coord_cache[query] = None
    save_cache()
    return None

def haversine(coord1, coord2):
    R = 6371
    lat1, lon1 = coord1
    lat2, lon2 = coord2
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1))*cos(radians(lat2))*sin(dlon/2)**2
    return R * (2 * atan2(sqrt(a), sqrt(1 - a)))

def calcular_saving(row):
    origem = row['CIDADE ORIGEM']
    destino = row['CIDADE DESTINO']
    uf_origem = row.get('UF ORIGEM', None)
    uf_destino = row.get('UF DESTINO', None)
    transportadora = row['TRANSPORTADORA']

    coleta_dt = pd.to_datetime(row['DATA COLETA'], dayfirst=True, errors='coerce')
    entrega_dt = pd.to_datetime(row['DATA AGENDA'], dayfirst=True, errors='coerce')

    coord_origem = get_coordinates(origem, uf_origem)
    coord_destino = get_coordinates(destino, uf_destino)

    if not coord_origem or not coord_destino:
        return pd.Series([None]*7, index=[
            'dist_rodoviaria_km', 'dias_transito', 'data_limite_coleta',
            'dias_antecipados', 'saving_diaria', 'saving_armazenagem', 'saving_total'
        ])

    dist_linha_reta = haversine(coord_origem, coord_destino)
    dist_rodoviaria = dist_linha_reta * 1.15

    dias_transito = ceil(dist_rodoviaria / 650)
    data_limite_coleta = entrega_dt - timedelta(days=dias_transito)

    delta = data_limite_coleta - coleta_dt
    dias_antecipados = max(0, ceil(delta.total_seconds() / 86400))

    qtd_paletes = 30
    vals = valores_transportadoras.get(transportadora, {'valor_por_plt': 0, 'transbordo': 0, 'diaria': 0})
    valor_por_plt = vals['valor_por_plt']
    transbordo = vals['transbordo']
    diaria_val = vals['diaria']

    # Armazenagem
    saving_armazenagem = 0
    if valor_por_plt > 0 and transbordo > 0:
        saving_armazenagem = transbordo + (qtd_paletes * valor_por_plt)
    elif valor_por_plt > 0 and transbordo == 0:
        saving_armazenagem = qtd_paletes * valor_por_plt
    elif valor_por_plt == 0 and transbordo > 0:
        saving_armazenagem = transbordo

    # Diárias
    saving_diaria = diaria_val * dias_antecipados

    saving_total = saving_armazenagem + saving_diaria

    return pd.Series({
        'dist_rodoviaria_km': round(dist_rodoviaria, 2),
        'dias_transito': dias_transito,
        'data_limite_coleta': data_limite_coleta.strftime('%d/%m/%Y'),
        'dias_antecipados': dias_antecipados,
        'saving_diaria': saving_diaria,
        'saving_armazenagem': saving_armazenagem,
        'saving_total': saving_total
    })

if __name__ == '__main__':
    df = pd.read_excel('ANTECIPADAS.xlsx', sheet_name='ANTECIPADAS')
    df.columns = df.columns.str.strip()
    df['DATA COLETA'] = pd.to_datetime(df['DATA COLETA'], dayfirst=True, errors='raise')
    df['DATA AGENDA'] = pd.to_datetime(df['DATA AGENDA'], dayfirst=True, errors='raise')

    resultados = df.apply(calcular_saving, axis=1)
    df_result = pd.concat([df, resultados], axis=1)

    df_result.to_excel('saving_gerado.xlsx', index=False)
    print("✅ Processo finalizado. Planilha gerada com sucesso!")
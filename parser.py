import requests
import pandas as pd
import sys
import os

ENDPOINT_URL = "https://query.wikidata.org/sparql"

def run_query(sparql_query):
    """Выполняет SPARQL-запрос к Wikidata и возвращает результаты в формате JSON."""
    headers = {"Accept": "application/sparql-results+json"}
    try:
        response = requests.get(ENDPOINT_URL, params={'query': sparql_query}, headers=headers, timeout=30)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Ошибка при запросе к Wikidata: {e}")
        sys.exit(1)
    return response.json()

# Чтение списка городов из файла cities.txt в текущей директории
filename = "cities.txt"
if not os.path.isfile(filename):
    print(f"Файл не найден: {filename}")
    sys.exit(1)

with open(filename, 'r', encoding='utf-8') as f:
    cities = [line.strip() for line in f if line.strip()]

if not cities:
    print("Список городов пуст.")
    sys.exit(0)

# Удаление дубликатов
cities = list(dict.fromkeys(cities))

# Формирование списка значений для SPARQL
values_list = " ".join([f'"{city}"@ru' for city in cities])

# SPARQL-запрос для получения районов и округов
subdivisions_query = f"""
SELECT DISTINCT ?cityLabel ?subdivisionLabel ?districtLabel WHERE {{
  VALUES ?cityName {{ {values_list} }}
  ?city rdfs:label ?cityName;
        wdt:P31/wdt:P279* wd:Q515.
  OPTIONAL {{
    ?city wdt:P150 ?sub.
    ?sub rdfs:label ?subdivisionLabel.
    FILTER(LANG(?subdivisionLabel) = "ru")
    OPTIONAL {{
      ?sub wdt:P150 ?district.
      ?district rdfs:label ?districtLabel.
      FILTER(LANG(?districtLabel) = "ru")
    }}
  }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "ru". ?city rdfs:label ?cityLabel. }}
}}
"""

# SPARQL-запрос для получения станций метро
metro_query = f"""
SELECT DISTINCT ?cityLabel ?metroLabel WHERE {{
  VALUES ?cityName {{ {values_list} }}
  ?city rdfs:label ?cityName;
        wdt:P31/wdt:P279* wd:Q515.
  OPTIONAL {{
    ?metro wdt:P31/wdt:P279* wd:Q928830;
           wdt:P131 ?city.
    ?metro rdfs:label ?metroLabel.
    FILTER(LANG(?metroLabel) = "ru")
  }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "ru". ?city rdfs:label ?cityLabel. }}
}}
"""

# Запросы к Wikidata
subdivisions_results = run_query(subdivisions_query)
metro_results = run_query(metro_query)

# Обработка результатов
subdivisions_data = {city: [("Нет данных", "Нет данных")] for city in cities}
for item in subdivisions_results.get('results', {}).get('bindings', []):
    city_name = item['cityLabel']['value']
    subdivision_name = item.get('subdivisionLabel', {}).get('value', "Нет данных")
    district_name = item.get('districtLabel', {}).get('value', "Нет данных")
    if subdivisions_data[city_name] == [("Нет данных", "Нет данных")]:
        subdivisions_data[city_name] = []
    subdivisions_data[city_name].append((subdivision_name, district_name))

metro_data = {city: ["метро нет"] for city in cities}
for item in metro_results.get('results', {}).get('bindings', []):
    city_name = item['cityLabel']['value']
    metro_name = item.get('metroLabel', {}).get('value', "метро нет")
    if metro_data[city_name] == ["метро нет"]:
        metro_data[city_name] = []
    metro_data[city_name].append(metro_name)

# Создание Excel-файла с результатами
output_file = "city_parser.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    for city in cities:
        subdivisions = subdivisions_data.get(city, [("Нет данных", "Нет данных")])
        districts, microdistricts = zip(*subdivisions)
        metro_stations = metro_data.get(city, ["метро нет"])
        
        max_length = max(len(districts), len(microdistricts), len(metro_stations))
        districts = list(districts) + ["Нет данных"] * (max_length - len(districts))
        microdistricts = list(microdistricts) + ["Нет данных"] * (max_length - len(microdistricts))
        metro_stations += ["метро нет"] * (max_length - len(metro_stations))
        
        df = pd.DataFrame({
            "Город": [city] * max_length,
            "Район": districts,
            "Округ/Микрорайон": microdistricts,
            "Метро": metro_stations
        })
        # Ограничиваем имя листа до 31 символа и убираем недопустимые символы
        sheet_name = city[:31].replace(":", "").replace("/", "")
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Файл сохранен: {output_file}")

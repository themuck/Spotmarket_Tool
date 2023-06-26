import requests
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import Workbook

def get_json_data(api_url):
    start_time = datetime.datetime.now()
    response = requests.get(api_url)
    end_time = datetime.datetime.now()
    duration = (end_time - start_time).total_seconds() * 1000
    
    if response.status_code == 200:
        json_data = response.json()
        return json_data, duration
    else:
        print("Fehler beim API-Aufruf.")
        return None, duration

def display_json_data(json_data):
    print("JSON-Daten:")
    headers = ['Zeitpunkt', 'Market Price (Eur/MWh)', 'Local Price (Eur/MWh)']
    print(f"{headers[0]:<20} {headers[1]:<25} {headers[2]:<25}")
    for data_point in json_data.get('data', []):
        start_time = datetime.datetime.fromtimestamp(data_point['start_timestamp'] / 1000).strftime('%d.%m.%Y %H:%M:%S')
        market_price = data_point['marketprice']
        local_price = data_point['localprice']
        print(f"{start_time:<20} {market_price:<25.2f} {local_price:<25.2f}")

def create_excel_file(json_data):
    excel_file = 'energy_prices.xlsx'
    workbook = Workbook()
    worksheet = workbook.active
    headers = ['Zeitpunkt', 'Market Price (Eur/MWh)', 'Local Price (Eur/MWh)']
    worksheet.append(headers)
    for data_point in json_data.get('data', []):
        start_time = datetime.datetime.fromtimestamp(data_point['start_timestamp'] / 1000).strftime('%d.%m.%Y %H:%M:%S')
        market_price = data_point['marketprice']
        local_price = data_point['localprice']
        worksheet.append([start_time, market_price, local_price])
    workbook.save(excel_file)

# Beispiel-API-URL
zip_code = "33829"
api_url = f"https://api.corrently.io/v2.0/gsi/marketdata?zip={zip_code}"

# API-Aufruf und Ausgabe der JSON-Daten
json_data, request_time = get_json_data(api_url)
if json_data:
    display_json_data(json_data)
    create_excel_file(json_data)

    # Extrahieren der relevanten Werte für die Diagrammerstellung
    data_points = json_data.get('data', [])  # Überprüfung, ob das Feld 'data' vorhanden ist
    data_points_sorted = sorted(data_points, key=lambda x: x['start_timestamp'])
    start_times = [datetime.datetime.fromtimestamp(point['start_timestamp'] / 1000) for point in data_points_sorted]
    local_prices = [point['localprice'] for point in data_points_sorted]
    
    # Erstellen des Diagramms
    plt.figure(dpi=1024)
    plt.plot(start_times, local_prices, label= f"Lokal-Preis in {zip_code}")

    # Datenpunkte markieren, die kleiner als 0 sind und den negativen Wert ausgeben
    negative_local_indices = [i for i, price in enumerate(local_prices) if price < 0]

    # Höchsten und niedrigsten Preis finden
    min_price = min(local_prices)
    max_price = max(local_prices)
    min_index = local_prices.index(min_price)
    max_index = local_prices.index(max_price)

    # Höchsten und niedrigsten Preis in der Grafik markieren
    plt.plot(start_times[min_index], min_price, 'go', label='Niedrigster Preis')
    plt.plot(start_times[max_index], max_price, 'ro', label='Höchster Preis')

    print("\nHöchster Preis:")
    formatted_max_time = start_times[max_index].strftime('%d.%m.%Y %H:%M:%S')
    print(f"Zeitpunkt: {formatted_max_time}")
    print(f"Lokal-Preis: {max_price:.2f} Eur/MWh")

    print("\nNiedrigster Preis:")
    formatted_min_time = start_times[min_index].strftime('%d.%m.%Y %H:%M:%S')
    print(f"Zeitpunkt: {formatted_min_time}")
    print(f"Lokal-Preis: {min_price:.2f} Eur/MWh")

    # Diagramm beschriften
    plt.xlabel('Zeit')
    plt.ylabel('Preis (Eur/MWh)')
    plt.title('Strompreise')
    plt.legend(fontsize='6', loc="best")  # Legende mit kleinerer Schriftgröße

    # X-Achse formatieren
    plt.xticks(rotation=45, ha='right')
    plt.gca().xaxis.set_major_locator(mdates.AutoDateLocator())
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d.%m %H:%M'))
    plt.gca().xaxis.set_tick_params(which='both', labelsize=8)

    # Gitter hinzufügen
    plt.grid(True)

    # Diagramm anzeigen
    plt.tight_layout()
    plt.show()

    # Zeit für den API-Request ausgeben
    print(f"\nDauer des API-Requests: {request_time:.2f} ms")

# Importeren requests en pandas
import pandas as pd
import requests

# Instellen API + omzetting werkbare vorm
URL = "https://eonet.gsfc.nasa.gov/api/v3/events/geojson"
response = requests.get(URL)
json_data = response.json()

# Data opslaan in rijen
rows = []
for events in json_data['features']:
    row = {}
    row['event_type'] = events['properties']['categories'][0]['title'] # Soort meteorologisch event
    row['name_event'] = events['properties']['title'] # Naam meteorologisch event
    row['coordinates'] = events['geometry']['coordinates'] # Coördinaten event + link naar locatie Google Maps (copy paste, webbrowser.open werkt niet in Colab!)
    coordinates = row['coordinates']
    coordinates.reverse() # Coördinaten (lat/lon) stonden omgekeerd in jsonfile (lon/lat)
    row['google_maps'] = (f'https://maps.google.com/?q={*coordinates,}')
    rows.append(row)

# Creëren DataFrame
df = pd.DataFrame(rows)
df
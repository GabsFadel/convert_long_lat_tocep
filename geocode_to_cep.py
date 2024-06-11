import pandas as pd
import requests
import time

def get_zipcode_google(lat, lon, api_key):
    try:
        url = f"https://maps.googleapis.com/maps/api/geocode/json?latlng={lat},{lon}&key={api_key}"
        response = requests.get(url)
        response_json = response.json()
        
        if response_json['status'] == 'OK':
            for component in response_json['results'][0]['address_components']:
                if 'postal_code' in component['types']:
                    return component['long_name']
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

print("Script started.")

api_key = 'google_api_key'

# Exemplo: C:\Users\User#01\Desktop\cep_output.xlsx
file_path = r'directory'

# Ler todas as linhas do arquivo Excel
sheet = pd.read_excel(file_path)

# Remover linhas onde latitude ou longitude são NaN
sheet = sheet.dropna(subset=['latitude', 'longitude'])

# Adicionar uma nova coluna para o código postal
sheet['Zipcode'] = None

# Contador de linhas com CEP
lines_with_zipcode = 0

# Processar as informações do código postal para cada linha
for index, row in sheet.iterrows():
    zipcode = get_zipcode_google(row['latitude'], row['longitude'], api_key)
    sheet.at[index, 'Zipcode'] = zipcode
    if zipcode:
        lines_with_zipcode += 1
    print(f"{lines_with_zipcode} linhas já foram criadas com CEP.")
    #time.sleep(0.1)

# Salvar o DataFrame atualizado com as informações do código postal na planilha original
sheet.to_excel(file_path, index=False)

print("Processing complete.")
print(f"Total de linhas com CEP: {lines_with_zipcode}")

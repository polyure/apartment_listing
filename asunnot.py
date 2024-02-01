from geopy.geocoders import Nominatim
import folium
import requests
import pandas as pd
import xlwings as xw

LOCATIONS = '[[14694,5,"00100, Helsinki"],\
            [14695,5,"00120, Helsinki"],\
            [14696,5,"00130, Helsinki"],\
            [14697,5,"00140, Helsinki"],\
            [14698,5,"00150, Helsinki"],\
            [14699,5,"00160, Helsinki"],\
            [14700,5,"00170, Helsinki"],\
            [14701,5,"00180, Helsinki"],\
            [5079889,5,"00220, Helsinki"],\
            [14705,5,"00250, Helsinki"],\
            [14706,5,"00260, Helsinki"],\
            [14709,5,"00290, Helsinki"],\
            [14725,5,"00500, Helsinki"],\
            [14726,5,"00510, Helsinki"],\
            [14728,5,"00530, Helsinki"],\
            [5079937,5,"00540, Helsinki"],\
            [14729,5,"00550, Helsinki"],\
            [14732,5,"00580, Helsinki"]]'
            
URL = "https://asunnot.oikotie.fi/vuokra-asunnot"
API_URL = "https://asunnot.oikotie.fi/api/cards"

PARAMS = {'cardType':101,
          'limit':1000, # cardType 101 for rentals, 100 for sale
          'locations':LOCATIONS,
          'offset':0, 
          'constructionYear[max]':2023,
          'roomCount[]':[2,3,4],
          'price[min]':800, 
          'price[max]':1500, 
          'size[min]':55, 
          'size[max]':250, 
          'sortBy':"published_sort_desc"} 

def get_headers():
    r = requests.get(url=URL)
    for r in r.text.split('\n'):
        if (r[:30] == '<meta name="api-token" content'):
            token = (r[32:-2])
        if (r[:28] == '<meta name="loaded" content='):
            loaded = r[29:-2]
        if (r[:26] == '<meta name="cuid" content='):
            cuid = r[27:-2]
    headers = {"OTA-cuid":cuid, "OTA-loaded":loaded, "OTA-token":token}
    return headers

def request_data(headers):
    r = requests.get(url=API_URL, params=PARAMS, headers=headers)
    data=r.json()
    return data

def create_datalist(data):
    fields = ["url","rooms","roomConfiguration","price","published","size","latitude","longitude","coordinates","buildingData"] # URL, huoneet, otsikko, hinta, päivämäärä, pinta-ala ja buildingData, jonka sisältö määritellään if-lauseessa
    datalist = []
    for i in data['cards']:
        row = []
        price = 0
        size = 0
        for j in i:
            if j in fields:
                if (j == "buildingData"):
                    row.append(i[j]['address'])
                    row.append(i[j]['district'])
                    row.append(i[j]['city'])
                    row.append(i[j]['year'])
                elif (j == "coordinates"):
                    row.append(i[j]['latitude'])
                    row.append(i[j]['longitude'])
                else:
                    row.append(i[j])
        datalist.append(row)
    return datalist

def create_dataframe(datalist):        
    df = pd.DataFrame(datalist, columns = ['url', 'rooms', 'roomConfiguration', 'price', 'published', 'size', 'address', 'district', 'city', 'buildYear', 'latitude', 'longitude'])
    return df

def create_CSV_sheet(df):

    wb = xw.Book('asunnot.xlsx')
    sheet = wb.sheets['CSV']

    sheet.clear_contents()
    sheet['A1'].options(index=False, header=False).value = df

    wb.save("asunnot.xlsx")
    wb.close()

def calculate_persqm(df):
    df['price'] = df['price'].replace(to_replace = "[^0-9]", value = "", regex = True)
    df['price'] = df['price'].apply(pd.to_numeric)
    df['perSquareMetre'] = df['price']/df['size']
    return df

def calculate_quintile(df):
    df['quintile'] = pd.qcut(df['perSquareMetre'], 5, labels=False)
    return df
    
def calculate_mean_rent(df):
    mean_rent = df['price'].mean() # For later implementations
    return mean_rent

def create_map(df):
    app = Nominatim(user_agent='tommy')
    apartment_map = folium.Map([60.1868, 24.933], zoom_start=12) # Center the map on Helsinki
    for i in range(len(df)):
        coords = [df.iloc[i]['latitude'], df.iloc[i]['longitude']]
        nametag = df.iloc[i]['address']
        nametag_url = df.iloc[i]['url']
        nametag_price = df.iloc[i]['price']
        nametag_size = df.iloc[i]['size']
        
        match df.iloc[i]['quintile']:
            case 0:
                colour = 'lightgreen'
            case 1:
                colour = 'green'
            case 2:
                colour = 'orange'
            case 3:
                colour = 'lightred'
            case 4:
                colour = 'red'
            case _:
                colour = 'black'
        
        folium.Marker(location=coords, popup='<a href=' + 
                    nametag_url + '>' + 
                    nametag +  '</a>' + '\n' +
                    str(nametag_price) + ' €/kk\n' +
                    str(nametag_size) + ' m^2',
                    icon=folium.Icon(color=colour)
                    ).add_to(apartment_map)

    apartment_map.save('asunnot.html')
    
headers = get_headers()
data = request_data(headers)
datalist = create_datalist(data)
df = create_dataframe(datalist)
#create_CSV_sheet(df)
df = calculate_persqm(df)
df = calculate_quintile(df)
#mean_rent = calculate_mean_rent(df)
create_map(df)
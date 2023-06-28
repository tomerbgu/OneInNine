import time
import urllib.request
import sys
import os
from tqdm import tqdm
import pandas as pd
from geopy.geocoders import Nominatim


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def calc_cij(etn_hard, etn_soft, lecturer, org, distances):
    language_match = lecturer['language'] == org['language']
    if not language_match:
        return 0

    is_hard = org['is_hard']
    hard_match = etn_hard.loc[lecturer['enthicity']][org['enthicity']]
    soft_match = etn_soft.loc[lecturer['enthicity']][org['enthicity']]
    hardness = 65 * (is_hard * hard_match + (1 - is_hard) * soft_match)
    fti = 10 if org['feature'] in lecturer['features'] else 0
    exp = 5 * lecturer['expi']
    row = distances[(distances['From'] == lecturer['address']) & (distances['To'] == org['address'])].reset_index()
    val = row['val'][0]
    dis = 30 * val - lecturer['mobility']
    return hardness + fti + exp + dis


def find_text_in_list(l, x, s=0):
    # Returns the index of l containing x; s indicates index the search starts from
    n = -1
    for i in range(s, len(l)):
        if l[i] == x:
            n = i
            break
    return n


def query_google_map(lat1, long1, lat2, long2):
    # receives two pairs of origin-destination, designated by longitude-latitude, queries the Google Maps API and returns the result string
    url = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=" + str(lat1) + "," + str(
        long1) + "&destinations=" + str(lat2) + "," + str(
        long2) + "&sensor=false&units=metric&key=AIzaSyBdsBLG56CFLVk_z4tqk6_gKi_O9gR6GWU"
    # print(url)
    f = urllib.request.urlopen(url)
    s = f.read()
    s = s.decode('UTF-8')
    f.close()
    # print(s)
    a = s.split()
    return a


def dist(locs):
    loc_df = pd.DataFrame(columns=['From', 'To', 'Meters', 'Seconds'])
    lon = locs['lon']
    lat = locs['lat']
    station_id = locs['address']

    new_rows = []  # Collect new rows in a list

    # queries each pair of stations and outputs distance and travel duration
    for i in tqdm(range(len(locs))):
        for j in range(len(locs)):
            a = query_google_map(lat[i], lon[i], lat[j], lon[j])
            from_loc = str(station_id[i])
            to_loc = str(station_id[j])

            ind1 = find_text_in_list(a, '"distance"')
            ind2 = find_text_in_list(a, '"value"', ind1)
            distance = a[ind2 + 2]
            ind1 = find_text_in_list(a, '"duration"', ind2)
            ind2 = find_text_in_list(a, '"value"', ind1)
            duration = a[ind2 + 2]
            new_row = {'From': from_loc, 'To': to_loc, 'Meters': distance, 'Seconds': duration}
            new_rows.append(new_row)  # Add the new row to the list

    loc_df = pd.concat([loc_df, pd.DataFrame(new_rows)], ignore_index=True)  # Concatenate all new rows at once
    return loc_df


"""def dist(locs):
    loc_df = pd.DataFrame(columns=['From', 'To', 'Meters', 'Seconds'])
    lon = locs['lon']
    lat = locs['lat']
    station_id = locs['address']

    # queries each pair of stations and outputs distance and travel duration
    for i in tqdm(range(len(locs))):
        for j in range(len(locs)):
            a = query_google_map(lat[i], lon[i], lat[j], lon[j])
            from_loc = str(station_id[i])
            to_loc = str(station_id[j])

            ind1 = find_text_in_list(a, '"distance"')
            ind2 = find_text_in_list(a, '"value"', ind1)
            distance = a[ind2 + 2]
            ind1 = find_text_in_list(a, '"duration"', ind2)
            ind2 = find_text_in_list(a, '"value"', ind1)
            duration = a[ind2 + 2]
            new_row = {'From': from_loc, 'To': to_loc, 'Meters': distance, 'Seconds': duration}
            # Convert the new row to a DataFrame
            new_row_df = pd.DataFrame([new_row])

            # Concatenate the new row DataFrame with loc_df
            loc_df = pd.concat([loc_df, new_row_df], ignore_index=True)
            #loc_df = loc_df.append(new_row, ignore_index=True)
            print(loc_df)
    return loc_df"""


def main():
    data_path = resource_path('data/data.xlsx')
    etn_hard = pd.read_excel(data_path, sheet_name='etn_matrix_hard')
    etn_hard.set_index("Etn", inplace=True)
    etn_soft = pd.read_excel(data_path, sheet_name='etn_matrix_soft')
    etn_soft.set_index("Etn", inplace=True)

    # get language
    lecturers = pd.read_excel(data_path, sheet_name='lecturer_data')
    lecturers['features'] = lecturers['features'].apply(lambda x: x.split(','))
    lecturers['expi'] = lecturers['expi'].astype(int)

    orgs = pd.read_excel(data_path, sheet_name='org_data')
    orgs['is_hard'] = orgs['is_hard'].astype(int)

    locs = pd.DataFrame()
    locs['address'] = pd.concat([lecturers['address'], orgs['address']]).drop_duplicates()
    locs = locs.reset_index()
    locs[['lat', 'lon']] = locs['address'].apply(get_coordinates).apply(pd.Series)
    distances = dist(locs)
    distances.to_csv(resource_path('data/output_dist.csv'), encoding='utf-8-sig')
    distances['val'] = distances['Seconds'].apply(lambda x: 10 if int(x) <= 1800 else 5 if 1800 < int(x) <= 2700 else 1)

    output_table = pd.DataFrame(columns=orgs['id'])

    for _, lecturer in lecturers.iterrows():
        lecturer_values = []  # List to store calculated values for each lecturer

        for _, org in orgs.iterrows():
            value = calc_cij(etn_hard, etn_soft, lecturer, org, distances)  # Calculate the value using my_func
            lecturer_values.append(value)

        output_table.loc[lecturer['id']] = lecturer_values
    output_table.to_csv(resource_path('data/cij_matrix.csv'))


def get_coordinates(city):
    geolocator = Nominatim(user_agent="my_app")
    location = geolocator.geocode(city)
    latitude = location.latitude
    longitude = location.longitude
    return latitude, longitude


if __name__ == '__main__':
    main()

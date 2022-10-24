import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os


def obtener_data_google_sheet(name_archivo, name_sheet, keys):
    gc = gspread.service_account(filename=keys)
    scopes = gc.auth.scopes
    credentials = ServiceAccountCredentials.from_json_keyfile_name(keys, scopes)
    servicio = gspread.authorize(credentials)
    libro = servicio.open(name_archivo)
    hoja = libro.worksheet(name_sheet)
    # get_all_records es más directo para extraer data sin los decimales que en excel están con ','
    listas_ = hoja.get_all_values()
    data = pd.DataFrame(listas_[1:], columns=listas_[0])
    return data


def comma_to_dot(df, var):
    df[var] = df[var].replace('\,', '.', regex=True)
    df[var] = df[var].astype('float')
    return df[var]


# data.replace(dict.fromkeys([5], {r"\,": "."}), inplace=True, regex=True)

def name_columns_normal(columns):
    normal_columns = columns.str.lower(). \
        str.replace('__', '_'). \
        str.replace('í', 'i'). \
        str.replace('ú', 'u'). \
        str.replace('á', 'a'). \
        str.replace('é', 'e'). \
        str.replace('ó', 'o'). \
        str.replace('ñ', 'ni'). \
        str.strip(). \
        str.replace(' ', '_')
    return normal_columns

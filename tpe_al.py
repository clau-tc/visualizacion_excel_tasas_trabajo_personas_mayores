
import xlsxwriter
import matplotlib.pyplot as plt
import pandas as pd
from fun_procesamiento import *

os.chdir('/home/clautc/DataspellProjects/tasas_participacion_economica')

# 1) Obtención data

archivo = 'tpALC_jdgo'
name_sheet = 'datos_planos'
with open('/home/clautc/DataspellProjects/tasas_participacion_economica/keys/keys.txt') as k:
    keys = k.readline()
data = obtener_data_google_sheet(name_archivo=archivo, name_sheet=name_sheet, keys=keys)


# 2) procesar data
# normalizar nombres de columnas
data.columns = name_columns_normal(data.columns)
# formato de separador de decimales
data['dato'] = comma_to_dot(data, 'dato')

# 3) acceder a xlsxwriter

# wb = writer.book
# ws = writer.sheets['data']
wb = xlsxwriter.Workbook('data/data_result.xlsx')

ambossexos = wb.add_worksheet('ambos_sexos')
mujeres = wb.add_worksheet('mujeres')
hombres = wb.add_worksheet('hombres')

n = 1
as_c = 0
m_c = 0
h_c = 0
columnas = ['sexo', 'pais_estandar','anios_estandar'] + data.tasa_de_ocupacion_por_grupo_de_edad.unique().tolist()

data_ = data[['sexo', 'pais_estandar', 'anios_estandar', 'tasa_de_ocupacion_por_grupo_de_edad', 'dato']].pivot(index=['sexo', 'pais_estandar', 'anios_estandar'], columns='tasa_de_ocupacion_por_grupo_de_edad').reset_index()

data_agrupada = data_.groupby(['sexo', 'pais_estandar'])

ubi_as ='A1 F1 K1 P1 A13 F13 K13 P13 A26 F26 K26 P26 A39 F39 K39 P39 A52 F52 K52 P52'.split()
ubi_m = 'A1 F1 K1 P1 A13 F13 K13 P13 A26 F26 K26 P26 A39 F39 K39 P39 A52 F52 K52 P52'.split()
ubi_h = 'A1 F1 K1 P1 A13 F13 K13 P13 A26 F26 K26 P26 A39 F39 K39 P39 A52 F52 K52 P52'.split()

for d in data_agrupada.groups.keys():
    d_table = data_agrupada.get_group(d)
    name = ''.join(list(d))
    name = ''.join(list(filter(lambda c: c.isupper(), name))) + str(n)
    wsd = wb.add_worksheet(name)
    row_max, col_max = d_table.shape
    datos = d_table.values.tolist()
    headers = [{'header': v} for v in columnas]
    wsd.add_table(0, 0, row_max, col_max,
                  {'data': d_table.values.tolist(),
                   'columns': headers})
    chart = wb.add_chart({'type': 'line'})
    chart.add_series({
        'name': "='" + name + "'!$D1",
        'categories': "='" + name + "'!$C$2:" + '$C$' + str(row_max),
        'values': "='" + name + "'!$D$2:" + '$D$' + str(row_max),
        'marker': {'type': 'circle',},
                   # 'color': 'red'},
        'line': {'color': 'red'}
    })
    chart.add_series({
        'name': "='" + name + "'!$E1",
        'categories': "='" + name + "'!$C$2:" + '$C$' + str(row_max),
        'values': "='" + name + "'!$E$2:" + '$E$' + str(row_max),
        'marker': {'type': 'circle',},
                   # 'color': '#0D47A1'},
        'line': {'color': '#0D47A1'}
    })
    chart.add_series({
        'name': "='" + name + "'!$F1",
        'categories': "='" + name + "'!$C$2:" + '$C$' + str(row_max),
        'values': "='" + name + "'!$F$2:" + '$F$' + str(row_max),
        'marker': {'type': 'circle',},
                   # 'color': '#43A047'},
        'line': {'color': '#43A047'}
    })
    chart.add_series({
        'name': "='" + name + "'!$G1",
        'categories': "='" + name + "'!$C$2:" + '$C$' + str(row_max),
        'values': "='" + name + "'!$G$2:" + '$G$' + str(row_max),
        'marker': {'type': 'circle'},
                   # 'color': '#9C27B0'},
        'line': {'color': '#9C27B0'}
    })
    pais = d_table.pais_estandar.unique().tolist()[0]
    # chart.set_chartarea({'fill': {'color': '#9E9E9E'}})
    chart.set_size({'width': 320, 'height': 240})
    chart.set_legend({'position': 'bottom',
                      'font': {'size': 8}})
    chart.set_plotarea({'fill': {'color': 'white'}})
    chart.set_title({'name': pais,
                     'name_font': {'color': 'black',
                                   'size': 10}})
    chart.set_x_axis({'name': 'Periodo de estudio',
                      'name_font': {'size': 8},
                      'num_font': {'size': 8}})
    chart.set_y_axis({'name': 'Tasa por cada 100 personas',
                     'name_font': {'size': 8},
                      'num_font': {'size': 8},
                      # 'minor_unit': 0,
                      # 'major_unit': 100,
                      # 'interval_unit': 10,
                      'visible': True})

    # chart.set_style(12)
    sexo = d_table.sexo.unique().tolist()[0]

    if sexo == 'Ambos sexos':
        ambossexos.insert_chart(ubi_as[as_c], chart)
        as_c += 1
    elif sexo == 'Mujeres':
        mujeres.insert_chart(ubi_m[m_c], chart)
        m_c += 1
    else:
        hombres.insert_chart(ubi_h[h_c], chart)
        h_c += 1

    n += 1

wb.close()



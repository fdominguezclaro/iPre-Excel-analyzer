import os
import sys
import unicodedata
from statistics import mean
from scipy.stats import ttest_ind

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import unicodecsv
import xlrd


#  Funcion no usada
def xls2csv(xls_filename, csv_directory):
    """

    Converts all sheets from an Excel file to CSV files.
    """

    wb = xlrd.open_workbook(xls_filename)
    sheet_names = wb.sheet_names()
    xlsx_name = xls_filename.split('.')[0]

    for sheet in sheet_names:
        sh = wb.sheet_by_name(sheet)
        print(os.getcwd())
        if not os.path.isdir('csv/' + xlsx_name):
            os.makedirs('csv/' + xlsx_name)
        with open('{0}/{1}/{2}.csv'.format(csv_directory, xlsx_name, sheet), "wb") as fh:
            csv_out = unicodecsv.writer(fh, encoding='utf-8')

            for row_number in range(sh.nrows):
                csv_out.writerow(sh.row_values(row_number))


def remove_accents(input_str):
    """
    Convert strings to ASCII
    """
    input_str = input_str.title()
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    only_ascii = nfkd_form.encode('ASCII', 'ignore')
    only_ascii = only_ascii.decode('ASCII')
    return only_ascii


def get_group_columns(dataf):
    """

    Returns list with [['group_name', [index of columns]], []]
    """

    groups_names = []
    count = -1
    item_actual = ['', []]

    for item in dataf.columns.values:

        if count == -1 and 'Unnamed' in item:
            # If first item has no group name
            count += 1
            item_actual[1].append(count)

        elif 'Unnamed' not in item:
            # Agrego el item anterior
            if count != -1:
                groups_names.append(item_actual)
                count += 1
                item_actual = [item, [count]]

            # Creo el nuevo grupo y empiezo la cuenta de columnas
            else:
                item_actual = [item, []]
                count += 1
                item_actual[1].append(count)

        else:
            count += 1
            item_actual[1].append(count)

    groups_names.append(item_actual)

    return groups_names


def read_excel(path):
    df_dict = pd.read_excel(path, sheetname=None, index_col=[0, 1], header=0)
    sheets = [sheet for sheet in df_dict.keys()]
    rdict = {'data': df_dict, 'sheets': sheets}
    return rdict


def clean_frame(df):
    df.replace('x', np.nan, regex=True, inplace=True)
    return df


def merge_groups_mean(data):
    """
    Asocia los grupos por con el mean y agrega esta columna
    Tambien agrega un mean al final
    Tambien selecciona las columnas promediadas para sacar archivos resumidos
    :param data: {'data': df_dict, 'sheets': sheets}
    """
    short_data = data
    for sheet in data['sheets']:
        groups = get_group_columns(data['data'][sheet])
        clean_frame(data['data'][sheet])
        df_short = []
        df_sheet = ''
        first = True
        for group in groups:
            columns_group = [data['data'][sheet].iloc[:, [nro]] for nro in group[1]]
            df = pd.concat(columns_group, axis=1)
            df['{} mean'.format(group[0])] = df.mean(axis=1)
            df.loc['mean'] = df.iloc[1:, :].mean()
            group_mean = df['{} mean'.format(group[0])]

            if first:
                df_sheet = df
                df_short = group_mean
                first = False
            else:
                df_sheet = pd.concat([df_sheet, df], axis=1)
                df_short = pd.concat([df_short, group_mean], axis=1)
        data['data'][sheet] = df_sheet
        short_data['data'][sheet] = df_short

    return data, short_data


def new_files(data, name):
    writer = pd.ExcelWriter('output_{}.xlsx'.format(name), engine='xlsxwriter')
    for sheet in data['sheets']:
        data['data'][sheet].to_excel(writer, sheet)

    workbook = writer.book

    # Add a header format.
    header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                                      'border': 1, 'align': 'center'})
    col_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'align': 'center'})

    for sheet in data['sheets']:
        worksheet = writer.sheets[sheet]
        ncols = len(data['data'][sheet].shape)
        nrows = data['data'][sheet].shape[0]
        if ncols != 1:
            ncols = data['data'][sheet].shape[1]
        worksheet.set_column(0, ncols, 30, col_fmt)
        worksheet.set_row(0, None, header_fmt)
        worksheet.set_row(nrows, None, header_fmt)

    writer.save()


# Funcion no usada
def plot(all_inf):
    for poll in all_inf.keys():
        if poll != 'FC':
            marzo = []
            junio = []
            mini = 0
            maxi = 0
            for item in all_inf[poll].keys():
                mean_m = mean(all_inf[poll][item][0])
                mean_j = mean(all_inf[poll][item][1])
                marzo.append((item, mean_m))
                junio.append((item, mean_j))

                if mini > mean_m:
                    mini = mean_m
                if maxi < mean_m:
                    maxi = mean_m

                if mini > mean_j:
                    mini = mean_j
                if maxi < mean_j:
                    maxi = mean_j

            titles, ys = zip(*marzo)
            width = 0.25

            fig, ax = plt.subplots(figsize=(10, 5))

            items_m, score_m = zip(*marzo)
            items_j, score_j = zip(*junio)

            x_pos = np.arange(len(items_m))

            plt.bar(x_pos, score_m, align='center', width=width, alpha=0.5, color='#EE3224', label=titles[0])
            plt.bar(x_pos + 0 + width, score_j, align='center', width=width, alpha=0.5, color='#F78F1E',
                    label=titles[0])

            plt.xticks(x_pos + width / 2, items_m)
            plt.xticks(rotation=45)
            plt.tick_params(axis='both', which='major', labelsize=6)
            plt.tick_params(axis='both', which='minor', labelsize=6)

            plt.ylabel('Means')

            # Setting the y-axis limit

            plt.ylim(0, maxi + width * 4)
            plt.xlim(-0.5, 5)

            ax.set_title(poll)
            plt.legend(['Marzo', 'Junio'], loc='upper left')
            plt.grid()

            plt.show()

            #
            # plt.xlabel('Means')
            # plt.ylabel('Items')
            # plt.title(poll)
            # plt.legend()
            # plt.tight_layout()
            # plt.show()


def load_poll_info(daf1, daf2):
    """

    Carga toda la informacion en un diccionario para poder hacer consultas por nombre
    :param daf1: dataframe 1
    :param daf2: dataframe 2

    :return dict with database

    """
    db = {}
    for sheet in daf1['sheets']:
        if isinstance(daf1['data'][sheet], pd.Series):
            daf1['data'][sheet] = daf1['data'][sheet].to_frame()

        groups = list(daf1['data'][sheet].keys())
        matrix = daf1['data'][sheet].reset_index().values

        # Limpio el nombre de sheet
        sheet = sheet.split(' - ')[1]

        for d in matrix[1:]:
            # Arreglo el formato de los nombres
            name = remove_accents('{} {}'.format(d[0][1].strip(), d[0][0].strip()))
            if name not in db.keys():
                # Creo el diccionario para esa persona
                db[name] = {'Marzo': {}, 'Junio': {}}

            new_dict = {}
            for i in range(len(groups)):
                group = groups[i]

                # Elimino la palabra mean de las columnas
                if 'mean' in group:
                    new_dict[group[:-5]] = d[i + 1]
                else:
                    new_dict[group] = d[i + 1]

            db[name]['Marzo'][sheet] = new_dict

    for sheet in daf2['sheets']:
        if isinstance(daf2['data'][sheet], pd.Series):
            daf2['data'][sheet] = daf2['data'][sheet].to_frame()

        groups = list(daf2['data'][sheet].keys())
        matrix = daf2['data'][sheet].reset_index().values

        # Limpio el nombre de sheet
        sheet = sheet.split(' - ')[1]

        for d in matrix[1:]:
            name = remove_accents('{} {}'.format(d[0][1], d[0][0]))

            if name not in db.keys():
                # Creo el diccionario para esa persona
                db[name] = {'Marzo': {}, 'Junio': {}}

            new_dict = {}
            for i in range(len(groups)):
                group = groups[i]

                # Elimino la palabra mean de las columnas
                if 'mean' in group:
                    new_dict[group[:-5]] = d[i + 1]

                else:
                    new_dict[group] = d[i + 1]

            db[name]['Junio'][sheet] = new_dict

    return db


def ask_db(db):
    nombre = True

    while nombre:
        print(' \n\n - Q para salir')
        print(' - T para ver toda la base de datos')
        print(' - Nombre alumno para ver cambios\n')

        nombre = str(input('Ingresa Opcion: '))

        if nombre == 'q' or nombre == 'Q':
            sys.exit()

        elif nombre == 'T' or nombre == 't':
            print(db)

        else:

            try:
                keys = []

                for key in db[nombre]['Marzo'].keys():
                    if key not in keys:
                        keys.append(key)
                for key in db[nombre]['Junio'].keys():
                    if key not in keys:
                        keys.append(key)

                print('\nEstadisticas para: {}\n'.format(nombre))

                for key in keys:
                    if key in db[nombre]["Marzo"].keys():
                        print('Marzo:', key, db[nombre]["Marzo"][key])
                    if key in db[nombre]["Junio"].keys():
                        print('Junio:', key, db[nombre]["Junio"][key], '\n')

            except KeyError as err:
                print('Error: {} no registrado en la base de datos'.format(err))


def get_names(name):
    """
    Como la base de datos tiene el formato: Apellido 1 Apellido 2, Nombre. Esta funcion obtiene el nombre en el formato
    correcto "nombre apellido"

    :param name: tuple ('last names, name', 'username')
    :return: str(name lastname)
    """
    nombre = name[0].split(', ')
    nombre = '{0} {1}'.format(nombre[1], nombre[0])
    return nombre


# noinspection PyTypeChecker
def load_grades_info(datab):
    df1 = read_excel('Notas/Detalle Notas -2017-1- Seccion 1 - Alejandra Meneses - Con Flipped.xls')
    df2 = read_excel('Notas/Detalle Notas -2017-1- Seccion 2 - Ana Maria Jorquera - Con Flipped.xls')
    df3 = read_excel('Notas/Notas EDU0330 1° semestre 2017.xls')

    # Tienen distintos formatos por lo tanto manejo cada archivo de manera distinta

    for index, row in df1['data']['NOTAS DEF'].iterrows():

        # Solo si no es fila vacia
        if not isinstance(index[0], int):

            name = remove_accents(get_names(index))

            # Maneja el error en caso de tener segundo apellido
            try:
                datab[name]['Promedio'] = row['Promedio']

            except KeyError:
                # Elimino el ultimo apellido
                new_name = ' '.join(name.split(' ')[:-1])
                datab[new_name]['Promedio'] = row['Promedio']

    for index, row in df2['data']['NOTAS DEF'].iterrows():

        # Solo si no es fila vacia
        if not isinstance(index[0], int):
            name = remove_accents('{} {}'.format(index[1], index[0]))

            # Maneja el error en caso de tener segundo apellido
            try:
                datab[name]['Promedio'] = row['Promedio']

            except KeyError:

                try:
                    # Elimino el ultimo apellido
                    new_name = ' '.join(name.split(' ')[:-1])
                    datab[new_name]['Promedio'] = row['Promedio']

                except KeyError:

                    try:
                        # Elimino el segundo nombre
                        new_name = remove_accents('{} {}'.format(index[1].split(' ')[0], index[0]))
                        datab[new_name]['Promedio'] = row['Promedio']

                    except KeyError:
                        # Elimino el segundo nombre y ultimo apellido
                        new_name = ' '.join(name.split(' ')[:-1])
                        datab[new_name]['Promedio'] = row['Promedio']

    for index, row in df3['data']['NOTAS DEF'].iterrows():

        # Solo si no es fila vacia
        if not isinstance(index[0], int):
            name = remove_accents('{} {}'.format(index[1], index[0]))

            # Maneja el error en caso de tener segundo apellido o segundo nombre
            try:
                datab[name]['Promedio'] = row['Promedio']

            except KeyError:

                try:
                    # Elimino el ultimo apellido
                    new_name = ' '.join(name.split(' ')[:-1])
                    datab[new_name]['Promedio'] = row['Promedio']

                except KeyError:

                    try:
                        # Elimino el segundo nombre
                        new_name = remove_accents('{} {}'.format(index[1].split(' ')[0], index[0]))
                        datab[new_name]['Promedio'] = row['Promedio']

                    except KeyError:
                        # Elimino el segundo nombre y ultimo apellido
                        new_name = ' '.join(name.split(' ')[:-1])
                        datab[new_name]['Promedio'] = row['Promedio']

    return datab


def variaciones(nota_quiebre, datab):
    """

    Funcion que analiza y separa en dos grupos segun la nota de quiebre. Retorna lista de notas de ambas encuestas por
    cada item.

    :param nota_quiebre: int
    :param datab: dict with database
    """
    group_1 = list(
        filter(lambda x: datab[x[0]]['Promedio'] >= nota_quiebre if 'Promedio' in datab[x[0]].keys() else None,
               datab.items()))
    group_1 = {item[0]: item[1] for item in group_1}

    group_2 = list(
        filter(lambda x: datab[x[0]]['Promedio'] < nota_quiebre if 'Promedio' in datab[x[0]].keys() else None,
               datab.items()))
    group_2 = {item[0]: item[1] for item in group_2}

    # Creo un diccionario con diccionarios que tiene [0, 0] para los valores de marzo y junio de cada item

    info_1 = {poll: {item: [[], []] for item in group_1['Maria Jose Alvarez']['Junio'][poll].keys()} for poll in
              group_1['Maria Jose Alvarez']['Junio'].keys()}
    info_2 = {poll: {item: [[], []] for item in group_1['Maria Jose Alvarez']['Junio'][poll].keys()} for poll in
              group_1['Maria Jose Alvarez']['Junio'].keys()}
    allinfo = {poll: {item: [[], []] for item in group_1['Maria Jose Alvarez']['Junio'][poll].keys()} for poll in
               group_1['Maria Jose Alvarez']['Junio'].keys()}

    for name in group_1.keys():
        for poll in group_1[name]['Marzo'].keys():
            for item in group_1[name]['Marzo'][poll].keys():
                if not np.isnan(group_1[name]['Marzo'][poll][item]):
                    info_1[poll][item][0].append(float(group_1[name]['Marzo'][poll][item]))
                    allinfo[poll][item][0].append(float(group_1[name]['Marzo'][poll][item]))

        for poll in group_1[name]['Junio'].keys():
            for item in group_1[name]['Junio'][poll].keys():
                if not np.isnan(group_1[name]['Junio'][poll][item]):
                    info_1[poll][item][1].append(float(group_1[name]['Junio'][poll][item]))
                    allinfo[poll][item][1].append(float(group_1[name]['Junio'][poll][item]))

    for name in group_2.keys():
        for poll in group_2[name]['Marzo'].keys():
            for item in group_2[name]['Marzo'][poll].keys():
                if not np.isnan(group_2[name]['Marzo'][poll][item]):
                    info_2[poll][item][0].append(float(group_2[name]['Marzo'][poll][item]))
                    allinfo[poll][item][0].append(float(group_2[name]['Marzo'][poll][item]))

        for poll in group_2[name]['Junio'].keys():
            for item in group_2[name]['Junio'][poll].keys():
                if not np.isnan(group_2[name]['Junio'][poll][item]):
                    info_2[poll][item][1].append(float(group_2[name]['Junio'][poll][item]))
                    allinfo[poll][item][1].append(float(group_2[name]['Junio'][poll][item]))

    return info_1, info_2, allinfo


def stats_vars(info_1, info_2, nota_quiebre):
    """

    :param info_1: diccionario con listas de notas sobre la nota de quiebre
    :param info_2: diccionario con listas de notas bajo la nota de quiebre

    :param nota_quiebre: int con nota de quiebre

    Analiza las variaciones, imprime los resultados de manera ordenada y retorna diccionarios reduciendo a los
    promedios por cada encuesta.
    """

    print('\n\n-----------\n'
          'Variaciones'
          '\n-----------\n')

    proms_1 = {poll: {item: [[], []] for item in info_1[poll].keys()} for poll in info_1.keys()}
    proms_2 = {poll: {item: [[], []] for item in info_1[poll].keys()} for poll in info_1.keys()}

    # Saco el promedio de cada uno de los items
    for poll in info_1.keys():
        for item in info_1[poll].keys():
            proms_1[poll][item][0] = round(mean(info_1[poll][item][0]), 4) if len(info_1[poll][item][0]) > 0 else 0
            proms_1[poll][item][1] = round(mean(info_1[poll][item][1]), 4) if len(info_1[poll][item][1]) > 0 else 0

    for poll in info_2.keys():
        for item in info_2[poll].keys():
            proms_2[poll][item][0] = round(mean(info_2[poll][item][0]), 4) if len(info_2[poll][item][0]) > 0 else 0
            proms_2[poll][item][1] = round(mean(info_2[poll][item][1]), 4) if len(info_2[poll][item][1]) > 0 else 0

    # Imprimo las estadisticas
    print('\nAlumnos con promedio sobre {}: '.format(nota_quiebre))
    for poll in proms_1.keys():
        aumento = []
        disminuyo = []

        # Esta encuesta fue tomada solo en junio por lo tanto no sirve.
        if poll == 'FC':
            pass
        else:
            for item in proms_1[poll].keys():
                aumento.append(item) if proms_1[poll][item][1] > proms_1[poll][item][0] else disminuyo.append(item)
            print('\nEn encuesta {}:'.format(poll))
            for item in aumento:
                print('--- Aumentaron {0} de {1} -> {2}'.format(item, proms_1[poll][item][0], proms_1[poll][item][1]))
            print()
            for item in disminuyo:
                print('--- Disminuyeron {0} de {1} -> {2}'.format(item, proms_1[poll][item][0], proms_1[poll][item][1]))

    print('\nAlumnos con promedio bajo {}: '.format(nota_quiebre))
    for poll in proms_2.keys():
        aumento = []
        disminuyo = []
        # Esta encuesta fue tomada solo en junio por lo tanto no sirve.
        if poll == 'FC':
            pass
        else:
            for item in proms_2[poll].keys():
                aumento.append(item) if proms_2[poll][item][1] > proms_2[poll][item][0] else disminuyo.append(item)
            print('\nEn encuesta {}:'.format(poll))
            for item in aumento:
                print('--- Aumentaron {0} de {1} -> {2}'.format(item, proms_2[poll][item][0], proms_2[poll][item][1]))
            print()
            for item in disminuyo:
                print('--- Disminuyeron {0} de {1} -> {2}'.format(item, proms_2[poll][item][0], proms_2[poll][item][1]))

    return proms_1, proms_2


def ttests(info_1, info_2):
    """
    :parameter info_1 : Diccionario con todas las encuestas y sus items con una lista de todos los resultados para
    alumnos sobre la media
    :parameter info_2 : Diccionario con todas las encuestas y sus items con una lista de todos los resultados para
    alumnos bajo la media

    Saco los t test para los resultados de junio para cada factor para los dos grupos de alumnos (separados por
    nota quiebre)
    """
    print('\n-----------\n'
          'T-Tests\n'
          '-----------\n')

    for poll in info_1.keys():
        if poll != 'FC':
            print('\nEncuesta:', poll)

            for item in info_1[poll].keys():
                # Use scipy.stats.ttest_ind.
                g1 = info_1[poll][item][1]
                g2 = info_2[poll][item][1]

                t, p = ttest_ind(g1, g2, nan_policy='omit', equal_var=False)

                print('En {0}: t = {1} y p = {2}'.format(item, t, p))


# Leo los archivos y agrego las columnas importantes
dataframe1 = read_excel('Encuestas/MARZO_2017.xlsx')
dataframe1, short_df1 = merge_groups_mean(dataframe1)

dataframe2 = read_excel('Encuestas/JUNIO_2017.xlsx')
dataframe2, short_df2 = merge_groups_mean(dataframe2)

"""
Como los archivos ya estan creados, entonces las 4 lineas siguientes estan comentadas
"""
# new_files(df1, 'MARZO')
# new_files(df2, 'JUNIO')
# new_files(short_df1, 'MARZO_short')
# new_files(short_df2, 'JUNIO_short')

"""
Cargo la informacion a la base de datos
"""
database = load_poll_info(short_df1, short_df2)
database = load_grades_info(database)

"""
Elijo la nota de quibre. El promedio en las secciones fue un 5.3 por lo tanto elijo ese valor por defecto
"""

quiebre = 5.3
info1, info2, all_info = variaciones(quiebre, database)

"""
Imprimo los cambios que cada grupo tuvo y retorno diccionarios con promedios
"""
proms1, proms2 = stats_vars(info1, info2, quiebre)

"""
Imprimo los resultados de los T-Tests
"""

ttests(info1, info2)

"""
Si se quiere preguntar por cada alumno entonces descomentar la siguiente linea
"""
# ask_db(database)

"""
Saco algunos graficos
"""
# plot(allinfo)

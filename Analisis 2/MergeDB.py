import pandas as pd


def read_excel(path):
    dframe = pd.read_excel(path, sheetname=None, index_col=[0, 1], header=0, parse_dates=False)
    return dframe


def concatenate(dfs):
    """
    This fuction concatenates dfs frames with the same headers.

    :rtype: pandas data frame
    :param dfs: pandas data frames
    :return: Complete DF
    """

    cmplt = pd.concat(dfs)
    return cmplt


def get_dfs(dfs):
    """
    Get a list with data frames of each sheet.

    :param dfs: Complete excel file
    :return: list with data frame
    """

    sep_dfs = [dfs[sheet] for sheet in dfs.keys()]
    return sep_dfs


def df_to_excel(dfs):
    """
    Saves a pandas data frame to excel.

    :param dfs: pandas data frame
    """
    writer = pd.ExcelWriter('Encuestas_merged.xlsx')
    dfs.to_excel(writer, 'Sheet1', na_rep='')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.set_zoom(90)
    total_fmt = workbook.add_format({'align': 'center'})
    total_fmt.set_align('vcenter')

    worksheet.set_column('A:B', 20, total_fmt)
    worksheet.set_column('C:T', 5, total_fmt)
    worksheet.set_column('U:V', 20, total_fmt)
    worksheet.set_column('W:AL', 5, total_fmt)
    worksheet.set_column('AM:AT', 20, total_fmt)

    writer.save()


"""
El programa no permite que existan celdas con formatos como fechas, etc...
"""

df = read_excel('Encuestas.xlsx')
sep_df = get_dfs(df)
complete = concatenate(sep_df)
df_to_excel(complete)

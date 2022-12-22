import pandas as pd
import os
import seaborn as sns
import numpy as np
import re
import math
import datetime

# Disabling warnings
import warnings
warnings.filterwarnings("ignore")

# To see all columns in DF
pd.set_option('display.max_columns', None)

# Getting files in folder (cause I need to work with several excel files)
files_list = os.listdir('C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/')
# Cleaning files list to only have Excel files
files = []
for i in files_list:
    if i[-5:] == '.xlsx' and i[:2] != '~$':  # For not to take the temporary files
        files.append(i)


# Creating 2 Dataframes: one with name and ID of client AND other with the data
# First Dataframe:
# "For" to create the all-in-one database and the ID database
nofat_data = pd.DataFrame()
id_col = []
pac_id = 0
for j in files:
    # Import only a section of the Excel file
    print(j)
    paciente = pd.read_excel(
        io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
        sheet_name='BITACORA',
        usecols='C:AF',
        skiprows=21,
        nrows=14
    )

    # Drop unusefull column
    paciente.drop('Unnamed: 3', inplace=True, axis=1)

    # Drop unusefull row
    paciente.drop(1, inplace=True, axis=0)

    # Swapping Cols and Rows
    paciente = paciente.T
    paciente.reset_index(drop=True, inplace=True)

    # Columns names
    new_header = paciente.iloc[0]  # grab the first row for the header
    paciente = paciente[1:]  # take the data less the header row
    paciente.columns = new_header  # set the header row as the df header

    # Adding Day column as first column
    days = []
    for i in range(1, len(paciente)+1):
        days.append(i)
    paciente.insert(0, 'Day', days)

    # Adding patient ID as first column
    pac_id = pac_id + 1
    id = []
    for i in range(1, len(paciente)+1):
        id.append(pac_id)
    paciente.insert(0, 'ID', id)

    # Concat with nofat_data
    nofat_data = pd.concat([nofat_data, paciente])

    # Creating columns for id_data DF
    id_col.append(pac_id)

# Reset index of nofat_data
nofat_data.reset_index(drop=True, inplace=True)

# Adding Glucosa column
glucosa_col = []
for j in files:
    # Import only a section of the Excel file
    print(j)
    glucosa = pd.read_excel(
        io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
        sheet_name='BITACORA',
        usecols='E:AF',
        skiprows=62,
        nrows=1
    )

    # Swapping Cols and Rows
    glucosa = glucosa.T

    # Creating glucosa_col
    for i in range(0, len(glucosa)):
        glucosa_col.append(glucosa.iloc[i, 0])

nofat_data['glucosa'] = glucosa_col

# Correcting data
# Changing data types and if it can't be change, convert it to nan value
def convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        return pd.np.nan

# Applying function to all DF, value by value
for col in range(2, len(nofat_data.columns)):  # The first two columns don't need to be change
    print(nofat_data.columns[col])
    for row in range(0, len(nofat_data)):  # Trying to change every value, one column at time
        nofat_data.iloc[row, col] = convert_to_float(nofat_data.iloc[row, col])

# Rename columns
new_cols = {'ID': 'id', 'Day': 'day', 'PESO': 'peso', 'IMC': 'imc', '% GRASA': 'porc_grasa', 'MASA MUSCULAR KG': 'masa_muscular_kg', '% AGUA': 'porc_agua',
       'GRASA VISCERAL ': 'grasa_visceral', 'KG HUESO': 'kg_hueso', 'METABOLISMO': 'metabolismo', 'PROTEINA %': 'porc_proteina',
       'OBESIDAD %': 'porc_obesidad', 'EDAD CORPORAL': 'edad_corporal', 'PESO SIN GRASA -LBM': 'peso_sin_grasa_lbm', '% MASA MUSCULAR': 'porc_masa_muscular'}

nofat_data.rename(columns=new_cols, inplace=True)

# Cleaning null. Most of them are from people who finish the program before 28 days
nofat_data.dropna(inplace=True)
nofat_data.reset_index(drop=True, inplace=True)

# Checking and correcting strange values
# peso
sns.scatterplot(
    x=nofat_data.index,
    y='peso',
    data=nofat_data
)

# porc_grasa
sns.scatterplot(
    x=nofat_data.index,
    y='porc_grasa',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['porc_grasa'][i] > 1.0:
        nofat_data['porc_grasa'][i] = nofat_data['porc_grasa'][i] / 100

# masa_muscular_kg
sns.scatterplot(
    x=nofat_data.index,
    y='masa_muscular_kg',
    data=nofat_data
)

# porc_agua
sns.scatterplot(
    x=nofat_data.index,
    y='porc_agua',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['porc_agua'][i] > 1.0:
        nofat_data['porc_agua'][i] = nofat_data['porc_agua'][i] / 100

# grasa_visceral
sns.scatterplot(
    x=nofat_data.index,
    y='grasa_visceral',
    data=nofat_data
)

# kg_hueso
sns.scatterplot(
    x=nofat_data.index,
    y='kg_hueso',
    data=nofat_data
)

# metabolismo
sns.scatterplot(
    x=nofat_data.index,
    y='metabolismo',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['metabolismo'][i] > 4000:
        nofat_data['metabolismo'][i] = pd.np.nan

# porc_proteina
sns.scatterplot(
    x=nofat_data.index,
    y='porc_proteina',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['porc_proteina'][i] > 1.0:
        nofat_data['porc_proteina'][i] = nofat_data['porc_proteina'][i] / 100

# porc_obesidad
sns.scatterplot(
    x=nofat_data.index,
    y='porc_obesidad',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['porc_obesidad'][i] > 1.0:
        nofat_data['porc_obesidad'][i] = nofat_data['porc_obesidad'][i] / 100

for i in range(0, len(nofat_data)):
    if nofat_data['porc_obesidad'][i] < -0.1:
        nofat_data['porc_obesidad'][i] = pd.np.nan

# edad_corporal
sns.scatterplot(
    x=nofat_data.index,
    y='edad_corporal',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['edad_corporal'][i] < 10:
        nofat_data['edad_corporal'][i] = pd.np.nan

# peso_sin_grasa_lbm
sns.scatterplot(
    x=nofat_data.index,
    y='peso_sin_grasa_lbm',
    data=nofat_data
)

for i in range(0, len(nofat_data)):
    if nofat_data['peso_sin_grasa_lbm'][i] < 30:
        nofat_data['peso_sin_grasa_lbm'][i] = pd.np.nan

# porc_masa_muscular
sns.scatterplot(
    x=nofat_data.index,
    y='porc_masa_muscular',
    data=nofat_data
)

# glucosa
sns.scatterplot(
    x=nofat_data.index,
    y='glucosa',
    data=nofat_data
)


# Extracting more data from the excel files: edad, altura y sexo.
# For this, I will create a new DF and then join it with the nofat_data DF.

# sexo column (indicated by the owner of the data)
sexo = [
    'F',
    'F',
    'F',
    'F',
    'F',
    'M',
    'F',
    'F',
    'M',
    'F',
    'M',
    'M',
    'F',
    'F',
    'M',
    'M',
    'M',
    'F',
    'M',
    'F',
    'F',
    'F',
    'M',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'F',
    'M',
    'M',
    'F',
    'M',
    'M',
    'M',
    'M',
    'M',
    'F',
    'F'
]

edad_col = []
alt_col = []
for j in files:
    # Import only a section of the Excel file
    print(j)
    paciente = pd.read_excel(
        io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
        sheet_name='BITACORA',
        usecols='AJ:AK',
        skiprows=2,
        nrows=3
    )
    edad_col.append(re.sub('[a-zA-Zñ]', '', str(paciente.iloc[0, 1])).split(' ', 1)[0])  # Keep only numbers and decimal point (comma) and delete everything after white space
    alt_col.append(re.sub('[a-zA-Zñ]', '', str(paciente.iloc[1, 1])).replace(',', '.').split(' ', 1)[0])  # Keep only numbers and decimal point (comma), replace comma with point and delete everything after white space

id_info = pd.DataFrame({
    'id': id_col,
    'sexo': sexo,
    'edad': edad_col,
    'altura': alt_col
})

# Converting data types
id_info.edad = id_info.edad.astype(np.float)

for row in range(0, len(id_info)):  # Trying to change every value
    id_info.iloc[row, 3] = convert_to_float(id_info.iloc[row, 3])

# Joining DF to have all the info in one DF
nofat_data = pd.merge(nofat_data, id_info, on='id', how='left')

# Replacing imc column with the correct calculation
nofat_data.imc = nofat_data.peso / nofat_data.altura**2

# Cleaning null. Most of them are from people who finish the program before 28 days
nofat_data.dropna(inplace=True)
nofat_data.reset_index(drop=True, inplace=True)


# Ciclos
ciclos = []
for i in range(0, len(nofat_data)):
    if nofat_data.day[i] < 15:
        ciclos.append('C1')
    else:
        ciclos.append('C2')
nofat_data['ciclo'] = ciclos

cycle_day_col = []
for i in range(0, len(nofat_data)):
    if nofat_data.day[i] > 14:
        cycle_day_col.append(nofat_data.day[i] - 14)
    else:
        cycle_day_col.append(nofat_data.day[i])
nofat_data['dia_ciclo'] = cycle_day_col


# imc_change and imc_change_accum
# Add column of diff and acumm_diff of IMC
imc_diff = []
imc_accum = []
for id_row in nofat_data.id.unique():
    print(id_row)
    imc_aux = []
    imc_accum_aux = []
    for i in range(0, len(nofat_data)):
        if nofat_data.id[i] == id_row:
            if len(imc_aux) == 0:
                imc_aux.append(0)  # First value of difference and accum is always 0
                imc_accum_aux.append(0)  # First value of difference and accum is always 0
                last_imc_value = nofat_data.imc[i]
                first_imc_value = nofat_data.imc[i]
            elif math.isnan(nofat_data.imc[i]):
                imc_aux.append(math.nan)
                imc_accum_aux.append(math.nan)
            else:
                diff = nofat_data.imc[i] - last_imc_value
                imc_aux.append(diff)
                diff_acum = nofat_data.imc[i] - first_imc_value
                imc_accum_aux.append(diff_acum)
                last_imc_value = nofat_data.imc[i]
    imc_diff = imc_diff + imc_aux
    imc_accum = imc_accum + imc_accum_aux

nofat_data['delta_imc'] = imc_diff
nofat_data['delta_imc_acum'] = imc_accum


# Glucosa_change and glucosa_change_accum
# Creating columns diff and accum_diff
glucosa_diff = []
glucosa_accum = []
for id_row in nofat_data.id.unique():
    print(id_row)
    gluco_aux = []
    gluco_accum_aux = []
    for i in range(0, len(nofat_data)):
        if nofat_data.id[i] == id_row:
            if len(gluco_aux) == 0:
                gluco_aux.append(0)  # First value of difference and accum is always 0
                gluco_accum_aux.append(0)  # First value of difference and accum is always 0
                last_gluco_value = nofat_data.glucosa[i]
                first_gluco_value = nofat_data.glucosa[i]
            elif math.isnan(nofat_data.glucosa[i]):
                gluco_aux.append(math.nan)
                gluco_accum_aux.append(math.nan)
            else:
                diff = nofat_data.glucosa[i] - last_gluco_value
                gluco_aux.append(diff)
                diff_acum = nofat_data.glucosa[i] - first_gluco_value
                gluco_accum_aux.append(diff_acum)
                last_gluco_value = nofat_data.glucosa[i]
    glucosa_diff = glucosa_diff + gluco_aux
    glucosa_accum = glucosa_accum + gluco_accum_aux

nofat_data['delta_glucosa'] = glucosa_diff
nofat_data['delta_glucosa_acum'] = glucosa_accum


# Creating new id: id_ciclo column
id_ciclo = []
for i in range(0, len(nofat_data)):
    id_ciclo.append(str(int(nofat_data.id[i])) + '_' + nofat_data.ciclo[i])
nofat_data['id_ciclo'] = id_ciclo


# Add column of diff and acumm_diff of IMC for every id_ciclo
imc_diff = []
imc_accum = []
for id_row in nofat_data.id_ciclo.unique():
    print(id_row)
    imc_aux = []
    imc_accum_aux = []
    for i in range(0, len(nofat_data)):
        if nofat_data.id_ciclo[i] == id_row:
            if len(imc_aux) == 0:
                imc_aux.append(0)  # First value of difference and accum is always 0
                imc_accum_aux.append(0)  # First value of difference and accum is always 0
                last_imc_value = nofat_data.imc[i]
                first_imc_value = nofat_data.imc[i]
            elif math.isnan(nofat_data.imc[i]):
                imc_aux.append(math.nan)
                imc_accum_aux.append(math.nan)
            else:
                diff = nofat_data.imc[i] - last_imc_value
                imc_aux.append(diff)
                diff_acum = nofat_data.imc[i] - first_imc_value
                imc_accum_aux.append(diff_acum)
                last_imc_value = nofat_data.imc[i]
    imc_diff = imc_diff + imc_aux
    imc_accum = imc_accum + imc_accum_aux

nofat_data['ciclo_delta_imc'] = imc_diff
nofat_data['ciclo_delta_imc_acum'] = imc_accum


# Add column of diff and acumm_diff of glucosa for every id_ciclo
glucosa_diff = []
glucosa_accum = []
for id_row in nofat_data.id_ciclo.unique():
    print(id_row)
    gluco_aux = []
    gluco_accum_aux = []
    for i in range(0, len(nofat_data)):
        if nofat_data.id_ciclo[i] == id_row:
            if len(gluco_aux) == 0:
                gluco_aux.append(0)  # First value of difference and accum is always 0
                gluco_accum_aux.append(0)  # First value of difference and accum is always 0
                last_gluco_value = nofat_data.glucosa[i]
                first_gluco_value = nofat_data.glucosa[i]
            elif math.isnan(nofat_data.glucosa[i]):
                gluco_aux.append(math.nan)
                gluco_accum_aux.append(math.nan)
            else:
                diff = nofat_data.glucosa[i] - last_gluco_value
                gluco_aux.append(diff)
                diff_acum = nofat_data.glucosa[i] - first_gluco_value
                gluco_accum_aux.append(diff_acum)
                last_gluco_value = nofat_data.glucosa[i]
    glucosa_diff = glucosa_diff + gluco_aux
    glucosa_accum = glucosa_accum + gluco_accum_aux

nofat_data['ciclo_delta_glucosa'] = glucosa_diff
nofat_data['ciclo_delta_glucosa_acum'] = glucosa_accum


# -------------------------- CREATING CETONA_DATA DB, WITH OTHER MEDIC VALUES ------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
cetona_data = pd.DataFrame()
pac_id = 1
for j in files:
    # Import only a section of the Excel file
    print(j)
    paciente = pd.read_excel(
        io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
        sheet_name='BITACORA',
        usecols='D:AF',
        skiprows=61,
        nrows=6
    )

    # Drop unusefull rows
    paciente.drop([0, 3], inplace=True, axis=0)

    # Swapping Cols and Rows
    paciente = paciente.T
    paciente.reset_index(drop=True, inplace=True)

    # Columns names
    new_header = paciente.iloc[0]  # grab the first row for the header
    paciente = paciente[1:]  # take the data less the header row
    paciente.columns = new_header  # set the header row as the df header

    # Creating id_col
    id_col = []
    for i in range(0, len(paciente)):
        id_col.append(pac_id)
    pac_id = pac_id + 1

    # Adding id and day columns at the beginning of the paciente dataframe
    paciente.insert(0, 'day', day_col, True)
    paciente.insert(0, 'id', id_col, True)

    # Concat with nofat_data
    cetona_data = pd.concat([cetona_data, paciente])

# Reset index of nofat_data
cetona_data.reset_index(drop=True, inplace=True)

# Rename columns
new_cols = {'GLUCOSA SENSOR': 'glucosa_sensor', 'CETONA SANGRE': 'cetona_sangre', 'BRISTOL': 'bristol',
            'DIURESIS': 'diuresis'}
cetona_data.rename(columns=new_cols, inplace=True)

# Cleaning null.
cetona_data.dropna(inplace=True)
cetona_data.reset_index(drop=True, inplace=True)

# Correcting data
# Bristol
# All upper
cetona_data.bristol = cetona_data.bristol.str.upper()
# Erase spaces after and before
cetona_data.bristol = cetona_data.bristol.str.strip()
# Replacing N/A with nan
cetona_data.bristol = cetona_data.bristol.replace('N/A', pd.np.nan)

# Diuresis
# All upper
cetona_data.diuresis = cetona_data.diuresis.str.upper()
# Erase spaces after and before
cetona_data.diuresis = cetona_data.diuresis.str.strip()
# Replacing N/A with nan
cetona_data.diuresis = cetona_data.diuresis.replace('N/A', pd.np.nan)

# Cleaning null.
cetona_data.dropna(inplace=True)
cetona_data.reset_index(drop=True, inplace=True)

# Convert all to number type. If can't, convert it in nan and then drop it (because there are some misspelling values)
def convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        return pd.np.nan

# Applying function to all DF, value by value
for col in range(0, 4):
    print(cetona_data.columns[col])
    for row in range(0, len(cetona_data)):  # Trying to change every value, one column at time
        cetona_data.iloc[row, col] = convert_to_float(cetona_data.iloc[row, col])

# Cleaning null
cetona_data.dropna(inplace=True)
cetona_data.reset_index(drop=True, inplace=True)

# Asign data type
cetona_data.id = cetona_data.id.astype(float)
cetona_data.day = cetona_data.day.astype(float)
cetona_data.glucosa_sensor = cetona_data.glucosa_sensor.astype(float)
cetona_data.cetona_sangre = cetona_data.cetona_sangre.astype(float)

# LEFT JOIN with cetona_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(nofat_data, cetona_data, on=['id', 'day'], how='left')


# -------------------------- CREATING AYUNO DB ---------------------------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
ayuno_data = pd.DataFrame()
pac_id = 1
for j in files:
    # Import only a section of the Excel file
    print(j)
    paciente = pd.read_excel(
        io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
        sheet_name='BITACORA',
        usecols='C:AF',
        skiprows=59,
        nrows=1
    )

    # Swapping Cols and Rows
    paciente = paciente.T
    paciente.reset_index(drop=True, inplace=True)

    # Drop unusefull row
    paciente.drop(1, inplace=True, axis=0)


    # Columns names
    new_header = paciente.iloc[0]  # grab the first row for the header
    paciente = paciente[1:]  # take the data less the header row
    paciente.columns = new_header  # set the header row as the df header

    # Creating id_col
    id_col = []
    for i in range(0, len(paciente)):
        id_col.append(pac_id)
    pac_id = pac_id + 1

    # Adding id and day columns at the beginning of the paciente dataframe
    paciente.insert(0, 'day', day_col, True)
    paciente.insert(0, 'id', id_col, True)

    # Concat with nofat_data
    ayuno_data = pd.concat([ayuno_data, paciente])

# Reset index of ayuno_data
ayuno_data.reset_index(drop=True, inplace=True)

# Rename columns
new_cols = {'AYUNO MÁXIMO': 'ayuno_max'}
ayuno_data.rename(columns=new_cols, inplace=True)

# Convert all to number type. If can't, convert it in nan (because there are some misspelling values)
# Applying function to all DF, value by value
for row in range(0, len(ayuno_data)):  # Trying to change every value
    ayuno_data.iloc[row, 2] = convert_to_float(ayuno_data.iloc[row, 2])

# LEFT JOIN: all_data with ayuno_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(all_data, ayuno_data, on=['id', 'day'], how='left')


# -------------------------- CREATING LIQUIDOS DB to join it with all_data --------------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
liq_data = pd.DataFrame()
pac_id = 1

for j in files:  # with try and except cause some files are messed up
    try:
        # Import only a section of the Excel file
        print(j)
        paciente = pd.read_excel(
            io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
            sheet_name='BITACORA',
            usecols='D:AF',
            skiprows=108,
            nrows=1
        )

        # Swapping Cols and Rows
        paciente = paciente.T
        paciente.reset_index(drop=True, inplace=True)

        # Columns names
        new_header = paciente.iloc[0]  # grab the first row for the header
        paciente = paciente[1:]  # take the data less the header row
        paciente.columns = new_header  # set the header row as the df header

        # Creating id_col
        id_col = []
        for i in range(0, len(paciente)):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    except IndexError:
        # Creating paciente DF
        paciente = pd.DataFrame()

        # Creating id_col
        id_col = []
        for i in range(1, 29):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Creating TOTAL column with only nan values
        tot_col = []
        for i in range(1, 29):
            tot_col.append(pd.np.nan)

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'TOTAL', tot_col, True)
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    # Concat with liq_data
    liq_data = pd.concat([liq_data, paciente])

# Reset index of liq_data
liq_data.reset_index(drop=True, inplace=True)

# Rename columns
new_cols = {'TOTAL': 'total_liq'}
liq_data.rename(columns=new_cols, inplace=True)

# Convert all to number type. If can't, convert it in nan (because there are some misspelling values)
# Applying function to all DF, value by value
for row in range(0, len(liq_data)):  # Trying to change every value
    liq_data.iloc[row, 2] = convert_to_float(liq_data.iloc[row, 2])

# LEFT JOIN: all_data with liq_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(all_data, liq_data, on=['id', 'day'], how='left')


# -------------------------- CREATING SUENHO DB to join it with all_data --------------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
sleep_data = pd.DataFrame()
pac_id = 1

for j in files:  # with try and except cause some files are messed up
    try:
        # Import only a section of the Excel file
        print(j)
        paciente = pd.read_excel(
            io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
            sheet_name='BITACORA',
            usecols='D:AF',
            skiprows=19,
            nrows=1
        )

        # Swapping Cols and Rows
        paciente = paciente.T
        paciente.reset_index(drop=True, inplace=True)

        # Columns names
        new_header = paciente.iloc[0]  # grab the first row for the header
        paciente = paciente[1:]  # take the data less the header row
        paciente.columns = new_header  # set the header row as the df header

        # Creating id_col
        id_col = []
        for i in range(0, len(paciente)):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

        # Getting hours of sleep (in number format, rounded).
        sleep_col = []
        for i in range(0, len(paciente)):
            try:
                sleep_time = paciente.iloc[i,2].hour + paciente.iloc[i,2].minute/60
                sleep_col.append(sleep_time)
            except AttributeError:
                sleep_col.append(pd.np.nan)

        # Adding sleep_col to paciente DF
        paciente.insert(3, 'hr_suenho', sleep_col, True)

        # Drop unusefull row
        paciente.drop('TOT. SUEÑO', inplace=True, axis=1)

    except IndexError:
        # Creating paciente DF
        paciente = pd.DataFrame()

        # Creating id_col
        id_col = []
        for i in range(1, 29):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Creating TOTAL column with only nan values
        hr_suenho = []
        for i in range(1, 29):
            hr_suenho.append(pd.np.nan)

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'hr_suenho', hr_suenho, True)
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    # Concat with liq_data
    sleep_data = pd.concat([sleep_data, paciente])

# Reset index of liq_data
sleep_data.reset_index(drop=True, inplace=True)

# Convert all to number type. If can't, convert it in nan (because there are some misspelling values)
# Applying function to all DF, value by value
for row in range(0, len(sleep_data)):  # Trying to change every value
    sleep_data.iloc[row, 2] = convert_to_float(sleep_data.iloc[row, 2])

# LEFT JOIN: all_data with liq_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(all_data, sleep_data, on=['id', 'day'], how='left')


# -------------------------- CREATING O2 DB to join it with all_data --------------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
o2_data = pd.DataFrame()
pac_id = 1

for j in files:  # with try and except cause some files are messed up
    try:
        # Import only a section of the Excel file
        print(j)
        paciente = pd.read_excel(
            io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
            sheet_name='BITACORA',
            usecols='D:AF',
            skiprows=37,
            nrows=1
        )

        # Swapping Cols and Rows
        paciente = paciente.T
        paciente.reset_index(drop=True, inplace=True)

        # Columns names
        new_header = paciente.iloc[0]  # grab the first row for the header
        paciente = paciente[1:]  # take the data less the header row
        paciente.columns = new_header  # set the header row as the df header

        # Creating id_col
        id_col = []
        for i in range(0, len(paciente)):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    except IndexError:
        # Creating paciente DF
        paciente = pd.DataFrame()

        # Creating id_col
        id_col = []
        for i in range(1, 29):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Creating TOTAL column with only nan values
        o2_col = []
        for i in range(1, 29):
            o2_col.append(pd.np.nan)

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'Spo2%', o2_col, True)
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    # Concat with liq_data
    o2_data = pd.concat([o2_data, paciente])

# Reset index of liq_data
o2_data.reset_index(drop=True, inplace=True)

# Rename columns
new_cols = {'Spo2%': 'o2'}
o2_data.rename(columns=new_cols, inplace=True)

# Convert all to number type. If can't, convert it in nan (because there are some misspelling values)
# Applying function to all DF, value by value
for row in range(0, len(o2_data)):  # Trying to change every value
    o2_data.iloc[row, 2] = convert_to_float(o2_data.iloc[row, 2])

# LEFT JOIN: all_data with liq_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(all_data, o2_data, on=['id', 'day'], how='left')


# -------------------------- CREATING MEDIDAS DB to join it with all_data --------------------------------------------
# First, create a day_column (days from 1 to 28)
day_col = []
for i in range(1,29):
    day_col.append(i)

# "For" to get the data of every excel
med_data = pd.DataFrame()
pac_id = 1

for j in files:  # with try and except cause some files are messed up
    try:
        # Import only a section of the Excel file
        print(j)
        paciente = pd.read_excel(
            io=f'C:/Harold Data/Projects/2022-11-14 PM/Data RAW/2022-11-21/NOFATER/{j}',
            sheet_name='BITACORA',
            usecols='D:AF',
            skiprows=43,
            nrows=3
        )

        # Swapping Cols and Rows
        paciente = paciente.T
        paciente.reset_index(drop=True, inplace=True)

        # Columns names
        new_header = paciente.iloc[0]  # grab the first row for the header
        paciente = paciente[1:]  # take the data less the header row
        paciente.columns = new_header  # set the header row as the df header

        # Creating id_col
        id_col = []
        for i in range(0, len(paciente)):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    except IndexError:
        # Creating paciente DF
        paciente = pd.DataFrame()

        # Creating id_col
        id_col = []
        for i in range(1, 29):
            id_col.append(pac_id)
        pac_id = pac_id + 1

        # Creating brazo, cintura and muslo columns with only nan values
        brazo_col = []
        cintura_col = []
        muslo_col = []
        for i in range(1, 29):
            brazo_col.append(pd.np.nan)
            cintura_col.append(pd.np.nan)
            muslo_col.append(pd.np.nan)

        # Adding id and day columns at the beginning of the paciente dataframe
        paciente.insert(0, 'MUSLO', muslo_col, True)
        paciente.insert(0, 'CINTURA', cintura_col, True)
        paciente.insert(0, 'BRAZO', brazo_col, True)
        paciente.insert(0, 'day', day_col, True)
        paciente.insert(0, 'id', id_col, True)

    # Concat with liq_data
    med_data = pd.concat([med_data, paciente])

# Reset index of liq_data
med_data.reset_index(drop=True, inplace=True)

# Rename columns
new_cols = {'BRAZO': 'brazo', 'CINTURA': 'cintura', 'MUSLO': 'muslo'}
med_data.rename(columns=new_cols, inplace=True)

# Convert all to number type. If can't, convert it in nan (because there are some misspelling values)
# Applying function to all DF, value by value
for row in range(0, len(med_data)):  # Trying to change every value
    med_data.iloc[row, 2] = convert_to_float(med_data.iloc[row, 2])
    med_data.iloc[row, 3] = convert_to_float(med_data.iloc[row, 3])
    med_data.iloc[row, 4] = convert_to_float(med_data.iloc[row, 4])

# LEFT JOIN: all_data with liq_data using id and day
# Use the `merge()` function to perform a left join
all_data = pd.merge(all_data, med_data, on=['id', 'day'], how='left')


# Saving advance
all_data.to_csv('C:/Harold Data/Projects/2022-11-14 PM/all_data.csv')
# loading DB
all_data = pd.read_csv('C:/Harold Data/Projects/2022-11-14 PM/all_data.csv')
all_data.drop('Unnamed: 0', axis=1, inplace=True)

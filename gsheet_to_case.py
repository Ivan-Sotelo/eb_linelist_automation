import pandas as pd
import redcap
import json
import numpy as np
import pandas as pd
import sys

g_sheet_file = sys.argv[1]
sheet_name = sys.argv[2]
output_dir = sys.argv[3]


def connect_to_subset(field):
    # Connects to and returns subset of project data frame
    # Uses gssotelo Api key
    api_url = 'http://192.168.20.33/redcap/api/'
    api_key = 'CA7768C863D7E1A5791BC222E92C3A0C'
    project = redcap.Project(api_url, api_key)
    fields_of_interest = field
    subset_df = project.export_records(fields=fields_of_interest)
    #project_df = project.export_records(format='df')
    return subset_df

def process_subset_df(field):
    redcap_df = connect_to_subset(field)
    
    #transform to pandas dataframe
    df = pd.DataFrame(redcap_df)

    #subset to only barcodes and local_id and central id
    barcodes = df[["central_id"] + field]
    barcodes = barcodes[["central_id"] + field]

    #omit index
    barcodes = barcodes.reset_index(drop=True)
    return barcodes
    
def import_gsheet(filename, sheet_name):
    #import gsheet
    g_sheet = pd.ExcelFile(filename)
    sheet_df = pd.read_excel(g_sheet, sheet_name)
    sheet_df = sheet_df[sheet_df['RITM Lab ID'].notna()]

    return sheet_df

def get_sheet_columns(sheet_df):
    # GET GSHEET COLUMNS
    g_sheet_columns_case = ['RITM Lab ID', 'DATE OF COLLECTION (MM-DD-YYYY)','AGE', 'SEX', 'SAMPLE TYPE',
                        'PATIENT ADDRESS (CITY)', 'PATIENT ADDRESS (PROVINCE)', 'PATIENT ADDRESS (REGION)',
                        'BARCODE', 'Health Status', 'DRU', 'DRU ADDRESS']
    sheet_df = sheet_df[g_sheet_columns_case]
    # sheet_df['AGE'] = sheet_df['AGE'].astype('Int64')       #casts age data to int
    # sheet_df['BARCODE'] = sheet_df['BARCODE'].astype('Int64')
    sheet_df = sheet_df.astype(str)

    return sheet_df

def transfer_data_from_sheet_to_case_df(filename, sheet_df):
    #IMPORT CASE CSV AND TRANSFER DATA FROM G SHEET TO CASE CSV DATAFRAME
    case_columns = ['redcap_repeat_instance', 'redcap_repeat_instrument', 'gisaid_name', 'ont_barcode', 'local_id', 'adm3', 'adm2', 'adm1', 'adm0', 'date_collected', 'sample_type_collected', 'age', 'sex', 'health_status', 'patient_outcome']
    case_df = pd.read_csv(filename, usecols = case_columns)
    case_df[['local_id', 'date_collected', 'age', 'sex', 'sample_type_collected', 'adm3', 'adm2', 'adm1', 'ont_barcode', 'health_status']] = sheet_df[['RITM Lab ID', 'DATE OF COLLECTION (MM-DD-YYYY)', 'AGE', 'SEX', 'SAMPLE TYPE', 'PATIENT ADDRESS (CITY)', 'PATIENT ADDRESS (PROVINCE)', 'PATIENT ADDRESS (REGION)', 'BARCODE', 'Health Status']]
    
    
    return case_df

def import_dict_json(filename):
    #IMPORTS DICTIONARY FROM TXT FILE
    with open(filename) as f:
        data = f.read()
    js = json.loads(data)

    return js

def import_column_dictionaries(adm3_file, adm2_file, adm1_file, sample_file, sex_file, health_file):
    #IMPORT dictionaries for each column# 	
    # city/municipality
    adm3_dict = import_dict_json(adm3_file)
    
    # province
    adm2_dict = import_dict_json(adm2_file)
    
    # region
    adm1_dict = import_dict_json(adm1_file)
    
    # TODO: make a dictionary for countries, currrently not done as codebook contains commas(,) and quotes(")   

    # sample_type_collected
    # this uses the right value: key mapping, everythine else will needed to be inverted 
    sample_type_dict = import_dict_json(sample_file)

    # sex
    # this uses the right value: key mapping, everythine else will needed to be inverted 
    sex_dict = import_dict_json(sex_file)

    # health status
    health_status_dict = import_dict_json(health_file)

    # patient outcome
    

    return adm3_dict, adm2_dict, adm1_dict, sample_type_dict, sex_dict, health_status_dict

def invert_dictionary(data_dict):
    data_dict = {v: k for k, v in data_dict.items()}
    return data_dict

def lower_dictionary(data_dict):
    data = {k.lower():v for k,v in data_dict.items()}
    return data

def lower_dataframe(case_df):
    #   lower the case of the gsheet values to make it case insensitive

    case_df['adm3'] = case_df['adm3'].str.lower()
    case_df['adm2'] = case_df['adm2'].str.lower()
    case_df['sample_type_collected'] = case_df['sample_type_collected'].str.lower()
    case_df['age'] = case_df['age'].str.lower()
    case_df['sex'] = case_df['sex'].str.lower()
    case_df['health_status'] = case_df['health_status'].str.lower()

    return case_df

def map_df_to_dictionary(case_df):
    case_df = case_df.replace({"adm1": adm1_dict}) 
    case_df = case_df.replace({"adm2": adm2_dict})
    case_df = case_df.replace({"sample_type_collected": sample_type_dict})
    case_df = case_df.replace({"sex": sex_dict})
    case_df = case_df.replace({"health_status": health_status_dict})


    return case_df

def subset_city_from_province(case_df, adm3_dict):
    for index, row in case_df.iterrows():
        adm3_subset = {key: value for key, value in adm3_dict.items() if(row['adm2'] in key)}  # subset adm3_dictionary that contains the adm2 string
        adm3_subset = invert_dictionary(adm3_subset)                                           # invert the subset dictionary
        adm3_subset = lower_dictionary(adm3_subset)                                            # lower the keys in the subset to match gsheet data
        case_df.at[index, 'adm3'] = adm3_subset.get(row['adm3'], row['adm3'])                                  # replace the adm3 value with corresponding value from the subset dictionary

    return case_df

def get_region_from_province(case_df):
    for index, row in case_df.iterrows():
        if row['adm1'] == 'nan' or row['adm1'].startswith('PH') == False :
            if row['adm2'].startswith('PH'):
                adm1_code = row['adm2'][0:4]                                                   # assigns first 4 letters of province code as the region code
                case_df.at[index, "adm1"] = adm1_code
            else:
                case_df.at[index, "adm1"] = ''
    return case_df

def put_PH_as_the_country_code(case_df):
    for index, row in case_df.iterrows():
        if row['local_id'].startswith('NC') or row['local_id'].startswith('NEC') or row['local_id'].startswith('NTC'):
            case_df.at[index, "adm0"] = ''
        else:
            case_df.at[index, "adm0"] = 'PH'
    return case_df
  
def get_central_id(case_df, barcodes):
    # MERGE case_df with barcodes on local id to get right central_id from redcap
    case_df = pd.merge(case_df, barcodes, on='local_id', how='inner')
    
    # make central_id the first column
    cols = ['central_id',
    'redcap_repeat_instrument',
    'redcap_repeat_instance',
    'local_id',
    'ont_barcode',
    'gisaid_name',
    'adm3',
    'adm2',
    'adm1',
    'adm0',
    'date_collected',
    'sample_type_collected',
    'age',
    'sex',
    'health_status',
    'patient_outcome']

    case_df = case_df[cols]
    return case_df

def populate_misc_columns(case_df):
    # POPOULATE NECESSARY COLUMNS
    dag = 'ritm'
    for index, row in case_df.iterrows():
        case_df.at[index, "redcap_repeat_instrument"] = 'case'
        case_df.at[index, "redcap_repeat_instance"] = '1'
        # Set gisaid_name of summary_df record using central_id column value.
        case_df.at[index, "gisaid_name"] = "PH"+"-"+dag.upper()+"-"+row["central_id"]
    return case_df

def replace_nan(case_df):
    # REPLACE ALL nan VALUES WITH EMPTY STRING
    case_df = case_df.replace("nan", "", regex=True)
    case_df = case_df.replace("<NA>", "", regex=True)
    case_df = case_df.replace("NaT", "", regex=True)
    case_df = case_df.replace("<na>", "", regex=True)
    case_df = case_df.replace("NAN", "", regex=True)
    return case_df

def put_5_as_patient_outcome(case_df):
    for index, row in case_df.iterrows():
        if row['local_id'].startswith('NC') or row['local_id'].startswith('NEC') or row['local_id'].startswith('NTC'):
            case_df.at[index, "patient_outcome"] = ''
        else:
            case_df.at[index, "patient_outcome"] = '5'
    return case_df
    
def transfer_data_from_sheet_to_diagnostic_df(filename, sheet_df):
    #IMPORT CASE CSV AND TRANSFER DATA FROM G SHEET TO CASE CSV DATAFRAME
    case_columns = ['diagnostic_local_id', 'originating_lab', 'originating_lab_address']
    diagnostic_df = pd.read_csv(filename, usecols = case_columns)
    diagnostic_df[['diagnostic_local_id', 'originating_lab', 'originating_lab_address' ]] = sheet_df[['RITM Lab ID', 'DRU', 'DRU ADDRESS']]
    
    return diagnostic_df 

def match_on_local_id_to_case(diagnostic_df, case_df):
    #MATCH DIAGNOSTIC DATAFRAME WITH CASE DATAFRAME
    diagnostic_df.rename(columns = {'diagnostic_local_id':'local_id'}, inplace = True)  #rename to local_id to merge
    
    diagnostic_df = pd.merge(diagnostic_df, case_df, on='local_id', how='inner')
    
    diagnostic_df.rename(columns = {'local_id':'diagnostic_local_id'}, inplace = True)  #rename back to diagnostic_local_id

    

    columns = ['central_id','redcap_repeat_instance', 'redcap_repeat_instrument', 'diagnostic_local_id', 'originating_lab', 'originating_lab_address']
    diagnostic_df = diagnostic_df[columns]
    return diagnostic_df
    


if __name__ == '__main__':

    print('Making redcap case and diagnostic instrument csv...')
    pd.options.mode.chained_assignment = None  # default='warn' removes pandas warnings
    local_id_df = process_subset_df(['local_id'])
    g_sheet_df = import_gsheet(g_sheet_file, sheet_name)
    g_sheet_df = get_sheet_columns(g_sheet_df)
    case_df = transfer_data_from_sheet_to_case_df('case.csv', g_sheet_df)
    adm3_dict, adm2_dict, adm1_dict, sample_type_dict, sex_dict, health_status_dict = \
        import_column_dictionaries('adm3.txt', 'adm2.txt', 'adm1.txt', 'sample_type.txt', 'sex.txt', 'health_status.txt')

    #invert dicitonaries
    adm2_dict = invert_dictionary(adm2_dict)
    adm1_dict = invert_dictionary(adm1_dict)
    health_status_dict = invert_dictionary(health_status_dict)

    #lower dictionaries
    adm2_dict = lower_dictionary(adm2_dict)
    adm1_dict = lower_dictionary(adm1_dict)
    health_status_dict = lower_dictionary(health_status_dict)
    
    #cast df as string
    case_df = case_df.astype(str)
    
    #lower case dataframe
    case_df = lower_dataframe(case_df)
    
    #substiture values from the dictionary to the dataframe using keys
    case_df = map_df_to_dictionary(case_df)
    
    # Populate the adm3 column with cities/ municipalities using province as context
    case_df = subset_city_from_province(case_df, adm3_dict)
    
    # Get region data from province code if region is not mapped by dictionary
    case_df = get_region_from_province(case_df)
    
    #  PUT PH as the country code (manually change if foreign sample)
    case_df = put_PH_as_the_country_code(case_df)

    # Get central_id from redcap
    case_df = get_central_id(case_df, local_id_df)

    # populate columns necessary for redcap
    case_df = populate_misc_columns(case_df)

    # replace nan values with empty string
    case_df = replace_nan(case_df)

    # assign unknown code = 6 to patient_outcome column
    case_df = put_5_as_patient_outcome(case_df)

    #export case_df to csv
    case_df.to_csv(output_dir + '/case_import.csv', index=False)

    print('Redcap case instrument csv generated!')

    diagnostic_df = transfer_data_from_sheet_to_diagnostic_df('diagnostic.csv', g_sheet_df)

    # get the central_id and by merging to case_df
    diagnostic_df = match_on_local_id_to_case(diagnostic_df, case_df)

    # replace all nan on diagnostic_df with empty string
    diagnostic_df = replace_nan(diagnostic_df)

    # set the redcap_repeat_instrument to diagnostic
    diagnostic_df['redcap_repeat_instrument'] = 'diagnostic'

    # export diagnostic_df to csv
    diagnostic_df.to_csv(output_dir + '/diagnostic_import.csv', index=False)
   
    print('Redcap diagnostic instrument csv generated!')



















    

    


    


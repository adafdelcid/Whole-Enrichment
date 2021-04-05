'''
Author: Ada Del Cid 
GitHub: @adafdelcid
Apr 2021
'''
import math
from openpyxl import Workbook
import pandas as pd
import numpy as np
import openpyxl

def main():

    # User inputs (only provide these values DO NOT change anything else on the code to prevent code from breaking)
    # Note the formulations sheet must be on a sheet with name "Formulations", otherwise error will occur
    run_enrichment_analysis(destination_folder, formulations_sheet, csv_filepath, sorted_cells)

def run_enrichment_analysis(destination_folder, formulations_sheet, csv_filepath, sorted_cells):
    '''
    run_enrichment_analysis : driver function, it uses all other functions to create enrichment analysis
        inputs:
                destination_folder : user specified path to the folder where the user wants the excel file created to be saved
                data_sheet : user specified file path to excel spreadsheet with formulation sheet only
                sorted_cells : user specified list of cells that were sorted
    '''

    # create excel destination file
    destination_file = create_excel_spreadsheet(destination_folder)

    # Read formulation sheet and save as dataframe
    df_formulations = create_df_formulation_sheet(formulations_sheet, destination_file)

    # Read CSV file and save as dataframe
    df_norm_counts = create_df_norm_counts(csv_filepath, destination_file)

    # Merge dataframes
    df_merged = merge_formulations_and_norm_counts(df_formulations, df_norm_counts, destination_file, "Formulations + norm counts")



def avg_cell_type(df_merged, sorted_cells):
    '''
    merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single
    data frame and appends it to excel spreadsheet named "Formulations + Norm Counts"
        inputs:
            df_merged : formulations datasheet
            sorted_cells : data frame of normalized counts
        output:
            df_merged : dataframe of merged formulations and normalized counts
    '''



def merge_formulations_and_norm_counts(df_one, df_two, destination_file = '', s_name=''): # pylint: disable=R1710
    '''
    merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single
    data frame and appends it to excel spreadsheet named "Formulations + Norm Counts"
        inputs:
            df_one : first data frame containing formulations
            df_two : second data frame containing norm counts
            destination_file : directory of the excel spreadsheet created
            s_name = name of sheet
        output:
            None, if no destination file and s_name given
            OR
            df_merged : dataframe of merged formulations and normalized counts
    '''
    organized_columns = organize_cell_type(df_two) 

    # inner merge of data frames around barcodes ("BC")
    df_merged = df_one.merge(df_two, on="BC")

    # ordered columns
    l1 = df_one.columns.tolist() # columns up to phospholipid%
    order_columns = l1 + organized_columns # formulation columns and organized sample columns

    # rearrange columns on df_merged
    df_merged = df_merged[order_columns]
    #print(df_merged.columns)

    if destination_file != '':
        # append merged data frames onto spreadsheet on sheet named Formulations + Norm Counts
        # with outliers
        with pd.ExcelWriter(destination_file, engine="openpyxl", mode="w")\
        as writer: # pylint: disable=abstract-class-instantiated
            df_merged.to_excel(writer, sheet_name= s_name, index=False)
    
    return df_merged

def organize_cell_type(df_norm_counts):
    '''
    organize_cell_type : gets dataframe with normalized counts, creates a dataframe without outliers
    based on given percentile
        inputs:
            df_norm_counts :  dataframe with normalized counts
        output:
            organized_columns : list of samples organized by cell types of sorted cells
    '''

    organized_columns = get_columns(df_norm_counts)

    organized_columns.sort()
    return organized_columns

def get_columns(dataframe):
    '''
    get_columns: gets data frame with normalized counts, creates a dataframe without outliers based on given percentile
        inputs:
                dataframe :  dataframe to get column names for
        output:
                sample_columns : list of the names of the columns on the dataframe (names of samples)
    '''

    # drop first column (barcodes)
    dataframe = dataframe.drop("BC", axis=1)
    # get row names and save barcode column
    sample_columns = dataframe.columns.tolist()

    return sample_columns

def create_df_norm_counts(csv_filepath, destination_file):
    '''
    create_df_norm_counts : gets csv file path with normalized counts, creates a dataframe and
    appends it to destination_file on a sheet named "Normalized Counts"
        inputs:
            csv_filepath : file path to csv file
            destination_file :  name of the destination excel file
        output:
            df_norm_counts : data frame with normalized counts
    '''

    # Read CSV file and save as data frame
    df_norm_counts = pd.read_csv(csv_filepath, sep=',', header=0)

    columns = df_norm_counts.columns.tolist() # get names of columns
    # rename first column to BC for barcodes
    df_norm_counts.rename(columns={columns[0]:"BC"}, inplace=True)

    return df_norm_counts

def create_df_formulation_sheet(formulations_sheet, destination_file):
    '''
    create_df_formulation_sheet : gets formulation sheet, creates a dataframe and appends it to
    destination_file on a sheet named " Formulations"
        inputs:
            formulations_sheet : file path to excel sheet of formulation sheet
            destination_file : name of the destination excel file
        output:
            df_formulations : data frame with formulations sheet
    '''

    # Turn formulation sheet into data frames
    df_formulations = pd.read_excel(formulations_sheet, sheet_name="Formulations", header=0)

    columns = df_formulations.columns.tolist()
    df_formulations.rename(columns={columns[0]:"LNP",columns[1]:"BC"}, inplace=True)

    return df_formulations

def create_excel_spreadsheet(destination_folder, file_name="Whole Enrichment Analysis"):
    '''
    create_excel_spreadsheet : creates an excel spreadsheet
        inputs:
            destination_folder : directory of the folder where the user wants the file stored
            file_name : name of the file being created (default = "Enrichment Analysis")
        output:
            destination_file : directory of the excel spreadsheet created
    '''
    if destination_folder[-1] != '/': # check to save file on correct folder
        destination_folder = destination_folder + '/'

    destination_file = destination_folder + file_name + ".xlsx"
    w_b = Workbook()
    w_b.save(destination_file)

    return destination_file

if __name__ == "__main__":
    main()

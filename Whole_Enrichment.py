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
    df_merged = merge_formulations_and_norm_counts(df_formulations, df_norm_counts, destination_file,\
sorted_cells, True)



def get_n_percentile(df_norm_counts, percentile):
    '''
    get_n_percentile : finds value at given percentile from all data
        inputs:
            df_norm_counts :  data frame of normalized counts
            percentile : percentile of values accepted (default = 99.9%)
        output:
            n_at_percentile : value at given percentile
    '''
    return np.percentile(df_norm_counts.to_numpy(), percentile)

def merge_formulations_and_norm_counts(df_formulations, df_norm_counts, destination_file,\
sorted_cells, add_to_excel=False): # pylint: disable=R1710
    '''
    merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single
    data frame and appends it to excel spreadsheet named "Formulations + Norm Counts"
        inputs:
            df_formulations : formulations datasheet
            df_norm_counts : data frame of normalized counts
            organized_columns : list of samples organized by cell types of sorted cells
            destination_file : directory of the excel spreadsheet created
        output:
            df_merged : dataframe of merged formulations and normalized counts
    '''

    organized_columns = organize_cell_type(df_norm_counts, sorted_cells)

    # inner merge of data frames around barcodes ("BC")
    df_merged = df_formulations.merge(df_norm_counts, on="BC")

    # ordered columns
    l1 = df_merged.columns.tolist()[:11] # columns up to phospholipid%
    order_columns = l1 + organized_columns # formulation columns and organized sample columns

    # rearrange columns on df_merged
    df_merged = df_merged[order_columns]

    if add_to_excel:
        # append merged data frames onto spreadsheet on sheet named Formulations + Norm Counts
        # with outliers
        with pd.ExcelWriter(destination_file, engine="openpyxl", mode="w")\
        as writer: # pylint: disable=abstract-class-instantiated
            df_merged.to_excel(writer, sheet_name="Formulations + Norm Counts", index=False)
    else: #save dataframe without outliers
        return df_merged

def organize_cell_type(df_norm_counts, sorted_cells):
    '''
    organize_cell_type : gets data fram with normalized counts, creates a dataframe without outliers
    based on given percentile
        inputs:
            df_norm_counts :  data frame with normalized counts
            sorted_cells : user specified list of cells that were sorted
        output:
            organized_columns : list of samples organized by cell types of sorted cells
    '''

    sample_columns = get_columns(df_norm_counts)

    # organize columns
    organized_columns = []
    for sort_by in sorted_cells:
        for column in sample_columns:
            # if current sort_by is in the name of the current column (e.g: if "SB" in "AD SB102")
            if sort_by in column:
                organized_columns.append(column)

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
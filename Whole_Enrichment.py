# Author: Ada Del Cid
# GitHub: @adafdelcid
# April 2021

import math
from openpyxl import Workbook
import pandas as pd
import numpy as np


def run_enrichment_analysis(destination_folder, formulations_sheet, csv_filepath, sorted_cells, number_naked_bcs,
                            x_percent):
    """
    run_enrichment_analysis : driver function, it uses all other functions to create enrichment analysis
        inputs:
                destination_folder : user specified path to the folder where the user wants the excel file created to be
                                     saved
                data_sheet : user specified file path to excel spreadsheet with formulation sheet only
                sorted_cells : user specified list of cells that were sorted
    """
    # order list of sorted cells alphabetically
    sorted_cells.sort()

    # create excel destination file
    destination_file = create_excel_spreadsheet(destination_folder)

    # Read formulation sheet and save as dataframe
    df_formulations = create_df_formulation_sheet(formulations_sheet)

    # list of components
    list_components = df_formulations.columns.tolist()
    list_components.pop(0)
    list_components.pop(0)
    list_components.pop()

    # Read CSV file and save as dataframe
    df_norm_counts = create_df_norm_counts(csv_filepath)

    # Merge dataframes
    df_merged = merge_formulations_and_norm_counts(df_formulations, df_norm_counts, destination_file,
                                                   "Formulations + norm counts")

    # get ordered list of all samples
    list_samples_by_cell_type = divide_samples_by_cell_type(df_merged, sorted_cells)

    # divide samples by cell types
    dict_df_avg_cell_type = df_cell_types(df_merged, list_samples_by_cell_type)

    dict_df_organs = df_by_organs(df_merged, sorted_cells, dict_df_avg_cell_type)
    df_overall = get_df_overall(dict_df_organs, df_formulations)

    d_df_components_top, d_df_components_bottom = top_bottom_enrichment(destination_file, df_overall, list_components,
                                                                        x_percent, number_naked_bcs)

def get_overall_enrichment(df_overall, dict_components):
    dict_df_component_enrichment = {}

    for __, (component, component_list) in enumerate(dict_components):
        dict_df_component_enrichment[component] = calculate_enrichment(component, component_list, df_overall)
    return


def top_bottom_enrichment(destination_file, df_averaged, list_components, x_percent, number_naked_bcs, sort_by=None):
    """
    top_bottom_enrichment: creates dataframes for best and worst performing LNPs, counts and their
                            formulations
        inputs:
            destination_file : directory of the excel spreadsheet
            sort_by : user specified cell type to sort by
            df_averaged : dataframe with averaged normalized counts by cell type
            x_percent : user specified integer to find top and bottom performing LNPs (0-100)
        output:
            d_df_components_top : dictionary containing dataframes with enrichment analysis of top
                                    performing LNPs
            d_df_components_bottom : dictionary containing dataframes with enrichment analysis of
                                        bottom performing LNPs
    """

    # sort normalized counts by cell type
    df_sorted = sort_norm_counts(df_averaged)
    df_top, df_bottom = top_and_bottom_percent(df_sorted, x_percent, number_naked_bcs)
    print(df_top)

    d_df_components_top = create_enrichment_tables(destination_file, df_averaged, list_components, number_naked_bcs, df_top, sort_by, "Top")
    d_df_components_bottom = create_enrichment_tables(destination_file, df_averaged,list_components, number_naked_bcs, df_bottom, sort_by, "Bottom")

    return d_df_components_top, d_df_components_bottom


def create_enrichment_tables(destination_file, df_averaged, list_components, number_naked_bcs,  df_top_bottom_sort_by=None, sort_by=None,
                             top_or_bottom=None):
    """
    create_enrichment_tables: creates excel sheet with formulation enrichment tables of averaged
                            normalized counts (top or bottom performing LNPs if
                            df_top_bottom_sort_by value is passed) named "Form Enrichment" (or
                            "Form Enrichment" + sort_by + top_or_bottom if df_top_bottom_sort_by
                            provided)
        inputs:
            destination_file : directory of the excel spreadsheet
            df_averaged : dataframe with averaged normalized counts by cell type
            df_top_bottom_sort_by : dataframe of either top or bottom performing LNPs by specified
                                    cell type (default = None)
            sort_by : user specified cell type to sort by (default = None)
            top_or_bottom : specifies if enrichment is for top or bottom performing LNPs by
                            specified cell type(default = None)
        output:
            dict_df_components : dictionary with all data frames of all enrichment calculations of
            df_averaged (or df_top_bottom_sort_by if inputted)
    """

    dict_df_components = get_all_enrichments(df_averaged, list_components, number_naked_bcs, df_top_bottom_sort_by)

    '''current_row_1 = 1  # variable to place formulation enrichments by mole ratio
    current_row_2 = 1  # variable to place formulation enrichments by component
    enrichment_sheet = "Form Enrichment"
    if sort_by is not None:
        enrichment_sheet += " " + sort_by + " " + top_or_bottom
    with pd.ExcelWriter(destination_file, engine="openpyxl", mode="a") \
            as writer:  # pylint: disable=abstract-class-instantiated
        if df_top_bottom_sort_by is None:
            df_averaged.to_excel(writer, sheet_name=enrichment_sheet, index=False)
            off_set = len(df_averaged.columns)
        else:
            df_sorted_by_sort_by = sort_norm_counts(df_averaged)
            df_sorted_by_sort_by.to_excel(writer, sheet_name=enrichment_sheet, index=False)
            off_set = len(df_sorted_by_sort_by.columns)

        for index in range(len(list_components) // 2):
            dict_df_components[list_components[index]].to_excel(writer, sheet_name=enrichment_sheet,
                                                                startrow=current_row_1, startcol=off_set + 2,
                                                                index=False)
            dict_df_components[list_components[index + 4]].to_excel(writer, sheet_name=enrichment_sheet,
                                                                    startrow=current_row_2, startcol=off_set + 6,
                                                                    index=False)
            current_row_1 += len(dict_df_components[list_components[index]]) + 2
            current_row_2 += len(dict_df_components[list_components[index + 4]]) + 2'''

    return dict_df_components


def top_and_bottom_percent(df_sorted, x_percent, number_naked_bcs):
    """
    CURRENTLY NOT CHECKING IF NAKED BARCODES NOT ON BOTTOM
    top_and_bottom_percent: creates dataframes for best and worst performing LNPs, counts and their\
                            formulations
        inputs:
            df_sorted : dataframe with normalized counts sorted in descending order by specified\
                        cell type
            x_percent : user specified integer to find top and bottom performing LNPs (0-100)
            number_naked_bcs : user specified number of naked barcodes
        output:
            df_top : dataframe top performing LNPs
            df_bottom : dataframe bottom performing LNPs
    """

    total_lnp = len(df_sorted.index) - number_naked_bcs  # subtract two because of naked barcodes
    values_x_percent = math.ceil(total_lnp * (x_percent / 100))

    # gets top x percent
    df_top = df_sorted.loc[range(0, values_x_percent)]
    # gets bottom x percent
    df_bottom = df_sorted.loc[range(total_lnp - values_x_percent, total_lnp + number_naked_bcs)]

    # if "NAKED1" not in df_bottom["LNP"].to_list() and "NAKED2" not in df_bottom["LNP"].to_list():
        # raise NameError("Error: Naked barcodes not on bottom " + str(x_percent) + "% + 2!")

    return df_top, df_bottom


def get_all_enrichments(df, list_components, number_naked_bcs, df_top_bottom_sort_by=None):
    """
    get_all_enrichments: calculated enrichment by component or component_ratio
        inputs:
            df : dataframe
            df_top_bottom_sort_by : dataframe of either top or bottom performing LNPs by specified
                                    sort_by (optional input)
        output:
            dict_df_components : dictionary with all dataframes of all enrichment calculations of
                                df_averaged (or df_top_bottom_sort_by if inputted)
    """

    dict_components = get_lists_of_components(df, list_components, number_naked_bcs)
    dict_df_components = {}

    if df_top_bottom_sort_by is None:
        for component in dict_df_components:
            dict_df_components[component] = calculate_enrichment(component, dict_components[component], df)
    else:
        for component in dict_components:
            dict_df_components[component] = calculate_enrichment(component, dict_components[component],
                                                                 df_top_bottom_sort_by)

    return dict_df_components


def calculate_enrichment(component, component_list, df):
    """
    calculate_enrichment: calculated enrichment by component or component_ratio
    """

    component_total = [0] * len(component_list)
    for bc_x in df[component].values:
        for index, value in enumerate(component_list):
            if bc_x == value:
                component_total[index] += 1
                break

    total = sum(component_total)

    component_percent_total = []
    for each_component in component_total:
        component_percent_total.append(round(each_component / total, 9))

    component_list.append("TOTAL")
    component_total.append(total)
    component_percent_total.append(round(sum(component_percent_total)))

    t_component_list = [component_list, component_total, component_percent_total]
    np_temporary = np.array(t_component_list)
    np_temporary = np_temporary.T
    df_component_list = pd.DataFrame(data=np_temporary, columns=[component, "Total #",
                                                                 "% of Total"])

    return df_component_list


# ***** END OF OLD CODE
def get_lists_of_components(df_formulations, list_components, number_naked_bcs):
    """
    get_lists_of_components : returns a dictionary with all component mole ratios and component types
        inputs:
            df_formulations : dataframe of formulations
            list_components : list of the components used to formulate LNPs
            number_naked_bcs : user specified number of naked barcodes
        output:
            dict_components : a dictionary containing list of all the component mole ratios
                            and types
    """

    dict_components = get_dict_components(list_components)

    for component in dict_components:
        dict_components[component] = retrieve_component_list(df_formulations, component, number_naked_bcs)

    return dict_components


def get_dict_components(list_components):
    """
    get_dict_components : creates a dictionary with empty lists for each component in formulation sheet
        input :
            list_components : list of the components used to formulate LNPs
        output :
            dict_components : creates dictionary with empty list to save each type of component used
    """
    dict_components = {}

    for __, item in enumerate(list_components):
        dict_components[item] = []

    return dict_components


def retrieve_component_list(df_formulations, component, number_naked_bcs):
    """
    retrieve_component_list : returns a list of all the different mole ratios or types of a specific
                            component used
        inputs:
            df_formulations : dataframe of formulations
            component : string of the component in question
        output:
            component_list : list of all the different mole ratios or types of a component used
    """
    component_list = []

    for index in range(len(df_formulations[component].values) - number_naked_bcs):
        if df_formulations[component].values[index] not in component_list:
            component_list.append(df_formulations[component].values[index])

    component_list.sort()

    return component_list


def sort_norm_counts(df):
    """
    sort_norm_counts : sorts dataframe in descending order of norm counts
        inputs :
            df : dataframe
        output :
            df_sorted : sorted dataframe
    """
    temp_list = df.columns.tolist()
    sort_by = temp_list[-1]
    df_sorted = df.sort_values(by=sort_by, ascending=False, ignore_index=True)
    return df_sorted


# def sort_norm_counts(sort_by, df):
"""
    sort_norm_counts : sorts dataframe in descending order of norm counts
        inputs :
            sort_by : name of column to sort
            df : dataframe
        output :
            df_sorted : sorted dataframe
    """
    # df_sorted = df.sort_values(by=sort_by, ascending=False, ignore_index=True)
    # return df_sorted


def get_df_overall(dict_df_organs, df_formulations):
    """
    get_df_overall : creates dataframe with overall average
        inputs :
            dict_df_organs : dictionary containing dataframes of all organs
            df_formulations : dataframe with formulations sheet
        output :
            df_overall : dataframe with overall average
    """

    df = pd.DataFrame()

    for key in dict_df_organs:
        df_temp = pd.concat([df, dict_df_organs[key][key + "-AVG"]], axis=1)
        df = df_temp

    avg = df.mean(axis=1)
    df_overall = pd.concat([df_formulations, df], axis=1)
    df_overall["Overall-AVG"] = avg

    return df_overall


def df_by_organs(df_merged, sorted_cells, dict_df_avg_cell_type):
    """
    df_by_organs : creates dictionary with dataframes for all organs
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            sorted_cells : user specified list of the sorted cell types
            dict_df_avg_cell_type : dictionary with averaged dataframes of each cell type
        output :
            dict_df_organs : dictionary containing dataframes of all organs
    """

    list_organs = get_list_organs(sorted_cells)
    dict_cells_by_organs = get_dict_cells_organs(sorted_cells, list_organs)
    dict_df_organs = {}

    for organ in list_organs:
        list_cells_by_organ = dict_cells_by_organs[organ]
        if len(list_cells_by_organ) == 1:
            df = dict_df_avg_cell_type[list_cells_by_organ[0]]
            df = df.rename(columns={list_cells_by_organ[0]: organ + "-AVG"})
            df = df.drop(['std'], axis=1)
            dict_df_organs[organ] = df
        else:
            dict_df_organs[organ] = build_df_organ(df_merged, dict_df_avg_cell_type,
                                                   list_cells_by_organ, organ)
    return dict_df_organs


def build_df_organ(df_merged, dict_df_avg_cell_type, list_cells_by_organ, organ):
    """
    build_df_organ : builder of dataframe for organ
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            dict_df_avg_cell_type : dictionary with averaged dataframes of each cell type
            list_cells_by_organ : list_cells_by_organs : list of cell types sorted of an organ
            organ : organ for which we wish to get dataframe
        output :
            df_organ : dataframe with data for specific organ
    """
    df_organ = df_merged["LNP"].to_frame()
    for cell_type in list_cells_by_organ:
        temp_df = dict_df_avg_cell_type[cell_type]
        df = pd.concat([df_organ, temp_df[cell_type]], axis=1)
        df_organ = df

    # get average of organ
    avg = df_organ.mean(axis=1)
    df_organ[organ + "-AVG"] = avg

    return df_organ


def get_dict_cells_organs(sorted_cells, list_organs):
    """
    get_dict_cells_organs : creates a dictionary containing cell types sorted by organ
        inputs :
            sorted_cells : user specified list of the sorted cell types
            list_organs : list of organs sorted
        output :
            dict_cells_by_organs : dictionary of cell types organized by organ
    """

    dict_cells_by_organs = {}

    for organ in list_organs:
        dict_cells_by_organs[organ] = get_list_cells_by_organ(sorted_cells, organ)

    return dict_cells_by_organs


def get_list_cells_by_organ(sorted_cells, organ):
    """
    get_list_cells_by_organ : gets list of cell types sorted for a specific organ
        inputs :
            sorted_cells : user specified list of the sorted cell types
            organ : specific organ for which we want to get the list of cell types sorted
        outputs :
            list_cells_by_organs : list of cell types sorted of an organ
    """

    list_cells_by_organ = []

    for cell_type in sorted_cells:
        if cell_type[0] == organ:
            list_cells_by_organ.append(cell_type)

    return list_cells_by_organ


def get_list_organs(sorted_cells):
    """
    get_list_organs : returns a list of the organs sorted
        input :
            sorted_cells :  user specified list of the sorted cell types
        output :
            list_organs : list of organs sorted
    """

    list_organs = []
    for cell_type in sorted_cells:
        if cell_type[0] not in list_organs:
            list_organs.append(cell_type[0])

    return list_organs


def df_cell_types(df_merged, list_samples_by_cell_type):
    """
    df_cell_types: gets dataframe of each cell type
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            list_samples_by_cell_type : lists of samples IDs by sorted cell type
        output :
            dict_df_avg_cell_type : dictionary with averaged dataframes of each cell type

    df columns titles like: LNP#   Sample1   Sample2   SampleN   Average   Stdev
    """
    dict_df_avg_cell_type = avg_cell_type(df_merged, list_samples_by_cell_type)
    df1 = df_merged["LNP"].to_frame()
    for key,value in dict_df_avg_cell_type.items():
        df = pd.concat([df1, value], axis=1)
        value.loc[:, key] = df

    return dict_df_avg_cell_type


def avg_cell_type(df_merged, dict_samples_by_cell_type):
    """
    avg_cell_type : calculates the average of each cell type
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            dict_samples_by_cell_type : dictionary containing lists of samples IDs by sorted cell type
        outputs :
            dict_df_by_cell_type : dictionary containing dataframes for each sorted cell type and its average
            and standard deviation
    """
    dict_df_by_cell_type = {}

    for cell_type in dict_samples_by_cell_type:
        dict_df_by_cell_type[cell_type] = get_df_cell_type(df_merged, dict_samples_by_cell_type[cell_type])

    for cell_type, df_cell_type in dict_df_by_cell_type.items():
        avg = df_cell_type.mean(axis=1)
        std = df_cell_type.std(axis=1)
        df_cell_type.loc[:, cell_type] = avg
        df_cell_type.loc[:, "std"] = std

    return dict_df_by_cell_type


def get_df_cell_type(df_merged, list_samples):
    """
    get_df_cell_type : returns dataframe with only samples specified
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            list_samples : list of samples from a given cell type
        outputs :
            df_cell_type : a dataframe containing all samples from a specific cell type
    """
    return df_merged[list_samples]


def divide_samples_by_cell_type(df_merged, sorted_cells):
    """
    divide_samples_by_cell_type : creates a dictionary containing cell types as keys and a list of sample IDs as the
                                    value
        inputs :
            df_merged : dataframe containing formulation information and normalized counts
            sorted_cells : user specified list of the sorted cell types
        output :
            dict_samples_by_cell_type : dictionary containing lists of samples IDs by sorted cell type
    """
    # will be a list containing list of samples organized by cell types
    dict_samples_by_cell_type = {}  # will have same length as sorted_cells

    columns_df_merged = df_merged.columns.tolist()
    columns_df_merged = columns_df_merged[11:]  # merged must include column for charge

    for cell_type in sorted_cells:
        samples_by_cell_type = []
        for sample in columns_df_merged:
            if cell_type in sample:
                samples_by_cell_type.append(sample)
        dict_samples_by_cell_type[cell_type] = samples_by_cell_type

    return dict_samples_by_cell_type


def merge_formulations_and_norm_counts(df_one, df_two, destination_file='', s_name=''):  # pylint: disable=R1710
    """
    merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single
    data frame and appends it to excel spreadsheet named "Formulations + Norm Counts"
        inputs :
            df_one : first dataframe containing formulations
            df_two : second dataframe containing norm counts
            destination_file : directory of the excel spreadsheet created
            s_name = name of sheet
        output :
            df_merged : dataframe containing formulation information and normalized counts
    """
    organized_columns = organize_cell_type(df_two)

    # inner merge of data frames around barcodes ("BC")
    df_merged = df_one.merge(df_two, on="BC")

    # ordered columns
    l1 = df_one.columns.tolist()  # columns up to phospholipid%
    order_columns = l1 + organized_columns  # formulation columns and organized sample columns

    # rearrange columns on df_merged
    df_merged = df_merged[order_columns]

    if destination_file != '':
        # append merged data frames onto spreadsheet on sheet named Formulations + Norm Counts
        # with outliers
        with pd.ExcelWriter(destination_file, engine="openpyxl", mode="w") as writer:
            df_merged.to_excel(writer, sheet_name=s_name, index=False)

    return df_merged


def organize_cell_type(df_norm_counts):
    """
    organize_cell_type : takes in a dataframe and organizes the samples alphabetically
        inputs :
            df_norm_counts :  dataframe of normalized counts
        output :
            list_organized_col : list of samples organized alphabetically
    """

    list_organized_col = get_columns(df_norm_counts)

    list_organized_col.sort()
    return list_organized_col


def get_columns(dataframe):
    """
    get_columns : gets dataframe, removes barcodes and returns list of the names of columns
        inputs :
                dataframe :  a dataframe
        output :
                sample_columns : list of the names of the columns on the dataframe (names of samples)
    """

    # drop first column (barcodes)
    dataframe = dataframe.drop("BC", axis=1)
    # get row names and save barcode column
    sample_columns = dataframe.columns.tolist()

    return sample_columns


def create_df_norm_counts(csv_filepath):
    """
    create_df_norm_counts : gets csv file path with normalized counts, creates a dataframe
        inputs :
            csv_filepath : file path to csv file
        output :
            df_norm_counts : dataframe with normalized counts
    """

    # Read CSV file and save as dataframe
    df_norm_counts = pd.read_csv(csv_filepath, sep=',', header=0)

    columns = df_norm_counts.columns.tolist()  # get names of columns
    # rename first column to BC for barcodes
    df_norm_counts.rename(columns={columns[0]: "BC"}, inplace=True)

    return df_norm_counts


def create_df_formulation_sheet(formulations_sheet):
    """
    create_df_formulation_sheet : gets formulation sheet, creates a dataframe for formulations
        inputs :
            formulations_sheet : file path to excel sheet of formulation sheet
        output :
            df_formulations : dataframe with formulations sheet
    """

    # Turn formulation sheet into dataframe
    df_formulations = pd.read_excel(formulations_sheet, sheet_name="Formulations", header=0)

    columns = df_formulations.columns.tolist()
    df_formulations.rename(columns={columns[0]: "LNP", columns[1]: "BC"}, inplace=True)

    return df_formulations


def create_excel_spreadsheet(destination_folder, file_name="Whole Enrichment Analysis"):
    """
    create_excel_spreadsheet : creates an excel spreadsheet
        inputs :
            destination_folder : directory of the folder where the user wants the file stored
            file_name : name of the file being created (default = "Whole Enrichment Analysis")
        output :
            destination_file : directory of the excel spreadsheet created
    """
    if destination_folder[-1] != '/':  # check to save file on correct folder
        destination_folder = destination_folder + '/'

    destination_file = destination_folder + file_name + ".xlsx"
    w_b = Workbook()
    w_b.save(destination_file)

    return destination_file


if __name__ == "__main__":
    main()

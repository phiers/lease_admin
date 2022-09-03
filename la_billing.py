
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
import os
import sys


def main():
    process = start_program()
    # 5: TODO save analysis (after altering) to new file in inputs (for next month)
    # 4: TODO Create csv file (from analysis, so any edits can be reflected)
    if process == 1:
        print("INSTRUCTIONS:  .....")  # TODO
        sys.exit()
    if process == 2:
        directories = [
            "1_Lx_files",
            "2_lease_files",
            "3_equip_files",
            "4_input_files",
            "5_output_files",
        ]
        check_dir_structure(directories)
        rename_and_move_files("1_Lx_files", "2_lease_files")
        create_separate_homage_and_express_file("2_lease_files", "express.xlsx")
    if process == 3:
        date = get_date("Enter the monthend date in MM/DD/YY format: ")
        add_data = add_additional_invoice_items(date, Path.cwd().joinpath("4_input_files", "additional_invoice_items.csv"))
        results = process_files_and_create_dict("2_lease_files", add_data, date)
        create_initial_analysis(results, date)
    if process == 4:
        date = get_date("Enter the monthend date in MM/DD/YY format: ")
        m, _, y = date.split("/")
        path = Path.cwd().joinpath('5_output_files', f'{m}{y}_initial_invoice_analysis.xlsx')
        create_final_analysis_files(path, date)
    if process == 5:
        # TODO: create csv file from final analysis file
        save_lm_analysis_file_to_input_dir()
    if process == 6:
        sys.exit()


def start_program():
    while True:
        try:
            process = input(
                """What would you like to do?
            1 - Get instructions
            2 - Process Lucernex files
            3 - Create initial analysis file
            4 - Create final analysis file
            5 - Create csv file and archive this month's files
            6 - Exit this program
                 \n Enter the appropriate number: """
            )
            if int(process) in [1, 2, 3, 4, 5, 6]:
                return int(process)
            else:
                print("\nENTER A NUMBER BETWEEN 1 AND 6! \n")
        except:
            continue


def check_dir_structure(paths):
    for path in paths:
        path = Path.cwd().joinpath(path)
        folder = str(path).split("/")[-1]
        if os.path.isdir(path) == False:
            sys.exit(
                f"ERROR: The {folder} folder does not exist. Please add one as a subfolder in the Lease_Admin folder."
            )



def rename_and_move_files(directory, target_dir):
    lease_file_count = 0
    try:
        for file in os.scandir(directory):
            if file.name.split(".")[1] == "xlsx":
                wb = load_workbook(file)
                ws = wb.active
                sheet_identifier_list = ws["A1"].value.split(" ")
                if "Equipment" in sheet_identifier_list:
                    name = sheet_identifier_list[-1].split(".")[0] + "_equipment.xlsx"
                    wb.save(Path.cwd().joinpath("3_equip_files", name))
                # Camuto and Town Shoes are on the dsw site, but come as separate files, so have to handle separately
                elif "Camuto" in sheet_identifier_list:
                    wb.save(Path.cwd().joinpath(target_dir, "camuto.xlsx"))
                    lease_file_count += 1
                elif (
                    "Town" in sheet_identifier_list and "Shoes" in sheet_identifier_list
                ):
                    wb.save(Path.cwd().joinpath(target_dir, "townshoes.xlsx"))
                    lease_file_count += 1
                # Handle case where DSW Project count gets copied into emailed files (do nothing)
                elif "Project" in sheet_identifier_list:
                    print(f"{file.name} is not a lease listing and was not processed")
                    continue
                else:
                    name = sheet_identifier_list[-1].split(".")[0] + ".xlsx"
                    wb.save(Path.cwd().joinpath(target_dir, name))
                    lease_file_count += 1
    except FileNotFoundError:
        print(f"ERROR: the '{directory}' folder does not exist or is empty.")

    if lease_file_count == 0:
        print(f"There are no files in the {directory} folder!")
    else:
        print(
            f"{lease_file_count} lease files were created and saved in the lease_files folder"
        )


# Function to separate homage and express (both in express file when sent from Lucernex)
def create_separate_homage_and_express_file(dir, file):
    df = pd.read_excel(Path.cwd().joinpath(dir, file), header=1)
    cols = df.columns
    # find homage leases and make new file
    df_homage = df[df["Contract Name"].str.startswith("Hom")]
    df_homage.to_excel(Path.cwd().joinpath(dir, "homage.xlsx"), columns=cols)

    # remove homage from df and overwrite express.xlsx
    df_express = df[~df["Contract Name"].str.startswith("Hom")]
    df_express.to_excel(Path.cwd().joinpath(dir, "express_only.xlsx"), columns=cols)


# Helper function to get user input for date
def get_date(query_str):
    while True:
        try:
            date = input(query_str)
            # test to make sure user entered correctly
            _, _, _ = date.split("/")
            return date
        except ValueError:
            continue


# Helper function to create customer names dictionary (file names xref to NS names)
def create_cust_name_dict():
    try:
        return (
            pd.read_csv("4_input_files/customer_names.csv", header=None, index_col=0)
            .squeeze()
            .to_dict()
        )
    except FileNotFoundError:
        sys.exit("customer_names.csv file must be in input_files directory")



# Helper function to create lists to add lease abstracts and other items TODO: use excel file instead?
def add_additional_invoice_items(date, file_path):
    # Read file and create df
    df = pd.read_csv(file_path, index_col="Customer_File_Name",usecols=[0,1,2]).dropna(thresh=2)
    customer_name_dict = create_cust_name_dict()
    clients = []
    memos = []
    descriptions = []
    lx_type_codes = []
    quantities = []
    dates = []
    external_ids = []

    for ind in df.index:
        #TODO this doesn't work if client has more than one line item because new_data would need to be iterated over
        new_data = df.loc[ind]
        print(new_data)
        try:
            client = customer_name_dict[ind]
        except KeyError:
            client = ind
        clients.append(client)
        memos.append("Lease Admin Services")
        dates.append(date)
        ext_id = client[:5] + date[-2:] + date[:2]
        external_ids.append(ext_id.replace(" ", "_")
                            .replace(",", "_")
                            .replace(".", "_"))
        for k, v in new_data.items():
            if k == "Description":
                descriptions.append(v)
                lx_type_codes.append(f"{client}_{v}")
            elif k == "Quantity":
                quantities.append(v)
    
    print(clients)
    return external_ids, clients, dates, memos, lx_type_codes, quantities, descriptions

def process_files_and_create_dict(directory, addl_invoice_items, date):
    customer_name_dict = create_cust_name_dict()
    # build lists for dictionary items
    clients = addl_invoice_items[1]
    print(clients)
    memos = addl_invoice_items[3]
    descriptions = addl_invoice_items[6]
    lx_type_codes = addl_invoice_items[4]
    quantities = addl_invoice_items[5]
    dates = addl_invoice_items[2]
    external_ids = addl_invoice_items[0]
    # blank dictionary
    results_dict = {}

    for file in os.scandir(directory):
        if file.name.split(".")[1] == "xlsx":
            try:
                client_name = customer_name_dict[file.name.split(".")[0]]
            except KeyError:
                # express file is split into homage and express_only, so exclude express from keys
                if file.name != "express.xlsx":
                    print(
                        f"ERROR! The customer name does not exist on customer_names.csv for {file.name}"
                    )

            # the files created for express and homage start at row 0, not 1
            header = 1
            if file.name == "homage.xlsx" or file.name == "express_only.xlsx":
                header = 0

            # exclude original express file
            if file.name != "express.xlsx":
                # read files and poplulate lists
                df = pd.read_excel(file, header=header)
                try:
                    # need to distinguish between domestic and international for Tory
                    if file.name == "tory.xlsx":
                        data = df.loc[:, ["Lease Status", "Region"]]
                        for key, value in data.value_counts().items():
                            if key[1] == "North America":
                                clients.append(client_name)
                                memos.append("Lease Admin Services")
                                descriptions.append(f"{key[0]} - Domestic")
                                lx_type_codes.append(
                                    f"{client_name}_{key[0]} - Domestic"
                                )
                                quantities.append(value)
                                dates.append(date)
                                external_ids.append(
                                    client_name[:5] + date[-2:] + date[:2]
                                )
                            else:
                                clients.append(client_name)
                                memos.append("Lease Admin Services")
                                descriptions.append(f"{key[0]} - International")
                                lx_type_codes.append(
                                    f"{client_name}_{key[0]} - International - {key[1]}"
                                )
                                quantities.append(value)
                                dates.append(date)
                                external_ids.append(
                                    client_name[:5] + date[-2:] + date[:2]
                                )
                    else:
                        data = df.loc[:, "Lease Status"]
                        for key, value in data.value_counts().items():
                            clients.append(client_name)
                            memos.append("Lease Admin Services")
                            descriptions.append(key)
                            lx_type_codes.append(f"{client_name}_{key}")
                            quantities.append(value)
                            dates.append(date)
                            ext_id = client_name[:5] + date[-2:] + date[:2]
                            external_ids.append(
                                ext_id.replace(" ", "_")
                                .replace(",", "_")
                                .replace(".", "_")
                            )
                # throw error if file is not processsed (i.e., doesn't have "Lease Status" column)
                except KeyError:
                    print(f"ERROR! {file.name} was not processed or had no data")
    
    results_dict["External_ID"] = external_ids
    results_dict["Customer"] = clients
    results_dict["Date"] = dates
    results_dict["Memo"] = memos
    results_dict["Lx_Type"] = descriptions
    results_dict["Lx_Type_Code"] = lx_type_codes
    results_dict["Quantity"] = quantities

    return results_dict


# helper function to create dataframe with prices and descriptions from maintainable csv file
def create_price_and_description_df():
    try:
        return pd.read_csv(
            Path.cwd().joinpath("4_input_files", "type_desc_price_matrix.csv")
        )
    except FileNotFoundError:
        print(
            "ERROR: A file named 'type_desc_price_matrix.csv' must be in the input files directory"
        )



# helper function to create dataframe for last month's data from excel
def create_lm_df():
    try:
        df = pd.read_excel(
            # TODO: change filename to grab prior month
            Path.cwd().joinpath("4_input_files", "lm_invoice_analysis.xlsx"),
            usecols=["Lx_Type_Code", "Quantity", "Price"],
        )
        # rename columns
        df.columns = ["Lx_Type_Code", "LM_Quantity", "LM_Price"]
        # get rid of rows without quantity
        df = df[df["LM_Quantity"] != 0]

        return df

    except FileNotFoundError:
        print(
            "A file named 'lm_invoice_analysis.xlsx' must be in the input files directory"
        )


def create_initial_analysis(dic, date):
    # create df with monthly invoice data
    df_monthly_data = pd.DataFrame.from_dict(dic)
    # rename price and description dataframe
    df_price_desc = create_price_and_description_df()
    
    # create df with last month's data
    df_lm = create_lm_df()
    #df_lm.to_excel(Path.cwd().joinpath('5_output_files', 'lm.xlsx'))
    initial_combined_df = pd.merge(
        df_monthly_data, df_lm, how="outer", on="Lx_Type_Code",
    )
    #TODO figure out a way where this can process without having to do a bunch of manual stuff on the backend! 
    #initial_combined_df.to_excel(Path.cwd().joinpath('5_output_files', 'initial_combined.xlsx'))
    # combine the dataframes so price and description is in monthly data
    combined_df = pd.merge(
        initial_combined_df, df_price_desc, how="outer", on="Lx_Type_Code", indicator=True
    )
    # eliminate unnecessary rows TODO
    #combined_df = combined_df.loc[(combined_df["Quantity"]) != 0 & (combined_df["LM_Quantity"] != 0)  ]
    # sort rows
    combined_df.sort_values(["Lx_Type_Code", "Invoice_Description"], inplace=True)
    
    # create total and difference columns vs lm
    total = combined_df["Quantity"] * combined_df["Price"]
    lm_total = combined_df["LM_Price"] * combined_df["LM_Quantity"]
    combined_df.insert(10, "Total", total)
    combined_df.insert(12, "LM_Total", lm_total)
    # if there is no quantity this month, this calc won't work

    qnty_vs_lm = combined_df["Quantity"] - combined_df["LM_Quantity"]
    price_vs_lm = combined_df["Price"] - combined_df["LM_Price"]
    total_vs_lm = combined_df["Total"] - combined_df["LM_Total"]

    combined_df.insert(13, "Qnty_vs_LM", qnty_vs_lm)
    combined_df.insert(14, "Price_vs_LM", price_vs_lm)
    combined_df.insert(15, "Total_vs_LM", total_vs_lm)
    
    # sort columns in better order
    col_order = [
        "External_ID",
        "Date",
        "Customer",
        "Lx_Type",
        "Lx_Type_Code",
        "Memo",
        "Invoice_Description",
        "Quantity",
        "Price",
        "Total",
        "LM_Quantity",
        "LM_Price",
        "LM_Total",
        "Qnty_vs_LM",
        "Price_vs_LM",
        "Total_vs_LM",
        "_merge"
    ]
    combined_df = combined_df[col_order]
    # total numeric columns
    columns_to_total = ["Quantity", "Total", "LM_Quantity", "LM_Total", "Qnty_vs_LM", "Total_vs_LM"]
    combined_df.loc["Totals"]= combined_df.loc[:, columns_to_total].sum(axis=0)
    # create a name to save the file under
    month, _, year = date.split("/")
    save_file_path = Path.cwd().joinpath("5_output_files", f"{month}{year}_initial_invoice_analysis.xlsx")
    combined_df.to_excel(save_file_path)
    
    print(combined_df.iloc[1,1])
    print(f"initial analysis file for {date} created")
    return save_file_path


def create_final_analysis_files(file_path, date):
    combined_df = pd.read_excel(file_path).round(2)
    sum_df = combined_df.groupby(["Customer", "Invoice_Description"]).agg(
        {
            "Quantity": "sum",
            "Price": "mean",
            "Total": "sum",
            "LM_Quantity": "sum",
            "LM_Price": "mean",
            "LM_Total": "sum",
            "Qnty_vs_LM": "sum",
            "Price_vs_LM": "sum",
            "Total_vs_LM": "sum",
        }
    ).round(2)

    cust_sum_df = sum_df.groupby("Customer").agg({"Total": "sum", "LM_Total": "sum", "Total_vs_LM": "sum"}).round(2)
    # add totals to columns
    cust_sum_df.loc["Totals"]= cust_sum_df.sum(axis=0)
    sum_df.loc["Totals", :]= sum_df.sum(axis=0).values

    m, _, y = date.split("/")
    save_file_name = f"{m}{y}_final_analysis file.xlsx"
    # write to excel as three separate sheets
    with pd.ExcelWriter(
        Path.cwd().joinpath("5_output_files", save_file_name)
    ) as writer:
        cust_sum_df.to_excel(writer, sheet_name="cust_summary")
        sum_df.to_excel(writer, sheet_name="summary")
        combined_df.to_excel(
            writer,
            sheet_name="detail",
        )
    
    print("Final analysis files created")


def save_lm_analysis_file_to_input_dir():

    file_name = get_file_name()
    try:
        wb = load_workbook(Path.cwd().joinpath("output_files", file_name))
        wb.save(Path.cwd().joinpath("input_files", f"lm_invoice_analysis.xlsx"))
        print("This month's invoice analysis saved for next month's processing")
    except FileNotFoundError:
        print(
            f"ERROR: The file {file_name} was not found in the output_files directory"
        )


def create_csv_from_analysis_file(f):
    """df = pd.DataFrame.from_dict(d)
    # sort rows
    df.sort_values(["Customer", "Lx_Type"], inplace=True)
    #  sort columns in necessary order
    df.loc[:, ["External_ID", "Customer", "Date", "Memo", "Lx_Type", "Quantity"]]
    # set index to External_ID
    df.set_index("External_ID", inplace=True)
    df.to_csv(Path.cwd().joinpath("output_files", "results.csv"))
    print("CSV file created")"""


if __name__ == "__main__":
    main()

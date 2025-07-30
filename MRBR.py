import pandas as pd
from datetime import datetime, date

def get_year():
    current_date = date.today()
    year = current_date.year
    return year

def get_todays_date():
    today = datetime.today()
    year = today.strftime('%Y')  # Use '%Y' for full year (e.g., 2024)
    date_str = today.strftime('%y%m%d')  # Use separate format for date
    return year, date_str

def get_file_path(file_name, today=None, path_type=None, year=None):
    if path_type == "server_imac":
        server_imac_path = f"//sgsind0nsifsv01a/IMAC Data/IMAC Senior or Teams/Europe & Other Asia Team - SP/MRBR/{year}/"
        if today:
            return f"{server_imac_path}MRBR {today}.xlsx"
        else:
            raise ValueError("'today' must be provided for server_imac path_type.")
    elif path_type == "server_master":
        return f"//sgsind0nsifsv01a/IPO DATA/Finance/SG05 MRBR/{file_name}.xlsx"


def filter_dataframe(df, col_name, desired_values):
    df['BUYER Comment'] = ''
    df['Elaine approval'] = ''
    df['SP Update'] = ''
    if col_name is not None:
        df[col_name] = df[col_name].astype(str)
        df = df[~df[col_name].str.contains('1803')]
        df = df.loc[df['Plnt'].isin(desired_values)]

    else:
        df = df.loc[df['Plnt'].isin(desired_values)]
    return df

def main():
    file_name = input("Enter file name: ")
    file_path = get_file_path(file_name, path_type="server_master")
    today_year, today_date = get_todays_date()  # Unpack year and date
    new_excel_server = get_file_path(file_name, today=today_date, path_type="server_imac", year=today_year)
    desired_values = ["HU07", "HU08", "IN07", "IT08", "PL01", "SG02", "VN01"]

    df_site = pd.read_excel(file_path, sheet_name='SITE MRBR')
    df_ipo = pd.read_excel(file_path, sheet_name='IPO MRBR')
    df_imac = pd.read_excel(file_path, sheet_name='IMAC MRBR')

    df_imac_filtered = filter_dataframe(df_imac, None, desired_values)
    df_site_filtered = filter_dataframe(df_site, None, desired_values)
    df_ipo_filtered = filter_dataframe(df_ipo, '  CoCd', desired_values)

    data_frames = [df_ipo_filtered, df_site_filtered, df_imac_filtered]
    sheet_names = ["IPO MRBR", "SITE MRBR", "IMAC MRBR"]

    with pd.ExcelWriter(new_excel_server) as writer:
        for i, df in enumerate(data_frames):
            df.to_excel(writer, sheet_name=sheet_names[i], index=False)

    print("MRBR files successfully processed and saved!")

if __name__ == "__main__":
    try:
        main()
        input("Complete press Enter to exit")
    except Exception:
        import traceback
        traceback.print_exc()
        input("Program crashed; press Enter to exit")
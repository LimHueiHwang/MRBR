```python
import pandas as pd
from datetime import datetime


# ============================================================
# Configuration
# ============================================================

MASTER_SERVER_PATH = r"<MASTER_SERVER_PATH>"
IMAC_OUTPUT_PATH = r"<IMAC_OUTPUT_PATH>"


# ============================================================
# Date and File Path Functions
# ============================================================

def get_todays_date():
    today = datetime.today()
    year = today.strftime("%Y")
    date_str = today.strftime("%y%m%d")
    return year, date_str


def get_file_path(file_name, path_type, today=None, year=None):
    if path_type == "server_imac":
        if not today or not year:
            raise ValueError(
                "'today' and 'year' must be provided for server_imac path_type."
            )

        return f"{IMAC_OUTPUT_PATH}/{year}/MRBR {today}.xlsx"

    if path_type == "server_master":
        return f"{MASTER_SERVER_PATH}/{file_name}.xlsx"

    raise ValueError(f"Unsupported path_type: {path_type}")


# ============================================================
# Data Processing
# ============================================================

def filter_dataframe(df, col_name, desired_values):
    df = df.copy()

    df["BUYER Comment"] = ""
    df["Elaine approval"] = ""
    df["SP Update"] = ""

    if col_name is not None:
        df[col_name] = df[col_name].astype(str)
        df = df[~df[col_name].str.contains("1803")]

    df = df.loc[df["Plnt"].isin(desired_values)]

    return df


# ============================================================
# Main Process
# ============================================================

def main():
    file_name = input("Enter file name: ")

    file_path = get_file_path(
        file_name,
        path_type="server_master"
    )

    today_year, today_date = get_todays_date()

    new_excel_server = get_file_path(
        file_name,
        today=today_date,
        year=today_year,
        path_type="server_imac"
    )

    desired_values = [
        "HU07",
        "HU08",
        "IN07",
        "IT08",
        "PL01",
        "SG02",
        "VN01"
    ]

    df_site = pd.read_excel(
        file_path,
        sheet_name="SITE MRBR"
    )

    df_ipo = pd.read_excel(
        file_path,
        sheet_name="IPO MRBR"
    )

    df_imac = pd.read_excel(
        file_path,
        sheet_name="IMAC MRBR"
    )

    df_imac_filtered = filter_dataframe(
        df_imac,
        None,
        desired_values
    )

    df_site_filtered = filter_dataframe(
        df_site,
        None,
        desired_values
    )

    df_ipo_filtered = filter_dataframe(
        df_ipo,
        "  CoCd",
        desired_values
    )

    data_frames = [
        df_ipo_filtered,
        df_site_filtered,
        df_imac_filtered
    ]

    sheet_names = [
        "IPO MRBR",
        "SITE MRBR",
        "IMAC MRBR"
    ]

    with pd.ExcelWriter(new_excel_server) as writer:
        for df, sheet_name in zip(data_frames, sheet_names):
            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    print("MRBR files successfully processed and saved!")


if __name__ == "__main__":
    try:
        main()
        input("Complete. Press Enter to exit.")

    except Exception:
        import traceback

        traceback.print_exc()
        input("Program crashed; press Enter to exit.")
```

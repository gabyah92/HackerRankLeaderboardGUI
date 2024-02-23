# Students were instructed to use their unique college ID as their username.
# This ID is of the format: PES0UG00XX000, where 0 is any digit and X is any letter.
# Some students have _s in between their ID, and I'll allow it.

import re
import os
import pandas as pd


def find_id_in_username(username):
    username = username.replace("_", "")
    username = username.strip()
    username = username.upper()

    pattern = r"PES\d{1}UG\d{2}[A-Z]{2}\d{3}"
    matches = re.findall(pattern, username)

    if len(matches) == 0 or len(matches) > 1:
        return "SRN not found in username."

    return matches[0]

def get_all_excel_filenames(dir, ignore_files = []):
    excel_files = []
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith(".xlsx") and file not in ignore_files:
                excel_files.append(os.path.join(root, file))
    return excel_files


def main():

    filenames = get_all_excel_filenames("Leaderboards", ignore_files=["TotalHackerrankLeaderBoard.xlsx"])

    results_dir = ".\\TestResults"
    os.makedirs(results_dir, exist_ok=True)

    for filename in filenames:
        df = pd.read_excel(filename, engine="openpyxl")
        df["SRN"] = df["Name"].apply(find_id_in_username)
        new_filename = results_dir + "\\" + filename.split("\\")[-1].replace(".xlsx", ".csv")
        df.to_csv(new_filename, index=False)

    final_results_dir = ".\\FinalResults"
    os.makedirs(final_results_dir, exist_ok=True)

    sections_path = ".\\Sections"
    for root, dirs, files in os.walk(sections_path):
        for file in files:
            if file.endswith(".csv"):
                section_name = file.removesuffix(".csv")
                section_df = pd.read_csv(os.path.join(root, file))

                for root_, dirs_, files_ in os.walk(results_dir):
                    for file_ in files_:
                        if file_.endswith(".csv"):
                            test_name = file_.removesuffix(".csv")
                            test_res = pd.read_csv(os.path.join(root_, file_))
                
                            section_df = section_df.merge(test_res, left_on="SRN", right_on="Name", how="left").rename(columns={"Score": test_name, "Name_x": "Name", "SRN_x": "SRN"}).drop(columns=["Name_y", "SRN_y", "Rank"])
                
                section_df.to_csv(final_results_dir + "\\" + section_name + ".csv", index=False)

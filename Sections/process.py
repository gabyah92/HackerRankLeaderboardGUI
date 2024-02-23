# replace \t with , in all files in this folder
# created this file just to clean up the section details files, not necessary if that file is already good

path = ".\\Sections"

import os
import pandas as pd

for root, dirs, files in os.walk(path):
    for file in files:
        if file.endswith(".csv"):
            with open(os.path.join(root, file), "r") as f:
                data = f.read()
            data = data.replace("\t", ",")
            with open(os.path.join(root, file), "w") as f:
                f.write(data)
            print(f"Replaced tabs with commas in {file}")

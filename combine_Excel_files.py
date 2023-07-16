import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
# file_path = filedialog.askopenfilename()
folder_path = filedialog.askdirectory()
print(folder_path)


path = folder_path
files = os.listdir(path)

df = pd.DataFrame()
df_new = pd.DataFrame()

for file in files:
    if file.endswith('.xlsx'):
        #print(file)
        df_new = pd.read_excel(file)
        print(file)
        #print(df)
        #print(df.shape)
        df = pd.concat([df,df_new])

df.to_excel('combined_file.xlsx')
import os
import string
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
root = os.path.split(file_path)

# create folder to hold outputs
plot_output = root[0] + '/plots'
if not os.path.exists(plot_output):
    os.makedirs(plot_output)

# load Excel file
xls_raw = pd.ExcelFile(file_path)

# empty dictionary to store DataFrames
dfs_raw = {}
for sheet_name in xls_raw.sheet_names:
    df_raw = pd.read_excel(xls_raw, sheet_name=sheet_name)
    dfs_raw[sheet_name] = df_raw

df_list = list(dfs_raw.keys())

for i in range(len(df_list)):
    full_df = dfs_raw[df_list[i]]
    full_df['mod Image name'] = full_df['Image name'].str[:6] + full_df['Image name'].str.extract(r'slice(.*?)$')[0]
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip('.oir')
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip(string.digits)
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip('_')

    # Pivot the new dataframe based on ROIs
    pivot_df = full_df.pivot_table(index=['mod Image name', 'Image feature', 'ROI #'],
                                    columns='Timepoint', values='Gray value average', fill_value=0)
    # Reset index to remove multi-index
    pivot_df = pivot_df.reset_index()
    pivot_df = pivot_df.replace(0, np.nan)

    pivot_filename = root[1].rstrip('.xlsx') + df_list[i] + ('_.xlsx')
    output_path = root[0] + '/plots/' + pivot_filename
    pivot_df.to_excel(output_path, index=True)
    print(f'DataFrame saved as: {output_path}')

    roi_type = {
        'feature1': 'cell body',
        'feature2': 'axon',
        'feature3': 'background'
    }

    feature_dfs = {}
    for key, feature in roi_type.items():
        df = pivot_df.loc[pivot_df['Image feature'].str.contains(feature),
            ['mod Image name', 'ROI #', 'native fluorescence', 'DMSO quench', 'post stain']]
        feature_dfs[feature] = df

        plt.figure(figsize=(4, 6))
        plt.scatter(np.zeros(len(df['DMSO quench'])), df['DMSO quench'])
        plt.scatter(np.ones(len(df['post stain'])), df['post stain'])

        # plotting the lines
        for j in range(len(df['post stain'])):
            plt.plot([0,1], [df['DMSO quench'].iloc[j], df['post stain'].iloc[j]], c='k')

        plot_title = df_list[i] + ' ' + feature
        plt.title(plot_title)
        plt.ylabel('Fluorescence intensity, a.u.')
        plt.xlim([-1, 2])
        plt.xticks([0,1], ['DMSO quench', 'post stain'], rotation=45)
        plt.tight_layout()
        plt.show()


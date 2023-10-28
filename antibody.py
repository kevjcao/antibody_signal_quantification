import os
import string
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# select file dialog
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
root = os.path.split(file_path)

# create folder in same directory as Excel sheet to hold outputs
plot_output = root[0] + '/plots'
if not os.path.exists(plot_output):
    os.makedirs(plot_output)

# load Excel file
xls_raw = pd.ExcelFile(file_path)

# empty dictionary to store DataFrames of each sheet in Excel file
dfs_raw = {}
for sheet_name in xls_raw.sheet_names:
    df_raw = pd.read_excel(xls_raw, sheet_name=sheet_name)
    dfs_raw[sheet_name] = df_raw

# create a list of the sheets
df_list = list(dfs_raw.keys())

for i in range(len(df_list)):
    # create dataframe of the active sheet
    full_df = dfs_raw[df_list[i]]
    # unify the data by 'Image name' column (remove .oir extension and image sequence #s)
    full_df['mod Image name'] = full_df['Image name'].str[:6] + full_df['Image name'].str.extract(r'slice(.*?)$')[0]
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip('.oir')
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip(string.digits)
    full_df['mod Image name'] = full_df['mod Image name'].str.rstrip('_')

    # pivot the unified dataframe based on ROIs - output: df with columns for each 'Timepoint' for a given ROI
    pivot_df = full_df.pivot_table(index=['mod Image name', 'Image feature', 'ROI #'],
                                    columns='Timepoint', values='Gray value average', fill_value=0)

    # reset index to remove multi-index, replace 0s (usually data missing) w/ NaNs
    pivot_df = pivot_df.reset_index()
    pivot_df = pivot_df.replace(0, np.nan)

    # save pivot_df as individual Excel file in output folder
    pivot_filename = root[1].rstrip('.xlsx') + df_list[i] + ('_.xlsx')
    pivot_output_path = root[0] + '/plots/' + pivot_filename
    pivot_df.to_excel(pivot_output_path, index=True)
    print(f'DataFrame saved as: {pivot_output_path}')

    # dictionary of ROI types
    roi_type = {
        'feature1': 'cell body',
        'feature2': 'axon',
        'feature3': 'background'
    }

    # create a dictionary of dataframes for each ROI type
    feature_dfs = {}
    fig, axes = plt.subplots(1, 3, figsize=(10, 6))
    for n, (key, feature) in enumerate(roi_type.items()):
        # create a dataframe for the active ROI type in the for loop and add to the features_dfs dictionary
        df = pivot_df.loc[pivot_df['Image feature'].str.contains(feature),
            ['mod Image name', 'ROI #', 'native fluorescence', 'DMSO quench', 'post stain']]
        feature_dfs[feature] = df

        # calc. stats of each dataframe column
        df_stats = df.describe()

        # plot a before/after plot for each ROI type, only plot quenched v. post stain
        ax = axes[n]
        ax.scatter(np.zeros(len(df['DMSO quench'])), df['DMSO quench'], s=40, facecolors='none', edgecolors='xkcd:grey')
        ax.scatter(np.ones(len(df['post stain'])), df['post stain'], s=40, facecolors='none', edgecolors='xkcd:grey')

        # Plot mean and standard deviation
        ax.errorbar([0, 1], df_stats.loc['mean', ['DMSO quench', 'post stain']],
                    yerr=df_stats.loc['std', ['DMSO quench', 'post stain']], fmt='.', linewidth=0.5, c='xkcd:red pink',
                    label='Mean +/- SD')

        # Add text annotation for mean value next to the mean data point
        ax.text(-0.1, df_stats.loc['mean', 'DMSO quench'], f'{df_stats.loc["mean", "DMSO quench"]:.2f}', ha='right',
                va='center', c='xkcd:red pink')
        ax.text(1.1, df_stats.loc['mean', 'post stain'], f'{df_stats.loc["mean", "post stain"]:.2f}', ha='left',
                va='center', c='xkcd:red pink')

        # formatting
        plot_title = df_list[i] + ' ' + feature
        ax.set_title(plot_title)
        ax.set_ylabel('Fluorescence intensity, a.u.')
        ax.set_xlim([-0.75, 1.75])
        ax.set_xticks([0,1], ['DMSO quench', 'post stain'], rotation=45)
        ax.legend()

        # connect paired datapoints
        for j in range(len(df['post stain'])):
            ax.plot([0, 1], [df['DMSO quench'].iloc[j], df['post stain'].iloc[j]], linewidth=1, c='xkcd:grey')

    plt.tight_layout()
    fig_output_path = plot_output + '/' + df_list[i] + '.png'
    plt.savefig(fig_output_path, dpi=300, bbox_inches='tight')
    plt.show()

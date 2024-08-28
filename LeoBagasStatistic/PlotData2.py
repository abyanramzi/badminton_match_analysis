import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Load the spreadsheet file
file_path = 'LeoBagasStatistic\JapanOpen_LeoBagas.xlsx'
df = pd.read_excel(file_path)

# Initialize a dictionary to store data by set
data_by_set = {}

# Extract and store the data for each set
for number_of_set in df['Set'].unique():
    filtered_df = df[df['Set'] == number_of_set]
    data_by_set[number_of_set] = {
        'Total Points': filtered_df['Numbers of Last Hit'].tolist(),
        'Leo/Bagas Points': filtered_df['Leo/Bagas Points'].tolist(),
        'Executor': filtered_df['By'].tolist(),
        'Point Status': filtered_df['Point Status'].tolist(),
        'Hit Type': filtered_df['Type of Last Hit'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("flare", 3)

# Plotting the data for each set
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    leobagas_point = data_by_set[set_number]['Leo/Bagas Points']
    executor = data_by_set[set_number]["Executor"]
    hit_type = data_by_set[set_number]['Hit Type']
    point_status = data_by_set[set_number]['Point Status']

    leohit = []
    bagashit = []

    # Collect hit types for Leo and Bagas
    for i in range(len(executor)):
        exec = executor[i]
        current_hit_type = hit_type[i]
        current_point_status = point_status[i]

        if exec == "Leo" and current_point_status != "Fail":
            leohit.append(current_hit_type)
        elif exec == "Bagas" and current_point_status != "Fail":
            bagashit.append(current_hit_type)

    # Convert lists to Pandas Series and calculate value counts
    leohit_series = pd.Series(leohit)
    bagashit_series = pd.Series(bagashit)
    
    count_leohit = leohit_series.value_counts()
    count_bagashit = bagashit_series.value_counts()

    # Create a common index for both Leo and Bagas for the bar chart
    hit_types = sorted(set(count_leohit.index).union(set(count_bagashit.index)))
    index = np.arange(len(hit_types))

    # Create a figure and axis
    plt.figure(figsize=(12, 6))
    plt.style.use('dark_background')  # Set a dark background style

    # Plot bar charts for Leo and Bagas points
    bar_width = 0.4
    
    plt.bar(index - bar_width/2, count_leohit.reindex(hit_types, fill_value=0), width=bar_width, label='Leo Hits', color=colors[0])
    plt.bar(index + bar_width/2, count_bagashit.reindex(hit_types, fill_value=0), width=bar_width, label='Bagas Hits', color=colors[1])

    # Adding titles and labels
    plt.title(f'Set {set_number} Hit Type Comparison', fontsize=16, fontweight='bold', color='white')
    plt.xlabel('Hit Type', fontsize=12, color='white')
    plt.ylabel('Count', fontsize=12, color='white')
    plt.xticks(index, hit_types, color='white')  # Set x-ticks to be hit types
    plt.legend(loc='upper left', fontsize=10)

    # Customize the grid
    plt.grid(True, linestyle='--', alpha=0.5)

    # Set grid lines to appear every 2 points
    plt.yticks(np.arange(0, max(count_leohit.max(), count_bagashit.max())+3, 1), color='white')

    # Set the background color
    plt.gca().set_facecolor('black')  # Black background for the plot area
    plt.gcf().patch.set_facecolor('black')  # Black background for the figure

    # Adjust plot layout
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    plt.savefig('LeoBagasStatistic/Images/' + f'set_{set_number}_type_hit.png', bbox_inches='tight', dpi=300)
    plt.close()
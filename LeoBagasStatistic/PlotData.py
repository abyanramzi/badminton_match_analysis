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
        'Kenya/Hiroki Points': filtered_df['Kenya/Hiroki Points'].tolist(),
        'Executor': filtered_df['By'].tolist(),
        'Point Status': filtered_df['Point Status'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("flare", 3)

# Plotting the data for each set
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    leobagas_point = data_by_set[set_number]['Leo/Bagas Points']
    kenyahiroki_point = data_by_set[set_number]['Kenya/Hiroki Points']
    executor = data_by_set[set_number]['Executor']
    status = data_by_set[set_number]['Point Status']

    leopoint = []
    bagaspoint = []
    leo_cumulative = 0
    bagas_cumulative = 0

    # Calculate cumulative points for Leo and Bagas
    for i in range(len(executor)):
        exec = executor[i]
        point_status = status[i]
    
        if exec == "Leo" and point_status != "Fail":
            leo_cumulative += 1
        elif exec == "Bagas" and point_status != "Fail":
            bagas_cumulative += 1

        # Append cumulative points to lists inside the loop
        leopoint.append(leo_cumulative)
        bagaspoint.append(bagas_cumulative)

    # Create a figure and axis
    plt.figure(figsize=(12, 6))
    plt.style.use('dark_background')  # Set a dark background style

    # Plot bar charts for Leo and Bagas points
    bar_width = 0.4
    index = np.arange(len(total_point))
    
    plt.bar(index + bar_width/2 , leopoint, width=bar_width, label='Leo Points', color=colors[0])
    plt.bar(index - bar_width/2 , bagaspoint, width=bar_width, label='Bagas Points', color=colors[1])
    plt.plot(index, leobagas_point, label='Leo/Bagas Points', color=colors[2], marker='o', linewidth=2)

    # Adding titles and labels
    plt.title(f'Set {set_number} Performance', fontsize=16, fontweight='bold', color='white')
    plt.xlabel('Total Points', fontsize=12, color='white')
    plt.ylabel('Points', fontsize=12, color='white')
    plt.xticks(index, total_point, color='white')
    plt.legend(loc='upper left', fontsize=10)

    # Customize the grid
    plt.grid(True, linestyle='--', alpha=0.5)

    # Set grid lines to appear every 2 points
    plt.xticks(np.arange(min(index), max(index)+3 , 2), color='white')
    plt.yticks(np.arange(0, max(leobagas_point)+3 , 2), color='white')

    # Set the background color
    plt.gca().set_facecolor('black')  # Black background for the plot area
    plt.gcf().patch.set_facecolor('black')  # Black background for the figure

    # Adjust plot layout
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    plt.savefig('LeoBagasStatistic/Images/' + f'set_{set_number}_performance.png', bbox_inches='tight', dpi=300)
    plt.close()

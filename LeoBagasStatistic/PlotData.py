import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Load the spreadsheet file
file_path = 'LeoBagasStatistic\KoreaOpen_LeoBagas.xlsx'
df = pd.read_excel(file_path)

# Initialize a dictionary to store data by set
data_by_set = {}

# Extract and store the data for each set
for number_of_set in df['Set'].unique():
    filtered_df = df[df['Set'] == number_of_set]
    data_by_set[number_of_set] = {
        'Total Points': filtered_df['Numbers of Last Hit'].tolist(),
        'Leo/Bagas Points': filtered_df['Leo/Bagas Points'].tolist(),
        'Kang/Seo Points': filtered_df['Kang/Seo Points'].tolist(),
        'Executor': filtered_df['By'].tolist(),
        'Point Status': filtered_df['Point Status'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("BrBG", 10)
colors2 = sns.color_palette("Reds", 10)

# Plotting the data for each set
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    leobagas_point = data_by_set[set_number]['Leo/Bagas Points']
    kangseo_point = data_by_set[set_number]['Kang/Seo Points']
    executor = data_by_set[set_number]['Executor']
    status = data_by_set[set_number]['Point Status']

    leopoint = []
    bagaspoint = []
    leototal = bagastotal = 0

    for i in range(len(total_point)):
        exec = executor[i]
        sts = status[i]

        if exec == "Leo" and sts != "Fail":
          leototal += 1
        elif exec == "Bagas" and sts != "Fail":
          bagastotal += 1
      
        leopoint.append(leototal)
        bagaspoint.append(bagastotal)
    
    # Create a figure and axis
    plt.figure(figsize=(12, 6))
    plt.style.use('dark_background')  # Set a dark background style

    # Plot each line
    plt.plot(total_point, leobagas_point, label='Leo/Bagas Points', color=colors[5], marker='o', linewidth=2)
    plt.bar(total_point, leopoint, label='Leo Point', color=colors[2])
    plt.bar(total_point, bagaspoint, bottom=leopoint, label='Bagas Point', color=colors[8])

    plt.plot(total_point, kangseo_point, label='Kang/Seo Points', color=colors2[5], marker='o', linewidth=2)

    if max(kangseo_point) > max(leobagas_point):
      plt.plot(max(total_point), max(kangseo_point), color = "yellow", marker = "*", markersize=12)
    if max(leobagas_point) > max(kangseo_point):
      plt.plot(max(total_point), max(leobagas_point), color = "yellow", marker = "*", markersize=12)

    # Adding titles and labels
    plt.title(f'Set {int(set_number)} Performance', fontsize=16, fontweight='bold', color='white')
    plt.xlabel('Total Points', fontsize=12, color='white')
    plt.ylabel('Points', fontsize=12, color='white')
    plt.legend(loc='upper left', fontsize=10)

    # Customize the grid
    plt.grid(True, linestyle='--', alpha=0.5)

    # Set grid lines to appear every 2 points
    plt.xticks(np.arange(min(total_point), max(total_point) + 3, 2), color='white')
    plt.yticks(np.arange(0, max(max(leobagas_point), max(kangseo_point)) + 5, 2), color='white')

    # Set the background color
    plt.gca().set_facecolor('black')  # Black background for the plot area
    plt.gcf().patch.set_facecolor('black')  # Black background for the figure

    # Adjust plot layout
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    # Save plot
    plt.savefig('LeoBagasStatistic/Images/' + f'Set {int(set_number)} Performance Stats.png', bbox_inches='tight', dpi=300)
    plt.close()
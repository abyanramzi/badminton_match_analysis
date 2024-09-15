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
        'Executor': filtered_df['By'].tolist(),
        'Point Status': filtered_df['Point Status'].tolist(),
        'Hit Type': filtered_df['Type of Last Hit'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("PuBu", 10)

# Step 1: Determine the global maximum value for the x-axis limit across all sets
global_max_hits = 0
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    executor = data_by_set[set_number]['Executor']
    hit_type = data_by_set[set_number]['Hit Type']

    hits_leo = {'Drive': 0, 'Smash': 0, 'Push Shot': 0, 'Lift': 0, 'Netting': 0, 'Serve': 0, 'Block': 0, 'Net Kill': 0, 'Clear': 0}
    hits_bagas = {'Drive': 0, 'Smash': 0, 'Push Shot': 0, 'Lift': 0, 'Netting': 0, 'Serve': 0, 'Block': 0, 'Net Kill': 0, 'Clear': 0}

    # Count hit types for each player
    for i in range(len(total_point)):
        exec = str(executor[i])
        current_hit_type = hit_type[i]

        if exec == "Leo":
            if current_hit_type in hits_leo:
                hits_leo[current_hit_type] += 1
        elif exec == "Bagas":
            if current_hit_type in hits_bagas:
                hits_bagas[current_hit_type] += 1

    # Update the global maximum value of hits
    max_hits_in_set = max(max(hits_leo.values()), max(hits_bagas.values()))
    global_max_hits = max(global_max_hits, max_hits_in_set)

# Step 2: Plot the data for each set using the same x-axis limit
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    executor = data_by_set[set_number]['Executor']
    point_status = data_by_set[set_number]['Point Status']
    hit_type = data_by_set[set_number]['Hit Type']

    # Initialize hit counts for each player
    hits_leo = {'Drive': 0, 'Smash': 0, 'Push Shot': 0, 'Lift': 0, 'Netting': 0, 'Serve': 0, 'Block': 0, 'Net Kill': 0, 'Clear': 0}
    hits_bagas = {'Drive': 0, 'Smash': 0, 'Push Shot': 0, 'Lift': 0, 'Netting': 0, 'Serve': 0, 'Block': 0, 'Net Kill': 0, 'Clear': 0}

    # Initialize cumulative points
    leopoint = 0
    bagaspoint = 0

    # Count hit types and points for Leo and Bagas
    for i in range(len(total_point)):
        exec = str(executor[i])
        current_hit_type = hit_type[i]
        current_point_status = point_status[i]

        if current_point_status != "Fail":
            if exec == "Leo":
                if current_hit_type in hits_leo:
                    hits_leo[current_hit_type] += 1
                leopoint += 1
            elif exec == "Bagas":
                if current_hit_type in hits_bagas:
                    hits_bagas[current_hit_type] += 1
                bagaspoint += 1
                
    # Create a figure and axis
    plt.figure(figsize=(18, 3))
    plt.style.use('dark_background')  # Set a dark background style

    # Create the list of hit types and their counts for both players
    hit_types = list(hits_leo.keys())
    leo_hits = list(hits_leo.values())
    bagas_hits = list(hits_bagas.values())

    # Convert Bagas hits to negative values for diverging bar chart
    bagas_hits_neg = [-x for x in bagas_hits]

    # Plot the bars for Leo and Bagas (hit types)
    bar_width = 0.8
    indices = np.arange(len(hit_types))  # Create a list of positions for the bars

    leo_bars = plt.barh(indices, leo_hits, height=bar_width, label="Leo Hit Types", color=colors[9])
    bagas_bars = plt.barh(indices, bagas_hits_neg, height=bar_width, label="Bagas Hit Types", color=colors[7])

    # Add annotations (hit values) on the bars
    for i, (leo_val, bagas_val) in enumerate(zip(leo_hits, bagas_hits)):
        # Place hit type names in the center between Leo and Bagas bars
        plt.text(0, i, hit_types[i], va='center', ha='center', color=colors[0], fontsize=12, fontweight='bold')
        if leo_val != 0:
            # Annotate Leo bars
            plt.text(leo_val - 0.05, i, f'{leo_val}', va='center', ha='right', color='white', fontsize=12, fontweight='bold')
        if bagas_val != 0:
            # Annotate Bagas bars (note the use of bagas_hits_neg for placement)
            plt.text(bagas_hits_neg[i] + 0.05, i, f'{bagas_val}', va='center', ha='left', color='white', fontsize=12, fontweight='bold')
        if leo_val == 0:
            # Annotate Leo bars
            plt.text(leo_val + 0.7, i, f'{leo_val}', va='center', ha='right', color='white', fontsize=12, fontweight='bold')
        if bagas_val == 0:
            # Annotate Bagas bars (note the use of bagas_hits_neg for placement)
            plt.text(bagas_hits_neg[i] - 0.7, i, f'{bagas_val}', va='center', ha='left', color='white', fontsize=12, fontweight='bold')

    # Adding titles and labels
    plt.title(f'Set {int(set_number)}: Diverging Hit Types & Points Comparison for Leo and Bagas', fontsize=16, fontweight='bold', color='white')

    # Step 3: Set the x-axis limits to the global maximum value for consistency
    plt.xlim([-global_max_hits - 1, global_max_hits + 1])  # Use global maximum value for all sets

    # Hide both x-axis and y-axis
    plt.xticks([])  # Hide x-axis ticks
    plt.yticks([])  # Hide y-axis ticks

    # Set background color
    plt.gca().set_facecolor('black')  # Black background for the plot area
    plt.gcf().patch.set_facecolor('black')  # Black background for the figure

    # Adjust plot layout
    plt.tight_layout()

    plt.savefig('LeoBagasStatistic/Images/' + f'set_{int(set_number)}_type_hit.png', bbox_inches='tight', dpi=300)
    plt.close()
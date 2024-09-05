import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.patches import Polygon
from matplotlib.colors import Normalize
from mpl_toolkits.axes_grid1.inset_locator import inset_axes

# Load the spreadsheet file
file_path = 'LeoBagasStatistic\KoreaOpen_LeoBagas.xlsx'
df = pd.read_excel(file_path)

length = 1080

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
        'Point Status': filtered_df['Point Status'].tolist(),
        'Zone': filtered_df['Zone'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("dark", 10)

# Plotting the data for each set
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    leobagas_point = data_by_set[set_number]['Leo/Bagas Points']
    kangseo_point = data_by_set[set_number]['Kang/Seo Points']
    executor = data_by_set[set_number]['Executor']
    status = data_by_set[set_number]['Point Status']
    zone = data_by_set[set_number]['Zone']

    # ---- Performance Graph (Main Plot) ----
    fig, ax = plt.subplots(figsize=(14, 8))
    plt.style.use('dark_background')  # Set a dark background style

    # Plot each line
    ax.plot(total_point, leobagas_point, label='Leo/Bagas Points', color=colors[3], marker='o', linewidth=3)
    ax.plot(total_point, kangseo_point, label='Kang/Seo Points', color=colors[9], marker='o', linewidth=3)

    # Adding titles and labels
    ax.set_title(f'Leo/Bagas Set {set_number} Performance', fontsize=16, fontweight='bold', color='white')
    ax.set_xlabel('Total Points', fontsize=12, color='white')
    ax.set_ylabel('Points', fontsize=12, color='white')
    ax.legend(loc='lower right', fontsize=10)

    # Customize the grid
    ax.grid(True, linestyle='--', alpha=0.3)

    # Set grid lines to appear every 2 points
    ax.set_xticks(np.arange(0, max(total_point) + 3, 2))
    ax.set_yticks(np.arange(0, max(max(leobagas_point), max(kangseo_point)) + 5, 2))
    ax.tick_params(colors='white')

    # Set the background color
    ax.set_facecolor('black')

    # ---- Distribution Plot (Inset Plot) ----
    # Create an inset axes
    ax_inset = inset_axes(ax, width="50%", height="50%", loc="upper left", borderpad = 0, bbox_to_anchor=(0.01, 0.04, 1, 1), bbox_transform=ax.transAxes)
    # Draw the badminton court (Court drawing inside the inset)
    ax_inset.set_facecolor('black')
    ax_inset.set_aspect('equal')
    ax_inset.axis('off')

    # Calculate zone percentages
    zone1total = zone2total = zone3total = 0

    for i in range(len(total_point)):
        if status[i] != "Fail":
            if zone[i] == 1:
                zone1total += 1
            elif zone[i] == 2:
                zone2total += 1
            elif zone[i] == 3:
                zone3total += 1

    total_points = zone1total + zone2total + zone3total

    zone1_percentage = (zone1total / total_points) * 100 if total_points > 0 else 0
    zone2_percentage = (zone2total / total_points) * 100 if total_points > 0 else 0
    zone3_percentage = (zone3total / total_points) * 100 if total_points > 0 else 0
    zoneOp_percentage = 0

    # Define the zones (Court coordinates)
    zones = {
        "Zone 1": [(522, 50), (720, 50), (720, 660), (522, 660)],
        "Zone 2": [(50, 355), (522, 355), (522, 660), (50, 660)],
        "Zone 3": [(50, 50), (522, 50), (522, 355), (50, 355)],
        "Zone Op": [(720,50), (1390,50), (1390,660), (720,660)]
    }

    zone_centers = {
        "Zone 1": (sum([p[0] for p in zones["Zone 1"]]) / len(zones["Zone 1"]),
                   sum([p[1] for p in zones["Zone 1"]]) / len(zones["Zone 1"])),
        "Zone 2": (sum([p[0] for p in zones["Zone 2"]]) / len(zones["Zone 2"]),
                   sum([p[1] for p in zones["Zone 2"]]) / len(zones["Zone 2"])),
        "Zone 3": (sum([p[0] for p in zones["Zone 3"]]) / len(zones["Zone 3"]),
                   sum([p[1] for p in zones["Zone 3"]]) / len(zones["Zone 3"])),
        "Zone Op": (sum([p[0] for p in zones["Zone Op"]]) / len(zones["Zone Op"]),
                   sum([p[1] for p in zones["Zone Op"]]) / len(zones["Zone Op"]))
    }

    # Create a color map and normalize
    cmap = plt.cm.Reds  # Red gradient colormap
    norm = Normalize(vmin=0, vmax=100)  # Normalize to percentage range

    # Plot each zone with gradient color based on percentage
    zone_percentages = [zone1_percentage, zone2_percentage, zone3_percentage, 0]
    zone_point = [zone1total, zone2total, zone3total, 0]
    for zone_name, vertices, percentage, total in zip(zones.keys(), zones.values(), zone_percentages, zone_point):
        if zone_name == "Zone Op":
            color = "black"
        else :
            color = cmap(norm(percentage))
        polygon = Polygon(vertices, closed=True, color=color)
        ax_inset.add_patch(polygon)

        # Annotate the zone with percentage
        if zone_name != "Zone Op":
            center = zone_centers[zone_name]
            # ax_inset.text(center[0], center[1], f'{total}' + " Points\n" +f'{percentage:.1f}%',
            #               ha='center', va='center', fontsize=10, color="black", weight="bold")
            # ax_inset.text(center[0], center[1], f'{percentage:.1f}%',
            #               ha='center', va='center', fontsize=10, color="black", weight="bold")
            ax_inset.text(center[0], center[1], f'{total}' + "\nPoints",
                          ha='center', va='center', fontsize=10, color="black", weight="bold")

    # Draw court lines
    coordinates = {
        "line 1": [(50, 50), (50, 96), (50, 355), (50, 614), (50, 660)],
        "line 2": [(126, 50), (126, 96), (126, 355), (126, 614), (126, 660)],
        "line 3": [(522, 50), (522, 96), (522, 355), (522, 614), (522, 660)],
        "line 4": [(720, 50), (720, 96), (720, 355), (720, 614), (720, 660)],
        "line 5": [(918, 50), (918, 96), (918, 355), (918, 614), (918, 660)],
        "line 6": [(1314, 50), (1314, 96), (1314, 355), (1314, 614), (1314, 660)],
        "line 7": [(1390, 50), (1390, 96), (1390, 355), (1390, 614), (1390, 660)]
    }

    for i in range(5):  # Iterate over each coordinate
        for name, coords in coordinates.items():
            if name == "line 1":
                if i != 2:  # Exclude third coordinate of line 1
                    ax_inset.plot([coords[i][0], coordinates["line 7"][i][0]], 
                                  [coords[i][1], coordinates["line 7"][i][1]], color='white')
                else:  # Connect third coordinate of line 1 with line 3
                    ax_inset.plot([coords[i][0], coordinates["line 3"][i][0]], 
                                  [coords[i][1], coordinates["line 3"][i][1]], color='white')
            elif name == "line 5" and i == 2:  # Connect third coordinate of line 5 with line 7
                ax_inset.plot([coords[i][0], coordinates["line 7"][i][0]], 
                              [coords[i][1], coordinates["line 7"][i][1]], color='white')

        # Connect first and last coordinates for every line
        for name, coords in coordinates.items():
            ax_inset.plot([coords[0][0], coords[-1][0]], 
                          [coords[0][1], coords[-1][1]], color='white')

    # Calculate aspect ratio and compute width
    aspect_ratio = fig.get_size_inches()[1] / fig.get_size_inches()[0]  # height / width
    width = length / aspect_ratio  # Calculate width

    # Save the figure with the computed dimensions
    fig.set_size_inches(width / 100, length / 100)  # Convert pixels to inches (100 dpi)
    plt.savefig('LeoBagasStatistic/Images/' + f'LeoBagas Point Distribution Set_{set_number}.png', bbox_inches='tight', pad_inches=0.1, dpi=100)

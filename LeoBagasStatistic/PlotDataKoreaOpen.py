import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable

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
        'Point Status': filtered_df['Point Status'].tolist(),
        'Zone': filtered_df['Zone'].tolist()
    }

for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Points']
    point_status = data_by_set[set_number]['Point Status']
    executor = data_by_set[set_number]['Executor']
    zone = data_by_set[set_number]['Zone']

    zone1total = 0
    zone2total = 0
    zone3total = 0

    for i in range(len(total_point)):
        Status = point_status[i]
        Zone = zone[i]

        if Status != "Fail":
            if Zone == 1:
                zone1total += 1
            elif Zone == 2:
                zone2total += 1
            elif Zone == 3:
                zone3total += 1

    total_points = zone1total + zone2total + zone3total

    zone1_percentage = (zone1total / total_points) * 100 if total_points > 0 else 0
    zone2_percentage = (zone2total / total_points) * 100 if total_points > 0 else 0
    zone3_percentage = (zone3total / total_points) * 100 if total_points > 0 else 0

    print(f"Total Point Zone 1: {zone1total} ({zone1_percentage:.2f}%)")
    print(f"Total Point Zone 2: {zone2total} ({zone2_percentage:.2f}%)")
    print(f"Total Point Zone 3: {zone3total} ({zone3_percentage:.2f}%)")

    # Define the coordinates
    coordinates = {
        "line 1": [(50, 50), (50, 96), (50, 355), (50, 614), (50, 660)],
        "line 2": [(126, 50), (126, 96), (126, 355), (126, 614), (126, 660)],
        "line 3": [(522, 50), (522, 96), (522, 355), (522, 614), (522, 660)],
        "line 4": [(720, 50), (720, 96), (720, 355), (720, 614), (720, 660)],
        "line 5": [(918, 50), (918, 96), (918, 355), (918, 614), (918, 660)],
        "line 6": [(1314, 50), (1314, 96), (1314, 355), (1314, 614), (1314, 660)],
        "line 7": [(1390, 50), (1390, 96), (1390, 355), (1390, 614), (1390, 660)]
    }

    # Define the zones
    zones = {
        "Zone 1": [(522, 50), (720, 50), (720, 660), (522, 660)],
        "Zone 2": [(50, 355), (522, 355), (522, 660), (50, 660)],
        "Zone 3": [(50, 50), (522, 50), (522, 355), (50, 355)]
    }

    zone_centers = {
        "Zone 1": (sum([p[0] for p in zones["Zone 1"]]) / len(zones["Zone 1"]),
                   sum([p[1] for p in zones["Zone 1"]]) / len(zones["Zone 1"])),
        "Zone 2": (sum([p[0] for p in zones["Zone 2"]]) / len(zones["Zone 2"]),
                   sum([p[1] for p in zones["Zone 2"]]) / len(zones["Zone 2"])),
        "Zone 3": (sum([p[0] for p in zones["Zone 3"]]) / len(zones["Zone 3"]),
                   sum([p[1] for p in zones["Zone 3"]]) / len(zones["Zone 3"]))
    }

    # Set the dark background style
    plt.style.use('dark_background')

    # Create a figure and axis
    fig, ax = plt.subplots(figsize=(12, 8))

    # Set aspect ratio to 1:1
    ax.set_aspect('equal')
    ax.axis('off')

    # Create a color map and normalize
    cmap = plt.cm.Reds  # Red gradient colormap
    norm = Normalize(vmin=0, vmax=100)  # Normalize to percentage range

    # Plot each zone with gradient color based on percentage
    zone_percentages = [zone1_percentage, zone2_percentage, zone3_percentage]
    for zone_name, vertices, percentage in zip(zones.keys(), zones.values(), zone_percentages):
        color = cmap(norm(percentage))
        polygon = Polygon(vertices, closed=True, color=color)
        ax.add_patch(polygon)

    # Plot each zone with gradient color based on percentage
    zone_percentages = [zone1_percentage, zone2_percentage, zone3_percentage]
    for zone_name, vertices, percentage in zip(zones.keys(), zones.values(), zone_percentages):
        color = cmap(norm(percentage))
        polygon = Polygon(vertices, closed=True, color=color)
        ax.add_patch(polygon)
        # Annotate the zone with percentage
        center = zone_centers[zone_name]
        ax.text(center[0], center[1], f'{percentage:.1f}%', 
                ha='center', va='center', fontsize=14, color='white', weight='bold')

    # Plot the lines
    for i in range(5):  # Iterate over each coordinate
        for name, coords in coordinates.items():
            if name == "line 1":
                if i != 2:  # Exclude third coordinate of line 1
                    ax.plot([coords[i][0], coordinates["line 7"][i][0]], [coords[i][1], coordinates["line 7"][i][1]], color='white')
                else:  # Connect third coordinate of line 1 with line 3
                    ax.plot([coords[i][0], coordinates["line 3"][i][0]], [coords[i][1], coordinates["line 3"][i][1]], color='white')
            elif name == "line 5" and i == 2:  # Connect third coordinate of line 5 with line 7
                ax.plot([coords[i][0], coordinates["line 7"][i][0]], [coords[i][1], coordinates["line 7"][i][1]], color='white')
        # Connect first and last coordinates for every line
        for name, coords in coordinates.items():
            ax.plot([coords[0][0], coords[-1][0]], [coords[0][1], coords[-1][1]], color='white')

    # Display the plot
    plt.title(f'Point Distribution in Set {set_number}', fontsize=16, fontweight='bold', color='white')
    plt.savefig('LeoBagasStatistic/Images/' + f'Point Distribution in Set{set_number}.png', bbox_inches='tight', dpi=300)
    plt.close()

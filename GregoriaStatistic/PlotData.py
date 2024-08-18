import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.offsetbox import OffsetImage, AnnotationBbox

# Load the spreadsheet file
file_path = 'GregoriaStatistic\QuarterFinal_GregoriaMariskaTunjungVsKimGaEun.xlsx'
df = pd.read_excel(file_path)

# Load support file
gregoria_marker = plt.imread('GregoriaStatistic\SourceSupport\Gregoria.png')
kim_marker = plt.imread('GregoriaStatistic\SourceSupport/Kim.png')

# Initialize a dictionary to store data by set
data_by_set = {}

# Extract and store the data for each set
for number_of_set in df['Set'].unique():
    filtered_df = df[df['Set'] == number_of_set]
    data_by_set[number_of_set] = {
        'Total Point': filtered_df['Total Point'].tolist(),
        'Gregoria Point': filtered_df['Gregoria Point'].tolist(),
        'Kim Point': filtered_df['Kim Point'].tolist()
    }

# Define a color palette suitable for dark backgrounds
colors = sns.color_palette("flare", 3)

# Custom marker 
def custom_marker(image, zoom=0.05):
    return OffsetImage(image, zoom=zoom)

# Plotting the data for each set
for set_number in data_by_set:
    total_point = data_by_set[set_number]['Total Point']
    gregoria_point = data_by_set[set_number]['Gregoria Point']
    kim_point = data_by_set[set_number]['Kim Point']

    # Create a figure and axis
    plt.figure(figsize=(12, 6))
    plt.style.use('dark_background')  # Set a dark background style

    # Plot each line
    plt.plot(total_point, gregoria_point, label='Gregoria Points', color=colors[0], marker='o', linewidth=2)
    plt.plot(total_point, kim_point, label='Kim Points', color=colors[1], marker='o', linewidth=2)

    if max(gregoria_point) > max(kim_point):
      max_index = gregoria_point.index(max(gregoria_point))
      ab = AnnotationBbox(custom_marker(gregoria_marker), (total_point[max_index], max(gregoria_point)), frameon=False)
      plt.gca().add_artist(ab)
    elif max(gregoria_point) < max(kim_point):
      max_index = kim_point.index(max(kim_point))
      ab = AnnotationBbox(custom_marker(kim_marker), (total_point[max_index], max(kim_point)), frameon=False)
      plt.gca().add_artist(ab)


    # Annotate the lines
    # for i, txt in enumerate(gregoria_point):
    #     plt.annotate(int(txt), (total_point[i], gregoria_point[i]), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8, color=colors[0])
    # for i, txt in enumerate(kim_point):
    #     plt.annotate(int(txt), (total_point[i], kim_point[i]), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8, color=colors[1])

    # Adding titles and labels
    plt.title(f'Set {int(set_number)} Performance', fontsize=16, fontweight='bold', color='white')
    plt.xlabel('Total Points', fontsize=12, color='white')
    plt.ylabel('Points', fontsize=12, color='white')
    plt.legend(loc='upper left', fontsize=10)

    # Customize the grid
    plt.grid(True, linestyle='--', alpha=0.5)

    # Set grid lines to appear every 3 points
    plt.xticks(np.arange(min(total_point), max(total_point)+1, 2), color='white')
    plt.yticks(np.arange(min(min(gregoria_point), min(kim_point)), max(max(gregoria_point), max(kim_point))+5, 2), color='white')

    # Set the background color
    plt.gca().set_facecolor('black')  # Black background for the plot area
    plt.gcf().patch.set_facecolor('black')  # Black background for the figure

    # Adjust plot layout
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    # Save the plot as an image
    plt.savefig('GregoriaStatistic/Images/' + f'set_{set_number}_performance.png', bbox_inches='tight', dpi=300)
    plt.close()

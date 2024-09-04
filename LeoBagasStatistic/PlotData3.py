import pandas as pd
import matplotlib as plt
import seaborn as sns
import numpy as np
from matplotlib.offsetbox import OffsetImage, AnnotationBbox

file_path = 'LeoBagasStatistic\JapanOpen_LeoBagas.xlsx'
df = pd.read_excel(file_path)

data_by_set={}

for number_of_set in df['Set'].unique:
    df_by_set = df[df['Set'] == number_of_set]
    data_by_set[number_of_set] = {
        'TotalPoints': df_by_set['Number of Last Hit'].tolist(),
        'LeoBagasPoints': df_by_set['Leo/Bagas Points'].tolist(),
        'Executor': df_by_set['By'].tolist(),
        'PointStatus': df_by_set['PointStatus'].tolist(),
        'HitType': df_by_set['Type of Last Hit'].tolist()
    }

colors = sns.color_palette("flare", 5)

for set_number in data_by_set:
    total_points = data_by_set[set_number]['TotalPoints']
    leobagas_points = data_by_set[set_number]['LeoBagasPoints']
    executor = data_by_set[set_number]['executor']
    hit_type = data_by_set[set_number]['HitType']
    point_Status = data_by_set[set_number]["PointStatus"]


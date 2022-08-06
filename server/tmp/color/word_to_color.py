import pandas as pd


def choice_color_direct(input_word, color_patterns, colors):
    impressions = color_patterns['name']
    indexes = [i for i, name in enumerate(impressions) if name == input_word]
    selected_impressions = color_patterns.iloc[indexes]
    selected_names = list(selected_impressions['name'])

    result_colors = get_color_images(
        colors, selected_impressions, indexes, selected_names)

    return result_colors


def get_color_images(colors, impressions, indexes, names):
    result_colors = []
    for i, _ in enumerate(indexes):
        used_colors = list(impressions.iloc[i][impressions.iloc[i] == 1].index)
        result_a_color = []
        for j, color in enumerate(used_colors):
            name_rgb = colors[colors['Hue/Tone'] == color]
            rgb = [int(name_rgb['R']), int(name_rgb['G']), int(name_rgb['B'])]
            result_a_color.append(rgb)
        result_colors.append(result_a_color)

    return result_colors


if __name__ == '__main__':
    colors = pd.read_csv('./data/color.csv')
    color_patterns = pd.read_csv('./data/color_img_scale.csv')
    color_patterns.rename(
        columns={'Unnamed: 0': 'name'},
        inplace=True)
    impression = '人工的な'
    print(choice_color_direct(impression, color_patterns, colors))

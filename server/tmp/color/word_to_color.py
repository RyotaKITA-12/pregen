import gensim
import MeCab
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


def extract_origin(word):
    m = MeCab.Tagger('-Ochasen')
    node = m.parseToNode(word)
    while node:
        origin = node.feature.split(",")[6]
        if origin == "*":
            node = node.next
        else:
            return origin


def create_similar(word):
    word_origin = extract_origin(word)
    impressions = list(dict.fromkeys(list(color_patterns['name'])))
    similar_dict = {}
    for impression in impressions:
        impression_origin = extract_origin(impression)
        try:
            similar_dict[impression] = model.wv.similarity(word_origin,
                                                           impression_origin)
        except KeyError:
            pass
    return sorted(similar_dict.items(), key=lambda x: x[1], reverse=True)


if __name__ == '__main__':
    colors = pd.read_csv('./data/color.csv')
    color_patterns = pd.read_csv('./data/color_img_scale.csv')
    color_patterns.rename(
        columns={'Unnamed: 0': 'name'},
        inplace=True)
    model = gensim.models.Word2Vec.load('./model/word2vec.gensim.model')
    word = input("word :")
    sim_dic = create_similar(word)
    print(sim_dic[0])
    print(choice_color_direct(sim_dic[0][0], color_patterns, colors)[0][0])
    print(choice_color_direct(sim_dic[0][0], color_patterns, colors)[0][1])
    print(choice_color_direct(sim_dic[0][0], color_patterns, colors)[0][2])

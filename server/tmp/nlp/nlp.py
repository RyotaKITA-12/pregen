#必要なライブラリ
import pptx 
import re
import termextract.janome 
import termextract.core
from janome.tokenizer import Tokenizer
import collections
from janome.tokenizer import Tokenizer 

#関数

#デフォルトの文字を省く
def sub(text): 
    text =re.sub('タイトル :', '', text)
    text =re.sub('サブ', '', text)
    text =re.sub('本文1 :', '', text)
    text =re.sub('本文2 :', '', text)
    text =re.sub('本文3 :', '', text)
    text =re.sub('図1 :', '', text)
    text =re.sub('図2 :', '', text)
    text =re.sub('図3 :', '', text)
    
    return text

#pptx(frame)の中身をテキスト抽出(sample.txtとして出力)
#形態素解析用に文ごとの配列を返す
def pptx_text(frame):
    f = open('sample.txt', 'w')
    #frame='sampleFile.pptx'
    txt = []
    prs = pptx.Presentation(fname)

    for i, sld in enumerate(prs.slides, start=1):

        print(f'-- Page {i} --')

        for shp in sld.shapes:

            if shp.has_text_frame:
                print (shp.text)
                shp.text = sub(shp.text)
                txt.append(shp.text)
                f.write(shp.text)

            if shp.has_table:
                tbl = shp.table
                row_count = len(tbl.rows)
                col_count = len(tbl.columns)
                for r in range(0, row_count):                 
                    text=''
                    for c in range(0, col_count):
                        cell = tbl.cell(r,c)
                        paragraphs = cell.text_frame.paragraphs 
                        for paragraph in paragraphs:
                            for run in paragraph.runs:
                                text+=run.text
                            text+=', '
                    print (text)
                    f.write(text)
            print ()
    f.close()
    sentences = [s for s in txt if s != ' ']
    sentences = [s for s in sentences if s != '']
    return sentences
    
#sample.txtの中から重要度が高い順にした配列を返す
def keyword():
    # ファイルパス
    file_path = "sample.txt"
    
    # 日本語テキストの読み込み
    f = open(file_path, "r", encoding="utf-8")
    text = f.read()
    f.close
    
    # 形態素解析器で日本語処理
    t = Tokenizer()
    tokenize_text = t.tokenize(text)
    
    # Frequency生成＝複合語抽出処理（ディクショナリとリストの両方可)
    frequency = termextract.janome.cmp_noun_dict(tokenize_text)
    
    # FrequencyからLRを生成する
    lr = termextract.core.score_lr(
        frequency,
        ignore_words=termextract.mecab.IGNORE_WORDS,
        lr_mode=1, average_rate=1)
    
    # FrequencyとLRを組み合わせFLRの重要度を出す
    term_imp = termextract.core.term_importance(frequency, lr)
    
    # collectionsを使って重要度が高い順に表示
    data_collection = collections.Counter(term_imp)
    for cmp_noun, value in data_collection.most_common():
        #print(termextract.core.modify_agglutinative_lang(cmp_noun), value, sep="\t")
        words.append(termextract.core.modify_agglutinative_lang(cmp_noun))
    return words


#形態素解析をした後、形容詞と名詞を格納した配列を返す
def morpheme(sentences):
    # Tokenizerインスタンスの生成 
    t = Tokenizer()
    #print(text)
    # テキストを引数として、形態素解析の結果、名詞・形容詞(原形)のみを配列で抽出する関数を定義 
    def extract_words(text):
        tokens = t.tokenize(text)
        return [token.base_form for token in tokens 
            if token.part_of_speech.split(',')[0] in['名詞','形容詞']] 

    # それぞれの文章を単語リストに変換
    word_list = [extract_words(s) for s in sentences]
    words=[]
    for word in word_list:
        #print(word)
        words.append(word)
    return words

def main():


if __name__ == '__main__':
    sentences = pptx_text('../pptx/data/test.pptx') #形態素解析に使用する配列
    key_words = keyword() #重要な順に単語を格納した配列
    words = morpheme(sentences)
    
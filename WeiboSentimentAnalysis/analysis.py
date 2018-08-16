from snownlp import SnowNLP
import matplotlib.pyplot as plt
import openpyxl
import jieba
import numpy as np
from wordcloud import WordCloud,STOPWORDS,ImageColorGenerator


class SemanticAnalysis:
    def __init__(self,startTime,keyWord,excelPath):
        self.startTime=startTime
        self.keyWord=keyWord
        self.excelPath=excelPath
        self.message=[]
        self.sentimentslist = []
        self.summary=[]
        self.getMessage()



    def getMessage(self):
        self.wb=openpyxl.load_workbook(self.excelPath)
        title=self.startTime+"-"+self.keyWord
        #title=self.startTime
        try:
            self.sheet=self.wb[title]
        except:
            print('不存在条件表格')
            exit(0)
        for row in self.sheet.rows:
            self.message.append(row[5].value)
        print(len(self.message))

    def snowanalysis(self):
        for li in self.message:
            s = SnowNLP(li)
            self.sentimentslist.append(s.sentiments)
            self.summary.append(s.summary(3))
        print(len(self.sentimentslist))
        plt.hist(self.sentimentslist, bins=np.arange(0, 1, 0.01))
        plt.savefig("./sentiment.png")
        plt.figure('情感分析图')
        plt.show()

    def getWordCloud(self):
        text=''
        for s in self.summary:
            for li in s:
                text+=li
                text+=","
        wordlist = jieba.cut(text, cut_all=False)
        cloud_text=" ".join(wordlist)

        backgroud_Image=plt.imread("./image/wallpaper_01.jpg")
        wc=WordCloud(
            background_color="white",  # 背景颜色
            max_words=1000,  # 显示最大词数
            font_path="./font/mi.ttf",  # 使用字体
            min_font_size=20,
            max_font_size=100,
            width=1000,  # 图幅宽度
            height=860,
            margin=2,#词语边缘距离
            random_state=50,
        )
        image_colors = ImageColorGenerator(backgroud_Image)
        wc.generate(cloud_text)
        #wc.recolor(color_func=image_colors)
        wc.to_file("./wordCloud.png")
        image = plt.imread("./wordCloud.png")
        plt.figure("词云图")
        plt.imshow(image)
        plt.show()

if __name__=="__main__":
    startTime='2018-08-13'
    keyWord='小米'
    excelPath = 'data/weibo.xlsx'
    sa=SemanticAnalysis(startTime,keyWord,excelPath)
    sa.snowanalysis()
    sa.getWordCloud()




import docx
import os
from simplify_docx import simplify
import csv
from nltk.corpus import stopwords
from textblob import TextBlob
from textblob import Word
import ssl
from nltk.stem import PorterStemmer
import nltk
import re
import string
import pandas as pd 

from wordcloud import WordCloud, STOPWORDS 
import matplotlib.pyplot as plt 

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')
stop = stopwords.words('english')



def clean_text(train):
    result = {}
    split_train = train.split()

    #1.2
    train = train.translate(str.maketrans('', '', string.punctuation)) #remove punctuation
    train = train.translate(str.maketrans('', '', string.digits)) #remove digits

    train = " ".join(x for x in train.split() if x not in stop) #remove stop words

    train = train.lower() #lowercase

    st = PorterStemmer()
    train = " ".join(st.stem(x) for x in train.split())

    train = " ".join(Word(x).lemmatize() for x in train.split())
    return train
def cellValue(cell):
    text = ""
    if type(cell['VALUE']) == list and len(cell['VALUE']) > 0:
        cell = cell['VALUE'][0]
        if type(cell['VALUE']) == list and len(cell['VALUE']) > 0:
            cell = cell['VALUE'][0]
            if type(cell['VALUE']) == list and len(cell['VALUE']) > 0:
                text = cell['VALUE'][0]['VALUE']
            elif type(cell['VALUE']) == str:
                text = cell['VALUE']
        elif type(cell['VALUE']) == str:
            text = cell['VALUE']
    elif type(cell['VALUE']) == str:
        text = cell['VALUE']
    return text

def extractContent(json_doc,searchString):
    content_list = []
    for body in json_doc['VALUE']:
        start_appending = False
        stop_appending = False
        for content in body['VALUE']:
            if content['TYPE'] is 'table':
                for row in content['VALUE']:
                    rowContent=[]
                    for cell in row['VALUE']:
                        cellText = cellValue(cell)
                        if searchString in cellText:
                            start_appending = True
                            break
                        if start_appending == True:
                            rowContent.append(cellText)
                    
                    if start_appending == True and len(rowContent) > 0:
                        index = rowContent[0].replace(".","")
                        if not index.isdigit() or int(index) != (len(content_list)) :
                            return content_list
                    if start_appending == True:
                        content_list.append(rowContent)
    return content_list

success_factors = []
lessions_learned = []
shortcomings = []
def write_csv_to_file(filename,list_of_rows):
    with open(filename, mode='w') as f:
        writer = csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for content in list_of_rows:
            if len(content) > 0:
                for index in range(len(content)):
                    if index > 0:
                        content[index]=clean_text(content[index])
                writer.writerow(content)
for (dirpath,dirnames,filenames) in os.walk("./docs"):
    for filename in filenames:
        filepath = os.path.join(dirpath,filename)
        my_doc = docx.Document(filepath)
        json_doc = simplify(my_doc,{"remove-leading-white-space":False})
        factors = extractContent(json_doc,"Factors that promoted this success")
        lessons = extractContent(json_doc,"Lessons learned")
        project_shortcomings = extractContent(json_doc,"Project Shortcomings")
        success_factors.extend(factors)
        lessions_learned.extend(lessons)
        shortcomings.extend(project_shortcomings)

def extract_factors(success_factors):
    factors = []
    factors_list = []
    for f in success_factors:
        if len(f) < 3:
          continue
        factors_list.append(f[2])
    factors_list = list(set(factors_list))
    temp_list = []
    for f in factors_list:
        if len(temp_list) < 3:
            temp_list.append(f)
        else:
            factors.append(temp_list)
            temp_list = [f]
    return factors

def extract_lessions(lessions_learned,index=3):
    lessions = []
    lessions_list = []
    for f in lessions_learned:
        if len(f) < 3:
          continue
        lessions_list.append(f[index])
    lessions_list = list(set(lessions_list))
    temp_list = []
    for f in lessions_list:
        if len(temp_list) < 3:
            temp_list.append(f)
        else:
            lessions.append(temp_list)
            temp_list = [f]
    return lessions

def display_word_document(file_name):
    picture_name = file_name.split(".")[0] + ".png"
    df = pd.read_csv(r"{}".format(file_name), encoding ="latin-1") 
    comment_words = '' 
    stopwords = set(STOPWORDS) 
    
    # iterate through the csv file 
    for val in df: 
        
        # typecaste each val to string 
        val = str(val) 
    
        # split the value 
        tokens = val.split() 
        
        # Converts each token into lowercase 
        for i in range(len(tokens)): 
            tokens[i] = tokens[i].lower() 
        
        comment_words += " ".join(tokens)+" "
    
    wordcloud = WordCloud(width = 800, height = 800, 
                    background_color ='white', 
                    stopwords = stopwords, 
                    min_font_size = 10).generate(comment_words) 
    
    # plot the WordCloud image                        
    plt.figure(figsize = (8, 8), facecolor = None) 
    plt.imshow(wordcloud) 
    plt.axis("off") 
    plt.tight_layout(pad = 0) 
    plt.savefig(picture_name)
factors_file_name="factors.csv"
lessions_file_name = 'lessions.csv'
shortcoming_file_name = 'shortcomings.csv'
write_csv_to_file(shortcoming_file_name,extract_lessions(shortcomings,2))
display_word_document(shortcoming_file_name)
write_csv_to_file(factors_file_name,extract_factors(success_factors))
write_csv_to_file(lessions_file_name,extract_lessions(lessions_learned))
display_word_document(factors_file_name)
display_word_document(lessions_file_name)
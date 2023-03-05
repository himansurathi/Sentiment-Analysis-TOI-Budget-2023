from striprtf.striprtf import rtf_to_text
import os
import re
import nltk
import pandas as pd
from xlwt import Workbook
from nltk.corpus import stopwords
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from datetime import datetime

def get_sentiment(text):
    sid = SentimentIntensityAnalyzer()
    scores = sid.polarity_scores(text)
    if scores['compound'] >= 0.9:
        return 'Strongly Positive'
    elif scores['compound'] >= 0.5:
        return 'Positive'
    elif scores['compound'] > -0.5 and scores['compound'] < 0.5:
        return 'Neutral'
    elif scores['compound'] > -0.9:
        return 'Negative'
    else:
        return 'Strongly Negative'

def log(s):
    if DEBUG:
        print(s)

def preprocess_text(text):
    # Remove double quotes
    text = text.replace('"', '')

    # Remove punctuation
    text = re.sub(r'[^\w\s]', '', text)

    # Convert text to lowercase
    text = text.lower()

    # Tokenize the text
    words = nltk.word_tokenize(text)

    # Remove stop words
    stop_words = set(stopwords.words('english'))
    filtered_words = [word for word in words if word not in stop_words]

    # Join the filtered words back into a string
    preprocessed_text = ' '.join(filtered_words)

    return preprocessed_text

def classify_political_leaning(text):
    # Preprocess the text
    preprocessed_text = preprocess_text(text)

    # Analyze sentiment using VADER
    sia = SentimentIntensityAnalyzer()
    sentiment_scores = sia.polarity_scores(preprocessed_text)

    # Determine the political leaning based on the compound score and polarity
    if sentiment_scores['compound'] >= 0.05 and sentiment_scores['pos'] > sentiment_scores['neg']:
        return 'Right-leaning'
    elif sentiment_scores['compound'] <= -0.05 and sentiment_scores['pos'] < sentiment_scores['neg']:
        return 'Left-leaning'
    else:
        return 'Neutral'

def get_metadata(tokens, flag):
    if flag:
        author = tokens[0].strip().split("@timesgroup.com")[0].replace(".", " ").title() if "@timesgroup.com" in tokens[0] else tokens[0].strip()
        words = tokens[1].split(" ")[0]
        date = tokens[2]
        publication = tokens[3]
        symbol = tokens[4]
        language = tokens[4]
        return author, words, date, publication, symbol, language;
    else:
        author = ''
        words = tokens[0].split(" ")[0]
        date = tokens[1]
        publication = tokens[2]
        symbol = tokens[3]
        language = tokens[4]
        return author, words, date, publication, symbol, language;

def get_headline(tokens, flag):
    if flag:
        category = tokens[0]
        headline = tokens[1]
        return category, headline
    else:
        category = ''
        headline = tokens[0]
        return category, headline

def get_news(tokens):
        return tokens[0].strip()        
    
def read_file(file_path, row_num, sheet):
    message = ''
    with open(file_path, 'r', encoding="utf-8") as f:
        log(file_path)
        try:
            log("MultiVerse \n")
            article_rtf = f.read()
            log("Universe \n")
            log(article_rtf)
            article_text = rtf_to_text(article_rtf, encoding="utf-8").strip()
            log("World \n")
            tokens = article_text.split("\n\n", 2)
            log("India \n")
            if len(tokens) == 3:
                title = tokens[0].split("\n")
                log("Bangalore \n")
                if len(title) == 1 or len(title) == 2:
                    category, headline = get_headline(title, len(title) == 2)
                    metadata = tokens[1].split("\n")
                    log("IIMB \n")
                    if len(metadata) == 6 or len(metadata) == 7:                                       
                        author, words, date, publication, symbol, language = get_metadata(metadata, len(metadata) == 7);
                        break_text = "Document "+ symbol
                        body = tokens[2].split(break_text, 1)
                        log("CCS \n")
                        if len(body) == 2:
                            news = get_news(body)
                            log(news)
                            political_leaning = classify_political_leaning(news)
                            polarity = get_sentiment(news)
                            sheet.write(row_num, 2, author)
                            sheet.write(row_num, 3, words)
                            sheet.write(row_num, 4, date)
                            sheet.write(row_num, 5, publication)
                            sheet.write(row_num, 6, symbol)
                            sheet.write(row_num, 7, language)
                            sheet.write(row_num, 8, headline)
                            sheet.write(row_num, 9, category)
                            sheet.write(row_num, 10, political_leaning)
                            sheet.write(row_num, 11, polarity)
                            message = "Success ----- Parsed file - " + file_path + "\n"
                        else:
                           message = "Error ---- Improper footnote - " + file_path + "\n" 
                    else:
                        message = "Error ---- Improper metadata - " + file_path + "\n"
                else:
                   message = "Error ---- Improper title/headline - " + file_path + "\n"
            else:
                message = "Error ---- Improper file format - " + file_path + "\n"
        except Exception as e:
            message = "Error ----- Could not read the file - " + file_path + "\n" + repr(e)
    return message;

def read_directory(sheet):       
    count = 1
    for file in os.listdir():
        if file.endswith(".rtf"):
            sheet.write(count, 0, count)
            sheet.write(count, 1, file)
            file_path = f"{path}\{file}"
            result = read_file(file_path, count, sheet)
            if result != '':
                print(result)
            count = count + 1

def get_polarity_groupBykey(table, groupby, sheet_name, writer):
    table.sort_values(groupby)
    result = pd.pivot_table(table, values=t_file_name, index=[groupby], columns=[t_polarity], aggfunc='count', fill_value=0)
    result.loc['Total']= result.sum()
    result.to_excel(writer, sheet_name=sheet_name)
    
def transform_table(table):
    table[t_date]= pd.to_datetime(table[t_date])
    table[t_words] = table[t_words].str.replace(',','');
    table[t_words]= pd.to_numeric(table[t_words])    
    print(table[t_words].dtypes)
    table.loc[table[t_polarity] == 'Strongly Positive', t_polarity] = 'Positive'
    table.loc[table[t_polarity] == 'Strongly Negative', t_polarity] = 'Negative'
    table.loc[table[t_words] <= 350, t_article_length] = '<= 350 words'
    table.loc[(table[t_words] <= 700) & (table[t_words] > 350), t_article_length] = '350 - 700 words'
    table.loc[(table[t_words] <= 1050) & (table[t_words] > 700), t_article_length] = '700 - 1050 words'
    table.loc[(table[t_words] <= 1400) & (table[t_words] > 1050), t_article_length] = '1050 - 1400 words'
    table.loc[table[t_words] > 1400, t_article_length] = '>= 1400 words'    

def get_statistics():
    table = pd.read_excel(output_file, sheet_name=0, index_col=0)
    transform_table(table)
    writer = pd.ExcelWriter(statistic_summary_file, engine='xlsxwriter')
    get_polarity_groupBykey(table, t_author, t_author, writer)
    get_polarity_groupBykey(table, t_publication, t_publication, writer)
    get_polarity_groupBykey(table, t_category, t_category, writer)
    get_polarity_groupBykey(table, t_date, t_date, writer)
    get_polarity_groupBykey(table, t_article_length, t_article_length, writer)
    writer.close()


def parse_articles():
    if os.path.exists(output_file):
        os.remove(output_file)
    if os.path.exists(statistic_summary_file):
        os.remove(statistic_summary_file)
    wb = Workbook()
    sheet = wb.add_sheet(sheet_name)
    sheet.write(0, 0, t_id)
    sheet.write(0, 1, t_file_name)
    sheet.write(0, 2, t_author)
    sheet.write(0, 3, t_words)
    sheet.write(0, 4, t_date)
    sheet.write(0, 5, t_publication)
    sheet.write(0, 6, t_symbol)
    sheet.write(0, 7, t_language)
    sheet.write(0, 8, t_headline)
    sheet.write(0, 9, t_category)
    sheet.write(0, 10, t_leaning)
    sheet.write(0, 11, t_polarity) 
    read_directory(sheet)
    wb.save(output_file)
    
            
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('vader_lexicon')

DEBUG = False
output_file= 'CCS.xls'
statistic_summary_file = 'Summary Statistics.xlsx' 
sheet_name = 'NewsDataset'
t_id="Id"
t_file_name = 'File Name'
t_author = 'Author'
t_words = 'Words'
t_article_length = 'Article Length'
t_date = 'Date'
t_publication = 'Publication'
t_symbol = 'Symbol'
t_language = 'Language'
t_headline = 'Headline'
t_category = 'Category'
t_leaning = 'Political Leaning'
t_polarity = 'Polarity'
path = r'D:\IIM Bangalore\CCS\articles'
os.chdir(path);
#parse_articles()
get_statistics()
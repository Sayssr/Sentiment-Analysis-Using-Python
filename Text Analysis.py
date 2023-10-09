#!/usr/bin/env python
# coding: utf-8

pip install beautifulsoup4
pip install requests
pip install --upgrade pandas
import sys
import time
from bs4 import BeautifulSoup
import requests
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import re
import nltk 
nltk.download('punkt')
from nltk.corpus import cmudict
nltk.download('cmudict')
import pandas as pd
#from textblob import TextBlob
import string
from string import punctuation

import os
#Read Excel File
dfile = pd.read_excel('input.xlsx')
#Iterate over each row in data file
for index , row in dfile.iterrows():
    url_id = str(row['URL_ID'])
    url = str(row['URL'])
    
    page = requests.get(url)
    
    
    # Create a BeautifulSoup object to parse the HTML content
    soup = BeautifulSoup(page.text, 'html.parser')
    
    #Find the article title and text 
    title_element =soup.find('h1')
    url_title = title_element.text.strip() if title_element else 'N/A'
    
    #Finding all the paragraph containing text
    paragraphs_element =soup.find_all('p')
  

    #Concatenating all paragraph into text article
    url_text = ' '.join([p.text.strip() for p in paragraphs_element])

    #Creating file for Storing Title and Text
    file_name =f"{url_id}.txt"
    
    folder_path = 'ExtractedData'  # Update with your folder path  
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    file_path = os.path.join(folder_path, file_name) 
    #Saving title and  text data to generated file name
    with open( file_path, 'w', encoding='utf-8') as file:
            file.write(f"Title: {url_title}\n\n")
            file.write(url_text)
    print(f"Article saved for URL_ID: {url_id}")        

print("Extraction complete!")    

# Loading Stop words from files
stop_words_file =['StopWords/StopWords_Auditor.txt','StopWords/StopWords_Currencies.txt',
                  
                  'StopWords/StopWords_DatesandNumbers.txt' ,'StopWords/StopWords_Generic.txt',
                  
                  'StopWords/StopWords_GenericLong.txt','StopWords/StopWords_Geographic.txt',
                  
                  'StopWords/StopWords_Names.txt']
#Loading Stop Words
stop_words=set()
for file in stop_words_file:
    with open(file, 'r') as f:
        words = f.read().splitlines()
        stop_words.update(words)

#Loading positive and negative text file
positive_words_dicitionary = 'MasterDicitionary/positive_words.txt'
negative_words_dicitionary = 'MasterDicitionary/negative_words.txt'

#Loading positive and negative words
positive_words = set()
with open(positive_words_dicitionary, 'r') as f:
    positive_words = set(f.read().splitlines())
    
negative_words = set()
with open(negative_words_dicitionary, 'r') as f:
    negative_words = set(f.read().splitlines())

#Create Variables for empty list    
negative_count = 0
#text_files =['37.txt' , '38.txt' , '39.txt' , '40.txt']
###################Defining Function for empty variables##################################

####Checking for Syllables Count################
def count_syllables(word):
    vowels = "aeiou"
    exceptions = ["es", "ed"]
    
    # Handling exceptions for words ending with "es" or "ed"
    if word[-2:] in exceptions:
        word = word[:-2]
    
    # Counting the number of vowels in the word
    count = 0
    for i in range(len(word)):
        if word[i].lower() in vowels:
            if i == len(word) - 1 or word[i + 1].lower() not in vowels:
                count += 1
    
    return count
###Removing Punctuation Mark##########################t 
import string as st
def text_without_punct(text):
    # Remove punctuation marks from the text
    text = text.translate(str.maketrans("", "", st.punctuation))
    
    # Split the text into words
    words = text.split()
    
    # Count the number of words
    word_count = len(words)
    
    return word_count

###Personal Pronoun Function
def count_personal_pronoun(text):
   
    personal_pronouns = ['i', 'me', 'my', 'mine', 'you', 'your', 'yours']
    words_without_punct = re.findall(r'\b\w+\b', text)
    count_pronoun=0

    for word  in words_without_punct:
        if word.lower() in personal_pronouns:
            count_pronoun =count_pronoun+ 1
    return count_pronoun
    
###############################Read the text file#########################################
# for file in text_files:
for i in range(37,150):
    file = f"ExtractedData/{i}.txt"
    polarity_score=0
    positive_score=0
    negative_score=0
    subjectivity_scores=0
    polarity_score=0
    avg_words_per_sentence =0
    per_complex_words=0 
    fog_index_scores=0 
    complex_word_count=0 
    word_count =0
    personal_pronouns_count =0
    syllable_count =0
    avg_word_length =0 
    positive_count = 0
    total_clean_words=0 
    positive_count = 0
    negative_count = 0
    total_clean_words=0
    with open(file,'r',encoding='utf-8')as f:
        lines=f.readlines() 
        string_words=' '.join(lines)
        expression = word_tokenize(string_words)                            #Converting Lines to Words
        number_of_words = len(expression)                                   #Total Number of Words in file  
       
    #Cleaning Word by removing stop words
        cleaned_words = [word for word in expression if word.lower() not in stop_words]
        cleaned_data =' '.join(cleaned_words)
        clean_words =word_tokenize(cleaned_data)
        total_clean_words =len(clean_words)
    
    #Counting postive and negative words    
    for expression_count in clean_words:
        if expression_count.lower() in positive_words:
             positive_count+=1
        elif expression_count.lower() in negative_words:
            negative_count+=1
    
    #Final Positive Score/Negative Score /Polarity Score/Subjective_score
    positive_score = positive_count
    negative_score = negative_count
    polarity_score = (positive_count - negative_count) / ((positive_count + negative_count) + 0.000001)
    subjectivity_scores = (positive_count  + negative_count)/ (( total_clean_words) + 0.000001)
    print(file, "positive_count", positive_score)
    print("negative_count", negative_score)
    print("polarity", polarity_score)
    print("subjectivity_scores", subjectivity_scores)
    print("total_clean_words", total_clean_words)
   #################Analysis of Readability############################
    #Counting Word after excluding stop word and punctuation 
    word_count=text_without_punct(cleaned_data)
    print("Clean Word After Punctuation:",word_count)
    #Average Number of Words Per Sentence
    sentances =sent_tokenize(string_words)                                           #Converting Lines to Sentences
    total_sentance_count =len(sentances)                                             #Finding Number of Sentences
    avg_words_per_sentence = (word_count/total_sentance_count)                       #Finding Average Sentence Length 
    ##############Counting Number of complex Words####################
    number_of_complex_words=0
    for word in expression:
        syllable_count = count_syllables(word)
        if syllable_count >=2:
            number_of_complex_words += 1
    complex_word_count = number_of_complex_words                                #Complex word Count
    syllable_count = complex_word_count/word_count                              #Counting Syllable in  words
    print("number_of_complex_words:",number_of_complex_words)
    per_complex_words =(number_of_complex_words/word_count)                     #Complex Words Percentage
    fog_index_scores =0.4*(avg_words_per_sentence+per_complex_words) #Calculating Fog Index
    print(file,"\naAverage Sentence Length ",avg_words_per_sentence)           
    print("number_of_comple_words",number_of_complex_words)
    print("perctentage_of_complex_words",per_complex_words)
    print("fog_index_scores",fog_index_scores)
    #####Personal Pronoun Count#########################
    personal_pronouns_count = count_personal_pronoun(string_words)
    print("Personal Pronoun Count:",personal_pronouns_count)
    #Creating a data fram with variables and column headers

    data={ 
          
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': [negative_score],
        'POLARITY SCORE': [polarity_score],
        'SUBJECTIVITY SCORE': [subjectivity_scores],
        'AVG SENTENCE LENGTH': [avg_words_per_sentence],
        'PERCENTAGE OF COMPLEX WORDS': [per_complex_words],
        'FOG INDEX': [fog_index_scores],
        'AVG NUMBER OF WORDS PER SENTENCE': [avg_words_per_sentence],
        'COMPLEX WORD COUNT': [complex_word_count],
        'WORD COUNT': [word_count],
        'SYLLABLE PER WORD': [syllable_count],
        'PERSONAL PRONOUNS': [personal_pronouns_count],
        'AVG WORD LENGTH': [avg_word_length]
    }
    df = pd.DataFrame(data)
    output_df = pd.read_excel("Output.xlsx")  # Read existing data from the file

    # Write data starting from the 3rd column and the second row
    output_df.iloc[1:, 2:] = df

    # Save the updated dataframe to the Excel file 
    output_df.to_excel("Output.xlsx", index=False)
    
    
    
    


#!/usr/bin/env python
# coding: utf-8

# # Import Statements

# In[182]:


import pandas as pd
import numpy as np
import requests
import time
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.tokenize import RegexpTokenizer
import re
from nltk.tokenize import sent_tokenize

pd.set_option('display.max_colwidth', -1)


# In[183]:


nltk.download('punkt')
nltk.download('stopwords')


# # Combining links

# In[235]:


df_source = pd.read_excel(r'K:\College\Project BlackCoffer\cik_list.xlsx')
link = "https://www.sec.gov/Archives/"
df_source['SECFNAME_2'] = link + df_source['SECFNAME']
df_source_links = df_source['SECFNAME_2']
df_source_links.head()


# In[181]:


# Issue while i was trying to read data

# data_source = [*range(0,len(df_source_links),1)]
# for i in data_source:
#     data_source[i] = i


# In[ ]:


# its showing me a HTML Error

# for i in range(0,len(df_source_links),1):
#     response = requests.get(df_source_links[i])
#      data_source[i] = response.text
#      time.sleep(1) 


# In[167]:


# for i in range(0,8,1):
#     response = requests.get(df_source_links[i])
#     data_source[i] = response.text
#     time.sleep(2) 
    
# for i in range(8,15,1):
#     response = requests.get(df_source_links[i])
#     data_source[i] = response.text
#     time.sleep(2) 
    
# for i in range(145,len(df_source_links),1):
#     response = requests.get(df_source_links[i])
#     data_source[i] = response.text
#     time.sleep(2)


# In[236]:


response = requests.get(df_source_links[0])
data = response.text


# In[237]:


data


# In[238]:


# Removing punctuations and numbers
data_alpha = re.sub('[^a-zA-Z]+', ' ', data)


# In[239]:


tokenized_word = list(map(str.upper,word_tokenize(data_alpha))) # for making everyword in uppercase
# print(tokenized_word)


# In[240]:


sentences = sent_tokenize(data)
len(sentences)


# In[242]:


# Creating stopwords master list 

s1 = open("K:\College\Project BlackCoffer\StopWords_Auditor.txt").readlines()
s2 = open("K:\College\Project BlackCoffer\StopWords_Currencies.txt").readlines()
s3 = open("K:\College\Project BlackCoffer\StopWords_DatesandNumbers.txt").readlines()
s4 = open("K:\College\Project BlackCoffer\StopWords_Generic.txt").readlines()
s5 = open("K:\College\Project BlackCoffer\StopWords_GenericLong.txt").readlines()
s6 = open("K:\College\Project BlackCoffer\StopWords_Geographic.txt").readlines()
s7 = open("K:\College\Project BlackCoffer\StopWords_Names.txt").readlines()

stopwords_master = s1+s2+s3+s4+s5+s6+s7

length = len(stopwords_master)

# Loop to remove new line character
for i in range(0,length-1,1):
    stopwords_master[i]=stopwords_master[i].strip()

stopwords_master_upper = list(map(str.upper,stopwords_master))


# In[243]:


# removing stopwords from tokenized words

filtered_list=[]
for w in tokenized_word:
    if w not in stopwords_master_upper:
        filtered_list.append(w)
len(filtered_list)
# print(filtered_list)


# In[244]:


# Reading master Dictionary

df_master_dict = pd.read_excel(r'K:\College\Project BlackCoffer\LoughranMcDonald_MasterDictionary_2020.xlsx')


# In[245]:


# Creating positive Dictionary 

positive_dict = df_master_dict[df_master_dict.Positive >0]['Word'].reset_index(drop=True)

# Creating negative Dictionary 

negative_dict = df_master_dict[df_master_dict.Negative >0]['Word'].reset_index(drop=True)


# In[247]:


# Reading Constraining Dictionary
constraining_dict = pd.read_excel(r'K:\College\Project BlackCoffer\constraining_dictionary.xlsx').reset_index(drop=True)


# In[246]:


# Reading Constraining Dictionary
uncertainty_dict = pd.read_excel(r'K:\College\Project BlackCoffer\uncertainty_dictionary.xlsx').reset_index(drop=True)


# In[248]:


# calculating uncertainty score

uncertainty_score = 0

for w in filtered_list:
    if w in uncertainty_dict['Word'].tolist():
        uncertainty_score = uncertainty_score + 1


# In[249]:


# calculating constraining score

constraining_score = 0

for w in filtered_list:
    if w in constraining_dict['Word'].tolist():
        constraining_score = constraining_score + 1


# In[250]:


# calculating positive score

positive_score = 0

for w in filtered_list:
    if w in positive_dict.tolist():
        positive_score = positive_score + 1


# In[251]:


# calculating negative score

negative_score = 0

for word in filtered_list:
    if word in negative_dict.tolist():
        negative_score = negative_score - 1

negative_score = negative_score*(-1)


# In[252]:


# calculating polarity score

polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
polarity_score


# In[253]:


# calculating subjectivity score

subjectivity_score = (positive_score + negative_score) / len(filtered_list) + 0.000001
subjectivity_score


# In[254]:


# determining type of sentiment through polarity score

if polarity_score < 0.5:
    print("Most Negative")
elif polarity_score >= 0.5 and polarity_score < 0:
    print("Negative")
elif polarity_score == 0:
    print("Neutral")
elif polarity_score >= 0 and polarity_score < 0.5:
    print("Positive")
elif polarity_score >= 0.5:
    print("Most Positive")


# # Analysis of Readability

# In[255]:


# Calculating AVERAGE SENTENCE LENGTH

average_sentence_length = len(filtered_list)/len(sentences)
average_sentence_length


# In[256]:


## function to find count of Syllable in a word
def syllable_count(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count


# In[257]:


# calculating complex_word_count

complex_word_count = 0

for i in range(0,len(filtered_list),1):
    answer = syllable_count(filtered_list[i])
    if answer >= 2:
        complex_word_count = complex_word_count + 1


# In[258]:


# calculating percentage of complex words
percentage_of_complex_words = complex_word_count/(len(filtered_list))


# In[259]:


# calculating fog index
fog_index = 0.4*(average_sentence_length+percentage_of_complex_words)


# # Proportions

# In[260]:


# Calculating word count, constraining_count, uncertainty_count
word_count = len(filtered_list)
constraining_count = len(constraining_dict)
uncertainty_count = len(uncertainty_dict)


# In[261]:


# calculating positive, negative, constraining, uncertainty proportions

positive_word_proportion = positive_score/word_count
negative_word_proportion = negative_score/word_count
constraining_word_proportion = constraining_score/constraining_count
uncertainty_word_proportion = uncertainty_score/uncertainty_count


# In[262]:


# calculating constraining_words_whole_report
constraining_words_whole_report = constraining_score/len(tokenized_word)


# # Adding to Output CSV

# In[263]:


df_output = df_source.copy()


# In[264]:


df_output['positive_score'] = positive_score
df_output['negative_score'] = negative_score
df_output['polarity_score'] = polarity_score
df_output['average_sentence_length'] = average_sentence_length
df_output['percentage_of_complex_words'] = percentage_of_complex_words
df_output['fog_index'] = fog_index
df_output['complex_word_count'] = complex_word_count
df_output['word_count'] = word_count
df_output['uncertainty_score'] = uncertainty_score
df_output['constraining_score'] = constraining_score
df_output['positive_word_proportion'] = positive_word_proportion
df_output['negative_word_proportion'] = negative_word_proportion
df_output['uncertainty_word_proportion'] = uncertainty_word_proportion
df_output['constraining_word_proportion'] = constraining_word_proportion
df_output['constraining_words_whole_report'] = constraining_words_whole_report


# In[265]:


df_output.head()


# In[266]:


file_name = "K:\College\Project BlackCoffer\output.xlsx"
df_output.to_excel(file_name)


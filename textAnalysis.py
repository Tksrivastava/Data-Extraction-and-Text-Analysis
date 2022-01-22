import pandas as pd
from nltk.tokenize import word_tokenize
import re

# Importing 'Output Data Structure.xlsx' file
output_df = pd.read_excel('C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/AssessmentDetails/'
                          'Output Data Structure.xlsx', sheet_name='Sheet1')
URL_ID_List = list(output_df['URL_ID'])
output_df.set_index('URL_ID', inplace=True)

# Importing the 'LoughranMcDonald_MasterDictionary_2020'
masterDict_df = pd.read_csv('C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/'
                            'LoughranMcDonald_MasterDictionary_2020.csv')

# Creating positive score and negative score dictionaries for calculating positive and negative scores
positive_dict = dict(zip(masterDict_df['Word'], masterDict_df['Positive']))
negative_dict = dict(zip(masterDict_df['Word'], masterDict_df['Negative']))

# Creating stop word list
file = open('C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/StopWords_GenericLong.txt')
stopWordList = []
for line in file:
    stripped_line = line.strip()
    stopWordList.append(stripped_line)

# Calculating all parameters of text Analysis
for index in range(0, len(URL_ID_List)):
    fileIndex = index
    print('Index number ----', fileIndex+1)
    filename = "%s.txt" % (fileIndex + 1)

    # Creating tokens of extracted text files
    address = 'C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/textData/%s' % filename
    textFile = open(address, 'r', encoding="utf-8").read()
    tokenize_text = word_tokenize(textFile)

    # Cleaning stop words from extracted text
    Text = []
    for i in range(0, len(tokenize_text)):
        if tokenize_text[i].lower() in stopWordList:
            pass
        else:
            Text.append(tokenize_text[i])

    # Calculating number of lines
    lines = 0
    for i in range(0, len(Text)):
        if Text[i] == '.':
            lines += 1

    # Removing punctuation from the 'Text'
    punctuation = []
    for i in range(0, len(Text)):
        if Text[i].isalnum() != True and Text[i] != '.':
            punctuation.append(Text[i])
    for j in range(0, len(punctuation)):
        Text.remove(punctuation[j])

    # Calculating number of words present in the 'Text'
    wordCount = len(Text)

    # Calculating positive and negative scores respectively
    positive_score = 0
    negative_score = 0
    for i in range(0, len(Text)):
        for key in positive_dict:
            if Text[i].upper() == key and positive_dict[key] == 2009:
                positive_score += 1

    for i in range(0, len(Text)):
        for key in negative_dict:
            if Text[i].upper() == key and negative_dict[key] == 2009:
                negative_score += 1
    negative_score = -negative_score

    # Calculating polarity score
    if positive_score + negative_score == 0:
        polarity_score = 0
    else:
        polarity_score = ((positive_score - negative_score) / (positive_score + negative_score)) + 0.000001

    # Calculating subjective score
    subjective_score = ((positive_score + negative_score) / len(Text)) + 0.000001

    # calculating average sentence length, average number of words per sentence
    averageSentenceLen = wordCount / lines

    # calculating number of complex words present in the extracted word
    vowel = ['a', 'e', 'i', 'o', 'u']
    complexCount = 0
    syllable = 0
    for i in range(0, len(Text)):
        wordList = list(Text[i])
        if wordList[-2:] == ['e', 'd'] or wordList[-2:] == ['e', 's']:
            for j in range(0, len(wordList)):
                for k in range(0, len(vowel)):
                    if wordList[j] == vowel[k]:
                        syllable += 1
            syllable -= 1
        else:
            for j in range(0, len(wordList)):
                for k in range(0, len(vowel)):
                    if wordList[j] == vowel[k]:
                        syllable += 1
        if syllable > 1:
            complexCount += 1
        syllable = 0

    # Calculating percentage of complex words present in the extracted text
    percentComplex = complexCount / wordCount

    # Calculating the fog index
    fogIndex = (averageSentenceLen + percentComplex) * 0.4

    # Calculating syllable per word
    syllable = 0
    for i in range(0, len(Text)):
        wordList = list(Text[i])
        if wordList[-2:] == ['e', 'd'] or wordList[-2:] == ['e', 's']:
            for j in range(0, len(wordList)):
                for k in range(0, len(vowel)):
                    if wordList[j] == vowel[k]:
                        syllable += 1
            syllable -= 1
        else:
            for j in range(0, len(wordList)):
                for k in range(0, len(vowel)):
                    if wordList[j] == vowel[k]:
                        syllable += 1
    syllablePerWord = syllable/wordCount

    # Calculating personal pronouns
    pattern1 = re.compile(r'\sI\s')
    pattern2 = re.compile(r'\swe\s|\sWe\s')
    pattern3 = re.compile(r'\smy\s|\sMy\s')
    pattern4 = re.compile(r'\sours\b|\sOurs\b')
    pattern5 = re.compile(r'\sus\s|\sUs\s')
    match1 = pattern1.findall(textFile)
    match2 = pattern2.findall(textFile)
    match3 = pattern3.findall(textFile)
    match4 = pattern4.findall(textFile)
    match5 = pattern5.findall(textFile)
    personalPronouns = len(match1) + len(match2) + len(match3) + len(match4) + len(match5)

    # Calculating average word length
    charecterCount = 0
    for i in range(0, len(Text)):
        wordList = list(Text[i])
        for j in range(0, len(wordList)):
            charecterCount += 1
    charecterWord = charecterCount/wordCount

    # Saving values to 'Output Data structure.xlsx'
    output_df['POSITIVE SCORE'].iloc[fileIndex] = positive_score
    output_df['NEGATIVE SCORE'].iloc[fileIndex] = negative_score
    output_df['POLARITY SCORE'].iloc[fileIndex] = polarity_score
    output_df['SUBJECTIVITY SCORE'].iloc[fileIndex] = subjective_score
    output_df['AVG SENTENCE LENGTH'].iloc[fileIndex] = averageSentenceLen
    output_df['PERCENTAGE OF COMPLEX WORDS'].iloc[fileIndex] = percentComplex
    output_df['FOG INDEX'].iloc[fileIndex] = fogIndex
    output_df['AVG NUMBER OF WORDS PER SENTENCE'].iloc[fileIndex] = averageSentenceLen
    output_df['COMPLEX WORD COUNT'].iloc[fileIndex] = complexCount
    output_df['WORD COUNT'].iloc[fileIndex] = wordCount
    output_df['SYLLABLE PER WORD'].iloc[fileIndex] = syllablePerWord
    output_df['PERSONAL PRONOUNS'].iloc[fileIndex] = personalPronouns
    output_df['AVG WORD LENGTH'].iloc[fileIndex] = charecterWord
output_df.to_csv('Output Data Structure Filled.csv')

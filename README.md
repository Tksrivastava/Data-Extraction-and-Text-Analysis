# Data-Extraction-and-Text-Analysis
The main objective of this project is to show that I can scrape websites, manipulate, clean, and analyze textual data. For this project I have used 170 URLs [which you can find on Input.xlsx fille] scraping and extracting textual data. After tokanizing and removing all stop words from the text data. I have calculated 13 parameters for each textual data. These parameters are:
1. Positive score
2. Negative score
3. Polarity score
4. Subjectivity score
5. Avg sentence length
6. Percentage of complex words
7. Fog index
8. Avg number of words per sentence
9. Complex word count
10. Word count
11. Syllable per word
12. Personal pronouns
13. Avg word length

For removing the stop words from the textual data I have used 'StopWords GenericLong.txt' and used 'LoughranMcDonald MasterDictionary 2020.csv' for calculating all the parameters.

I had wrote two distinct scripts for this project:
1. dataExtraction.py - This script will extracts text data from the given URLs.
2. textAnalysis.py - This script will be used to analyse the extracted text data from from URLs.

When you execute dataExtraction.py after updating the Input.xlsx address and the place where the extracted text file will be saved, it will automatically process and save 170 text files. After modifying the locations of Input.xlsx, LoughranMcDonald MasterDictionary 2020.csv, StopWords GenericLong.txt, and extracted text, execute texAnalysis.py. It will calculate all of the essential parameters for all 170 text files automatically and save it to 'Output Data Structure Filled.csv'.

• This project is completed with the help of pandas, requests, bs4, nltk, and re packages installed on Python 3.8 and PyCharm IDE. 

from newspaper import Article
import pandas as pd
import xlsxwriter
from nltk.tokenize import word_tokenize
from nltk.corpus import words
from nltk.corpus import cmudict
import re

# Read the Excel file into a DataFrame
df = pd.read_excel('Input.xlsx')

#Store data in the form of dictionary where url_id is key and url is value
dt = df.set_index('URL_ID')['URL'].to_dict()




# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook("Output_Data_Structure.xlsx")
worksheet = workbook.add_worksheet()

cv = [] #for row data
cn = ['URL_ID','URL','POSITIVE SCORE','NEGATIVE_SCORE','POLARITY SCORE','SUBJECTIVITY SCORE','AVG SENTENCE LENGTH','PERCENTAGE OF COMPLEX WORDS','FOG INDEX',
      'AVG NUMBER OF WORDS PER SENTENCE','COMPLEX WORD COUNT','SYLLABLE PER WORD','PERSONAL PRONOUNS','AVG WORD LENGTH']
cv.append(cn)

k = 0
for i in cn:                  #insert row heading
  worksheet.write(0,k,i)
  k += 1



#Extract data from the url including heading and text
def extract_text_and_heading_from_url(url):
    article = Article(url)
    article.download()
    article.parse()

    heading = article.title
    text = article.text

    return heading, text





#Calculate Positive and Negative Score
def score(text):
  Positive_Score = 0
  Negative_Score = 0


  # Open the positive word file
  with open('positive-words.txt', 'r') as file1:
    file_contents1 = ' '.join(line.strip() for line in file1)
    file_contents1  = file_contents1.split()

  # Open the negative word file
  # Initialize an empty list to store words
  word_list = []

  # Open the text file with specific encoding
  with open('negative-words.txt', 'r', encoding='latin-1') as file:
    # Read each line
    for line in file:
        # Use regular expression to split the line into words
        words = re.findall(r'\b(?:\w+-\w+|\w+)\b', line)
        # Extend the word_list with the words from the current line
        word_list.extend(words)

    for i in text:
      if i in file_contents1:
        Positive_Score = Positive_Score+1

      else:
        if i in word_list:
          Negative_Score = Negative_Score+1


  return Positive_Score,Negative_Score



#Calculate average sentence length and average word per sentence
from nltk.tokenize import sent_tokenize, word_tokenize

def average_sentence_length(paragraph):
    # Split the paragraph into sentences
    sentences = sent_tokenize(paragraph)

    # Calculate the total number of characters in all sentences
    total_chars = sum(len(sentence) for sentence in sentences)

    # Calculate the average sentence length
    if sentences:  # Check if there are any sentences to avoid division by zero
        avg_sentence_length = total_chars / len(sentences)
    else:
        avg_sentence_length = 0

    return avg_sentence_length

def average_words_per_sentence(paragraph):
    # Split the paragraph into sentences
    sentences = sent_tokenize(paragraph)

    # Calculate the total number of words in all sentences
    total_words = sum(len(word_tokenize(sentence)) for sentence in sentences)

    # Calculate the average words per sentence
    if sentences:  # Check if there are any sentences to avoid division by zero
        avg_words_per_sentence = total_words / len(sentences)
    else:
        avg_words_per_sentence = 0

    return avg_words_per_sentence




#calculate total word count
def calculate_total_word_count(text):
    words = word_tokenize(text)
    total_word_count = len(words)
    return total_word_count



# Download necessary NLTK resources
nltk.download('punkt')  # for tokenization
nltk.download('cmudict')  # CMU Pronouncing Dictionary

# Load CMU Pronouncing Dictionary
d = cmudict.dict()

def count_complex_words(paragraph):
    # Split the paragraph into words
    words = word_tokenize(paragraph)

    # Calculate the number of complex words
    complex_words = [word for word in words if nsyl(word) >= 3]  # A word is considered complex if it has 3 or more syllables
    complex_word_count = len(complex_words)

    return complex_word_count, complex_words

def calculate_percentage_complex_words(complex_word_count, total_word_count):
    # Calculate the percentage of complex words
    if total_word_count != 0:
        percentage_complex_words = (complex_word_count / total_word_count) * 100
    else:
        percentage_complex_words = 0

    return percentage_complex_words

def calculate_syllables_per_word(paragraph):
    # Split the paragraph into words
    words = word_tokenize(paragraph)

    # Calculate the total number of syllables
    total_syllables = sum(nsyl(word) for word in words)

    # Calculate the average syllables per word
    if words:  # Check if there are any words to avoid division by zero
        syllables_per_word = total_syllables / len(words)
    else:
        syllables_per_word = 0

    return syllables_per_word

def calculate_fog_index(average_sentence_length, percentage_complex_words):
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    return fog_index

def nsyl(word):
    if word.lower() in d:
        return [len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]][0]  # Return syllable count of first pronunciation
    else:
        return 0  # If the word is not found, return 0

def analyze_paragraph(paragraph):
    # Count total words in the paragraph
    words = word_tokenize(paragraph)
    total_word_count = len(words)

    # Calculate complex word count and percentage
    complex_word_count, _ = count_complex_words(paragraph)
    percentage_complex_words = calculate_percentage_complex_words(complex_word_count, total_word_count)

    # Calculate syllables per word
    syllables_per_word = calculate_syllables_per_word(paragraph)

    # Calculate average sentence length
    sentences = sent_tokenize(paragraph)
    average_sentence_length = sum(len(word_tokenize(sentence)) for sentence in sentences) / len(sentences)

    # Calculate FOG index
    fog_index = calculate_fog_index(average_sentence_length, percentage_complex_words)

    return complex_word_count, percentage_complex_words, syllables_per_word, fog_index



#Calculate count of personal pronoun
def count_personal_pronouns(paragraph):
    # Define a list of personal pronouns
    personal_pronouns = ['I', 'me', 'my', 'mine', 'myself', 'you', 'your', 'yours', 'yourself', 'he', 'him', 'his', 'himself', 'she', 'her', 'hers', 'herself', 'it', 'its', 'itself', 'we', 'us', 'our', 'ours', 'ourselves', 'you', 'your', 'yours', 'yourselves', 'they', 'them', 'their', 'theirs', 'themselves']

    # Count occurrences of personal pronouns
    total_count = 0
    for pronoun in personal_pronouns:
        count = len(re.findall(r'\b' + pronoun + r'\b', paragraph, re.IGNORECASE))
        total_count += count

    return total_count





#Calculae average word length
def average_word_length(paragraph):
    # Tokenize the paragraph into words
    words = word_tokenize(paragraph)

    # Calculate the total number of characters in all words
    total_chars = sum(len(word) for word in words)

    # Calculate the average word length
    if words:  # Check if there are any words to avoid division by zero
        avg_word_length = total_chars / len(words)
    else:
        avg_word_length = 0

    return avg_word_length



row_index = 1    #because 1st row fill as heading above
for url_id in dt:
  try:
    column_index = 0
    row_data = []     #collect data for each row(each file data)
    heading,text = extract_text_and_heading_from_url(dt[url_id])
    f_content = heading+text
    row_data.append(url_id)
    row_data.append(dt[url_id])



    g,h = score(f_content.split())
    #Calculate positive score
    Positive_Score = g

    #Calculate negative score
    Negative_Score = h

    #Calculate Polarity Score
    Polarity_Score = (Positive_Score - Negative_Score)/ ((Positive_Score + Negative_Score) + 0.000001)
    row_data.append(Positive_Score)
    row_data.append(Negative_Score)
    row_data.append(Polarity_Score)


    #calculate SUBJECTIVITY_SCORE
    total_word_count = calculate_total_word_count(f_content)
    SUBJECTIVITY_SCORE = (Positive_Score + Negative_Score)/((total_word_count) + 0.000001)

    row_data.append(SUBJECTIVITY_SCORE)


    #Calculate and append the average sentence length
    avg_sentence_len = average_sentence_length(f_content)
    row_data.append(avg_sentence_len)

    #Calculate complex_word_count, percentage_complex_words, syllables_per_word and fog_index
    complex_word_count, percentage_complex_words, syllables_per_word, fog_index = analyze_paragraph(f_content)

    #Append percentage of complex sentence
    row_data.append(percentage_complex_words)

    #Append percentage of fog_index
    row_data.append(fog_index)


    #Calculate average words per sentence
    avg_words_per_sent = average_words_per_sentence(f_content)
    row_data.append(avg_words_per_sent)


    #Append percentage of complex_word_count
    row_data.append(complex_word_count)

    #Append percentage of  syllables_per_word
    row_data.append(syllables_per_word)



    #Calculate count of personal pronoun
    personal_pro = count_personal_pronouns(f_content)
    row_data.append(personal_pro)
    

    #Calculate average word length
    avg_length  = average_word_length(f_content)
    row_data.append(avg_length)
    



    #Entry every row wise
    for row_insert in range(len(row_data)):
      worksheet.write(row_index,column_index,row_data[row_insert])
      column_index += 1

    row_index += 1

  except:
    row_data = []
    row_data.append(url_id)
    row_data.append(dt[url_id])


    #Entry only url _id and url because data not found from the link
    for row_insert in range(len(row_data)):
      worksheet.write(row_index,column_index,row_data[row_insert])
      column_index += 1

    row_index += 1 




#save and close the file
workbook.close()
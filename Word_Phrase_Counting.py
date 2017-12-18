
import sys
import codecs
import nltk
import csv
from nltk.corpus import stopwords


# NLTK's default German stopwords
default_stopwords = set(nltk.corpus.stopwords.words('english'))


stopwords_file = 'C:\\Users\\ja052464\AppData\\Roaming\\nltk_data\\corpora\\stopwords\\english'

all_stopwords = default_stopwords

# input_file = sys.argv[1]


fp = codecs.open("C:\\Users\\ja052464\\Downloads\\All_Validated_Counts.txt", 'r', encoding='utf-8', errors='ignore')


words = nltk.word_tokenize(fp.read())

scrubbed_words = []

# Remove single-character tokens (mostly punctuation)
# words = [word for word in words if len(word) > 1]

for word in words:
    if len(word) > 1 and not word.isnumeric() and word not in all_stopwords:
        scrubbed_words.append(word.lower())



# Calculates count of individual words
fdist = nltk.FreqDist(scrubbed_words)

# creates word combination pairs
bgs = nltk.bigrams(scrubbed_words)

bglist = nltk.FreqDist(bgs)


with open("final_counts.csv", "w") as fp:
    writer = csv.writer(fp, quoting=csv.QUOTE_ALL, lineterminator='\n')
    writer.writerows(fdist.items())

with open("final_combinations.csv", "w") as bp:
    writer = csv.writer(bp, quoting=csv.QUOTE_MINIMAL,lineterminator='\n')
    writer.writerows(bglist.items())

print("done!:)")
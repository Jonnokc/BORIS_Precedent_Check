import Levenshtein
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import timeit
import sys
import os


def sys_check(s1, s2):
    if s1 in s2:
        return True
    else:
        return False


def keyword_lookup(s1):
    for word in s1:
        for keyword in keywords_d['Keywords']:
            if word.lower() in keyword.lower():
                return True
    return False


def code_sys_check(s1, s2):
    for code in s1:
        if code.lower() in s2.lower():
            return True
    return False


# headers representing data needed
exc_headers = ['RAW_CODE_DISPLAY', 'RAW_CODE_SYSTEM_NAME', 'PRECEDENT_ANALYSIS']
acc_headers = ['Raw Code Display', 'Map Count', 'Precedent Raw Code System ID(s)']

# start time used to determine program completion
start_time = timeit.default_timer()

# TODO prompt user for path to files
# user_input = raw_input("Enter the path of your file: ")
#
# assert os.path.exists(user_input), "I did not find the file at, "+str(user_input)
# f = open(user_input,'r+')
# print("Hooray we found your file!")


# file paths
excel_workbook_location = (
    "C:/Users/ja052464/Downloads/Testing_Python_Similarity4.csv")
access_workbook_location = (
    "C:/Users/ja052464/Downloads/All_Validated_With_Counts.csv")
keywords_workbook_location = (
    "Y:\\Data Intelligence\\Code_Database\\Proprietary_Code_Display_Keywords.csv")

# save path
path = "C:\\Users\\ja052464\\Downloads\\"

# read in the files
exc_d = pd.read_csv(excel_workbook_location).astype(str)
acc_d = pd.read_csv(access_workbook_location, encoding='iso-8859-1').astype(str)
keywords_d = pd.read_csv(keywords_workbook_location).astype(str)


# TODO loop through pandas headers for excel and identify if a column is missing. If so, ask user for actual column name.
# loops through the headers of the unmapped codes CSV file to check for any header changes.
# If a header is different, prompt user for the header name and use that.
for header in range(len(exc_headers)):
    if exc_headers[header] not in list(exc_d.columns.values):
        print('Could not find a column with the header %s. I will need you to enter the header representing this column within the data file.\n Your options are.. \n ' % exc_headers[header])
        for i in exc_d.columns.values:
            print(i)

        while True:
            try:
                column_name = input("What is the name of the header representing %s? " % exc_headers[header])
                if column_name not in exc_d.columns.values:
                    print("That column name is not within the file. Please try again. Capitalization counts.")
                else:
                    print("Perfect. I will use that column header instead")
                    exc_headers[header] = column_name
                    break
            except ValueError:
                print("Not a valid input. Try again.\n")
                continue

# row count used to determine completion time
total_access_rows = acc_d['Raw Code Display'].count()
total_excel_rows = exc_d['RAW_CODE_DISPLAY'].count()

# calculate expected completion time
expected_time = (total_excel_rows * 5) / 60

# tell user completion time.
print("Oh hai. So this might take a bit. I am checking %s unmapped codes against %s validated codes...."
      "That is a lot of math. I expect to be done in roughly %.2f minutes"
      % (total_excel_rows, total_access_rows, expected_time))

# list(my_dataframe.columns.values)'

# To add a new header
# exc_d['new header name'] = " "

# Loops throw the currently unmapped raw code displays and checks it against the previously mapped displays.
for row in exc_d[exc_headers[0]].items():
    i = row[0]
    unmapped = row[1]
    un_code_sys = exc_d.loc[i][exc_headers[1]]
    code_time = timeit.default_timer()
    unmapped_code_split = unmapped.split()
    best_score = 0
    best_display = ""

    for j in acc_d['Raw Code Display']:

        # Very quick way of looping through to find exact and very close matches.
        score = Levenshtein.ratio(unmapped, j)
        # score = fuzz.ratio(unmapped, j)

        if score is 1:
            best_score = score
            best_display = j
            break
        elif score > best_score:
            best_score = score
            best_display = j

    # TODO - add in loop for key words to see if code contains any of the high priority key words
    keyword_check = keyword_lookup(unmapped_code_split)


    # TODO - unmapped code score is over 50 and contains a keyword, then use the fuzzy score

    if best_score > .5 and keyword_check is True:
        results = process.extractOne(
            unmapped, acc_d['Raw Code Display'], scorer=fuzz.token_sort_ratio, score_cutoff=50)

        best_display = results[0]
        best_score = (results[1] / 100)

        # checks if the best match has a code system within the string.
        val_code_sys = acc_d.loc[i, acc_headers[2]]
        code_sys_check = code_sys_check(un_code_sys, val_code_sys)



    # TODO Set loop to evaluate code. If match from Levenshtein is less than 100%, but is greater than XXX then run the fuzzy check to determine if there is a closer match.
    # Slower way to evaluate closer matches to get better result
    # results = process.extractOne(
    #     unmapped, acc_d['Raw Code Display'], scorer=fuzz.token_sort_ratio, score_cutoff=50)
    # best_display = results[0]
    # best_score = (results[1] / 100)
    # best_concept_alias = acc_d.loc([results[2],'Concept_Alias'])

    Code_elapsed = timeit.default_timer() - code_time

    exc_d.set_value(i, ['Python_Match'], best_display)
    exc_d.set_value(i, ['Python_Similarity'], best_score)
    exc_d.set_value(i, ['PRECEDENT_ANALYSIS'], keyword_check)
    # exc_d.set_value(i, ['Python_Concept_Alias'], best_concept_alias)

    print("Best match for %s is %s. Score was %.2f. It took %.2f sec to complete" %
          (unmapped, best_display, best_score, Code_elapsed))


elapsed = (timeit.default_timer() - start_time) / 60

print("I am done doing math. Hold tight while I update the file!")

exc_d.to_csv(path + "testing_export.csv")

print("All done! Finished in %.2f minutes" % elapsed)

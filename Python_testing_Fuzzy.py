import Levenshtein
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import timeit


# checks unmapped code display words against keywords table
def keyword_lookup(s1):
    for word in s1:
        for keyword in keywords_d['Keywords']:
            if word.lower() in keyword.lower():
                return True
    return False


# checks unmapped raw code system against the code systems within the best match
def code_sys_check(s1, s2):
    for code in s1:
        if code.lower() in s2.lower():
            return True
    return False


# uses token sort fuzzywuzzy lookup
def fuzzy_matching(s1, array):
    return process.extractOne(s1, array, scorer=fuzz.token_sort_ratio)


# headers representing data needed
exc_headers = [
    'RAW_CODE_DISPLAY', 'RAW_CODE_SYSTEM_NAME', 'PRECEDENT_ANALYSIS',
    'PRECEDENT_MATCH', 'PRECEDENT_SIMILARITY'
]
acc_headers = [
    'Raw Code Display', 'Map Count', 'Precedent Raw Code System ID(s)'
]

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
    "Y:\\Data Intelligence\\Code_Database\\Proprietary_Code_Display_Keywords.csv"
)

# save path
path = "C:\\Users\\ja052464\\Downloads\\"

# read in the needed files
exc_d = pd.read_csv(
    excel_workbook_location,
    dtype={exc_headers[0]: str,
           exc_headers[1]: str,
           exc_headers[2]: str}).astype(str)
acc_d = pd.read_csv(
    access_workbook_location,
    encoding='iso-8859-1',
    dtype={
        acc_headers[0]: str,
        acc_headers[1]: str,
        acc_headers[2]: str,
    }).astype(str)
keywords_d = pd.read_csv(keywords_workbook_location, dtype=str).astype(str)

# loops through the headers of the unmapped codes CSV file to check for any header changes.
# If a header is different, prompt user for the header name and use that.
for header in range(len(exc_headers)):
    if exc_headers[header] not in list(exc_d.columns.values):
        print(
            'Could not find a column with the header %s. I will need you to enter the header representing this column within the data file.\n Your options are.. \n '
            % exc_headers[header])
        for i in exc_d.columns.values:
            print(i)

        while True:
            try:
                column_name = input(
                    "What is the name of the header representing %s? " %
                    exc_headers[header])
                if column_name not in exc_d.columns.values:
                    print(
                        "That column name is not within the file. Please try again. Capitalization counts."
                    )
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
print(
    "Oh hai. So this might take a bit. I am checking %s unmapped codes against %s validated codes...."
    "That is a lot of math. I expect to be done in roughly %.2f minutes" %
    (total_excel_rows, total_access_rows, expected_time))

# list(my_dataframe.columns.values)'

# To add a new header
# exc_d['new header name'] = " "

# Loops throw the currently unmapped raw code displays and checks it against the previously mapped displays.
for row in exc_d.itertuples():
    i = row[0]
    unmapped = row[1]
    un_code_sys = row[2]
    unmapped_code_split = row[1].split()
    best_score = 0
    best_display = ""
    prec_analysis = ""
    code_time = timeit.default_timer()

    for j in acc_d.itertuples():

        # Very quick way of looping and excluding low probability matches
        score = Levenshtein.ratio(unmapped, j[1])

        if score is 1:
            best_score = score
            best_display = j[1]
            break
        elif score > best_score:
            best_score = score
            best_display = j[1]

    # Checks each word against the keyword table for any hits.
    keyword_check = keyword_lookup(unmapped_code_split)

    if best_score == 1 and keyword_check:
        pass

    elif .8 <= best_score < 1 or keyword_check:

        # calculate the best match using fuzzy
        results = fuzzy_matching(unmapped, acc_d['Raw Code Display'])

        # store results
        best_display = results[0]
        best_score = (results[1] / 100)

    # Checks if code system is within the best match.
    code_sys_check_result = code_sys_check(un_code_sys,
                                           acc_d.loc[i, acc_headers[2]])

    # determines analysis result
    if best_score == 1 and code_sys_check_result:
        prec_analysis = "100% Match"
    elif best_score >= .88 and code_sys_check_result and keyword_check:
        prec_analysis = "Strong Match & Display Contains Keyword"
    elif .7 >= best_score < .85 and keyword_check:
        prec_analysis = "Possible Match. Contains Keyword"
    elif best_score < .7 and keyword_check is False:
        prec_analysis = "Weak Match. No Keyword"
    elif best_score <= .7 and keyword_check is False:
        prec_analysis = "Low Probability. No Keyword"

    Code_elapsed = timeit.default_timer() - code_time

    # write to dataFrame
    exc_d.set_value(i, ['Python_Match'], best_display)
    exc_d.set_value(i, ['Python_Similarity'], best_score)
    exc_d.set_value(i, ['PRECEDENT_ANALYSIS'], prec_analysis)

elapsed = (timeit.default_timer() - start_time) / 60

print("I am done doing math. Hold tight while I update the file!")

exc_d.to_csv(path + "testing_export.csv")

print("All done! Finished in %.2f minutes" % elapsed)

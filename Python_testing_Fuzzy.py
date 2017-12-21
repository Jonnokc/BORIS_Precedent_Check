import Levenshtein
import pandas as pd
import timeit
import new_util
from fuzzywuzzy import process
from fuzzywuzzy import fuzz


# checks unmapped code display words against keywords table
def keyword_lookup(s1):
    for word in s1:
        for keyword in key_d['Keywords']:
            if word.lower() in keyword.lower():
                return True
    return False


# checks unmapped raw code system against the code systems within the best match
def code_sys_check(s1, s2):
    for code in s1:
        if code.lower() in s2.lower():
            return True
    return False


# yes no prompt
def yes_or_no(question):
    while "the answer is invalid":
        reply = str(input(question + ' (y/n): ')).lower().strip()
        if reply[0] == 'y':
            return True
        if reply[0] == 'n':
            return False


def get_best_match(s1, array):
    highest_score = 0
    best_row = ""
    for j in array:
        score = Levenshtein.ratio(s1, j[1])
        if score is 1:
            return j, score
        elif score > highest_score:
            best_row = j
            highest_score = score
    return best_row, highest_score


def medication_check(s1, array):
    if s1 in array:
        return True
    return False


def process_and_sort(s, force_ascii, full_process=True):
    """Return a cleaned string with token sorted."""
    # pull tokens
    ts = new_util.full_process(s, force_ascii=force_ascii) if full_process else s
    tokens = ts.split()

    # sort tokens and join
    sorted_string = u" ".join(sorted(tokens))
    return sorted_string.strip()


@new_util.check_for_none
def token_clean(s1, force_ascii=True, full_process=True):
    sorted1 = process_and_sort(s1, force_ascii, full_process=full_process)
    return sorted1


def column_locations(preset_headers, file_name):
    for header in range(len(preset_headers)):
        if preset_headers[header] in list(file_name.columns.values):
            preset_headers[header] = file_name.columns.get_loc(preset_headers[header])

        else:
            print(
                'Could not find a column with the header %s. I will need you to enter the header representing this column within the data file.\n Your options are.. \n '
                % exc_headers[header])
            for j in file_name.columns.values:
                print(j)

            while True:
                try:
                    column_name = input(
                        "What is the name of the header representing %s? " %
                        exc_headers[header])
                    if column_name not in file_name.columns.values:
                        print(
                            "That column name is not within the file. Please try again. Capitalization counts."
                        )
                    else:
                        print("Perfect. I will use that column header instead")
                        exc_headers[header] = file_name.columns.get_loc(column_name)
                        break
                except ValueError:
                    print("Not a valid input. Try again.\n")
                    continue


# headers representing data needed
exc_headers = [
    'RAW_CODE_DISPLAY', 'RAW_CODE_SYSTEM_NAME', 'PRECEDENT_ANALYSIS',
    'PRECEDENT_MATCH', 'PRECEDENT_SIMILARITY', 'PRECEDENT_MAP_COUNT', 'PRECEDENT_RAW_CODE_SYSTEM_IDs', 'PRECEDENT_CONCEPT_ALIAS', 'un_cleaned_display']
acc_headers = [
    'Raw Code Display', 'Map Count', 'Precedent Raw Code System ID(s)', 'Concept Alias', 'val_cleaned_display']

# start time used to determine program completion
# start_time = timeit.default_timer()

# TODO prompt user for path to files
# user_input = raw_input("Enter the path of your file: ")
#
# assert os.path.exists(user_input), "I did not find the file at, "+str(user_input)
# f = open(user_input,'r+')
# print("Hooray we found your file!")

# file paths
excel_workbook_location = (
    "C:\\Users\\ja052464\\OneDrive - Cerner Corporation\\Documents\\NMHS_NE Phase 1 Unmapped Coes To Review.csv")
access_workbook_location = (
    "C:\\Users\\ja052464\\OneDrive - Cerner Corporation\\Documents\\Q_Validated_To_Concept_No_Medications.csv")
keywords_workbook_location = (
    "Y:\\Data Intelligence\\Code_Database\\Proprietary_Code_Display_Keywords.csv")
medication_workbook_location = (
    "Y:\\Data Intelligence\\Code_Database\\Medications_Bank.csv")

# save path
path = "C:\\Users\\ja052464\\Downloads\\"

# read in the needed files
exc_d = pd.read_csv(
    excel_workbook_location,
    encoding='iso-8859-1',
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

key_d = pd.read_csv(keywords_workbook_location, encoding='iso-8859-1', dtype=str).astype(str)

meds_d = pd.read_csv(medication_workbook_location, encoding='iso-8859-1', dtype=str).astype(str)


# todo - Determine column index from column headers
# loops through the headers of the unmapped codes CSV file to check for any header changes.
# If a header is different, prompt user for the header name and use that.

# creates new cleaned data in advance of string comparison
exc_d['un_cleaned_display'] = ""
acc_d['val_cleaned_display'] = ""

# get excel file column locations
column_locations(exc_headers, exc_d)

# get access file column locations
column_locations(acc_headers, acc_d)


# row count used to determine completion time
total_access_rows = acc_d['Raw Code Display'].count()
total_excel_rows = exc_d['RAW_CODE_DISPLAY'].count()


# calculate expected completion time
expected_time = (total_excel_rows * 2) / 60


# tell user completion time.
print(
    "Oh hai. So this might take a bit. I am checking %s unmapped codes against %s validated codes...."
    "That is a lot of math. I expect to be done in roughly %.2f minutes" %
    (total_excel_rows, total_access_rows, expected_time))


print("Cleaning the validated code displays....\n")

for row in acc_d.itertuples():
        percent_completed = (row[0] / total_access_rows) * 100

        if percent_completed % 5 == 0:
            print("%i%% complete...." % percent_completed)

        acc_d.loc[row[0], ['val_cleaned_display']] = token_clean(row[acc_headers[0] + 1], force_ascii=True, full_process=True)

print("Done cleaning validated code displays....")


print("Cleaning the unmapped code displays....\n")
for row in exc_d.itertuples():
    percent_completed = (row[0] / total_excel_rows) * 100

    if percent_completed % 5 == 0:
        print("%i%% complete...." % percent_completed)

    exc_d.loc[row[0], ['un_cleaned_display']] = token_clean(row[exc_headers[0] + 1], force_ascii=True, full_process=True)

print("Done cleaning unmapped code displays....")


# Loops throw the currently unmapped raw code displays and checks it against the previously mapped displays.
for row in exc_d.itertuples():
    i = row[0]
    unmapped = row[exc_headers[8] + 1]
    un_code_sys = row[exc_headers[1] + 1]
    unmapped_code_split = row[exc_headers[8] + 1].split()
    best_display = ""
    prec_analysis = ""
    best_map_count = 0
    best_concept_alias = ""
    code_time = timeit.default_timer()

    # Checks each word against the keyword table for any hits.
    keyword_check = keyword_lookup(unmapped_code_split)

    # gets lev quick score
    lev_match = get_best_match(unmapped, acc_d.itertuples())

    meds_check = medication_check(unmapped, meds_d['drug_name'])

    best_score = lev_match[1]

    # if it is a perfect match, then end loop and write results
    if best_score == 1 and keyword_check:
        best_display = lev_match[0][acc_headers[0]+1]
        best_map_count = lev_match[0][acc_headers[1]+1]
        best_code_system = lev_match[0][acc_headers[2]+1]
    # if it is a close match, run fuzzy match to identify closest legit match
    elif .9 <= best_score < 1 or keyword_check and meds_check is False:

        # calculate the best match using fuzzy
        results = process.extractOne(unmapped, acc_d.iloc[:, acc_headers[4]], scorer=fuzz.ratio)

        # store results
        match_index = results[2]
        best_score = results[1] / 100
        best_display = acc_d.iloc[match_index, acc_headers[0]]
        best_map_count = acc_d.iloc[match_index, acc_headers[1]]
        best_code_system = acc_d.iloc[match_index, acc_headers[2]]
        best_concept_alias = acc_d.iloc[match_index, acc_headers[3]]

    # Checks if code system is within the best match.
    code_sys_check_result = code_sys_check(un_code_sys, acc_d.iloc[i, acc_headers[2]])

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
        prec_analysis = "Low Probability . No Keyword"

    rows_remaining = total_excel_rows - i

    if i % 4 == 0:
        Code_elapsed = timeit.default_timer() - code_time
        print("5 rows completed. It took %.2f sec to calculate. %i rows remaining." %
            (Code_elapsed, rows_remaining))

    # write to dataFrame
    exc_d.at[i, exc_headers[0]] = best_display
    exc_d.at[i, exc_headers[4]] = best_score
    exc_d.at[i, exc_headers[2]] = prec_analysis
    exc_d.at[i, exc_headers[5]] = best_map_count
    exc_d.at[i, exc_headers[6]] = best_concept_alias


# elapsed = (timeit.default_timer() - start_time) / 60

print("I am done doing math. Hold tight while I update the file!")

exc_d.to_csv(path + "testing_export.csv")

print("All done!")

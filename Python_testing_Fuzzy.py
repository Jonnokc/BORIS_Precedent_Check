#!/usr/bin/env python
# encoding: utf-8

import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import new_util
import text

# headers representing data needed
exc_headers = [
    'RAW_CODE_DISPLAY', 'RAW_CODE_SYSTEM_NAME', 'PRECEDENT_ANALYSIS',
    'PRECEDENT_MATCH', 'PRECEDENT_SIMILARITY', 'PRECEDENT_MAP_COUNT', 'PRECEDENT_RAW_CODE_SYSTEM_IDs', 'PRECEDENT_CONCEPT_ALIAS', 'un_cleaned_display']
exc_headers_names = [
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

key_d = pd.read_csv(keywords_workbook_location,
                    encoding='iso-8859-1', dtype=str).astype(str)

meds_d = pd.read_csv(medication_workbook_location,
                     encoding='iso-8859-1', dtype=str).astype(str)


# loops through the headers of the unmapped codes CSV file to check for any header changes.
# If a header is different, prompt user for the header name and use that.

# creates new cleaned data in advance of string comparison
exc_d['un_cleaned_display'] = ""
acc_d['val_cleaned_display'] = ""

# get excel file column locations
new_util.column_locations(exc_headers, exc_d)

# get access file column locations
new_util.column_locations(acc_headers, acc_d)


# row count used to determine completion time
total_access_rows = acc_d['Raw Code Display'].count()
total_excel_rows = exc_d['RAW_CODE_DISPLAY'].count()
total_med_rows = meds_d['drug_name'].count()
total_key_rows = key_d['Keywords'].count()


# tell user completion time.

text.intro(total_excel_rows, total_access_rows, total_med_rows, total_key_rows)


# clean unmapped headers
new_util.row_cleaning(exc_headers, exc_d, total_excel_rows, 8, "unmapped")

# cleans access displays
new_util.row_cleaning(acc_headers, acc_d, total_access_rows,
                      4, "previously validated")

# removes duplicate rows from previously matched
acc_d.drop_duplicates()

print("Analyzing data.... aka of math....")

bar = new_util.progressbar.ProgressBar(maxval=total_excel_rows, widgets=[new_util.progressbar.Bar('=', '[', ']'),
                                                                         ' ', new_util.progressbar.Percentage(),
                                                                         ' ', new_util.progressbar.AdaptiveETA()])

# Loops throw the currently unmapped raw code displays and checks it against the previously mapped displays.
for row in exc_d.itertuples():
    i = row[0]
    unmapped = exc_d.iloc[i, exc_headers[8]]
    un_code_sys = exc_d.iloc[i, exc_headers[1]]
    unmapped_code_split = exc_d.iloc[i, exc_headers[8]].split()
    best_display = ""
    prec_analysis = ""
    best_map_count = 0
    best_concept_alias = ""

    # Checks each word against the keyword table for any hits.
    keyword_check = new_util.keyword_lookup(
        unmapped_code_split, key_d['Keywords'])

    # gets lev quick score
    lev_match = new_util.get_best_match(unmapped, acc_d.itertuples())

    meds_check = new_util.medication_check(unmapped, meds_d['drug_name'])

    best_score = lev_match[1]

    # if it is a perfect match, then end loop and write results
    if best_score == 1 and keyword_check:
        best_display = acc_d.iloc[i, acc_headers[0]]
        best_map_count = acc_d.iloc[i, acc_headers[1]]
        best_code_system = acc_d.iloc[i, acc_headers[2]]
    # if it is a close match, run fuzzy match to identify closest legit match
    elif .9 <= best_score < 1 or keyword_check and meds_check is False:

        # calculate the best match using fuzzy
        results = process.extractOne(
            unmapped, acc_d.iloc[:, acc_headers[4]], scorer=fuzz.ratio)

        # store results
        match_index = results[2]
        best_score = results[1]
        best_display = acc_d.iloc[match_index, acc_headers[0]]
        best_map_count = acc_d.iloc[match_index, acc_headers[1]]
        best_code_system = acc_d.iloc[match_index, acc_headers[2]]
        best_concept_alias = acc_d.iloc[match_index, acc_headers[3]]

    # Checks if code system is within the best match.
    code_sys_check_result = new_util.code_sys_check(
        un_code_sys, acc_d.iloc[i, acc_headers[2]])

    # determines analysis result
    if best_score == 1 and code_sys_check_result:
        prec_analysis = "100% Match"
    elif best_score >= .85 and keyword_check is True:
        prec_analysis = "Possible Match with Keyword"
    elif best_score > .85 and keyword_check is False:
        prec_analysis = "Possible Match no Keyword"
    else:
        prec_analysis = "Low Match Probability"

    # write to dataFrame
    exc_d.at[i, exc_headers_names[3]] = best_display
    exc_d.at[i, exc_headers_names[4]] = best_score
    exc_d.at[i, exc_headers_names[2]] = prec_analysis
    exc_d.at[i, exc_headers_names[5]] = best_map_count
    exc_d.at[i, exc_headers_names[6]] = best_concept_alias

    bar.update(i + 1),

bar.update(total_excel_rows, True),


print("Analysis completed. Hold tight while I write the results to file.\n....\n....")

exc_d.to_csv(path + "testing_export.csv")

print("All done!")

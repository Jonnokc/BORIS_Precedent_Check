import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import timeit
import Levenshtein


start_time = timeit.default_timer()

excel_workbook_location = (
    "C:/Users/ja052464/Downloads/Testing_Python_Similarity4.csv")
access_workbook_location = (
    "C:/Users/ja052464/Downloads/All_Validated_With_Counts.csv")

path = "C:\\Users\\ja052464\\Downloads\\"


exc_d = pd.read_csv(excel_workbook_location).astype(str)
acc_d = pd.read_csv(access_workbook_location,
                    encoding='iso-8859-1').astype(str)


total_access_rows = acc_d['Raw Code Display'].count()
total_excel_rows = exc_d['RAW_CODE_DISPLAY'].count()

expected_time = (total_excel_rows * 5) / 60

print("Oh hai. So this might take a bit. I am checking %s unmapped codes against %s validated codes...."
      "That is a lot of math. I expect to be done in roughly %.2f minutes"
      % (total_excel_rows, total_access_rows, expected_time))

# list(my_dataframe.columns.values)'

# To add a new header
# exc_d['new header name'] = " "

for row in exc_d['RAW_CODE_DISPLAY'].iteritems():
    i = row[0]
    unmapped = row[1]
    test = "random code"
    code_time = timeit.default_timer()
    best_display = ""
    best_score = 0

    for j in acc_d['Raw Code Display']:

        # Very quick way of looping through to find exact and very close matches.
        score = Levenshtein.ratio(unmapped, j)
        # score = fuzz.ratio(unmapped, j)

        if score is 100:
            best_score = score
            best_display = j
            break
        elif score > best_score:
            best_score = score
            best_display = j

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
    # exc_d.set_value(i, ['Python_Concept_Alias'], best_concept_alias)

    print("Best match for %s is %s. Score was %.2f. It took %.2f sec to complete" %
          (unmapped, best_display, best_score, Code_elapsed))


elapsed = (timeit.default_timer() - start_time) / 60

print("I am done doing math. Hold tight while I update the file!")

exc_d.to_csv(path + "testing_export.csv")

print("All done! Finished in %.2f minutes" % elapsed)

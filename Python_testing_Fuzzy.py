import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import timeit


start_time = timeit.default_timer()

excel_workbook_location = (
    "C:/Users/ja052464/Downloads/Testing_Python_Similarity.csv")
access_workbook_location = (
    "C:/Users/ja052464/Downloads/All_Validated_With_Counts.csv")

path = "C:\\Users\\ja052464\\Downloads\\"


exc_d = pd.read_csv(excel_workbook_location).astype(str)
acc_d = pd.read_csv(access_workbook_location,
                    encoding='iso-8859-1').astype(str)


total_access_rows = acc_d['Raw Code Display'].count()
total_excel_rows = exc_d['RAW_CODE_DISPLAY'].count()

expected_time = (total_excel_rows * 14) / 60

print("Oh hai. So this might take a bit. I am checking %s unmapped codes against %s validated codes...."
      "That is a lot of math. I expect to be done in roughly %s minutes"
      % (total_excel_rows, total_access_rows, expected_time))

# list(my_dataframe.columns.values)'

# To add a new header
# exc_d['new header name'] = " "

for row in exc_d['RAW_CODE_DISPLAY'].iteritems():
    i = row[0]
    unmapped = row[1]
    code_time = timeit.default_timer()

    results = process.extractOne(
        unmapped, acc_d['Raw Code Display'], scorer=fuzz.ratio, score_cutoff=50)
    best_display = results[0]
    best_score = results[1]

    Code_elapsed = timeit.default_timer() - code_time

    exc_d.set_value(i, ['Python_Match'], best_display)
    exc_d.set_value(i, ['Python_Similarity'], best_score)

    print("Best match for %s is %s. Score was %d. It took %d sec to complete" %
          (unmapped, best_display, best_score, Code_elapsed))


elapsed = (timeit.default_timer() - start_time) / 60

print("I am done doing math. Hold tight while I update the file!")

exc_d.to_csv(path + "testing_export.csv")

print("All done! Finished in %s minutes" % elapsed)

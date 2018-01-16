#!/usr/bin/env python
# encoding: utf-8

from __future__ import unicode_literals
import functools
from fuzzywuzzy import utils
import Levenshtein
import progressbar
from time import sleep


def progress_with_automatic_max(s1):
    # Progressbar can guess max_value automatically.
    bar = progressbar.ProgressBar()
    for i in bar(range(s1)):
        sleep(0.01)


def token_clean(s1, force_ascii=True, full_process=True):
    sorted1 = process_and_sort(s1, force_ascii, full_process=full_process)
    sorted1 = utils.asciidammit(sorted1)
    return sorted1


def check_for_none(func):
    @functools.wraps(func)
    def decorator(*args, **kwargs):
        if args[0] is None:
            return 0
        return func(*args, **kwargs)
    return decorator


def process_and_sort(s, force_ascii, full_process=True):
    """Return a cleaned string with token sorted."""
    # pull tokens
    ts = utils.full_process(s, force_ascii=force_ascii) if full_process else s
    tokens = ts.split()

    # sort tokens and join
    sorted_string = u" ".join(sorted(tokens))
    return sorted_string.strip()


def medication_check(s1, array):
    if s1 in array:
        return True
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


# yes no prompt
def yes_or_no(question):
    while "the answer is invalid":
        reply = str(input(question + ' (y/n): ')).lower().strip()
        if reply[0] == 'y':
            return True
        if reply[0] == 'n':
            return False


# checks unmapped raw code system against the code systems within the best match
def code_sys_check(s1, s2):
    for code in s1:
        if code.lower() in s2.lower():
            return True
    return False


# checks unmapped code display words against keywords table
def keyword_lookup(s1, array):
    for word in s1:
        for keyword in array:
            if word.lower() in keyword.lower():
                return True
    return False


def column_locations(preset_headers, file_name):
    for header in range(len(preset_headers)):
        if preset_headers[header] in list(file_name.columns.values):
            preset_headers[header] = file_name.columns.get_loc(
                preset_headers[header])

        else:
            print(
                'Could not find a column with the header %s. I will need you to enter the header representing this column within the data file.\n Your options are.. \n '
                % preset_headers[header])
            for j in file_name.columns.values:
                print(j)

            while True:
                try:
                    column_name = input(
                        "What is the name of the header representing %s? " %
                        preset_headers[header])
                    if column_name not in file_name.columns.values:
                        print(
                            "That column name is not within the file. Please try again. Capitalization counts."
                        )
                    else:
                        print("Perfect. I will use that column header instead")
                        preset_headers[header] = file_name.columns.get_loc(
                            column_name)
                        break
                except ValueError:
                    print("Not a valid input. Try again.\n")
                    continue


def row_cleaning(headers, df, rows, output, file):
    print("Cleaning the %s code displays...." % file)
    bar = progressbar.ProgressBar(maxval=rows, widgets=[progressbar.Bar(
        '=', '[', ']'), ' ', progressbar.Percentage()])
    for row in df.itertuples():

        df.iloc[row[0], headers[output]] = token_clean(
            row[headers[0] + 1], force_ascii=True, full_process=True)
        bar.update(row[0] + 1)

    bar.update(rows, True)
    bar.update(rows, True)
    bar.update(rows, True)

    print("\nDone cleaning the %s code displays!\n....\n...." % file)

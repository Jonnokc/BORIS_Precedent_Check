import sys
import time
from time import sleep

# wpm
typing_speed = 90


def print_slow(text):
    for letter in text:
        time.sleep(0.03)
        sys.stdout.write(letter)
        sys.stdout.flush()

def intro(total_excel_rows, total_access_rows, total_med_rows, total_key_rows):
    print_slow(
        "Oh hai there. My name is BORIS. Looks like you want to run the precedent check. I can do that....\n")

    sleep(2)

    print_slow("Okiedokie so we have " + str(total_excel_rows) + " unmapped code displays to check. This will take me a bit... \n"
                    "I will need to check each and every one of these displays against" +
                    str(total_access_rows) + " previously mapped displays, "
                    + str(total_med_rows) + " medications, and " + str(total_key_rows) + "....\n" +
                    "Needless to say that's a lot of math!\n\n")

    sleep(4)

    print_slow("Anyway... let's get to it. For the most part I will handle things. You can just watch the pretty status bars below. \n""
               "If I need something from you I will ask you in the text area below.\n\n")

    sleep(4)

    print_slow("Oh one quick thing.... Kind of important.... See the speed at which I can do math (aka finish this) is directly related to how hard your PC is working. \n"
                    "For the most part I will be fine, but watching a video while I am working will make this take twice as long. \n"
                    "So if you want me to finish as quickly as I can hold off on your cat videos and other cpu heavy stuff until I am done... Okay that was the last thing. Let's get started!\n\n")

    sleep(2)

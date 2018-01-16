#!/usr/bin/env python
# encoding: utf-8

from . import fuzz
from fuzzywuzzy import utils
import logging
from functools import partial


default_scorer = fuzz.WRatio

default_processor = utils.full_process


def new_extract_without_order(query, choices, processor=default_processor, scorer=default_scorer, score_cutoff=0):
    """Select the best match in a list or dictionary of choices.

    Find best matches in a list or dictionary of choices, return a
    generator of tuples containing the match and its score. If a dictionary
    is used, also returns the key for each match.

    Arguments:
        query: An object representing the thing we want to find.
        choices: An iterable or dictionary-like object containing choices
            to be matched against the query. Dictionary arguments of
            {key: value} pairs will attempt to match the query against
            each value.
        processor: Optional function of the form f(a) -> b, where a is the query or
            individual choice and b is the choice to be used in matching.

            This can be used to match against, say, the first element of
            a list:

            lambda x: x[0]

            Defaults to fuzzywuzzy.utils.full_process().
        scorer: Optional function for scoring matches between the query and
            an individual processed choice. This should be a function
            of the form f(query, choice) -> int.

            By default, fuzz.WRatio() is used and expects both query and
            choice to be strings.
        score_cutoff: Optional argument for score threshold. No matches with
            a score less than this number will be returned. Defaults to 0.

    Returns:
        Generator of tuples containing the match and its score.

        If a list is used for choices, then the result will be 2-tuples.
        If a dictionary is used, then the result will be 3-tuples containing
        the key for each match.

        For example, searching for 'bird' in the dictionary

        {'bard': 'train', 'dog': 'man'}

        may return

        ('train', 22, 'bard'), ('man', 0, 'dog')
    """
    # Catch generators without lengths
    def no_process(x):
        return x

    try:
        if choices is None or len(choices) == 0:
            raise StopIteration
    except TypeError:
        pass

    # If the processor was removed by setting it to None
    # perfom a noop as it still needs to be a function
    if processor is None:
        processor = no_process

    # Run the processor on the input query.
    processed_query = processor(query)

    if len(processed_query) == 0:
        logging.warning(u"blank returning 0 "
                        "[Query: \'{0}\']".format(query))

    # Don't run full_process twice
    if scorer in [fuzz.WRatio, fuzz.QRatio,
                  fuzz.token_set_ratio, fuzz.token_sort_ratio,
                  fuzz.partial_token_set_ratio, fuzz.partial_token_sort_ratio,
                  fuzz.UWRatio, fuzz.UQRatio] \
            and processor == utils.full_process:
        processor = no_process

    # Only process the query once instead of for every choice
    if scorer in [fuzz.UWRatio, fuzz.UQRatio]:
        pre_processor = partial(utils.full_process, force_ascii=False)
        scorer = partial(scorer, full_process=False)
    elif scorer in [fuzz.WRatio, fuzz.QRatio,
                    fuzz.token_set_ratio, fuzz.token_sort_ratio,
                    fuzz.partial_token_set_ratio, fuzz.partial_token_sort_ratio]:
        pre_processor = partial(utils.full_process, force_ascii=True)
        scorer = partial(scorer, full_process=False)
    else:
        pre_processor = no_process
    processed_query = pre_processor(processed_query)

    try:
        # See if choices is a dictionary-like object.
        for key, choice in choices.items():
            processed = pre_processor(processor(choice))
            score = scorer(processed_query, processed)
            if score == 1:
                yield (choice, score)
            elif score >= score_cutoff:
                yield (choice, score, key)
    except AttributeError:
        # It's a list; just iterate over it.
        for choice in choices:
            processed = pre_processor(processor(choice))
            score = scorer(processed_query, processed)
            if score == 1:
                yield (choice, score)
            elif score >= score_cutoff:
                yield (choice, score)


def new_extract_one(query, choices, processor=default_processor, scorer=default_scorer, score_cutoff=0):
    """Find the single best match above a score in a list of choices.

    This is a convenience method which returns the single best choice.
    See extract() for the full arguments list.

    Args:
        query: A string to match against
        choices: A list or dictionary of choices, suitable for use with
            extract().
        processor: Optional function for transforming choices before matching.
            See extract().
        scorer: Scoring function for extract().
        score_cutoff: Optional argument for score threshold. If the best
            match is found, but it is not greater than this number, then
            return None anyway ("not a good enough match").  Defaults to 0.

    Returns:
        A tuple containing a single match and its score, if a match
        was found that was above score_cutoff. Otherwise, returns None.
    """
    best_list = new_extract_without_order(
        query, choices, processor, scorer, score_cutoff)
    try:
        return max(best_list, key=lambda i: i[1])
    except ValueError:
        return None

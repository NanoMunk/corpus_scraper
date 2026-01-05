This is a web scraper script I made to fetch example sentences from the 2019 Finnish corpus of Leipzig University. The sentences are imported into the vocabulary learning app I use.

The words to be looked up are in column C of search_file.xlsx. The example sentences are pasted into column A. Columb B is left empty for the translation of the sentences.

It requires no additional input. Once the python file is run it goes through all the words in the .xlsx file.

Known issues:

If the corpus loads too slow the scraper takes it as no results were found. Most of the time, the current explicit wait is enough.

Sometimes the website itself gets confused and serves other corpi instead of the Finnish corpus. RUS, DE, ES, etc. It's rare, but can happen. Very rarely it gets stuck in the incorrect corpus too. This can be caught in the print statements.
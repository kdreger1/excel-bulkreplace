# excel-bulkreplace-odk

XLSM (macro-enabled Excel spreadsheet) to do a bulk find-and-replace in the survey tab of an ODK Excel spreadsheet using patterns defined in another Excel spreadsheet. Windows OS only.

Important: When you open the PMA_findAndReplace spreadsheet, you must allow macros to function.

In the spreadsheet, you will be able to modify two cells with the path and filenames for the pattern file and the ODK file. To run the find-and-replace, press the blue button that says "Run Find-and-Replace"

The pattern file is an Excel (.xlsx only) document you will need to create. An example file called "testingInputPatterns.xlsx" is included in this repository. The first column are the words to search for; the second column are what each word is replaced by. Do not include column headers.

A "word" in the context of this module is a series of characters that form a discrete unit and is separated from surrounding text by a space or non-alphanumeric character (such as a period). For example, "101" will not match part of the string "101001" nor will it match "101a" but will match the string "101.". The only non-alphanumeric character that can be matched on is an underscore. Rows in the Excel spreadsheet that are missing a change-from field or a change-to field are ignored. Also ignored are rows with non-changes; example: 101a -> 101a.

This program will warn the user if there are nonblank find-for patterns in the Excel document that are not in the ODK document. This does not produce an error however, and the find-and-replace will continue after the warning message.

There are three modes this find-and-replace feature can run in:

 1. Highlight cells with any changes, but do not make them
 2. Highlight cells with any changes and make any such changes
 3. Make the changes without highlighting the cells

In any case, the changes are applied to a copy of the ODK document that is created and becomes the active document. It is up to the user to name the new file and save it.

# PhoneNumberCleaner
VBA word macro to clean up phone numbers

# Set Up
Import PhoneCleaner.bas into Normal, run BuildPhoneCleanerForm once, then use the PhoneCleaner macro to open the pop-up.

The tool ignores dates, percentages, and labels. It also recognizes common formats such as 2107900232, 210.790.0232, and 1-210-790-0232, and it removes duplicates.

Limits to be aware of: 
The tool only recognizes 10-digit U.S. numbers, so it skips international numbers and extensions. If a line contains no number it can recognize, the status bar shows how many lines were skipped so you can review them by hand rather than lose them silently.

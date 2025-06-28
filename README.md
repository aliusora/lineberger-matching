Lineberger Service Requests Matching
====================================

This project contains an R script to filter, clean, and match service requests 
to the Lineberger Comprehensive Cancer Center Directory. The goal is to keep 
only the requests that matter for this year and make sure the Principal 
Investigators (PIs) are in the directory, even if they used nicknames or 
different spellings.

------------------------------------
What this pipeline does
------------------------------------

1. Clean the Lineberger Directory
   - Make all names uppercase and trim extra spaces.
   - Remove extra parts like “Jr.”, “Sr.”, or initials.
   - Extract nicknames in parentheses.
   - Extract the username part of the email address (before the @).

2. Filter Service Requests
   - Convert date fields to calendar dates.
   - Keep only requests that were:
     * Submitted on or after Jan 1, 2024, OR
     * Closed on or after Jan 1, 2024.
   - This keeps any active or recent request for this year.

3. Clean PI Names & Emails
   - Uppercase PI names and remove punctuation.
   - Extract username part of PI emails.
   - Use a nickname list (like Bob = Robert, Jake = Jacob) to expand names 
     in both directions.

4. Build Match Keys
   - Create multiple keys for matching:
     * Clean name
     * Nickname version
     * Reverse nickname version
     * Email username

5. Fuzzy Match
   - Cross join all request keys with directory keys.
   - Use fuzzy matching to catch typos and small spelling differences.
   - Keep matches within a 20% distance threshold.

6. Export Final Matches
   - Keep only requests with at least one valid match.
   - Save final matched requests to an Excel file for review.

------------------------------------
Key Files & Scripts
------------------------------------

- matching_script.R : Main script to run everything.
- Input files:
  * directory (CSV or Excel)
  * requests (Excel)
- Output file:
  * matched_requests.xlsx

------------------------------------
Libraries Used
------------------------------------

readr, readxl, dplyr, stringr, stringdist, janitor, lubridate, tidyr, 
writexl, purrr

------------------------------------
How to Run
------------------------------------

1. Open matching_script.R in RStudio Cloud or RStudio Desktop.
2. Load your directory and requests data.
3. Run the script step by step or all at once.
4. Check the output: matched_requests.xlsx will be in your working folder.

------------------------------------
Why This Matters
------------------------------------

This process:
- Keeps only requests from this year.
- Handles nicknames, initials, and typos.
- Makes your data accurate and consistent.
- Creates a clear, reproducible record for reporting.

------------------------------------
Contact
------------------------------------

Internal use only.
Contact the Lineberger data team if you have questions.

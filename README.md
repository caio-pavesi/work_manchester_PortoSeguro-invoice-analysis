# Porto Seguro pdf invoice analysis
## A fully functional EXAMPLE project written in Python to search for invoice in a inbox within the Outlook DESKTOP application and analyze it as demanded!

This project is just a demonstration of a similir project of mine, developed at the "Manchester Investimentos" financial department. Every part of this project is sample code which shows how to do the following:
* Search for an invoice with pywin32.
* Obtain additional data in metabase.
* Read the invoice using pdfplumber.
* Use pandas to read support data.
* Compare data for the analysis*.
* Report the results by e-mail.

Our analysis was not that complex but the time spent on it was huge, thats why this automation come in handy as it saves time doing comparisons between data. pywin32 / win32com (similar to VBA) is used, in this case, to automate the Outlook DESKTOP application retrieving the target invoice based on current month. pdfplumber automates the gathering of data from the invoice while we use RegEx to filter dow only the information that we need. pandas is used to create DataFrames to store data from both the invoice and the support documents and compare them for the analysis*. To finish the process we again use pywin32 to send the results by e-mail.
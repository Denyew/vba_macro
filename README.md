#Excel VBA Macros

Macros created under own initiative to streamline repetitve data manipulation/repetitive processes that have minimal variability in results

**GLDI Summarizer**

[GLDI_Summarizer](gldi_summary.vba)

This macro helps to manipulate the data file from Oracle and creates a separate Pivot Table for a quick Profit and Loss statement view, based on account code and cost center.

Original Process requires column adding to obtain the total amount (nett of Credit and Debit), along with looking up account descriptions from a separate Chart of Accounts file, and creating a Pivot Table to drag and drop required fields

Current process only requires selecting the downloaded file using a File Dialog, and immediately presenting it in the Pivot Table, with an added option to export the file with timestamps

Process time - 4 mins -> 15 seconds

**Sales Report Generator**

[Sales Report](sales_report.vba)

This macro removes the need to manually delimit sales data and append it in another file, plus it helps to pivot data from a .pbix output, which is this macro [.pbix_output](pbi_output.vba)

Time savings come from not needing to manually manipulate data and removes waiting time to open files

Process time - 10 minutes -> 3 minutes

**Journal Processor**

[Journals](journal.vba)

This macro automatically populates data into the Oracle Journal template based on the nature of the journal. However, it's all based on the same source so manipulation is done via the macro

Current process requires less time in reconciling as it is populating a blank template rather than reusing the previous month's file as data may not be exactly the same

Process time - 15 minutes (longer if there are inconsistencies) -> 6 minutes (waiting time after running the macro)

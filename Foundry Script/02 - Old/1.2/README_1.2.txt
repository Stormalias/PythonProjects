*********************************************
Foundry Monitoring Script by Praval Visvanath
*********************************************

***SUMMARY DESCRIPTION***
This script ingests data from one csv file, checking for errors in the MTD and YTD for different benches, and prints a detailed
table of the errors in the portfolios of each bench. This output is stored in a word document, the results of which can be directly
pasted into the Teams chat to report to CR. This is done for the morning and afternoon reporting, with the latter updating any
information about the first.

***USAGE***
1.	To ensure the script can run, the system will need python3 installed, as well as the python-docx package.
Python can be installed from the Microsoft Store. Once you have installed Python 3, open command prompt (cmd)
and run the following command:
>pip3 install python-docx

2.	Once the relevant items are installed, the script can be run. For the script to run, it will require two
folders in the directory it is placed in. Note that the script can run from anywhere, but the output files will
be in that same directory, and the input files must be in folders in that directory.(i.e. if the script is as "C:/Script.py", the output
doc will be as "C:/report.docx" and the input csv must be placed as "C:/Morning/export.csv" or "C:/Afternoon/export.csv")

3.	Ensure 2 folders in the directory you intend to run the script from are named "Morning" and "Afternoon" (capitalisation
important) [If you downloaded the entire zipfile containing this README and the script, the folders should already exist here]

4.	When in the foundry environment, navigate to 'PNL050_Portfolio_DTD_Check and select every row in the Bench_error_check_pivot_table.
bring up the data tab if it isn't already expanded. 

5. On this data tab, hit 'export' as well. Refer to CapA.png and CapB.png for a screenshot of the table selection, this tab and where 
to find the export button

6. You should now have the file downloaded, namely 'contour-export-(TODAYS DATE IN MM-DD-YYY)'. Move this file
to the Morning and Afternoon folder respectively. If running a Singular report (explanined below), you only need to have file in the 
Morning folder.

7.	Run Foundry Monitoring Script.py from the folder. Follow the instructions on the CLI screen. The Morning report and Afternoon report 
are self explanatory. The Singular report option is for the case where there are issues with files and pipeline building due to 
dependencies on other platforms. In this case, due to the delay you would only need to send the one report in the afternoon (UTC+8 1300).
In this case, you only need the relevant file in the "Morning" folder. 

8.	After the dialog box confirming the output, check the directory where the scipt is to find the .docx file. 

9.	Open the file. The output table may not show up with lines but rest assured it is in table format. You may select all (control + A) and copy(control + C)
and thereafter paste (control + V) into the teams chat to report to CR. 

10. Enjoy not needing to filter/type/screenshot for the report anymore.



If there are any issues/suggestions please contact the author
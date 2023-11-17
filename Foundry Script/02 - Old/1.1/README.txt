*********************************************
Foundry Monitoring Script by Praval Visvanath
*********************************************

***SUMMARY DESCRIPTION***
This script ingests data from 2 csv files, checking for errors in the MTD and YTD for different benches, and prints a detailed
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
be in that same directory, and the input files must be in folders in that directory. 

3.	Ensure 2 folders in the directory you intend to run the script from are named "Morning" and "Afternoon" (capitalisation
important) [If you downloaded the entire zipfile containing this README and the script, the folders should already exist here]

4.	When in the foundry environment, when needing to check which benches have errors, instead hit 'export' from the 
Bench_error_check_pivot_table. After this, select each row with an 'ERROR' in the overall_check column and bring up the data tab
if it isn't already expanded. 

5. On this data tab, hit 'export' as well. Refer to data_tab_closed.png and data_tab_open.png for a screenshot of this tab and where 
to find the export buttons

6. You should now have 2 files downloaded, namely 'export.csv' and 'contour-export-(TODAYS DATE IN MM-DD-YYY)'. Move these files
to the Morning and Afternoon folder respectively. If running a Singular report (explanined below), you only need to have files in the 
Morning folder.

7.	Run Foundry Monitoring Script.py from the folder. Follow the instructions on the CLI screen. The Morning report and Afternoon report 
are self explanatory. The Singular report option is for the case where there are issues with file and pipeline building due to 
dependencies on other platforms. In this case, due to the delay you would only need to send the one report in the afternoon (UTC+8 1300).
In this case, you only need the relevant files in the "Morning" folder. 

8.	After the dialog box confirming the output, check the directory where the scipt is to find the .docx file. 

9.	Open the file. The output table will not show lines. But it is in table format. You may select all (control + A) and copy(control + C)
and thereafter paste (control + V) into the teams chat to report to CR. 

10. Enjoy not needing to filter/type/screenshot for the report anymore. NOTE THAT when using it for consecutive days, that you should replace
the 'export.csv' file in the folders, as this name does not change per day. The Contour export is dated and thus will not have any issues. 
If it makes it neater for you, it is recomended to delete old .csv files before starting reporting for the day. 



If there are any issues/suggestions please contact the author
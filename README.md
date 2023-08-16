# Automated-Labeling-Tool
This automated labeling tool was created to streamline the data scrubbing process which looks for common words used in an excel file and replaces empty or N/A labels cells with the correct information.


# Introduction
This repository displays the automated labeling tool that I created to help me in my day job as a Data Analyst for The Disneyland Resort. <B>The data used for this repository is public data of IMDB’s Disney+ shows’ information. The reason is I’m not able to use the actual excel file due to its sensitive nature.</B> This tool was created to help deal with my ongoing process of data scrubbing by updating cells with the correct labels and filling in empty cells in an automated way. I have included the original excel file with the python file of the automated labeling tool.


# Abstract
For anyone who deals with data, it is understood that most of your time will be spent cleaning data (e.g., data scrubbing) and arranging the data to find trends. In my day as a Data Analyst, the original process of data scrubbing took about five hours to complete. After researching, a solution was found to use python to work with Microsoft Excel and help streamline and cut down the time spent on data scrubbing. This tool has saved my department hours per week, which would be spent on data scrubbing.


# Summary of Skills
This tool was created in the Python programming language and uses the Openpyxl Python module to work with Microsoft Excel. I will also use a pattern fill feature from the Openpyxl Python module to highlight the updated cells. I have Python 3.10 as my environment and it can work with any Python 3 environment.


# Preview
![Preview of version 2 tool.](https://github.com/micgonzalez/Automated-Labeling-Tool/blob/main/Automated_Labeling_Tool_Images/preview%20of%20version%2002%20tool.jpg)

This is a preview of version two results of Automated Labeling tool.


# Findings
It was amazing to find out about python modules that allow you to work with excel files. The use of the Openpyxl module made a world of difference for me. It allowed me to create this automated labeling tool to help in data scrubbing. There is still more room for features that I want to include for this tool. 


# Challenges
The challenge for version two of automated labeling tool was researching on an efficient method to read keywords from a list or from a different column and update cells with correct information. This was an informative endeavor, there are some sections of the code that will be updated with a similar method. Another challenge was to find information related to having python tell Excel which cell will be highlighted. The point of highlighting cells is to show which cells were updated with the correct by the index column and other multiple columns. One last challenge was allotting time to work on this passion project, while working my day job.


# Conclusion
Updating the code of the automated labeling tool was a continue motivation to better this code. Slowly adding features, such as searching for keywords in a list or in another cell from a different column. Once a match is found, the cell is updated with the correct information. After the cell is updated, the same cell and index cell will be highlighted to show which row and cell was updated. This tool has cut down the time spent on cleaning data and gave me more time to explore the data.

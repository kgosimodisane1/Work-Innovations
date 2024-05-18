# The Private Equity Fund Finder

## How it works

The Fund Finder is designed to locate specific funds from a long list of funds and then hide everything else. The file labelled "List1" has a list of funds that any analyst 
would use to add names of funds that they are familiar with. The larger list that would be searched through can be found in the FundFinder excel file. For relatability, 
I have instead replaced the fund names in the files with South African and American-listed stocks. 

The VBA works by looking for specific words **from** "List1" **in** the FundFinder file. This means that the VBA can locate the same name twice and select both, even if the 
names slightly differ. For example, one name in "List1" is Sasol, but in the FundFinder file you can see that Sasol is listed twice as **Sasol Limited - XNYSE: SSL** and 
**Sasol Limited - XJSE: SOL**. Once the file is run, it will be able to identify both names despite being listed in different stock exchanges.

## The motivation

The idea is to improve the analysts' productivity by exposing them to funds that they are familiar with and that they can analyse and interpret more efficiently. If widely 
adopted, this should result in a significant increase in the team's productivity since the team is exposed to work that they somewhat _specialise_ in. 

## How to use it

To run the file, both files will need to be downloaded or saved in your cloud drive and you will need to open your VBA programme in the FundFinder.xlsm. A shortcut to opening the 
VBA is Alt + F11 on Windows. From there, you will need to change the file location of "List1" in the VBA to match where you have stored it. 
My "List1" file is currently stored in my documents folder in my OneDrive as follows:

wsList1 = Workbooks.Open("C:\Users\lenovo\OneDrive\Documents\List1.xlsx").Sheets("Sheet1")

Finally, you will press the run button which is the green play icon in that top tab and watch the magic.
I have attached a picture the icon below.


![image](https://github.com/kgosimodisane1/Work-Innovations/assets/159646111/46bb47d3-de45-49f3-a9c6-45f902c7576f)


## Known issues (Work in progress)

In order to locate the funds/stocks from List1, I would need to copy the VBA code from the FundFinder macro and paste it into the macro of the **daily workflow allocation** 
(which is an excel file of work that gets assigned to the team at the start of every day) and then run it. This can be a bit time-consuming and tedious. The final improvement to
this idea would be to include it as an **add-in** in excel. With this, no matter what excel file I open, I will be able to run the VBA code with just a click of a button. 

Unfortunately, the company I work for has IT restrictions that prevent me from making this an excel add-in, although any idea to improve on this idea's productivity will be 
explored.

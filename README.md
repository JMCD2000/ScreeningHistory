# ScreeningHistory
This is a VBA project that opens weekly excel reports and then updates the Access Database and or a remote database found under the Screencatcher_Analyst project. This automates the populating of the weekly data to build combined tables and metrics in Access with VBA

I would have loved to have done this project in Python, but my work doesn't allow it. This runs using VBA written in Modules in MS Access 2016. The more tables there are the slower this runs. This is just an issue with using VBA and 32bit office. Again, it is what my work allows and has configured.

I run weekly exports from a database that exports to excel to provide a status of open items, a burn rate, a class comparision, and responsibity for action. These weekly reports are used to build the data tables. This ScreeningHistory code was developed in 2015 to automate the aggration of screening data. It has a slow run time and the project met the goals but was not activly used until the Screencatcher_Analyst was written in Feburary 2020 to find the curn of rescreenings, changes in responsibility from Government to Contractor, Number of screenings and combinations, and other metrics. The manual cleaning and loading of Excel reports into Access was time consuming and could introduce errors. So this project was dusted of and modified to be able to use the Screencatcher_Analyst database.

The code contained in ScreeningHistory makes it so that the below excel data file manual clean up is not needed before importing the Excel into the Access Screencatcher_Analyst. 

Column header cleanup (used on the Final only) Trial_Card Star Pri Safety Screening Act_1 Act_2 Status Action_Taken Date_Discovered Date_Closed Trial_ID Event TC_Screening TC_Screening_AC1_AC2 Final_Sts_A_T =CONCATENATE(E2,"/",F2,"/",G2," ",H2,"/",I2) =CONCATENATE(E2,"/",F2,"/",G2) =CONCATENATE(,H2,"/",I2)
(used on all other reports) Trial_Card Star Pri Safety Screening Act_1 Act_2 Status Action_Taken Date_Discovered Date_Closed Trial_ID Event TC_Screening TC_Screening_AC1_AC2 =CONCATENATE(E2,"/",F2,"/",G2," ",H2,"/",I2) =CONCATENATE(E2,"/",F2,"/",G2)

Clean up cell values to be compatible with Access CONVERT STAR="" TO "STAR" SCREEN="**" TO "AST"; clear out any bad symboles, "-", "@" AC1/2="***" TO "AST"; clear out any bad symboles, "-", "@" "ORACLE DATES" DASHes TO "MICROSOFT DATES" SLASHes CLEAR OUT EMPTY DATE FIELDS Columns A:I and L:P format as Text Columns J:K format as Date Sort A-Z on Trial_Card and Save

On import to Access save the new table with this format. Nontrial beans YYYY/MM/DD_LPD17

Trial beans YYYY/MM/DD_LPD17_BT YYYY/MM/DD_LPD17_AT YYYY/MM/DD_LPD17_FCT YYYY/MM/DD_LPD17_OWLD YYYY/MM/DD_LPD17_Final

Option Compare Database
Option Explicit

Public Function OpenXls_BeanReport(ByVal strPath As String, ByVal strPath_Database As String, ByVal myFileName As String, ByVal hullNum As String, ByVal curClassHull As String, ByVal reportDate As Date, ByVal ScrCtchr_reportDate As String, ByVal fileReport As String, ByVal dbTarget As String) As Integer
'This function opens the file name and path of all the passed in file and writes the data to the table
    'strPath: The path to get the current file
    'strPath_Database: The path to the target DB
    'myFileName: The current excel or text file name
    'hullNum: the current ship hull number
    'reportDate: the Month/Day/Year
    'ScrCtchr_reportDate: the Year/Month/Day
    'fileReport: this determins the presumed format of the cells
    'dbTarget: is the targeted database that has a specfic function and data structure
    'Function returns an integer count when completed
    
    Dim dbs As DAO.Database 'This is this database
    Dim rst As DAO.Recordset 'This is this database
    Dim dbs_remote As DAO.Database 'This is the remote Screencatcher database
    Dim rst_remote As DAO.Recordset 'This is the remote Screencatcher database
    Dim myTC_Number As String
    Dim myLenTC As Integer
    Dim mySTAR As String
    Dim myPRI As String
    Dim mySAFETY As String
    Dim mySCREEN As String
    Dim myAC1 As String
    Dim myAC2 As String
    Dim mySTATUS As String
    Dim myActTkn As String
    Dim myMM_Disc As String '
    Dim myDD_Disc As String '
    Dim myYYYY_Disc As String '
    Dim myDateDisc As Date
    Dim myMM_Close As String '
    Dim myDD_Close As String '
    Dim myYYYY_Close As String '
    Dim myDateClose As Date
    Dim myTrial_ID As String
    Dim myLenEvent As Integer
    Dim myEVENT As String
    Dim myScrnConCat As String ' "XX/XXXX/XXXX/X/X"
    Dim myScrnConCatLong As String ' "XX/XXXX/XXXX X/X"
    Dim myScrnConCatShort As String ' "XX/XXXX/XXXX"
    Dim tableFinal As Boolean ' the Final table captures the Status and Action Taken
    Dim myScrnConCatFinal As String ' the Final table captures the Status and Action Taken
    'Dim myPriConCat As String
    Dim recsetDate As Date
    
''''Read worksheet range into array
    'uses late binding to open excel workbook and open it line by line
    'make reference to Microsoft Excel xx.x Object Model to use native functions, requires early binding however
    Dim xlApp As Object 'Excel.Application
    Dim xlWrk As Object 'Excel.Workbook
    Dim xlSheet As Object 'Excel.Worksheet
    Dim myWkShtName As String 'Excel.Worksheet
    
    Dim currentFile As String 'Path and File name combined
    
    'Dim mySQL_search As String
    Dim mySearch As String
        
    Dim myTblName As String 'ScreeningHistory Database table name
        
    Dim tblScreencatcher As String 'Screeningcatcher Database table name
            
    Dim rowPointer As Integer 'Current row copy from, paste too
    rowPointer = 0 'Set to zero, in the For loop it is set to 2, header row is 1

' Excel File
    Set xlApp = VBA.CreateObject("Excel.Application")
    'toggle visibility for debugging (True, False)
    xlApp.Visible = False
    'assign the path to the Excel Workbook
    currentFile = strPath & "\" & myFileName
    '~~Debug.Print "currentFile : " & currentFile
    'Set the path to the Excel Workbook and Open it
    Set xlWrk = xlApp.Workbooks.Open(currentFile)
    'Set xlSheet = xlWrk.Sheets("Sheet1")
    'Set xlSheet = xlWrk.Sheets("TSM EXPORT")
    If fileReport = "TSM_EXPORT" Then
        myWkShtName = "TSM EXPORT"
    ElseIf fileReport = "Bean_Data" Then
        myWkShtName = hullNum & " Cards"
    End If
    'Set the Excel Work book to the target work sheet
    Set xlSheet = xlWrk.Sheets(myWkShtName)
    
''''Build current worksheet range row number
    Dim myCurRng_Dwn As Long 'Get the last non-blank cell in the current column
        myCurRng_Dwn = xlSheet.Range("A1").CurrentRegion.Rows.Count
    'Dim lastRow As Long 'Get the last non-blank cell in the current column
    '    lastRow = xlSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print "   Row values of myCurRng_Dwn : " & myCurRng_Dwn
    'Debug.Print "   Row values of lastRow : " & lastRow
        
''''Build current worksheet range column number
    Dim myCurRng_Acr As Long 'Get the last non-blank cell in the current row
        myCurRng_Acr = xlSheet.Range("A1").CurrentRegion.Columns.Count
    'Dim lastCol As Long 'Get the last non-blank cell in the current row
    '    lastCol = xlSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Dim col_Ltr As String 'Get the column leter of the last column
        col_Ltr = Split(xlSheet.Cells(1, myCurRng_Acr).Address, "$")(1)
    Debug.Print "   Column values of myCurRng_Acr : " & myCurRng_Acr
    'Debug.Print "   Column values of lastCol : " & lastCol
    Debug.Print "   Column leter values of col_Ltr : " & col_Ltr
    
''''Build current worksheet range reference
    Dim myCurRange As String
    myCurRange = ("A1:" & col_Ltr & myCurRng_Dwn) '("A1:B2") is format of Var
    Debug.Print "   Range value of myCurRange : " & myCurRange
    
''''Read worksheet range into array
    Dim mySheetArray As Variant 'read whole data area as one array
    mySheetArray = xlSheet.Range(myCurRange).Value
    
    'Right here the Excel could be closed because the worksheet has been read into an array
    xlWrk.Close savechanges:=False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlWrk = Nothing
    Set xlApp = Nothing
    
''''Read rows from array
    Dim myCol_A As String 'Trial Card Number
    Dim myCol_B As String 'Star
    Dim myCol_C As String 'Priority
    Dim myCol_D As String 'Safety
    Dim myCol_E As String 'Screen
    Dim myCol_F As String 'Action Code 1
    Dim myCol_G As String 'Action Code 2
    Dim myCol_H As String 'Status
    Dim myCol_I As String 'Action Taken
    Dim myCol_J As String 'Date Disc
    Dim myCol_K As String 'Date Closed
    Dim myCol_L As String 'Trial ID
    Dim myCol_M As String 'Event
    
    
    'Open the database objects
    If dbTarget = "Screencatcher" Then
        'Open the Screencatcher Database for a single instance
        Set dbs_remote = DBEngine.Workspaces(0).OpenDatabase(strPath_Database)
        'Open the ScreeningHistory Database for logging of file actions
        Set dbs = CurrentDb
    ElseIf dbTarget = "ScreeningHistory" Then
        'Open the ScreeningHistory Database for logging and screening
        Set dbs = CurrentDb
    End If
        
    For rowPointer = 2 To (myCurRng_Dwn - 1)
        'rowPointer = 1 is the column headers and is not used here
        'rowPointer = row number that is zero based index _
        row 0 is the column header found in A1:Col(X)Row(X) _
        if there are no column headers used then .Offset((rowPointer - 1), 0) to index correctly _
        if there is column headers used then .Offset(rowPointer, 0) to index correctly
        'Zero based index pointer is used in the "With" & "Offset" Statements
        '.Offset(rowPointer, 0) is (Varible row pointer Int, Fixed column pointer)
    
        'do my work in the array
        myCol_A = mySheetArray(rowPointer, 1) 'Trial Card Number
        myCol_B = mySheetArray(rowPointer, 2) 'Star
        myCol_C = mySheetArray(rowPointer, 3) 'Priority
        myCol_D = mySheetArray(rowPointer, 4) 'Safety
        myCol_E = mySheetArray(rowPointer, 5) 'Screen
        myCol_F = mySheetArray(rowPointer, 6) 'Action Code 1
        myCol_G = mySheetArray(rowPointer, 7) 'Action Code 2
        myCol_H = mySheetArray(rowPointer, 8) 'Status
        myCol_I = mySheetArray(rowPointer, 9) 'Action Taken
        myCol_J = mySheetArray(rowPointer, 10) 'Date Disc
        myCol_K = mySheetArray(rowPointer, 11) 'Date Closed
        myCol_L = mySheetArray(rowPointer, 12) 'Trial ID
        myCol_M = mySheetArray(rowPointer, 13) 'Event
        
        myLenEvent = (Len(myCol_M)) 'Col M EVENT Len AT or BT = 2, FCT = 3
        myLenTC = 7 + myLenEvent + 9 ' 7 = LPD0017, 9 = -DC000101
        myTC_Number = myCol_A 'Col A DSP or TC_Number
        
        If Len(myTC_Number) = 8 Then
            'This is only Dept/Serial/Part
            'DC000101
            myTC_Number = hullNum & myCol_M & "-" & myTC_Number
            '~~Debug.Print "Len = 8. myTC_Number : " & myTC_Number
            
        ElseIf Len(myTC_Number) = (myLenTC - 2) Then
            'This does not have leading zeros with hull number
            'LPD17BT-DC000101
            '~~Debug.Print "Len = (myLenTC - 2). myTC_Number : " & myTC_Number
            
        ElseIf Len(myTC_Number) = myLenTC Then
            'This is the unique TC number
            'LPD0017BT-DC000101
            'do nothing to it, GTG.
            '~~Debug.Print "Len = myLenTC. GTG myTC_Number : " & myTC_Number
            
        ElseIf Len(myTC_Number) > myLenTC Then
            'This is the FOR OFFICIAL USE ONLY, end of report/export
            Debug.Print ("~~~~~~~~Len > myLenTC. Possible (FOR OFFICIAL USE ONLY) myTC_Number : " & myTC_Number & " ~~~~~~~~")
            Debug.Print ("~~~~~~~~currentFile: " & currentFile)
            Debug.Print ("~~~~~~~~myTC_Number: " & myTC_Number)
            Debug.Print ("~~~~~~~~rowPointer: " & rowPointer)
            Exit For
            
        Else
            'Error in the TC number length
            Debug.Print ("~~~~~~~~Trial Card number Read error~~~~~~~~")
            Debug.Print ("~~~~~~~~currentFile: " & currentFile)
            Debug.Print ("~~~~~~~~myTC_Number: " & myTC_Number)
            Debug.Print ("~~~~~~~~rowPointer: " & rowPointer)
            Exit For
            
        End If
            
        mySTAR = myCol_B 'Col B STAR
        'Convert literal asteric to STAR
        If mySTAR = "*" Then
            mySTAR = "STAR"
        End If
        
        myPRI = myCol_C 'Col C PRI
        mySAFETY = myCol_D 'Col D SAFETY
        mySCREEN = myCol_E 'Col E SCREEN
        
        'Convert literal asteric to AST
        If mySCREEN = "**" Then
            mySCREEN = "AST"
        ElseIf mySCREEN = "AST" Then
            'mySCREEN = "AST"
            'Do nothing
        End If
        
        myAC1 = myCol_F 'Col F ACTION CODE 1
        'Convert literal asteric to AST
        If myAC1 = "**" Then
            myAC1 = "AST"
        ElseIf myAC1 = "****" Then
            myAC1 = "AST"
        ElseIf myAC1 = "AST" Then
            'myAC1 = "AST"
            'Do nothing
        ElseIf myAC1 = "-" Then
            myAC1 = ""
        ElseIf myAC1 = "@" Then
            myAC1 = ""
        End If
        
        myAC2 = myCol_G 'Col G ACTION CODE 2
        'Convert literal asteric to AST
        If myAC2 = "**" Then
            myAC2 = "AST"
        ElseIf myAC2 = "****" Then
            myAC2 = "AST"
        ElseIf myAC2 = "AST" Then
            myAC2 = "AST"
            'Do nothing
        ElseIf myAC2 = "-" Then
            myAC2 = ""
        ElseIf myAC2 = "@" Then
            myAC2 = ""
        End If
        
        mySTATUS = myCol_H 'Col H STATUS
        myActTkn = myCol_I 'Col I ACTION TAKEN
        
        'Need to do a check on the date format to see
        'if it is yyyy-mm-dd or mm/dd/yyyy
        If (InStr(1, myCol_J, "/", 1)) = 0 Then
            If fileReport = "TSM_EXPORT" Then
                myMM_Disc = Mid(myCol_J, 6, 2) 'yyyy-mm-dd Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                myDD_Disc = Mid(myCol_J, 9, 2) 'yyyy-mm-dd Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                myYYYY_Disc = Mid(myCol_J, 1, 4) 'yyyy-mm-dd Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                myDateDisc = myMM_Disc & "/" & myDD_Disc & "/" & myYYYY_Disc 'Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                
                If Len(myCol_K) = 10 Then
                    myMM_Close = Mid(myCol_K, 6, 2) 'yyyy-mm-dd Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                    myDD_Close = Mid(myCol_K, 9, 2) 'yyyy-mm-dd Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                    myYYYY_Close = Mid(myCol_K, 1, 4) 'yyyy-mm-dd Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                    myDateClose = myMM_Close & "/" & myDD_Close & "/" & myYYYY_Close 'Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                    
                Else
                    myDateClose = Empty
                    
                End If
            'May need to do a check on the date format to see
            'if it is yyyy-mm-dd or mm/dd/yyyy
            ElseIf fileReport = "Bean_Data" Then
                myDateDisc = myCol_J 'Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                If Len(myCol_K) = 10 Then
                    myDateClose = myCol_K 'Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                Else
                    myDateClose = Empty
                End If
            End If
        ElseIf (InStr(1, myCol_J, "-", 1)) = 0 Then
            If fileReport = "TSM_EXPORT" Then
                myDateDisc = myCol_J 'Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                If Len(myCol_K) = 10 Then
                    myDateClose = myCol_K 'Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                Else
                    myDateClose = Empty
                End If
            ElseIf fileReport = "Bean_Data" Then
                myDateDisc = myCol_J 'Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
                If Len(myCol_K) = 10 Then
                    myDateClose = myCol_K 'Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
                Else
                    myDateClose = Empty
                End If
            End If
        Else
            Debug.Print ("Date read error. Col J DATE DISCOVERED: " & myCol_J)
            'Untrapped Date read error
            myDateDisc = #2/2/1902#
            myDateClose = #2/2/1902#
        End If
        
        myTrial_ID = myCol_L 'Col L TRIAL ID
        myEVENT = myCol_M 'Col M EVENT
        
        myScrnConCat = mySCREEN & "/" & myAC1 & "/" & myAC2 & "/" & mySTATUS & "/" & myActTkn 'Scrn/AC1/AC2/Stat/AT
        myScrnConCatLong = mySCREEN & "/" & myAC1 & "/" & myAC2 & " " & mySTATUS & "/" & myActTkn 'Scrn/AC1/AC2 Stat/AT
        myScrnConCatShort = mySCREEN & "/" & myAC1 & "/" & myAC2 'Scrn/AC1/AC2
        'myPriConCat = mySTAR & "/" & myPRI & "/" & mySAFETY 'Star/Pri/Safety
        
'Write the results into the database and table
        If dbTarget = "Screencatcher" Then
            'Look for table, check to make sure that tables match availible excel reports
            'Hard coded LPD shipclass reference, for now.
            tblScreencatcher = ScrCtchr_reportDate & "_LPD" & hullNum ' "2016/03/06_LPD26"
            
            tableFinal = False ' this will get set to true when Event Table Final is found
            
            'Need to check for BT, AT, FCT, OWLD, Final Events Tables
            If Table_Exists(tblScreencatcher, strPath_Database) = True Then
                'Table is a non-event table
                'Do nothing to table name, PASS
            ElseIf Table_Exists((tblScreencatcher & "_BT"), strPath_Database) = True Then
                'The Table is the BT Event Table
                'Append "_BT" to table var
                tblScreencatcher = tblScreencatcher & "_BT"
            ElseIf Table_Exists((tblScreencatcher & "_AT"), strPath_Database) = True Then
                'The Table is the AT Event Table
                'Append "_AT" to table var
                tblScreencatcher = tblScreencatcher & "_AT"
            ElseIf Table_Exists((tblScreencatcher & "_FCT"), strPath_Database) = True Then
                'The Table is the FCT Event Table
                'Append "_FCT" to table var
                tblScreencatcher = tblScreencatcher & "_FCT"
            ElseIf Table_Exists((tblScreencatcher & "_OWLD"), strPath_Database) = True Then
                'The Table is the OWLD Event Table
                'Append "_OWLD" to table var
                tblScreencatcher = tblScreencatcher & "_OWLD"
            ElseIf Table_Exists((tblScreencatcher & "_Final"), strPath_Database) = True Then
                'The Table is the Final Event Table
                'Append "_Final" to table var
                tblScreencatcher = tblScreencatcher & "_Final"
                tableFinal = True
                myScrnConCatFinal = mySTATUS & "/" & myActTkn 'Stat/AT
            ElseIf Table_Exists(tblScreencatcher, strPath_Database) = False Then
                'Table was not found
                'Exit
                Debug.Print ("Table not found in Screencatcher Database: " & tblScreencatcher)
            Else
                'The Table was not found
                'Un-trapped error
            End If
            
            'Open the Database
            'This is opened outside of this "For" Loop, leave this connection Open
            'Set dbs_remote = DBEngine.Workspaces(0).OpenDatabase(strPath_Database)
            
            'Open the Recordset
            'This is opened inside of this "For" loop, this will close at the bottom of the "For" loop
            Set rst_remote = dbs_remote.OpenRecordset(tblScreencatcher, dbOpenDynaset)
            
            'Look for TC number, the table only allows Trial Card Number
            mySearch = "[Trial_Card]= '" & myTC_Number & "'"
            'Debug.Print "Screencatcher mySearch : " & mySearch
            rst_remote.FindFirst (mySearch) 'returns a bool .NoMatch property
            
            If rst_remote.NoMatch = True Then
                'write Excel row into Access Table Record Row
                rst_remote.AddNew 'Doesn't exist
                rst_remote.Fields("Trial_Card") = myTC_Number
                rst_remote.Fields("TC_Screening") = myScrnConCatLong 'Scrn/AC1/AC2/Stat/AT
                rst_remote.Fields("TC_Screening_AC1_AC2") = myScrnConCatShort 'Scrn/AC1/AC2
                
                If tableFinal = True Then
                    'Write the final column
                    rst_remote.Fields("Final_Sts_A_T") = myScrnConCatFinal
                Else
                    'Do nothing
                End If
                
                rst_remote.Fields("Star") = mySTAR
                rst_remote.Fields("Pri") = myPRI
                rst_remote.Fields("Safety") = mySAFETY
                rst_remote.Fields("Screening") = mySCREEN
                rst_remote.Fields("Act_1") = myAC1
                rst_remote.Fields("Act_2") = myAC2
                rst_remote.Fields("Status") = mySTATUS
                rst_remote.Fields("Action_Taken") = myActTkn
                rst_remote.Fields("Date_Discovered") = myDateDisc
                rst_remote.Fields("Date_Closed") = myDateClose
                rst_remote.Fields("Trial_ID") = myTrial_ID
                rst_remote.Fields("Event") = myEVENT
                
                rst_remote.Update ' save record to the table
    
            ElseIf rst_remote.NoMatch = False Then
                'assumed that there is a duplicate TC in the report or other report
                Debug.Print ("Duplicate Trial Card Number found: " & myTC_Number)
            Else
                'Do nothing, for now.
                'recordset error with find first
            End If
    
            'Close the Recordset
            rst_remote.Close
            Set rst_remote = Nothing
        
        ElseIf dbTarget = "ScreeningHistory" Then
            'Set the table name
            myTblName = "LPD_" & hullNum & "_SCRN"
            'Open the Recordset
            Set rst = dbs.OpenRecordset(myTblName, dbOpenDynaset)
            'Look for TC number with screening combination
            mySearch = "[TrialCardNumber]= '" & myTC_Number & "' And [ScrningConCat]= '" & myScrnConCat & "'"
            'Debug.Print "ScreeningHistory mySearch : " & mySearch
            rst.FindFirst (mySearch) 'returns a bool .NoMatch property
            
            If rst.NoMatch = True Then
                'write Excel row into Access Table Record Row
                rst.AddNew 'Doesn't exist
                rst.Fields("TrialCardNumber") = myTC_Number
                rst.Fields("ScrningConCat") = myScrnConCat 'Scrn/AC1/AC2/Stat/AT
                'rst.Fields("PriConCat") = myPriConCat 'Star/Pri/Safety
                rst.Fields("Star") = mySTAR
                rst.Fields("Pri") = myPRI
                rst.Fields("Safety") = mySAFETY
                rst.Fields("Scrn") = mySCREEN
                rst.Fields("AC1") = myAC1
                rst.Fields("AC2") = myAC2
                rst.Fields("Status") = mySTATUS
                rst.Fields("ActTkn") = myActTkn
                rst.Fields("DateDisc") = myDateDisc
                rst.Fields("DateClosed") = myDateClose
                rst.Fields("Trial_ID") = myTrial_ID
                rst.Fields("Event") = myEVENT
                rst.Fields("MyReport_date") = reportDate
                
                rst.Update 'This writes back to the recordset/table
    
            Else
                'test for earlist data date
                If reportDate < rst.Fields("MyReport_date") Then
                    '~~Debug.Print "rst.Update: reportDate < rst.Fields(MyReport_date)"
                    rst.Edit 'Exists
                    rst.Fields("MyReport_date") = reportDate
                    
                    rst.Update
                
                ElseIf reportDate >= rst.Fields("MyReport_date") Then
                    recsetDate = rst.Fields("MyReport_date")
                    '~~Debug.Print ">> Do nothing: reportDate[" & reportDate & "] > [" & recsetDate & "]rst.Fields(MyReport_date)"
                    'current report has same or later data by date
                    
                Else
                    'Do nothing
                    
                End If
    
            End If
            
            'Close the Recordset
            rst.Close
            Set rst = Nothing
            
        Else
            'Assumed, No database target was passed
            Debug.Print ("No Target Database was passed in: " & dbTarget & " var value.")
            'for now doing nothing
        End If
        
''''Can be used as current record row number for current action status
        '~~Debug.Print "rowPointer : " & rowPointer
                
    Next rowPointer
    
    '~~Debug.Print "Done reading in Values"
    
    'Close the database objects
    If dbTarget = "Screencatcher" Then
        'Close the Screencatcher Database for a single instance
        dbs_remote.Close
        Set dbs_remote = Nothing
        'Close the ScreeningHistory Database for logging of file actions
        dbs.Close
        Set dbs = Nothing
    Set rst = Nothing
    ElseIf dbTarget = "ScreeningHistory" Then
        'Close the ScreeningHistory Database for logging and screening
        dbs.Close
        Set dbs = Nothing
    End If
    
    'Completed reading in the Excel File, Return the row counter
    OpenXls_BeanReport = rowPointer
    
End Function

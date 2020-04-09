Option Compare Database
Option Explicit

Public Function WriteCurFile_meta(ByVal strPath As String, ByVal myFileName As String, ByVal curClassHull As String, ByVal reportDate As Date, ByVal dbTarget As String) ' As Integer
'This function opens the file name and path of all the passed in file and writes the data to the table
    'strPath: The path to get the current file
    'myFileName: The current file name
    'curClassHull: passed in value from user form
    'reportDate: The date of the report
    'New_xl_to_txt_Name:= newTextName (not used yet, IsNull)
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim rst_inc As DAO.Recordset

    Dim currentFile As String
    Dim className As String '
    Dim hullNum As String '
    Dim mySearch As String
    Dim mySearch_inc As String
    Dim rdc_var As Long
    Dim myRDC_success As String
    
    Dim myTblName As String
    
    
    'Split out Hull Num and Class Name
    className = Left(curClassHull, 3)
    '~~Debug.Print "className : " & className
    hullNum = Right(curClassHull, 2)
    '~~Debug.Print "hullNum : " & hullNum
    
    'myTblName = "MyReportData"
    myTblName = "LPD_" & hullNum & "_MyReportData"
    
    Set dbs = CurrentDb
    'commented out using below with search statement
    'Set rst = dbs.OpenRecordset(myTblName, dbOpenDynaset)
    
    currentFile = strPath & "\" & myFileName
    '~~Debug.Print "currentFile : " & currentFile
    
    rdc_var = 1
        
    Set rst = dbs.OpenRecordset(myTblName, dbOpenDynaset)
    'mySearch = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= '" & reportDate & "' And [ReportDateCount]= '" & rdc_var & "'"
    mySearch = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= #" & reportDate & "# And [ReportDateCount]= " & rdc_var & ""
    '~~Debug.Print "mySearch : " & mySearch
    rst.FindFirst (mySearch)
        
    If rst.NoMatch = True Then
        rst.AddNew 'Doesn't exist
        rst.Fields("ClassHull") = curClassHull
        rst.Fields("ReportDate") = reportDate
        rst.Fields("ReportDateCount") = rdc_var
        rst.Fields("FileName") = myFileName
        rst.Fields("FilePath") = currentFile
        rst.Fields("New_xl_to_txt_Name") = "" 'empty string for now, later will export excel data to text file
        rst.Fields("DateReadIn") = Now()
        rst.Fields("Target_Database") = dbTarget
        rst.Fields("Target_Table") = "" 'empty string for now
        
        rst.Update
        rst.Close

    Else
        'test for next increment ReportDateCount
        myRDC_success = ""
        Do Until ((myRDC_success = "Yes") Or (myRDC_success = "Fail"))
            rdc_var = rdc_var + 1
            '~~Debug.Print "rdc_var : " & rdc_var
            
            Set rst_inc = dbs.OpenRecordset(myTblName, dbOpenDynaset)
            'mySearch = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= '" & reportDate & "' And [ReportDateCount]= '" & rdc_var & "'"
            mySearch_inc = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= #" & reportDate & "# And [ReportDateCount]= " & rdc_var & ""
            '~~Debug.Print "mySearch_inc : " & mySearch_inc
            rst_inc.FindFirst (mySearch_inc)
            
            If rst_inc.NoMatch = True Then
                rst_inc.AddNew 'Doesn't exist
                rst_inc.Fields("ClassHull") = curClassHull
                rst_inc.Fields("ReportDate") = reportDate
                rst_inc.Fields("ReportDateCount") = rdc_var
                rst_inc.Fields("FileName") = myFileName
                rst_inc.Fields("FilePath") = currentFile
                rst_inc.Fields("New_xl_to_txt_Name") = "" 'empty string for now, later will export excel data to text file
                rst_inc.Fields("DateReadIn") = Now()
                rst_inc.Update
                rst_inc.Close
                myRDC_success = "Yes"
                
            ElseIf rst_inc.NoMatch = False Then
                '~~Debug.Print "rst_inc.NoMatch = False"
                rst_inc.Close
                myRDC_success = "No"
                
            Else
                '~~Debug.Print "something went wrong in rst_inc search."
                rst_inc.Close
                myRDC_success = "Fail"
                
            End If
            
        Loop
        
    End If

    '~~Debug.Print "Done adding report data to table"
    
    'make sure to dispose of objects
    Set rst = Nothing
    Set rst_inc = Nothing
    myRDC_success = ""
    
End Function

Public Function CheckCurFile_meta(ByVal strPath As String, ByVal myFileName As String, ByVal curClassHull As String, ByVal reportDate As Date) As Boolean
'This function opens the file name and path of all the passed in file and writes the data to the table
    'strPath: The path to get the current file
    'myFileName: The current file name
    'curClassHull: passed in value from user form
    'reportDate: The date of the report
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset

    Dim currentFile As String
    Dim className As String '
    Dim hullNum As String '
    Dim mySearch As String
    Dim rdc_var As Long
    
    Dim myTblName As String
    
    'Split out Hull Num and Class Name
    className = Left(curClassHull, 3)
    '~~Debug.Print "className : " & className
    hullNum = Right(curClassHull, 2)
    '~~Debug.Print "hullNum : " & hullNum
    
    'myTblName = "MyReportData"
    myTblName = "LPD_" & hullNum & "_MyReportData"
    
    Set dbs = CurrentDb
    'commented out using below with search statement
    'Set rst = dbs.OpenRecordset(myTblName, dbOpenDynaset)
    
    currentFile = strPath & "\" & myFileName
    '~~Debug.Print "currentFile : " & currentFile
    
    rdc_var = 1
        
    Set rst = dbs.OpenRecordset(myTblName, dbOpenDynaset)
    'mySearch = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= '" & reportDate & "' And [ReportDateCount]= '" & rdc_var & "'"
    'mySearch = "[ClassHull]= '" & curClassHull & "' And [ReportDate]= #" & reportDate & "# And [ReportDateCount]= " & rdc_var & ""
    mySearch = "[ClassHull]= '" & curClassHull & "' And [FileName]= '" & myFileName & "'"
    '~~Debug.Print "mySearch : " & mySearch
    rst.FindFirst (mySearch)
    
    If rst.NoMatch = True Then
        'make sure to dispose of objects
        rst.Close
        Set rst = Nothing
        CheckCurFile_meta = False 'This file has not been read in
    
    ElseIf rst.NoMatch = False Then
        'make sure to dispose of objects
        rst.Close
        Set rst = Nothing
        CheckCurFile_meta = True 'This file has been read in skip it
        
    End If

    '~~Debug.Print "Done checking report data in table"
    
    'make sure to dispose of objects
    'rst.Close
    'Set rst = Nothing
    
End Function

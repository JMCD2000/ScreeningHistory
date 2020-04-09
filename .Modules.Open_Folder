Option Compare Database
Option Explicit

Public Function openFolder(curClassHull As String, dbTarget As String)
'This function gets the target folder and steps through all the Files
    '-----------------
    'assertion Var Val
    'Dim curClassHull As String
    'curClassHull = "LPD0025"
    'dbTarget is "Screencatcher" or "ScreeningHistory", This is the target DataBase that gets the data
    '-----------------
    Dim intResult_Folder As Integer 'determines if the user selects a directory from the folder dialog
    Dim intResult_Database As Integer 'determines if the user selects a database from the file dialog
    Dim strPath As String 'the path selected by the user from the folder dialog
    Dim strPath_Database As String 'the database path selected by the user from the file dialog
    Dim objFSO As Object 'Filesystem object
    Dim intCountRows As Integer 'the current number of rows
    
    intCountRows = 0
    
    'Call to the folder picker
    'https://analystcave.com/vba-application-filedialog-select-file/
    'https://docs.microsoft.com/en-us/office/vba/api/overview/library-reference/filedialog-members-office
    Application.FileDialog(4).Title = "Select a data folder path" '4=msoFileDialogFolderPicker
    intResult_Folder = Application.FileDialog(4).Show '4=msoFileDialogFolderPicker
    strPath = Application.FileDialog(4).SelectedItems(1) '4=msoFileDialogFolderPicker
    
    'Call to the Database application picker
    If dbTarget = "f" Then
        Application.FileDialog(3).Title = "Select the database file path" '3=msoFileDialogFilePicker
        intResult_Database = Application.FileDialog(3).Show '3=msoFileDialogFilePicker
        strPath_Database = Application.FileDialog(3).SelectedItems(1) '3=msoFileDialogFolderPicker
        Debug.Print "strPath_Database : " & strPath_Database
    Else
        'Do nothing
        strPath_Database = Empty
        Debug.Print "strPath_Database : " & strPath_Database
    End If
    
    'checks if user has cancled the dialogs
    If (intResult_Folder <> 0 And intResult_Database <> 0) Then
        
        'Folder Path
        Debug.Print "strPath : " & strPath
        
        'Selected ship Class and Hull value
        Debug.Print "curClassHull : " & curClassHull
        
        'Create an instance of the FileSystemObject
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        'Loop through each file in the directory and print the names and path
        ' GetAllFiles() is a private function call
        intCountRows = GetAllFiles(strPath, strPath_Database, objFSO, curClassHull, dbTarget)
        
        'Total number of files read in folder
        Debug.Print "Number of files read : " & intCountRows
        
        'loops through all the files and folder in the input path
        'Call GetAllFolders(strPath, objFSO, intCountRows)
        
    ElseIf (intResult_Folder <> 0 And intResult_Database = 0) Then
        
        'Folder Path
        Debug.Print "strPath : " & strPath
        
        'Selected ship Class and Hull value
        Debug.Print "curClassHull : " & curClassHull
        
        'Create an instance of the FileSystemObject
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        'Loop through each file in the directory and print the names and path
        ' GetAllFiles() is a private function call
        intCountRows = GetAllFiles(strPath, strPath_Database, objFSO, curClassHull, dbTarget)
        
        'Total number of files read in folder
        Debug.Print "Number of files read : " & intCountRows
        
        'loops through all the files and folder in the input path
        'Call GetAllFolders(strPath, objFSO, intCountRows)
    Else:
        'Assumed file picker was closed, or nothing was selected
        'Should have a pop up that tells user nothing selected
        
    End If
    
End Function

Private Function GetAllFiles(ByVal strPath As String, ByVal strPath_Database As String, ByRef objFSO As Object, ByVal curClassHull As String, ByVal dbTarget As String) As Integer
'This function prints the name and path of all the files in the directory strPath
    'strPath: The path to the folder to get the list of files from
    'strPath_Database: The path with fill name to the target DB
    'objFSO: An instance of the FileSystemObject
    'curClassHull: The current ship Class and Hull selected
    Dim inf As Integer
    Dim RDC As Boolean
    Dim rdc_var As Integer 'Report Date Count
    Dim DCT_cs As Integer 'DCount check sum
    Dim mySQL_ReadInsert As String
    Dim curRepRowCnt As Integer
    Dim myCurFile As String
    Dim myCurPath As String
    Dim dcount_var As Integer
    Dim fileExt As String
    Dim openWith As String
    Dim fileReport As String
    
    'Used when creating the text file exported from the Excel File
    Dim newTextName As String
    newTextName = Empty
    
    'Used to limit the reports to the current hull number
    Dim myHullCk As String
    myHullCk = Empty
    
    'Increment the current file operations count
    Dim iCount As Integer
    iCount = 0
    
    'Split out Hull Number
    Dim className As String '
    className = Left(curClassHull, 3)
    '~~Debug.Print "className : " & className
    
    'Split out Class Name
    Dim hullNum As String '
    hullNum = Right(curClassHull, 2)
    '~~Debug.Print "hullNum : " & hullNum
    
    'Set the target report folder file path
    'Dim objFSO As Object
    'Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder As Object '
    Set objFolder = objFSO.GetFolder(strPath)
    
    Dim objFile As Object 'Current object file in Current object folder
    ''''Dim vars used in loop
    Dim skipAction As Boolean 'Used for flagging that current file is not what I am looking for
    Dim myfileNameVar1 As String '
    Dim myfileNameVar2 As String '
    Dim myfileNameVar3 As String '
    Dim myfileNameVar4 As String '
    Dim myfileNameVar5 As String '
    Dim myfileNameVar6 As String '
    Dim myfileNameVar7 As String '
    Dim myfileNameVar8 As String '
    Dim myfileNameVar9 As String '
    Dim reportDate As Date ' Month/Day/Year
    Dim ScrCtchr_reportDate As String ' Year/Month/Day
    Dim myMM As String '
    Dim myDD As String '
    Dim myYYYY As String '
    
   
    For Each objFile In objFolder.Files
        
        'This is used to check if the target file is a Text file or Excel file
        ' It is hardcoded in the myfileNameVar
        'fileExt = objFile.GetExtensionName(objFile)
        'Debug.Print "fileExt : " & fileExt
        
        'All file names need to have MM.DD.YYYY
        reportDate = #1/1/1900# 'each start reset as test value
        skipAction = False
        
        'This is where the Beans have been ordered
        
        myfileNameVar1 = "(??)_LPD" & hullNum & "Bean(DATA)*" 'two digit serial order
        myfileNameVar2 = "(?)_LPD" & hullNum & "Bean(DATA)*" 'one digit serial order
        
        'This is for the beans that are un-ordered
        
        myfileNameVar3 = "LPD" & hullNum & "Bean(DATA)(FCT)*"
        myfileNameVar4 = "LPD" & hullNum & "Bean(DATA)(INSURV)*"
        myfileNameVar5 = "LPD" & hullNum & "Bean(DATA)*" 'Must be last for this name series
        
        'This is for the TSM export file
        
        myfileNameVar6 = "????????_? LPD nu *"
        myfileNameVar7 = "????????_? LPD " & hullNum & " *"
        myfileNameVar8 = "???????? LPD nu *"
        myfileNameVar9 = "???????? LPD " & hullNum & " *"
        
        myCurFile = objFile.Name
        myCurPath = objFile.Path
'        fileExt = objFile.GetExtensionName(objFile)
        '~~Debug.Print "myCurFile : " & myCurFile
        '~~Debug.Print "myCurPath : " & myCurPath
'        '~~Debug.Print "fileExt : " & fileExt
        
        'The file name should be like (00)_LPD17Bean(DATA)MM.DD.YYYY.xlsx
        If myCurFile Like myfileNameVar1 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 21, 2) '(00)_LPD17Bean(DATA)01.15.2008.xlsx
            myDD = Mid(objFile.Name, 24, 2) '(00)_LPD17Bean(DATA)01.15.2008.xlsx
            myYYYY = Mid(objFile.Name, 27, 4) '(00)_LPD17Bean(DATA)01.15.2008.xlsx
            myHullCk = Mid(objFile.Name, 9, 2) '(00)_LPD17Bean(DATA)01.15.2008.xlsx
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var1: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar1 : " & myfileNameVar1
            fileReport = "Bean_Data"
        
        'The file name should be like (0)_LPD17Bean(DATA)MM.DD.YYYY.xlsx
        ElseIf myCurFile Like myfileNameVar2 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 20, 2) '(0)_LPD17Bean(DATA)01.15.2008.xlsx
            myDD = Mid(objFile.Name, 23, 2) '(0)_LPD17Bean(DATA)01.15.2008.xlsx
            myYYYY = Mid(objFile.Name, 26, 4) '(0)_LPD17Bean(DATA)01.15.2008.xlsx
            myHullCk = Mid(objFile.Name, 8, 2) '(0)_LPD17Bean(DATA)01.15.2008.xlsx
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var2: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar2 : " & myfileNameVar2
            fileReport = "Bean_Data"
            
        'The file name should be like LPD17Bean(DATA)(FCT)MM.DD.YYYY.xlsx
        ElseIf myCurFile Like myfileNameVar3 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 21, 2) 'LPD17Bean(DATA)(FCT)01.15.2008.xlsx
            myDD = Mid(objFile.Name, 24, 2) 'LPD17Bean(DATA)(FCT)01.15.2008.xlsx
            myYYYY = Mid(objFile.Name, 27, 4) 'LPD17Bean(DATA)(FCT)01.15.2008.xlsx
            myHullCk = Mid(objFile.Name, 4, 2) 'LPD17Bean(DATA)(FCT)01.15.2008.xlsx
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var3: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar3 : " & myfileNameVar3
            fileReport = "Bean_Data"
            
        'The file name should be like LPD17Bean(DATA)(INSURV)MM.DD.YYYY.xlsx
        ElseIf myCurFile Like myfileNameVar4 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 24, 2) 'LPD17Bean(DATA)(INSURV)01.15.2008.xlsx
            myDD = Mid(objFile.Name, 27, 2) 'LPD17Bean(DATA)(INSURV)01.15.2008.xlsx
            myYYYY = Mid(objFile.Name, 30, 4) 'LPD17Bean(DATA)(INSURV)01.15.2008.xlsx
            myHullCk = Mid(objFile.Name, 4, 2) 'LPD17Bean(DATA)(INSURV)01.15.2008.xlsx
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var4: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar4 : " & myfileNameVar4
            fileReport = "Bean_Data"
            
        'The file name should be like LPD17Bean(DATA)MM.DD.YYYY.xlsx
        ElseIf myCurFile Like myfileNameVar5 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 16, 2) 'LPD17Bean(DATA)01.15.2008.xlsx
            myDD = Mid(objFile.Name, 19, 2) 'LPD17Bean(DATA)01.15.2008.xlsx
            myYYYY = Mid(objFile.Name, 22, 4) 'LPD17Bean(DATA)01.15.2008.xlsx
            myHullCk = Mid(objFile.Name, 4, 2) 'LPD17Bean(DATA)01.15.2008.xlsx
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var5: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar5 : " & myfileNameVar5
            fileReport = "Bean_Data"
        
        'The file name should be like 20161114_1 LPD nu ALL bean bu.xls
        ElseIf myCurFile Like myfileNameVar6 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 5, 2) '20161114_1 LPD nu ALL bean bu.xls
            myDD = Mid(objFile.Name, 7, 2) '20161114_1 LPD nu ALL bean bu.xls
            myYYYY = Mid(objFile.Name, 1, 4) '20161114_1 LPD nu ALL bean bu.xls
            myHullCk = Mid(objFile.Name, 16, 2) '20161114_1 LPD nu ALL bean bu.xls
            '~~Debug.Print "myHullCk : " & myHullCk
            myHullCk = hullNum
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var6: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar6 : " & myfileNameVar6
            fileReport = "TSM_EXPORT"
        
        'The file name should be like 20161114_1 LPD 17 ALL bean bu.xls
        ElseIf myCurFile Like myfileNameVar7 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 5, 2) '20161114_1 LPD 17 ALL bean bu.xls
            myDD = Mid(objFile.Name, 7, 2) '20161114_1 LPD 17 ALL bean bu.xls
            myYYYY = Mid(objFile.Name, 1, 4) '20161114_1 LPD 17 ALL bean bu.xls
            myHullCk = Mid(objFile.Name, 16, 2) '20161114_1 LPD 17 ALL bean bu.xls
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var7: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar7 : " & myfileNameVar7
            fileReport = "TSM_EXPORT"
        
        'The file name should be like 20161114 LPD nu ALL bean bu.xls
        ElseIf myCurFile Like myfileNameVar8 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 5, 2) '20161114 LPD nu ALL bean bu.xls
            myDD = Mid(objFile.Name, 7, 2) '20161114 LPD nu ALL bean bu.xls
            myYYYY = Mid(objFile.Name, 1, 4) '20161114 LPD nu ALL bean bu.xls
            myHullCk = Mid(objFile.Name, 16, 2) '20161114 LPD nu ALL bean bu.xls
            '~~Debug.Print "myHullCk : " & myHullCk
            myHullCk = hullNum
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var8: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar8 : " & myfileNameVar8
            fileReport = "TSM_EXPORT"
        
        'The file name should be like 20161114 LPD 17 ALL bean bu.xls
        ElseIf myCurFile Like myfileNameVar9 Then
            'extract date from report source name
            myMM = Mid(objFile.Name, 5, 2) '20161114 LPD 17 ALL bean bu.xls
            myDD = Mid(objFile.Name, 7, 2) '20161114 LPD 17 ALL bean bu.xls
            myYYYY = Mid(objFile.Name, 1, 4) '20161114 LPD 17 ALL bean bu.xls
            myHullCk = Mid(objFile.Name, 16, 2) '20161114 LPD 17 ALL bean bu.xls
            reportDate = myMM & "/" & myDD & "/" & myYYYY
            ScrCtchr_reportDate = myYYYY & "/" & myMM & "/" & myDD
            openWith = "OpenXls_BeanReport"
            Debug.Print "Var9: LPD " & hullNum & " reportDate : " & reportDate
            '~~Debug.Print "myHullCk : " & myHullCk
            '~~Debug.Print "myfileNameVar9 : " & myfileNameVar9
            fileReport = "TSM_EXPORT"
            
        Else:
            'File doesnt match my file name patern
            'will goto next because reportDate is empty
            skipAction = True
            '~~Debug.Print "Hull " & hullNum & " file name NOT LIKE or NO MATCH with Vars"
            
        End If
        
        'Should reset date place holders
        myMM = Empty
        myDD = Empty
        myYYYY = Empty
        
        If myHullCk <> hullNum Then
            skipAction = True
            '~~Debug.Print "Report Hull number[ " & myHullCk & " ] and Current Action Hull number[ " & hullNum & " ] NOT EQUAL"
        End If
        
        'Check if current file has already been read in. CheckCurFile_meta returns a Boolean
        '~~Debug.Print "<<< skipAction before file name check : " & skipAction
        skipAction = (CheckCurFile_meta(strPath, myCurFile, curClassHull, reportDate))
        '~~Debug.Print ">>>    skipAction after file name check : " & skipAction
        
        If skipAction = False Then
            If (Not (reportDate = #1/1/1900#)) Then
                'Read Sheet object and write into database table
                
                '''Excel file opperation
                If openWith = "OpenXls_BeanReport" Then
                    'Call to Module Read_XLS.OpenXls_BeanReport()
                    curRepRowCnt = (OpenXls_BeanReport(strPath, strPath_Database, objFile.Name, hullNum, curClassHull, reportDate, ScrCtchr_reportDate, fileReport, dbTarget))
                    '~~Debug.Print "Current Report : " & objFile.Name
                    '~~Debug.Print "Current report row count : " & curRepRowCnt
                
                '''Text file opperation
'                ElseIf openWith = "OpenText_BeanReport" Then
'                    curRepRowCnt = (OpenText_BeanReport(strPath, objFile.Name, hullNum, reportDate))
'                    '~~Debug.Print "Current Report : " & objFile.Name
'                    '~~Debug.Print "Current report row count : " & curRepRowCnt
                    
                End If
                
                'Read Sheet object and write in text file
                'newTextName = returned value from function call (not used yet, is empty string)
                
                'Read current file meta and write to table
                WriteCurFile_meta strPath, myCurFile, curClassHull, reportDate, dbTarget
                
            End If
        End If
        
        'Increment the current file operations count
        iCount = iCount + 1
        
    Next objFile
    
    GetAllFiles = iCount
    
End Function

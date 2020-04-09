Option Compare Database
Option Explicit


Public Function myTestTable_Exists()
    'Test function
    'testTableName: The name of the Table to check for
    'strPath_Database: The path to the target DB
    Dim strMyPath As String 'The path to the tested database
    Dim strDBName As String 'The target database
    Dim strDB As String 'The string path and database
    Dim dbs_remote As DAO.Database
    Dim testTableName As String 'This is the Table being checked
    Dim tdf As DAO.TableDef
    Dim myTableIsReal As Boolean
    myTableIsReal = False
    
    'Set my varibles to check for the table
    strMyPath = "C:\Users\jon\Desktop\MyWorkData"
    strDBName = "LPD26 Screencatcher_Analyst.accdb"
    'testTableName = "2016/03/06_LPD26" 'Will not find
    'testTableName = "2016/03/06_LPD26_BT" 'Is first Table
    testTableName = "2016/09/24_LPD26" ' Is in the middle
    
    'Set the string variable to the Database:
    strDB = strMyPath & "\" & strDBName

    'Open the Database
    Set dbs_remote = DBEngine.Workspaces(0).OpenDatabase(strDB)
    
    'Compare test table against DB.TableDefs
    For Each tdf In dbs_remote.TableDefs
        If tdf.Name = testTableName Then
            myTableIsReal = True
            Debug.Print ("myTableIsReal: " & myTableIsReal)
            Exit For ' I found a match
        Else
            myTableIsReal = False
        End If
        Debug.Print ("myTableIsReal: " & myTableIsReal)
        
    Next tdf
    
    'This will cause an error if the table is not found
    'tdf = IsObject(dbs_remote.QueryDefs(testTableName))
    'Debug.Print ("tdf: " & tdf)
    
    dbs_remote.Close
    'Table_Exists = myTableIsReal ' Table was or was not found
     
End Function

Public Function Table_Exists(testTableName As String, ByVal strPath_Database As String) As Boolean
    'This function checks if the table exists in the target database
    'testTableName: The name of the Table to check for
    'strPath_Database: The path to the target DB
    
    Dim strDB As String 'The string path and database
    Dim dbs_remote As DAO.Database
    Dim tdf As DAO.TableDef
    
    'This is the results to be returned
    Dim myTableIsReal As Boolean
    myTableIsReal = False
    
    'Set the string variable to the Database:
    strDB = strPath_Database

    'Open the Database
    Set dbs_remote = DBEngine.Workspaces(0).OpenDatabase(strDB)
    
    'Compare test table against DB.TableDefs
    For Each tdf In dbs_remote.TableDefs
        If tdf.Name = testTableName Then
            myTableIsReal = True
            'Debug.Print ("myTableIsReal: " & myTableIsReal)
            Exit For ' I found a match
        Else
            myTableIsReal = False
        End If
        'Debug.Print ("myTableIsReal: " & myTableIsReal)
        
    Next tdf
    
    'Close the database connection
    dbs_remote.Close
    
    'Return the result
    Table_Exists = myTableIsReal ' Table was or was not found
     
End Function

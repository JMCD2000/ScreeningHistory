Option Compare Database
Option Explicit

Private Sub cmdGetReports_Click()
Dim curClassHull As String 'Current Class Hull target LPD0017
Dim curTargetDataBase As String 'This selects the target data base

curClassHull = Me!cboHullSelect.Column(2) 'This is the selected value from the combo list
Debug.Print "curClassHull : " & curClassHull

curTargetDataBase = Me!cboDataBaseSelect.Column(0) 'This is the selected targeted database value from the combo list
Debug.Print "curTargetDataBase : " & curTargetDataBase

'This is a Module and Function call to Modules.Open_Folder.openFolder()
openFolder curClassHull, curTargetDataBase

End Sub

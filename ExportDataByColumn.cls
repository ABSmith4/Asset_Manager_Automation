VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExtractData()
  
    Dim sourceFile As String
    Dim targetFile As String
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Prompt the user to select the source file
    sourceFile = Application.GetOpenFilename("Excel files (*.xlsx), *.xlsx", , "Select the source file")
    
    ' Exit the sub if the user cancels the file selection
    If sourceFile = "False" Then
        Exit Sub
    End If
    
    ' Prompt the user to select the target file
    targetFile = Application.GetOpenFilename("Excel files (*.xlsx), *.xlsx", , "Select the target file")
    
    
    ' Exit the sub if the user cancels the file selection
    If targetFile = "False" Then
        Exit Sub
    End If
    
    ' Open the source file and set the source sheet and source range
    Workbooks.Open sourceFile
    Set sourceSheet = ActiveSheet
    Set sourceRange = Application.InputBox("Select the source range", "Source Range", Type:=8)
    
    ' Open the target file and set the target sheet and target range
    Workbooks.Open targetFile
    Set targetSheet = ActiveSheet
    Set targetRange = Application.InputBox("Select the target range", "Target Range", Type:=8)
    
    ' Copy the data from the source range to the target range
    sourceRange.Copy targetRange
    
    ' Close the source file
    Workbooks(sourceFile).Close SaveChanges:=False

End Sub



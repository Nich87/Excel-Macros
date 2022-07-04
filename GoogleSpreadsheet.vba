Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

'Googleスプレッドシートをインポートする
Sub Main()
    Dim URL As String
    Dim strFile As String
    URL = "https://docs.google.com/spreadsheets/d/"
    strFile = GetSpreadsheet(strURL)

    Application.ScreenUpdating = False
    Call GetAllSheets(ThisWorkbook, strFile)
    Application.ScreenUpdating = True
    MsgBox "最新のデータに更新が完了しました。"
End Sub

Function GetSpreadsheet(ByVal argURL As String) As String
    Dim outFile As String
    outFile = ThisWorkbook.Path & "\" & "temp.xlsx"

    If argURL Like "*edit?usp=sharing" Then
        argURL = Replace(argURL, "edit?usp=sharing", "")
    End If
    argURL = argURL & "export?format=xlsx"

    Call URLDownloadToFile(0, argURL, outFile, 0, 0)
    GetSpreadsheet = outFile
End Function

Sub GetAllSheets(targetBook As Workbook, ByVal strFile As String)
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Workbooks.Open(Filename:=strFile, ReadOnly:=True)

    For Each ws In wb.Sheets
        ws.UsedRange.Copy (targetBook.Worksheets("Sheet1").Range("A1"))
    Next

    wb.Close SaveChanges:=False
    Kill strFile
End Sub


Private Sub Worksheet_Activate()
Call Main
End Sub

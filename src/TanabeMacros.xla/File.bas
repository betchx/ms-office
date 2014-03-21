Attribute VB_Name = "File"
Option Explicit


' CSV file import into current book
''' Template
Sub CSVインポート()
Attribute CSVインポート.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    
    ' change current directory to current workbook if it has been saved.
    
    If Not IsNull(ActiveWorkbook.Path) And ActiveWorkbook.Path <> "" Then
        Dim filedir
        
        'no need   filedir = Left(ActiveWorkbook.Path, InStrRev(ActiveWorkbook.Path, "\"))
        
        'Sub Sample2()
            With CreateObject("WScript.Shell")
                .CurrentDirectory = ActiveWorkbook.Path ' filedir
            End With
        '    MsgBox CurDir
        'End Sub
        
    End If
    
    
    Dim targets
    targets = Application.GetOpenFilename( _
      FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
      Title:="対象となるcsv fileを選択", MultiSelect:=True)
    
    
    If IsArray(targets) Then
    
      Dim book As Workbook
      Set book = ActiveWorkbook
    
      Dim txtfile
      For Each txtfile In targets
        Workbooks.OpenText txtfile, DataType:=xlDelimited, Comma:=True
        
        Dim Sheet As Worksheet
        Set Sheet = ActiveSheet
        
        Sheet.Move After:=book.Sheets.item(book.Sheets.Count)
      Next txtfile
      
    End If

eee:
  Application.ScreenUpdating = True

End Sub


Sub CSVとしてエクスポート()
Attribute CSVとしてエクスポート.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim sht As Worksheet
'  Dim FSO As New FileSystemObject
  
  Dim dir_name As String
  Dim file_name As String

  Set sht = ActiveSheet
  
  dir_name = ActiveWorkbook.Path  'FSO.GetParentFolderName(ActiveWorkbook.Path)
  file_name = dir_name & "\" & sht.name & ".csv" 'FSO.BuildPath(dir_name, sht.name & ".csv")
  
 ' Set FSO = Nothing

  sht.Copy
  ActiveSheet.SaveAs file_name, xlCSV, AddToMru:=False
  ActiveWorkbook.Close SaveChanges:=False

  sht.Activate


End Sub

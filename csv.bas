Attribute VB_Name = "csv"
Option Compare Database
Option Explicit


Sub main()

    Dim csvPath As String
    csvPath = Application.CurrentProject.Path & "\"
    Const CsvFile = "table.csv"
    
    Call GetDataTableFromCSV(csvPath, CsvFile)

End Sub

'   csvレイアウトを元にcreateSQL作成
Sub GetDataTableFromCSV(csvPath As String, CsvFile As String)
    
    Dim cn As New ADODB.Connection
'    'CSVの接続文字列 Connectionstring
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Text;HDR=Yes;FMT=Delimited"
        .Open csvPath
    End With

    Dim sSQL As String
    sSQL = "SELECT * FROM [" & CsvFile & "]"

    Dim rs1    As New ADODB.Recordset
    rs1.Open sSQL, cn: rs1.MoveFirst
    
    
    '項目毎の最大サイズを取得
    Dim dicItem As Dictionary: Set dicItem = New Dictionary
    Dim idxi As Long, idxj As Long
    
    Do Until rs1.EOF = True
        
        For idxi = 0 To rs1.Fields.Count - 1
        
            If dicItem.Count <= rs1.Fields.Count - 1 Then dicItem.Add rs1(idxi).Name, Len(rs1(idxi))
            
            If dicItem.Exists(rs1(idxi).Name) = False Then
                dicItem.Add rs1(idxi).Name, Len(rs1(idxi))
            Else
            
                If dicItem.Item(rs1(idxi).Name) < Len(rs1(idxi)) Then dicItem.Item(rs1(idxi).Name) = Len(rs1(idxi))
                    
            End If

        Next idxi

        rs1.MoveNext
    Loop

    Dim SQL As String
    SQL = "CREATE TABLE " & "table" & " ("

    Dim arKeys, vKey
    arKeys = dicItem.Keys
    For Each vKey In arKeys
        
        If dicItem.Item(vKey) > 255 Then
            SQL = SQL & vKey & " MEMO,"
        Else
            SQL = SQL & vKey & " TEXT(" & dicItem.Item(vKey) & "),"
        End If
'        Debug.Print vKey
'        Debug.Print "=" & dicItem.Item(vKey)
    Next

    SQL = SQL & ")"
Debug.Print SQL

End Sub




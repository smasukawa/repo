Attribute VB_Name = "mod_DB"
Option Compare Database
Option Explicit

'*******************************************************************************
'   ＣＳＶファイルよりレコードセットを作成する関数
'*******************************************************************************
Public Function Ret_rsCSV(CsvName As String) As ADODB.Recordset

    Dim cn      As New ADODB.Connection
    Dim rs      As New ADODB.Recordset
    Dim strSQL  As String

    With cn
'        ' ヘッダなし
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cmCls.CurrentDir & ";" _
'                            & "Extended Properties='text;HDR=No;FMT=Delimited'"

        ' ヘッダあり
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cmCls.CurrentDir & ";" _
                            & "Extended Properties='text;HDR=YES;FMT=Delimited'"
        .Open
    End With

    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & CsvName & " "

    rs.Open strSQL, cn, adOpenStatic

    Set Ret_rsCSV = rs

End Function

'*******************************************************************************
'   投げられたＳＱＬを元にレコードセットを作成する関数
'*******************************************************************************
Public Function Ret_rsSql(strSQL As String) As ADODB.Recordset

    Dim cn  As ADODB.Connection
    Dim rs As New ADODB.Recordset

    Set cn = New ADODB.Connection
    cn.ConnectionString = CurrentProject.BaseConnectionString
    cn.Open

    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly
    Set Ret_rsSql = rs

End Function

'*******************************************************************************
'ローカルテーブルをセットしたレコードセットを作成する関数（ReadOnly）
'*******************************************************************************
Public Function Ret_rsTableReadOnly(TblName As String) As ADODB.Recordset

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set cn = CurrentProject.Connection

    Set rs = cn.OpenRecordset(TblName, adOpenForwardOnly, adLockReadOnly)
    Set Ret_rsTableReadOnly = rs

End Function

'*******************************************************************************
'   ローカルテーブルをセットしたレコードセットを作成する関数（更新可能）
'*******************************************************************************
Public Function Ret_rsTable(TblName As String) As ADODB.Recordset

    Dim cn  As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set cn = CurrentProject.Connection

    Set rs = New ADODB.Recordset
    rs.Open TblName, cn, adOpenForwardOnly, adLockOptimistic
'    rs.Open TblName, cn, adOpenKeyset, adLockOptimistic
    Set Ret_rsTable = rs

End Function

'*******************************************************************************
'   DoCmd.RunSQLで直接実行
'*******************************************************************************
Public Function RunSQL(strSQL As String) As Boolean

    RunSQL = False

    With DoCmd
        .SetWarnings False
        .RunSQL strSQL
        .SetWarnings True
    End With

    RunSQL = True

End Function

'*******************************************************************************
'   対象テーブルのレコードを削除する
'*******************************************************************************
Public Function DeleteTbl(TblName) As Boolean

    DeleteTbl = False

    RunSQL ("DELETE * FROM " & TblName)

    DeleteTbl = True

End Function

'******************************************************************************
'   レコードセットのクローン作成
'       クローン元のレコードセットがCloseされても連動してCloseされない
'******************************************************************************
Public Function cloneRs(rsSrc As ADODB.Recordset) As ADODB.Recordset

    Dim fld As ADODB.Field

    Set cloneRs = New ADODB.Recordset

    If rsSrc.EOF Then Exit Function

    rsSrc.MoveFirst

    For Each fld In rsSrc.Fields
        cloneRs.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Next

    cloneRs.Open

    Do Until rsSrc.EOF
        cloneRs.AddNew
        For Each fld In rsSrc.Fields
            cloneRs.Fields(fld.Name).Value = fld.Value
        Next
        cloneRs.Update
        rsSrc.MoveNext
    Loop

End Function

'*******************************************************************************
'   リンクテーブルの接続先パスの更新
'   srcTable    ：リンク元のテーブルが格納されたmdbのフルパス名
'*******************************************************************************
Public Function Update_LinkTables(srcTable) As Boolean

    Const procName = "Update_LinkTables"

    Update_LinkTables = False

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

On Error GoTo Err

    Set dbs = CurrentDb

    '-- ローカルの全テーブルを検索
    For Each tdf In dbs.TableDefs
        With tdf
            If .Attributes And dbAttachedTable Then
                '-- リンテーブルを更新
                .Connect = ";DATABASE=" & srcTable & _
                           ";PWD=" & PWD
                .RefreshLink
            End If
        End With
    Next tdf

    Update_LinkTables = True

    Exit Function

Err:
    Call ThrowErr(procName, Err.Number, Err.Description)

End Function

''*******************************************************************************
''   対象テーブルのデータ有無を返す
''    true = あり: False =なし
''*******************************************************************************
'Public Function IsTblDataExist(TblName) As Boolean
'
'    IsTblDataExist = False
'
'On Error GoTo Err
'
'    If DCount("*", cLayout1.ErrTblName) <> 0 Then
'        IsTblDataExist = True
'    End If
'
'Err:
'
'End Function


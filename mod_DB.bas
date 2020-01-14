Attribute VB_Name = "mod_DB"
Option Compare Database
Option Explicit

'*******************************************************************************
'   �b�r�u�t�@�C����背�R�[�h�Z�b�g���쐬����֐�
'*******************************************************************************
Public Function Ret_rsCSV(CsvName As String) As ADODB.Recordset

    Dim cn      As New ADODB.Connection
    Dim rs      As New ADODB.Recordset
    Dim strSQL  As String

    With cn
'        ' �w�b�_�Ȃ�
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cmCls.CurrentDir & ";" _
'                            & "Extended Properties='text;HDR=No;FMT=Delimited'"

        ' �w�b�_����
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
'   ������ꂽ�r�p�k�����Ƀ��R�[�h�Z�b�g���쐬����֐�
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
'���[�J���e�[�u�����Z�b�g�������R�[�h�Z�b�g���쐬����֐��iReadOnly�j
'*******************************************************************************
Public Function Ret_rsTableReadOnly(TblName As String) As ADODB.Recordset

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set cn = CurrentProject.Connection

    Set rs = cn.OpenRecordset(TblName, adOpenForwardOnly, adLockReadOnly)
    Set Ret_rsTableReadOnly = rs

End Function

'*******************************************************************************
'   ���[�J���e�[�u�����Z�b�g�������R�[�h�Z�b�g���쐬����֐��i�X�V�\�j
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
'   DoCmd.RunSQL�Œ��ڎ��s
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
'   �Ώۃe�[�u���̃��R�[�h���폜����
'*******************************************************************************
Public Function DeleteTbl(TblName) As Boolean

    DeleteTbl = False

    RunSQL ("DELETE * FROM " & TblName)

    DeleteTbl = True

End Function

'******************************************************************************
'   ���R�[�h�Z�b�g�̃N���[���쐬
'       �N���[�����̃��R�[�h�Z�b�g��Close����Ă��A������Close����Ȃ�
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
'   �����N�e�[�u���̐ڑ���p�X�̍X�V
'   srcTable    �F�����N���̃e�[�u�����i�[���ꂽmdb�̃t���p�X��
'*******************************************************************************
Public Function Update_LinkTables(srcTable) As Boolean

    Const procName = "Update_LinkTables"

    Update_LinkTables = False

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

On Error GoTo Err

    Set dbs = CurrentDb

    '-- ���[�J���̑S�e�[�u��������
    For Each tdf In dbs.TableDefs
        With tdf
            If .Attributes And dbAttachedTable Then
                '-- �����e�[�u�����X�V
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
''   �Ώۃe�[�u���̃f�[�^�L����Ԃ�
''    true = ����: False =�Ȃ�
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


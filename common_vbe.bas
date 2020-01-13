Attribute VB_Name = "common_vbe"
Option Explicit
'******************************************************************************
'   �����e�i���X���ɗL�p�ȃc�[�����܂Ƃ߂����W���[��
'
'   �ύX�����̔c�����͂Ɏg����ȉ��̋@�\��񋟂��܂�
'
'   ExportAccessObjects �F�J�����gMDB�̃I�u�W�F�N�g���G�N�X�|�[�g����@�\
'   ExportModules       �F�\�[�X�R�[�h���ꊇ�o�͂���@�\
'   ExportModules_List  �FVBA�\�[�X�R�[�h�֐��ꗗ���o�͂���@�\
'   ExportQuery         �F�N�G���Ƃ��ď����ꂽ�r�p�k���o�͂���@�\
'   ExportTableObjects  �F�e�[�u���E���R�[�h�ꗗ���o�͂���@�\
'   PrintReferenceTable �F�Q�Ɛݒ�̈ꗗ���C�~�f�B�G�C�g�ɏo�͂���@�\
'
'   ��ExportModules,ExportModules_List�ɂ��Ă�Access�łȂ��Ă��g���܂��B
'   ���̏ꍇ��MyPath�֐������ɍ��킹�ďC�����Ă�������
'
'   �Q�Ɛݒ�F  Microsoft Visual Basic for Application Extensibility
'               Microsoft Scripting Runtime
'******************************************************************************

'Public Const extCsv             As String = ".csv"  '�g���q�@CSV�t�@�C��

'******************************************************************************
'   �V�X�e�����ɈقȂ�Path�w��̐ؑ֗p
'******************************************************************************
Public Function MyPath() As String
'    MyPath = App.Path               'VB6
'    MyPath = ThisWorkbook.Path      'ExcelVBA
    MyPath = CurrentProject.Path    'AccessVBA
End Function

'******************************************************************************
'   �J�����gMDB�̃I�u�W�F�N�g���G�N�X�|�[�g����@�\
'******************************************************************************
Public Sub ExportAccessObjects()

    Dim curDat  As Object
    Dim curPrj  As Object
    Dim outDir  As String

    Debug.Print "-- �I�u�W�F�N�g �G�N�X�|�[�g�J�n"

    outDir = outputDir & "\Obj"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    Set curDat = Application.CurrentData
    Set curPrj = Application.CurrentProject

    ExportObjectType acQuery, curDat.AllQueries, outDir, ".qry"     '�N�G��
    ExportObjectType acForm, curPrj.AllForms, outDir, ".frm"        '�t�H�[��
    ExportObjectType acReport, curPrj.AllReports, outDir, ".rpt"    '���|�[�g
    ExportObjectType acMacro, curPrj.AllMacros, outDir, ".mcr"      '�}�N��
    ExportObjectType acModule, curPrj.AllModules, outDir, ".bas"    '�W�����W���[������уN���X

    Debug.Print "-- �I�u�W�F�N�g �G�N�X�|�[�g����"

End Sub

'******************************************************************************
'(ExportAccessObject��p�̃T�u���W���[��)
'******************************************************************************
Private Sub ExportObjectType(objType As Integer, _
                             ObjCollection As Variant, _
                             Path As String, _
                             ext As String _
                            )
    Dim Obj As Variant
    Dim filePath As String

    For Each Obj In ObjCollection
        filePath = Path & "\" & Obj.Name & ext
        SaveAsText objType, Obj.Name, filePath
        Debug.Print "Save " & Obj.Name
    Next

End Sub

'******************************************************************************
'  �\�[�X�R�[�h���ꊇ�o�͂���@�\
'******************************************************************************
Public Sub ExportModules()

    Dim vbcComp As VBComponent
    Dim ext     As String
    Dim outDir  As String

    outDir = outputDir & "\src"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    Debug.Print "-- �\�[�X�R�[�h �G�N�X�|�[�g�J�n"
    For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
        Debug.Print vbcComp.Name

        Select Case vbcComp.Type
            Case vbext_ct_StdModule
                ext = ".bas"

            Case vbext_ct_MSForm, vbext_ct_Document
                ext = ".frm"

            Case vbext_ct_ClassModule
                ext = ".cls"
        End Select
        vbcComp.Export (outDir & "\" & vbcComp.Name & ext)
    Next
    Debug.Print "-- �\�[�X�R�[�h �G�N�X�|�[�g����"

End Sub

'******************************************************************************
'  VBA�\�[�X�R�[�h�֐��ꗗ���o�͂���@�\
'******************************************************************************
Public Sub ExportModules_List()

    Dim fso         As New Scripting.FileSystemObject
    Dim ts          As TextStream
    Dim vbcComp     As VBComponent

    Dim connStr         As String   'csv������
    Dim DecFlg          As Boolean  '�錾�Z�N�V�����̔���p�t���O

    Debug.Print "VBA�\�[�X�R�[�h�֐��ꗗ�o�͏���"

    DecFlg = False
    Set ts = fso.CreateTextFile(outputDir & "\" & "ModuleList_" & Format(Now, "yyyymmdd_hhmmss") & ".csv", True)

    connStr = ""
    connStr = connStr & "���W���[����"
    connStr = connStr & ",���W���[�����"
    connStr = connStr & ",�v���V�[�W����"
    connStr = connStr & ",�֐��X�R�[�v"
    connStr = connStr & ",���X�e�b�v��"
    ts.WriteLine connStr

    On Error Resume Next

    '�����̃l�X�g�[�����E�E�ʖڂ����̍�ґ������Ƃ����Ȃ��ƁE�E
    For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
        
        Dim lineCount       As Long     '�R�[�h�̍s�J��グ�J�E���^
        Dim JudgeProcName   As String   '�v���V�[�W�����i����p�j
        Dim ProcName        As String   '�v���V�[�W����()
        Dim BeforeProc      As String   '�v���V�[�W���؂�ւ��^�C�~���O�̔���p
        Dim TotalCount      As Long     '�v���V�[�W�������J�E���^
        
        '// ���W���[���̑S�s
        For lineCount = 1 To vbcComp.CodeModule.CountOfLines
        
            Dim Step            As Long     '���X�e�b�v��
        
            JudgeProcName = vbcComp.CodeModule.ProcOfLine(lineCount, vbext_pk_Proc)
            If JudgeProcName <> Empty And JudgeProcName <> BeforeProc Then
                
                Dim Status          As String   '�֐��X�R�[�v

                '// �錾�Z�N�V�����̔���
                If DecFlg Then
                    ProcName = "Declarations"
                    Status = ""

                    Dim procCount       As Long     '�v���V�[�W���s���J�E���^


                    '�錾�Z�N�V�����s�����̂��������s�����擾
                    For procCount = 1 To vbcComp.CodeModule.CountOfDeclarationLines

                        '// �s���J�E���g
                        GoSub StepCount
                    
                    Next procCount

                    DecFlg = False
                    
                    '// ���W���[����ޔ���
                    GoSub DetermineModuleClasses
                    
                    '// ��������
                    GoSub WriteLine
                    
                End If
                
                ProcName = JudgeProcName

                Dim ProcClasses     As Long     '�v���V�[�W�����
                '�� �v���V�[�W����ނ̈�����\�ߎ擾�ł��Ȃ����߁A����ȓD�L��������������K�v������
                For ProcClasses = 0 To 3
                    
                    Dim ProcString      As String   '�v���V�[�W���錾��������

                    '// �֐��X�R�[�v�̔���
                    ProcString = vbcComp.CodeModule.Lines _
                        (vbcComp.CodeModule.ProcBodyLine(JudgeProcName, ProcClasses), 1)

                    Select Case Err.Number
                        Case 0
                            If InStr(ProcString, "Public") > 0 Then
                                Status = "Public"
                            Else
                                Status = "Private"
                            End If

                            '// �v���V�[�W���s�����̂��������s�����擾
                            For procCount = vbcComp.CodeModule.ProcStartLine(JudgeProcName, ProcClasses) _
                                To vbcComp.CodeModule.ProcStartLine(JudgeProcName, ProcClasses) + _
                                   vbcComp.CodeModule.ProcCountLines(JudgeProcName, ProcClasses)

                                '// �s���J�E���g
                                GoSub StepCount

                            Next procCount

                            '// ���W���[����ޔ���
                            GoSub DetermineModuleClasses

                        Case 35 '�v���V�[�W����ނ̈����Ǝ��ۂ̃v���V�[�W����ނ��s��v�̏ꍇ�A�X���[���đ�������i�߂�
                            Err.Clear
                        Case Else
                            GoTo ErrProc
                    End Select

                Next ProcClasses

                '// ��������
                GoSub WriteLine
                
                Step = 0: TotalCount = TotalCount + 1
            
            End If
            BeforeProc = JudgeProcName
        Next
        DecFlg = True
    Next
    ts.Close: Set ts = Nothing: Set fso = Nothing
    Debug.Print "�o�͊����F" & outputDir: Debug.Print "�v���V�[�W�������F" & TotalCount

    Exit Sub

'// ���W���[����ޔ���
DetermineModuleClasses:
    Dim ModClasses      As String   '���W���[�����
    Select Case vbcComp.Type
        Case vbext_ct_StdModule '�W�����W���[��
            ModClasses = "Module" '
        Case vbext_ct_ClassModule '�N���X���W���[��
            ModClasses = "Class" '
        Case vbext_ct_MSForm, vbext_ct_Document '���[�U�[�t�H�[���A�I�u�W�F�N�g���W���[��
            ModClasses = "Object"
        Case Else
            ModClasses = "-"
    End Select
Return

'// �s���J�E���g
StepCount:
    Dim codeString      As String   '�v���V�[�W���s���̃R�[�h������
    codeString = vbcComp.CodeModule.Lines(procCount, 1)
    
    '�R�����g�Ƌ�s���J�E���g���珜�O����
    If Left(Trim(codeString), 1) <> "'" _
    And Left(StrConv(Trim(codeString), vbUpperCase), 3) <> "REM" _
    And Trim(codeString) <> "" Then
        Step = Step + 1
    End If

Return

'// ��������
WriteLine:
    connStr = ""
    connStr = connStr & vbcComp.CodeModule.Name
    connStr = connStr & "," & ModClasses
    connStr = connStr & "," & ProcName
    connStr = connStr & "," & Status
    connStr = connStr & "," & Step
    ts.WriteLine connStr: Step = 0: ModClasses = ""
Return

ErrProc:
    Debug.Print Err.Number & ":"; Err.Description
End Sub

'******************************************************************************
'   �N�G���Ƃ��ď����ꂽ�r�p�k���o�͂���@�\
'******************************************************************************
Public Sub ExportQuery()

    Dim fso     As New Scripting.FileSystemObject
    Dim ts      As TextStream
    Dim qdf     As DAO.QueryDef
    Dim outDir  As String
    Dim qdfName As String

    Debug.Print "-- �N�G���r�p�k �G�N�X�|�[�g�J�n"

    outDir = outputDir & "\Query"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    For Each qdf In CurrentDb.QueryDefs
        qdfName = Replace(qdf.Name, "/", "�^") '�N�G�����ɑS�p�^�g���Ă�Əo�͏o���Ȃ��̂ŕϊ�����K�v����
        Set ts = fso.CreateTextFile(outDir & "\" & qdfName & ".sql")
        ts.Write qdf.SQL
        ts.Close
        Debug.Print "Save " & qdfName
    Next

    Debug.Print "-- �N�G���r�p�k �G�N�X�|�[�g����"

End Sub

'******************************************************************************
'   �e�[�u���E���R�[�h�ꗗ�o�͏���
'   DAO�ł͏����_�ȉ��������E���܂��ʁE�E
'   �T�C�Y�͕����T�C�Y�Ō����ł͂Ȃ����A���܂�L�����܂���E�E
'   �C���������炻�̂���ADO�x�[�X�ō�蒼������
'******************************************************************************
Public Sub ExportTableObjects()

    Dim fso         As New Scripting.FileSystemObject
    Dim ts          As TextStream
    Dim i           As Long             '���ԗp
    Dim connStr     As String           '�A��������p�o�b�t�@
    
    Dim tdf         As DAO.TableDef
    Dim dbs         As DAO.Database
    Dim fld         As DAO.Field
    
    '��L�[�擾�p
    Dim idxLoop     As DAO.Index        '�C���f�b�N�X�I�u�W�F�N�g����L�[��T��
    Dim idxFld      As Object           '�C���f�b�N�X���ڂ��i�[
    Dim MainKey     As Variant          '��L�[����ɗ��p
    Dim KeyNames    As Collection       '��L�[���̂��i�[
    Dim blnKey      As Boolean          '��L�[����t���O

    Debug.Print "-- �e�[�u���E���R�[�h�ꗗ�o�͏���"

    Set ts = fso.CreateTextFile(outputDir & "\" & "TableObjects_" & _
                                Format(Now, "YYYYMMDD_hhmmss") & ".csv", True)

'    Set ts = fso.CreateTextFile(outputDir & "\" & "TableObjects" & extCsv, True)

    '//�@�b�r�u�w�b�_
    connStr = ""
    connStr = connStr & "�e�[�u����"
    connStr = connStr & ",����"
    connStr = connStr & ",���O"
    connStr = connStr & ",�^"
    connStr = connStr & ",�T�C�Y"
    connStr = connStr & ",��L�["
    connStr = connStr & ",NOT NULL"

    ts.WriteLine connStr

    Set dbs = CurrentDb

    '//�e�[�u����
    For Each tdf In dbs.TableDefs
        '�V�X�e���e�[�u���ȊO���o�͑ΏۂƂ���
        If Left(tdf.Name, 4) <> "MSys" Then

            Set KeyNames = New Collection

            '// ��L�[�̈ꗗ���擾
            For Each idxLoop In tdf.Indexes
                If idxLoop.Primary = True Then
                    For Each idxFld In idxLoop.Fields
                        KeyNames.Add Item:=idxFld.Name
                    Next idxFld
                End If
            Next idxLoop

            i = 1
            
            '//���ږ�
            For Each fld In tdf.Fields
                
                '��L�[�̔���
                For Each MainKey In KeyNames
                    If fld.Name = MainKey Then blnKey = True
                Next MainKey

                '// ��������
                connStr = ""
                connStr = connStr & tdf.Name                '�e�[�u����
                connStr = connStr & "," & i                 '����
                connStr = connStr & "," & fld.Name          '���O
                connStr = connStr & "," & Mid(fld.Name, 3)  '�^
                connStr = connStr & "," & fld.Size          '�T�C�Y
                
                If blnKey Then                              '��L�[
                    connStr = connStr & ",TRUE"
                Else
                    connStr = connStr & ","""""
                End If

                If fld.Required Then                        'NOT NULL
                    connStr = connStr & ",TRUE"
                Else
                    connStr = connStr & ","""""
                End If

                ts.WriteLine connStr

                i = i + 1
                blnKey = False

            Next fld

        End If

    Next tdf

    dbs.Close: Set dbs = Nothing
    Debug.Print "-- �o�͊����F" & outputDir
    Exit Sub

End Sub

'******************************************************************************
'   ��ƃt�H���_����Ԃ��i�Ȃ���������j�֐�
'       �l�b�g���[�N�p�X���Ǝg���Ȃ��E�E�E
'******************************************************************************
Private Function outputDir() As String

    Dim fso     As New Scripting.FileSystemObject
    Dim tp      As Variant
    Dim pp      As Variant
    Dim i       As Variant

    outputDir = MyPath & "\Source\" & Format(date, "yyyymmdd")

    If (fso.FileExists(outputDir)) = False Then
        pp = ""
        tp = Split(outputDir, "\")
        For Each i In tp
            pp = pp & IIf(pp = "", "", "\") & i
            If Not fso.FolderExists(pp & "\") Then
                fso.CreateFolder pp
            End If
        Next i
    End If

    Set fso = Nothing

End Function

'*******************************************************************************
'   �w�肵���e�[�u���̑S���ڂ�Uncod���k��L���ɂ���
'   TblName �F�Ώۂ̃e�[�u����
'       �e�[�u���V�ݎ��ɂ́A�e�[�u����Uncod���k���f�t�H���g�����Ȃ̂�
'       �S���ڂ̈��k���ꊇ�ŗL���ɂ���X�N���v�g������Ă݂�
'
'       �������A���҂����t�@�C���T�C�Y�ጸ���ʂ͑S��������ꂸ�E�E
'       ���̐ݒ���ĈӖ�����́H
'*******************************************************************************
Public Sub compTable(TblName As String)

    Dim voDb    As DAO.Database
    Dim voFld   As DAO.Field
    Dim voPrp   As DAO.Property

On Error GoTo Err:

    Set voDb = CurrentDb

    For Each voFld In voDb.TableDefs(TblName).Fields
        If voFld.Type = dbText Or voFld.Type = dbMemo Then

            Set voPrp = voFld.CreateProperty(Name:="UnicodeCompression", _
                                            Type:=dbBoolean, _
                                            Value:=True)
            voFld.Properties.Append Object:=voPrp
            Set voPrp = Nothing
            Debug.Print "���k�ρF" & voFld.Name
        End If
    Next voFld

    voDb.Close: Set voDb = Nothing

    Exit Sub

Err:
    Select Case Err.Number

    Case 3367 '���Ɉ��k����Ă���ꍇ�̃G���[�Ȃ̂ŃX���[
        Resume Next
    Case Else
        MsgBox Err.Number & ":" & Err.Description
    End Select

End Sub

''*******************************************************************************
''�Q�Ɛݒ�̏��擾
''*******************************************************************************
Public Sub PrintReferenceTable()

    Dim Ref  As Object
    
    For Each Ref In Application.VBE.ActiveVBProject.References
        With Ref
            Debug.Print .Description
    '        Debug.Print .Name
    '        Debug.Print .FullPath
    '        Debug.Print .Guid
    '        Debug.Print .Major
    '        Debug.Print .Minor
    '        Debug.Print .IsBroken
        End With
    Next Ref

End Sub

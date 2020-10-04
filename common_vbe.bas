Attribute VB_Name = "common_vbe"
Option Explicit
'******************************************************************************
'   メンテナンス時に有用なツールをまとめたモジュール
'   ver.20201004
'
'   変更差分の把握や解析に使える以下の機能を提供します
'
'   ExportAccessObjects ：カレントMDBのオブジェクトをエクスポートする機能
'   ExportModules       ：ソースコードを一括出力する機能
'   ExportModules_List  ：VBAソースコード関数一覧を出力する機能
'   ExportQuery         ：クエリとして書かれたＳＱＬを出力する機能
'   ExportTableObjects  ：テーブル・レコード一覧を出力する機能
'   PrintReferenceTable ：参照設定の一覧をイミディエイトに出力する機能
'
'   ※ExportModules,ExportModules_ListについてはAccessでなくても使えます。
'   その場合はMyPath関数を環境に合わせて修正してください
'
'   参照設定：  Microsoft Visual Basic for Application Extensibility
'               Microsoft Scripting Runtime
'******************************************************************************

'Public Const extCsv             As String = ".csv"  '拡張子　CSVファイル

'******************************************************************************
'   システム毎に異なるPath指定の切替用
'******************************************************************************
Public Function MyPath() As String
'    MyPath = App.Path               'VB6
'    MyPath = ThisWorkbook.Path      'ExcelVBA
    MyPath = CurrentProject.Path    'AccessVBA
End Function

'******************************************************************************
'   カレントMDBのオブジェクトをエクスポートする機能
'******************************************************************************
Public Sub ExportAccessObjects()

    Dim curDat  As Object
    Dim curPrj  As Object
    Dim outDir  As String

    Debug.Print "-- オブジェクト エクスポート開始"

    outDir = outputDir & "\Obj"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    Set curDat = Application.CurrentData
    Set curPrj = Application.CurrentProject

    ExportObjectType acQuery, curDat.AllQueries, outDir, ".qry"     'クエリ
    ExportObjectType acForm, curPrj.AllForms, outDir, ".frm"        'フォーム
    ExportObjectType acReport, curPrj.AllReports, outDir, ".rpt"    'レポート
    ExportObjectType acMacro, curPrj.AllMacros, outDir, ".mcr"      'マクロ
    ExportObjectType acModule, curPrj.AllModules, outDir, ".bas"    '標準モジュールおよびクラス

    Debug.Print "-- オブジェクト エクスポート完了"

End Sub

'******************************************************************************
'(ExportAccessObject専用のサブモジュール)
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
'  ソースコードを一括出力する機能
'******************************************************************************
Public Sub ExportModules()

    Dim vbcComp As VBComponent
    Dim ext     As String
    Dim outDir  As String

    outDir = outputDir & "\src"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    Debug.Print "-- ソースコード エクスポート開始"
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
    Debug.Print "-- ソースコード エクスポート完了"

End Sub

'******************************************************************************
'  VBAソースコード関数一覧を出力する機能
'******************************************************************************
Public Sub ExportModules_List()

    Dim fso         As New Scripting.FileSystemObject
    Dim ts          As TextStream
    Dim vbcComp     As VBComponent

    Dim connStr         As String   'csv文字列
    Dim DecFlg          As Boolean  '宣言セクションの判定用フラグ

    Debug.Print "VBAソースコード関数一覧出力処理"

    DecFlg = False
    Set ts = fso.CreateTextFile(outputDir & "\" & "ModuleList_" & Format(Now, "yyyymmdd_hhmmss") & ".csv", True)

    connStr = ""
    connStr = connStr & "モジュール名"
    connStr = connStr & ",モジュール種類"
    connStr = connStr & ",プロシージャ名"
    connStr = connStr & ",関数スコープ"
    connStr = connStr & ",実ステップ数"
    ts.WriteLine connStr

    On Error Resume Next

    'ここのネスト深すぎ・・駄目だこの作者早く何とかしないと・・
    For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
        
        Dim lineCount       As Long     'コードの行繰り上げカウンタ
        Dim JudgeProcName   As String   'プロシージャ名（判定用）
        Dim ProcName        As String   'プロシージャ名()
        Dim BeforeProc      As String   'プロシージャ切り替わりタイミングの判定用
        Dim TotalCount      As Long     'プロシージャ総数カウンタ
        
        '// モジュールの全行
        For lineCount = 1 To vbcComp.CodeModule.CountOfLines
        
            Dim Step            As Long     '実ステップ数
        
            JudgeProcName = vbcComp.CodeModule.ProcOfLine(lineCount, vbext_pk_Proc)
            If JudgeProcName <> Empty And JudgeProcName <> BeforeProc Then
                
                Dim Status          As String   '関数スコープ

                '// 宣言セクションの判定
                If DecFlg Then
                    ProcName = "Declarations"
                    Status = ""

                    Dim procCount       As Long     'プロシージャ行数カウンタ


                    '宣言セクション行総数のうち実効行数を取得
                    For procCount = 1 To vbcComp.CodeModule.CountOfDeclarationLines

                        '// 行数カウント
                        GoSub StepCount
                    
                    Next procCount

                    DecFlg = False
                    
                    '// モジュール種類判定
                    GoSub DetermineModuleClasses
                    
                    '// 書き込み
                    GoSub WriteLine
                    
                End If
                
                ProcName = JudgeProcName

                Dim ProcClasses     As Long     'プロシージャ種類
                '↓ プロシージャ種類の引数を予め取得できないため、こんな泥臭い総当たりをやる必要がある
                For ProcClasses = 0 To 3
                    
                    Dim ProcString      As String   'プロシージャ宣言部文字列

                    '// 関数スコープの判定
                    ProcString = vbcComp.CodeModule.Lines _
                        (vbcComp.CodeModule.ProcBodyLine(JudgeProcName, ProcClasses), 1)

                    Select Case Err.Number
                        Case 0
                            If InStr(ProcString, "Public") > 0 Then
                                Status = "Public"
                            Else
                                Status = "Private"
                            End If

                            '// プロシージャ行総数のうち実効行数を取得
                            For procCount = vbcComp.CodeModule.ProcStartLine(JudgeProcName, ProcClasses) _
                                To vbcComp.CodeModule.ProcStartLine(JudgeProcName, ProcClasses) + _
                                   vbcComp.CodeModule.ProcCountLines(JudgeProcName, ProcClasses)

                                '// 行数カウント
                                GoSub StepCount

                            Next procCount

                            '// モジュール種類判定
                            GoSub DetermineModuleClasses

                        Case 35 'プロシージャ種類の引数と実際のプロシージャ種類が不一致の場合、スルーして総当たり進める
                            Err.Clear
                        Case Else
                            GoTo ErrProc
                    End Select

                Next ProcClasses

                '// 書き込み
                GoSub WriteLine
                
                Step = 0: TotalCount = TotalCount + 1
            
            End If
            BeforeProc = JudgeProcName
        Next
        DecFlg = True
    Next
    ts.Close: Set ts = Nothing: Set fso = Nothing
    Debug.Print "出力完了：" & outputDir: Debug.Print "プロシージャ総数：" & TotalCount

    Exit Sub

'// モジュール種類判定
DetermineModuleClasses:
    Dim ModClasses      As String   'モジュール種類
    Select Case vbcComp.Type
        Case vbext_ct_StdModule '標準モジュール
            ModClasses = "Module" '
        Case vbext_ct_ClassModule 'クラスモジュール
            ModClasses = "Class" '
        Case vbext_ct_MSForm, vbext_ct_Document 'ユーザーフォーム、オブジェクトモジュール
            ModClasses = "Object"
        Case Else
            ModClasses = "-"
    End Select
Return

'// 行数カウント
StepCount:
    Dim codeString      As String   'プロシージャ行内のコード文字列
    codeString = vbcComp.CodeModule.Lines(procCount, 1)
    
    'コメントと空行をカウントから除外する
    If Left(Trim(codeString), 1) <> "'" _
    And Left(StrConv(Trim(codeString), vbUpperCase), 3) <> "REM" _
    And Trim(codeString) <> "" Then
        Step = Step + 1
    End If

Return

'// 書き込み
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
'   クエリとして書かれたＳＱＬを出力する機能
'******************************************************************************
Public Sub ExportQuery()

    Dim fso     As New Scripting.FileSystemObject
    Dim ts      As TextStream
    Dim qdf     As DAO.QueryDef
    Dim outDir  As String
    Dim qdfName As String

    Debug.Print "-- クエリＳＱＬ エクスポート開始"

    outDir = outputDir & "\Query"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    For Each qdf In CurrentDb.QueryDefs
        qdfName = Replace(qdf.Name, "/", "／") 'クエリ名に全角／使ってると出力出来ないので変換する必要あり
        Set ts = fso.CreateTextFile(outDir & "\" & qdfName & ".sql")
        ts.Write qdf.SQL
        ts.Close
        Debug.Print "Save " & qdfName
    Next

    Debug.Print "-- クエリＳＱＬ エクスポート完了"

End Sub

'******************************************************************************
'   テーブル・レコード一覧出力処理
'   DAOでは小数点以下桁数が拾えませぬ・・
'   サイズは物理サイズで桁長ではないし、あまり有難くありません・・
'   気が向いたらそのうちADOベースで作り直すかも
'******************************************************************************
Public Sub ExportTableObjects()

    Dim fso         As New Scripting.FileSystemObject
    Dim ts          As TextStream
    Dim i           As Long             '項番用
    Dim connStr     As String           '連結文字列用バッファ
    
    Dim tdf         As DAO.TableDef
    Dim dbs         As DAO.Database
    Dim fld         As DAO.Field
    
    '主キー取得用
    Dim idxLoop     As DAO.Index        'インデックスオブジェクトより主キーを探す
    Dim idxFld      As Object           'インデックス項目を格納
    Dim MainKey     As Variant          '主キー判定に利用
    Dim KeyNames    As Collection       '主キー名称を格納
    Dim blnKey      As Boolean          '主キー判定フラグ

    Debug.Print "-- テーブル・レコード一覧出力処理"

    Set ts = fso.CreateTextFile(outputDir & "\" & "TableObjects_" & _
                                Format(Now, "YYYYMMDD_hhmmss") & ".csv", True)

'    Set ts = fso.CreateTextFile(outputDir & "\" & "TableObjects" & extCsv, True)

    '//　ＣＳＶヘッダ
    connStr = ""
    connStr = connStr & "テーブル名"
    connStr = connStr & ",項番"
    connStr = connStr & ",名前"
    connStr = connStr & ",型"
    connStr = connStr & ",サイズ"
    connStr = connStr & ",主キー"
    connStr = connStr & ",NOT NULL"

    ts.WriteLine connStr

    Set dbs = CurrentDb

    '//テーブル毎
    For Each tdf In dbs.TableDefs
        'システムテーブル以外を出力対象とする
        If Left(tdf.Name, 4) <> "MSys" Then

            Set KeyNames = New Collection

            '// 主キーの一覧を取得
            For Each idxLoop In tdf.Indexes
                If idxLoop.Primary = True Then
                    For Each idxFld In idxLoop.Fields
                        KeyNames.Add Item:=idxFld.Name
                    Next idxFld
                End If
            Next idxLoop

            i = 1
            
            '//項目毎
            For Each fld In tdf.Fields
                
                '主キーの判定
                For Each MainKey In KeyNames
                    If fld.Name = MainKey Then blnKey = True
                Next MainKey

                '// 書き込み
                connStr = ""
                connStr = connStr & tdf.Name                'テーブル名
                connStr = connStr & "," & i                 '項番
                connStr = connStr & "," & fld.Name          '名前
                connStr = connStr & "," & Mid(fld.Name, 3)  '型
                connStr = connStr & "," & fld.Size          'サイズ
                
                If blnKey Then                              '主キー
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
    Debug.Print "-- 出力完了：" & outputDir
    Exit Sub

End Sub

'******************************************************************************
'   作業フォルダ名を返す（なかったら作る）関数
'******************************************************************************
Private Function outputDir() As String

    Dim fso     As New Scripting.FileSystemObject

    outputDir = MyPath & "\Source\" & Format(Date, "yyyymmdd") & "\" & _
                    fso.GetBaseName(Application.CurrentProject.Name)
    
    If (fso.FileExists(outputDir)) = False Then
    
        Dim tp As Variant, pp As Variant
        pp = ""
        tp = Split(outputDir, "\")
        
        Dim var As Variant
        Dim i   As Long: i = 0
        For Each var In tp
        
            If Left(outputDir, 1) = "\" And i = 3 Then 'ネットワークパス対応
                pp = "\\" & pp & IIf(pp = "", "", "\") & var
            Else
                pp = pp & IIf(pp = "", "", "\") & var
            End If
            
            If Not fso.FolderExists(pp & "\") Then
                fso.CreateFolder pp
            End If
            
            i = i + 1

        Next var
    
    End If

    Set fso = Nothing

End Function
            
'*******************************************************************************
'   指定したテーブルの全項目のUncod圧縮を有効にする
'   TblName ：対象のテーブル名
'       テーブル新設時には、テーブルのUncod圧縮がデフォルト無効なので
'       全項目の圧縮を一括で有効にするスクリプトを作ってみた
'
'       しかし、期待したファイルサイズ低減効果は全く感じられず・・
'       この設定って意味あるの？
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
            Debug.Print "圧縮済：" & voFld.Name
        End If
    Next voFld

    voDb.Close: Set voDb = Nothing

    Exit Sub

Err:
    Select Case Err.Number

    Case 3367 '既に圧縮されている場合のエラーなのでスルー
        Resume Next
    Case Else
        MsgBox Err.Number & ":" & Err.Description
    End Select

End Sub

''*******************************************************************************
''参照設定の情報取得
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

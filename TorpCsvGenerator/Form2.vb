Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

Public Class PatternEditor


#Region "定数"

    'パターン種類コンボ　「新規」名
    Private Const PATTERNCOMBO_NEW = "(新規作成)"

    'プリセット格納パス
    Private Const PRESET_PATH = "\preset"

    'DataGridView カラム名
    Private Const COL_VTIME = "vtime"
    Private Const COL_VDIR = "vdir"
    Private Const COL_VSPD = "vspd"
#End Region

#Region "クラス変数"

    'パターン一覧保存DataSet(画面ロード時に保存し、キャンセルボタン押下時に使用)
    Private _patternDataSetTemporary As DataSet

    'パターン一覧保存DataSet(グリッドを編集する度リアルタイム更新)
    Public _editingPatternDataSet As DataSet

    'パターン一覧変更履歴情報
    Private _patternChangeHistory As Dictionary(Of String, String)

    'プリセット一覧(画面ロード時に取得)
    Private _presetDataSet As DataSet

    '保存ボタン押下フラグ
    Private _isSaved As Boolean

    '選択中パターン名（キーとして利用）
    Private _selectedPatternName As String

#End Region



#Region "静的変数"

    ''' <summary>
    ''' 自身インスタンス
    ''' </summary>
    Private Shared _PatternEditorInstance As PatternEditor

    ''' <summary>
    ''' 自身インスタンス
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property PatternEditorInstance() As PatternEditor
        Get
            Return _PatternEditorInstance
        End Get
        Set(value As PatternEditor)
            _PatternEditorInstance = value
        End Set
    End Property

    '''' <summary>
    '''' パターンDataSet プロパティ
    '''' </summary>
    '''' <returns></returns>
    'Public Property PatternDataSet() As DataSet
    '    Get
    '        Return _patternDataSet
    '    End Get
    '    Set(value As DataSet)
    '        _patternDataSet = value
    '    End Set
    'End Property

    ''' <summary>
    ''' パターン保存実行フラグ　プロパティ
    ''' </summary>
    ''' <returns></returns>
    Public Property IsSaved() As Boolean
        Get
            Return _isSaved
        End Get
        Set(value As Boolean)
            _isSaved = value
        End Set
    End Property

#End Region


#Region "画面起動時"


    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Public Sub New(patternDataSet As DataSet)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _editingPatternDataSet = patternDataSet.Copy()
        _patternDataSetTemporary = patternDataSet.Copy()

    End Sub


    ''' <summary>
    ''' 画面ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PatternEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'タイトル
        Me.Text = TorpCsvGenerator.TITLEBAR_TAILTEXT

        'パターン一覧のロード
        If Not _editingPatternDataSet Is Nothing AndAlso
           Not _editingPatternDataSet.Tables.Count = 0 Then

            'パターン一覧をリストボックスにセット
            LoadPattern()

            '先頭パターンにフォーカス
            ListBox.SetSelected(0, True)

            'グリッド使用可能
            DataGridViewPattern.AllowUserToAddRows = True

        Else
            'パターンが存在しない場合はDataSetから新規作成
            _editingPatternDataSet = New DataSet()

        End If

        'パターン一覧変更履歴情報を生成
        _patternChangeHistory = New Dictionary(Of String, String)
        'パターンDataSetからパターン一覧をリストボックスにセット
        For Each dataTable As DataTable In _editingPatternDataSet.Tables
            _patternChangeHistory.Add(dataTable.TableName, dataTable.TableName)
        Next


        'プリセットファイル読み込み
        LoadPreset()

        'プリセット一覧をコンボボックスにセット
        '先頭は新規作成で固定
        ComboBoxPattern.Items.Add(PATTERNCOMBO_NEW)
        For Each dataTable As DataTable In _presetDataSet.Tables
            ComboBoxPattern.Items.Add(dataTable.TableName)
        Next
        '先頭行を選択状態にする
        ComboBoxPattern.SelectedIndex = 0


    End Sub


    ''' <summary>
    ''' パターン読み込み
    ''' </summary>
    Private Sub LoadPattern()

        'リストを初期化
        ListBox.Items.Clear()

        'パターンDataSetからパターン一覧をリストボックスにセット
        For Each dataTable As DataTable In _editingPatternDataSet.Tables
            ListBox.Items.Add(dataTable.TableName)
        Next

    End Sub


    ''' <summary>
    ''' プリセット読み込み
    ''' </summary>
    Private Sub LoadPreset()

        _presetDataSet = New DataSet()
        Dim csvFiles As String()

        Try
            csvFiles =
                System.IO.Directory.GetFiles(
                System.IO.Directory.GetCurrentDirectory & PRESET_PATH,
                "*.csv",
                System.IO.SearchOption.TopDirectoryOnly)

        Catch ex As DirectoryNotFoundException

            'MessageBox.Show("presetフォルダが見つかりません。")
            Exit Sub

        Catch ex As Exception

            MessageBox.Show("想定外エラーが発生しました:" & ex.Message)
            Exit Sub

        End Try

        For Each file As String In csvFiles

            Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(file)
            Dim dataTable As DataTable = GetPresetDataTable(fileName)
            Dim parser As TextFieldParser = New TextFieldParser(file, System.Text.Encoding.UTF8)
            With parser
                .TextFieldType = FieldType.Delimited
                .SetDelimiters(",")
                .HasFieldsEnclosedInQuotes = False
            End With
            Dim data() As String
            Dim csvRowNumber As Integer = 1

            Try

                While Not parser.EndOfData

                    data = parser.ReadFields()
                    Dim cols As Integer = data.Length

                    If cols <> 3 Then

                        Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                            & "　ファイル名：" & fileName & vbCrLf _
                                            & "　エラー行数：" & csvRowNumber & vbCrLf _
                                            & "　エラー内容：カンマ（,）を2つ記述してください。")

                    ElseIf csvRowNumber = 1 AndAlso
                           Not data(0).Equals("0") Then

                        Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                            & "　ファイル名：" & fileName & vbCrLf _
                                            & "　エラー行数：" & csvRowNumber & vbCrLf _
                                            & "　エラー内容：1行目の時間は0を指定してください。")

                    ElseIf Not Regex.IsMatch(data(0), "^(|0|[1-9]\d{0,3})$") Then

                        Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                            & "　ファイル名：" & fileName & vbCrLf _
                                            & "　エラー行数：" & csvRowNumber & vbCrLf _
                                            & "　エラー内容：1列目（時間）を0～9999の範囲で指定してください。")

                    ElseIf Not Regex.IsMatch(data(1), "^(|0|1)$") Then

                        Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                            & "　ファイル名：" & fileName & vbCrLf _
                                            & "　エラー行数：" & csvRowNumber & vbCrLf _
                                            & "　エラー内容：2列目（方向）を0または1で指定してください。")

                    ElseIf Not Regex.IsMatch(data(2), "^(|0|[1-9]\d{0,1}|100)$") Then

                        Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                            & "　ファイル名：" & fileName & vbCrLf _
                                            & "　エラー行数：" & csvRowNumber & vbCrLf _
                                            & "　エラー内容：3列目（速度）を0～100の範囲で指定してください。")

                    Else

                        Dim row As DataRow = dataTable.NewRow()
                        row(COL_VTIME) = data(0)
                        row(COL_VDIR) = data(1)
                        row(COL_VSPD) = data(2)

                        dataTable.Rows.Add(row)

                    End If
                    csvRowNumber += 1

                End While

                If csvRowNumber <= 2 Then

                    Throw New Exception("presetフォルダ内のcsvファイルを正常に読み込めませんでした。" & vbCrLf _
                                        & "　ファイル名：" & fileName & vbCrLf _
                                        & "　エラー内容：2行以上を記述してください。")

                End If

                _presetDataSet.Tables.Add(dataTable)

            Catch ex As Exception

                MessageBox.Show(ex.Message)

            Finally

                '読み込んだプリセットファイルをクローズ
                parser.Close()
                parser.Dispose()

            End Try

        Next

    End Sub

#End Region


#Region "DataGridView"


    ''' <summary>
    ''' 編集開始時、IMEモードを設定する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewPattern.CellEnter

        Select Case e.ColumnIndex

            Case DataGridViewPattern.Columns(COL_VTIME).Index,
                 DataGridViewPattern.Columns(COL_VDIR).Index,
                 DataGridViewPattern.Columns(COL_VSPD).Index

                DataGridViewPattern.ImeMode = System.Windows.Forms.ImeMode.Disable

        End Select

    End Sub


    ''' <summary>
    ''' DataGridViewへの入力文字制限
    ''' 時間：数字のみ
    ''' 方向：0か1のみ
    ''' 速度：数字のみ
    ''' 
    ''' 参考：Dobon.net「DataGridViewでセルが編集中の時にキーイベントを捕捉する」
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridViewPattern.EditingControlShowing

        'DataGridViewのうち、TextBox型のもののみを対象に入力制限をかける
        If Not TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Exit Sub
        End If

        'DataGridViewオブジェクト取得
        Dim dgv As DataGridView = CType(sender, DataGridView)

        '現在のコントロールを取得
        Dim tb As DataGridViewTextBoxEditingControl =
            CType(e.Control, DataGridViewTextBoxEditingControl)

        'イベントハンドラを削除
        RemoveHandler tb.KeyPress, AddressOf AllowKeyPressDigit
        RemoveHandler tb.KeyPress, AddressOf AllowKeyPressZeroOne

        '該当する列か判定し、ハンドラを追加
        Select Case dgv.CurrentCell.OwningColumn.Name
            Case COL_VTIME, COL_VSPD
                AddHandler tb.KeyPress, AddressOf AllowKeyPressDigit
            Case COL_VDIR
                AddHandler tb.KeyPress, AddressOf AllowKeyPressZeroOne
        End Select

    End Sub


    ''' <summary>
    ''' 数字のみ入力できる制御
    ''' 
    ''' 参考：Dobon.net「DataGridViewでセルが編集中の時にキーイベントを捕捉する」
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AllowKeyPressDigit(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        If (Not Char.IsDigit(e.KeyChar)) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If

    End Sub


    ''' <summary>
    ''' 0か1のみ入力できる制御
    ''' 
    ''' 参考：Dobon.net「DataGridViewでセルが編集中の時にキーイベントを捕捉する」
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AllowKeyPressZeroOne(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        If e.KeyChar <> "0"c AndAlso "1"c <> e.KeyChar AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If

    End Sub


    ''' <summary>
    ''' ビュー行頭番号の設定
    ''' 
    ''' 参考：Dobon.net「DataGridViewの行ヘッダーに行番号を表示する」
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridViewPattern.CellPainting

        If e.ColumnIndex < 0 AndAlso e.RowIndex >= 0 Then
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            TextRenderer.DrawText(e.Graphics,
                                  (e.RowIndex + 1).ToString,
                                  e.CellStyle.Font,
                                  indexRect,
                                  e.CellStyle.ForeColor,
                                  TextFormatFlags.Right Or TextFormatFlags.VerticalCenter
                                  )
            e.Handled = True
        End If

    End Sub


    ''' <summary>
    ''' 編集確定後、入力内容を調整する。
    ''' 時間：先頭0埋めを削除
    ''' 方向：（特になし）
    ''' 速度：先頭0埋めを削除、最大を100とする（101以上は100に変換）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewPattern.CellEndEdit

        '入力内容の調整
        Select Case e.ColumnIndex

            Case DataGridViewPattern.Columns(COL_VTIME).Index

                Dim setValue As String = DataGridViewPattern.Rows(e.RowIndex).Cells(COL_VTIME).Value
                If Not String.IsNullOrEmpty(setValue) Then
                    DataGridViewPattern.Rows(e.RowIndex).Cells(COL_VTIME).Value = CInt(setValue)
                End If

            Case DataGridViewPattern.Columns(COL_VSPD).Index

                Dim setValue As String = DataGridViewPattern.Rows(e.RowIndex).Cells(COL_VSPD).Value
                If Not String.IsNullOrEmpty(setValue) Then
                    If CInt(setValue) > 100 Then
                        DataGridViewPattern.Rows(e.RowIndex).Cells(COL_VSPD).Value = 100
                    Else
                        DataGridViewPattern.Rows(e.RowIndex).Cells(COL_VSPD).Value = CInt(setValue)
                    End If

                End If

        End Select

        '描画を更新
        DrawDataGridViewRow(e.RowIndex)

    End Sub


    ''' <summary>
    ''' DataGridViewの行のアクティブ・非アクティブ・色を設定
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub DrawDataGridViewRow(row As Integer)

        Dim foreColorSpd As Color
        Dim backColorSpd As Color

        Dim spd As Double

        '先頭行（ヘッダ）クリック時は何もしない
        If row = -1 Then
            Exit Sub
        End If

        '速度
        If String.IsNullOrEmpty(DataGridViewPattern.Rows(row).Cells(COL_VSPD).Value) Then

            '未入力の場合
            foreColorSpd = Color.Empty
            backColorSpd = Color.Empty

        Else

            '速度の数値ごとの配色定義
            spd = CDbl(DataGridViewPattern.Rows(row).Cells(COL_VSPD).Value)

            If spd = 0 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HFF, &HFF)

            ElseIf spd < 25 Then
                foreColorSpd = Color.Blue
                backColorSpd = Color.FromArgb(&HFF, &HF0, &HFF)

            ElseIf spd < 40 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HE6, &HFF)

            ElseIf spd < 50 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HDC, &HFF)

            ElseIf spd < 60 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HD2, &HFF)

            ElseIf spd < 70 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HC8, &HFF)

            ElseIf spd < 80 Then
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HBE, &HFF)

            Else
                foreColorSpd = Color.Empty
                backColorSpd = Color.FromArgb(&HFF, &HB4, &HFF)

            End If

        End If

        '速度
        DataGridViewPattern(COL_VSPD, row).Style.ForeColor = foreColorSpd '文字色
        DataGridViewPattern(COL_VSPD, row).Style.BackColor = backColorSpd '背景色

        '1行目の時間は編集不可
        If row = 0 Then
            '時間
            DataGridViewPattern(COL_VTIME, row).ReadOnly = True   '読み込みのみ
            DataGridViewPattern(COL_VTIME, row).Style.ForeColor = Color.Empty '文字色
            DataGridViewPattern(COL_VTIME, row).Style.BackColor = Color.FromArgb(&HE0, &HE0, &HE0) '背景色
            DataGridViewPattern(COL_VTIME, row).Style.SelectionForeColor = Color.Black '選択時文字色
            DataGridViewPattern(COL_VTIME, row).Style.SelectionBackColor = Color.FromArgb(&HE0, &HE0, &HE0) '選択時背景色

        End If

    End Sub


#End Region


#Region "DataGridView以外イベント"

    ''' <summary>
    ''' 追加ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonPatternAdd_Click(sender As Object, e As EventArgs) Handles ButtonPatternAdd.Click

        '既存パターン名一覧を取得
        Dim patternListNames() As String = GetPatternListNames()

        If ComboBoxPattern.SelectedItem.Equals(PATTERNCOMBO_NEW) Then
            '新規作成の場合

            'パターン名編集ダイヤログを表示する
            Dim defaultPatternName As String = String.Empty
            Dim form As New Form4(patternListNames, defaultPatternName, "add")
            Form4.PatternNameFormInstance = form
            form.StartPosition = FormStartPosition.CenterParent
            form.ShowDialog()

            '「OK」が押された場合、新規パターンとしてパターンDataSetに追加
            If form.IsOK = True Then
                Dim dataTable As DataTable = GetPresetDataTable(form.PatternName)

                '1行目の時間を0、方向と速を空白とする。
                Dim row As DataRow = dataTable.NewRow()
                row(COL_VTIME) = "0"
                row(COL_VDIR) = ""
                row(COL_VSPD) = ""
                dataTable.Rows.Add(row)

                _editingPatternDataSet.Tables.Add(dataTable)
            End If

            'フォームオブジェクト削除
            form.Dispose()

        Else
            'プリセット追加の場合

            'パターン名の決定
            Dim patternName As String = ComboBoxPattern.SelectedItem.ToString
            For Each patternListName As String In patternListNames
                If Not String.IsNullOrEmpty(patternListName) AndAlso
                   patternListName.Equals(patternName) Then

                    '重複しないパターン名を取得する
                    Dim j As Integer = 2
                    While True
                        Dim isNameFound As Boolean = True
                        patternName = ComboBoxPattern.SelectedItem.ToString & "_" & j.ToString

                        For Each patternListName_rename As String In patternListNames
                            If patternListName_rename.Equals(patternName) Then
                                isNameFound = False
                                Exit For
                            End If
                        Next
                        If isNameFound = True Then
                            Exit While
                        Else
                            j += 1
                        End If
                    End While

                End If
            Next

            'プリセット内容をコピーし、決定したパターン名でリネーム、パターンDataSetに追加
            Dim dataTable As DataTable = _presetDataSet.Tables(ComboBoxPattern.SelectedItem.ToString).Copy()
            dataTable.TableName = patternName
            _editingPatternDataSet.Tables.Add(dataTable)

        End If

        'パターンリスト再読み込み
        LoadPattern()

    End Sub


    ''' <summary>
    ''' 既存パターン名一覧を取得
    ''' </summary>
    ''' <returns></returns>
    Private Function GetPatternListNames() As String()
        Dim patternListNames(0) As String
        Dim i As Integer = 0
        If Not _editingPatternDataSet Is Nothing Then
            For Each dataTable As DataTable In _editingPatternDataSet.Tables
                ReDim Preserve patternListNames(i)
                patternListNames(i) = dataTable.TableName
                i += 1
            Next
        End If

        Return patternListNames
    End Function


    ''' <summary>
    ''' DataTableの型を取得する
    ''' </summary>
    ''' <param name="dataTableName"></param>
    ''' <returns></returns>
    Private Function GetPresetDataTable(dataTableName As String) As DataTable
        Dim dataTable As DataTable = New DataTable(dataTableName)
        dataTable.Columns.Add(COL_VTIME, Type.GetType("System.String"))
        dataTable.Columns.Add(COL_VDIR, Type.GetType("System.String"))
        dataTable.Columns.Add(COL_VSPD, Type.GetType("System.String"))
        Return dataTable
    End Function


    ''' <summary>
    ''' パターンリストボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox.SelectedIndexChanged

        'リストに対する選択がない場合、処理を終了（初期表示で通過することがある）
        If ListBox.SelectedItem Is Nothing Then
            'グリッド操作不可能
            DataGridViewPattern.AllowUserToAddRows = False
            Exit Sub
        Else
            'グリッド操作可能
            DataGridViewPattern.AllowUserToAddRows = True
        End If

        '前の選択パターンに対して保存を実行
        SavePattern()

        'パターンリストから選択中のパターン名を更新
        _selectedPatternName = ListBox.SelectedItem.ToString
        Dim displayDataTable As DataTable = _editingPatternDataSet.Tables(_selectedPatternName).Copy()

        'DataGridView再読み込み
        DataGridViewPattern.Rows.Clear()
        For Each row As DataRow In displayDataTable.Rows

            Dim newRow As Integer = DataGridViewPattern.Rows.Add()
            DataGridViewPattern.Rows(newRow).Cells(COL_VTIME).Value = row.Item(COL_VTIME).ToString
            DataGridViewPattern.Rows(newRow).Cells(COL_VDIR).Value = row.Item(COL_VDIR).ToString
            DataGridViewPattern.Rows(newRow).Cells(COL_VSPD).Value = row.Item(COL_VSPD).ToString

        Next

        '描画を更新
        For i As Integer = 0 To DataGridViewPattern.Rows.Count - 2
            DrawDataGridViewRow(i)
        Next

    End Sub


    ''' <summary>
    ''' 選択中のパターンを保存する
    ''' </summary>
    Private Sub SavePattern()

        If Not _selectedPatternName Is Nothing Then

            Dim saveDataTable As DataTable = _editingPatternDataSet.Tables(_selectedPatternName)
            saveDataTable.Rows.Clear()
            For rowNum As Integer = 0 To DataGridViewPattern.Rows.Count - 2 '最終行は空白のため、-2　とする

                Dim row As DataRow = saveDataTable.NewRow()
                row(COL_VTIME) = DataGridViewPattern.Rows(rowNum).Cells(COL_VTIME).Value
                row(COL_VDIR) = DataGridViewPattern.Rows(rowNum).Cells(COL_VDIR).Value
                row(COL_VSPD) = DataGridViewPattern.Rows(rowNum).Cells(COL_VSPD).Value
                saveDataTable.Rows.Add(row)

            Next

        End If

    End Sub


    ''' <summary>
    ''' パターン上移動ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonUp_Click(sender As Object, e As EventArgs) Handles ButtonUp.Click
        Dim isUpButton As Boolean = True
        ReplacePattern(isUpButton, ListBox.SelectedIndex)
    End Sub


    ''' <summary>
    ''' パターン下移動ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonDown_Click(sender As Object, e As EventArgs) Handles ButtonDown.Click
        Dim isUpButton As Boolean = False
        ReplacePattern(isUpButton, ListBox.SelectedIndex)
    End Sub


    ''' <summary>
    ''' パターン位置移動
    ''' </summary>
    Private Sub ReplacePattern(ByVal isUpButton As Boolean,
                               ByVal selectedIndex As Integer)

        If selectedIndex = -1 Then
            Exit Sub
        End If

        Dim firstReplaceIndex As Integer

        If isUpButton Then
            firstReplaceIndex = selectedIndex - 1
        Else
            firstReplaceIndex = selectedIndex
        End If

        '一番上のパターンを上に移動、一番下のパターンを下に移動させない
        If firstReplaceIndex < 0 OrElse firstReplaceIndex > _editingPatternDataSet.Tables.Count - 2 Then
            Exit Sub
        End If

        Dim sortedDataSet As New DataSet
        Dim firstReplaceDataTable As New DataTable
        Dim selectedPatternName As String = _editingPatternDataSet.Tables(selectedIndex).TableName

        'パターンDataSetを並べ替え
        For i As Integer = 0 To _editingPatternDataSet.Tables.Count - 1

            Dim selectingDataTable As DataTable = _editingPatternDataSet.Tables(i).Copy

            If firstReplaceIndex = i Then

                '並べ替え対象の1つめを保持
                firstReplaceDataTable = _editingPatternDataSet.Tables(i).Copy

            ElseIf firstReplaceIndex + 1 = i Then

                '並べ替え対象の2つめを格納した後で1つめを格納し、並べ替えを完了する
                sortedDataSet.Tables.Add(selectingDataTable.Copy)
                sortedDataSet.Tables.Add(firstReplaceDataTable.Copy)

            Else

                '並べ替え対象外
                sortedDataSet.Tables.Add(_editingPatternDataSet.Tables(i).Copy)

            End If

        Next

        '並べ替えするDataTableが存在しなかった場合、処理を中断（通らない想定）
        If firstReplaceDataTable.Columns.Count = 0 Then
            Exit Sub
        End If

        '並べ替えしたDataSetでパターン一覧保存DataSetを更新
        _editingPatternDataSet = sortedDataSet.Copy

        'パターンリスト再読み込み
        LoadPattern()

        '並べ替え時に選択していたパターンにフォーカスを当てる
        RemoveHandler ListBox.SelectedIndexChanged, AddressOf ListBox_SelectedIndexChanged
        ListBox.SelectedItem = selectedPatternName
        _selectedPatternName = selectedPatternName
        AddHandler ListBox.SelectedIndexChanged, AddressOf ListBox_SelectedIndexChanged

    End Sub


    ''' <summary>
    ''' パターン名変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonRename_Click(sender As Object, e As EventArgs) Handles ButtonRename.Click

        'リストに対する選択がない場合、処理を終了
        If ListBox.SelectedItem Is Nothing Then
            Exit Sub
        End If

        '選択されているパターンを取得
        Dim renameDataTable As DataTable = _editingPatternDataSet.Tables(ListBox.SelectedItem.ToString)

        '既存パターン名一覧を取得
        Dim patternListNames() As String = GetPatternListNames()

        'パターン名編集ダイヤログを表示する
        Dim defaultPatternName As String = ListBox.SelectedItem.ToString
        Dim form As New Form4(patternListNames, defaultPatternName, "change")
        Form4.PatternNameFormInstance = form
        form.StartPosition = FormStartPosition.CenterParent
        form.ShowDialog()

        '「OK」が押された場合、選択されたパターンをリネーム
        '「キャンセル」が押された場合、処理を終了
        Dim newPatternName As String
        If form.IsOK = True Then
            newPatternName = form.PatternName
            renameDataTable.TableName = newPatternName
        Else
            Exit Sub
        End If

        'フォームオブジェクト削除
        form.Dispose()

        'パターンリスト再読み込み
        LoadPattern()

        'リネーム後のパターンを選択する
        RemoveHandler ListBox.SelectedIndexChanged, AddressOf ListBox_SelectedIndexChanged
        ListBox.SelectedItem = newPatternName
        _selectedPatternName = newPatternName
        AddHandler ListBox.SelectedIndexChanged, AddressOf ListBox_SelectedIndexChanged

        'パターン変更履歴情報を更新
        'Value同士は一意なパターン名を持つはずなので、
        'ValueからKeyを検索し、そのKeyのValueを新しいパターン名で上書きする
        If _patternChangeHistory.ContainsValue(defaultPatternName) Then
            Dim key As String = _patternChangeHistory.FirstOrDefault(Function(k) k.Value = defaultPatternName).Key
            If key Is Nothing Then
                '本来は通らない想定。
                MessageBox.Show("DEBUG:パターン変更履歴の保存に失敗しました。")
            Else
                _patternChangeHistory(key) = newPatternName
            End If
        End If


    End Sub


    ''' <summary>
    ''' ×ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click

        'リストに対する選択がない場合、処理を終了
        If ListBox.SelectedItem Is Nothing Then
            Exit Sub
        End If

        'パターンDataSetから選択されたパターンを削除
        Dim defaultPatternName As String = ListBox.SelectedItem.ToString
        _editingPatternDataSet.Tables.Remove(defaultPatternName)

        '選択中のパターンは存在しなくなる
        _selectedPatternName = Nothing

        'グリッド操作不可能
        DataGridViewPattern.AllowUserToAddRows = False

        'DataGridViewクリア
        DataGridViewPattern.Rows.Clear()

        'パターンリスト再読み込み
        LoadPattern()

        'パターン変更履歴情報を更新
        'Value同士は一意なパターン名を持つはずなので、
        'ValueからKeyを検索し、そのKeyのValueをEmptyとする
        If _patternChangeHistory.ContainsValue(defaultPatternName) Then
            Dim key As String = _patternChangeHistory.FirstOrDefault(Function(k) k.Value = defaultPatternName).Key
            If key Is Nothing Then
                '本来は通らない想定。
                MessageBox.Show("DEBUG:パターン変更履歴の保存に失敗しました。")
            Else
                _patternChangeHistory(key) = String.Empty
            End If
        End If
    End Sub


    ''' <summary>
    ''' キャンセルボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        _isSaved = False
        Me.Close()
    End Sub


    ''' <summary>
    ''' 保存ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click

        '現在の選択パターンの内容を保存
        SavePattern()

        'パターンを保存し、メイン画面へ戻る
        _isSaved = True
        Me.Close()

    End Sub

#End Region


End Class
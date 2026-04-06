
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.Remoting.Contexts
Imports System.Text
Imports System.Xml
Imports OxyPlot
Imports OxyPlot.Series
Imports OxyPlot.Axes


Public Class Scry1CsvGenerator


#Region "定数"

    'アプリケーション名、バージョン
    Public Const TITLEBAR_TAILTEXT = "Scry1CsvGenerator v0.0.1"

    Private Const COL_CHECK = "check"
    Private Const COL_VTIME = "vtime"
    Private Const COL_VTIMEHMS = "vtimehms"
    Private Const COL_VTIMEDIFF = "vtimediff"
    Private Const COL_VDIR = "vdir"
    Private Const COL_VSPD = "vspd"
    Private Const COL_VCMT = "vcmt"
    Private Const COL_PATN1 = "patn1"
    Private Const COL_PATN2 = "patn2"
    Private Const COL_PATN3 = "patn3"
    Private Const COL_PATN4 = "patn4"
    Private Const COL_DSPD1 = "dspd1"
    Private Const COL_DSPD2 = "dspd2"
    Private Const COL_DSPD3 = "dspd3"
    Private Const COL_DSPD4 = "dspd4"
    Private Const COL_RATE1 = "rate1"
    Private Const COL_RATE2 = "rate2"
    Private Const COL_RATE3 = "rate3"
    Private Const COL_RATE4 = "rate4"
    Private Const COL_USERDM = "userdm"
    Private Const COL_USEITV = "useitv"
    Private Const COL_ITVMIN = "itvmin"
    Private Const COL_ITVMAX = "itvmax"

    Private Const XML_TORPCSVGENERATOR = "torpCsvGenerator"
    Private Const XML_INFO = "info"
    Private Const XML_ROWCOUNT = "rowCount"
    Private Const XML_TIMESHEET = "timesheet"
    Private Const XML_TIMESHEET_ROW = "timesheetRow"
    Private Const XML_NO = "no"
    Private Const XML_USEPTN = "useptn"
    Private Const XML_VTIME = "vtime"
    Private Const XML_VDIR = "vdir"
    Private Const XML_VSPD = "vspd"
    Private Const XML_VCMT = "vcmt"
    Private Const XML_PATN1 = "patn1"
    Private Const XML_DSPD1 = "dspd1"
    Private Const XML_RATE1 = "rate1"
    Private Const XML_PATN2 = "patn2"
    Private Const XML_DSPD2 = "dspd2"
    Private Const XML_RATE2 = "rate2"
    Private Const XML_PATN3 = "patn3"
    Private Const XML_DSPD3 = "dspd3"
    Private Const XML_RATE3 = "rate3"
    Private Const XML_PATN4 = "patn4"
    Private Const XML_DSPD4 = "dspd4"
    Private Const XML_RATE4 = "rate4"
    Private Const XML_USERDM = "userdm"
    Private Const XML_USEITV = "useitv"
    Private Const XML_ITVMIN = "itvmin"
    Private Const XML_ITVMAX = "itvmax"
    Private Const XML_PATTERNS = "patterns"
    Private Const XML_PATTERN = "pattern"
    Private Const XML_PATTERN_ROW = "patternRow"
    Private Const XML_PATTERN_NAME = "pattern_name"

    Private Const EXTENSION = "tcga10cycsa"

    Private Const BOOLSTR_TRUE = "1"
    Private Const BOOLSTR_FALSE = "0"

    'ダイヤログパス　デフォルト
    Private Const DEFAULT_PATH = ""

    '時間キャプチャ ディレイ設定 デフォルト
    Private Const CAPTURE_DELAY As Integer = 0

#End Region


#Region "クラス変数"

    '動画再生フォーム
    Private _videoForm As Form3
    Private _fileDialogVideoFileName As String
    Private _fileDialogVideoFilePath As String

    'DataGridView 選択行の一時保存
    Private _selectedRow As Integer

    'パターン一覧
    Private _patternDataSet As DataSet

    Private _IsLoading As Boolean = False

    'トラックバー値
    Private _pattern1_dspd0 As String
    Private _pattern1_dspd1 As String
    Private _pattern1_dspd2 As String
    Private _pattern1_dspd3 As String
    Private _pattern1_dspd4 As String
    Private _pattern1_dspd5 As String
    Private _pattern1_dspd6 As String
    Private _pattern1_dspd7 As String
    Private _pattern1_dspd8 As String
    Private _pattern2_dspd0 As String
    Private _pattern2_dspd1 As String
    Private _pattern2_dspd2 As String
    Private _pattern2_dspd3 As String
    Private _pattern2_dspd4 As String
    Private _pattern2_dspd5 As String
    Private _pattern2_dspd6 As String
    Private _pattern2_dspd7 As String
    Private _pattern2_dspd8 As String
    Private _pattern3_dspd0 As String
    Private _pattern3_dspd1 As String
    Private _pattern3_dspd2 As String
    Private _pattern3_dspd3 As String
    Private _pattern3_dspd4 As String
    Private _pattern3_dspd5 As String
    Private _pattern3_dspd6 As String
    Private _pattern3_dspd7 As String
    Private _pattern3_dspd8 As String
    Private _pattern4_dspd0 As String
    Private _pattern4_dspd1 As String
    Private _pattern4_dspd2 As String
    Private _pattern4_dspd3 As String
    Private _pattern4_dspd4 As String
    Private _pattern4_dspd5 As String
    Private _pattern4_dspd6 As String
    Private _pattern4_dspd7 As String
    Private _pattern4_dspd8 As String

    Private _pattern_rate0 As String
    Private _pattern_rate1 As String
    Private _pattern_rate2 As String
    Private _pattern_rate3 As String
    Private _pattern_rate4 As String

    Private _interval_time00 As String
    Private _interval_time01 As String
    Private _interval_time02 As String
    Private _interval_time03 As String
    Private _interval_time04 As String
    Private _interval_time05 As String
    Private _interval_time06 As String
    Private _interval_time07 As String
    Private _interval_time08 As String
    Private _interval_time09 As String
    Private _interval_time10 As String
    Private _interval_time11 As String
    Private _interval_time12 As String
    Private _interval_time13 As String
    Private _interval_time14 As String
    Private _interval_time15 As String
    Private _interval_time16 As String
    Private _interval_time17 As String
    Private _interval_time18 As String
    Private _interval_time19 As String
    Private _interval_time20 As String

    '時間キャプチャ ディレイ設定
    Private _capture_delay As Integer

    '編集中タイムシート定義ファイル
    Private _editing_fileName As String

    'タイムシート定義ファイル ダイヤログパス
    Private _dialog_path_timesheet_def As String

    'タイムシート出力先　ダイヤログパス
    Private _dialog_path_output As String

    'コンテンツ再生　ダイヤログパス
    Private _dialog_path_video As String
    Private Enum LogLevel
        INFO
        WARN
        ERR
        FATAL
    End Enum

#End Region


#Region "画面起動時"

    ''' <summary>
    ''' 画面起動時動作
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TorpCsvGenerator_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '速度補正のメモリ設定
        _pattern1_dspd0 = "-20"
        _pattern1_dspd1 = "-15"
        _pattern1_dspd2 = "-10"
        _pattern1_dspd3 = "-5"
        _pattern1_dspd4 = "0"
        _pattern1_dspd5 = "+5"
        _pattern1_dspd6 = "+10"
        _pattern1_dspd7 = "+15"
        _pattern1_dspd8 = "+20"

        _pattern2_dspd0 = "-20"
        _pattern2_dspd1 = "-15"
        _pattern2_dspd2 = "-10"
        _pattern2_dspd3 = "-5"
        _pattern2_dspd4 = "0"
        _pattern2_dspd5 = "+5"
        _pattern2_dspd6 = "+10"
        _pattern2_dspd7 = "+15"
        _pattern2_dspd8 = "+20"

        _pattern3_dspd0 = "-20"
        _pattern3_dspd1 = "-15"
        _pattern3_dspd2 = "-10"
        _pattern3_dspd3 = "-5"
        _pattern3_dspd4 = "0"
        _pattern3_dspd5 = "+5"
        _pattern3_dspd6 = "+10"
        _pattern3_dspd7 = "+15"
        _pattern3_dspd8 = "+20"

        _pattern4_dspd0 = "-20"
        _pattern4_dspd1 = "-15"
        _pattern4_dspd2 = "-10"
        _pattern4_dspd3 = "-5"
        _pattern4_dspd4 = "0"
        _pattern4_dspd5 = "+5"
        _pattern4_dspd6 = "+10"
        _pattern4_dspd7 = "+15"
        _pattern4_dspd8 = "+20"

        '実行割合のメモリ設定
        _pattern_rate0 = "0"
        _pattern_rate1 = "1"
        _pattern_rate2 = "2"
        _pattern_rate3 = "3"
        _pattern_rate4 = "4"

        '休止実行割合のメモリ設定
        _interval_time00 = "0"
        _interval_time01 = "0.1"
        _interval_time02 = "0.2"
        _interval_time03 = "0.3"
        _interval_time04 = "0.4"
        _interval_time05 = "0.5"
        _interval_time06 = "0.6"
        _interval_time07 = "0.8"
        _interval_time08 = "1"
        _interval_time09 = "1.5"
        _interval_time10 = "2"
        _interval_time11 = "2.5"
        _interval_time12 = "3"
        _interval_time13 = "4"
        _interval_time14 = "5"
        _interval_time15 = "6"
        _interval_time16 = "8"
        _interval_time17 = "10"
        _interval_time18 = "15"
        _interval_time19 = "20"
        _interval_time20 = "30"

        '画面表示初期化

        'タイトル
        Me.Text = TITLEBAR_TAILTEXT

        '速度補正の数値
        LabelPattern1dspd.Text = _pattern_rate1
        LabelPattern2dspd.Text = _pattern_rate1
        LabelPattern3dspd.Text = _pattern_rate1
        LabelPattern4dspd.Text = _pattern_rate1

        '実行割合の数値
        LabelPattern1rate.Text = _pattern1_dspd4
        LabelPattern2rate.Text = _pattern2_dspd4
        LabelPattern3rate.Text = _pattern3_dspd4
        LabelPattern4rate.Text = _pattern4_dspd4

        '休止の実行確率の数値
        LabelMinTime.Text = _interval_time00
        LabelMaxTime.Text = _interval_time00

        ''ディレイ設定 デフォルト（±0秒）
        '_capture_delay = 0

        '編集中ファイルパス
        _editing_fileName = String.Empty

        'オブジェクト初期化
        _patternDataSet = New DataSet()

        '1行目時間を選択
        DataGridViewTimesheet.CurrentCell =
            DataGridViewTimesheet(COL_VTIME, 0)

        'メニューアクティブ更新
        UpdateMenuActivate()

        '設定ファイル読み込み 定義ファイルディレクトリ
        _dialog_path_timesheet_def = My.Settings.TimesheetDefPath
        If Not Directory.Exists(_dialog_path_timesheet_def) Then
            _dialog_path_timesheet_def = DEFAULT_PATH
        End If

        '設定ファイル読み込み コンテンツディレクトリ
        _dialog_path_video = My.Settings.VideoPath
        If Not Directory.Exists(_dialog_path_video) Then
            _dialog_path_video = DEFAULT_PATH
        End If


        '設定ファイル読み込み ディレイ設定
        _capture_delay = My.Settings.CaptureDelay
        Dim infoTextCaptureDelay As String
        Select Case _capture_delay
            Case 5
                ToolStripMenuItem_plus5s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus5s.Text
            Case 4
                ToolStripMenuItem_plus4s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus4s.Text
            Case 3
                ToolStripMenuItem_plus3s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus3s.Text
            Case 2
                ToolStripMenuItem_plus2s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus2s.Text
            Case 1
                ToolStripMenuItem_plus1s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus1s.Text
            Case 0
                ToolStripMenuItem_plus0s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus0s.Text
            Case -1
                ToolStripMenuItem_minus1s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_minus1s.Text
            Case -2
                ToolStripMenuItem_minus2s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_minus2s.Text
            Case -3
                ToolStripMenuItem_minus3s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_minus3s.Text
            Case -4
                ToolStripMenuItem_minus4s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_minus4s.Text
            Case -5
                ToolStripMenuItem_minus5s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_minus5s.Text
            Case Else
                _capture_delay = CAPTURE_DELAY
                ToolStripMenuItem_plus0s.CheckState = CheckState.Indeterminate
                infoTextCaptureDelay = ToolStripMenuItem_plus0s.Text
        End Select

        TextBoxLog.Text =
            MakeInfoText("時間キャプチャ　ディレイ時間：" & infoTextCaptureDelay, LogLevel.INFO)

        Dim model As PlotModel = New PlotModel
        model.Axes.Add(New LinearAxis With {.Position = AxisPosition.Left, .Minimum = -100, .Maximum = 100})
        model.Axes.Add(New LinearAxis With {.Position = AxisPosition.Bottom})

        OxyPlotView.Model = model
    End Sub


    ''' <summary>
    ''' 閉じる時、設定ファイル保存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TorpCsvGenerator_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        My.Settings.TimesheetDefPath = _dialog_path_timesheet_def
        My.Settings.VideoPath = _dialog_path_video
        My.Settings.CaptureDelay = _capture_delay

    End Sub

#End Region



#Region "メニュー"

    ''' <summary>
    ''' 開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemOpen_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemOpen.Click

        Dim ofd As New OpenFileDialog()
        ofd.FileName = ""
        ofd.InitialDirectory = _dialog_path_timesheet_def
        ofd.Filter = "タイムシート定義ファイル(*." & EXTENSION & ")|*." & EXTENSION
        ofd.Title = "読み込むファイルを選択してください"

        If ofd.ShowDialog() = DialogResult.OK Then

            OpenFile(ofd.FileName)

        End If

    End Sub
    Private Sub OpenFile(ByVal fileName As String)

        If LoadGeneratorFile(fileName) = True Then

            '編集中ファイルパス更新
            _editing_fileName = fileName

            'ダイヤログパス更新
            _dialog_path_timesheet_def = Path.GetDirectoryName(fileName)

            'タイトルバーを更新
            Me.Text = Path.GetFileNameWithoutExtension(_editing_fileName) & " - " & TITLEBAR_TAILTEXT

            'メニューアクティブ更新
            UpdateMenuActivate()

            SetPlot()
        End If

    End Sub

    ''' <summary>
    ''' 開く実行時、タイムシート生成ファイルを読み込む
    ''' </summary>
    ''' <param name="fileName"></param>
    Private Function LoadGeneratorFile(ByVal fileName As String) As Boolean
        Try
            _IsLoading = True

            'xmlファイル読み込み
            Dim xmlDoc As XmlDocument = New XmlDocument
            Dim elmTorpCsvGenerator As XmlElement

            xmlDoc.Load(fileName)
            elmTorpCsvGenerator = xmlDoc.DocumentElement

            'DataGridViewの初期化
            DataGridViewTimesheet.Rows.Clear()
            ClearPatternCombo()

            '--XML読み込み ここから---------------------------------------------------

            '--タイムシート --------------------------------------------------------------------
            Dim nodeTimesheet As XmlNodeList
            nodeTimesheet = elmTorpCsvGenerator.GetElementsByTagName(XML_TIMESHEET)

            For Each elmTimesheet As XmlElement In nodeTimesheet

                Dim nodeTimesheetRow As XmlNodeList
                nodeTimesheetRow = elmTorpCsvGenerator.GetElementsByTagName(XML_TIMESHEET_ROW)

                For i = 0 To nodeTimesheetRow.Count - 1

                    Dim elmRow As XmlElement = CType(nodeTimesheetRow.Item(i), XmlElement)
                    Dim newRow As Integer = DataGridViewTimesheet.Rows.Add()

                    '行のデフォルト値設定
                    DefineRow(DataGridViewTimesheet.Rows(newRow))

                    With DataGridViewTimesheet.Rows(newRow)

                        If CInt(elmRow.Item(XML_USEPTN).InnerText) = BOOLSTR_TRUE Then
                            .Cells(COL_CHECK).Value = True
                        Else
                            .Cells(COL_CHECK).Value = False
                        End If
                        .Cells(COL_VTIME).Value = elmRow.Item(XML_VTIME).InnerText

                        '換算時間を求める
                        If Not String.IsNullOrEmpty(elmRow.Item(XML_VTIME).InnerText) Then
                            InputVtimeHms(newRow)
                        End If

                        '時間差分を求める
                        If Not newRow = 0 Then
                            InputVtimeDiff(newRow - 1)
                        End If

                        .Cells(COL_VDIR).Value = elmRow.Item(XML_VDIR).InnerText
                        .Cells(COL_VSPD).Value = elmRow.Item(XML_VSPD).InnerText
                        .Cells(COL_VCMT).Value = elmRow.Item(XML_VCMT).InnerText

                        If Not elmRow.Item(XML_PATN1) Is Nothing Then
                            .Cells(COL_PATN1).Value = elmRow.Item(XML_PATN1).InnerText
                        End If
                        If Not elmRow.Item(XML_PATN2) Is Nothing Then
                            .Cells(COL_PATN2).Value = elmRow.Item(XML_PATN2).InnerText
                        End If
                        If Not elmRow.Item(XML_PATN3) Is Nothing Then
                            .Cells(COL_PATN3).Value = elmRow.Item(XML_PATN3).InnerText
                        End If
                        If Not elmRow.Item(XML_PATN4) Is Nothing Then
                            .Cells(COL_PATN4).Value = elmRow.Item(XML_PATN4).InnerText
                        End If
                        If Not elmRow.Item(XML_DSPD1) Is Nothing Then
                            .Cells(COL_DSPD1).Value = elmRow.Item(XML_DSPD1).InnerText
                        End If
                        If Not elmRow.Item(XML_DSPD2) Is Nothing Then
                            .Cells(COL_DSPD2).Value = elmRow.Item(XML_DSPD2).InnerText
                        End If
                        If Not elmRow.Item(XML_DSPD3) Is Nothing Then
                            .Cells(COL_DSPD3).Value = elmRow.Item(XML_DSPD3).InnerText
                        End If
                        If Not elmRow.Item(XML_DSPD4) Is Nothing Then
                            .Cells(COL_DSPD4).Value = elmRow.Item(XML_DSPD4).InnerText
                        End If
                        If Not elmRow.Item(XML_RATE1) Is Nothing Then
                            .Cells(COL_RATE1).Value = elmRow.Item(XML_RATE1).InnerText
                        End If
                        If Not elmRow.Item(XML_RATE2) Is Nothing Then
                            .Cells(COL_RATE2).Value = elmRow.Item(XML_RATE2).InnerText
                        End If
                        If Not elmRow.Item(XML_RATE3) Is Nothing Then
                            .Cells(COL_RATE3).Value = elmRow.Item(XML_RATE3).InnerText
                        End If
                        If Not elmRow.Item(XML_RATE4) Is Nothing Then
                            .Cells(COL_RATE4).Value = elmRow.Item(XML_RATE4).InnerText
                        End If
                        If Not elmRow.Item(XML_USERDM) Is Nothing Then
                            If CInt(elmRow.Item(XML_USERDM).InnerText) = BOOLSTR_TRUE Then
                                .Cells(COL_USERDM).Value = True
                            Else
                                .Cells(COL_USERDM).Value = False
                            End If
                        End If
                        If Not elmRow.Item(XML_USEITV) Is Nothing Then
                            If CInt(elmRow.Item(XML_USEITV).InnerText) = BOOLSTR_TRUE Then
                                .Cells(COL_USEITV).Value = True
                            Else
                                .Cells(COL_USEITV).Value = False
                            End If
                        End If
                        If Not elmRow.Item(XML_ITVMIN) Is Nothing Then
                            .Cells(COL_ITVMIN).Value = elmRow.Item(XML_ITVMIN).InnerText
                        End If
                        If Not elmRow.Item(XML_ITVMAX) Is Nothing Then
                            .Cells(COL_ITVMAX).Value = elmRow.Item(XML_ITVMAX).InnerText
                        End If
                    End With

                Next

            Next
            '--タイムシート ここまで------------------------------------------------------------

            '--パターン-------------------------------------------------------------------------
            If Not _patternDataSet Is Nothing AndAlso
           Not _patternDataSet.Tables.Count = 0 Then

                'DataSetが存在する場合は中身をクリアする
                _patternDataSet.Tables.Clear()

            Else
                '存在しない場合は新規作成
                _patternDataSet = New DataSet()

            End If

            Dim nodePatterns As XmlNodeList
            nodePatterns = elmTorpCsvGenerator.GetElementsByTagName(XML_PATTERNS)
            For Each elmPatterns As XmlElement In nodePatterns

                Dim nodePattern As XmlNodeList
                nodePattern = elmPatterns.GetElementsByTagName(XML_PATTERN)

                For Each elmPattern As XmlElement In nodePattern

                    'nodePatternの要素数が0の場合はスキップ
                    If nodePattern.Count = 0 Then
                        Continue For
                    End If

                    Dim tableName As String = elmPattern.Item(XML_PATTERN_NAME).InnerText
                    Dim dataTable As DataTable = GetPresetDataTable(tableName)

                    Dim nodePatternRow As XmlNodeList
                    nodePatternRow = elmPattern.GetElementsByTagName(XML_PATTERN_ROW)

                    For i = 0 To nodePatternRow.Count - 1

                        Dim elmRow As XmlElement = CType(nodePatternRow.Item(i), XmlElement)
                        Dim dataRow As DataRow = dataTable.NewRow()

                        dataRow(COL_VTIME) = elmRow.Item(XML_VTIME).InnerText
                        dataRow(COL_VDIR) = elmRow.Item(XML_VDIR).InnerText
                        dataRow(COL_VSPD) = elmRow.Item(XML_VSPD).InnerText

                        dataTable.Rows.Add(dataRow)

                    Next

                    _patternDataSet.Tables.Add(dataTable)

                Next

            Next
            '--パターン ここまで----------------------------------------------------------------

            '--XML読み込み ここまで---------------------------------------------------

            '行のアクティベイト、描画を更新
            For i As Integer = 0 To DataGridViewTimesheet.Rows.Count - 2
                DrawDataGridViewRow(i)
            Next

            'パターンコンボボックス更新
            LoadPatternToCombo()

            '1行目時間を選択
            DataGridViewTimesheet.CurrentCell =
                DataGridViewTimesheet(COL_VTIME, 0)
            _selectedRow = 0
            RefreshDetail(_selectedRow)

            TextBoxLog.Text = MakeInfoText("タイムシート定義ファイルを開きました。" & "(" & DateTime.Now.ToString("HH:mm:ss") & ")", LogLevel.INFO)

            _IsLoading = False
            Return True

        Catch ex As Exception

            TextBoxLog.Text = MakeInfoText("タイムシート定義ファイル読み込み時に不測エラーが発生しました。", LogLevel.FATAL)

            'DataGridViewの初期化
            DataGridViewTimesheet.Rows.Clear()
            ClearPatternCombo()

            _IsLoading = False
            Return False

        End Try

    End Function


    ''' <summary>
    ''' メニューアクティブ更新
    ''' </summary>
    Private Sub UpdateMenuActivate()

        If Not String.IsNullOrEmpty(_editing_fileName) Then
            Me.ToolStripMenuItemOverwriteSave.Enabled = True
        Else
            Me.ToolStripMenuItemOverwriteSave.Enabled = False
        End If

    End Sub


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
    ''' 名前をつけて保存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemSave_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemSave.Click

        Dim isOverWriteSave As Boolean = False
        Dim isSavedSuccessfully As Boolean = False

        isSavedSuccessfully = SaveFile(isOverWriteSave)

        If isSavedSuccessfully Then

            'タイトルバーを更新
            Me.Text = Path.GetFileNameWithoutExtension(_editing_fileName) & " - " & TITLEBAR_TAILTEXT

            'メニューアクティブ更新
            UpdateMenuActivate()

        End If

    End Sub


    ''' <summary>
    ''' 上書き保存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemOverwriteSave_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemOverwriteSave.Click

        Dim isOverWriteSave As Boolean = True
        Dim isSavedSuccessfully As Boolean = False

        isSavedSuccessfully = SaveFile(isOverWriteSave)

        If isSavedSuccessfully Then
            'タイトルバーを更新
            Me.Text = Path.GetFileNameWithoutExtension(_editing_fileName) & " - " & TITLEBAR_TAILTEXT
        End If

    End Sub


    ''' <summary>
    ''' ファイル保存ダイヤログ表示、XMLファイル保存
    ''' </summary>
    Private Function SaveFile(ByVal isOverWriteSave As Boolean) As Boolean

        Dim isSavedSuccessfully As Boolean = False

        If isOverWriteSave = True AndAlso
           Not String.IsNullOrEmpty(_editing_fileName) AndAlso
           File.Exists(_editing_fileName) Then

            '上書き保存
            isSavedSuccessfully = SaveGeneratorFile(_editing_fileName)

        Else

            '名前をつけて保存
            Dim sfd As New SaveFileDialog()

            sfd.FileName = "新しいファイル." & EXTENSION
            sfd.InitialDirectory = _dialog_path_timesheet_def
            sfd.Filter = "タイムシート定義ファイル(*." & EXTENSION & ")|*." & EXTENSION
            sfd.Title = "保存先を選択してください"

            If sfd.ShowDialog() = DialogResult.OK Then

                '編集中ファイルパス更新
                _editing_fileName = sfd.FileName

                'ダイヤログパス更新
                _dialog_path_timesheet_def = Path.GetDirectoryName(sfd.FileName)

                '保存
                isSavedSuccessfully = SaveGeneratorFile(sfd.FileName)

            End If

        End If

        Return isSavedSuccessfully

    End Function


    ''' <summary>
    ''' 保存、上書き保存を実行時、タイムシート生成ファイルを保存する
    ''' </summary>
    ''' <param name="fileName"></param>
    Private Function SaveGeneratorFile(ByVal fileName As String) As Boolean

        Dim setting As New XmlWriterSettings()
        setting.Indent = True
        setting.IndentChars = "    "
        setting.Encoding = System.Text.Encoding.UTF8
        Dim writer As XmlWriter = Nothing

        Try

            writer = XmlWriter.Create(fileName, setting)

            '--XML書き込み ここから---------------------------------------------------
            writer.WriteStartElement(XML_TORPCSVGENERATOR)

            writer.WriteStartElement(XML_INFO)
            writer.WriteElementString(XML_ROWCOUNT, (DataGridViewTimesheet.Rows.Count - 1).ToString)
            writer.WriteEndElement()

            'タイムシート書き込み
            writer.WriteStartElement(XML_TIMESHEET)
            For i = 0 To DataGridViewTimesheet.Rows.Count - 2

                writer.WriteStartElement(XML_TIMESHEET_ROW)

                writer.WriteStartAttribute(XML_NO)
                writer.WriteString((i + 1).ToString)
                writer.WriteEndAttribute()

                If CType(DataGridViewTimesheet.Rows(i).Cells(COL_CHECK).Value, Boolean) = True Then
                    writer.WriteElementString(XML_USEPTN, BOOLSTR_TRUE)
                Else
                    writer.WriteElementString(XML_USEPTN, BOOLSTR_FALSE)
                End If

                writer.WriteElementString(XML_VTIME, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_VTIME).Value))
                writer.WriteElementString(XML_VDIR, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_VDIR).Value))
                writer.WriteElementString(XML_VSPD, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_VSPD).Value))
                writer.WriteElementString(XML_VCMT, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_VCMT).Value))

                writer.WriteElementString(XML_PATN1, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_PATN1).Value))
                writer.WriteElementString(XML_DSPD1, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_DSPD1).Value))
                writer.WriteElementString(XML_RATE1, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_RATE1).Value))

                writer.WriteElementString(XML_PATN2, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_PATN2).Value))
                writer.WriteElementString(XML_DSPD2, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_DSPD2).Value))
                writer.WriteElementString(XML_RATE2, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_RATE2).Value))

                writer.WriteElementString(XML_PATN3, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_PATN3).Value))
                writer.WriteElementString(XML_DSPD3, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_DSPD3).Value))
                writer.WriteElementString(XML_RATE3, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_RATE3).Value))

                writer.WriteElementString(XML_PATN4, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_PATN4).Value))
                writer.WriteElementString(XML_DSPD4, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_DSPD4).Value))
                writer.WriteElementString(XML_RATE4, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_RATE4).Value))


                If CType(DataGridViewTimesheet.Rows(i).Cells(COL_USERDM).Value, Boolean) = True Then
                    writer.WriteElementString(XML_USERDM, BOOLSTR_TRUE)
                Else
                    writer.WriteElementString(XML_USERDM, BOOLSTR_FALSE)
                End If

                If CType(DataGridViewTimesheet.Rows(i).Cells(COL_USEITV).Value, Boolean) = True Then
                    writer.WriteElementString(XML_USEITV, BOOLSTR_TRUE)
                Else
                    writer.WriteElementString(XML_USEITV, BOOLSTR_FALSE)
                End If

                writer.WriteElementString(XML_ITVMIN, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_ITVMIN).Value))
                writer.WriteElementString(XML_ITVMAX, CStr(DataGridViewTimesheet.Rows(i).Cells(COL_ITVMAX).Value))

                writer.WriteEndElement() 'XML_TIMESHEET

            Next
            writer.WriteEndElement()


            'パターン書き込み
            writer.WriteStartElement(XML_PATTERNS)
            For Each table As DataTable In _patternDataSet.Tables

                writer.WriteStartElement(XML_PATTERN)
                writer.WriteElementString(XML_PATTERN_NAME,
                                          table.TableName)

                For i = 0 To table.Rows.Count - 1

                    Dim row As DataRow = table.Rows(i)

                    writer.WriteStartElement(XML_PATTERN_ROW)

                    writer.WriteStartAttribute(XML_NO)
                    writer.WriteString((i + 1).ToString)
                    writer.WriteEndAttribute()

                    writer.WriteElementString(XML_VTIME, row(XML_VTIME).ToString)
                    writer.WriteElementString(XML_VDIR, row(XML_VDIR).ToString)
                    writer.WriteElementString(XML_VSPD, row(XML_VSPD).ToString)

                    writer.WriteEndElement()

                Next

                writer.WriteEndElement()

            Next
            writer.WriteEndElement()


            writer.WriteEndElement() 'XML_TORPCSVGENERATOR
            '--XML書き込み ここまで---------------------------------------------------

            '終了処理
            writer.Flush()
            writer.Close()

            TextBoxLog.Text = MakeInfoText("保存に成功しました。" & "(" & DateTime.Now.ToString("HH:mm:ss") & ")", LogLevel.INFO)

            Return True

        Catch ex As IOException

            TextBoxLog.Text = MakeInfoText("タイムシート定義ファイルにアクセスできません。" & vbCrLf &
                              "ファイルを閉じてから保存を実行してください。", LogLevel.FATAL)
            Return False

        Catch ex As Exception

            TextBoxLog.Text = MakeInfoText("保存中に不測エラーが発生しました。", LogLevel.FATAL)

            Return False

        End Try

    End Function


    ''' <summary>
    ''' クリアボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemClose_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemClear.Click

        '1行も書かれていない場合は処理しない
        If DataGridViewTimesheet.Rows.Count = 1 Then
            Exit Sub
        End If

        Dim doSave As DialogResult = MessageBox.Show("クリア前に保存しますか？",
                                                     TITLEBAR_TAILTEXT,
                                                     MessageBoxButtons.YesNoCancel,
                                                     MessageBoxIcon.Exclamation,
                                                     MessageBoxDefaultButton.Button1)

        Dim isSavedSuccessfully As Boolean = False

        If doSave = DialogResult.Yes Then

            Dim isOverWriteSave As Boolean = True
            isSavedSuccessfully = SaveFile(isOverWriteSave)

        End If

        If doSave = DialogResult.No OrElse isSavedSuccessfully Then

            '編集中ファイルパスをクリア
            _editing_fileName = String.Empty

            'タイトルバーを更新
            Me.Text = TITLEBAR_TAILTEXT

            'DataGridViewの初期化
            DataGridViewTimesheet.Rows.Clear()
            'パターンオブジェクト初期化
            _patternDataSet = New DataSet()
            'コンボボックス初期化
            LoadPatternToCombo()

            '1行目時間を選択
            DataGridViewTimesheet.CurrentCell =
                DataGridViewTimesheet(COL_VTIME, 0)
            _selectedRow = 0
            RefreshDetail(_selectedRow)

            'メニューアクティブ更新
            UpdateMenuActivate()

            If doSave = DialogResult.No Then

                TextBoxLog.Text = ""

            End If

        End If

    End Sub


    ''' <summary>
    ''' パターンコンボボックスクリア
    ''' </summary>
    Private Sub ClearPatternCombo()

        ComboBoxPattern1.Items.Clear()
        ComboBoxPattern2.Items.Clear()
        ComboBoxPattern3.Items.Clear()
        ComboBoxPattern4.Items.Clear()

    End Sub

    ''' <summary>
    ''' 終了ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemEnd_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemEnd.Click

        Dim result As DialogResult = MessageBox.Show("終了しますか？",
                                                     TITLEBAR_TAILTEXT,
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Information,
                                                     MessageBoxDefaultButton.Button2)

        If result = DialogResult.Yes Then

            Me.Close()

        End If

    End Sub


    ''' <summary>
    ''' コンテンツ再生
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemContentsPlay_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemContentsPlay.Click

        Dim ofd As New OpenFileDialog()
        ofd.FileName = ""
        ofd.InitialDirectory = _dialog_path_video
        ofd.Filter = "動画または音声ファイル|*.mp3;*.wav;*.wma;*.mp4;*.wmv;*.avi;*.mpg;*.mpeg;*.asf;*.m4a;*.aac|すべてのファイル(*.*)|*.*"
        ofd.Title = "動画または音声ファイルを選択してください"

        If ofd.ShowDialog() = DialogResult.OK Then

            UpdateVideoForm(ofd.FileName)

        End If

    End Sub

    Private Sub UpdateVideoForm(ByVal fileName As String)

        'ダイヤログパス更新
        _dialog_path_video = Path.GetDirectoryName(fileName)

        'ビデオ再生フォーム 初回起動時
        If _videoForm Is Nothing Then
            _videoForm = New Form3 With {
                .StartPosition = FormStartPosition.CenterScreen
            }
            _videoForm.Show()

            '時間キャプチャボタンをアクティブにする
            ButtonTimeCapture.Enabled = True
        End If

        '再生フォーム タイトル
        _videoForm.Text = Path.GetFileName(fileName)

        'ビデオ再生フォームをアクティブにし、最前面に出す
        _videoForm.Activate()
        _videoForm.TopMost = True
        _videoForm.TopMost = False

        '再生ファイルを指定
        _videoForm.AxWindowsMediaPlayer1.URL = fileName
        '再生位置は0秒
        _videoForm.AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = 0

        '再生倍率メニュー活性化
        Me.コンテンツ再生速度ToolStripMenuItem.Enabled = True
        Me.X050ToolStripMenuItem.Enabled = True
        Me.X075ToolStripMenuItem.Enabled = True
        Me.X100ToolStripMenuItem.Enabled = True
        Me.X125ToolStripMenuItem.Enabled = True
        Me.X150ToolStripMenuItem.Enabled = True
        Me.X200ToolStripMenuItem.Enabled = True

        '再生倍率x1.0
        ChangeContentsPlayRate(1.0, X100ToolStripMenuItem)

        '戻る・スキップメニュー活性化
        Me.back5sToolStripMenuItem.Enabled = True
        Me.skip5sToolStripMenuItem.Enabled = True

    End Sub


    ''' <summary>
    ''' CSV出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemCsvGenerate_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemCsvGenerate.Click

        If CheckBeforeGenerate() = False Then
            Exit Sub
        End If

        Dim fileName As String
        If Not String.IsNullOrEmpty(_editing_fileName) Then
            fileName = Path.GetFileNameWithoutExtension(_editing_fileName)
        Else
            fileName = "新しいファイル"
        End If

        Dim sfd As New SaveFileDialog()

        sfd.FileName = fileName & ".csv"
        sfd.InitialDirectory = _dialog_path_timesheet_def
        sfd.Filter = "CSVファイル(*.csv)|*.csv"
        sfd.Title = "CSVファイルの出力先を選択してください"

        If sfd.ShowDialog() = DialogResult.OK Then

            'ダイヤログパス更新
            _dialog_path_timesheet_def = Path.GetDirectoryName(sfd.FileName)

            Try

                GenerateTimesheet(sfd.FileName)

            Catch ex As IOException

                TextBoxLog.Text = MakeInfoText("CSVファイルにアクセスできません。" & vbCrLf &
                                  "ファイルを閉じてから出力してください。", LogLevel.FATAL)
            Catch ex As Exception

                TextBoxLog.Text = MakeInfoText("CSV出力時に不測エラーが発生しました。", LogLevel.FATAL)

            End Try

        End If

    End Sub


    ''' <summary>
    ''' タイムシート出力前チェック
    ''' </summary>
    ''' <returns>出力可能成否</returns>
    Private Function CheckBeforeGenerate() As Boolean

        Dim errText As String() = Nothing
        Dim errCount As Integer = 0

        '1行も書かれていない場合
        If DataGridViewTimesheet.Rows.Count = 1 Then
            ReDim Preserve errText(errCount)
            errText(errCount) = MakeInfoText("行がありません。", LogLevel.ERR)
            errCount += 1

        Else

            'パターン名一覧を取得
            Dim patterns As New List(Of String)
            For Each dataTable As DataTable In _patternDataSet.Tables
                patterns.Add(dataTable.TableName)
            Next

            '各行チェック
            For row = 0 To DataGridViewTimesheet.Rows.Count - 2

                Dim time As String = DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value
                Dim lastTime As String
                Dim dir As String
                Dim spd As String
                Dim useptn As Boolean = CBool(DataGridViewTimesheet.Rows(row).Cells(COL_CHECK).Value)
                Dim patn1 As String
                Dim patn2 As String
                Dim patn3 As String
                Dim patn4 As String

                '時間の記述チェック
                If String.IsNullOrEmpty(time) Then
                    ReDim Preserve errText(errCount)
                    errText(errCount) = MakeInfoText("{1}行目：時間が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                    errCount += 1
                Else
                    '時間の増加チェック
                    If row <> 0 Then
                        lastTime = DataGridViewTimesheet.Rows(row - 1).Cells(COL_VTIME).Value
                        If Not String.IsNullOrEmpty(lastTime) AndAlso
                       CInt(time) <= CInt(lastTime) Then
                            ReDim Preserve errText(errCount)
                            errText(errCount) = MakeInfoText("{1}行目：前の行より時間が増加していません。", LogLevel.ERR, (row + 1).ToString())
                            errCount += 1
                        End If
                    End If
                End If

                If useptn = False Then

                    dir = DataGridViewTimesheet.Rows(row).Cells(COL_VDIR).Value
                    spd = DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value

                    '方向の記述チェック
                    If String.IsNullOrEmpty(dir) Then
                        ReDim Preserve errText(errCount)
                        errText(errCount) = MakeInfoText("{1}行目：方向が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                        errCount += 1
                    End If

                    '速度の記述チェック
                    If String.IsNullOrEmpty(spd) Then
                        ReDim Preserve errText(errCount)
                        errText(errCount) = MakeInfoText("{1}行目：速度が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                        errCount += 1
                    End If

                Else

                    patn1 = DataGridViewTimesheet.Rows(row).Cells(COL_PATN1).Value
                    patn2 = DataGridViewTimesheet.Rows(row).Cells(COL_PATN2).Value
                    patn3 = DataGridViewTimesheet.Rows(row).Cells(COL_PATN3).Value
                    patn4 = DataGridViewTimesheet.Rows(row).Cells(COL_PATN4).Value

                    'パターンの記述チェック
                    If String.IsNullOrEmpty(patn1) AndAlso
                   String.IsNullOrEmpty(patn2) AndAlso
                   String.IsNullOrEmpty(patn3) AndAlso
                   String.IsNullOrEmpty(patn4) Then
                        ReDim Preserve errText(errCount)
                        errText(errCount) = MakeInfoText("{1}行目：パターンが1つも選択されていません。", LogLevel.ERR, (row + 1).ToString())
                        errCount += 1
                    Else

                        Dim dummyCnt As Integer = 0
                        If Not String.IsNullOrEmpty(patn1) Then
                            dummyCnt += CInt(DataGridViewTimesheet.Rows(row).Cells(COL_RATE1).Value)

                            '指定されたパターン名の存在チェック
                            If patterns.IndexOf(patn1) = -1 Then
                                ReDim Preserve errText(errCount)
                                errText(errCount) = MakeInfoText("{1}行目：パターン""{2}""は作成されていません。", LogLevel.ERR, (row + 1).ToString(), patn1)
                                errCount += 1
                            End If

                        End If
                        If Not String.IsNullOrEmpty(patn2) Then
                            dummyCnt += CInt(DataGridViewTimesheet.Rows(row).Cells(COL_RATE2).Value)

                            '指定されたパターン名の存在チェック
                            If patterns.IndexOf(patn2) = -1 Then
                                ReDim Preserve errText(errCount)
                                errText(errCount) = MakeInfoText("{1}行目：パターン""{2}""は作成されていません。", LogLevel.ERR, (row + 1).ToString(), patn2)
                                errCount += 1
                            End If

                        End If
                        If Not String.IsNullOrEmpty(patn3) Then
                            dummyCnt += CInt(DataGridViewTimesheet.Rows(row).Cells(COL_RATE3).Value)

                            '指定されたパターン名の存在チェック
                            If patterns.IndexOf(patn3) = -1 Then
                                ReDim Preserve errText(errCount)
                                errText(errCount) = MakeInfoText("{1}行目：パターン""{2}""は作成されていません。", LogLevel.ERR, (row + 1).ToString(), patn3)
                                errCount += 1
                            End If

                        End If
                        If Not String.IsNullOrEmpty(patn4) Then
                            dummyCnt += CInt(DataGridViewTimesheet.Rows(row).Cells(COL_RATE4).Value)

                            '指定されたパターン名の存在チェック
                            If patterns.IndexOf(patn4) = -1 Then
                                ReDim Preserve errText(errCount)
                                errText(errCount) = MakeInfoText("{1}行目：パターン""{2}""は作成されていません。", LogLevel.ERR, (row + 1).ToString(), patn4)
                                errCount += 1
                            End If

                        End If

                        '実行比率チェック（パターンが選択されたもののみチェックする）
                        If dummyCnt = 0 Then
                            ReDim Preserve errText(errCount)
                            errText(errCount) = MakeInfoText("{1}行目：実行比率がすべて0です。", LogLevel.ERR, (row + 1).ToString())
                            errCount += 1
                        End If
                    End If

                End If
            Next

            '最後の行パターン非使用チェック
            Dim lastRowUseptn As Boolean = CBool(DataGridViewTimesheet.Rows(DataGridViewTimesheet.Rows.Count - 2).Cells(COL_CHECK).Value)
            If lastRowUseptn = True Then
                ReDim Preserve errText(errCount)
                errText(errCount) = MakeInfoText("{1}行目：最後の行にパターンは使用できません。", LogLevel.ERR, ((DataGridViewTimesheet.Rows.Count - 2) + 1).ToString())
                errCount += 1
            End If

        End If


        'パターン一覧チェック
        For Each dataTable As DataTable In _patternDataSet.Tables

            'パターンの行数が2行未満の場合
            If dataTable.Rows.Count <= 1 Then

                ReDim Preserve errText(errCount)
                errText(errCount) = MakeInfoText("パターン""" & dataTable.TableName & """ 2行以上記述されていません。", LogLevel.ERR)
                errCount += 1

            Else

                Dim time As String
                Dim lastTime As String
                Dim dir As String
                Dim spd As String

                '各行チェック
                For row = 0 To dataTable.Rows.Count - 1

                    time = dataTable.Rows(row).Item(COL_VTIME).ToString
                    dir = dataTable.Rows(row).Item(COL_VDIR).ToString
                    spd = dataTable.Rows(row).Item(COL_VSPD).ToString

                    '時間の記述チェック
                    If String.IsNullOrEmpty(time) Then
                        ReDim Preserve errText(errCount)
                        errText(errCount) = MakeInfoText("パターン""" & dataTable.TableName & """ {1}行目：時間が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                        errCount += 1
                    Else
                        If row <> 0 Then
                            lastTime = dataTable.Rows(row - 1).Item(COL_VTIME).ToString
                            If Not String.IsNullOrEmpty(lastTime) AndAlso
                           CInt(time) <= CInt(lastTime) Then
                                ReDim Preserve errText(errCount)
                                errText(errCount) = MakeInfoText("パターン""" & dataTable.TableName & """ {1}行目：前の行より時間が増加していません。", LogLevel.ERR, (row + 1).ToString())
                                errCount += 1
                            End If
                        End If
                    End If

                    '方向の記述チェック（最終行は次のパターンの数値を使用するため、チェック不要）
                    If row <> dataTable.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(dir) Then
                            ReDim Preserve errText(errCount)
                            errText(errCount) = MakeInfoText("パターン""" & dataTable.TableName & """ {1}行目：方向が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                            errCount += 1
                        End If
                    End If

                    '速度の記述チェック（最終行は次のパターンの数値を使用するため、チェック不要）
                    If row <> dataTable.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(spd) Then
                            ReDim Preserve errText(errCount)
                            errText(errCount) = MakeInfoText("パターン""" & dataTable.TableName & """ {1}行目：速度が記述されていません。", LogLevel.ERR, (row + 1).ToString())
                            errCount += 1
                        End If
                    End If

                Next

            End If

        Next

        'エラーテキストを出力
        Dim outputText As String = ""
        If Not errText Is Nothing Then
            For i As Integer = 0 To errText.Length - 1
                outputText = outputText & errText(i) & vbCrLf
            Next
        End If
        TextBoxLog.Text = outputText

        '出力可否判断
        If errCount = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    ''' <summary>
    ''' テキストを作成
    ''' </summary>
    ''' <param name="infoText"></param>
    ''' <param name="textArgs1"></param>
    ''' <param name="textArgs2"></param>
    ''' <param name="textArgs3"></param>
    ''' <returns></returns>
    Private Function MakeInfoText(ByVal infoText As String,
                                 ByVal logLevel As LogLevel,
                                 ByVal Optional textArgs1 As String = "",
                                 ByVal Optional textArgs2 As String = "",
                                 ByVal Optional textArgs3 As String = "") As String

        Dim returnText As String = infoText

        If logLevel = LogLevel.INFO Then
            returnText = "[INFO]" & returnText
        ElseIf logLevel = LogLevel.WARN Then
            returnText = "[WARNING]" & returnText
        ElseIf logLevel = LogLevel.ERR Then
            returnText = "[ERROR]" & returnText
        ElseIf logLevel = LogLevel.FATAL Then
            returnText = "[FATAL]" & returnText
        End If

        If Not textArgs1.Equals("") Then
            returnText = returnText.Replace("{1}", textArgs1)
        End If
        If Not textArgs2.Equals("") Then
            returnText = returnText.Replace("{2}", textArgs2)
        End If
        If Not textArgs3.Equals("") Then
            returnText = returnText.Replace("{3}", textArgs3)
        End If

        Return returnText

    End Function


    Private Function GenerateTimeSheetRowsFromDataGridView() As List(Of TimeSheetRow)
        Dim rdmObj As New Random(Environment.TickCount)

        Dim result As New List(Of TimeSheetRow)

        Try

            For row = 0 To DataGridViewTimesheet.Rows.Count - 2

                Dim time As Integer = CInt(DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value)

                Dim dir As Integer
                Dim spd As Integer
                Dim useptn As Boolean
                Dim userdm As Boolean
                Dim useitv As Boolean


                useptn = CBool(DataGridViewTimesheet.Rows(row).Cells(COL_CHECK).Value)

                If useptn = False Then

                    dir = CInt(DataGridViewTimesheet.Rows(row).Cells(COL_VDIR).Value)
                    spd = CInt(DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value)

                    result.Add(New TimeSheetRow With {
                    .Time = time,
                    .Dir = dir,
                    .Spd = spd
                })

                    Continue For

                Else

                    Dim nextTime As Integer
                    Dim reachedEndTime As Boolean = False
                    Dim selectedPattern As String
                    Dim selectedPatternNumber As Integer = 0
                    Dim lastPatternNumber As Integer = 0

                    nextTime = CInt(DataGridViewTimesheet.Rows(row + 1).Cells(COL_VTIME).Value)
                    userdm = CBool(DataGridViewTimesheet.Rows(row).Cells(COL_USERDM).Value)
                    useitv = CBool(DataGridViewTimesheet.Rows(row).Cells(COL_USEITV).Value)

                    Do Until reachedEndTime = True

                        If userdm = False Then
                            selectedPattern = GetNextPattern(row, lastPatternNumber, selectedPatternNumber)
                        Else
                            selectedPattern = GetRandomPatterns(row, rdmObj, selectedPatternNumber)
                        End If

                        Dim r = GenerateTimeSheetRowsWithPattern(row, selectedPattern, time, nextTime, selectedPatternNumber)
                        result.AddRange(r.Item1)
                        reachedEndTime = r.Item2

                        If reachedEndTime = True Then
                            Exit Do
                        End If

                        If useitv = True Then
                            Dim r2 = GenerateTimeSheetRowsWithInterval(row, rdmObj, time, nextTime)
                            result.AddRange(r2.Item1)
                            reachedEndTime = r2.Item2
                        End If

                    Loop

                End If

            Next


            Return result
        Catch ex As OutOfMemoryException
            TextBoxLog.Text = "TimeSheetRowsの生成に失敗しました。"
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' タイムシート生成
    ''' </summary>
    ''' <param name="fileName"></param>
    Private Sub GenerateTimesheet(ByVal fileName As String)
        Dim timeSheetRows = GenerateTimeSheetRowsFromDataGridView()

        Dim sb = New StringBuilder

        For Each row As TimeSheetRow In timeSheetRows
            sb.AppendLine(row.ToCsvLine)
        Next

        File.WriteAllText(fileName, sb.ToString(), System.Text.Encoding.UTF8)
    End Sub

    Private Function GetNextPattern(ByVal row As Integer,
                                    ByRef lastPatternNumber As Integer,
                                    ByRef thisPatternNumber As Integer) As String

        Dim patternName As String = String.Empty
        Dim patn1 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN1).Value, String)
        Dim patn2 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN2).Value, String)
        Dim patn3 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN3).Value, String)
        Dim patn4 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN4).Value, String)

        Select Case lastPatternNumber

            Case 0 '代入していない場合の初期値は0
                If Not String.IsNullOrEmpty(patn1) Then
                    patternName = patn1
                    thisPatternNumber = 1
                ElseIf Not String.IsNullOrEmpty(patn2) Then
                    patternName = patn2
                    thisPatternNumber = 2
                ElseIf Not String.IsNullOrEmpty(patn3) Then
                    patternName = patn3
                    thisPatternNumber = 3
                Else
                    patternName = patn4
                    thisPatternNumber = 4
                End If

            Case 1
                If Not String.IsNullOrEmpty(patn2) Then
                    patternName = patn2
                    thisPatternNumber = 2
                ElseIf Not String.IsNullOrEmpty(patn3) Then
                    patternName = patn3
                    thisPatternNumber = 3
                ElseIf Not String.IsNullOrEmpty(patn4) Then
                    patternName = patn4
                    thisPatternNumber = 4
                Else
                    patternName = patn1
                    thisPatternNumber = 1
                End If

            Case 2
                If Not String.IsNullOrEmpty(patn3) Then
                    patternName = patn3
                    thisPatternNumber = 3
                ElseIf Not String.IsNullOrEmpty(patn4) Then
                    patternName = patn4
                    thisPatternNumber = 4
                ElseIf Not String.IsNullOrEmpty(patn1) Then
                    patternName = patn1
                    thisPatternNumber = 1
                Else
                    patternName = patn2
                    thisPatternNumber = 2
                End If

            Case 3
                If Not String.IsNullOrEmpty(patn4) Then
                    patternName = patn4
                    thisPatternNumber = 4
                ElseIf Not String.IsNullOrEmpty(patn1) Then
                    patternName = patn1
                    thisPatternNumber = 1
                ElseIf Not String.IsNullOrEmpty(patn2) Then
                    patternName = patn2
                    thisPatternNumber = 2
                Else
                    patternName = patn3
                    thisPatternNumber = 3
                End If

            Case 4
                If Not String.IsNullOrEmpty(patn1) Then
                    patternName = patn1
                    thisPatternNumber = 1
                ElseIf Not String.IsNullOrEmpty(patn2) Then
                    patternName = patn2
                    thisPatternNumber = 2
                ElseIf Not String.IsNullOrEmpty(patn3) Then
                    patternName = patn3
                    thisPatternNumber = 3
                Else
                    patternName = patn4
                    thisPatternNumber = 4
                End If

        End Select

        lastPatternNumber = thisPatternNumber
        Return patternName

    End Function


    Private Function GetRandomPatterns(ByVal row As Integer,
                                       ByRef rdmObj As Random,
                                       ByRef nextPatternNumber As Integer) As String

        Dim patn1 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN1).Value, String)
        Dim patn2 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN2).Value, String)
        Dim patn3 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN3).Value, String)
        Dim patn4 As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN4).Value, String)

        Dim rate1 As Integer = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE1).Value, String)
        Dim rate2 As Integer = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE2).Value, String)
        Dim rate3 As Integer = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE3).Value, String)
        Dim rate4 As Integer = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE4).Value, String)

        Dim patternList As New List(Of String)
        Dim rateList As New List(Of Integer)

        If Not String.IsNullOrEmpty(patn1) Then
            patternList.Add(patn1)
            rateList.Add(rate1)
        End If
        If Not String.IsNullOrEmpty(patn2) Then
            patternList.Add(patn2)
            rateList.Add(rate2)
        End If
        If Not String.IsNullOrEmpty(patn3) Then
            patternList.Add(patn3)
            rateList.Add(rate3)
        End If
        If Not String.IsNullOrEmpty(patn4) Then
            patternList.Add(patn4)
            rateList.Add(rate4)
        End If

        Dim totalRate As Integer = 1
        For Each rate As Integer In rateList
            totalRate += rate
        Next

        Dim randomRate As Integer
        randomRate = rdmObj.Next(1, totalRate) '1以上totalRate未満

        Dim addedRate As Integer = 0
        Dim patternNumber As Integer = 0

        For Each ptnrate As Integer In rateList
            addedRate += ptnrate
            If addedRate >= randomRate Then
                nextPatternNumber = patternNumber + 1
                Return patternList(patternNumber)
            End If
            patternNumber += 1
        Next

        Return String.Empty 'エラー

    End Function



    Private Function GenerateTimeSheetRowsWithPattern(ByVal row As Integer,
                                         ByVal patternName As String,
                                         ByRef time As Integer,
                                         ByVal nextTime As Integer,
                                         ByVal patternNumber As Integer) As Tuple(Of List(Of TimeSheetRow), Boolean)

        Dim result As New List(Of TimeSheetRow)
        Dim baseTime As Integer = time
        Dim dspd As Integer = GetDspd(row, patternNumber) '補正速度
        Dim patternDataTable As DataTable = _patternDataSet.Tables(patternName)

        '「次の行の時間 - 0.1秒」以上ならば、CSV行を記述しない
        Dim endTime As Integer = nextTime - 1

        For i = 0 To patternDataTable.Rows.Count - 2

            time = baseTime + CInt(patternDataTable.Rows(i).Item(COL_VTIME))


            If time >= endTime Then
                Return Tuple.Create(result, True)
            End If

            Dim dir As Integer = CInt(patternDataTable.Rows(i).Item(COL_VDIR))
            Dim spd As Integer = CInt(patternDataTable.Rows(i).Item(COL_VSPD))

            If spd <> 0 Then
                If (spd + dspd) < 0 Then
                    spd = 0
                ElseIf (spd + dspd) > 100 Then
                    spd = 100
                Else
                    spd = spd + dspd
                End If
            End If

            result.Add(New TimeSheetRow With {
                .Time = time,
                .Dir = dir,
                .Spd = spd
            })
        Next

        time = baseTime + CInt(patternDataTable.Rows(patternDataTable.Rows.Count - 1).Item(COL_VTIME))
        If time >= endTime Then
            Return Tuple.Create(result, True)
        End If
        Return Tuple.Create(result, False)

    End Function


    Private Function GenerateTimeSheetRowsWithInterval(ByVal row As Integer,
                                          ByRef rdmObj As Random,
                                          ByRef Time As Integer,
                                          ByVal nextTime As Integer) As Tuple(Of List(Of TimeSheetRow), Boolean)

        Dim itvminValue As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_ITVMIN).Value, Integer)
        Dim itvmaxValue As String = CType(DataGridViewTimesheet.Rows(row).Cells(COL_ITVMAX).Value, Integer)
        Dim itvminStr As String = String.Empty
        Dim itvmaxStr As String = String.Empty

        Dim result As New List(Of TimeSheetRow)

        Select Case itvminValue
            Case 0
                itvminStr = _interval_time00
            Case 1
                itvminStr = _interval_time01
            Case 2
                itvminStr = _interval_time02
            Case 3
                itvminStr = _interval_time03
            Case 4
                itvminStr = _interval_time04
            Case 5
                itvminStr = _interval_time05
            Case 6
                itvminStr = _interval_time06
            Case 7
                itvminStr = _interval_time07
            Case 8
                itvminStr = _interval_time08
            Case 9
                itvminStr = _interval_time09
            Case 10
                itvminStr = _interval_time10
            Case 11
                itvminStr = _interval_time11
            Case 12
                itvminStr = _interval_time12
            Case 13
                itvminStr = _interval_time13
            Case 14
                itvminStr = _interval_time14
            Case 15
                itvminStr = _interval_time15
            Case 16
                itvminStr = _interval_time16
            Case 17
                itvminStr = _interval_time17
            Case 18
                itvminStr = _interval_time18
            Case 19
                itvminStr = _interval_time19
            Case 20
                itvminStr = _interval_time20
        End Select

        Select Case itvmaxValue
            Case 0
                itvmaxStr = _interval_time00
            Case 1
                itvmaxStr = _interval_time01
            Case 2
                itvmaxStr = _interval_time02
            Case 3
                itvmaxStr = _interval_time03
            Case 4
                itvmaxStr = _interval_time04
            Case 5
                itvmaxStr = _interval_time05
            Case 6
                itvmaxStr = _interval_time06
            Case 7
                itvmaxStr = _interval_time07
            Case 8
                itvmaxStr = _interval_time08
            Case 9
                itvmaxStr = _interval_time09
            Case 10
                itvmaxStr = _interval_time10
            Case 11
                itvmaxStr = _interval_time11
            Case 12
                itvmaxStr = _interval_time12
            Case 13
                itvmaxStr = _interval_time13
            Case 14
                itvmaxStr = _interval_time14
            Case 15
                itvmaxStr = _interval_time15
            Case 16
                itvmaxStr = _interval_time16
            Case 17
                itvmaxStr = _interval_time17
            Case 18
                itvmaxStr = _interval_time18
            Case 19
                itvmaxStr = _interval_time19
            Case 20
                itvmaxStr = _interval_time20
        End Select

        Dim itvmin As Double
        Dim itvmax As Double

        '「_interval_time」は秒単位のため、0.1秒単位に変換する
        itvmin = CDbl(itvminStr) * 10
        itvmax = CDbl(itvmaxStr) * 10

        Dim itvTime As Integer

        If itvmin = itvmax Then

            itvTime = itvmin

        Else
            'itvmin以上、(itvmax+1)未満
            itvTime = rdmObj.Next(itvmin, itvmax + 1)
        End If

        '0秒のインターバルタイムが指定された場合、CSVを記述しない
        If itvTime = 0 Then
            Return Tuple.Create(Of List(Of TimeSheetRow), Boolean)(Nothing, False)
        End If

        result.Add(New TimeSheetRow With {
            .Time = Time,
            .Dir = False,
            .Spd = 0
        })

        '「次の行の時間 - 0.1秒」以上ならば、CSV行を記述しない
        Dim endTime As Integer = nextTime - 1

        Time = Time + itvTime
        If Time > endTime Then
            Return Tuple.Create(result, True)
        End If
        Return Tuple.Create(result, False)

    End Function

    Private Function GetDspd(ByVal Row As Integer,
                             ByVal patternNumber As Integer) As Integer
        Dim dspdValue As Integer
        Dim dspd As Integer

        Dim rate0 As String = String.Empty
        Dim rate1 As String = String.Empty
        Dim rate2 As String = String.Empty
        Dim rate3 As String = String.Empty
        Dim rate4 As String = String.Empty
        Dim rate5 As String = String.Empty
        Dim rate6 As String = String.Empty
        Dim rate7 As String = String.Empty
        Dim rate8 As String = String.Empty

        Select Case patternNumber

            Case 1
                dspdValue = CType(DataGridViewTimesheet.Rows(Row).Cells(COL_DSPD1).Value, String)
                rate0 = _pattern1_dspd0
                rate1 = _pattern1_dspd1
                rate2 = _pattern1_dspd2
                rate3 = _pattern1_dspd3
                rate4 = _pattern1_dspd4
                rate5 = _pattern1_dspd5
                rate6 = _pattern1_dspd6
                rate7 = _pattern1_dspd7
                rate8 = _pattern1_dspd8
            Case 2
                dspdValue = CType(DataGridViewTimesheet.Rows(Row).Cells(COL_DSPD2).Value, String)
                rate0 = _pattern2_dspd0
                rate1 = _pattern2_dspd1
                rate2 = _pattern2_dspd2
                rate3 = _pattern2_dspd3
                rate4 = _pattern2_dspd4
                rate5 = _pattern2_dspd5
                rate6 = _pattern2_dspd6
                rate7 = _pattern2_dspd7
                rate8 = _pattern2_dspd8
            Case 3
                dspdValue = CType(DataGridViewTimesheet.Rows(Row).Cells(COL_DSPD3).Value, String)
                rate0 = _pattern3_dspd0
                rate1 = _pattern3_dspd1
                rate2 = _pattern3_dspd2
                rate3 = _pattern3_dspd3
                rate4 = _pattern3_dspd4
                rate5 = _pattern3_dspd5
                rate6 = _pattern3_dspd6
                rate7 = _pattern3_dspd7
                rate8 = _pattern3_dspd8
            Case 4
                dspdValue = CType(DataGridViewTimesheet.Rows(Row).Cells(COL_DSPD4).Value, String)
                rate0 = _pattern4_dspd0
                rate1 = _pattern4_dspd1
                rate2 = _pattern4_dspd2
                rate3 = _pattern4_dspd3
                rate4 = _pattern4_dspd4
                rate5 = _pattern4_dspd5
                rate6 = _pattern4_dspd6
                rate7 = _pattern4_dspd7
                rate8 = _pattern4_dspd8
        End Select

        Select Case dspdValue
            Case 0
                dspd = rate0
            Case 1
                dspd = rate1
            Case 2
                dspd = rate2
            Case 3
                dspd = rate3
            Case 4
                dspd = rate4
            Case 5
                dspd = rate5
            Case 6
                dspd = rate6
            Case 7
                dspd = rate7
            Case 8
                dspd = rate8
        End Select

        Return dspd

    End Function



    ''' <summary>
    ''' 行追加ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemAddRow_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemAddRow.Click

        If _selectedRow <> -1 Then
            DataGridViewTimesheet.Rows.Insert(_selectedRow)

            '行のデフォルト値設定
            DefineRow(DataGridViewTimesheet.Rows(_selectedRow - 1))

        End If

    End Sub


    ''' <summary>
    ''' 行削除ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemDeleteRow_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemDeleteRow.Click

        '行未選択時、または最終行は削除を実行しない
        If _selectedRow <> -1 AndAlso _selectedRow <> DataGridViewTimesheet.Rows.Count - 1 Then
            DataGridViewTimesheet.Rows.RemoveAt(_selectedRow)
        End If

    End Sub


    ''' <summary>
    ''' 時間キャプチャ調整 クリック時 排他チェック制御(dobon参考)、ディレイ設定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ToolStripMenuItem_plus5s.Click,
                ToolStripMenuItem_plus4s.Click,
                ToolStripMenuItem_plus3s.Click,
                ToolStripMenuItem_plus2s.Click,
                ToolStripMenuItem_plus1s.Click,
                ToolStripMenuItem_plus0s.Click,
                ToolStripMenuItem_minus1s.Click,
                ToolStripMenuItem_minus2s.Click,
                ToolStripMenuItem_minus3s.Click,
                ToolStripMenuItem_minus4s.Click,
                ToolStripMenuItem_minus5s.Click

        'グループのToolStripMenuItemを配列にしておく
        Dim groupMenuItems As ToolStripMenuItem() = New ToolStripMenuItem() {
            ToolStripMenuItem_plus5s,
            ToolStripMenuItem_plus4s,
            ToolStripMenuItem_plus3s,
            ToolStripMenuItem_plus2s,
            ToolStripMenuItem_plus1s,
            ToolStripMenuItem_plus0s,
            ToolStripMenuItem_minus1s,
            ToolStripMenuItem_minus2s,
            ToolStripMenuItem_minus3s,
            ToolStripMenuItem_minus4s,
            ToolStripMenuItem_minus5s}

        'グループのToolStripMenuItemを列挙する
        For Each item As ToolStripMenuItem In groupMenuItems
            If Object.ReferenceEquals(item, sender) Then
                'ClickされたToolStripMenuItemならば、Indeterminateにする
                item.CheckState = CheckState.Indeterminate
            Else
                'ClickされたToolStripMenuItemでなければ、Uncheckedにする
                item.CheckState = CheckState.Unchecked
            End If
        Next

        'ディレイ設定　変更
        Dim infoTextCaptureDelay As String
        Select Case True
            Case sender Is ToolStripMenuItem_plus5s
                _capture_delay = 5
                infoTextCaptureDelay = ToolStripMenuItem_plus5s.Text
            Case sender Is ToolStripMenuItem_plus4s
                _capture_delay = 4
                infoTextCaptureDelay = ToolStripMenuItem_plus4s.Text
            Case sender Is ToolStripMenuItem_plus3s
                _capture_delay = 3
                infoTextCaptureDelay = ToolStripMenuItem_plus3s.Text
            Case sender Is ToolStripMenuItem_plus2s
                _capture_delay = 2
                infoTextCaptureDelay = ToolStripMenuItem_plus2s.Text
            Case sender Is ToolStripMenuItem_plus1s
                _capture_delay = 1
                infoTextCaptureDelay = ToolStripMenuItem_plus1s.Text
            Case sender Is ToolStripMenuItem_plus0s
                _capture_delay = 0
                infoTextCaptureDelay = ToolStripMenuItem_plus0s.Text
            Case sender Is ToolStripMenuItem_minus1s
                _capture_delay = -1
                infoTextCaptureDelay = ToolStripMenuItem_minus1s.Text
            Case sender Is ToolStripMenuItem_minus2s
                _capture_delay = -2
                infoTextCaptureDelay = ToolStripMenuItem_minus2s.Text
            Case sender Is ToolStripMenuItem_minus3s
                _capture_delay = -3
                infoTextCaptureDelay = ToolStripMenuItem_minus3s.Text
            Case sender Is ToolStripMenuItem_minus4s
                _capture_delay = -4
                infoTextCaptureDelay = ToolStripMenuItem_minus4s.Text
            Case sender Is ToolStripMenuItem_minus5s
                _capture_delay = -5
                infoTextCaptureDelay = ToolStripMenuItem_minus5s.Text
            Case Else
                _capture_delay = 0
                infoTextCaptureDelay = ToolStripMenuItem_plus0s.Text
        End Select

        TextBoxLog.Text =
            MakeInfoText("時間キャプチャ　ディレイ時間：" & infoTextCaptureDelay, LogLevel.INFO)

    End Sub

#End Region



#Region "DataGridView"


    ''' <summary>
    ''' 編集開始時、IMEモードを設定する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTimesheet.CellEnter

        Select Case e.ColumnIndex

            Case DataGridViewTimesheet.Columns(COL_VTIME).Index,
                 DataGridViewTimesheet.Columns(COL_VDIR).Index,
                 DataGridViewTimesheet.Columns(COL_VSPD).Index

                DataGridViewTimesheet.ImeMode = System.Windows.Forms.ImeMode.Disable

            Case DataGridViewTimesheet.Columns(COL_VCMT).Index

                DataGridViewTimesheet.ImeMode = System.Windows.Forms.ImeMode.Inherit

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
    Private Sub DataGridViewTimesheet_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridViewTimesheet.EditingControlShowing

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
        RemoveHandler tb.KeyPress, AddressOf AllowKeyPressDigitOrSpace
        RemoveHandler tb.KeyPress, AddressOf AllowKeyPressDigit
        RemoveHandler tb.KeyPress, AddressOf AllowKeyPressZeroOne

        '該当する列か判定し、ハンドラを追加
        Select Case dgv.CurrentCell.OwningColumn.Name
            Case COL_VTIME
                AddHandler tb.KeyPress, AddressOf AllowKeyPressDigitOrSpace
            Case COL_VSPD
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
    ''' 数字かスペースキーのみ入力できる制御
    ''' 
    ''' 参考：Dobon.net「DataGridViewでセルが編集中の時にキーイベントを捕捉する」
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AllowKeyPressDigitOrSpace(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        If (Not Char.IsDigit(e.KeyChar)) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
            'もしスペースの入力なら時間キャプチャを行う
            If e.KeyChar = " " Then
                If ButtonTimeCapture.Enabled = True Then

                    CaptureTime(DataGridViewTimesheet.CurrentRow.Index)

                End If
            End If
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
    Private Sub DataGridViewTimesheet_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridViewTimesheet.CellPainting

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
    ''' 選択行変更時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTimesheet.RowEnter

        _selectedRow = e.RowIndex

        '行頭の選択の場合は何もしない
        If _selectedRow = -1 Then
            Exit Sub
        End If

        '詳細欄の表示を更新
        RefreshDetail(_selectedRow)

    End Sub


    ''' <summary>
    ''' DataGridViewの選択中の行により右エリアの表示を更新する
    ''' </summary>
    ''' <param name="row"></param>
    Public Sub RefreshDetail(row As Integer)

        '選択行が存在しない場合は処理を終了
        If row = -1 Then
            Exit Sub
        End If

        'パターンを使用するチェックボックス　へ反映
        If CType(DataGridViewTimesheet.Rows(row).Cells(COL_CHECK).Value, Boolean) = True Then
            CheckBoxUsePattern.CheckState = CheckState.Checked
        Else
            CheckBoxUsePattern.CheckState = CheckState.Unchecked
        End If

        'パターン名コンボボックス　へ反映
        ComboBoxPattern1.Text = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN1).Value, String)
        ComboBoxPattern2.Text = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN2).Value, String)
        ComboBoxPattern3.Text = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN3).Value, String)
        ComboBoxPattern4.Text = CType(DataGridViewTimesheet.Rows(row).Cells(COL_PATN4).Value, String)

        '速度補正トラックバー　へ反映
        TrackBarPattern1dspd.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_DSPD1).Value, Integer)
        TrackBarPattern2dspd.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_DSPD2).Value, Integer)
        TrackBarPattern3dspd.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_DSPD3).Value, Integer)
        TrackBarPattern4dspd.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_DSPD4).Value, Integer)

        '速度補正ラベル　へ反映
        DisplayDSPD(TrackBarPattern1dspd)
        DisplayDSPD(TrackBarPattern2dspd)
        DisplayDSPD(TrackBarPattern3dspd)
        DisplayDSPD(TrackBarPattern4dspd)

        '実行割合トラックバー　へ反映
        TrackBarPattern1rate.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE1).Value, Integer)
        TrackBarPattern2rate.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE2).Value, Integer)
        TrackBarPattern3rate.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE3).Value, Integer)
        TrackBarPattern4rate.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_RATE4).Value, Integer)

        '実行割合ラベル　へ反映
        DisplayRATE(TrackBarPattern1rate)
        DisplayRATE(TrackBarPattern2rate)
        DisplayRATE(TrackBarPattern3rate)
        DisplayRATE(TrackBarPattern4rate)

        '実行モードラジオボタン　へ反映
        If CType(DataGridViewTimesheet.Rows(row).Cells(COL_USERDM).Value, Boolean) = True Then
            RadioButtonRandom.Checked = True
        Else
            RadioButtonOrder.Checked = True
        End If

        'パターン間に休止をはさむチェックボックス　へ反映
        If CType(DataGridViewTimesheet.Rows(row).Cells(COL_USEITV).Value, Boolean) = True Then
            CheckBoxUseInterval.CheckState = CheckState.Checked
        Else
            CheckBoxUseInterval.CheckState = CheckState.Unchecked
        End If

        '最小時間ラベル　へ反映
        TrackBarItvMinTime.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_ITVMIN).Value, Integer)
        DisplayItvTime(TrackBarItvMinTime)

        '最大時間ラベル　へ反映
        TrackBarItvMaxTime.Value = CType(DataGridViewTimesheet.Rows(row).Cells(COL_ITVMAX).Value, Integer)
        DisplayItvTime(TrackBarItvMaxTime)

        '詳細欄　選択可能/不可能　設定
        If CType(DataGridViewTimesheet.Rows(row).Cells(COL_CHECK).Value, Boolean) = True Then

            'パターン名コンボボックス
            ComboBoxPattern1.Enabled = True
            ComboBoxPattern2.Enabled = True
            ComboBoxPattern3.Enabled = True
            ComboBoxPattern4.Enabled = True

            '速度補正ラベル
            TrackBarPattern1dspd.Enabled = True
            TrackBarPattern2dspd.Enabled = True
            TrackBarPattern3dspd.Enabled = True
            TrackBarPattern4dspd.Enabled = True

            '実行割合トラックバー（ランダム実行の場合のみ有効）
            If CType(DataGridViewTimesheet.Rows(row).Cells(COL_USERDM).Value, Boolean) = True Then
                TrackBarPattern1rate.Enabled = True
                TrackBarPattern2rate.Enabled = True
                TrackBarPattern3rate.Enabled = True
                TrackBarPattern4rate.Enabled = True
            Else
                TrackBarPattern1rate.Enabled = False
                TrackBarPattern2rate.Enabled = False
                TrackBarPattern3rate.Enabled = False
                TrackBarPattern4rate.Enabled = False
            End If

            '実行モードラジオボタン
            RadioButtonOrder.Enabled = True
            RadioButtonRandom.Enabled = True

            'パターン間に休止をはさむチェックボックス
            CheckBoxUseInterval.Enabled = True

            '最小時間ラベル
            TrackBarItvMinTime.Enabled = True

            '最大時間ラベル
            TrackBarItvMaxTime.Enabled = True

        Else

            'パターン名コンボボックス
            ComboBoxPattern1.Enabled = False
            ComboBoxPattern2.Enabled = False
            ComboBoxPattern3.Enabled = False
            ComboBoxPattern4.Enabled = False

            '速度補正ラベル
            TrackBarPattern1dspd.Enabled = False
            TrackBarPattern2dspd.Enabled = False
            TrackBarPattern3dspd.Enabled = False
            TrackBarPattern4dspd.Enabled = False

            '実行割合トラックバー
            TrackBarPattern1rate.Enabled = False
            TrackBarPattern2rate.Enabled = False
            TrackBarPattern3rate.Enabled = False
            TrackBarPattern4rate.Enabled = False

            '実行モードラジオボタン
            RadioButtonOrder.Enabled = False
            RadioButtonRandom.Enabled = False

            'パターン間に休止をはさむチェックボックス
            CheckBoxUseInterval.Enabled = False

            '最小時間ラベル
            TrackBarItvMinTime.Enabled = False

            '最大時間ラベル
            TrackBarItvMaxTime.Enabled = False

        End If

    End Sub


    ''' <summary>
    ''' DataGridViewのパターンチェックボックスがクリックされた時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTimesheet.CellContentClick

        '選択行更新
        _selectedRow = e.RowIndex

        If _selectedRow = -1 Then
            Exit Sub
        End If

        If Not e.ColumnIndex = DataGridViewTimesheet.Columns(COL_CHECK).Index Then
            Exit Sub
        End If

        '編集状態を強制的に終了し、以降の描画の更新、詳細欄の更新ができるようにする
        DataGridViewTimesheet.EndEdit()

        '行のアクティベイト、描画を更新
        DrawDataGridViewRow(_selectedRow)

        '詳細欄の表示を更新
        RefreshDetail(_selectedRow)

    End Sub


    ''' <summary>
    ''' DataGridViewの行のアクティブ・非アクティブ・色を設定
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub DrawDataGridViewRow(row As Integer)

        Dim isReadOnlyDirSpd As Boolean
        Dim foreColorDir As Color
        Dim backColorDir As Color
        Dim foreColorSpd As Color
        Dim backColorSpd As Color
        Dim selectionForeColorDirSpd As Color
        Dim selectionBackColorDirSpd As Color

        Dim isReadOnlyPtn As Boolean
        Dim foreColorPtn As Color
        Dim backColorPtn As Color
        Dim selectionForeColorPtn As Color
        Dim selectionBackColorPtn As Color

        Dim spd As Double

        '先頭行（ヘッダ）クリック時は何もしない
        If row = -1 Then
            Exit Sub
        End If

        '方向・速度の文字色・背景色の定義
        If CBool(DataGridViewTimesheet.Rows(row).Cells(COL_CHECK).Value) = True Then

            '方向
            isReadOnlyDirSpd = True
            foreColorDir = Color.FromArgb(&HFF, &HFF, &HFF)
            backColorDir = Color.FromArgb(&HFF, &HFF, &HFF)

            '速度
            isReadOnlyDirSpd = True
            foreColorSpd = Color.FromArgb(&HFF, &HFF, &HFF)
            backColorSpd = Color.FromArgb(&HFF, &HFF, &HFF)

            '方向・速度（選択時）
            selectionForeColorDirSpd = Color.FromArgb(&HFF, &HFF, &HFF)
            selectionBackColorDirSpd = Color.FromArgb(&HFF, &HFF, &HFF)

            'パターン
            isReadOnlyPtn = False
            foreColorPtn = Color.Black
            backColorPtn = Color.FromArgb(&HFF, &HFF, &HFF)

            'パターン（選択時）
            selectionForeColorPtn = Color.Empty
            selectionBackColorPtn = Color.Empty

        Else

            '方向
            isReadOnlyDirSpd = False
            foreColorDir = Color.Empty
            backColorDir = Color.Empty

            '方向・速度（選択時）
            selectionForeColorDirSpd = Color.Empty
            selectionBackColorDirSpd = Color.Empty

            'パターン
            isReadOnlyPtn = True
            foreColorPtn = Color.FromArgb(&HEC, &HEC, &HEC)
            backColorPtn = Color.FromArgb(&HEC, &HEC, &HEC)

            'パターン（選択時）
            selectionForeColorPtn = Color.FromArgb(&HEC, &HEC, &HEC)
            selectionBackColorPtn = Color.FromArgb(&HEC, &HEC, &HEC)

            '速度
            If String.IsNullOrEmpty(DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value) Then

                '未入力の場合
                foreColorSpd = Color.Empty
                backColorSpd = Color.Empty

            Else

                '速度の数値ごとの配色定義
                spd = CDbl(DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value)

                If spd = 0 Then
                    foreColorSpd = Color.Empty
                    backColorSpd = Color.FromArgb(&HFF, &HFF, &HFF)

                ElseIf spd < 25 Then
                    foreColorSpd = Color.Empty
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

        End If

        '方向
        DataGridViewTimesheet(COL_VDIR, row).ReadOnly = isReadOnlyDirSpd   '読み込みのみ
        DataGridViewTimesheet(COL_VDIR, row).Style.ForeColor = foreColorDir '文字色
        DataGridViewTimesheet(COL_VDIR, row).Style.BackColor = backColorDir '背景色
        DataGridViewTimesheet(COL_VDIR, row).Style.SelectionForeColor = selectionForeColorDirSpd '選択時背景色
        DataGridViewTimesheet(COL_VDIR, row).Style.SelectionBackColor = selectionBackColorDirSpd '選択時文字色

        '速度
        DataGridViewTimesheet(COL_VSPD, row).ReadOnly = isReadOnlyDirSpd   '読み込みのみ
        DataGridViewTimesheet(COL_VSPD, row).Style.ForeColor = foreColorSpd '文字色
        DataGridViewTimesheet(COL_VSPD, row).Style.BackColor = backColorSpd '背景色
        DataGridViewTimesheet(COL_VSPD, row).Style.SelectionForeColor = selectionForeColorDirSpd '選択時背景色
        DataGridViewTimesheet(COL_VSPD, row).Style.SelectionBackColor = selectionBackColorDirSpd '選択時文字色

        'パターン
        DataGridViewTimesheet(COL_PATN1, row).ReadOnly = isReadOnlyPtn '読み込みのみ
        DataGridViewTimesheet(COL_PATN1, row).Style.ForeColor = foreColorPtn '文字色
        DataGridViewTimesheet(COL_PATN1, row).Style.BackColor = backColorPtn '背景色
        DataGridViewTimesheet(COL_PATN1, row).Style.SelectionForeColor = selectionForeColorPtn '選択時文字色
        DataGridViewTimesheet(COL_PATN1, row).Style.SelectionBackColor = selectionBackColorPtn '選択時背景色
        DataGridViewTimesheet(COL_PATN2, row).ReadOnly = isReadOnlyPtn
        DataGridViewTimesheet(COL_PATN2, row).Style.ForeColor = foreColorPtn
        DataGridViewTimesheet(COL_PATN2, row).Style.BackColor = backColorPtn
        DataGridViewTimesheet(COL_PATN2, row).Style.SelectionForeColor = selectionForeColorPtn
        DataGridViewTimesheet(COL_PATN2, row).Style.SelectionBackColor = selectionBackColorPtn
        DataGridViewTimesheet(COL_PATN3, row).ReadOnly = isReadOnlyPtn
        DataGridViewTimesheet(COL_PATN3, row).Style.ForeColor = foreColorPtn
        DataGridViewTimesheet(COL_PATN3, row).Style.BackColor = backColorPtn
        DataGridViewTimesheet(COL_PATN3, row).Style.SelectionForeColor = selectionForeColorPtn
        DataGridViewTimesheet(COL_PATN3, row).Style.SelectionBackColor = selectionBackColorPtn
        DataGridViewTimesheet(COL_PATN4, row).ReadOnly = isReadOnlyPtn
        DataGridViewTimesheet(COL_PATN4, row).Style.ForeColor = foreColorPtn
        DataGridViewTimesheet(COL_PATN4, row).Style.BackColor = backColorPtn
        DataGridViewTimesheet(COL_PATN4, row).Style.SelectionForeColor = selectionForeColorPtn
        DataGridViewTimesheet(COL_PATN4, row).Style.SelectionBackColor = selectionBackColorPtn

        '動作時間の文字色・背景色の定義
        Dim foreColorVTimeDiff As Color
        Dim backColorVTimeDiff As Color
        If String.IsNullOrEmpty(DataGridViewTimesheet.Rows(row).Cells(COL_VTIMEDIFF).Value) Then
            foreColorVTimeDiff = Color.Black
            backColorVTimeDiff = Color.FromArgb(&HEC, &HEC, &HEC)
        Else
            If CDbl(DataGridViewTimesheet.Rows(row).Cells(COL_VTIMEDIFF).Value) > 0 Then
                foreColorVTimeDiff = Color.Black
                backColorVTimeDiff = Color.FromArgb(&HEC, &HEC, &HEC)
            Else
                foreColorVTimeDiff = Color.Red
                backColorVTimeDiff = Color.FromArgb(&HFF, &HDC, &HFF)
            End If
        End If

        DataGridViewTimesheet(COL_VTIMEDIFF, row).Style.ForeColor = foreColorVTimeDiff '文字色
        DataGridViewTimesheet(COL_VTIMEDIFF, row).Style.BackColor = backColorVTimeDiff '背景色

    End Sub


    ''' <summary>
    ''' 編集確定後、入力内容を調整する。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTimesheet.CellEndEdit

        EditInput(e.ColumnIndex, e.RowIndex)

    End Sub


    ''' 入力内容を調整する。
    ''' 時間：先頭0埋めを削除
    ''' 方向：（特になし）
    ''' 速度：先頭0埋めを削除、最大を100とする（101以上は100に変換）
    ''' コメント：（特になし）
    ''' </summary>
    ''' <param name="col"></param>
    ''' <param name="row"></param>
    Private Sub EditInput(ByVal col As Integer, ByVal row As Integer)

        '入力内容の調整
        Select Case col

            Case DataGridViewTimesheet.Columns(COL_VTIME).Index

                Dim setValue As String = DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value
                If Not String.IsNullOrEmpty(setValue) Then
                    DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value = CInt(setValue)

                End If

                '換算時間を求める
                InputVtimeHms(row)

            Case DataGridViewTimesheet.Columns(COL_VSPD).Index

                Dim setValue As String = DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value
                If Not String.IsNullOrEmpty(setValue) Then
                    If CInt(setValue) > 100 Then
                        DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value = 100
                    Else
                        DataGridViewTimesheet.Rows(row).Cells(COL_VSPD).Value = CInt(setValue)
                    End If

                End If

        End Select

        '行のアクティベイト、描画を更新
        'DrawDataGridViewRow(_selectedRow)

        '速度と時間差分の自動計算（編集した行 最下行なら計算しない）
        If Not row = DataGridViewTimesheet.Rows.Count - 1 Then

            InputVtimeDiff(row)

            '描画を更新
            DrawDataGridViewRow(row)

        End If

        '速度と時間差分の自動計算（編集した行より一つ前の行 最初の行なら計算しない）
        If Not row = 0 Then

            InputVtimeDiff(row - 1)

            '描画を更新
            DrawDataGridViewRow(row - 1)

        End If

        '詳細欄の表示を更新
        RefreshDetail(_selectedRow)

    End Sub


    ''' <summary>
    ''' 換算時間を求める
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub InputVtimeHms(ByVal row As Integer)

        Dim setTimeValue As String
        Dim time As String = DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value

        If String.IsNullOrEmpty(time) Then
            setTimeValue = String.Empty
        Else
            Dim milliSec As Integer = CInt(time) * 100
            Dim ts As TimeSpan = New TimeSpan(0, 0, 0, 0, milliSec)

            If milliSec < 3600000 Then
                setTimeValue = ts.ToString("mm':'ss'.'f")
            ElseIf milliSec < 86400000 Then
                '1時間を超える時間からhhを表示
                setTimeValue = ts.ToString("hh':'mm':'ss'.'f")
            Else
                '1日を超える時間ならば'1.hh:mm:ss.f'で表示させる
                setTimeValue = ts.ToString("d'.'hh':'mm':'ss'.'f")
            End If

        End If

        DataGridViewTimesheet.Rows(row).Cells(COL_VTIMEHMS).Value = setTimeValue

    End Sub


    ''' <summary>
    ''' 時間差分を求める
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub InputVtimeDiff(ByVal row As Integer)

        Dim setTimeValue As String = String.Empty
        Dim timeFrom As String = DataGridViewTimesheet.Rows(row).Cells(COL_VTIME).Value
        Dim timeTo As String = DataGridViewTimesheet.Rows(row + 1).Cells(COL_VTIME).Value

        If Not String.IsNullOrEmpty(timeFrom) And
            Not String.IsNullOrEmpty(timeTo) Then
            setTimeValue = CDbl(timeTo) - CDbl(timeFrom)
        End If

        DataGridViewTimesheet.Rows(row).Cells(COL_VTIMEDIFF).Value = setTimeValue

    End Sub

    ''' <summary>
    ''' 新規行追加時、デフォルト値を設定する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_DefaultValuesNeeded(sender As Object, e As DataGridViewRowEventArgs) Handles DataGridViewTimesheet.DefaultValuesNeeded

        DefineRow(e.Row)

    End Sub


    ''' <summary>
    ''' 新規行の定義
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub DefineRow(ByRef row As DataGridViewRow)

        row.Cells(COL_CHECK).Value = False
        row.Cells(COL_VTIME).Value = String.Empty
        row.Cells(COL_VDIR).Value = String.Empty
        row.Cells(COL_VSPD).Value = String.Empty
        row.Cells(COL_VCMT).Value = String.Empty
        row.Cells(COL_PATN1).Value = String.Empty
        row.Cells(COL_PATN2).Value = String.Empty
        row.Cells(COL_PATN3).Value = String.Empty
        row.Cells(COL_PATN4).Value = String.Empty
        row.Cells(COL_DSPD1).Value = 4
        row.Cells(COL_DSPD2).Value = 4
        row.Cells(COL_DSPD3).Value = 4
        row.Cells(COL_DSPD4).Value = 4
        row.Cells(COL_RATE1).Value = 1
        row.Cells(COL_RATE2).Value = 1
        row.Cells(COL_RATE3).Value = 1
        row.Cells(COL_RATE4).Value = 1
        row.Cells(COL_USERDM).Value = False
        row.Cells(COL_USEITV).Value = False
        row.Cells(COL_ITVMIN).Value = 0
        row.Cells(COL_ITVMAX).Value = 0

    End Sub


    ''' <summary>
    ''' 行追加イベント発生時、速度を再計算する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridViewTimesheet.RowsAdded

        '1行目追加時、または最終行編集による行追加時は何もしない
        If e.RowIndex = 0 OrElse e.RowIndex = DataGridViewTimesheet.Rows.Count - 1 Then
            Exit Sub
        End If

        '追加行より1つ前の行の時間差分を再計算し、描画を更新
        InputVtimeDiff(e.RowIndex - 1)
        DrawDataGridViewRow(e.RowIndex - 1)

    End Sub


    ''' <summary>
    ''' 行削除イベント発生時、速度を再計算し、描画を更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridViewTimesheet_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridViewTimesheet.RowsRemoved

        '1行目削除時は何もしない
        '（または「開く」「クリア」実行時にe.RowIndex=0で通過するため、この時も何もしない）
        If e.RowIndex = 0 Then
            Exit Sub
        End If

        '削除行より1つ前の行の時間差分を再計算し、描画を更新
        InputVtimeDiff(e.RowIndex - 1)
        DrawDataGridViewRow(e.RowIndex - 1)

    End Sub

#End Region



#Region "DataGridView以外イベント"

    ''' <summary>
    ''' 動画・音声から時間キャプチャボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonTimeCapture_Click(sender As Object, e As EventArgs) Handles ButtonTimeCapture.Click

        CaptureTime(DataGridViewTimesheet.CurrentRow.Index)

    End Sub


    ''' <summary>
    ''' 動画・音声から時間キャプチャ
    ''' </summary>
    Private Sub CaptureTime(ByVal row As Integer)

        Try

            'ビデオから時間の取得
            Dim time As Double = _videoForm.AxWindowsMediaPlayer1.Ctlcontrols.currentPosition

            '時間の編集
            Dim translatedTime As String = TlansrateTime(time)

            '最下行なら行追加
            If DataGridViewTimesheet.Rows.Count - 1 = DataGridViewTimesheet.CurrentRow.Index Then

                '「動画・音声から時間キャプチャボタン」フォーカスロストイベントを一時的に抑止
                '（DataGridViewTimesheet.BeginEditイベントで発生）
                RemoveHandler ButtonTimeCapture.Leave, AddressOf ButtonTimeCapture_Leave

                DataGridViewTimesheet.BeginEdit(False)
                DataGridViewTimesheet.NotifyCurrentCellDirty(True)
                DataGridViewTimesheet.EndEdit()

                '無効化していたハンドラを復活させる
                AddHandler ButtonTimeCapture.Leave, AddressOf ButtonTimeCapture_Leave

            End If

            'DataTableに時間を記録
            DataGridViewTimesheet(COL_VTIME, DataGridViewTimesheet.CurrentRow.Index).Value = translatedTime

            'CellEndEditイベントを一時的に無効化する
            '（理由）ここを通過する機会として２通りの方法がある。
            '1つは、グリッドで時間列にフォーカスを当てた状態でスペースキーが押下された場合。
            'この場合、下の行に移動(CurrentCell)を実行すると、CellEndEditが走る。
            '1つは、「動画・音声から時間キャプチャ」ボタンにフォーカスが当たった状態でスペースキーが押下された場合。
            'この場合、そもそもセル編集が始まってない為、CellEndEditが走らない。
            '両者の場合で動作を一様にするため、CellEndEditは抑止し、
            '統一してEditInputイベントで編集内容を調整する
            RemoveHandler DataGridViewTimesheet.CellEndEdit, AddressOf DataGridViewTimesheet_CellEndEdit

            '下の行に移動
            DataGridViewTimesheet.CurrentCell =
            DataGridViewTimesheet(COL_VTIME, DataGridViewTimesheet.CurrentRow.Index + 1)

            '無効化していたハンドラを復活させる
            AddHandler DataGridViewTimesheet.CellEndEdit, AddressOf DataGridViewTimesheet_CellEndEdit

            'キャプチャボタンへフォーカス
            ButtonTimeCapture.Focus()

        Catch ex As Exception

            TextBoxLog.Text = MakeInfoText("時間キャプチャ時に不測エラーが発生しました。", LogLevel.FATAL)

        End Try

    End Sub


    ''' <summary>
    ''' WMPから取得した時間をタイムシート時間形式にフォーマット
    ''' </summary>
    ''' <param name="time"></param>
    ''' <returns></returns>
    Private Function TlansrateTime(ByVal time As Double) As String

        Dim separateStr As String() = time.ToString.Split("."c)

        '0秒の場合
        If separateStr.Length = 1 Then
            Return "0"
        Else
            Return CDbl(separateStr(0)) * 10 + CDbl(separateStr(1).Substring(0, 1)) + _capture_delay
        End If

    End Function


    ''' <summary>
    ''' '「動画・音声から時間キャプチャボタン」フォーカスロストイベント
    ''' 時間キャプチャ操作を終えてから描画を一斉に更新し、時間キャプチャの動作を軽くする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonTimeCapture_Leave(sender As Object, e As EventArgs) Handles ButtonTimeCapture.Leave

        For row As Integer = 0 To DataGridViewTimesheet.Rows.Count - 2
            '時分秒
            InputVtimeHms(row)
            '色など
            DrawDataGridViewRow(row)
        Next

    End Sub


    ''' <summary>
    ''' パターン編集画面をモーダル表示する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonPatternSetting_Click(sender As Object, e As EventArgs) Handles ButtonPatternSetting.Click

        'パターン編集フォーム表示
        Dim form As New PatternEditor(_patternDataSet)
        PatternEditor.PatternEditorInstance = form
        form.StartPosition = FormStartPosition.CenterParent
        form.ShowDialog()

        '保存ボタンが押された場合、パターンDataSetを上書きし、パターンコンボボックスを更新する。
        If form.IsSaved = True Then
            _patternDataSet = form._editingPatternDataSet.Copy()
            LoadPatternToCombo()
        End If

        'フォームオブジェクト削除
        form.Dispose()

    End Sub


    ''' <summary>
    ''' コンボボックス更新
    ''' </summary>
    Private Sub LoadPatternToCombo()

        ComboBoxPattern1.Items.Clear()
        ComboBoxPattern2.Items.Clear()
        ComboBoxPattern3.Items.Clear()
        ComboBoxPattern4.Items.Clear()


        'コンボボックスへの追加
        ComboBoxPattern1.BeginUpdate()
        ComboBoxPattern2.BeginUpdate()
        ComboBoxPattern3.BeginUpdate()
        ComboBoxPattern4.BeginUpdate()


        '先頭行は空白
        ComboBoxPattern1.Items.Add("")
        ComboBoxPattern2.Items.Add("")
        ComboBoxPattern3.Items.Add("")
        ComboBoxPattern4.Items.Add("")

        For Each t As DataTable In _patternDataSet.Tables

            Dim fileName As String = t.TableName
            ComboBoxPattern1.Items.Add(fileName)
            ComboBoxPattern2.Items.Add(fileName)
            ComboBoxPattern3.Items.Add(fileName)
            ComboBoxPattern4.Items.Add(fileName)

        Next

        ComboBoxPattern1.EndUpdate()
        ComboBoxPattern2.EndUpdate()
        ComboBoxPattern3.EndUpdate()
        ComboBoxPattern4.EndUpdate()
    End Sub



    ''' <summary>
    ''' パターンを使用するチェックボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CheckBoxUsePattern_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxUsePattern.CheckedChanged
        If Me.CheckBoxUsePattern.CheckState = CheckState.Checked Then
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_CHECK).Value = True
        Else
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_CHECK).Value = False
        End If

        '行のアクティベイト、描画を更新
        DrawDataGridViewRow(_selectedRow)

        '詳細欄の表示を更新
        RefreshDetail(_selectedRow)

    End Sub


    ''' <summary>
    ''' パターン１コンボボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBoxPattern1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPattern1.SelectedIndexChanged
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_PATN1).Value = ComboBoxPattern1.Text
    End Sub

    ''' <summary>
    ''' パターン２コンボボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBoxPattern2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPattern2.SelectedIndexChanged
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_PATN2).Value = ComboBoxPattern2.Text
    End Sub

    ''' <summary>
    ''' パターン３コンボボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBoxPattern3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPattern3.SelectedIndexChanged
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_PATN3).Value = ComboBoxPattern3.Text
    End Sub

    ''' <summary>
    ''' パターン４コンボボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBoxPattern4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPattern4.SelectedIndexChanged
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_PATN4).Value = ComboBoxPattern4.Text
    End Sub


    ''' <summary>
    ''' パターン１速度補正トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern1dspd_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern1dspd.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_DSPD1).Value = TrackBarPattern1dspd.Value
        DisplayDSPD(TrackBarPattern1dspd)
    End Sub

    ''' <summary>
    ''' パターン２速度補正トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern2dspd_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern2dspd.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_DSPD2).Value = TrackBarPattern2dspd.Value
        DisplayDSPD(TrackBarPattern2dspd)
    End Sub

    ''' <summary>
    ''' パターン３速度補正トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern3dspd_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern3dspd.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_DSPD3).Value = TrackBarPattern3dspd.Value
        DisplayDSPD(TrackBarPattern3dspd)
    End Sub

    ''' <summary>
    ''' パターン４速度補正トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern4dspd_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern4dspd.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_DSPD4).Value = TrackBarPattern4dspd.Value
        DisplayDSPD(TrackBarPattern4dspd)
    End Sub


    ''' <summary>
    ''' 変更されたトラックバーの状態から、速度補正の値をラベルに表示する
    ''' </summary>
    ''' <param name="sender"></param>
    Private Sub DisplayDSPD(sender As TrackBar)
        Dim trackBar As TrackBar = Nothing
        Dim label As Label = Nothing
        Dim rate0 As String = String.Empty
        Dim rate1 As String = String.Empty
        Dim rate2 As String = String.Empty
        Dim rate3 As String = String.Empty
        Dim rate4 As String = String.Empty
        Dim rate5 As String = String.Empty
        Dim rate6 As String = String.Empty
        Dim rate7 As String = String.Empty
        Dim rate8 As String = String.Empty
        Select Case True
            Case sender Is TrackBarPattern1dspd
                trackBar = TrackBarPattern1dspd
                label = LabelPattern1dspd
                rate0 = _pattern1_dspd0
                rate1 = _pattern1_dspd1
                rate2 = _pattern1_dspd2
                rate3 = _pattern1_dspd3
                rate4 = _pattern1_dspd4
                rate5 = _pattern1_dspd5
                rate6 = _pattern1_dspd6
                rate7 = _pattern1_dspd7
                rate8 = _pattern1_dspd8
            Case sender Is TrackBarPattern2dspd
                trackBar = TrackBarPattern2dspd
                label = LabelPattern2dspd
                rate0 = _pattern2_dspd0
                rate1 = _pattern2_dspd1
                rate2 = _pattern2_dspd2
                rate3 = _pattern2_dspd3
                rate4 = _pattern2_dspd4
                rate5 = _pattern2_dspd5
                rate6 = _pattern2_dspd6
                rate7 = _pattern2_dspd7
                rate8 = _pattern2_dspd8
            Case sender Is TrackBarPattern3dspd
                trackBar = TrackBarPattern3dspd
                label = LabelPattern3dspd
                rate0 = _pattern3_dspd0
                rate1 = _pattern3_dspd1
                rate2 = _pattern3_dspd2
                rate3 = _pattern3_dspd3
                rate4 = _pattern3_dspd4
                rate5 = _pattern3_dspd5
                rate6 = _pattern3_dspd6
                rate7 = _pattern3_dspd7
                rate8 = _pattern3_dspd8
            Case sender Is TrackBarPattern4dspd
                trackBar = TrackBarPattern4dspd
                label = LabelPattern4dspd
                rate0 = _pattern4_dspd0
                rate1 = _pattern4_dspd1
                rate2 = _pattern4_dspd2
                rate3 = _pattern4_dspd3
                rate4 = _pattern4_dspd4
                rate5 = _pattern4_dspd5
                rate6 = _pattern4_dspd6
                rate7 = _pattern4_dspd7
                rate8 = _pattern4_dspd8
        End Select

        Select Case trackBar.Value
            Case 0
                label.Text = rate0
            Case 1
                label.Text = rate1
            Case 2
                label.Text = rate2
            Case 3
                label.Text = rate3
            Case 4
                label.Text = rate4
            Case 5
                label.Text = rate5
            Case 6
                label.Text = rate6
            Case 7
                label.Text = rate7
            Case 8
                label.Text = rate8
        End Select

    End Sub



    ''' <summary>
    ''' パターン１実行割合トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern1rate_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern1rate.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_RATE1).Value = TrackBarPattern1rate.Value
        DisplayRATE(TrackBarPattern1rate)
    End Sub

    ''' <summary>
    ''' パターン２１実行割合トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern2rate_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern2rate.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_RATE2).Value = TrackBarPattern2rate.Value
        DisplayRATE(TrackBarPattern2rate)
    End Sub

    ''' <summary>
    ''' パターン３実行割合トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern3rate_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern3rate.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_RATE3).Value = TrackBarPattern3rate.Value
        DisplayRATE(TrackBarPattern3rate)
    End Sub

    ''' <summary>
    ''' パターン４実行割合トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarPattern4rate_Scroll(sender As Object, e As EventArgs) Handles TrackBarPattern4rate.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_RATE4).Value = TrackBarPattern4rate.Value
        DisplayRATE(TrackBarPattern4rate)
    End Sub


    ''' <summary>
    ''' 変更されたトラックバーの状態から、実行割合の値をラベルに表示する
    ''' </summary>
    ''' <param name="sender"></param>
    Private Sub DisplayRATE(sender As TrackBar)
        Dim trackBar As TrackBar = Nothing
        Dim label As Label = Nothing
        Select Case True
            Case sender Is TrackBarPattern1rate
                trackBar = TrackBarPattern1rate
                label = LabelPattern1rate
            Case sender Is TrackBarPattern2rate
                trackBar = TrackBarPattern2rate
                label = LabelPattern2rate
            Case sender Is TrackBarPattern3rate
                trackBar = TrackBarPattern3rate
                label = LabelPattern3rate
            Case sender Is TrackBarPattern4rate
                trackBar = TrackBarPattern4rate
                label = LabelPattern4rate
        End Select

        Select Case trackBar.Value
            Case 0
                label.Text = _pattern_rate0
            Case 1
                label.Text = _pattern_rate1
            Case 2
                label.Text = _pattern_rate2
            Case 3
                label.Text = _pattern_rate3
            Case 4
                label.Text = _pattern_rate4
        End Select
    End Sub


    ''' <summary>
    ''' 実行モードラジオボタン→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RadioButtonRandom_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonRandom.CheckedChanged

        If RadioButtonOrder.Checked = True Then

            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_USERDM).Value = False

            '実行割合トラックバー　無効
            TrackBarPattern1rate.Enabled = False
            TrackBarPattern2rate.Enabled = False
            TrackBarPattern3rate.Enabled = False
            TrackBarPattern4rate.Enabled = False

        Else

            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_USERDM).Value = True

            '実行割合トラックバー　有効
            TrackBarPattern1rate.Enabled = True
            TrackBarPattern2rate.Enabled = True
            TrackBarPattern3rate.Enabled = True
            TrackBarPattern4rate.Enabled = True

        End If

    End Sub



    ''' <summary>
    ''' パターン間に休止をはさむチェックボックス→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CheckBoxUseInterval_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxUseInterval.CheckedChanged
        If Me.CheckBoxUseInterval.CheckState = CheckState.Checked Then
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_USEITV).Value = True
        Else
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_USEITV).Value = False
        End If
    End Sub


    ''' <summary>
    ''' 休止最小時間トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarItvMinTime_Scroll(sender As Object, e As EventArgs) Handles TrackBarItvMinTime.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_ITVMIN).Value = TrackBarItvMinTime.Value
        DisplayItvTime(TrackBarItvMinTime)

        '休止最大時間トラックバーよりも大きい値の場合、
        '休止最大時間トラックバーの値を休止最小時間の値に合わせる
        If TrackBarItvMinTime.Value > TrackBarItvMaxTime.Value Then
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_ITVMAX).Value = TrackBarItvMinTime.Value
            TrackBarItvMaxTime.Value = TrackBarItvMinTime.Value
            DisplayItvTime(TrackBarItvMaxTime)
        End If

        'パターン間に休止をはさむチェックボックス更新
        UpdateUseItvCheckboxByTrackbar()

    End Sub


    ''' <summary>
    ''' 休止最大時間トラックバー→グリッドへ反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TrackBarItvMaxTime_Scroll(sender As Object, e As EventArgs) Handles TrackBarItvMaxTime.Scroll
        DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_ITVMAX).Value = TrackBarItvMaxTime.Value
        DisplayItvTime(TrackBarItvMaxTime)

        '休止最小時間トラックバーよりも小さい値の場合、
        '休止最小時間トラックバーの値を休止最大時間の値に合わせる
        If TrackBarItvMinTime.Value > TrackBarItvMaxTime.Value Then
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_ITVMIN).Value = TrackBarItvMaxTime.Value
            TrackBarItvMinTime.Value = TrackBarItvMaxTime.Value
            DisplayItvTime(TrackBarItvMinTime)
        End If

        'パターン間に休止をはさむチェックボックス更新
        UpdateUseItvCheckboxByTrackbar()

    End Sub


    ''' <summary>
    ''' 変更されたトラックバーの状態から、休止時間の値をラベルに表示する
    ''' </summary>
    ''' <param name="sender"></param>
    Private Sub DisplayItvTime(sender As TrackBar)
        Dim trackBar As TrackBar = Nothing
        Dim label As Label = Nothing
        Select Case True
            Case sender Is TrackBarItvMinTime
                trackBar = TrackBarItvMinTime
                label = LabelMinTime
            Case sender Is TrackBarItvMaxTime
                trackBar = TrackBarItvMaxTime
                label = LabelMaxTime
        End Select

        Select Case trackBar.Value
            Case 0
                label.Text = _interval_time00
            Case 1
                label.Text = _interval_time01
            Case 2
                label.Text = _interval_time02
            Case 3
                label.Text = _interval_time03
            Case 4
                label.Text = _interval_time04
            Case 5
                label.Text = _interval_time05
            Case 6
                label.Text = _interval_time06
            Case 7
                label.Text = _interval_time07
            Case 8
                label.Text = _interval_time08
            Case 9
                label.Text = _interval_time09
            Case 10
                label.Text = _interval_time10
            Case 11
                label.Text = _interval_time11
            Case 12
                label.Text = _interval_time12
            Case 13
                label.Text = _interval_time13
            Case 14
                label.Text = _interval_time14
            Case 15
                label.Text = _interval_time15
            Case 16
                label.Text = _interval_time16
            Case 17
                label.Text = _interval_time17
            Case 18
                label.Text = _interval_time18
            Case 19
                label.Text = _interval_time19
            Case 20
                label.Text = _interval_time20
        End Select
    End Sub


    ''' <summary>
    ''' パターン間に休止をはさむチェックボックスを更新（トラックバーの変更後）
    ''' </summary>
    Private Sub UpdateUseItvCheckboxByTrackbar()

        'パターン間に休止をはさむチェックボックスがOFFの状態、
        'かつ休止最大時間が0秒以上の場合、
        'パターン間に休止をはさむチェックボックスをONにする。
        If Me.CheckBoxUseInterval.CheckState = CheckState.Unchecked AndAlso
           TrackBarItvMaxTime.Value > 0 Then
            Me.CheckBoxUseInterval.CheckState = CheckState.Checked
            DataGridViewTimesheet.Rows(_selectedRow).Cells(COL_USEITV).Value = True
        End If

    End Sub

    Private Sub 全体を01秒進めるToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 全体を01秒進めるToolStripMenuItem.Click

        For i = 0 To DataGridViewTimesheet.Rows.Count - 2
            Dim time As String
            time = DataGridViewTimesheet.Rows(i).Cells(COL_VTIME).Value
            If time IsNot Nothing AndAlso Not time.Equals(String.Empty) Then
                DataGridViewTimesheet.Rows(i).Cells(COL_VTIME).Value = (CInt(time) + 1).ToString()
                InputVtimeHms(i)
            End If
        Next

    End Sub

    Private Sub 全体を01秒遅らせるToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 全体を01秒遅らせるToolStripMenuItem.Click

        For i = 0 To DataGridViewTimesheet.Rows.Count - 2
            Dim time As String
            time = DataGridViewTimesheet.Rows(i).Cells(COL_VTIME).Value
            If time IsNot Nothing AndAlso Not time.Equals(String.Empty) AndAlso time <> "0" Then
                DataGridViewTimesheet.Rows(i).Cells(COL_VTIME).Value = (CInt(time) - 1).ToString()
                InputVtimeHms(i)
            End If
        Next

    End Sub

    Private Sub 末尾に100行追加ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 末尾に100行追加ToolStripMenuItem.Click
        Add100Rows(False)
    End Sub

    Private Sub 末尾に100行追加パターンありToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 末尾に100行追加パターンありToolStripMenuItem.Click
        Add100Rows(True)
    End Sub

    Private Sub Add100Rows(ByVal usePattern As Boolean)
        Dim maxRow As Integer = DataGridViewTimesheet.Rows.Count

        For i = 0 To 99
            DataGridViewTimesheet.Rows.Add()
        Next

        For i = 0 To 99
            DefineRow(DataGridViewTimesheet.Rows(i + maxRow - 1))
        Next

        If usePattern Then
            For i = 0 To 99
                DataGridViewTimesheet.Rows(i + maxRow - 1).Cells(COL_CHECK).Value = True
            Next
        End If

        '描画を更新
        For i = 0 To 99
            DrawDataGridViewRow(i + maxRow - 1)
        Next
    End Sub

    Private Sub X050ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X050ToolStripMenuItem.Click
        ChangeContentsPlayRate(0.5, X050ToolStripMenuItem)
    End Sub

    Private Sub X075ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X075ToolStripMenuItem.Click
        ChangeContentsPlayRate(0.75, X075ToolStripMenuItem)
    End Sub

    Private Sub X100ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X100ToolStripMenuItem.Click
        ChangeContentsPlayRate(1.0, X100ToolStripMenuItem)
    End Sub

    Private Sub X125ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X125ToolStripMenuItem.Click
        ChangeContentsPlayRate(1.25, X125ToolStripMenuItem)
    End Sub

    Private Sub X150ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X150ToolStripMenuItem.Click
        ChangeContentsPlayRate(1.5, X150ToolStripMenuItem)
    End Sub

    Private Sub X200ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles X200ToolStripMenuItem.Click
        ChangeContentsPlayRate(2.0, X200ToolStripMenuItem)

    End Sub

#End Region

    Private Sub ChangeContentsPlayRate(ByVal rate As Double,
                                       ByVal tsmi As ToolStripMenuItem)
        If Not _videoForm Is Nothing Then
            _videoForm.AxWindowsMediaPlayer1.settings.rate = rate
            EditPlayRateCheckState(tsmi)
        End If
    End Sub

    Private Sub EditPlayRateCheckState(ByVal tsmi As ToolStripMenuItem)
        X050ToolStripMenuItem.CheckState = CheckState.Unchecked
        X075ToolStripMenuItem.CheckState = CheckState.Unchecked
        X100ToolStripMenuItem.CheckState = CheckState.Unchecked
        X125ToolStripMenuItem.CheckState = CheckState.Unchecked
        X150ToolStripMenuItem.CheckState = CheckState.Unchecked
        X200ToolStripMenuItem.CheckState = CheckState.Unchecked

        tsmi.CheckState = CheckState.Indeterminate

    End Sub
    Private Sub back5sToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles back5sToolStripMenuItem.Click
        If Not _videoForm Is Nothing Then
            With _videoForm.AxWindowsMediaPlayer1.Ctlcontrols
                .currentPosition = Math.Max(0, .currentPosition - 5)
            End With
        End If
    End Sub

    Private Sub skip5sToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles skip5sToolStripMenuItem.Click
        If Not _videoForm Is Nothing Then
            With _videoForm.AxWindowsMediaPlayer1.Ctlcontrols
                .currentPosition = .currentPosition + 5
            End With
        End If
    End Sub
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) _
    Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) _
    Handles Me.DragDrop

        Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())

        '定義ファイル
        For Each filePath In files
            If IO.Path.GetExtension(filePath).ToLower() = "." & EXTENSION Then
                OpenFile(filePath)
                Exit Sub
            End If
        Next

        'メディアファイル
        Dim allowedVideoExtensions As String() = {
            ".mp3", ".wav", ".wma",
            ".mp4", ".wmv", ".avi",
            ".mpg", ".mpeg", ".asf",
            ".m4a", ".aac"
        }
        For Each filePath In files
            If allowedVideoExtensions.Contains(IO.Path.GetExtension(filePath).ToLower()) Then
                UpdateVideoForm(filePath)
                Exit Sub
            End If
        Next

    End Sub

    Private Sub TSVエクスポートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TSVエクスポートToolStripMenuItem.Click

        Dim fileName As String
        If Not String.IsNullOrEmpty(_editing_fileName) Then
            fileName = Path.GetFileNameWithoutExtension(_editing_fileName)
        Else
            fileName = "新しいファイル"
        End If

        '名前をつけて保存
        Dim sfd As New SaveFileDialog()

        sfd.FileName = fileName & ".tsv"
        sfd.InitialDirectory = _dialog_path_timesheet_def
        sfd.Filter = "TSVファイル (*.tsv)|*.tsv"
        sfd.Title = "出力先を指定してください"

        If sfd.ShowDialog() = DialogResult.OK Then
            ExportDataGridViewToTsv(sfd.FileName)
        End If

    End Sub

    Private Sub ExportDataGridViewToTsv(filePath As String)

        Try
            Using sw As New StreamWriter(filePath, False, New UTF8Encoding(True))

                ' ヘッダー
                Dim headers = DataGridViewTimesheet.Columns.
                          Cast(Of DataGridViewColumn)().
                          Select(Function(c) c.HeaderText)

                sw.WriteLine(String.Join(vbTab, headers))

                ' データ
                For Each row As DataGridViewRow In DataGridViewTimesheet.Rows
                    If Not row.IsNewRow Then
                        Dim cellValues As New List(Of String)

                        For Each cell As DataGridViewCell In row.Cells
                            If cell.Value IsNot Nothing Then
                                cellValues.Add(cell.Value.ToString())
                            Else
                                cellValues.Add("")
                            End If
                        Next

                        sw.WriteLine(String.Join(vbTab, cellValues))
                    End If
                Next
            End Using

            TextBoxLog.Text = MakeInfoText("TSVエクスポートに成功しました。" & "(" & DateTime.Now.ToString("HH:mm:ss") & ")", LogLevel.INFO)

        Catch ex As Exception

            TextBoxLog.Text = MakeInfoText("TSVエクスポート中に不測エラーが発生しました。", LogLevel.FATAL)

        End Try

    End Sub

    Private Sub TSVインポートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TSVインポートToolStripMenuItem.Click

        Dim ofd As New OpenFileDialog()
        ofd.InitialDirectory = _dialog_path_timesheet_def
        ofd.Filter = "TSVファイル (*.tsv)|*.tsv"
        ofd.Title = "インポートするTSVを選択してください"

        If ofd.ShowDialog() = DialogResult.OK Then
            ImportTsvToDataGridView(ofd.FileName)
        End If

    End Sub

    Private Sub ImportTsvToDataGridView(filePath As String)

        Dim errList As New List(Of String)

        Try
            DataGridViewTimesheet.Rows.Clear()

            Dim lines = File.ReadAllLines(filePath)

            ' 1行目がヘッダーならスキップ
            For i As Integer = 1 To lines.Length - 1

                Dim cols = lines(i).Split(vbTab)

                ' 必要列数チェック（23列以上）
                If cols.Length < 23 Then
                    errList.Add(MakeInfoText($"{i + 1}行目 列数が不足しています。", LogLevel.ERR))
                    Continue For
                End If

                Try

                    Dim time As Integer? = ParseNullableInt(cols(1), "時間", i)
                    Dim direction As Integer? = ParseNullableRangeInt(cols(3), 0, 1, "方向", i)
                    Dim speed As Integer? = ParseNullableRangeInt(cols(4), 0, 100, "速度", i)

                    Dim pattern As Boolean = ParseBool(cols(5), "パターン", i)
                    Dim pattern1 As String = cols(6)
                    Dim pattern2 As String = cols(7)
                    Dim pattern3 As String = cols(8)
                    Dim pattern4 As String = cols(9)

                    Dim comment As String = cols(10)

                    Dim dspd1 As Integer = ParseRangeInt(cols(11), 0, 8, "速度補正1", i)
                    Dim dspd2 As Integer = ParseRangeInt(cols(12), 0, 8, "速度補正2", i)
                    Dim dspd3 As Integer = ParseRangeInt(cols(13), 0, 8, "速度補正3", i)
                    Dim dspd4 As Integer = ParseRangeInt(cols(14), 0, 8, "速度補正4", i)

                    Dim rate1 As Integer = ParseRangeInt(cols(15), 0, 4, "実行比率1", i)
                    Dim rate2 As Integer = ParseRangeInt(cols(16), 0, 4, "実行比率2", i)
                    Dim rate3 As Integer = ParseRangeInt(cols(17), 0, 4, "実行比率3", i)
                    Dim rate4 As Integer = ParseRangeInt(cols(18), 0, 4, "実行比率4", i)


                    Dim userdm As Boolean = ParseBool(cols(19), "実行モード", i)
                    Dim useitv As Boolean = ParseBool(cols(20), "パターン間に休止をはさむ", i)

                    Dim itvmin As Integer = ParseRangeInt(cols(21), 0, 20, "休止最小時間", i)
                    Dim itvmax As Integer = ParseRangeInt(cols(22), 0, 20, "休止最大時間", i)

                    If itvmin > itvmax Then
                        Throw New Exception($"{i + 1}行目 itvmin > itvmax")
                    End If

                    ' DataGridViewへ追加
                    DataGridViewTimesheet.Rows.Add(
                        Nothing,
                        If(time.HasValue, time.Value, ""),
                        Nothing,
                        If(direction.HasValue, direction.Value, ""),
                        If(speed.HasValue, speed.Value, ""),
                        pattern,
                        pattern1,
                        pattern2,
                        pattern3,
                        pattern4,
                        comment,
                        dspd1,
                        dspd2,
                        dspd3,
                        dspd4,
                        rate1,
                        rate2,
                        rate3,
                        rate4,
                        userdm,
                        useitv,
                        itvmin,
                        itvmax)

                Catch ex As Exception

                    errList.Add(MakeInfoText(ex.Message, LogLevel.ERR))
                    Continue For

                End Try

            Next

            '描画を更新
            For i As Integer = 0 To DataGridViewTimesheet.Rows.Count - 2
                DrawDataGridViewRow(i)
                InputVtimeHms(i)
            Next

            Dim outputText As String = ""
            If errList.Count > 0 Then
                outputText = MakeInfoText("エラー行があり、TSVファイルの一部の行取り込みをスキップします。", LogLevel.INFO) & vbCrLf &
                    String.Join(vbCrLf, errList)
            End If
            TextBoxLog.Text =
                    MakeInfoText("インポートが完了しました。" & "(" & DateTime.Now.ToString("HH:mm:ss") & ")", LogLevel.INFO) & vbCrLf &
                    outputText

        Catch ex As Exception

            TextBoxLog.Text = MakeInfoText("TSVインポート中に不測エラーが発生しました。", LogLevel.FATAL)

        End Try

    End Sub

    Private Function ParseInt(value As String, colName As String, rowIndex As Integer) As Integer
        Dim result As Integer
        If Not Integer.TryParse(value, result) Then
            Throw New Exception($"{rowIndex + 1}行目 列[{colName}] が整数ではありません: {value}")
        End If
        Return result
    End Function

    Private Function ParseNullableInt(value As String, colName As String, rowIndex As Integer) As Integer?

        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        Dim result As Integer
        If Not Integer.TryParse(value, result) Then
            Throw New Exception($"{rowIndex + 1}行目 列[{colName}] が整数ではありません: {value}")
        End If

        Return result
    End Function

    Private Function ParseNullableRangeInt(value As String, min As Integer, max As Integer, colName As String, rowIndex As Integer) As Integer?

        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        Dim result = ParseInt(value, colName, rowIndex)

        If result < min OrElse result > max Then
            Throw New Exception($"{rowIndex + 1}行目 列[{colName}] が範囲外({min}-{max}): {value}")
        End If

        Return result

    End Function
    Private Function ParseRangeInt(value As String, min As Integer, max As Integer, colName As String, rowIndex As Integer) As Integer
        Dim result = ParseInt(value, colName, rowIndex)
        If result < min OrElse result > max Then
            Throw New Exception($"{rowIndex + 1}行目 列[{colName}] が範囲外({min}-{max}): {value}")
        End If
        Return result
    End Function

    Private Function ParseBool(value As String, colName As String, rowIndex As Integer) As Boolean
        Dim result As Boolean
        If Not Boolean.TryParse(value, result) Then
            Throw New Exception($"{rowIndex + 1}行目 列[{colName}] がBooleanではありません: {value}")
        End If
        Return result
    End Function

    Private Sub DataGridViewTimesheet_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTimesheet.CellValueChanged
        If OxyPlotView.Model Is Nothing Then
            Return
        End If
        If _IsLoading Then
            Return
        End If

        SetPlot()
    End Sub

    Private Sub SetPlot()
        Dim model As New PlotModel
        model.Axes.Add(New LinearAxis With {
                       .Position = AxisPosition.Left,
                       .Minimum = -100,
                       .Maximum = 100,
                       .AbsoluteMinimum = -100,
                       .AbsoluteMaximum = 100,
                       .MajorGridlineColor = OxyColors.Silver,
                       .MajorGridlineStyle = LineStyle.Solid,
                       .MinorGridlineColor = OxyColors.Silver,
                       .MinorGridlineStyle = LineStyle.Dot,
                       .IsZoomEnabled = False
        })
        model.Axes.Add(New LinearAxis With {.Position = AxisPosition.Bottom, .AbsoluteMinimum = 0})

        Dim timeSheet = GenerateTimeSheetRowsFromDataGridView()
        If timeSheet Is Nothing Then
            Return
        End If

        Dim series = New LineSeries
        series.Points.Add(New DataPoint(0, 0))
        Dim prevSpd As Integer = 0

        For Each row In timeSheet
            Dim spd As Integer
            If row.Dir = True Then
                spd = row.Spd
            Else
                spd = -row.Spd
            End If

            series.Points.Add(New DataPoint(row.Time, prevSpd))
            series.Points.Add(New DataPoint(row.Time, spd))
            prevSpd = spd
        Next

        model.Series.Add(series)
        OxyPlotView.Model = model
    End Sub
End Class


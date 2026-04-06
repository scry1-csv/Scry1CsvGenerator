<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Scry1CsvGenerator
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Scry1CsvGenerator))
        Me.DataGridViewTimesheet = New System.Windows.Forms.DataGridView()
        Me.vtimehms = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vtime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vtimediff = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vdir = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vspd = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.check = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.patn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.patn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.patn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.patn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vcmt = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dspd1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dspd2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dspd3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dspd4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rate1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rate2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rate3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rate4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.userdm = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.useitv = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.itvmin = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.itvmax = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ButtonPatternSetting = New System.Windows.Forms.Button()
        Me.CheckBoxUsePattern = New System.Windows.Forms.CheckBox()
        Me.ComboBoxPattern1 = New System.Windows.Forms.ComboBox()
        Me.TrackBarPattern1dspd = New System.Windows.Forms.TrackBar()
        Me.TrackBarPattern1rate = New System.Windows.Forms.TrackBar()
        Me.LabelPattern1dspd = New System.Windows.Forms.Label()
        Me.LabelPattern1rate = New System.Windows.Forms.Label()
        Me.LabelPatternDetail = New System.Windows.Forms.Label()
        Me.LabelPattern1 = New System.Windows.Forms.Label()
        Me.LabelPattern2 = New System.Windows.Forms.Label()
        Me.LabelPattern2rate = New System.Windows.Forms.Label()
        Me.LabelPattern2dspd = New System.Windows.Forms.Label()
        Me.TrackBarPattern2rate = New System.Windows.Forms.TrackBar()
        Me.TrackBarPattern2dspd = New System.Windows.Forms.TrackBar()
        Me.ComboBoxPattern2 = New System.Windows.Forms.ComboBox()
        Me.LabelPattern3 = New System.Windows.Forms.Label()
        Me.LabelPattern4 = New System.Windows.Forms.Label()
        Me.LabelPattern3rate = New System.Windows.Forms.Label()
        Me.LabelPattern3dspd = New System.Windows.Forms.Label()
        Me.TrackBarPattern3rate = New System.Windows.Forms.TrackBar()
        Me.TrackBarPattern3dspd = New System.Windows.Forms.TrackBar()
        Me.ComboBoxPattern3 = New System.Windows.Forms.ComboBox()
        Me.LabelPattern4rate = New System.Windows.Forms.Label()
        Me.LabelPattern4dspd = New System.Windows.Forms.Label()
        Me.TrackBarPattern4rate = New System.Windows.Forms.TrackBar()
        Me.TrackBarPattern4dspd = New System.Windows.Forms.TrackBar()
        Me.ComboBoxPattern4 = New System.Windows.Forms.ComboBox()
        Me.LabelMode = New System.Windows.Forms.Label()
        Me.RadioButtonOrder = New System.Windows.Forms.RadioButton()
        Me.RadioButtonRandom = New System.Windows.Forms.RadioButton()
        Me.CheckBoxUseInterval = New System.Windows.Forms.CheckBox()
        Me.LabelMin = New System.Windows.Forms.Label()
        Me.LabelMax = New System.Windows.Forms.Label()
        Me.TextBoxLog = New System.Windows.Forms.TextBox()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItemFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemOverwriteSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSVエクスポートToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSVインポートToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItemClear = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemEnd = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemAddRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemDeleteRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.一括編集ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.末尾に100行追加ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.末尾に100行追加パターンありToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItemDelay = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus5s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus4s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus3s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus2s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus1s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_plus0s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_minus1s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_minus2s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_minus3s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_minus4s = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_minus5s = New System.Windows.Forms.ToolStripMenuItem()
        Me.全体を01秒進めるToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.全体を01秒遅らせるToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.コンテンツ再生速度ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X050ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X075ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X100ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X125ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X150ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.X200ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.back5sToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.skip5sToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemContentsPlay = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItemCsvGenerate = New System.Windows.Forms.ToolStripMenuItem()
        Me.ButtonTimeCapture = New System.Windows.Forms.Button()
        Me.LabelLog = New System.Windows.Forms.Label()
        Me.TrackBarItvMinTime = New System.Windows.Forms.TrackBar()
        Me.TrackBarItvMaxTime = New System.Windows.Forms.TrackBar()
        Me.LabelMinTime = New System.Windows.Forms.Label()
        Me.LabelMaxTime = New System.Windows.Forms.Label()
        Me.LabelSec1 = New System.Windows.Forms.Label()
        Me.LabelSec2 = New System.Windows.Forms.Label()
        Me.OxyPlotView = New OxyPlot.WindowsForms.PlotView()
        CType(Me.DataGridViewTimesheet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern1dspd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern1rate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern2rate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern2dspd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern3rate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern3dspd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern4rate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarPattern4dspd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip.SuspendLayout()
        CType(Me.TrackBarItvMinTime, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBarItvMaxTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewTimesheet
        '
        Me.DataGridViewTimesheet.AllowUserToResizeRows = False
        Me.DataGridViewTimesheet.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridViewTimesheet.ColumnHeadersHeight = 58
        Me.DataGridViewTimesheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridViewTimesheet.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.vtimehms, Me.vtime, Me.vtimediff, Me.vdir, Me.vspd, Me.check, Me.patn1, Me.patn2, Me.patn3, Me.patn4, Me.vcmt, Me.dspd1, Me.dspd2, Me.dspd3, Me.dspd4, Me.rate1, Me.rate2, Me.rate3, Me.rate4, Me.userdm, Me.useitv, Me.itvmin, Me.itvmax})
        Me.DataGridViewTimesheet.Location = New System.Drawing.Point(13, 33)
        Me.DataGridViewTimesheet.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridViewTimesheet.MultiSelect = False
        Me.DataGridViewTimesheet.Name = "DataGridViewTimesheet"
        Me.DataGridViewTimesheet.RowHeadersWidth = 102
        Me.DataGridViewTimesheet.RowTemplate.Height = 21
        Me.DataGridViewTimesheet.Size = New System.Drawing.Size(880, 714)
        Me.DataGridViewTimesheet.TabIndex = 0
        '
        'vtimehms
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        Me.vtimehms.DefaultCellStyle = DataGridViewCellStyle1
        Me.vtimehms.Frozen = True
        Me.vtimehms.HeaderText = ""
        Me.vtimehms.MinimumWidth = 12
        Me.vtimehms.Name = "vtimehms"
        Me.vtimehms.ReadOnly = True
        Me.vtimehms.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vtimehms.Width = 60
        '
        'vtime
        '
        DataGridViewCellStyle2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.vtime.DefaultCellStyle = DataGridViewCellStyle2
        Me.vtime.Frozen = True
        Me.vtime.HeaderText = "時間"
        Me.vtime.MaxInputLength = 6
        Me.vtime.MinimumWidth = 12
        Me.vtime.Name = "vtime"
        Me.vtime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vtime.ToolTipText = "時間 [10=1秒]"
        Me.vtime.Width = 65
        '
        'vtimediff
        '
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black
        Me.vtimediff.DefaultCellStyle = DataGridViewCellStyle3
        Me.vtimediff.Frozen = True
        Me.vtimediff.HeaderText = "動作時間"
        Me.vtimediff.MinimumWidth = 12
        Me.vtimediff.Name = "vtimediff"
        Me.vtimediff.ReadOnly = True
        Me.vtimediff.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vtimediff.ToolTipText = "差[10=1秒]"
        Me.vtimediff.Width = 80
        '
        'vdir
        '
        DataGridViewCellStyle4.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.vdir.DefaultCellStyle = DataGridViewCellStyle4
        Me.vdir.Frozen = True
        Me.vdir.HeaderText = "方向"
        Me.vdir.MaxInputLength = 1
        Me.vdir.MinimumWidth = 12
        Me.vdir.Name = "vdir"
        Me.vdir.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vdir.ToolTipText = "回転方向 [0=右回転, 1=左回転]"
        Me.vdir.Width = 65
        '
        'vspd
        '
        DataGridViewCellStyle5.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.vspd.DefaultCellStyle = DataGridViewCellStyle5
        Me.vspd.Frozen = True
        Me.vspd.HeaderText = "速度"
        Me.vspd.MaxInputLength = 3
        Me.vspd.MinimumWidth = 12
        Me.vspd.Name = "vspd"
        Me.vspd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vspd.ToolTipText = "回転速度 [推奨値・・・A10サイクロンSA：40以上　U.F.O.SA：25以上]"
        Me.vspd.Width = 65
        '
        'check
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        DataGridViewCellStyle6.NullValue = False
        Me.check.DefaultCellStyle = DataGridViewCellStyle6
        Me.check.Frozen = True
        Me.check.HeaderText = "パターン"
        Me.check.MinimumWidth = 12
        Me.check.Name = "check"
        Me.check.Width = 65
        '
        'patn1
        '
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.patn1.DefaultCellStyle = DataGridViewCellStyle7
        Me.patn1.HeaderText = "パターン1"
        Me.patn1.MaxInputLength = 100
        Me.patn1.MinimumWidth = 12
        Me.patn1.Name = "patn1"
        Me.patn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.patn1.Width = 80
        '
        'patn2
        '
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.patn2.DefaultCellStyle = DataGridViewCellStyle8
        Me.patn2.HeaderText = "パターン2"
        Me.patn2.MaxInputLength = 100
        Me.patn2.MinimumWidth = 12
        Me.patn2.Name = "patn2"
        Me.patn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.patn2.Width = 80
        '
        'patn3
        '
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.patn3.DefaultCellStyle = DataGridViewCellStyle9
        Me.patn3.HeaderText = "パターン3"
        Me.patn3.MaxInputLength = 100
        Me.patn3.MinimumWidth = 12
        Me.patn3.Name = "patn3"
        Me.patn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.patn3.Width = 80
        '
        'patn4
        '
        DataGridViewCellStyle10.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.patn4.DefaultCellStyle = DataGridViewCellStyle10
        Me.patn4.HeaderText = "パターン4"
        Me.patn4.MaxInputLength = 100
        Me.patn4.MinimumWidth = 12
        Me.patn4.Name = "patn4"
        Me.patn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.patn4.Width = 80
        '
        'vcmt
        '
        DataGridViewCellStyle11.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(236, Byte), Integer))
        DataGridViewCellStyle11.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.vcmt.DefaultCellStyle = DataGridViewCellStyle11
        Me.vcmt.HeaderText = "コメント"
        Me.vcmt.MaxInputLength = 100
        Me.vcmt.MinimumWidth = 12
        Me.vcmt.Name = "vcmt"
        Me.vcmt.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vcmt.Width = 285
        '
        'dspd1
        '
        Me.dspd1.HeaderText = "dspd1"
        Me.dspd1.MaxInputLength = 3
        Me.dspd1.MinimumWidth = 12
        Me.dspd1.Name = "dspd1"
        Me.dspd1.Visible = False
        Me.dspd1.Width = 20
        '
        'dspd2
        '
        Me.dspd2.HeaderText = "dspd2"
        Me.dspd2.MaxInputLength = 3
        Me.dspd2.MinimumWidth = 12
        Me.dspd2.Name = "dspd2"
        Me.dspd2.Visible = False
        Me.dspd2.Width = 20
        '
        'dspd3
        '
        Me.dspd3.HeaderText = "dspd3"
        Me.dspd3.MaxInputLength = 3
        Me.dspd3.MinimumWidth = 12
        Me.dspd3.Name = "dspd3"
        Me.dspd3.Visible = False
        Me.dspd3.Width = 20
        '
        'dspd4
        '
        Me.dspd4.HeaderText = "dspd4"
        Me.dspd4.MaxInputLength = 3
        Me.dspd4.MinimumWidth = 12
        Me.dspd4.Name = "dspd4"
        Me.dspd4.Visible = False
        Me.dspd4.Width = 20
        '
        'rate1
        '
        Me.rate1.HeaderText = "rate1"
        Me.rate1.MaxInputLength = 1
        Me.rate1.MinimumWidth = 12
        Me.rate1.Name = "rate1"
        Me.rate1.Visible = False
        Me.rate1.Width = 20
        '
        'rate2
        '
        Me.rate2.HeaderText = "rate2"
        Me.rate2.MaxInputLength = 1
        Me.rate2.MinimumWidth = 12
        Me.rate2.Name = "rate2"
        Me.rate2.Visible = False
        Me.rate2.Width = 20
        '
        'rate3
        '
        Me.rate3.HeaderText = "rate3"
        Me.rate3.MaxInputLength = 1
        Me.rate3.MinimumWidth = 12
        Me.rate3.Name = "rate3"
        Me.rate3.Visible = False
        Me.rate3.Width = 20
        '
        'rate4
        '
        Me.rate4.HeaderText = "rate4"
        Me.rate4.MaxInputLength = 1
        Me.rate4.MinimumWidth = 12
        Me.rate4.Name = "rate4"
        Me.rate4.Visible = False
        Me.rate4.Width = 20
        '
        'userdm
        '
        Me.userdm.HeaderText = "userdm"
        Me.userdm.MinimumWidth = 12
        Me.userdm.Name = "userdm"
        Me.userdm.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.userdm.Visible = False
        Me.userdm.Width = 20
        '
        'useitv
        '
        Me.useitv.HeaderText = "useitv"
        Me.useitv.MinimumWidth = 12
        Me.useitv.Name = "useitv"
        Me.useitv.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.useitv.Visible = False
        Me.useitv.Width = 20
        '
        'itvmin
        '
        Me.itvmin.HeaderText = "itvmin"
        Me.itvmin.MaxInputLength = 4
        Me.itvmin.MinimumWidth = 12
        Me.itvmin.Name = "itvmin"
        Me.itvmin.Visible = False
        Me.itvmin.Width = 20
        '
        'itvmax
        '
        Me.itvmax.HeaderText = "itvmax"
        Me.itvmax.MaxInputLength = 4
        Me.itvmax.MinimumWidth = 12
        Me.itvmax.Name = "itvmax"
        Me.itvmax.Visible = False
        Me.itvmax.Width = 20
        '
        'ButtonPatternSetting
        '
        Me.ButtonPatternSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonPatternSetting.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonPatternSetting.Location = New System.Drawing.Point(1183, 50)
        Me.ButtonPatternSetting.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonPatternSetting.Name = "ButtonPatternSetting"
        Me.ButtonPatternSetting.Size = New System.Drawing.Size(230, 31)
        Me.ButtonPatternSetting.TabIndex = 1
        Me.ButtonPatternSetting.Text = "パターンを編集する"
        Me.ButtonPatternSetting.UseVisualStyleBackColor = True
        '
        'CheckBoxUsePattern
        '
        Me.CheckBoxUsePattern.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxUsePattern.AutoSize = True
        Me.CheckBoxUsePattern.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.CheckBoxUsePattern.Location = New System.Drawing.Point(705, 95)
        Me.CheckBoxUsePattern.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxUsePattern.Name = "CheckBoxUsePattern"
        Me.CheckBoxUsePattern.Size = New System.Drawing.Size(353, 44)
        Me.CheckBoxUsePattern.TabIndex = 2
        Me.CheckBoxUsePattern.Text = "パターンを使用する"
        Me.CheckBoxUsePattern.UseVisualStyleBackColor = True
        '
        'ComboBoxPattern1
        '
        Me.ComboBoxPattern1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxPattern1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPattern1.FormattingEnabled = True
        Me.ComboBoxPattern1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ComboBoxPattern1.Location = New System.Drawing.Point(932, 123)
        Me.ComboBoxPattern1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBoxPattern1.Name = "ComboBoxPattern1"
        Me.ComboBoxPattern1.Size = New System.Drawing.Size(204, 48)
        Me.ComboBoxPattern1.TabIndex = 3
        '
        'TrackBarPattern1dspd
        '
        Me.TrackBarPattern1dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern1dspd.Location = New System.Drawing.Point(1169, 127)
        Me.TrackBarPattern1dspd.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern1dspd.Maximum = 8
        Me.TrackBarPattern1dspd.Name = "TrackBarPattern1dspd"
        Me.TrackBarPattern1dspd.Size = New System.Drawing.Size(136, 114)
        Me.TrackBarPattern1dspd.TabIndex = 4
        Me.TrackBarPattern1dspd.Value = 4
        '
        'TrackBarPattern1rate
        '
        Me.TrackBarPattern1rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern1rate.Location = New System.Drawing.Point(1333, 127)
        Me.TrackBarPattern1rate.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern1rate.Maximum = 4
        Me.TrackBarPattern1rate.Name = "TrackBarPattern1rate"
        Me.TrackBarPattern1rate.Size = New System.Drawing.Size(80, 114)
        Me.TrackBarPattern1rate.TabIndex = 5
        Me.TrackBarPattern1rate.Value = 1
        '
        'LabelPattern1dspd
        '
        Me.LabelPattern1dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern1dspd.AutoSize = True
        Me.LabelPattern1dspd.Location = New System.Drawing.Point(1144, 127)
        Me.LabelPattern1dspd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern1dspd.Name = "LabelPattern1dspd"
        Me.LabelPattern1dspd.Size = New System.Drawing.Size(115, 40)
        Me.LabelPattern1dspd.TabIndex = 6
        Me.LabelPattern1dspd.Text = "dspd1"
        Me.LabelPattern1dspd.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPattern1rate
        '
        Me.LabelPattern1rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern1rate.AutoSize = True
        Me.LabelPattern1rate.Location = New System.Drawing.Point(1321, 127)
        Me.LabelPattern1rate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern1rate.Name = "LabelPattern1rate"
        Me.LabelPattern1rate.Size = New System.Drawing.Size(104, 40)
        Me.LabelPattern1rate.TabIndex = 7
        Me.LabelPattern1rate.Text = "rate1"
        Me.LabelPattern1rate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPatternDetail
        '
        Me.LabelPatternDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPatternDetail.AutoSize = True
        Me.LabelPatternDetail.Location = New System.Drawing.Point(1169, 99)
        Me.LabelPatternDetail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPatternDetail.Name = "LabelPatternDetail"
        Me.LabelPatternDetail.Size = New System.Drawing.Size(580, 40)
        Me.LabelPatternDetail.TabIndex = 8
        Me.LabelPatternDetail.Text = "速度補正　　　　　　　　　実行比率"
        '
        'LabelPattern1
        '
        Me.LabelPattern1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern1.AutoSize = True
        Me.LabelPattern1.Location = New System.Drawing.Point(915, 127)
        Me.LabelPattern1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern1.Name = "LabelPattern1"
        Me.LabelPattern1.Size = New System.Drawing.Size(37, 40)
        Me.LabelPattern1.TabIndex = 9
        Me.LabelPattern1.Text = "1"
        '
        'LabelPattern2
        '
        Me.LabelPattern2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern2.AutoSize = True
        Me.LabelPattern2.Location = New System.Drawing.Point(915, 169)
        Me.LabelPattern2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern2.Name = "LabelPattern2"
        Me.LabelPattern2.Size = New System.Drawing.Size(37, 40)
        Me.LabelPattern2.TabIndex = 15
        Me.LabelPattern2.Text = "2"
        '
        'LabelPattern2rate
        '
        Me.LabelPattern2rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern2rate.AutoSize = True
        Me.LabelPattern2rate.Location = New System.Drawing.Point(1321, 170)
        Me.LabelPattern2rate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern2rate.Name = "LabelPattern2rate"
        Me.LabelPattern2rate.Size = New System.Drawing.Size(104, 40)
        Me.LabelPattern2rate.TabIndex = 14
        Me.LabelPattern2rate.Text = "rate2"
        Me.LabelPattern2rate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPattern2dspd
        '
        Me.LabelPattern2dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern2dspd.AutoSize = True
        Me.LabelPattern2dspd.Location = New System.Drawing.Point(1144, 170)
        Me.LabelPattern2dspd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern2dspd.Name = "LabelPattern2dspd"
        Me.LabelPattern2dspd.Size = New System.Drawing.Size(115, 40)
        Me.LabelPattern2dspd.TabIndex = 13
        Me.LabelPattern2dspd.Text = "dspd2"
        Me.LabelPattern2dspd.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TrackBarPattern2rate
        '
        Me.TrackBarPattern2rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern2rate.Location = New System.Drawing.Point(1333, 170)
        Me.TrackBarPattern2rate.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern2rate.Maximum = 4
        Me.TrackBarPattern2rate.Name = "TrackBarPattern2rate"
        Me.TrackBarPattern2rate.Size = New System.Drawing.Size(80, 114)
        Me.TrackBarPattern2rate.TabIndex = 12
        Me.TrackBarPattern2rate.Value = 1
        '
        'TrackBarPattern2dspd
        '
        Me.TrackBarPattern2dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern2dspd.Location = New System.Drawing.Point(1169, 170)
        Me.TrackBarPattern2dspd.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern2dspd.Maximum = 8
        Me.TrackBarPattern2dspd.Name = "TrackBarPattern2dspd"
        Me.TrackBarPattern2dspd.Size = New System.Drawing.Size(136, 114)
        Me.TrackBarPattern2dspd.TabIndex = 11
        Me.TrackBarPattern2dspd.Value = 4
        '
        'ComboBoxPattern2
        '
        Me.ComboBoxPattern2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxPattern2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPattern2.FormattingEnabled = True
        Me.ComboBoxPattern2.Location = New System.Drawing.Point(932, 166)
        Me.ComboBoxPattern2.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBoxPattern2.Name = "ComboBoxPattern2"
        Me.ComboBoxPattern2.Size = New System.Drawing.Size(204, 48)
        Me.ComboBoxPattern2.TabIndex = 10
        '
        'LabelPattern3
        '
        Me.LabelPattern3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern3.AutoSize = True
        Me.LabelPattern3.Location = New System.Drawing.Point(915, 212)
        Me.LabelPattern3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern3.Name = "LabelPattern3"
        Me.LabelPattern3.Size = New System.Drawing.Size(37, 40)
        Me.LabelPattern3.TabIndex = 16
        Me.LabelPattern3.Text = "3"
        '
        'LabelPattern4
        '
        Me.LabelPattern4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern4.AutoSize = True
        Me.LabelPattern4.Location = New System.Drawing.Point(915, 254)
        Me.LabelPattern4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern4.Name = "LabelPattern4"
        Me.LabelPattern4.Size = New System.Drawing.Size(37, 40)
        Me.LabelPattern4.TabIndex = 17
        Me.LabelPattern4.Text = "4"
        '
        'LabelPattern3rate
        '
        Me.LabelPattern3rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern3rate.AutoSize = True
        Me.LabelPattern3rate.Location = New System.Drawing.Point(1321, 213)
        Me.LabelPattern3rate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern3rate.Name = "LabelPattern3rate"
        Me.LabelPattern3rate.Size = New System.Drawing.Size(104, 40)
        Me.LabelPattern3rate.TabIndex = 22
        Me.LabelPattern3rate.Text = "rate3"
        Me.LabelPattern3rate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPattern3dspd
        '
        Me.LabelPattern3dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern3dspd.AutoSize = True
        Me.LabelPattern3dspd.Location = New System.Drawing.Point(1144, 213)
        Me.LabelPattern3dspd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern3dspd.Name = "LabelPattern3dspd"
        Me.LabelPattern3dspd.Size = New System.Drawing.Size(115, 40)
        Me.LabelPattern3dspd.TabIndex = 21
        Me.LabelPattern3dspd.Text = "dspd3"
        Me.LabelPattern3dspd.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TrackBarPattern3rate
        '
        Me.TrackBarPattern3rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern3rate.Location = New System.Drawing.Point(1333, 213)
        Me.TrackBarPattern3rate.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern3rate.Maximum = 4
        Me.TrackBarPattern3rate.Name = "TrackBarPattern3rate"
        Me.TrackBarPattern3rate.Size = New System.Drawing.Size(80, 114)
        Me.TrackBarPattern3rate.TabIndex = 20
        Me.TrackBarPattern3rate.Value = 1
        '
        'TrackBarPattern3dspd
        '
        Me.TrackBarPattern3dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern3dspd.Location = New System.Drawing.Point(1169, 213)
        Me.TrackBarPattern3dspd.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern3dspd.Maximum = 8
        Me.TrackBarPattern3dspd.Name = "TrackBarPattern3dspd"
        Me.TrackBarPattern3dspd.Size = New System.Drawing.Size(136, 114)
        Me.TrackBarPattern3dspd.TabIndex = 19
        Me.TrackBarPattern3dspd.Value = 4
        '
        'ComboBoxPattern3
        '
        Me.ComboBoxPattern3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxPattern3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPattern3.FormattingEnabled = True
        Me.ComboBoxPattern3.Location = New System.Drawing.Point(932, 209)
        Me.ComboBoxPattern3.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBoxPattern3.Name = "ComboBoxPattern3"
        Me.ComboBoxPattern3.Size = New System.Drawing.Size(204, 48)
        Me.ComboBoxPattern3.TabIndex = 18
        '
        'LabelPattern4rate
        '
        Me.LabelPattern4rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern4rate.AutoSize = True
        Me.LabelPattern4rate.Location = New System.Drawing.Point(1321, 255)
        Me.LabelPattern4rate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern4rate.Name = "LabelPattern4rate"
        Me.LabelPattern4rate.Size = New System.Drawing.Size(104, 40)
        Me.LabelPattern4rate.TabIndex = 27
        Me.LabelPattern4rate.Text = "rate4"
        Me.LabelPattern4rate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPattern4dspd
        '
        Me.LabelPattern4dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelPattern4dspd.AutoSize = True
        Me.LabelPattern4dspd.Location = New System.Drawing.Point(1144, 255)
        Me.LabelPattern4dspd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPattern4dspd.Name = "LabelPattern4dspd"
        Me.LabelPattern4dspd.Size = New System.Drawing.Size(115, 40)
        Me.LabelPattern4dspd.TabIndex = 26
        Me.LabelPattern4dspd.Text = "dspd4"
        Me.LabelPattern4dspd.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TrackBarPattern4rate
        '
        Me.TrackBarPattern4rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern4rate.Location = New System.Drawing.Point(1333, 255)
        Me.TrackBarPattern4rate.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern4rate.Maximum = 4
        Me.TrackBarPattern4rate.Name = "TrackBarPattern4rate"
        Me.TrackBarPattern4rate.Size = New System.Drawing.Size(80, 114)
        Me.TrackBarPattern4rate.TabIndex = 25
        Me.TrackBarPattern4rate.Value = 1
        '
        'TrackBarPattern4dspd
        '
        Me.TrackBarPattern4dspd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarPattern4dspd.Location = New System.Drawing.Point(1169, 255)
        Me.TrackBarPattern4dspd.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarPattern4dspd.Maximum = 8
        Me.TrackBarPattern4dspd.Name = "TrackBarPattern4dspd"
        Me.TrackBarPattern4dspd.Size = New System.Drawing.Size(136, 114)
        Me.TrackBarPattern4dspd.TabIndex = 24
        Me.TrackBarPattern4dspd.Value = 4
        '
        'ComboBoxPattern4
        '
        Me.ComboBoxPattern4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxPattern4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPattern4.FormattingEnabled = True
        Me.ComboBoxPattern4.Location = New System.Drawing.Point(932, 251)
        Me.ComboBoxPattern4.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBoxPattern4.Name = "ComboBoxPattern4"
        Me.ComboBoxPattern4.Size = New System.Drawing.Size(204, 48)
        Me.ComboBoxPattern4.TabIndex = 23
        '
        'LabelMode
        '
        Me.LabelMode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelMode.AutoSize = True
        Me.LabelMode.Location = New System.Drawing.Point(929, 293)
        Me.LabelMode.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelMode.Name = "LabelMode"
        Me.LabelMode.Size = New System.Drawing.Size(189, 40)
        Me.LabelMode.TabIndex = 28
        Me.LabelMode.Text = "実行モード"
        '
        'RadioButtonOrder
        '
        Me.RadioButtonOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButtonOrder.AutoSize = True
        Me.RadioButtonOrder.Checked = True
        Me.RadioButtonOrder.Location = New System.Drawing.Point(899, 291)
        Me.RadioButtonOrder.Name = "RadioButtonOrder"
        Me.RadioButtonOrder.Size = New System.Drawing.Size(195, 44)
        Me.RadioButtonOrder.TabIndex = 29
        Me.RadioButtonOrder.TabStop = True
        Me.RadioButtonOrder.Text = "上から順"
        Me.RadioButtonOrder.UseVisualStyleBackColor = True
        '
        'RadioButtonRandom
        '
        Me.RadioButtonRandom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButtonRandom.AutoSize = True
        Me.RadioButtonRandom.Location = New System.Drawing.Point(1005, 291)
        Me.RadioButtonRandom.Name = "RadioButtonRandom"
        Me.RadioButtonRandom.Size = New System.Drawing.Size(174, 44)
        Me.RadioButtonRandom.TabIndex = 30
        Me.RadioButtonRandom.Text = "ランダム"
        Me.RadioButtonRandom.UseVisualStyleBackColor = True
        '
        'CheckBoxUseInterval
        '
        Me.CheckBoxUseInterval.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxUseInterval.AutoSize = True
        Me.CheckBoxUseInterval.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.CheckBoxUseInterval.Location = New System.Drawing.Point(644, 341)
        Me.CheckBoxUseInterval.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxUseInterval.Name = "CheckBoxUseInterval"
        Me.CheckBoxUseInterval.Size = New System.Drawing.Size(455, 44)
        Me.CheckBoxUseInterval.TabIndex = 31
        Me.CheckBoxUseInterval.Text = "パターン間に休止をはさむ"
        Me.CheckBoxUseInterval.UseVisualStyleBackColor = True
        '
        'LabelMin
        '
        Me.LabelMin.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelMin.AutoSize = True
        Me.LabelMin.Location = New System.Drawing.Point(934, 373)
        Me.LabelMin.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelMin.Name = "LabelMin"
        Me.LabelMin.Size = New System.Drawing.Size(177, 40)
        Me.LabelMin.TabIndex = 33
        Me.LabelMin.Text = "最小時間"
        '
        'LabelMax
        '
        Me.LabelMax.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelMax.AutoSize = True
        Me.LabelMax.Location = New System.Drawing.Point(934, 416)
        Me.LabelMax.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelMax.Name = "LabelMax"
        Me.LabelMax.Size = New System.Drawing.Size(177, 40)
        Me.LabelMax.TabIndex = 34
        Me.LabelMax.Text = "最大時間"
        '
        'TextBoxLog
        '
        Me.TextBoxLog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxLog.BackColor = System.Drawing.Color.White
        Me.TextBoxLog.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBoxLog.Location = New System.Drawing.Point(913, 475)
        Me.TextBoxLog.Multiline = True
        Me.TextBoxLog.Name = "TextBoxLog"
        Me.TextBoxLog.ReadOnly = True
        Me.TextBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBoxLog.Size = New System.Drawing.Size(499, 330)
        Me.TextBoxLog.TabIndex = 39
        '
        'MenuStrip
        '
        Me.MenuStrip.Font = New System.Drawing.Font("Yu Gothic UI", 12.0!)
        Me.MenuStrip.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.MenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemFile, Me.ToolStripMenuItemEdit, Me.ToolStripMenuItemContentsPlay, Me.ToolStripMenuItemCsvGenerate})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(1426, 62)
        Me.MenuStrip.TabIndex = 44
        Me.MenuStrip.Text = "MenuStrip1"
        '
        'ToolStripMenuItemFile
        '
        Me.ToolStripMenuItemFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemOpen, Me.ToolStripMenuItemSave, Me.ToolStripMenuItemOverwriteSave, Me.ToolStripSeparator1, Me.TSVエクスポートToolStripMenuItem, Me.TSVインポートToolStripMenuItem, Me.ToolStripSeparator3, Me.ToolStripMenuItemClear, Me.ToolStripMenuItemEnd})
        Me.ToolStripMenuItemFile.Name = "ToolStripMenuItemFile"
        Me.ToolStripMenuItemFile.Size = New System.Drawing.Size(159, 58)
        Me.ToolStripMenuItemFile.Text = "ファイル"
        '
        'ToolStripMenuItemOpen
        '
        Me.ToolStripMenuItemOpen.Name = "ToolStripMenuItemOpen"
        Me.ToolStripMenuItemOpen.Size = New System.Drawing.Size(518, 62)
        Me.ToolStripMenuItemOpen.Text = "開く"
        '
        'ToolStripMenuItemSave
        '
        Me.ToolStripMenuItemSave.Name = "ToolStripMenuItemSave"
        Me.ToolStripMenuItemSave.Size = New System.Drawing.Size(518, 62)
        Me.ToolStripMenuItemSave.Text = "名前をつけて保存"
        '
        'ToolStripMenuItemOverwriteSave
        '
        Me.ToolStripMenuItemOverwriteSave.Name = "ToolStripMenuItemOverwriteSave"
        Me.ToolStripMenuItemOverwriteSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.ToolStripMenuItemOverwriteSave.Size = New System.Drawing.Size(518, 62)
        Me.ToolStripMenuItemOverwriteSave.Text = "上書き保存"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(515, 6)
        '
        'TSVエクスポートToolStripMenuItem
        '
        Me.TSVエクスポートToolStripMenuItem.Name = "TSVエクスポートToolStripMenuItem"
        Me.TSVエクスポートToolStripMenuItem.Size = New System.Drawing.Size(518, 62)
        Me.TSVエクスポートToolStripMenuItem.Text = "エクスポート(.tsv)"
        Me.TSVエクスポートToolStripMenuItem.ToolTipText = "タブ区切りファイルにエクスポートします。"
        '
        'TSVインポートToolStripMenuItem
        '
        Me.TSVインポートToolStripMenuItem.Name = "TSVインポートToolStripMenuItem"
        Me.TSVインポートToolStripMenuItem.Size = New System.Drawing.Size(518, 62)
        Me.TSVインポートToolStripMenuItem.Text = "インポート(.tsv)"
        Me.TSVインポートToolStripMenuItem.ToolTipText = "タブ区切りファイルをインポートします。"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(515, 6)
        '
        'ToolStripMenuItemClear
        '
        Me.ToolStripMenuItemClear.Name = "ToolStripMenuItemClear"
        Me.ToolStripMenuItemClear.Size = New System.Drawing.Size(518, 62)
        Me.ToolStripMenuItemClear.Text = "クリア"
        '
        'ToolStripMenuItemEnd
        '
        Me.ToolStripMenuItemEnd.Name = "ToolStripMenuItemEnd"
        Me.ToolStripMenuItemEnd.Size = New System.Drawing.Size(518, 62)
        Me.ToolStripMenuItemEnd.Text = "終了"
        '
        'ToolStripMenuItemEdit
        '
        Me.ToolStripMenuItemEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemAddRow, Me.ToolStripMenuItemDeleteRow, Me.一括編集ToolStripMenuItem, Me.ToolStripSeparator2, Me.ToolStripMenuItemDelay, Me.全体を01秒進めるToolStripMenuItem, Me.全体を01秒遅らせるToolStripMenuItem, Me.ToolStripSeparator4, Me.コンテンツ再生速度ToolStripMenuItem, Me.back5sToolStripMenuItem, Me.skip5sToolStripMenuItem})
        Me.ToolStripMenuItemEdit.Name = "ToolStripMenuItemEdit"
        Me.ToolStripMenuItemEdit.Size = New System.Drawing.Size(127, 58)
        Me.ToolStripMenuItemEdit.Text = "編集"
        '
        'ToolStripMenuItemAddRow
        '
        Me.ToolStripMenuItemAddRow.Name = "ToolStripMenuItemAddRow"
        Me.ToolStripMenuItemAddRow.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.ToolStripMenuItemAddRow.Size = New System.Drawing.Size(672, 62)
        Me.ToolStripMenuItemAddRow.Text = "行追加"
        '
        'ToolStripMenuItemDeleteRow
        '
        Me.ToolStripMenuItemDeleteRow.Name = "ToolStripMenuItemDeleteRow"
        Me.ToolStripMenuItemDeleteRow.Size = New System.Drawing.Size(672, 62)
        Me.ToolStripMenuItemDeleteRow.Text = "行削除(Delete)"
        '
        '一括編集ToolStripMenuItem
        '
        Me.一括編集ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.末尾に100行追加ToolStripMenuItem, Me.末尾に100行追加パターンありToolStripMenuItem})
        Me.一括編集ToolStripMenuItem.Name = "一括編集ToolStripMenuItem"
        Me.一括編集ToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.一括編集ToolStripMenuItem.Text = "一括編集"
        '
        '末尾に100行追加ToolStripMenuItem
        '
        Me.末尾に100行追加ToolStripMenuItem.Name = "末尾に100行追加ToolStripMenuItem"
        Me.末尾に100行追加ToolStripMenuItem.Size = New System.Drawing.Size(709, 62)
        Me.末尾に100行追加ToolStripMenuItem.Text = "末尾に100行追加"
        '
        '末尾に100行追加パターンありToolStripMenuItem
        '
        Me.末尾に100行追加パターンありToolStripMenuItem.Name = "末尾に100行追加パターンありToolStripMenuItem"
        Me.末尾に100行追加パターンありToolStripMenuItem.Size = New System.Drawing.Size(709, 62)
        Me.末尾に100行追加パターンありToolStripMenuItem.Text = "末尾に100行追加　パターンあり"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(669, 6)
        '
        'ToolStripMenuItemDelay
        '
        Me.ToolStripMenuItemDelay.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem_plus5s, Me.ToolStripMenuItem_plus4s, Me.ToolStripMenuItem_plus3s, Me.ToolStripMenuItem_plus2s, Me.ToolStripMenuItem_plus1s, Me.ToolStripMenuItem_plus0s, Me.ToolStripMenuItem_minus1s, Me.ToolStripMenuItem_minus2s, Me.ToolStripMenuItem_minus3s, Me.ToolStripMenuItem_minus4s, Me.ToolStripMenuItem_minus5s})
        Me.ToolStripMenuItemDelay.Name = "ToolStripMenuItemDelay"
        Me.ToolStripMenuItemDelay.Size = New System.Drawing.Size(672, 62)
        Me.ToolStripMenuItemDelay.Text = "時間キャプチャ　ディレイ設定"
        '
        'ToolStripMenuItem_plus5s
        '
        Me.ToolStripMenuItem_plus5s.CheckOnClick = True
        Me.ToolStripMenuItem_plus5s.Name = "ToolStripMenuItem_plus5s"
        Me.ToolStripMenuItem_plus5s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus5s.Text = "+0.5s"
        '
        'ToolStripMenuItem_plus4s
        '
        Me.ToolStripMenuItem_plus4s.CheckOnClick = True
        Me.ToolStripMenuItem_plus4s.Name = "ToolStripMenuItem_plus4s"
        Me.ToolStripMenuItem_plus4s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus4s.Text = "+0.4s"
        '
        'ToolStripMenuItem_plus3s
        '
        Me.ToolStripMenuItem_plus3s.CheckOnClick = True
        Me.ToolStripMenuItem_plus3s.Name = "ToolStripMenuItem_plus3s"
        Me.ToolStripMenuItem_plus3s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus3s.Text = "+0.3s"
        '
        'ToolStripMenuItem_plus2s
        '
        Me.ToolStripMenuItem_plus2s.CheckOnClick = True
        Me.ToolStripMenuItem_plus2s.Name = "ToolStripMenuItem_plus2s"
        Me.ToolStripMenuItem_plus2s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus2s.Text = "+0.2s"
        '
        'ToolStripMenuItem_plus1s
        '
        Me.ToolStripMenuItem_plus1s.CheckOnClick = True
        Me.ToolStripMenuItem_plus1s.Name = "ToolStripMenuItem_plus1s"
        Me.ToolStripMenuItem_plus1s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus1s.Text = "+0.1s"
        '
        'ToolStripMenuItem_plus0s
        '
        Me.ToolStripMenuItem_plus0s.CheckOnClick = True
        Me.ToolStripMenuItem_plus0s.Name = "ToolStripMenuItem_plus0s"
        Me.ToolStripMenuItem_plus0s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_plus0s.Text = "±0s"
        '
        'ToolStripMenuItem_minus1s
        '
        Me.ToolStripMenuItem_minus1s.CheckOnClick = True
        Me.ToolStripMenuItem_minus1s.Name = "ToolStripMenuItem_minus1s"
        Me.ToolStripMenuItem_minus1s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_minus1s.Text = "-0.1s"
        '
        'ToolStripMenuItem_minus2s
        '
        Me.ToolStripMenuItem_minus2s.CheckOnClick = True
        Me.ToolStripMenuItem_minus2s.Name = "ToolStripMenuItem_minus2s"
        Me.ToolStripMenuItem_minus2s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_minus2s.Text = "-0.2s"
        '
        'ToolStripMenuItem_minus3s
        '
        Me.ToolStripMenuItem_minus3s.CheckOnClick = True
        Me.ToolStripMenuItem_minus3s.Name = "ToolStripMenuItem_minus3s"
        Me.ToolStripMenuItem_minus3s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_minus3s.Text = "-0.3s"
        '
        'ToolStripMenuItem_minus4s
        '
        Me.ToolStripMenuItem_minus4s.CheckOnClick = True
        Me.ToolStripMenuItem_minus4s.Name = "ToolStripMenuItem_minus4s"
        Me.ToolStripMenuItem_minus4s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_minus4s.Text = "-0.4s"
        '
        'ToolStripMenuItem_minus5s
        '
        Me.ToolStripMenuItem_minus5s.CheckOnClick = True
        Me.ToolStripMenuItem_minus5s.Name = "ToolStripMenuItem_minus5s"
        Me.ToolStripMenuItem_minus5s.Size = New System.Drawing.Size(291, 62)
        Me.ToolStripMenuItem_minus5s.Text = "-0.5s"
        '
        '全体を01秒進めるToolStripMenuItem
        '
        Me.全体を01秒進めるToolStripMenuItem.Name = "全体を01秒進めるToolStripMenuItem"
        Me.全体を01秒進めるToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.全体を01秒進めるToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.全体を01秒進めるToolStripMenuItem.Text = "全体を0.1秒遅らせる"
        Me.全体を01秒進めるToolStripMenuItem.ToolTipText = "時間を全体的に0.1秒遅らせます。"
        '
        '全体を01秒遅らせるToolStripMenuItem
        '
        Me.全体を01秒遅らせるToolStripMenuItem.Name = "全体を01秒遅らせるToolStripMenuItem"
        Me.全体を01秒遅らせるToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.全体を01秒遅らせるToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.全体を01秒遅らせるToolStripMenuItem.Text = "全体を0.1秒早める"
        Me.全体を01秒遅らせるToolStripMenuItem.ToolTipText = "時間を全体的に0.1秒早めます。"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(669, 6)
        '
        'コンテンツ再生速度ToolStripMenuItem
        '
        Me.コンテンツ再生速度ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.X050ToolStripMenuItem, Me.X075ToolStripMenuItem, Me.X100ToolStripMenuItem, Me.X125ToolStripMenuItem, Me.X150ToolStripMenuItem, Me.X200ToolStripMenuItem})
        Me.コンテンツ再生速度ToolStripMenuItem.Enabled = False
        Me.コンテンツ再生速度ToolStripMenuItem.Name = "コンテンツ再生速度ToolStripMenuItem"
        Me.コンテンツ再生速度ToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.コンテンツ再生速度ToolStripMenuItem.Text = "コンテンツ再生速度"
        '
        'X050ToolStripMenuItem
        '
        Me.X050ToolStripMenuItem.Enabled = False
        Me.X050ToolStripMenuItem.Name = "X050ToolStripMenuItem"
        Me.X050ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X050ToolStripMenuItem.Text = "0.5x"
        '
        'X075ToolStripMenuItem
        '
        Me.X075ToolStripMenuItem.Enabled = False
        Me.X075ToolStripMenuItem.Name = "X075ToolStripMenuItem"
        Me.X075ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X075ToolStripMenuItem.Text = "0.75x"
        '
        'X100ToolStripMenuItem
        '
        Me.X100ToolStripMenuItem.Enabled = False
        Me.X100ToolStripMenuItem.Name = "X100ToolStripMenuItem"
        Me.X100ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X100ToolStripMenuItem.Text = "1.0x"
        '
        'X125ToolStripMenuItem
        '
        Me.X125ToolStripMenuItem.Enabled = False
        Me.X125ToolStripMenuItem.Name = "X125ToolStripMenuItem"
        Me.X125ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X125ToolStripMenuItem.Text = "1.25x"
        '
        'X150ToolStripMenuItem
        '
        Me.X150ToolStripMenuItem.Enabled = False
        Me.X150ToolStripMenuItem.Name = "X150ToolStripMenuItem"
        Me.X150ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X150ToolStripMenuItem.Text = "1.5x"
        '
        'X200ToolStripMenuItem
        '
        Me.X200ToolStripMenuItem.Enabled = False
        Me.X200ToolStripMenuItem.Name = "X200ToolStripMenuItem"
        Me.X200ToolStripMenuItem.Size = New System.Drawing.Size(287, 62)
        Me.X200ToolStripMenuItem.Text = "2.0x"
        '
        'back5sToolStripMenuItem
        '
        Me.back5sToolStripMenuItem.Enabled = False
        Me.back5sToolStripMenuItem.Name = "back5sToolStripMenuItem"
        Me.back5sToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.back5sToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.back5sToolStripMenuItem.Text = "5秒戻る"
        '
        'skip5sToolStripMenuItem
        '
        Me.skip5sToolStripMenuItem.Enabled = False
        Me.skip5sToolStripMenuItem.Name = "skip5sToolStripMenuItem"
        Me.skip5sToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.G), System.Windows.Forms.Keys)
        Me.skip5sToolStripMenuItem.Size = New System.Drawing.Size(672, 62)
        Me.skip5sToolStripMenuItem.Text = "5秒スキップ"
        '
        'ToolStripMenuItemContentsPlay
        '
        Me.ToolStripMenuItemContentsPlay.Name = "ToolStripMenuItemContentsPlay"
        Me.ToolStripMenuItemContentsPlay.Size = New System.Drawing.Size(275, 58)
        Me.ToolStripMenuItemContentsPlay.Text = "コンテンツ再生"
        '
        'ToolStripMenuItemCsvGenerate
        '
        Me.ToolStripMenuItemCsvGenerate.Name = "ToolStripMenuItemCsvGenerate"
        Me.ToolStripMenuItemCsvGenerate.Size = New System.Drawing.Size(198, 58)
        Me.ToolStripMenuItemCsvGenerate.Text = "CSV出力"
        '
        'ButtonTimeCapture
        '
        Me.ButtonTimeCapture.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonTimeCapture.Enabled = False
        Me.ButtonTimeCapture.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonTimeCapture.Location = New System.Drawing.Point(13, 755)
        Me.ButtonTimeCapture.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonTimeCapture.Name = "ButtonTimeCapture"
        Me.ButtonTimeCapture.Size = New System.Drawing.Size(880, 50)
        Me.ButtonTimeCapture.TabIndex = 45
        Me.ButtonTimeCapture.Text = "動画・音声から時間キャプチャ(Space)"
        Me.ButtonTimeCapture.UseVisualStyleBackColor = True
        '
        'LabelLog
        '
        Me.LabelLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelLog.AutoSize = True
        Me.LabelLog.Location = New System.Drawing.Point(915, 456)
        Me.LabelLog.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelLog.Name = "LabelLog"
        Me.LabelLog.Size = New System.Drawing.Size(97, 40)
        Me.LabelLog.TabIndex = 46
        Me.LabelLog.Text = "情報"
        '
        'TrackBarItvMinTime
        '
        Me.TrackBarItvMinTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarItvMinTime.Location = New System.Drawing.Point(1080, 373)
        Me.TrackBarItvMinTime.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarItvMinTime.Maximum = 20
        Me.TrackBarItvMinTime.Name = "TrackBarItvMinTime"
        Me.TrackBarItvMinTime.Size = New System.Drawing.Size(332, 114)
        Me.TrackBarItvMinTime.TabIndex = 47
        Me.TrackBarItvMinTime.Value = 4
        '
        'TrackBarItvMaxTime
        '
        Me.TrackBarItvMaxTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBarItvMaxTime.Location = New System.Drawing.Point(1080, 416)
        Me.TrackBarItvMaxTime.Margin = New System.Windows.Forms.Padding(4)
        Me.TrackBarItvMaxTime.Maximum = 20
        Me.TrackBarItvMaxTime.Name = "TrackBarItvMaxTime"
        Me.TrackBarItvMaxTime.Size = New System.Drawing.Size(332, 114)
        Me.TrackBarItvMaxTime.TabIndex = 48
        Me.TrackBarItvMaxTime.Value = 4
        '
        'LabelMinTime
        '
        Me.LabelMinTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelMinTime.AutoSize = True
        Me.LabelMinTime.Location = New System.Drawing.Point(1022, 373)
        Me.LabelMinTime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelMinTime.Name = "LabelMinTime"
        Me.LabelMinTime.Size = New System.Drawing.Size(74, 40)
        Me.LabelMinTime.TabIndex = 49
        Me.LabelMinTime.Text = "min"
        Me.LabelMinTime.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelMaxTime
        '
        Me.LabelMaxTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelMaxTime.AutoSize = True
        Me.LabelMaxTime.Location = New System.Drawing.Point(1022, 416)
        Me.LabelMaxTime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelMaxTime.Name = "LabelMaxTime"
        Me.LabelMaxTime.Size = New System.Drawing.Size(83, 40)
        Me.LabelMaxTime.TabIndex = 50
        Me.LabelMaxTime.Text = "max"
        Me.LabelMaxTime.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelSec1
        '
        Me.LabelSec1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelSec1.AutoSize = True
        Me.LabelSec1.Location = New System.Drawing.Point(1049, 373)
        Me.LabelSec1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelSec1.Name = "LabelSec1"
        Me.LabelSec1.Size = New System.Drawing.Size(57, 40)
        Me.LabelSec1.TabIndex = 51
        Me.LabelSec1.Text = "秒"
        Me.LabelSec1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelSec2
        '
        Me.LabelSec2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelSec2.AutoSize = True
        Me.LabelSec2.Location = New System.Drawing.Point(1049, 416)
        Me.LabelSec2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelSec2.Name = "LabelSec2"
        Me.LabelSec2.Size = New System.Drawing.Size(57, 40)
        Me.LabelSec2.TabIndex = 52
        Me.LabelSec2.Text = "秒"
        Me.LabelSec2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'OxyPlotView
        '
        Me.OxyPlotView.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OxyPlotView.Location = New System.Drawing.Point(13, 812)
        Me.OxyPlotView.Name = "OxyPlotView"
        Me.OxyPlotView.PanCursor = System.Windows.Forms.Cursors.Hand
        Me.OxyPlotView.Size = New System.Drawing.Size(1401, 220)
        Me.OxyPlotView.TabIndex = 53
        Me.OxyPlotView.Text = "PlotView1"
        Me.OxyPlotView.ZoomHorizontalCursor = System.Windows.Forms.Cursors.SizeWE
        Me.OxyPlotView.ZoomRectangleCursor = System.Windows.Forms.Cursors.SizeNWSE
        Me.OxyPlotView.ZoomVerticalCursor = System.Windows.Forms.Cursors.SizeNS
        '
        'Scry1CsvGenerator
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(21.0!, 40.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1426, 1044)
        Me.Controls.Add(Me.OxyPlotView)
        Me.Controls.Add(Me.LabelSec2)
        Me.Controls.Add(Me.LabelSec1)
        Me.Controls.Add(Me.LabelMaxTime)
        Me.Controls.Add(Me.LabelMinTime)
        Me.Controls.Add(Me.TrackBarItvMaxTime)
        Me.Controls.Add(Me.TrackBarItvMinTime)
        Me.Controls.Add(Me.LabelLog)
        Me.Controls.Add(Me.ButtonTimeCapture)
        Me.Controls.Add(Me.TextBoxLog)
        Me.Controls.Add(Me.LabelMax)
        Me.Controls.Add(Me.LabelMin)
        Me.Controls.Add(Me.CheckBoxUseInterval)
        Me.Controls.Add(Me.RadioButtonRandom)
        Me.Controls.Add(Me.RadioButtonOrder)
        Me.Controls.Add(Me.LabelMode)
        Me.Controls.Add(Me.LabelPattern4rate)
        Me.Controls.Add(Me.LabelPattern4dspd)
        Me.Controls.Add(Me.TrackBarPattern4rate)
        Me.Controls.Add(Me.TrackBarPattern4dspd)
        Me.Controls.Add(Me.ComboBoxPattern4)
        Me.Controls.Add(Me.LabelPattern3rate)
        Me.Controls.Add(Me.LabelPattern3dspd)
        Me.Controls.Add(Me.TrackBarPattern3rate)
        Me.Controls.Add(Me.TrackBarPattern3dspd)
        Me.Controls.Add(Me.ComboBoxPattern3)
        Me.Controls.Add(Me.LabelPattern4)
        Me.Controls.Add(Me.LabelPattern3)
        Me.Controls.Add(Me.LabelPattern2)
        Me.Controls.Add(Me.LabelPattern2rate)
        Me.Controls.Add(Me.LabelPattern2dspd)
        Me.Controls.Add(Me.TrackBarPattern2rate)
        Me.Controls.Add(Me.TrackBarPattern2dspd)
        Me.Controls.Add(Me.ComboBoxPattern2)
        Me.Controls.Add(Me.LabelPattern1)
        Me.Controls.Add(Me.LabelPatternDetail)
        Me.Controls.Add(Me.LabelPattern1rate)
        Me.Controls.Add(Me.LabelPattern1dspd)
        Me.Controls.Add(Me.TrackBarPattern1rate)
        Me.Controls.Add(Me.TrackBarPattern1dspd)
        Me.Controls.Add(Me.ComboBoxPattern1)
        Me.Controls.Add(Me.CheckBoxUsePattern)
        Me.Controls.Add(Me.ButtonPatternSetting)
        Me.Controls.Add(Me.DataGridViewTimesheet)
        Me.Controls.Add(Me.MenuStrip)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "Scry1CsvGenerator"
        Me.Text = "Scry1CsvGenerator"
        CType(Me.DataGridViewTimesheet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern1dspd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern1rate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern2rate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern2dspd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern3rate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern3dspd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern4rate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarPattern4dspd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        CType(Me.TrackBarItvMinTime, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBarItvMaxTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridViewTimesheet As DataGridView
    Friend WithEvents ButtonPatternSetting As Button
    Friend WithEvents CheckBoxUsePattern As CheckBox
    Friend WithEvents ComboBoxPattern1 As ComboBox
    Friend WithEvents TrackBarPattern1dspd As TrackBar
    Friend WithEvents TrackBarPattern1rate As TrackBar
    Friend WithEvents LabelPattern1dspd As Label
    Friend WithEvents LabelPattern1rate As Label
    Friend WithEvents LabelPatternDetail As Label
    Friend WithEvents LabelPattern1 As Label
    Friend WithEvents LabelPattern2 As Label
    Friend WithEvents LabelPattern2rate As Label
    Friend WithEvents LabelPattern2dspd As Label
    Friend WithEvents TrackBarPattern2rate As TrackBar
    Friend WithEvents TrackBarPattern2dspd As TrackBar
    Friend WithEvents ComboBoxPattern2 As ComboBox
    Friend WithEvents LabelPattern3 As Label
    Friend WithEvents LabelPattern4 As Label
    Friend WithEvents LabelPattern3rate As Label
    Friend WithEvents LabelPattern3dspd As Label
    Friend WithEvents TrackBarPattern3rate As TrackBar
    Friend WithEvents TrackBarPattern3dspd As TrackBar
    Friend WithEvents ComboBoxPattern3 As ComboBox
    Friend WithEvents LabelPattern4rate As Label
    Friend WithEvents LabelPattern4dspd As Label
    Friend WithEvents TrackBarPattern4rate As TrackBar
    Friend WithEvents TrackBarPattern4dspd As TrackBar
    Friend WithEvents ComboBoxPattern4 As ComboBox
    Friend WithEvents LabelMode As Label
    Friend WithEvents RadioButtonOrder As RadioButton
    Friend WithEvents RadioButtonRandom As RadioButton
    Friend WithEvents CheckBoxUseInterval As CheckBox
    Friend WithEvents LabelMin As Label
    Friend WithEvents LabelMax As Label
    Friend WithEvents TextBoxLog As TextBox
    Friend WithEvents MenuStrip As MenuStrip
    Friend WithEvents ToolStripMenuItemFile As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemContentsPlay As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemOpen As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemSave As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItemClear As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemEnd As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemEdit As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemAddRow As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemDeleteRow As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ButtonTimeCapture As Button
    Friend WithEvents LabelLog As Label
    Friend WithEvents ToolStripMenuItemCsvGenerate As ToolStripMenuItem
    Friend WithEvents TrackBarItvMinTime As TrackBar
    Friend WithEvents TrackBarItvMaxTime As TrackBar
    Friend WithEvents LabelMinTime As Label
    Friend WithEvents LabelMaxTime As Label
    Friend WithEvents LabelSec1 As Label
    Friend WithEvents LabelSec2 As Label
    Friend WithEvents ToolStripMenuItemDelay As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus5s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus4s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus3s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus2s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus1s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_plus0s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_minus1s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_minus2s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_minus3s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_minus4s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_minus5s As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItemOverwriteSave As ToolStripMenuItem
    Friend WithEvents 全体を01秒進めるToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 全体を01秒遅らせるToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents 一括編集ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents コンテンツ再生速度ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X050ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X075ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X100ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X125ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X150ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents X200ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 末尾に100行追加ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 末尾に100行追加パターンありToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents back5sToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents skip5sToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TSVエクスポートToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TSVインポートToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents vtimehms As DataGridViewTextBoxColumn
    Friend WithEvents vtime As DataGridViewTextBoxColumn
    Friend WithEvents vtimediff As DataGridViewTextBoxColumn
    Friend WithEvents vdir As DataGridViewTextBoxColumn
    Friend WithEvents vspd As DataGridViewTextBoxColumn
    Friend WithEvents check As DataGridViewCheckBoxColumn
    Friend WithEvents patn1 As DataGridViewTextBoxColumn
    Friend WithEvents patn2 As DataGridViewTextBoxColumn
    Friend WithEvents patn3 As DataGridViewTextBoxColumn
    Friend WithEvents patn4 As DataGridViewTextBoxColumn
    Friend WithEvents vcmt As DataGridViewTextBoxColumn
    Friend WithEvents dspd1 As DataGridViewTextBoxColumn
    Friend WithEvents dspd2 As DataGridViewTextBoxColumn
    Friend WithEvents dspd3 As DataGridViewTextBoxColumn
    Friend WithEvents dspd4 As DataGridViewTextBoxColumn
    Friend WithEvents rate1 As DataGridViewTextBoxColumn
    Friend WithEvents rate2 As DataGridViewTextBoxColumn
    Friend WithEvents rate3 As DataGridViewTextBoxColumn
    Friend WithEvents rate4 As DataGridViewTextBoxColumn
    Friend WithEvents userdm As DataGridViewCheckBoxColumn
    Friend WithEvents useitv As DataGridViewCheckBoxColumn
    Friend WithEvents itvmin As DataGridViewTextBoxColumn
    Friend WithEvents itvmax As DataGridViewTextBoxColumn
    Friend WithEvents OxyPlotView As OxyPlot.WindowsForms.PlotView
End Class

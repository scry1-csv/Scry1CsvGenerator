<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PatternEditor
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PatternEditor))
        Me.ComboBoxPattern = New System.Windows.Forms.ComboBox()
        Me.ButtonPatternAdd = New System.Windows.Forms.Button()
        Me.ListBox = New System.Windows.Forms.ListBox()
        Me.ButtonDelete = New System.Windows.Forms.Button()
        Me.DataGridViewPattern = New System.Windows.Forms.DataGridView()
        Me.vtime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vdir = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.vspd = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.ButtonRename = New System.Windows.Forms.Button()
        Me.LabelTip = New System.Windows.Forms.Label()
        Me.ButtonUp = New System.Windows.Forms.Button()
        Me.ButtonDown = New System.Windows.Forms.Button()
        CType(Me.DataGridViewPattern, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBoxPattern
        '
        Me.ComboBoxPattern.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPattern.FormattingEnabled = True
        Me.ComboBoxPattern.Location = New System.Drawing.Point(13, 13)
        Me.ComboBoxPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBoxPattern.Name = "ComboBoxPattern"
        Me.ComboBoxPattern.Size = New System.Drawing.Size(199, 24)
        Me.ComboBoxPattern.TabIndex = 4
        '
        'ButtonPatternAdd
        '
        Me.ButtonPatternAdd.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonPatternAdd.Location = New System.Drawing.Point(220, 9)
        Me.ButtonPatternAdd.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonPatternAdd.Name = "ButtonPatternAdd"
        Me.ButtonPatternAdd.Size = New System.Drawing.Size(78, 31)
        Me.ButtonPatternAdd.TabIndex = 5
        Me.ButtonPatternAdd.Text = "追加"
        Me.ButtonPatternAdd.UseVisualStyleBackColor = True
        '
        'ListBox
        '
        Me.ListBox.FormattingEnabled = True
        Me.ListBox.ItemHeight = 16
        Me.ListBox.Location = New System.Drawing.Point(13, 44)
        Me.ListBox.Name = "ListBox"
        Me.ListBox.Size = New System.Drawing.Size(247, 532)
        Me.ListBox.TabIndex = 6
        '
        'ButtonDelete
        '
        Me.ButtonDelete.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonDelete.Location = New System.Drawing.Point(265, 161)
        Me.ButtonDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonDelete.Name = "ButtonDelete"
        Me.ButtonDelete.Size = New System.Drawing.Size(33, 31)
        Me.ButtonDelete.TabIndex = 9
        Me.ButtonDelete.Text = "×"
        Me.ButtonDelete.UseVisualStyleBackColor = True
        '
        'DataGridViewPattern
        '
        Me.DataGridViewPattern.AllowUserToAddRows = False
        Me.DataGridViewPattern.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewPattern.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.vtime, Me.vdir, Me.vspd})
        Me.DataGridViewPattern.Location = New System.Drawing.Point(306, 44)
        Me.DataGridViewPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridViewPattern.MultiSelect = False
        Me.DataGridViewPattern.Name = "DataGridViewPattern"
        Me.DataGridViewPattern.RowTemplate.Height = 21
        Me.DataGridViewPattern.Size = New System.Drawing.Size(339, 532)
        Me.DataGridViewPattern.TabIndex = 30
        '
        'vtime
        '
        Me.vtime.HeaderText = "時間"
        Me.vtime.MaxInputLength = 4
        Me.vtime.Name = "vtime"
        Me.vtime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vtime.ToolTipText = "時間 [10=1秒]"
        '
        'vdir
        '
        Me.vdir.HeaderText = "方向"
        Me.vdir.MaxInputLength = 1
        Me.vdir.Name = "vdir"
        Me.vdir.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vdir.ToolTipText = "回転方向 [0=右回転, 1=左回転]"
        Me.vdir.Width = 80
        '
        'vspd
        '
        Me.vspd.HeaderText = "速度"
        Me.vspd.MaxInputLength = 3
        Me.vspd.Name = "vspd"
        Me.vspd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.vspd.ToolTipText = "回転速度 [推奨値・・・A10サイクロンSA：40以上　U.F.O.SA：25以上]"
        Me.vspd.Width = 80
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonSave.Location = New System.Drawing.Point(492, 587)
        Me.ButtonSave.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(153, 31)
        Me.ButtonSave.TabIndex = 31
        Me.ButtonSave.Text = "保存"
        Me.ButtonSave.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonCancel.Location = New System.Drawing.Point(13, 587)
        Me.ButtonCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(132, 31)
        Me.ButtonCancel.TabIndex = 32
        Me.ButtonCancel.Text = "キャンセル"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'ButtonRename
        '
        Me.ButtonRename.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonRename.Location = New System.Drawing.Point(265, 122)
        Me.ButtonRename.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonRename.Name = "ButtonRename"
        Me.ButtonRename.Size = New System.Drawing.Size(33, 31)
        Me.ButtonRename.TabIndex = 33
        Me.ButtonRename.Text = "..."
        Me.ButtonRename.UseVisualStyleBackColor = True
        '
        'LabelTip
        '
        Me.LabelTip.AutoSize = True
        Me.LabelTip.Font = New System.Drawing.Font("MS UI Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LabelTip.Location = New System.Drawing.Point(303, 12)
        Me.LabelTip.Name = "LabelTip"
        Me.LabelTip.Size = New System.Drawing.Size(174, 28)
        Me.LabelTip.TabIndex = 34
        Me.LabelTip.Text = "行削除：Deleteキー" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "最後の行の方向・速度は不要"
        '
        'ButtonUp
        '
        Me.ButtonUp.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonUp.Location = New System.Drawing.Point(265, 44)
        Me.ButtonUp.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonUp.Name = "ButtonUp"
        Me.ButtonUp.Size = New System.Drawing.Size(33, 31)
        Me.ButtonUp.TabIndex = 37
        Me.ButtonUp.Text = "▲"
        Me.ButtonUp.UseVisualStyleBackColor = True
        '
        'ButtonDown
        '
        Me.ButtonDown.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonDown.Location = New System.Drawing.Point(265, 83)
        Me.ButtonDown.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonDown.Name = "ButtonDown"
        Me.ButtonDown.Size = New System.Drawing.Size(33, 31)
        Me.ButtonDown.TabIndex = 38
        Me.ButtonDown.Text = "▼"
        Me.ButtonDown.UseVisualStyleBackColor = True
        '
        'PatternEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 631)
        Me.Controls.Add(Me.ButtonDown)
        Me.Controls.Add(Me.ButtonUp)
        Me.Controls.Add(Me.LabelTip)
        Me.Controls.Add(Me.ButtonRename)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonSave)
        Me.Controls.Add(Me.DataGridViewPattern)
        Me.Controls.Add(Me.ButtonDelete)
        Me.Controls.Add(Me.ListBox)
        Me.Controls.Add(Me.ButtonPatternAdd)
        Me.Controls.Add(Me.ComboBoxPattern)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "PatternEditor"
        Me.Text = "パターン編集"
        CType(Me.DataGridViewPattern, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBoxPattern As ComboBox
    Friend WithEvents ButtonPatternAdd As Button
    Friend WithEvents ListBox As ListBox
    Friend WithEvents ButtonDelete As Button
    Friend WithEvents DataGridViewPattern As DataGridView
    Friend WithEvents ButtonSave As Button
    Friend WithEvents ButtonCancel As Button
    Friend WithEvents ButtonRename As Button
    Friend WithEvents LabelTip As Label
    Friend WithEvents vtime As DataGridViewTextBoxColumn
    Friend WithEvents vdir As DataGridViewTextBoxColumn
    Friend WithEvents vspd As DataGridViewTextBoxColumn
    Friend WithEvents ButtonUp As Button
    Friend WithEvents ButtonDown As Button
End Class

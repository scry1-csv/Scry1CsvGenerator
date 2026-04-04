<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form4))
        Me.TextBoxPatternName = New System.Windows.Forms.TextBox()
        Me.LabelPatternName = New System.Windows.Forms.Label()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBoxPatternName
        '
        Me.TextBoxPatternName.Location = New System.Drawing.Point(136, 24)
        Me.TextBoxPatternName.Name = "TextBoxPatternName"
        Me.TextBoxPatternName.Size = New System.Drawing.Size(222, 27)
        Me.TextBoxPatternName.TabIndex = 33
        '
        'LabelPatternName
        '
        Me.LabelPatternName.AutoSize = True
        Me.LabelPatternName.Location = New System.Drawing.Point(28, 27)
        Me.LabelPatternName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelPatternName.Name = "LabelPatternName"
        Me.LabelPatternName.Size = New System.Drawing.Size(101, 20)
        Me.LabelPatternName.TabIndex = 34
        Me.LabelPatternName.Text = "パターン名："
        '
        'ButtonOK
        '
        Me.ButtonOK.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.ButtonOK.Location = New System.Drawing.Point(288, 73)
        Me.ButtonOK.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(153, 31)
        Me.ButtonOK.TabIndex = 35
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.Button1.Location = New System.Drawing.Point(127, 73)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(153, 31)
        Me.Button1.TabIndex = 36
        Me.Button1.Text = "キャンセル"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(454, 117)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.LabelPatternName)
        Me.Controls.Add(Me.TextBoxPatternName)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form4"
        Me.Text = "パターン名変更"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBoxPatternName As TextBox
    Friend WithEvents LabelPatternName As Label
    Friend WithEvents ButtonOK As Button
    Friend WithEvents Button1 As Button
End Class

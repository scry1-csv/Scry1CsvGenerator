Public Class Form4

#Region "クラス変数"

    Private _PatternListNames() As String
    Private _defaultPatternName As String
    Private _isOK As Boolean

#End Region


#Region "静的変数"

    ''' <summary>
    ''' 自身インスタンス
    ''' </summary>
    Private Shared _PatternNameFormInstance As Form4

    ''' <summary>
    ''' 自身インスタンス
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property PatternNameFormInstance() As Form4
        Get
            Return _PatternNameFormInstance
        End Get
        Set(value As Form4)
            _PatternNameFormInstance = value
        End Set
    End Property

    ''' <summary>
    ''' パターン保存実行フラグ　プロパティ
    ''' </summary>
    ''' <returns></returns>
    Public Property IsOK() As Boolean
        Get
            Return _isOK
        End Get
        Set(value As Boolean)
            _isOK = value
        End Set
    End Property

    Public Property PatternName() As String
        Get
            Return TextBoxPatternName.Text
        End Get
        Set(value As String)
            TextBoxPatternName.Text = value
        End Set
    End Property

#End Region

#Region "イベント"

    Public Sub New(patternListNames As String(), defaultPatternName As String, mode As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _PatternListNames = patternListNames
        _defaultPatternName = defaultPatternName
        TextBoxPatternName.Text = defaultPatternName

        If mode.Equals("add") Then
            Me.Text = "追加"
        ElseIf mode.Equals("change") Then
            Me.Text = "名称変更"
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        _isOK = False
        Me.Close()
    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click

        If TextBoxPatternName.Text.Equals(String.Empty) Then
            MessageBox.Show("パターン名を入力してください")
            Exit Sub
        End If

        '名称が変更されていない場合、実質は何もせずフォームを閉じる
        If _defaultPatternName.Equals(TextBoxPatternName.Text) Then
            _isOK = False
            Me.Close()
            Exit Sub
        End If

        'パターン名重複チェック
        'パターンが0の場合、配列の1つ目にNULLが格納されている
        If Not String.IsNullOrEmpty(_PatternListNames(0)) Then
            For Each patternListName As String In _PatternListNames
                If patternListName.Equals(TextBoxPatternName.Text) Then
                    MessageBox.Show("パターン名が重複しています。")
                    Exit Sub
                End If
            Next
        End If

        _isOK = True
        Me.Close()

    End Sub

#End Region

End Class
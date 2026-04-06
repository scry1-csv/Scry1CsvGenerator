Public Class TimeSheetRow
    Public Time As Integer
    Public Dir As Boolean
    Public Spd As Integer

    Function ToCsvLine() As String
        Return Time.ToString() & "," & If(Dir, "1", "0") & "," & Spd.ToString()
    End Function
End Class
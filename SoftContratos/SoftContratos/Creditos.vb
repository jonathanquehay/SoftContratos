Public Class Creditos

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Panel1.Top -= 1

        If (Panel1.Location.Y - Panel1.Height <= -760) Then
            Me.Dispose()
        End If

    End Sub

    Private Sub Creditos_Click(sender As Object, e As EventArgs) Handles Me.Click
        Me.Dispose()
    End Sub

    Private Sub Creditos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
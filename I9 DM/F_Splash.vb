Public Class F_Splash
    Private Sub F_Splash_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Form1.Show()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Timer1.Enabled = False
        Timer2.Enabled = True

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        'Fade out the SplashScreen
        Me.Opacity = Me.Opacity - 0.2

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        If Me.Opacity = 0 Then
            Me.Hide()
            Timer1.Stop()
            Timer2.Stop()
            Timer3.Stop()
        End If

    End Sub

End Class
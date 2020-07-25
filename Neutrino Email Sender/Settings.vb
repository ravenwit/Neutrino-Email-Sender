Public Class Settings

    Private Sub Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFont.Text = (My.Settings.DeFont).ToString
        txtFoColor.Text = (My.Settings.DeForColor).ToString
        txtBaColor.Text = (My.Settings.DeBaColor).ToString

    End Sub

    Private Sub btnFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFont.Click
        Dim font As New FontDialog
        Try
            If font.ShowDialog = DialogResult.OK Then
                My.Settings.DeFont = font.Font

            End If
        Catch ex As Exception
            MsgBox("An error occured. Please try again." & vbCrLf & ex.Message)
        End Try
        My.Settings.Save()
        txtFont.Text = (My.Settings.DeFont).ToString

        Neutrino_Email_Sender.rtxtBody.Font = My.Settings.DeFont
        Neutrino_Email_Sender.txtSubject.Font = My.Settings.DeFont
        Neutrino_Email_Sender.cmbTo.Font = My.Settings.DeFont

    End Sub

    Private Sub btnFoColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFoColor.Click
        Dim color As New ColorDialog
        Try
            If color.ShowDialog = DialogResult.OK Then
                My.Settings.DeForColor = color.Color

            End If
        Catch ex As Exception
            MsgBox("An error occured. Please try again." & vbCrLf & ex.Message)
        End Try
        My.Settings.Save()
        txtFoColor.Text = (My.Settings.DeForColor).ToString

       
        Neutrino_Email_Sender.rtxtBody.ForeColor = My.Settings.DeForColor
        Neutrino_Email_Sender.txtSubject.ForeColor = My.Settings.DeForColor
        Neutrino_Email_Sender.cmbTo.ForeColor = My.Settings.DeForColor
    End Sub

    Private Sub btnBaColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBaColor.Click
        Dim color As New ColorDialog
        Try
            If color.ShowDialog = DialogResult.OK Then
                My.Settings.DeBaColor = color.Color

            End If
        Catch ex As Exception
            MsgBox("An error occured. Please try again." & vbCrLf & ex.Message)
        End Try
        My.Settings.Save()
        txtBaColor.Text = (My.Settings.DeBaColor).ToString

        Neutrino_Email_Sender.rtxtBody.BackColor = My.Settings.DeBaColor
        Neutrino_Email_Sender.txtSubject.BackColor = My.Settings.DeBaColor
        Neutrino_Email_Sender.cmbTo.BackColor = My.Settings.DeBaColor
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim font As New FontDialog
        Try
            If font.ShowDialog = DialogResult.OK Then
                My.Settings.PrintFont = font.Font

            End If
        Catch ex As Exception
            MsgBox("An error occured. Please try again." & vbCrLf & ex.Message)
        End Try
        My.Settings.Save()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        My.Settings.Save()
        Me.Close()

    End Sub
End Class
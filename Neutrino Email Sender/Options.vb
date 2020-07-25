Imports System.Net.Mail
Public Class Options

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim num As Integer
        Dim mail As String
        If Not num = InputBox("How many emails you want to add in cc", "Numbers of CC") = vbCancel Then
            For i = 1 To num
                mail = InputBox("Enter cc : ", "CC")
                message.CC.Add(New MailAddress(mail))
                cmbCC.Items.Add(mail)
            Next
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim num As Integer
        Dim mail As String
        num = InputBox("How many emails you want to add in Bcc", "Numbers of Bcc")
        For i = 1 To num
            mail = InputBox("Enter cc : ", "Bcc")
            message.Bcc.Add(New MailAddress(mail))
            cmbBcc.Items.Add(mail)
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim attachfile As New OpenFileDialog
        attachfile.Title = "Attachment File"
        attachfile.Multiselect = False
        Try
            If attachfile.ShowDialog = DialogResult.OK Then

                attachfile1 = attachfile.FileName
                txtAttach1.Text = attachfile1
            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to attach the file. Please try again" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim attachfile As New OpenFileDialog
        attachfile.Title = "Attachment File"
        attachfile.Multiselect = False
        Try
            If attachfile.ShowDialog = DialogResult.OK Then
                
                attachfile2 = attachfile.FileName
                txtAttach2.Text = attachfile2
            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to attach the file. Please try again" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim attachfile As New OpenFileDialog
        attachfile.Title = "Attachment File"
        attachfile.Multiselect = False
        Try
            If attachfile.ShowDialog = DialogResult.OK Then
               
                attachfile3 = attachfile.FileName
                txtAttach3.Text = attachfile3
            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to attach the file. Please try again" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim attachfile As New OpenFileDialog
        attachfile.Title = "Attachment File"
        attachfile.Multiselect = False
        Try
            If attachfile.ShowDialog = DialogResult.OK Then
                
                attachfile4 = attachfile.FileName
                txtAttach4.Text = attachfile4
            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to attach the file. Please try again" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If txtAttach1.Text <> vbNullString And txtAttach1.Text = attachfile1 Then
            Dim attach1 As New Attachment(attachfile1)
            message.Attachments.Add(attach1)
        ElseIf txtAttach1.Text <> vbNullString And txtAttach1.Text <> attachfile1 Then
            If MsgBox("Check your 1st file path.", MsgBoxStyle.OkCancel) = vbOK Then
                Exit Sub
            Else
                Resume
            End If
        End If
        If txtAttach2.Text <> vbNullString And txtAttach2.Text = attachfile2 Then
            Dim attach2 As New Attachment(attachfile2)
            message.Attachments.Add(attach2)
        ElseIf txtAttach2.Text <> vbNullString And txtAttach2.Text <> attachfile2 Then
            If MsgBox("Check your 2nd file path.", MsgBoxStyle.OkCancel) = vbOK Then
                Exit Sub
            Else
                Resume
            End If
        End If
        If txtAttach3.Text <> vbNullString And txtAttach3.Text = attachfile3 Then
            Dim attach3 As New Attachment(attachfile3)
            message.Attachments.Add(attach3)
        ElseIf txtAttach3.Text <> vbNullString And txtAttach3.Text <> attachfile3 Then
            If MsgBox("Check your 3rd file path.", MsgBoxStyle.OkCancel) = vbOK Then
                Exit Sub
            Else
                Resume
            End If
        End If
        If txtAttach4.Text <> vbNullString And txtAttach4.Text = attachfile4 Then
            Dim attach4 As New Attachment(attachfile4)
            message.Attachments.Add(attach4)
        ElseIf txtAttach4.Text <> vbNullString And txtAttach4.Text <> attachfile4 Then
            If MsgBox("Check your 4th file path.", MsgBoxStyle.OkCancel) = vbOK Then
                Exit Sub
            Else
                Resume
            End If
        End If
        Me.Close()
    End Sub

    Private Sub txtAttach1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAttach1.TextChanged

    End Sub
End Class
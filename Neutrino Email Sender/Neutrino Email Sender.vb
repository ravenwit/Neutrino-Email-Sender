Imports System.Net.Mail
Imports System.Drawing.Printing
Imports System.IO

Public Class Neutrino_Email_Sender
   
    Public isSave As Integer = 0
    Dim webpre As New WebBrowser
    Dim i As Integer = 0


    

    Private Sub NewEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewEmailToolStripMenuItem.Click
        Dim want
        If rtxtBody.Modified = True Then
            want = MsgBox("Do you want to save current mail ?", MsgBoxStyle.Information, "Save")

            If want = vbYes Then SaveEmailToolStripMenuItem.PerformClick()
            If want = vbNo Then rtxtBody.Clear()
            If want = vbCancel Then Exit Sub
            isSave = 0

        End If


    End Sub

    Private Sub OpenEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenEmailToolStripMenuItem.Click
        Dim openmail As New OpenFileDialog
        Dim Reader As System.IO.StreamReader

        openmail.DefaultExt = "Text |*.txt"
        openmail.AddExtension = True
        openmail.Multiselect = False
        openmail.RestoreDirectory = True
        openmail.Title = "Open Mail"
        Try
            If openmail.ShowDialog = DialogResult.OK Then
                Reader = New System.IO.StreamReader(openmail.FileName)
                Me.rtxtBody.Text = Reader.ReadToEnd
                'rtxtBody.LoadFile(openmail.FileName, RichTextBoxStreamType.RichText)
                Reader.Close()
                Reader = Nothing
            End If

        Catch ex As Exception
            MsgBox("Sorry ! An error occured to open file. Please try again" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")

        End Try


    End Sub


    Private Sub SaveEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveEmailToolStripMenuItem.Click
        Dim savemail As New SaveFileDialog
        Dim Writer As System.IO.StreamWriter

        savemail.AddExtension = True
        savemail.DefaultExt = "rtf"
        savemail.FileName = "Neutrino Email Sender - Email"
        savemail.RestoreDirectory = True
        savemail.Title = "Save Email"
        Try
            If savemail.ShowDialog = DialogResult.OK Then
                Writer = New System.IO.StreamWriter(savemail.FileName)
                Writer.Write(Me.rtxtBody.Text)
                'rtxtBody.SaveFile(savemail.FileName, RichTextBoxStreamType.RichText)
                Writer.Close()
                Writer = Nothing
            End If

        Catch ex As Exception
            MsgBox("Sorry ! An error occured to save mail" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
        isSave = 1
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        rtxtBody.Undo()

    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        rtxtBody.Redo()

    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        rtxtBody.Cut()

    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        rtxtBody.Copy()

    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        rtxtBody.Paste()

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        rtxtBody.SelectedRtf = vbNullString

    End Sub

    Private Sub MainToolbarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MainToolbarToolStripMenuItem.Click
        If MainToolbarToolStripMenuItem.Checked = True Then
            MainToolbarToolStripMenuItem.Checked = False
            MainToolbar.Visible = False
        ElseIf MainToolbarToolStripMenuItem.Checked = False Then
            MainToolbarToolStripMenuItem.Checked = True
            MainToolbar.Visible = True
        End If
    End Sub

    Private Sub EditorToolbarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditorToolbarToolStripMenuItem.Click
        If EditorToolbarToolStripMenuItem.Checked = True Then
            EditorToolbarToolStripMenuItem.Checked = False
            EditToolbar.Visible = False
        ElseIf EditorToolbarToolStripMenuItem.Checked = False Then
            EditorToolbarToolStripMenuItem.Checked = True
            EditToolbar.Visible = True
        End If
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click
        If StatusBarToolStripMenuItem.Checked = True Then
            StatusBarToolStripMenuItem.Checked = False
            Statusbar.Visible = False
        ElseIf StatusBarToolStripMenuItem.Checked = False Then
            StatusBarToolStripMenuItem.Checked = True
            Statusbar.Visible = True
        End If
    End Sub

    Private Sub BoldToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BoldToolStripMenuItem.Click
        If BoldToolStripMenuItem.Checked = True Then
            BoldToolStripMenuItem.Checked = False
            btnToolBold.CheckState = CheckState.Unchecked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Bold)
        ElseIf BoldToolStripMenuItem.Checked = False Then
            BoldToolStripMenuItem.Checked = True
            btnToolBold.CheckState = CheckState.Checked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Bold)
        End If

    End Sub

    Private Sub ItalicToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItalicToolStripMenuItem.Click
        If ItalicToolStripMenuItem.Checked = True Then
            ItalicToolStripMenuItem.Checked = False
            btnToolItalic.CheckState = CheckState.Unchecked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Italic)
        ElseIf ItalicToolStripMenuItem.Checked = False Then
            ItalicToolStripMenuItem.Checked = True
            btnToolItalic.CheckState = CheckState.Checked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Italic)
        End If
    End Sub

    Private Sub UnderlineToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnderlineToolStripMenuItem.Click
        If UnderlineToolStripMenuItem.Checked = True Then
            UnderlineToolStripMenuItem.Checked = False
            btnToolUnderline.CheckState = CheckState.Unchecked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Underline)
        ElseIf UnderlineToolStripMenuItem.Checked = False Then
            UnderlineToolStripMenuItem.Checked = True
            btnToolUnderline.CheckState = CheckState.Checked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Underline)
        End If
    End Sub

    Private Sub RegularToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegularToolStripMenuItem.Click
        btnToolBold.CheckState = CheckState.Unchecked
        btnToolItalic.CheckState = CheckState.Unchecked
        btnToolUnderline.CheckState = CheckState.Unchecked
        btnToolStrike.CheckState = CheckState.Unchecked
        BoldToolStripMenuItem.Checked = False
        ItalicToolStripMenuItem.Checked = False
        UnderlineToolStripMenuItem.Checked = False
        StrikethroughToolStripMenuItem.Checked = False
        rtxtBody.Font = New Font(rtxtBody.Font, FontStyle.Regular)
    End Sub

    Private Sub StrikethroughToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StrikethroughToolStripMenuItem.Click
        If StrikethroughToolStripMenuItem.Checked = True Then
            StrikethroughToolStripMenuItem.Checked = False
            btnToolStrike.CheckState = CheckState.Unchecked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Strikeout)
        ElseIf StrikethroughToolStripMenuItem.Checked = False Then
            StrikethroughToolStripMenuItem.Checked = True
            btnToolStrike.CheckState = CheckState.Checked
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Strikeout)
        End If
    End Sub

    Private Sub FontForegroundColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontForegroundColorToolStripMenuItem.Click
        Dim forecolor As New ColorDialog
        Try
            If forecolor.ShowDialog = DialogResult.OK Then
                If Not IsNothing(rtxtBody.SelectedRtf) Then rtxtBody.SelectionColor = forecolor.Color Else rtxtBody.ForeColor = forecolor.Color

            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to select color" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub AlignLeftToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlignLeftToolStripMenuItem.Click
        If AlignLeftToolStripMenuItem.Checked = True Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf AlignLeftToolStripMenuItem.Checked = False Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignRightToolStripMenuItem.Checked = False
            AlignCenterToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        End If
        If AlignLeftToolStripMenuItem.Checked = True Then
            btnToolAlignLeft.CheckState = CheckState.Checked
        ElseIf AlignCenterToolStripMenuItem.Checked = True Then
            btnToolAlignCenter.CheckState = CheckState.Checked
        ElseIf AlignRightToolStripMenuItem.Checked = True Then
            btnToolAlignRight.CheckState = CheckState.Checked
        End If
    End Sub

    Private Sub AlignCenterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlignCenterToolStripMenuItem.Click
        If AlignCenterToolStripMenuItem.Checked = True Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf AlignCenterToolStripMenuItem.Checked = False Then
            AlignCenterToolStripMenuItem.Checked = True
            AlignLeftToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Center

        End If
        If AlignLeftToolStripMenuItem.Checked = True Then
            btnToolAlignLeft.CheckState = CheckState.Checked
        ElseIf AlignCenterToolStripMenuItem.Checked = True Then
            btnToolAlignCenter.CheckState = CheckState.Checked
        ElseIf AlignRightToolStripMenuItem.Checked = True Then
            btnToolAlignRight.CheckState = CheckState.Checked
        End If
    End Sub

    Private Sub AlignRightToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlignRightToolStripMenuItem.Click
        If AlignRightToolStripMenuItem.Checked = True Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf AlignRightToolStripMenuItem.Checked = False Then
            AlignRightToolStripMenuItem.Checked = True
            AlignLeftToolStripMenuItem.Checked = False
            AlignCenterToolStripMenuItem.Checked = False
            rtxtBody.SelectionAlignment = HorizontalAlignment.Right


        End If
        If AlignLeftToolStripMenuItem.Checked = True Then
            btnToolAlignLeft.CheckState = CheckState.Checked
        ElseIf AlignCenterToolStripMenuItem.Checked = True Then
            btnToolAlignCenter.CheckState = CheckState.Checked
        ElseIf AlignRightToolStripMenuItem.Checked = True Then
            btnToolAlignRight.CheckState = CheckState.Checked
        End If
    End Sub

    Private Sub BackgroundColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackgroundColorToolStripMenuItem.Click
        Dim backcolor As New ColorDialog
        Try
            If backcolor.ShowDialog = DialogResult.OK Then
                If rtxtBody.SelectedRtf <> "" Then rtxtBody.SelectionBackColor = backcolor.Color Else rtxtBody.BackColor = backcolor.Color
            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to select color. Please try again." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub SendEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendEmailToolStripMenuItem.Click
        If My.Settings.ToEmail.Contains(cmbTo.Text) = False Then
            My.Settings.ToEmail.Add(cmbTo.Text)
        End If

        Dim msg
        Try
            smtpServer.Credentials = New Net.NetworkCredential(My.Settings.ID, My.Settings.Pass)
            smtpServer.EnableSsl = True
            smtpServer.Port = My.Settings.Port
            smtpServer.Host = (My.Settings.Host).ToString


            message.From = New MailAddress(My.Settings.ID, My.Settings.UserName)
            message.To.Add(cmbTo.Text.ToString)
            message.Body = rtxtBody.Text
            message.Subject = txtSubject.Text

            Try
                lblStatus.Text = "Sending Email..."
                lblProgress.Visible = True
                smtpServer.Send(message)
                msg = MsgBox("Sent Successful")
                If msg = vbOK Then
                    lblProgress.Visible = False
                    lblStatus.Text = "Ready"
                End If

            Catch ex As Exception
                MsgBox("Error" & vbCrLf & ex.Message)
            End Try
        Catch ex As Exception
            MsgBox("Error" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        lblProgress.Visible = False
        lblStatus.Text = "Ready"
        cmbTo.Items.Clear()

        For Each item As String In My.Settings.ToEmail
            cmbTo.Items.Add(item)
        Next
        
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        My.Settings.Login = 0
        LogIn()


    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        NewEmailToolStripMenuItem.PerformClick()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        OpenEmailToolStripMenuItem.PerformClick()

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        SaveEmailToolStripMenuItem.PerformClick()

    End Sub

    Private Sub Neutrino_Email_Sender_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim want
        If rtxtBody.Modified = True Then
            want = MsgBox("Do you want to save current mail ?", MsgBoxStyle.Information + vbYesNoCancel, "Save")
            If want = vbYes Then SaveEmailToolStripMenuItem.PerformClick()
            If want = vbNo Then Exit Sub
            If want = vbCancel Then
                e.Cancel = True

            End If
        End If
    End Sub

    Private Sub Neutrino_Email_Sender_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LogIn()
        lblToolEmail.Text = MainID

        For Each Item As String In My.Settings.ToEmail
            cmbTo.Items.Add(Item)
        Next
        lblStatus.Text = "Ready"
        lblProgress.Visible = False
        chkHTML.Checked = False

        rtxtBody.DetectUrls = True
        rtxtBody.Multiline = True
        rtxtBody.HideSelection = True

        rtxtBody.BackColor = My.Settings.DeBaColor
        rtxtBody.Font = My.Settings.DeFont
        rtxtBody.ForeColor = My.Settings.DeForColor

        txtSubject.BackColor = My.Settings.DeBaColor
        txtSubject.Font = My.Settings.DeFont
        txtSubject.ForeColor = My.Settings.DeForColor

        cmbTo.BackColor = My.Settings.DeBaColor
        cmbTo.Font = My.Settings.DeFont
        cmbTo.ForeColor = My.Settings.DeForColor

        btnToolBold.CheckOnClick = True
        btnToolBold.CheckState = False
        btnToolItalic.CheckOnClick = True
        btnToolItalic.CheckState = False
        btnToolUnderline.CheckOnClick = True
        btnToolUnderline.CheckState = False
        btnToolAlignCenter.CheckOnClick = True
        btnToolAlignCenter.CheckState = False
        btnToolAlignLeft.CheckOnClick = True
        btnToolAlignLeft.CheckState = True
        btnToolAlignRight.CheckOnClick = True
        btnToolAlignRight.CheckState = False
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        SendEmailToolStripMenuItem.PerformClick()

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        LogoutToolStripMenuItem.PerformClick()


    End Sub

    Private Sub btnToolUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolUndo.Click
        UndoToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolRedo.Click
        RedoToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolDelete.Click
        DeleteToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolCut.Click
        CutToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolCopy.Click
        CopyToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolPaste.Click
        PasteToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolFont.Click
        FontToolStripMenuItem.PerformClick()

    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        Dim font As New FontDialog

        Try
            If font.ShowDialog = DialogResult.OK Then
                If Not IsNothing(rtxtBody.SelectedRtf) Then Me.rtxtBody.SelectionFont = font.Font Else rtxtBody.Font = font.Font

            End If
        Catch ex As Exception
            MsgBox("Sorry ! An error occured to select font. Please try again." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub btnToolFColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolFColor.Click
        FontForegroundColorToolStripMenuItem.PerformClick()

    End Sub

    Private Sub btnToolBold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolBold.Click
        If btnToolBold.CheckState = CheckState.Unchecked Then
            BoldToolStripMenuItem.Checked = False
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Bold)
        ElseIf btnToolBold.CheckState = CheckState.Checked Then
            BoldToolStripMenuItem.Checked = True
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Bold)
        End If

    End Sub

    Private Sub btnToolItalic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolItalic.Click
        If btnToolItalic.CheckState = CheckState.Unchecked Then
            ItalicToolStripMenuItem.Checked = False
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Italic)

        ElseIf btnToolItalic.CheckState = CheckState.Checked Then

            ItalicToolStripMenuItem.Checked = True
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Italic)
        End If

    End Sub

    Private Sub btnToolUnderline_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolStrike.Click
        If btnToolStrike.CheckState = CheckState.Unchecked Then
            StrikethroughToolStripMenuItem.Checked = False
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Strikeout)

        ElseIf btnToolStrike.CheckState = CheckState.Checked Then

            StrikethroughToolStripMenuItem.Checked = True
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Strikeout)
        End If


    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolUnderline.Click
        If btnToolUnderline.CheckState = CheckState.Unchecked Then
            UnderlineToolStripMenuItem.Checked = False
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style - FontStyle.Underline)

        ElseIf btnToolUnderline.CheckState = CheckState.Checked Then

            UnderlineToolStripMenuItem.Checked = True
            rtxtBody.SelectionFont = New Font(rtxtBody.SelectionFont, rtxtBody.SelectionFont.Style + FontStyle.Underline)
        End If
    End Sub

    Private Sub btnToolAlignLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolAlignLeft.Click
        If btnToolAlignLeft.CheckState = CheckState.Unchecked Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            btnToolAlignLeft.CheckState = CheckState.Checked
            btnToolAlignCenter.CheckState = CheckState.Checked
            btnToolAlignRight.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf btnToolAlignLeft.CheckState = CheckState.Checked Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            btnToolAlignCenter.CheckState = CheckState.Unchecked
            btnToolAlignRight.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        End If

    End Sub

    Private Sub btnToolAlignCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolAlignCenter.Click
        If btnToolAlignCenter.CheckState = CheckState.Unchecked Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            btnToolAlignLeft.CheckState = CheckState.Checked
            btnToolAlignRight.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf btnToolAlignCenter.CheckState = CheckState.Checked Then
            AlignLeftToolStripMenuItem.Checked = False
            AlignCenterToolStripMenuItem.Checked = True
            AlignRightToolStripMenuItem.Checked = False
            btnToolAlignLeft.CheckState = CheckState.Unchecked
            btnToolAlignRight.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Center
        End If


    End Sub

    Private Sub btnToolAlignRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolAlignRight.Click
        If btnToolAlignRight.CheckState = CheckState.Unchecked Then
            AlignLeftToolStripMenuItem.Checked = True
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = False
            btnToolAlignLeft.CheckState = CheckState.Checked
            btnToolAlignCenter.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Left
        ElseIf btnToolAlignRight.CheckState = CheckState.Checked Then
            AlignLeftToolStripMenuItem.Checked = False
            AlignCenterToolStripMenuItem.Checked = False
            AlignRightToolStripMenuItem.Checked = True
            btnToolAlignLeft.CheckState = CheckState.Unchecked
            btnToolAlignCenter.CheckState = CheckState.Unchecked
            rtxtBody.SelectionAlignment = HorizontalAlignment.Right
        End If

    End Sub

    Private Sub btnToolBackColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToolBackColor.Click
        BackgroundColorToolStripMenuItem.PerformClick()

    End Sub

    Private Sub chkHTML_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHTML.CheckedChanged
        On Error Resume Next
        webpre.Name = "Preview Browser"
        webpre.Dock = DockStyle.Fill
        If chkHTML.Checked = True Then
            message.IsBodyHtml = True
            TabControl1.TabPages.Add("Browser Preview")
            TabControl1.TabPages.Item(1).Controls.Add(webpre)
            CType(TabControl1.TabPages.Item(1).Controls.Item(0), WebBrowser).DocumentText = rtxtBody.Text

        ElseIf chkHTML.Checked = False Then
            message.IsBodyHtml = False
            TabControl1.TabPages.RemoveAt(1)

        End If
    End Sub

    Private Sub rtxtBody_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rtxtBody.TextChanged
        On Error Resume Next
        If chkHTML.Checked = True Then CType(TabControl1.TabPages.Item(1).Controls.Item(0), WebBrowser).DocumentText = (rtxtBody.Text).ToString
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

   
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        txtSubject.Width = Panel1.Width - 140
    End Sub

    Private Sub EditToolbar_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles EditToolbar.ItemClicked

    End Sub

    Private Sub ToolStripButton6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Options.Show()

    End Sub

    Private Sub DateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateToolStripMenuItem.Click
        rtxtBody.SelectedText = DateAndTime.DateString
    End Sub

   
    Private Sub TimeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeToolStripMenuItem.Click
        rtxtBody.SelectedText = DateAndTime.TimeString
    End Sub

    Private Sub UserNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserNameToolStripMenuItem.Click
        rtxtBody.SelectedText = (My.Settings.UserName).ToString
    End Sub

    Private Sub EmailAddressToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailAddressToolStripMenuItem.Click
        rtxtBody.SelectedText = (My.Settings.ID).ToString
    End Sub

    Private Sub DateAndTimeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateAndTimeToolStripMenuItem.Click
        rtxtBody.SelectedText = (Now).ToString
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        lbltoolDate.Text = FormatDateTime(DateAndTime.Today, DateFormat.LongDate)
        lbltooltime.Text = FormatDateTime(DateAndTime.TimeOfDay, DateFormat.LongTime)

        If rtxtBody.CanUndo = True Then
            UndoToolStripMenuItem.Enabled = True
            btnToolUndo.Enabled = True
        Else
            UndoToolStripMenuItem.Enabled = False
            btnToolUndo.Enabled = False
        End If
        If rtxtBody.CanRedo = True Then
            RedoToolStripMenuItem.Enabled = True
            btnToolRedo.Enabled = True
        Else
            RedoToolStripMenuItem.Enabled = False
            btnToolRedo.Enabled = False
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        If SelectAllToolStripMenuItem.Checked = True Then
            SelectAllToolStripMenuItem.Checked = False
            rtxtBody.DeselectAll()
        ElseIf SelectAllToolStripMenuItem.Checked = False Then
            SelectAllToolStripMenuItem.Checked = True
            rtxtBody.SelectAll()

        End If

    End Sub

    Private Sub FindToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PlainTextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlainTextToolStripMenuItem.Click
        Dim savemail As New SaveFileDialog
        savemail.AddExtension = True
        savemail.DefaultExt = "txt"
        savemail.FileName = "Neutrino Email Sender - Email"
        savemail.RestoreDirectory = True
        savemail.Title = "Save Email"
        Try
            If savemail.ShowDialog = DialogResult.OK Then
                rtxtBody.SaveFile(savemail.FileName, RichTextBoxStreamType.PlainText)
            End If

        Catch ex As Exception
            MsgBox("Sorry ! An error occured to save mail" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
        isSave = 1
    End Sub

    Private Sub RichTextFormatToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextFormatToolStripMenuItem.Click
        Dim savemail As New SaveFileDialog
        savemail.AddExtension = True
        savemail.DefaultExt = "rtf"
        savemail.FileName = "Neutrino Email Sender - Email"
        savemail.RestoreDirectory = True
        savemail.Title = "Save Email"
        Try
            If savemail.ShowDialog = DialogResult.OK Then
                rtxtBody.SaveFile(savemail.FileName, RichTextBoxStreamType.RichText)
            End If

        Catch ex As Exception
            MsgBox("Sorry ! An error occured to save mail" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
        isSave = 1
    End Sub

    Private Sub WordDocumentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordDocumentToolStripMenuItem.Click
        Dim savemail As New SaveFileDialog
        savemail.AddExtension = True
        savemail.DefaultExt = "doc"
        savemail.FileName = "Neutrino Email Sender - Email"
        savemail.RestoreDirectory = True
        savemail.Title = "Save Email"
        Try
            If savemail.ShowDialog = DialogResult.OK Then
                rtxtBody.SaveFile(savemail.FileName, RichTextBoxStreamType.RichText)
            End If

        Catch ex As Exception
            MsgBox("Sorry ! An error occured to save mail" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
        isSave = 1
    End Sub



    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

    End Sub

  
    Private Sub pdoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Declare a variable to hold the position of the last printed char. Declare
        ' as static so that subsequent PrintPage events can reference it.
        Static intCurrentChar As Int32
        ' Initialize the font to be used for printing.
        Dim font As Font = My.Settings.PrintFont

        Dim intPrintAreaHeight, intPrintAreaWidth, marginLeft, marginTop As Int32
        With PrintDocument1.DefaultPageSettings
            ' Initialize local variables that contain the bounds of the printing 
            ' area rectangle.
            intPrintAreaHeight = .PaperSize.Height - .Margins.Top - .Margins.Bottom
            intPrintAreaWidth = .PaperSize.Width - .Margins.Left - .Margins.Right

            ' Initialize local variables to hold margin values that will serve
            ' as the X and Y coordinates for the upper left corner of the printing 
            ' area rectangle.
            marginLeft = .Margins.Left ' X coordinate
            marginTop = .Margins.Top ' Y coordinate
        End With

        ' If the user selected Landscape mode, swap the printing area height 
        ' and width.
        If PrintDocument1.DefaultPageSettings.Landscape Then
            Dim intTemp As Int32
            intTemp = intPrintAreaHeight
            intPrintAreaHeight = intPrintAreaWidth
            intPrintAreaWidth = intTemp
        End If

        ' Calculate the total number of lines in the document based on the height of
        ' the printing area and the height of the font.
        Dim intLineCount As Int32 = CInt(intPrintAreaHeight / font.Height)
        ' Initialize the rectangle structure that defines the printing area.
        Dim rectPrintingArea As New RectangleF(marginLeft, marginTop, intPrintAreaWidth, intPrintAreaHeight)

        ' Instantiate the StringFormat class, which encapsulates text layout 
        ' information (such as alignment and line spacing), display manipulations 
        ' (such as ellipsis insertion and national digit substitution) and OpenType 
        ' features. Use of StringFormat causes MeasureString and DrawString to use
        ' only an integer number of lines when printing each page, ignoring partial
        ' lines that would otherwise likely be printed if the number of lines per 
        ' page do not divide up cleanly for each page (which is usually the case).
        ' See further discussion in the SDK documentation about StringFormatFlags.
        Dim fmt As New StringFormat(StringFormatFlags.LineLimit)
        ' Call MeasureString to determine the number of characters that will fit in
        ' the printing area rectangle. The CharFitted Int32 is passed ByRef and used
        ' later when calculating intCurrentChar and thus HasMorePages. LinesFilled 
        ' is not needed for this sample but must be passed when passing CharsFitted.
        ' Mid is used to pass the segment of remaining text left off from the 
        ' previous page of printing (recall that intCurrentChar was declared as 
        ' static.
        Dim intLinesFilled, intCharsFitted As Int32
        e.Graphics.MeasureString(Mid(rtxtBody.Text, intCurrentChar + 1), font, New SizeF(intPrintAreaWidth, intPrintAreaHeight), fmt, intCharsFitted, intLinesFilled)

        ' Print the text to the page.
        e.Graphics.DrawString(Mid(rtxtBody.Text, intCurrentChar + 1), font, Brushes.Black, rectPrintingArea, fmt)

        ' Advance the current char to the last char printed on this page. As 
        ' intCurrentChar is a static variable, its value can be used for the next
        ' page to be printed. It is advanced by 1 and passed to Mid() to print the
        ' next page (see above in MeasureString()).
        intCurrentChar += intCharsFitted

        ' HasMorePages tells the printing module whether another PrintPage event
        ' should be fired.
        If intCurrentChar < rtxtBody.Text.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            ' You must explicitly reset intCurrentChar as it is static.
            intCurrentChar = 0
        End If
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Try
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
        Catch exp As Exception
            MsgBox(exp.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub PrintDialogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintDialogToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub PageSetupToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageSetupToolStripMenuItem1.Click
        With PageSetupDialog1
            .Document = PrintDocument1
            .PageSettings = PrintDocument1.DefaultPageSettings
        End With

        If PageSetupDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings
        End If
    End Sub

  
    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        rtxtBody.Refresh()

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        RefreshToolStripMenuItem.PerformClick()

    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Settings.Show()

    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Options.Show()

    End Sub

    Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        Settings.Show()

    End Sub

    Private Sub FindToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindToolStripMenuItem.Click

        Find.Show()

    End Sub
End Class

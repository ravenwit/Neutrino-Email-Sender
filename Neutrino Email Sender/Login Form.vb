Public Class Login_Form

   
    Private Sub cmbHost_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHost.SelectedIndexChanged
        Select Case cmbHost.SelectedItem
            Case "Gmail"
                My.Settings.Host = "smtp.gmail.com"
                My.Settings.Port = 587
            Case "Yahoo"
                My.Settings.Host = "smtp.mail.yahoo.com"
                My.Settings.Port = 25
            Case "Hotmail"
                My.Settings.Host = "smtp.live.com"
                My.Settings.Port = 25
            Case "Lycos"
                My.Settings.Host = "smtp.mail.lycos.com"
                My.Settings.Port = 25
            Case "Aol"
                My.Settings.Host = "smtp.aol.com"
                My.Settings.Port = 587
            Case "Rambler.ru"
                My.Settings.Host = "smtp.rambler.ru"
                My.Settings.Port = 25
            Case "Mail.ru"
                My.Settings.Host = "smtp.mail.ru"
                My.Settings.Port = 2525
            Case "Yandex.ru"
                My.Settings.Host = "smtp.yandex.ru"
                My.Settings.Port = 25
            Case "Meta.ua"
                My.Settings.Host = "smtp.meta.ua"
                My.Settings.Port = 465
            Case "Ukr.net"
                My.Settings.Host = "smtp.ukr.net"
                My.Settings.Port = 465
        End Select
        My.Settings.Save()

    End Sub

    Private Sub Login_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbHost.SelectedItem = "Gmail"
        txtUser.Text = My.Settings.UserName
        If GmID <> vbNullString Then cmbID.Items.Add(GmID)
        If YahID <> vbNullString Then cmbID.Items.Add(YahID)
        If HotID <> vbNullString Then cmbID.Items.Add(HotID)
        If LyID <> vbNullString Then cmbID.Items.Add(LyID)
        If AolID <> vbNullString Then cmbID.Items.Add(AolID)
        If RamID <> vbNullString Then cmbID.Items.Add(RamID)
        If MailID <> vbNullString Then cmbID.Items.Add(MailID)
        If YanID <> vbNullString Then cmbID.Items.Add(YanID)
        If MetaID <> vbNullString Then cmbID.Items.Add(MetaID)
        If UkID <> vbNullString Then cmbID.Items.Add(UkID)

    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If cmbID.Text <> vbNullString And txtPass.Text <> vbNullString Then
            My.Settings.ID = cmbID.Text
            My.Settings.Pass = txtPass.Text
            My.Settings.UserName = (txtUser.Text).ToString
            Neutrino_Email_Sender.Show()
            Neutrino_Email_Sender.lblToolEmail.Text = My.Settings.ID
            My.Settings.Login = 1
            My.Settings.Save()
            Me.Close()
        Else
            If cmbID.Text = vbNullString Then
                MsgBox("Email ID is required.")
            ElseIf txtPass.Text = vbNullString Then
                MsgBox("Password is required.")
            End If
        End If
        
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Select Case cmbHost.SelectedItem
            Case "Gmail"
                My.Settings.GMID = cmbID.Text
                My.Settings.GMPass = txtPass.Text

            Case "Yahoo"
                My.Settings.YahID = cmbID.Text
                My.Settings.YahPass = txtPass.Text
            Case "Hotmail"
                My.Settings.HotID = cmbID.Text
                My.Settings.HotPass = txtPass.Text
            Case "Lycos"
                My.Settings.LyID = cmbID.Text
                My.Settings.LyPass = txtPass.Text
            Case "Aol"
                My.Settings.AolID = cmbID.Text
                My.Settings.AolPass = txtPass.Text
            Case "Rambler.ru"
                My.Settings.RamID = cmbID.Text
                My.Settings.RamPass = txtPass.Text
            Case "Mail.ru"
                My.Settings.MailID = cmbID.Text
                My.Settings.MailPass = txtPass.Text
            Case "Yandex.ru"
                My.Settings.YanID = cmbID.Text
                My.Settings.YanPass = txtPass.Text
            Case "Meta.ua"
                My.Settings.MetaID = cmbID.Text
                My.Settings.MetaPass = txtPass.Text
            Case "Ukr.net"
                My.Settings.UkID = cmbID.Text
                My.Settings.UkPass = txtPass.Text
        End Select
        My.Settings.Save()
    End Sub

    Private Sub cmbID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbID.SelectedIndexChanged
        CheckEnter()
        My.Settings.Save()


    End Sub
End Class
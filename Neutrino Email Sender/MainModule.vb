Imports System.Net.Mail
Module MainModule
    Public message As New MailMessage
    Public smtpServer As New SmtpClient

    Public MainID As String = My.Settings.ID
    Public MainPass As String = My.Settings.Pass

    Public attachfile1 As String
    Public attachfile2 As String
    Public attachfile3 As String
    Public attachfile4 As String

    Public GmID As String = My.Settings.GMID
    Public GmPass As String = My.Settings.GMPass
    Public YahID As String = My.Settings.YahID
    Public YahPass As String = My.Settings.YahPass
    Public HotID As String = My.Settings.HotID
    Public HotPass As String = My.Settings.HotPass
    Public LyID As String = My.Settings.LyID
    Public LyPass As String = My.Settings.LyPass
    Public AolID As String = My.Settings.AolID
    Public AolPass As String = My.Settings.AolPass
    Public RamID As String = My.Settings.RamID
    Public RamPass As String = My.Settings.RamPass
    Public MailID As String = My.Settings.MailID
    Public MailPass As String = My.Settings.MailPass
    Public YanID As String = My.Settings.YanID
    Public YanPass As String = My.Settings.YanPass
    Public MetaID As String = My.Settings.MetaID
    Public MetaPass As String = My.Settings.MetaPass
    Public UkID As String = My.Settings.UkID
    Public UkPass As String = My.Settings.UkPass


    Public Function sendMail(ByVal Email As String, ByVal Password As String, ByVal Email_To As String, ByVal Body As String, ByVal Attachment As String, ByVal Port As Integer, ByVal Host As String)
        message.From = New MailAddress(Email)
        message.To.Add(Email_To)
        message.Body = Body
        smtpServer.EnableSsl = True
        smtpServer.Port = Port
        smtpServer.Host = Host
        smtpServer.Credentials = New Net.NetworkCredential(Email, Password)
        If Not IsNothing(Attachment) Then
            Dim attachfile As New Attachment(Attachment)
            message.Attachments.Add(attachfile)

        End If
        Try
            smtpServer.Send(message)
            MsgBox("Sent Successful")

        Catch ex As Exception
            MsgBox("Error")
        End Try
        Return True
    End Function
    Public Function CheckEnter()

        If Login_Form.cmbID.SelectedItem = GmID Then
            Login_Form.cmbHost.SelectedItem = "Gmail"
            Login_Form.txtPass.Text = GmPass
            My.Settings.Host = "smtp.gmail.com"
            My.Settings.Port = 587
        End If
        If Login_Form.cmbID.SelectedItem = YahID Then
            Login_Form.cmbHost.SelectedItem = "Yahoo"
            Login_Form.txtPass.Text = YahPass
            My.Settings.Host = "smtp.mail.yahoo.com"
            My.Settings.Port = 25
        End If
        If Login_Form.cmbID.SelectedItem = HotID Then
            Login_Form.cmbHost.SelectedItem = "Hotmail"
            Login_Form.txtPass.Text = HotPass
            My.Settings.Host = "smtp.live.com"
            My.Settings.Port = 25
        End If
        If Login_Form.cmbID.SelectedItem = LyID Then
            Login_Form.cmbHost.SelectedItem = "Lycos"
            Login_Form.txtPass.Text = LyPass
            My.Settings.Host = "smtp.mail.lycos.com"
            My.Settings.Port = 25
        End If
        If Login_Form.cmbID.SelectedItem = AolID Then
            Login_Form.cmbHost.SelectedItem = "Aol"
            Login_Form.Text = AolPass
            My.Settings.Host = "smtp.aol.com"
            My.Settings.Port = 587
        End If
        If Login_Form.cmbID.SelectedItem = RamID Then
            Login_Form.cmbHost.SelectedItem = "Rambler.ru"
            Login_Form.Text = RamPass
            My.Settings.Host = "smtp.rambler.ru"
            My.Settings.Port = 25
        End If
        If Login_Form.cmbID.SelectedItem = MailID Then
            Login_Form.cmbHost.SelectedItem = "Mail.ru"
            Login_Form.Text = MailPass
            My.Settings.Host = "smtp.mail.ru"
            My.Settings.Port = 2525
        End If
        If Login_Form.cmbID.SelectedItem = YanID Then
            Login_Form.cmbHost.SelectedItem = "Yandex.ru"
            Login_Form.Text = YanPass
            My.Settings.Host = "smtp.yandex.ru"
            My.Settings.Port = 25
        End If
        If Login_Form.cmbID.SelectedItem = MetaID Then
            Login_Form.cmbHost.SelectedItem = "Meta.ua"
            Login_Form.Text = MetaPass
            My.Settings.Host = "smtp.meta.ua"
            My.Settings.Port = 465
        End If
        If Login_Form.cmbID.SelectedItem = UkID Then
            Login_Form.cmbHost.SelectedItem = "Ukr.net"
            Login_Form.Text = UkPass
            My.Settings.Host = "smtp.ukr.net"
            My.Settings.Port = 465
        End If
        Return True
    End Function

    Public Function LogIn()
        If My.Settings.Login = 0 Then
            Neutrino_Email_Sender.Hide()
            Login_Form.Show()
        Else
            Neutrino_Email_Sender.Show()
            Login_Form.Close()
        End If
        Return True
    End Function
End Module

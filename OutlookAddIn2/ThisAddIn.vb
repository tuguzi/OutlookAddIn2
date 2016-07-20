'名前空間
Imports System.Windows.Forms


Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'MessageBox.Show("ThisAddIn_Startup")
        'MessageBox.Show(System.Environment.GetEnvironmentVariable("path"))
        'MessageBox.Show(System.Environment.GetEnvironmentVariable("nitoms_shared_address"))
    End Sub



    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Handlesでメール送信イベントを利用する
    Private Sub Application_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend

        Dim sharedAdr As String = "sb-nitoms-it-servicedesk@nitto.com"
        'Dim sharedAdr As String = System.Environment.GetEnvironmentVariable("nitoms_shared_address")
        'MessageBox.Show(sharedAdr)

        Dim myAdr As String = Application.GetNamespace("MAPI").Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        Dim mail As Outlook.MailItem = TryCast(Item, Outlook.MailItem)

        'メールオブジェクトの場合は以下継続するのです
        If Not IsNothing(mail) Then
            'MessageBox.Show("Application ItemSend As Mail")


            If (MessageBox.Show("共有アドレスとして送信しますか？" & vbNewLine & sharedAdr, "メール送信確認①", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                'MessageBox.Show("selected YES")
                'objMail As Outlook.MailItem
                'Dim objRecip As Outlook.Recipient
                'Dim SharedAddr As String
                '既存のItemは送信元が確定しているので代理送信ができない。従ってメールオブジェクトをコピーしてそれを使う
                'objMail = TryCast(Item, Outlook.MailItem).Copy

                'コピー後 Item は削除・送信キャンセルする
                'TryCast(Item, Outlook.MailItem).Delete()
                'Cancel = True

                'objMail.SentOnBehalfOfName = "sb-nitoms-it-servicedesk@nitto.com"
                'objMail.Save()
                'System.Windows.Forms.MessageBox.Show(objMail.SentOnBehalfOfName)
                'objRecip = objMail.Recipients.Add("sb-nitoms-it-servicedesk@nitto.com")
                'objRecip.Resolve()
                'If objRecip.Resolved Then
                ' objRecip.Type = Outlook.OlMailRecipientType.olCC    ' 2
                'End If
                'objRecip = Nothing
                '

                '既存のItemは送信元が確定しているので代理送信ができない。従ってメールオブジェクトをコピーしてそれを使う
                'Dim objMail As Outlook.MailItem = TryCast(Item, Outlook.MailItem).Copy
                Dim objMail As Outlook.MailItem = mail.Copy

                'コピー後 Item は削除・送信キャンセルする
                'TryCast(Item, Outlook.MailItem).Delete()
                mail.Delete()
                Cancel = True

                objMail.SentOnBehalfOfName = sharedAdr
                objMail.CC = objMail.CC + ";" + sharedAdr
                objMail.Send()

                'Dim mailItem As Outlook.MailItem = TryCast(Item, Outlook.MailItem)
                'mailItem.BCC = SharedAddr
                'mailItem.SentOnBehalfOfName = SharedAddr
                'mailItem.Save()
                '       mailItem.Send()
                '       Cancel = True
            ElseIf (MessageBox.Show("個人アドレスとして送信しますか？" & vbNewLine & myAdr, "メール送信確認②", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                Dim objMail As Outlook.MailItem = mail.Copy
                mail.Delete()
                Cancel = True
                objMail.SentOnBehalfOfName = myAdr
                objMail.Send()
            Else
                MessageBox.Show("メールは未送信です")
                Cancel = True
            End If
        End If
    End Sub

End Class

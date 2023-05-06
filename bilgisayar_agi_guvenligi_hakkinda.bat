Option Explicit

Sub Main()
    Dim title As String: title = "Bilgisayar Ağının Güvenliği Hakkında"
    Dim message1 As String: message1 = "Merhaba çalışanlar,"
    Dim message2 As String: message2 = "Bu yazılım, bilgisayar ağının güvenliği hakkında bilgi almanızı sağlayacak."
    Dim message3 As String: message3 = "Okumak istediğiniz metne tıklayın."
    Dim message4 As String: message4 = "Sizce bilgisayar ağının güvenliği neden önemlidir? Lütfen düşüncelerinizi paylaşın."
    Dim message5 As String: message5 = "Teşekkür ederiz! Bilgisayarınızın güvenliği için aldığınız önlemlerden dolayı takdir ediyoruz."
    Dim message6 As String: message6 = "Bilgisayarınızın kapatılması gerekiyor. Devam etmek istiyor musunuz?"

    MsgBox message1 & vbCrLf & message2 & vbCrLf & message3, vbInformation, title
    MsgBox message4, vbQuestion, title

    Dim response As Integer
    response = MsgBox("Bilgisayarınızın güvenliği hakkında bilgi almak için 'Evet', kapatmak için 'Hayır' seçeneğini tıklayın.", vbQuestion + vbYesNo, title)

    If response = vbYes Then
        MsgBox message5, vbInformation, title
    Else
        MsgBox message6, vbExclamation, title, 30
        Shell "shutdown 600 -s -t 0", vbHide
    End If
End Sub

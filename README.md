# mail-using-Excel-VBA
Automate  mail using Excel VBA 

**below Script will send my local image file as report mai**

Sub Automail()

Dim OutApp As Object
Dim OutMail As Object
Dim table As Range
Dim pic As Picture
Dim ws As Worksheet
Dim wordDoc



Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = "<BODY style=font-size:14pt;font-family:Consolas>Dear Sir<p>Please look into below General Monitoring<p> </BODY> "
emailMessage = "<BODY style=font-size:11pt;font-family:Calibri>Dear Team " & titleName & " " & FullName & "," & _
                    "<p>Please check below <strong>General Monitoring</strong> Report ." & _
                    "<p>"
'create email message
On Error Resume Next
    With OutMail
        .To = "abc@gmail.com"
        .CC = ""
        .BCC = ""
        .Subject = "General Monitoring SCP " & Format(Date, "dd-mm-yy")
        .Height = 50
        .Width = 50
        .Attachments.Add "C:\xampp\htdocs\st22.png", olByValue, 0
        sImgName1 = "st22.png"
        .Attachments.Add "C:\xampp\htdocs\st22d.png", olByValue, 0
        sImgName2 = "st22d.png"
        .Attachments.Add "C:\xampp\htdocs\db01.png", olByValue, 0
        sImgName3 = "db01.png"
        .Attachments.Add "C:\xampp\htdocs\al08.png", olByValue, 0
        sImgName4 = "al08.png"
        .Attachments.Add "C:\xampp\htdocs\sm12.png", olByValue, 0
        sImgName5 = "sm12.png"
        .Attachments.Add "C:\xampp\htdocs\sm13.png", olByValue, 0
        sImgName6 = "sm13.png"
        .Attachments.Add "C:\xampp\htdocs\sost.png", olByValue, 0
        sImgName7 = "sost.png"
        .Attachments.Add "C:\xampp\htdocs\sm51.png", olByValue, 0
        sImgName8 = "sm51.png"
        .Attachments.Add "C:\xampp\htdocs\sm66.png", olByValue, 0
        sImgName9 = "sm66.png"
        .Attachments.Add "C:\xampp\htdocs\sm66spo.png", olByValue, 0
        sImgName10 = "sm66spo.png"
        .Attachments.Add "C:\xampp\htdocs\sp01.png", olByValue, 0
        sImgName11 = "sp01.png"
        .Attachments.Add "C:\xampp\htdocs\sm37cnl.png", olByValue, 0
        sImgName12 = "sm37cnl.png"
        .Attachments.Add "C:\xampp\htdocs\sm37act.png", olByValue, 0
        sImgName13 = "sm37act.png"
        .HTMLBody = "<html>" & _
                "<p <h1 style=""text-align:center"">General Monitoring Report</h1></p>" & _
                "<p <h3 style=""text-align:center""> ST22 Dumps Count </h3></p>" & "<p style=""text-align:center""><img src=""cid:" & sImgName1 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> ST22 Dumps Status </h3></p>" & "<p style=""text-align:center""><img src=""cid:" & sImgName2 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Blocked Transactions </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName3 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> User Sessions  </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName4 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Lock Entries </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName5 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Update Request </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName6 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> SOST </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName7 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> SM51 </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName8 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Active Work Process </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName9 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Active Work Process SPO </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName10 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Spools </h3></p>" & "<p style=""text-align:center""><img src=""cid:" & sImgName11 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Cancelled Background Jobs </h3></p>" & "<p style=""text-align:center""><img src=""cid:" & sImgName12 & """ height=520 width=750></p>" & _
                "<p <h3 style=""text-align:center""> Active Background Jobs </h3></p>" & _
                "<p style=""text-align:center""><img src=""cid:" & sImgName13 & """ height=520 width=750></p>" & _
                "<p <h2 style=""text-align:center"">Thank You </h2></p>"
        .Display
        .send
    End With
    On Error GoTo 0
    
Set OutApp = Nothing
Set OutMail = Nothing

MsgBox Completed
End Sub

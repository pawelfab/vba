
  Private Sub STATUS_RMA_AfterUpdate()
Dim statusRMA As String
 statusRMA = DLookup("[status]", "status_rma", "[id_status] =" & Me.status_rma)

If (Not IsNull(Me.EMAIL)) Then
Set szablon = Nothing
sToAddress = Me.EMAIL
sBCCAdress = "***@***.pl"
sSubject = "Status naprawy  " + Me.MODEL + " , " + Me.NUMER_SERYJNY + ": " + statusRMA
'dwa razy kod html stopki bo sa dwa szablony i w kazdym inaczej outlook zalacza nazwy zdjec np w jednym logo ma image001 w drugim zminil na image003
sStopka = "<br><br><p style=""font-weight: lighter; font:Verdana, 10px""> Pozdrawiamy serdecznie, <br><br>Autoryzowany Serwis <br>tel. 800 800 800<br>faks: (22) ***<br>***@***.pl<br><br><img src=""image001.jpg""><br><br>POLAND S.A.<br>ul. **** 10<br> Warszawa<br>NIP </p>"
sStopkaAparat = "<br><br><p style=""font-weight: lighter; font:Verdana, 10px""> Pozdrawiamy serdecznie, <br><br>Autoryzowany Serwis <br>tel. 800 800 800<br>faks: (22) 8***<br>drukarki@***.pl<br><br><img src=""image003.jpg""><br><br>POLAND S.A.<br>ul. *** 10<br>Warszawa<br>NIP </p>"
sPromoAparat = "<br><p style=""font-weight: lighter; font:Verdana, 10px;""> <table><tr><th rowspan=""3""><img src=""image004.jpg"" height=""48"" width=""62""></th><th style=""text-align:left;"">Wygraj cyfrowy aparat fotograficzny  !</th></tr><tr><td>Wypełnij ankietę pod adresem <a href=""www. -europe.com/surveya"">www. -europe.com/surveya</td></tr><tr><td>i wygraj aparat </td></tr></table></p>"
'załaczniki jpg są z pliku szablonu oft katalog na dysku c ale outlook sam zmienia nazwy zdjec w szablonie na img003.jpg itd. Szablon nalezy tworzyc pusty z tylko zalaczonymi zdjeciami

If IsNull(Me.infomail) Then
        sInfomail = " "
    Else
        sInfomail = Me.infomail & " ."
End If

If IsNull(Me.ADRES) Then
        sRMAklienta = " "
    Else
        sRMAklienta = "<br><br>Nr RMA Klienta: " + Me.ADRES + "<br>"
End If

If IsNull(Me.nowy_Nr_seryjny) Then
        sNowyNrSer = " "
    Else
        sNowyNrSer = "<br><br>Nowy numer seryjny urządzenia: " + Me.nowy_Nr_seryjny + "<br>"
End If


If (Not IsNull(Me.NR_LISTU_WYCH)) Then
    szablon = 1
    sBody = "Dzień dobry!<br><br>Informujemy, że naprawa Twojego urządzenia <b>" + Me.MODEL + "</b> o numerze seryjnym: <b>" + Me.NUMER_SERYJNY + "</b><br>Otrzymała status: <b>" + statusRMA + "</b><br>Numer listu przewozowego kuriera DPD: <b>" + NR_LISTU_WYCH + "</b><BR><br>Po otrzymaniu  przesyłki prosimy rozpakować ją w obecności kuriera i sprawdzić stan towaru. W przypadku uszkodzenia należy spisać protokół reklamacyjny (druk posiada kurier), podpisać i przekazać kurierowi do podpisania. Rozpocznie to procedurę reklamacji w firmie kurierskiej.<br><br>Status przesyłki znajdziesz tutaj: <a href=""https://tracktrace.dpd.com.pl/parcelDetails?p1=" + NR_LISTU_WYCH + """>Monitoring przesyłek</a>" + sNowyNrSer + sRMAklienta + sStopkaAparat + sPromoAparat
 Else
    szablon = 0
   sBody = "Dzień dobry!<br><br>Informujemy, że naprawa Twojego urządzenia <b>" + Me.MODEL + "</b> o numerze seryjnym: <b>" + Me.NUMER_SERYJNY + "</b><br>Otrzymała status: <b>" + statusRMA + "<br><br>" + sInfomail + sRMAklienta + sStopka
 
End If
  
  DoCmd.SetWarnings False
  
  
  
  
   If MsgBox("Zmienić status naprawy? Zostanie wysłany email do klienta!", vbYesNo, "Uwaga") = vbYes Then



        'wysyłanie maila
        Dim outl As Outlook.Application
        Set outl = New Outlook.Application
        Dim mi As Outlook.MailItem



        If (szablon = 1) Then
            Set mi = outl.CreateItemFromTemplate("C:\outlooktemplate\drukarki_aparat.oft")
        Else
            Set mi = outl.CreateItemFromTemplate("C:\outlooktemplate\drukarki.oft")
        End If
        mi.HTMLBody = sBody
        mi.Subject = sSubject
        mi.To = sToAddress
        mi.BCC = sBCCAdress
        'mi.SentOnBehalfOfName = "***@***.pl"
        mi.Send
        'mi.Display
        Set mi = Nothing
        Set outl = Nothing
        '//wysylanie
        Set szablon = Nothing
        

   End If
   DoCmd.SetWarnings False


   End If
End Sub
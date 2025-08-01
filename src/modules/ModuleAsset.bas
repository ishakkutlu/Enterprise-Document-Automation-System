Attribute VB_Name = "ModuleAsset"
Option Explicit
Dim GelenTemaGlobal As String
Dim ctl As MSForms.Control
Dim SayEsasVarlikGlobal As Long, SiraNoEsasVarlikGlobal As Long
Dim WsVarliklarGlobal As Object

Sub hideheadings()
    Dim wsSheet As Worksheet
    
    ThisWorkbook.Unprotect "123"
    
    Application.ScreenUpdating = False
    For Each wsSheet In ThisWorkbook.Worksheets
        wsSheet.Activate
        With ActiveWindow
            .DisplayHeadings = False 'True 'False
        End With
    Next wsSheet
    
    ThisWorkbook.Protect "123"
    
    Application.ScreenUpdating = True
End Sub


Sub GenelVarlikEsas()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim ItemBul As Range

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "General Primary Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

On Error Resume Next 'Operation içinde General Primary Asset Report.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "General Primary Asset Report.xlsx"
'General Primary Asset Report açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("General Primary Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "General Primary Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("General Primary Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
'            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR2_2 (GİRİŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If

                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                Else
                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
    '                WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                Else
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
    
                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                            
                            'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                End If
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklar.Range("D" & i) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR3_2 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
            End If
            
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If
                
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR3_1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
            End If
            
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If

            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            
            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi(DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira))
            End If
        End If
        'RAPOR (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            
            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
            End If
        End If
        'RAPOR2_2 (ÇIKIŞLAR, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                    End If
                    WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    End If
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                Else
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
    
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
                    End If

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        End If
                        WsVarliklar.Range("H" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        
                        'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                        If ctl.BackColor = &H80000003 Then
                            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & j))
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                    End If
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    End If
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                Else
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
    
                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                            
                            'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            End If
                            WsVarliklar.Range("H" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
    
                            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                            If ctl.BackColor = &H80000003 Then
                                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & j))
                            End If
                            
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                End If
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklar.Range("D" & i) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR3_2 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
            End If
            
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If

            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1

            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira))
            End If
        End If
        'RAPOR3_1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
            End If
            
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            
            WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If

            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1

            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira))
            End If
        End If
    End If
Next ctl



'GoTo DuzenlemeyiAtla
'_____________________________________________________________BAŞLANGIÇ (Direkt kaldırsan yine de sorunsuz çalışır.)

Dim KontB As Integer, KontR As Integer, SayEsasVarlikx As Long
Dim SayEsasVarlikTakip As Long, a As Integer, b As Integer
Dim NumRaporNo As Variant

SayEsasVarlikTakip = 0
KontB = 0
KontR = 0
SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'General Primary Asset Report düzeltme faktörü (Technique A ve diğerlerinin sıra numarasını birleştir.)
If SayEsasVarlikx - 1 >= 7 Then
'    'RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox "2: " & StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir1
                    End If
                Next j
DonguBitir1:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontB = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontB = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon1
                    End If
                Next x
xDonguSon1:
                WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklar.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon2
                            End If
                        Next x
xDonguSon2:
                        WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                        WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip
    
    'RAPOR NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo

                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox StrRaporNo
        
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir2
                    End If
                Next j
DonguBitir2:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontR = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontR = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon3
                    End If
                Next x
xDonguSon3:
                WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklar.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon4
                            End If
                        Next x
xDonguSon4:

                        'WsVarliklar.Range("L" & j).Value = NumRaporNo

                        WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                        WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklar.Range("L" & SayEsasVarlikTakip & ":L" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = NumRaporNo
                        'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip

    'Sıra no.lar
    SiraNoEsasVarlik = 0
    'MsgBox SayEsasVarlik
    'WsVarliklar.Range("C" & 7 & ":C" & SayEsasVarlik).ClearContents
    For i = 7 To SayEsasVarlik
        If WsVarliklar.Range("C" & i) <> "" And WsVarliklar.Range("I" & i) = "" Then
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            WsVarliklar.Range("C" & i) = SiraNoEsasVarlik
        End If
        If WsVarliklar.Range("I" & i) <> "" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=WsVarliklar.Range("I" & i), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=WsVarliklar.Range("I" & i), SearchDirection:=xlPrevious, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'MsgBox IlkSira & " : " & SonSira
            If IlkSira = SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklar.Range("C" & i) = SiraNoEsasVarlik
                i = SonSira
            ElseIf IlkSira <> SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklar.Range("C" & IlkSira) = SiraNoEsasVarlik
                WsVarliklar.Range("C" & IlkSira & ":C" & SonSira).Merge
                i = SonSira
            End If

        End If
    Next i
End If

SatirSilTekrar:
SayEsasVarlik = SayEsasVarlik - 1
If SayEsasVarlik < 7 Then
    GoTo SatirSilTekrarAtla
End If
For i = 7 To SayEsasVarlik
    If WsVarliklar.Range("K" & i) = "" Then
        WsVarliklar.Rows(i).EntireRow.Delete
        GoTo SatirSilTekrar
    End If
Next i
SatirSilTekrarAtla:

If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
Else
    SayEsasVarlik = SayEsasVarlik + 1
End If

'_____________________________________________________________BİTİŞ



'____________________İkinci düzenleme bölümü, BAŞLANGIÇ (Bu bölüm sadece XXXMud'den gelen raporları/öğelerin çıkışı içindir.)

SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'Raporların rapro no.'ları ile onun alt kırılımını sayısal bir değere dönüştür. Bu özellik raporları kendi içinde sıralamada kullanılacak.
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR ve RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'R
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklar.Range("D" & i)
            End If
        End If
        'B
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo
        '_________________________
        For a = 1 To 50
            StrRaporNo = Replace(StrRaporNo, " ", "") 'Rapor no içinde varsa boşlukları kaldır
        Next a
        NumRaporNo = Replace(StrRaporNo, "-", "0000") '1-1'in 100001 gerçek rapor no ile çakışma ihtimali 0'a yakın.

        '_________________________

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir21
            End If
        Next j
DonguBitir21:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik1
        End If

        For j = IlkSiraVarlik To SonSiraVarlik
            If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 And WsVarliklar.Range("L" & j).Value = "" Then
                WsVarliklar.Range("L" & j).Value = NumRaporNo
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklar.Range("D" & a + 1) = "" And WsVarliklar.Range("K" & a + 1) <> "" And WsVarliklar.Range("L" & a + 1).Value = "" Then
                            WsVarliklar.Range("L" & a + 1).Value = NumRaporNo
                        Else
                            GoTo aDonguSonu2
                        End If
                    End If
                Next a
            End If
        Next j
aDonguSonu2:
DonguSonuIlkSiraVarlik1:
    Next i
End If


'XXXMud'den gelen ve varlıkdaki raporlar. SIRALAMA ve BİRLEŞTİRME
For i = 7 To SayEsasVarlik - 1
    If WsVarliklar.Range("M" & i) = "" Then 'And WsVarliklar.Range("H" & i) <> "" Then
        If Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklar.Range("D" & i)
            End If
        End If
        If Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir22
            End If
        Next j
DonguBitir22:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo
        '_______________


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik
        End If

        WsVarliklar.Range("N" & IlkSiraVarlik) = IlkSiraVarlik
        WsVarliklar.Range("O" & SonSiraVarlik) = SonSiraVarlik
        WsVarliklar.Range("M" & IlkSiraVarlik & ":M" & SonSiraVarlik).Value = "*" 'Aynı satırlarda tekrar işlem yapmasın diye
        'Sorting Z to A by NumRaporNo
        WsVarliklar.Range("D" & IlkSiraVarlik & ":L" & SonSiraVarlik).UnMerge
        WsVarliklar.Range("D" & IlkSiraVarlik & ":L" & SonSiraVarlik).Sort key1:=Range("L" & IlkSiraVarlik & ":L" & SonSiraVarlik), order1:=xlAscending, Header:=xlNo

        'Rapor no.'ları ve Package A/Package B/Package C satırlarını birleştir.
        For j = IlkSiraVarlik To SonSiraVarlik
            If WsVarliklar.Range("L" & j).Value <> "" Then
                If WsVarliklar.Range("L" & j).Value = WsVarliklar.Range("L" & j + 1).Value Then
                    WsVarliklar.Range("D" & j & ":D" & j + 1).Merge
                End If
            End If
            If WsVarliklar.Range("E" & j).Value = "Package A" Or WsVarliklar.Range("E" & j).Value = "Package B" Or WsVarliklar.Range("E" & j).Value = "Package C" Then
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklar.Range("E" & a + 1).Value = "" Then
                            WsVarliklar.Range("E" & a & ":E" & a + 1).Merge
                            WsVarliklar.Range("F" & a & ":F" & a + 1).Merge
                            WsVarliklar.Range("G" & a & ":G" & a + 1).Merge
                            WsVarliklar.Range("H" & a & ":H" & a + 1).Merge
                        Else
                            GoTo aDonguSonu
                        End If
                    End If
                Next a
            End If
aDonguSonu:
        Next j
    End If
DonguSonuIlkSiraVarlik:
Next i

'Artıkları temizle
If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
End If
WsVarliklar.Range("I" & 7 & ":O" & SayEsasVarlik).ClearContents

'____________________İkinci düzenleme bölümü, BİTİŞ


DuzenlemeyiAtla:

'GoTo Son

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira))
        End If
        'RAPOR (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
        End If
        'RAPOR2_2 (DEVİRLER, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                Else
                    Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                For i = IlkSira To SonSira
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & i).Value, 8) <> "Technique A" Then
                        Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & i))
                    End If
                Next i
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                Else
                    For i = IlkSira To SonSira
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & i).Value, 11) = "Technique A" Then
                            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & i))
                        End If
                    Next i
                End If
            End If
        End If
        'RAPOR3_2 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira))
        End If
        'RAPOR3_1 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira))
        End If
    End If
Next ctl

ToplamName = "Total (Qty)"
DevirName = "Carried-Over Balance from the Previous Day (Qty)"
GunSonuName = "End-of-Day Balance (Qty)"
WsVarliklar.Range("C" & SayEsasVarlik) = ToplamName
WsVarliklar.Range("C" & SayEsasVarlik + 2) = DevirName
WsVarliklar.Range("C" & SayEsasVarlik + 3) = GunSonuName
WsVarliklar.Range("C" & SayEsasVarlik & ":H" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":H" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":H" & SayEsasVarlik + 3).Font.Bold = True

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":H" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).Merge

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).HorizontalAlignment = xlLeft

If SayEsasVarlik - 1 >= 7 Then
    WsVarliklar.Range("G" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 7), WsVarliklar.Cells(SayEsasVarlik - 1, 7))) 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 8), WsVarliklar.Cells(SayEsasVarlik - 1, 8))) 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = WsVarliklar.Range("G" & SayEsasVarlik) + Devir - WsVarliklar.Range("H" & SayEsasVarlik) 'Gün sonu toplamı 'Gün sonu toplamı
Else
    WsVarliklar.Range("G" & SayEsasVarlik) = 0 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = 0 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = 0 + Devir - 0 'Gün sonu toplamı
End If

'Kontrol edilmiştir metni.
WsVarliklar.Range("C" & SayEsasVarlik + 5) = "The above records have been reviewed by us."
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).HorizontalAlignment = xlCenter 'xlLeft
'Anahtar sahipleri metni.
WsVarliklar.Range("C" & SayEsasVarlik + 7) = "Relevant Officers"
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).HorizontalAlignment = xlCenter 'xlLeft



'imzalar
WsVarliklar.Range("C" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza1.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza1Atla
End If
WsVarliklar.Range("C" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza1.Value <> "" Then
    WsVarliklar.Range("C" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza1Atla:

WsVarliklar.Range("E" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza2.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza2Atla
End If
WsVarliklar.Range("E" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza2.Value <> "" Then
    WsVarliklar.Range("E" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza2Atla:

WsVarliklar.Range("G" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza3.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza3Atla
End If
WsVarliklar.Range("G" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza3.Value <> "" Then
    WsVarliklar.Range("G" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza3Atla:


'İmza satırlarını ayarla
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter


'Kenarlıklar.
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":H" & SayEsasVarlik + 3)
'Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
'Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .Color = RGB(217, 217, 217)
End With
With Kenarlar.Borders(xlEdgeBottom)
    .Color = RGB(217, 217, 217)
End With
With Kenarlar.Borders(xlInsideVertical)
    .Color = RGB(217, 217, 217)
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .Color = RGB(217, 217, 217)
End With


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub Rapor1Varlik()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 1 – Supporting Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Varlıklar.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 1 – Supporting Asset Report.xlsx"
'Varlıklar açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 1 – Supporting Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 1 – Supporting Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 1 – Supporting Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
'Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            Else
                WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            End If
            WsVarliklar.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

If WsVarliklar.Range("C7") = "" Then
    GoTo Son
End If

'Öğelere göre toplamlar
WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 3).Value = "*"
x = SayEsasVarlik + 3
For i = 7 To SayEsasVarlik - 1
    'WsVarliklar.Range("C" & i & ":N" & i).EntireRow.AutoFit
    'Mevcutlar
    If WsVarliklar.Range("J" & i) <> "" And WsVarliklar.Range("K" & i) <> "" Then
        Say2 = WsVarliklar.Range("J100000").End(xlUp).Row '+ 1
        'MsgBox "SayEsasVarlik: " & SayEsasVarlik & " Say2: " & Say2
        For j = SayEsasVarlik + 2 To Say2
            'Öğe aynı
            If WsVarliklar.Range("J" & j) = WsVarliklar.Range("J" & i) Then 'Öğe aynı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    'Öğe ve öğe değeri aynı
                    If WsVarliklar.Range("K" & j) = WsVarliklar.Range("K" & i) Then
                        WsVarliklar.Range("L" & j) = WsVarliklar.Range("L" & j) + WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j) = WsVarliklar.Range("M" & j) + WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j) = WsVarliklar.Range("N" & j) + WsVarliklar.Range("N" & i)
                        GoTo jDonguSon1
                    Else 'Öğe aynı, öğe değeri farklı
                        If WsVarliklar.Range("J" & j + 1) = "" Then
                            WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                            WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                            WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                            WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                            WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                            x = x + 1
                        End If
                    End If
                End If
            Else 'Öğe farklı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    If WsVarliklar.Range("J" & j + 1) = "" Then
                        WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                        WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                        WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                        x = x + 1
                    End If
                End If
            End If
        Next j
jDonguSon1:
    End If
Next i

'Boş satırlara 0 yaz.
For j = SayEsasVarlik + 2 To x
    If WsVarliklar.Range("L" & j) = "" Then
        WsVarliklar.Range("L" & j).Value = 0
    End If
    If WsVarliklar.Range("M" & j) = "" Then
        WsVarliklar.Range("M" & j).Value = 0
    End If
    If WsVarliklar.Range("N" & j) = "" Then
        WsVarliklar.Range("N" & j).Value = 0
    End If
Next j

WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 2).ClearContents
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":O" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik & ":K" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik) = "Total"
WsVarliklar.Range("C" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik).HorizontalAlignment = xlRight
WsVarliklar.Range("L" & SayEsasVarlik & ":O" & SayEsasVarlik).Font.Bold = True
'Toplamlar (Sol üstteki toplam)
For i = 7 To SayEsasVarlik - 1
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & SayEsasVarlik).Value = WsVarliklar.Range("L" & SayEsasVarlik).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & SayEsasVarlik).Value = WsVarliklar.Range("M" & SayEsasVarlik).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & SayEsasVarlik).Value = WsVarliklar.Range("N" & SayEsasVarlik).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & SayEsasVarlik).Value = WsVarliklar.Range("O" & SayEsasVarlik).Value + WsVarliklar.Range("O" & i).Value
Next i

'ÖZET TABLO BAŞLIĞI
WsVarliklar.Range("J" & SayEsasVarlik + 2).Value = "Item-Based Totals (Qty)"
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft

WsVarliklar.Range("J" & SayEsasVarlik + 3).Value = "Item Type"
WsVarliklar.Range("K" & SayEsasVarlik + 3).Value = "Item Value"
WsVarliklar.Range("L" & SayEsasVarlik + 3).Value = "Current Qty"
WsVarliklar.Range("M" & SayEsasVarlik + 3).Value = "Inbound Qty"
WsVarliklar.Range("N" & SayEsasVarlik + 3).Value = "Outbound Qty"
WsVarliklar.Range("O" & SayEsasVarlik + 3).Value = "Remaining Qty"
WsVarliklar.Range("J" & SayEsasVarlik + 3 & ":O" & SayEsasVarlik + 3).Font.Bold = True
y = SayEsasVarlik + 3

'Toplamlar (öğelere göre, sağ alttaki)
For i = SayEsasVarlik + 4 To x
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & x + 1).Value = WsVarliklar.Range("L" & x + 1).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & x + 1).Value = WsVarliklar.Range("M" & x + 1).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & x + 1).Value = WsVarliklar.Range("N" & x + 1).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & x + 1).Value = WsVarliklar.Range("O" & x + 1).Value + WsVarliklar.Range("O" & i).Value
Next i
'Sort Z to A
WsVarliklar.Range("J" & SayEsasVarlik + 4 & ":O" & x).Sort key1:=Range("K" & SayEsasVarlik + 4 & ":K" & x), order1:=xlDescending, Header:=xlNo

WsVarliklar.Range("L" & x + 1).Font.Bold = True
WsVarliklar.Range("M" & x + 1).Font.Bold = True
WsVarliklar.Range("N" & x + 1).Font.Bold = True
WsVarliklar.Range("O" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1) = "Total"
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Merge
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).HorizontalAlignment = xlLeft

'Kenarlıklar.
SayEsasVarlik = x + 1
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":O" & SayEsasVarlik)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
WsVarliklar.Range("C" & y - 2 & ":I" & x + 2).Borders.LineStyle = xlNone
WsVarliklar.Range("J" & y - 1 & ":J" & x + 1).Borders((xlLeft)).LineStyle = xlNone


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub RaporVarlik()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 2.1 – Supporting Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Varlıklar.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 2.1 – Supporting Asset Report.xlsx"
'Varlıklar açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 2.1 – Supporting Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 2.1 – Supporting Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 2.1 – Supporting Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
'Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal

            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            Else
                WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            End If
            WsVarliklar.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Row height
            If IlkSira = SonSira Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklar.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

If WsVarliklar.Range("C7") = "" Then
    GoTo Son
End If


'Öğelere göre toplamlar
WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 3).Value = "*"
x = SayEsasVarlik + 3
For i = 7 To SayEsasVarlik - 1
    'WsVarliklar.Range("C" & i & ":N" & i).EntireRow.AutoFit
    'Mevcutlar
    If WsVarliklar.Range("J" & i) <> "" And WsVarliklar.Range("K" & i) <> "" Then
        Say2 = WsVarliklar.Range("J100000").End(xlUp).Row '+ 1
        'MsgBox "SayEsasVarlik: " & SayEsasVarlik & " Say2: " & Say2
        For j = SayEsasVarlik + 2 To Say2
            'Öğe aynı
            If WsVarliklar.Range("J" & j) = WsVarliklar.Range("J" & i) Then 'Öğe aynı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    'Öğe ve öğe değeri aynı
                    If WsVarliklar.Range("K" & j) = WsVarliklar.Range("K" & i) Then
                        WsVarliklar.Range("L" & j) = WsVarliklar.Range("L" & j) + WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j) = WsVarliklar.Range("M" & j) + WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j) = WsVarliklar.Range("N" & j) + WsVarliklar.Range("N" & i)
                        GoTo jDonguSon1
                    Else 'Öğe aynı, öğe değeri farklı
                        If WsVarliklar.Range("J" & j + 1) = "" Then
                            WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                            WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                            WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                            WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                            WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                            x = x + 1
                        End If
                    End If
                End If
            Else 'Öğe farklı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    If WsVarliklar.Range("J" & j + 1) = "" Then
                        WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                        WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                        WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                        x = x + 1
                    End If
                End If
            End If
        Next j
jDonguSon1:
    End If
Next i

'Boş satırlara 0 yaz.
For j = SayEsasVarlik + 2 To x
    If WsVarliklar.Range("L" & j) = "" Then
        WsVarliklar.Range("L" & j).Value = 0
    End If
    If WsVarliklar.Range("M" & j) = "" Then
        WsVarliklar.Range("M" & j).Value = 0
    End If
    If WsVarliklar.Range("N" & j) = "" Then
        WsVarliklar.Range("N" & j).Value = 0
    End If
Next j

WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 2).ClearContents
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":O" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik & ":K" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik) = "Total"
WsVarliklar.Range("C" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik).HorizontalAlignment = xlRight
WsVarliklar.Range("L" & SayEsasVarlik & ":O" & SayEsasVarlik).Font.Bold = True
'Toplamlar (Sol üstteki toplam)
For i = 7 To SayEsasVarlik - 1
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & SayEsasVarlik).Value = WsVarliklar.Range("L" & SayEsasVarlik).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & SayEsasVarlik).Value = WsVarliklar.Range("M" & SayEsasVarlik).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & SayEsasVarlik).Value = WsVarliklar.Range("N" & SayEsasVarlik).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & SayEsasVarlik).Value = WsVarliklar.Range("O" & SayEsasVarlik).Value + WsVarliklar.Range("O" & i).Value
Next i

'ÖZET TABLO BAŞLIĞI
WsVarliklar.Range("J" & SayEsasVarlik + 2).Value = "Item-Based Totals (Qty)"
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft

WsVarliklar.Range("J" & SayEsasVarlik + 3).Value = "Item Type"
WsVarliklar.Range("K" & SayEsasVarlik + 3).Value = "Item Value"
WsVarliklar.Range("L" & SayEsasVarlik + 3).Value = "Current Qty"
WsVarliklar.Range("M" & SayEsasVarlik + 3).Value = "Inbound Qty"
WsVarliklar.Range("N" & SayEsasVarlik + 3).Value = "Outbound Qty"
WsVarliklar.Range("O" & SayEsasVarlik + 3).Value = "Remaining Qty"
WsVarliklar.Range("J" & SayEsasVarlik + 3 & ":O" & SayEsasVarlik + 3).Font.Bold = True
y = SayEsasVarlik + 3

'Toplamlar (öğelere göre, sağ alttaki)
For i = SayEsasVarlik + 4 To x
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & x + 1).Value = WsVarliklar.Range("L" & x + 1).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & x + 1).Value = WsVarliklar.Range("M" & x + 1).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & x + 1).Value = WsVarliklar.Range("N" & x + 1).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & x + 1).Value = WsVarliklar.Range("O" & x + 1).Value + WsVarliklar.Range("O" & i).Value
Next i
'Sort Z to A
WsVarliklar.Range("J" & SayEsasVarlik + 4 & ":O" & x).Sort key1:=Range("K" & SayEsasVarlik + 4 & ":K" & x), order1:=xlDescending, Header:=xlNo

WsVarliklar.Range("L" & x + 1).Font.Bold = True
WsVarliklar.Range("M" & x + 1).Font.Bold = True
WsVarliklar.Range("N" & x + 1).Font.Bold = True
WsVarliklar.Range("O" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1) = "Total"
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Merge
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).HorizontalAlignment = xlLeft

'Kenarlıklar.
SayEsasVarlik = x + 1
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":O" & SayEsasVarlik)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
WsVarliklar.Range("C" & y - 2 & ":I" & x + 2).Borders.LineStyle = xlNone
WsVarliklar.Range("J" & y - 1 & ":J" & x + 1).Borders((xlLeft)).LineStyle = xlNone


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub
Sub Rapor2_2VarlikDevam()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object ', WsVarliklarGlobal As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String, XXXMudGiden As String, XXXMudGelen As String

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long


SayEsasVarlik = SayEsasVarlikGlobal
SiraNoEsasVarlik = SiraNoEsasVarlikGlobal

'GoTo DuzenlemeyiAtla

'_____________________________________________________________BAŞLANGIÇ (Direkt kaldırsan yine de sorunsuz çalışır.)

Dim KontB As Integer, KontR As Integer, SayEsasVarlikx As Long
Dim SayEsasVarlikTakip As Long, a As Integer, b As Integer
Dim NumRaporNo As Variant

SayEsasVarlikTakip = 0
KontB = 0
KontR = 0
SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik



'MsgBox SayEsasVarlik

'varlıklar düzeltme faktörü (Technique A ve diğerlerinin sıra numarasını birleştir.)
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox "2: " & StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir1
                    End If
                Next j
DonguBitir1:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontB = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontB = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon1
                    End If
                Next x
xDonguSon1:

'                'Row height
'                If SonAltSatirNo = IlkAltSatirNo Then
'                    WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                End If

                WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklarGlobal.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon2
                            End If
                        Next x
xDonguSon2:

'                        'Row height
'                        If SonAltSatirNo = IlkAltSatirNo Then
'                            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                        End If

                        WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                        WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip

    'RAPOR NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo

                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir2
                    End If
                Next j
DonguBitir2:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontR = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontR = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon3
                    End If
                Next x
xDonguSon3:

'                'Row height
'                If SonAltSatirNo = IlkAltSatirNo Then
'                    WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                End If

                WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklarGlobal.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon4
                            End If
                        Next x
xDonguSon4:

'                        'Row height
'                        If SonAltSatirNo = IlkAltSatirNo Then
'                            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                        End If

                        WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                        WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

'GoTo Son

     SayEsasVarlik = SayEsasVarlikTakip

    'Sıra no.lar
    SiraNoEsasVarlik = 0
    'WsVarliklarGlobal.Range("C" & 7 & ":C" & SayEsasVarlik).ClearContents
    For i = 7 To SayEsasVarlik
        If WsVarliklarGlobal.Range("C" & i) <> "" And WsVarliklarGlobal.Range("P" & i) = "" Then
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            WsVarliklarGlobal.Range("C" & i) = SiraNoEsasVarlik
        End If
        If WsVarliklarGlobal.Range("P" & i) <> "" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=WsVarliklarGlobal.Range("P" & i), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=WsVarliklarGlobal.Range("P" & i), SearchDirection:=xlPrevious, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
'            MsgBox IlkSira & " : " & SonSira
            If IlkSira = SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklarGlobal.Range("C" & i) = SiraNoEsasVarlik
                'Row height
                WsVarliklarGlobal.Range("C" & IlkSira & ":O" & IlkSira).EntireRow.AutoFit
                i = SonSira
            ElseIf IlkSira <> SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklarGlobal.Range("C" & IlkSira) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("C" & IlkSira & ":C" & SonSira).Merge
                WsVarliklarGlobal.Range("E" & IlkSira & ":E" & SonSira).Merge
                WsVarliklarGlobal.Range("F" & IlkSira & ":F" & SonSira).Merge
                WsVarliklarGlobal.Range("G" & IlkSira & ":G" & SonSira).Merge
                WsVarliklarGlobal.Range("H" & IlkSira & ":H" & SonSira).Merge
                WsVarliklarGlobal.Range("I" & IlkSira & ":I" & SonSira).Merge
                i = SonSira
            End If

        End If
    Next i
End If

SatirSilTekrar:
SayEsasVarlik = SayEsasVarlik - 1
If SayEsasVarlik < 7 Then
    GoTo SatirSilTekrarAtla
End If
For i = 7 To SayEsasVarlik
    If WsVarliklarGlobal.Range("R" & i) = "" Then
        WsVarliklarGlobal.Rows(i).EntireRow.Delete
        GoTo SatirSilTekrar
    End If
Next i
SatirSilTekrarAtla:

If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
Else
    SayEsasVarlik = SayEsasVarlik + 1
End If


'_____________________________________________________________BİTİŞ


'____________________İkinci düzenleme bölümü, BAŞLANGIÇ (Bu bölüm sadece XXXMud'den gelen raporları/öğelerin çıkışı içindir.)

SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'Raporların rapro no.'ları ile onun alt kırılımını sayısal bir değere dönüştür. Bu özellik raporları kendi içinde sıralamada kullanılacak.
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR ve RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'R
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklarGlobal.Range("D" & i)
            End If
        End If
        'B
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo
        '_________________________
        For a = 1 To 50
            StrRaporNo = Replace(StrRaporNo, " ", "") 'Rapor no içinde varsa boşlukları kaldır
        Next a
        NumRaporNo = Replace(StrRaporNo, "-", "0000") '1-1'in 100001 gerçek rapor no ile çakışma ihtimali 0'a yakın.

        '_________________________

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir21
            End If
        Next j
DonguBitir21:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik1
        End If

        For j = IlkSiraVarlik To SonSiraVarlik
            If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 And WsVarliklarGlobal.Range("S" & j).Value = "" Then
                WsVarliklarGlobal.Range("S" & j).Value = NumRaporNo
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklarGlobal.Range("D" & a + 1) = "" And WsVarliklarGlobal.Range("R" & a + 1) <> "" And WsVarliklarGlobal.Range("S" & a + 1).Value = "" Then
                            WsVarliklarGlobal.Range("S" & a + 1).Value = NumRaporNo
                        Else
                            GoTo aDonguSonu2
                        End If
                    End If
                Next a
            End If
        Next j
aDonguSonu2:
DonguSonuIlkSiraVarlik1:
    Next i
End If


'XXXMud'den gelen ve varlıkdaki raporlar. SIRALAMA ve BİRLEŞTİRME
For i = 7 To SayEsasVarlik - 1
    If WsVarliklarGlobal.Range("T" & i) = "" Then 'And WsVarliklarGlobal.Range("H" & i) <> "" Then
        If Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklarGlobal.Range("D" & i)
            End If
        End If
        If Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir22
            End If
        Next j
DonguBitir22:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo
        '_______________


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik
        End If

        WsVarliklarGlobal.Range("U" & IlkSiraVarlik) = IlkSiraVarlik
        WsVarliklarGlobal.Range("V" & SonSiraVarlik) = SonSiraVarlik
        WsVarliklarGlobal.Range("T" & IlkSiraVarlik & ":T" & SonSiraVarlik).Value = "*" 'Aynı satırlarda tekrar işlem yapmasın diye
        'Sorting Z to A by NumRaporNo
        WsVarliklarGlobal.Range("D" & IlkSiraVarlik & ":S" & SonSiraVarlik).UnMerge
        WsVarliklarGlobal.Range("D" & IlkSiraVarlik & ":S" & SonSiraVarlik).Sort key1:=Range("S" & IlkSiraVarlik & ":S" & SonSiraVarlik), order1:=xlAscending, Header:=xlNo

        'Rapor no.'ları ve Package A/Package B/Package C satırlarını birleştir.
        For j = IlkSiraVarlik To SonSiraVarlik
            If WsVarliklarGlobal.Range("S" & j).Value <> "" Then
                If WsVarliklarGlobal.Range("S" & j).Value = WsVarliklarGlobal.Range("S" & j + 1).Value Then
                    WsVarliklarGlobal.Range("D" & j & ":D" & j + 1).Merge
                End If
            End If
            If WsVarliklarGlobal.Range("J" & j).Value = "Package A" Or WsVarliklarGlobal.Range("J" & j).Value = "Package B" Or WsVarliklarGlobal.Range("J" & j).Value = "Package C" Then
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklarGlobal.Range("J" & a + 1).Value = "" Then
                            WsVarliklarGlobal.Range("J" & a & ":J" & a + 1).Merge
                            WsVarliklarGlobal.Range("K" & a & ":K" & a + 1).Merge
                            WsVarliklarGlobal.Range("L" & a & ":L" & a + 1).Merge
                            WsVarliklarGlobal.Range("M" & a & ":M" & a + 1).Merge
                            WsVarliklarGlobal.Range("N" & a & ":N" & a + 1).Merge
                            WsVarliklarGlobal.Range("O" & a & ":O" & a + 1).Merge
                        Else
                            GoTo aDonguSonu
                        End If
                    End If
                Next a
            End If
aDonguSonu:
        Next j
    End If
DonguSonuIlkSiraVarlik:
Next i

IlkSira = 0
SonSira = 0
For i = 7 To SayEsasVarlik - 1
    If WsVarliklarGlobal.Range("T" & i) <> "" Then
'        IlkSira = 0
'        SonSira = 0
        For j = i To SayEsasVarlik - 1
            If IlkSira = 0 And WsVarliklarGlobal.Range("U" & j) <> "" Then
                IlkSira = j
            End If
            If SonSira = 0 And WsVarliklarGlobal.Range("V" & j) <> "" Then
                SonSira = j
            End If

            If IlkSira <> 0 And SonSira <> 0 Then
                If IlkSira <> SonSira Then
                    WsVarliklarGlobal.Range("C" & IlkSira & ":C" & SonSira).Merge
                    WsVarliklarGlobal.Range("E" & IlkSira & ":E" & SonSira).Merge
                    WsVarliklarGlobal.Range("F" & IlkSira & ":F" & SonSira).Merge
                    WsVarliklarGlobal.Range("G" & IlkSira & ":G" & SonSira).Merge
                    WsVarliklarGlobal.Range("H" & IlkSira & ":H" & SonSira).Merge
                    WsVarliklarGlobal.Range("I" & IlkSira & ":I" & SonSira).Merge
                    IlkSira = 0
                    SonSira = 0
                Else
                    WsVarliklarGlobal.Range("C" & IlkSira & ":O" & IlkSira).EntireRow.AutoFit
                    IlkSira = 0
                    SonSira = 0
                End If
                GoTo iDonguDevamBirlestirmeTamam
            Else
                GoTo iDonguDevamBirlestirmeYok
            End If
        Next j
    End If
iDonguDevamBirlestirmeTamam:
iDonguDevamBirlestirmeYok:
Next i


            
'Artıkları temizle
If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
End If
WsVarliklarGlobal.Range("P" & 7 & ":V" & SayEsasVarlik).ClearContents

'____________________İkinci düzenleme bölümü, BİTİŞ


DuzenlemeyiAtla:


'Öğelere göre toplamlar
WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 3).Value = "*"
x = SayEsasVarlik + 3
For i = 7 To SayEsasVarlik - 1
    'WsVarliklarGlobal.Range("C" & i & ":N" & i).EntireRow.AutoFit
    'Mevcutlar
    If WsVarliklarGlobal.Range("J" & i) <> "" And WsVarliklarGlobal.Range("K" & i) <> "" Then
        Say2 = WsVarliklarGlobal.Range("J100000").End(xlUp).Row '+ 1
        'MsgBox "SayEsasVarlik: " & SayEsasVarlik & " Say2: " & Say2
        For j = SayEsasVarlik + 2 To Say2
            'Öğe aynı
            If WsVarliklarGlobal.Range("J" & j) = WsVarliklarGlobal.Range("J" & i) Then 'Öğe aynı
                If WsVarliklarGlobal.Range("L" & i) <> "" Or WsVarliklarGlobal.Range("M" & i) <> "" Or WsVarliklarGlobal.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    'Öğe ve öğe değeri aynı
                    If WsVarliklarGlobal.Range("K" & j) = WsVarliklarGlobal.Range("K" & i) Then
                        WsVarliklarGlobal.Range("L" & j) = WsVarliklarGlobal.Range("L" & j) + WsVarliklarGlobal.Range("L" & i)
                        WsVarliklarGlobal.Range("M" & j) = WsVarliklarGlobal.Range("M" & j) + WsVarliklarGlobal.Range("M" & i)
                        WsVarliklarGlobal.Range("N" & j) = WsVarliklarGlobal.Range("N" & j) + WsVarliklarGlobal.Range("N" & i)
                        GoTo jDonguSon1
                    Else 'Öğe aynı, öğe değeri farklı
                        If WsVarliklarGlobal.Range("J" & j + 1) = "" Then
                            WsVarliklarGlobal.Range("J" & j + 1) = WsVarliklarGlobal.Range("J" & i)
                            WsVarliklarGlobal.Range("K" & j + 1) = WsVarliklarGlobal.Range("K" & i)
                            WsVarliklarGlobal.Range("L" & j + 1) = WsVarliklarGlobal.Range("L" & i)
                            WsVarliklarGlobal.Range("M" & j + 1) = WsVarliklarGlobal.Range("M" & i)
                            WsVarliklarGlobal.Range("N" & j + 1) = WsVarliklarGlobal.Range("N" & i)
                            x = x + 1
                        End If
                    End If
                End If
            Else 'Öğe farklı
                If WsVarliklarGlobal.Range("L" & i) <> "" Or WsVarliklarGlobal.Range("M" & i) <> "" Or WsVarliklarGlobal.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    If WsVarliklarGlobal.Range("J" & j + 1) = "" Then
                        WsVarliklarGlobal.Range("J" & j + 1) = WsVarliklarGlobal.Range("J" & i)
                        WsVarliklarGlobal.Range("K" & j + 1) = WsVarliklarGlobal.Range("K" & i)
                        WsVarliklarGlobal.Range("L" & j + 1) = WsVarliklarGlobal.Range("L" & i)
                        WsVarliklarGlobal.Range("M" & j + 1) = WsVarliklarGlobal.Range("M" & i)
                        WsVarliklarGlobal.Range("N" & j + 1) = WsVarliklarGlobal.Range("N" & i)
                        x = x + 1
                    End If
                End If
            End If
        Next j
jDonguSon1:
    End If
Next i

'Boş satırlara 0 yaz.
For j = SayEsasVarlik + 2 To x
    If WsVarliklarGlobal.Range("L" & j) = "" Then
        WsVarliklarGlobal.Range("L" & j).Value = 0
    End If
    If WsVarliklarGlobal.Range("M" & j) = "" Then
        WsVarliklarGlobal.Range("M" & j).Value = 0
    End If
    If WsVarliklarGlobal.Range("N" & j) = "" Then
        WsVarliklarGlobal.Range("N" & j).Value = 0
    End If
Next j

WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 2).ClearContents
WsVarliklarGlobal.Range("C" & SayEsasVarlik + 1 & ":O" & SayEsasVarlik + 1).Merge
WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":K" & SayEsasVarlik).Merge
WsVarliklarGlobal.Range("C" & SayEsasVarlik) = "Total"
WsVarliklarGlobal.Range("C" & SayEsasVarlik).Font.Bold = True
WsVarliklarGlobal.Range("C" & SayEsasVarlik).HorizontalAlignment = xlRight
WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":O" & SayEsasVarlik).Font.Bold = True
'Toplamlar (Sol üstteki toplam)
For i = 7 To SayEsasVarlik - 1
    WsVarliklarGlobal.Range("O" & i).Value = WsVarliklarGlobal.Range("L" & i).Value + WsVarliklarGlobal.Range("M" & i).Value - WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("L" & i).Value
    WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("M" & i).Value
    WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("O" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("O" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("O" & i).Value
Next i

'ÖZET TABLO BAŞLIĞI
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2).Value = "Item-Based Totals (Qty)"
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Merge
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft

WsVarliklarGlobal.Range("J" & SayEsasVarlik + 3).Value = "Item Type"
WsVarliklarGlobal.Range("K" & SayEsasVarlik + 3).Value = "Item Value"
WsVarliklarGlobal.Range("L" & SayEsasVarlik + 3).Value = "Current Qty"
WsVarliklarGlobal.Range("M" & SayEsasVarlik + 3).Value = "Inbound Qty"
WsVarliklarGlobal.Range("N" & SayEsasVarlik + 3).Value = "Outbound Qty"
WsVarliklarGlobal.Range("O" & SayEsasVarlik + 3).Value = "Remaining Qty"
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 3 & ":O" & SayEsasVarlik + 3).Font.Bold = True
y = SayEsasVarlik + 3

'Toplamlar (öğelere göre, sağ alttaki)
For i = SayEsasVarlik + 4 To x
    WsVarliklarGlobal.Range("O" & i).Value = WsVarliklarGlobal.Range("L" & i).Value + WsVarliklarGlobal.Range("M" & i).Value - WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("L" & x + 1).Value = WsVarliklarGlobal.Range("L" & x + 1).Value + WsVarliklarGlobal.Range("L" & i).Value
    WsVarliklarGlobal.Range("M" & x + 1).Value = WsVarliklarGlobal.Range("M" & x + 1).Value + WsVarliklarGlobal.Range("M" & i).Value
    WsVarliklarGlobal.Range("N" & x + 1).Value = WsVarliklarGlobal.Range("N" & x + 1).Value + WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("O" & x + 1).Value = WsVarliklarGlobal.Range("O" & x + 1).Value + WsVarliklarGlobal.Range("O" & i).Value
Next i
'Sort Z to A
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 4 & ":O" & x).Sort key1:=Range("K" & SayEsasVarlik + 4 & ":K" & x), order1:=xlDescending, Header:=xlNo

''Özet bölümünde yer alan mevcut adet ve giren adedi sil.
'WsVarliklarGlobal.Range("L" & SayEsasVarlik + 2 & ":M" & x + 1).ClearContents
''Özet satırlarını birleştir
'For i = SayEsasVarlik + 2 To x + 1
'    WsVarliklarGlobal.Range("K" & i & ":M" & i).Merge
'    WsVarliklarGlobal.Range("K" & i & ":M" & i).HorizontalAlignment = xlLeft
'Next i

WsVarliklarGlobal.Range("L" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("M" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("N" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("O" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("J" & x + 1) = "Total"
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).Merge
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).HorizontalAlignment = xlLeft

'Kenarlıklar.
SayEsasVarlik = x + 1
Set Kenarlar = WsVarliklarGlobal.Range("C" & 7 & ":O" & SayEsasVarlik)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
WsVarliklarGlobal.Range("C" & y - 2 & ":I" & x + 2).Borders.LineStyle = xlNone
WsVarliklarGlobal.Range("J" & y - 1 & ":J" & x + 1).Borders((xlLeft)).LineStyle = xlNone

Son:


End Sub

Sub Rapor2_2Varlik()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object ', WsVarliklarGlobal As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String, XXXMudGiden As String, XXXMudGelen As String

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 2.2 – Supporting Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

    


On Error Resume Next 'Operation içinde Varlıklar.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 2.2 – Supporting Asset Report.xlsx"
'Varlıklar açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 2.2 – Supporting Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 2.2 – Supporting Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklarGlobal = Workbooks("Report 2.2 – Supporting Asset Report.xlsx").Worksheets(1)
WsVarliklarGlobal.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklarGlobal.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

'XXXMudGiden = " (XXXMud-Giden)"
XXXMudGelen = "ORGANIZATION A XXX Directorate"
            
SayEsasVarlik = 7
SiraNoEsasVarlik = 1
'Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
'NOT: Tüm yardımcı varlıklarda giriş işleminde çıkış tarihi varsa kaldır.
'Çıkış işleminde giriş tarihi olması normal; ama giriş işleminde çıkış tarihi olamaz.
'Benzer şekilde mevcutlar da çıkış tarihi olamaz. Mevcutlarda da çıkış tarihi varsa kaldır.
'07.10.2019, 01:30 İshak.
'YAPILDI!
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (GİRİŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
            
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik
           
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
    
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i

                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
        
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                                
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If

            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            'WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge
    
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl
   
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (ÇIKIŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    Else
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        SayEsasVarlikDongu = SayEsasVarlik
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)

                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        Else
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        End If
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                        SayEsasVarlikDongu = SayEsasVarlik
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
                        
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        Else
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        End If
                        WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    Else
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        Else
                            WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        End If
                        WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            Else
                                WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            End If
                            WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        SayEsasVarlikDongu = SayEsasVarlik
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)

                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        Else
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        End If
                        
                        WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        SayEsasVarlikDongu = SayEsasVarlik
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                                
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
    
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                    WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                Else
                                    WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                End If
                                WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge

            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (MEVCUT, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik

                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
                        
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1

                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik

                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                                
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
    
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


If WsVarliklarGlobal.Range("C7") = "" Then
    GoTo Son
End If


'Refere prosedür
SayEsasVarlikGlobal = SayEsasVarlik 'SayEsasVarlik değişkenini aktar
SiraNoEsasVarlikGlobal = SiraNoEsasVarlik
Call Rapor2_2VarlikDevam


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub Rapor3Varlik()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 3 – Supporting Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Varlıklar.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 3 – Supporting Asset Report.xlsx"
'Varlıklar açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 3 – Supporting Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 3 – Supporting Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 3 – Supporting Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
'Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor3 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i
                                
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
            
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
            
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
 
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
            
'            'ÖR yaz
'            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
'                If WsVarliklar.Range("D" & i) <> "" Then
'                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
'                End If
'            Next i

            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor3 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
                
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                    WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
                Else
                    WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
                End If
                WsVarliklar.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
            
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                    WsVarliklar.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
                Else
                    WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
                End If
                WsVarliklar.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
            
'            'ÖR yaz
'            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
'                If WsVarliklar.Range("D" & i) <> "" Then
'                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
'                End If
'            Next i

            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'Rapor3 (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
                
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
                
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklar.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklar.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklar.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                'WsVarliklar.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
                'WsVarliklar.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                WsVarliklar.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklar.Range("D" & i) <> "" Then
                            WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
'            'ÖR yaz
'            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
'                If WsVarliklar.Range("D" & i) <> "" Then
'                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
'                End If
'            Next i

            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklar.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

If WsVarliklar.Range("C7") = "" Then
    GoTo Son
End If


'Öğelere göre toplamlar
WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 3).Value = "*"
x = SayEsasVarlik + 3
For i = 7 To SayEsasVarlik - 1
    'WsVarliklar.Range("C" & i & ":N" & i).EntireRow.AutoFit
    'Mevcutlar
    If WsVarliklar.Range("J" & i) <> "" And WsVarliklar.Range("K" & i) <> "" Then
        Say2 = WsVarliklar.Range("J100000").End(xlUp).Row '+ 1
        'MsgBox "SayEsasVarlik: " & SayEsasVarlik & " Say2: " & Say2
        For j = SayEsasVarlik + 2 To Say2
            'Öğe aynı
            If WsVarliklar.Range("J" & j) = WsVarliklar.Range("J" & i) Then 'Öğe aynı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    'Öğe ve öğe değeri aynı
                    If WsVarliklar.Range("K" & j) = WsVarliklar.Range("K" & i) Then
                        WsVarliklar.Range("L" & j) = WsVarliklar.Range("L" & j) + WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j) = WsVarliklar.Range("M" & j) + WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j) = WsVarliklar.Range("N" & j) + WsVarliklar.Range("N" & i)
                        GoTo jDonguSon1
                    Else 'Öğe aynı, öğe değeri farklı
                        If WsVarliklar.Range("J" & j + 1) = "" Then
                            WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                            WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                            WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                            WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                            WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                            x = x + 1
                        End If
                    End If
                End If
            Else 'Öğe farklı
                If WsVarliklar.Range("L" & i) <> "" Or WsVarliklar.Range("M" & i) <> "" Or WsVarliklar.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    If WsVarliklar.Range("J" & j + 1) = "" Then
                        WsVarliklar.Range("J" & j + 1) = WsVarliklar.Range("J" & i)
                        WsVarliklar.Range("K" & j + 1) = WsVarliklar.Range("K" & i)
                        WsVarliklar.Range("L" & j + 1) = WsVarliklar.Range("L" & i)
                        WsVarliklar.Range("M" & j + 1) = WsVarliklar.Range("M" & i)
                        WsVarliklar.Range("N" & j + 1) = WsVarliklar.Range("N" & i)
                        x = x + 1
                    End If
                End If
            End If
        Next j
jDonguSon1:
    End If
Next i

'Boş satırlara 0 yaz.
For j = SayEsasVarlik + 2 To x
    If WsVarliklar.Range("L" & j) = "" Then
        WsVarliklar.Range("L" & j).Value = 0
    End If
    If WsVarliklar.Range("M" & j) = "" Then
        WsVarliklar.Range("M" & j).Value = 0
    End If
    If WsVarliklar.Range("N" & j) = "" Then
        WsVarliklar.Range("N" & j).Value = 0
    End If
Next j

WsVarliklar.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 2).ClearContents
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":O" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik & ":K" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik) = "Total"
WsVarliklar.Range("C" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik).HorizontalAlignment = xlRight
WsVarliklar.Range("L" & SayEsasVarlik & ":O" & SayEsasVarlik).Font.Bold = True
'Toplamlar (Sol üstteki toplam)
For i = 7 To SayEsasVarlik - 1
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & SayEsasVarlik).Value = WsVarliklar.Range("L" & SayEsasVarlik).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & SayEsasVarlik).Value = WsVarliklar.Range("M" & SayEsasVarlik).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & SayEsasVarlik).Value = WsVarliklar.Range("N" & SayEsasVarlik).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & SayEsasVarlik).Value = WsVarliklar.Range("O" & SayEsasVarlik).Value + WsVarliklar.Range("O" & i).Value
Next i

'ÖZET TABLO BAŞLIĞI
WsVarliklar.Range("J" & SayEsasVarlik + 2).Value = "Item-Based Totals (Qty)"
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft

WsVarliklar.Range("J" & SayEsasVarlik + 3).Value = "Item Type"
WsVarliklar.Range("K" & SayEsasVarlik + 3).Value = "Item Value"
WsVarliklar.Range("L" & SayEsasVarlik + 3).Value = "Current Qty"
WsVarliklar.Range("M" & SayEsasVarlik + 3).Value = "Inbound Qty"
WsVarliklar.Range("N" & SayEsasVarlik + 3).Value = "Outbound Qty"
WsVarliklar.Range("O" & SayEsasVarlik + 3).Value = "Remaining Qty"
WsVarliklar.Range("J" & SayEsasVarlik + 3 & ":O" & SayEsasVarlik + 3).Font.Bold = True
y = SayEsasVarlik + 3

'Toplamlar (öğelere göre, sağ alttaki)
For i = SayEsasVarlik + 4 To x
    WsVarliklar.Range("O" & i).Value = WsVarliklar.Range("L" & i).Value + WsVarliklar.Range("M" & i).Value - WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("L" & x + 1).Value = WsVarliklar.Range("L" & x + 1).Value + WsVarliklar.Range("L" & i).Value
    WsVarliklar.Range("M" & x + 1).Value = WsVarliklar.Range("M" & x + 1).Value + WsVarliklar.Range("M" & i).Value
    WsVarliklar.Range("N" & x + 1).Value = WsVarliklar.Range("N" & x + 1).Value + WsVarliklar.Range("N" & i).Value
    WsVarliklar.Range("O" & x + 1).Value = WsVarliklar.Range("O" & x + 1).Value + WsVarliklar.Range("O" & i).Value
Next i
'Sort Z to A
WsVarliklar.Range("J" & SayEsasVarlik + 4 & ":O" & x).Sort key1:=Range("K" & SayEsasVarlik + 4 & ":K" & x), order1:=xlDescending, Header:=xlNo

WsVarliklar.Range("L" & x + 1).Font.Bold = True
WsVarliklar.Range("M" & x + 1).Font.Bold = True
WsVarliklar.Range("N" & x + 1).Font.Bold = True
WsVarliklar.Range("O" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1) = "Total"
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Font.Bold = True
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).Merge
WsVarliklar.Range("J" & x + 1 & ":K" & x + 1).HorizontalAlignment = xlLeft

'Kenarlıklar.
SayEsasVarlik = x + 1
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":O" & SayEsasVarlik)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
WsVarliklar.Range("C" & y - 2 & ":I" & x + 2).Borders.LineStyle = xlNone
WsVarliklar.Range("J" & y - 1 & ":J" & x + 1).Borders((xlLeft)).LineStyle = xlNone


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub GenelVarlik()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object ', WsVarliklarGlobal As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String, XXXMudGiden As String, XXXMudGelen As String

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "General Supporting Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Varlıklar.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "General Supporting Asset Report.xlsx"
'Varlıklar açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("General Supporting Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "General Supporting Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklarGlobal = Workbooks("General Supporting Asset Report.xlsx").Worksheets(1)
WsVarliklarGlobal.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklarGlobal.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
'Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "ÖR/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'RAPOR2_2 (GİRİŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
            
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik
           
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
    
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i

                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
        
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                                
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If

            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            'WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge
    
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor3 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i
                    
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If
                
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i
                
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            Else
                WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            End If
            WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "ÖR/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal

            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            Else
                WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            End If
            WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'RAPOR2_2 (ÇIKIŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    Else
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        SayEsasVarlikDongu = SayEsasVarlik
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
    '                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GL" & IlkSira)
    '                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GK" & IlkSira)
    '                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GI" & IlkSira)
    '                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)

                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        Else
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        End If
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                        SayEsasVarlikDongu = SayEsasVarlik
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
                        
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
    '                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GL" & IlkSira)
    '                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GK" & IlkSira)
    '                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("GI" & IlkSira)
    '                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        Else
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        End If
                        WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    Else
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        Else
                            WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        End If
                        WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DA" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            Else
                                WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            End If
                            WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        SayEsasVarlikDongu = SayEsasVarlik
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)

                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        Else
                            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                                WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                            Else
                                WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                            End If
                        End If
                        
                        WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        SayEsasVarlikDongu = SayEsasVarlik
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("DB" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                                
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
    
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                    WsVarliklarGlobal.Range("M" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                Else
                                    WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                End If
                                WsVarliklarGlobal.Range("N" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge

            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor3 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
                End If
                WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If
                
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                    WsVarliklarGlobal.Range("M" & SayEsasVarlik & ":M" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
                End If
                WsVarliklarGlobal.Range("N" & SayEsasVarlik & ":N" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
                
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

SayEsasVarlikGlobal = SayEsasVarlik 'SayEsasVarlik değişkenini aktar
SiraNoEsasVarlikGlobal = SiraNoEsasVarlik
Call GenelVarlikMevcutDevam


If WsVarliklarGlobal.Range("C7") = "" Then
    GoTo Son
End If

SayEsasVarlik = SayEsasVarlikGlobal
SiraNoEsasVarlik = SiraNoEsasVarlikGlobal
SayEsasVarlikGlobal = SayEsasVarlik 'SayEsasVarlik değişkenini aktar
SiraNoEsasVarlikGlobal = SiraNoEsasVarlik
Call GenelVarlikDevam

Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub
Sub GenelVarlikMevcutDevam()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object ', WsVarliklarGlobal As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String, XXXMudGiden As String, XXXMudGelen As String

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long

SayEsasVarlik = SayEsasVarlikGlobal
SiraNoEsasVarlik = SiraNoEsasVarlikGlobal

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor1GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i


            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("U" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("T" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("AE" & IlkSira)
            'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(3).Range("CR" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "ÖR/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call RaporGelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
            WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
            WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
            WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
            'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
            WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            
            WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) <> "" Then
                    WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'RAPOR2_2 (MEVCUT, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If

            'Row height
            If IlkSira = SonSira Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            
            Call Rapor2_2GelenTema
            GelenTema = GelenTemaGlobal
            
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik

                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    Else
                    
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
    
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomSay = 0
                        NomDeger = 0
                        For i = IlkSira To SonSira
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                            NomSay = NomSay + 1
                        Next i
                        
                        WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                        'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                        WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                    End If
                Else
                    SayEsasVarlikDongu = SayEsasVarlik
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                        'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    j = IlkSira - 1
                    WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                    WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                    WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                    WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                    'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger

                            'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                ElseIf Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    SayEsasVarlikDongu = SayEsasVarlik
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        WsVarliklarGlobal.Range("K" & SayEsasVarlik).Value = "-"
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklarGlobal.Range("J" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    Else
                        j = IlkSira - 1
                        WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                        WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema 'XXXMudGelen
                        WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AC" & IlkSira)
                        WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AB" & IlkSira)
                        WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Range("AM" & IlkSira)
                        'WsVarliklarGlobal.Range("I" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("CZ" & IlkSira)
                        
                        NomDeger = 0
                        For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                            j = j + 1
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                                WsVarliklarGlobal.Range("J" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
    
                                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                NomDeger = CLng(NomDeger)
                                WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu & ":K" & SayEsasVarlikDongu).Value = NomDeger
    
                                'WsVarliklarGlobal.Range("K" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                                WsVarliklarGlobal.Range("L" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                                SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                            End If
                        Next i
                        
                        WsVarliklarGlobal.Range("Q" & SayEsasVarlik & ":Q" & SayEsasVarlikDongu - 1).Value = "*"
                    End If
                End If
            End If
            'Row height
            If SayEsasVarlik = SayEsasVarlikDongu - 1 Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik).EntireRow.AutoFit
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklarGlobal.Range("D" & i) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklarGlobal.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklarGlobal.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlikDongu - 1).Merge

            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("Q" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If

            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If

        'Rapor3 (MEVCUT)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Or Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Call Rapor1GelenTema
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                GelenTema = "Report 3.2"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                GelenTema = "Report 3.1"
            End If
            
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("W" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
                WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If

            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
                'Aktarımı başlat
                WsVarliklarGlobal.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                'Öğe Değeri string'ten long'a dönüştür ve yaz.
                NomSay = 0
                NomDeger = 0
                For i = IlkSira To SonSira
                    NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                    NomDeger = CLng(NomDeger)
                    WsVarliklarGlobal.Range("K" & SayEsasVarlik + NomSay & ":K" & SayEsasVarlik + NomSay).Value = NomDeger
                    NomSay = NomSay + 1
                Next i
  
                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
                Else
                    WsVarliklarGlobal.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
                End If
                
                WsVarliklarGlobal.Range("E" & SayEsasVarlik).Value = GelenTema
                WsVarliklarGlobal.Range("F" & SayEsasVarlik).Value = "#"
                WsVarliklarGlobal.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                WsVarliklarGlobal.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("CQ" & IlkSira)
                'WsVarliklarGlobal.Range("I" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(5).Range("FT" & IlkSira)
                WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value
                'WsVarliklarGlobal.Range("K" & SayEsasVarlik & ":K" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
                WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":L" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            
                WsVarliklarGlobal.Range("R" & SayEsasVarlik & ":R" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"

                If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                    'R yaz
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        If WsVarliklarGlobal.Range("D" & i) <> "" Then
                            WsVarliklarGlobal.Range("D" & i) = "R/" & WsVarliklarGlobal.Range("D" & i)
                        End If
                    Next i
                End If
                
            End If
'            'ÖR yaz
'            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
'                If WsVarliklarGlobal.Range("D" & i) <> "" Then
'                    WsVarliklarGlobal.Range("D" & i) = "ÖR/" & WsVarliklarGlobal.Range("D" & i)
'                End If
'            Next i

            'Sıra ve kayıt nolar ile gönderen, belge no ve tarih, giriş ve çıkış tarihini dikeyde birleştir.
            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            WsVarliklarGlobal.Range("I" & SayEsasVarlik & ":I" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklarGlobal.Range("D" & i) = "" And WsVarliklarGlobal.Range("J" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklarGlobal.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl

SayEsasVarlikGlobal = SayEsasVarlik
SiraNoEsasVarlikGlobal = SiraNoEsasVarlik

End Sub

Sub GenelVarlikDevam()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object ', WsVarliklarGlobal As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim Say2 As Long, y As Long
Dim GelenTema As String, XXXMudGiden As String, XXXMudGelen As String

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long


'GoTo DuzenlemeyiAtla

'_____________________________________________________________BAŞLANGIÇ (Direkt kaldırsan yine de sorunsuz çalışır.)

Dim KontB As Integer, KontR As Integer, SayEsasVarlikx As Long
Dim SayEsasVarlikTakip As Long, a As Integer, b As Integer
Dim NumRaporNo As Variant

'GoTo Son

SayEsasVarlik = SayEsasVarlikGlobal
SiraNoEsasVarlik = SiraNoEsasVarlikGlobal

SayEsasVarlikTakip = 0
KontB = 0
KontR = 0
SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'MsgBox SayEsasVarlik

'varlıklar düzeltme faktörü (Technique A ve diğerlerinin sıra numarasını birleştir.)
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox "2: " & StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir1
                    End If
                Next j
DonguBitir1:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontB = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontB = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon1
                    End If
                Next x
xDonguSon1:

'                'Row height
'                If SonAltSatirNo = IlkAltSatirNo Then
'                    WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                End If

                WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklarGlobal.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon2
                            End If
                        Next x
xDonguSon2:

'                        'Row height
'                        If SonAltSatirNo = IlkAltSatirNo Then
'                            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                        End If

                        WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                        WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip

    'RAPOR NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo

                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir2
                    End If
                Next j
DonguBitir2:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontR = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontR = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                        GoTo xDonguSon3
                    End If
                Next x
xDonguSon3:

'                'Row height
'                If SonAltSatirNo = IlkAltSatirNo Then
'                    WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                End If

                WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklarGlobal.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) <> "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklarGlobal.Range("D" & x) <> "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklarGlobal.Range("D" & x) = "" And WsVarliklarGlobal.Range("Q" & x) = "" Then
                                GoTo xDonguSon4
                            End If
                        Next x
xDonguSon4:

'                        'Row height
'                        If SonAltSatirNo = IlkAltSatirNo Then
'                            WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":N" & SayEsasVarlik + SonAltSatirNo - IlkAltSatirNo).EntireRow.AutoFit
'                        End If

                        WsVarliklarGlobal.Range("C" & SayEsasVarlikTakip & ":O" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":O" & SonAltSatirNo).Value
                        WsVarliklarGlobal.Range("P" & SayEsasVarlikTakip & ":P" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklarGlobal.Range("R" & SayEsasVarlikTakip & ":R" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklarGlobal.Range("P" & IlkAltSatirNo & ":P" & SonAltSatirNo).Value = WsVarliklarGlobal.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).UnMerge
                        WsVarliklarGlobal.Range("C" & IlkAltSatirNo & ":Q" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

'GoTo Son

     SayEsasVarlik = SayEsasVarlikTakip

    'Sıra no.lar
    SiraNoEsasVarlik = 0
    'WsVarliklarGlobal.Range("C" & 7 & ":C" & SayEsasVarlik).ClearContents
    For i = 7 To SayEsasVarlik
        If WsVarliklarGlobal.Range("C" & i) <> "" And WsVarliklarGlobal.Range("P" & i) = "" Then
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            WsVarliklarGlobal.Range("C" & i) = SiraNoEsasVarlik
        End If
        If WsVarliklarGlobal.Range("P" & i) <> "" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=WsVarliklarGlobal.Range("P" & i), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=WsVarliklarGlobal.Range("P" & i), SearchDirection:=xlPrevious, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
'            MsgBox IlkSira & " : " & SonSira
            If IlkSira = SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklarGlobal.Range("C" & i) = SiraNoEsasVarlik
                'Row height
                WsVarliklarGlobal.Range("C" & IlkSira & ":O" & IlkSira).EntireRow.AutoFit
                i = SonSira
            ElseIf IlkSira <> SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklarGlobal.Range("C" & IlkSira) = SiraNoEsasVarlik
                WsVarliklarGlobal.Range("C" & IlkSira & ":C" & SonSira).Merge
                WsVarliklarGlobal.Range("E" & IlkSira & ":E" & SonSira).Merge
                WsVarliklarGlobal.Range("F" & IlkSira & ":F" & SonSira).Merge
                WsVarliklarGlobal.Range("G" & IlkSira & ":G" & SonSira).Merge
                WsVarliklarGlobal.Range("H" & IlkSira & ":H" & SonSira).Merge
                WsVarliklarGlobal.Range("I" & IlkSira & ":I" & SonSira).Merge
                i = SonSira
            End If

        End If
    Next i
End If

SatirSilTekrar:
SayEsasVarlik = SayEsasVarlik - 1
If SayEsasVarlik < 7 Then
    GoTo SatirSilTekrarAtla
End If
For i = 7 To SayEsasVarlik
    If WsVarliklarGlobal.Range("R" & i) = "" Then
        WsVarliklarGlobal.Rows(i).EntireRow.Delete
        GoTo SatirSilTekrar
    End If
Next i
SatirSilTekrarAtla:

If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
Else
    SayEsasVarlik = SayEsasVarlik + 1
End If


'_____________________________________________________________BİTİŞ


'____________________İkinci düzenleme bölümü, BAŞLANGIÇ (Bu bölüm sadece XXXMud'den gelen raporları/öğelerin çıkışı içindir.)

SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'Raporların rapro no.'ları ile onun alt kırılımını sayısal bir değere dönüştür. Bu özellik raporları kendi içinde sıralamada kullanılacak.
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR ve RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'R
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklarGlobal.Range("D" & i)
            End If
        End If
        'B
        If WsVarliklarGlobal.Range("D" & i) <> "" And Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo
        '_________________________
        For a = 1 To 50
            StrRaporNo = Replace(StrRaporNo, " ", "") 'Rapor no içinde varsa boşlukları kaldır
        Next a
        NumRaporNo = Replace(StrRaporNo, "-", "0000") '1-1'in 100001 gerçek rapor no ile çakışma ihtimali 0'a yakın.

        '_________________________

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir21
            End If
        Next j
DonguBitir21:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik1
        End If

        For j = IlkSiraVarlik To SonSiraVarlik
            If InStr(WsVarliklarGlobal.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 And WsVarliklarGlobal.Range("S" & j).Value = "" Then
                WsVarliklarGlobal.Range("S" & j).Value = NumRaporNo
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklarGlobal.Range("D" & a + 1) = "" And WsVarliklarGlobal.Range("R" & a + 1) <> "" And WsVarliklarGlobal.Range("S" & a + 1).Value = "" Then
                            WsVarliklarGlobal.Range("S" & a + 1).Value = NumRaporNo
                        Else
                            GoTo aDonguSonu2
                        End If
                    End If
                Next a
            End If
        Next j
aDonguSonu2:
DonguSonuIlkSiraVarlik1:
    Next i
End If


'XXXMud'den gelen ve varlıkdaki raporlar. SIRALAMA ve BİRLEŞTİRME
For i = 7 To SayEsasVarlik - 1
    If WsVarliklarGlobal.Range("T" & i) = "" Then 'And WsVarliklarGlobal.Range("H" & i) <> "" Then
        If Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "B" Then
                StrRaporNo = Left(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklarGlobal.Range("D" & i)
            End If
        End If
        If Left(WsVarliklarGlobal.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklarGlobal.Range("D" & i), InStr(WsVarliklarGlobal.Range("D" & i), "(") + 1, Len(WsVarliklarGlobal.Range("D" & i)) - InStr(WsVarliklarGlobal.Range("D" & i), "(") - 1) 'R/raporno şeklinde
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir22
            End If
        Next j
DonguBitir22:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo
        '_______________


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklarGlobal.Range("P6:P100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik
        End If

        WsVarliklarGlobal.Range("U" & IlkSiraVarlik) = IlkSiraVarlik
        WsVarliklarGlobal.Range("V" & SonSiraVarlik) = SonSiraVarlik
        WsVarliklarGlobal.Range("T" & IlkSiraVarlik & ":T" & SonSiraVarlik).Value = "*" 'Aynı satırlarda tekrar işlem yapmasın diye
        'Sorting Z to A by NumRaporNo
        WsVarliklarGlobal.Range("D" & IlkSiraVarlik & ":S" & SonSiraVarlik).UnMerge
        WsVarliklarGlobal.Range("D" & IlkSiraVarlik & ":S" & SonSiraVarlik).Sort key1:=Range("S" & IlkSiraVarlik & ":S" & SonSiraVarlik), order1:=xlAscending, Header:=xlNo

        'Rapor no.'ları ve Package A/Package B/Package C satırlarını birleştir.
        For j = IlkSiraVarlik To SonSiraVarlik
            If WsVarliklarGlobal.Range("S" & j).Value <> "" Then
                If WsVarliklarGlobal.Range("S" & j).Value = WsVarliklarGlobal.Range("S" & j + 1).Value Then
                    WsVarliklarGlobal.Range("D" & j & ":D" & j + 1).Merge
                End If
            End If
            If WsVarliklarGlobal.Range("J" & j).Value = "Package A" Or WsVarliklarGlobal.Range("J" & j).Value = "Package B" Or WsVarliklarGlobal.Range("J" & j).Value = "Package C" Then
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklarGlobal.Range("J" & a + 1).Value = "" Then
                            WsVarliklarGlobal.Range("J" & a & ":J" & a + 1).Merge
                            WsVarliklarGlobal.Range("K" & a & ":K" & a + 1).Merge
                            WsVarliklarGlobal.Range("L" & a & ":L" & a + 1).Merge
                            WsVarliklarGlobal.Range("M" & a & ":M" & a + 1).Merge
                            WsVarliklarGlobal.Range("N" & a & ":N" & a + 1).Merge
                            WsVarliklarGlobal.Range("O" & a & ":O" & a + 1).Merge
                        Else
                            GoTo aDonguSonu
                        End If
                    End If
                Next a
            End If
aDonguSonu:
        Next j
    End If
DonguSonuIlkSiraVarlik:
Next i

IlkSira = 0
SonSira = 0
For i = 7 To SayEsasVarlik - 1
    If WsVarliklarGlobal.Range("T" & i) <> "" Then
'        IlkSira = 0
'        SonSira = 0
        For j = i To SayEsasVarlik - 1
            If IlkSira = 0 And WsVarliklarGlobal.Range("U" & j) <> "" Then
                IlkSira = j
            End If
            If SonSira = 0 And WsVarliklarGlobal.Range("V" & j) <> "" Then
                SonSira = j
            End If

            If IlkSira <> 0 And SonSira <> 0 Then
                If IlkSira <> SonSira Then
                    WsVarliklarGlobal.Range("C" & IlkSira & ":C" & SonSira).Merge
                    WsVarliklarGlobal.Range("E" & IlkSira & ":E" & SonSira).Merge
                    WsVarliklarGlobal.Range("F" & IlkSira & ":F" & SonSira).Merge
                    WsVarliklarGlobal.Range("G" & IlkSira & ":G" & SonSira).Merge
                    WsVarliklarGlobal.Range("H" & IlkSira & ":H" & SonSira).Merge
                    WsVarliklarGlobal.Range("I" & IlkSira & ":I" & SonSira).Merge
                    IlkSira = 0
                    SonSira = 0
                Else
                    WsVarliklarGlobal.Range("C" & IlkSira & ":O" & IlkSira).EntireRow.AutoFit
                    IlkSira = 0
                    SonSira = 0
                End If
                GoTo iDonguDevamBirlestirmeTamam
            Else
                GoTo iDonguDevamBirlestirmeYok
            End If
        Next j
    End If
iDonguDevamBirlestirmeTamam:
iDonguDevamBirlestirmeYok:
Next i


            
'Artıkları temizle
If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
End If
WsVarliklarGlobal.Range("P" & 7 & ":V" & SayEsasVarlik).ClearContents

'____________________İkinci düzenleme bölümü, BİTİŞ


DuzenlemeyiAtla:


'Öğelere göre toplamlar
WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 3).Value = "*"
x = SayEsasVarlik + 3
For i = 7 To SayEsasVarlik - 1
    'WsVarliklarGlobal.Range("C" & i & ":N" & i).EntireRow.AutoFit
    'Mevcutlar
    If WsVarliklarGlobal.Range("J" & i) <> "" And WsVarliklarGlobal.Range("K" & i) <> "" Then
        Say2 = WsVarliklarGlobal.Range("J100000").End(xlUp).Row '+ 1
        'MsgBox "SayEsasVarlik: " & SayEsasVarlik & " Say2: " & Say2
        For j = SayEsasVarlik + 2 To Say2
            'Öğe aynı
            If WsVarliklarGlobal.Range("J" & j) = WsVarliklarGlobal.Range("J" & i) Then 'Öğe aynı
                If WsVarliklarGlobal.Range("L" & i) <> "" Or WsVarliklarGlobal.Range("M" & i) <> "" Or WsVarliklarGlobal.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    'Öğe ve öğe değeri aynı
                    If WsVarliklarGlobal.Range("K" & j) = WsVarliklarGlobal.Range("K" & i) Then
                        WsVarliklarGlobal.Range("L" & j) = WsVarliklarGlobal.Range("L" & j) + WsVarliklarGlobal.Range("L" & i)
                        WsVarliklarGlobal.Range("M" & j) = WsVarliklarGlobal.Range("M" & j) + WsVarliklarGlobal.Range("M" & i)
                        WsVarliklarGlobal.Range("N" & j) = WsVarliklarGlobal.Range("N" & j) + WsVarliklarGlobal.Range("N" & i)
                        GoTo jDonguSon1
                    Else 'Öğe aynı, öğe değeri farklı
                        If WsVarliklarGlobal.Range("J" & j + 1) = "" Then
                            WsVarliklarGlobal.Range("J" & j + 1) = WsVarliklarGlobal.Range("J" & i)
                            WsVarliklarGlobal.Range("K" & j + 1) = WsVarliklarGlobal.Range("K" & i)
                            WsVarliklarGlobal.Range("L" & j + 1) = WsVarliklarGlobal.Range("L" & i)
                            WsVarliklarGlobal.Range("M" & j + 1) = WsVarliklarGlobal.Range("M" & i)
                            WsVarliklarGlobal.Range("N" & j + 1) = WsVarliklarGlobal.Range("N" & i)
                            x = x + 1
                        End If
                    End If
                End If
            Else 'Öğe farklı
                If WsVarliklarGlobal.Range("L" & i) <> "" Or WsVarliklarGlobal.Range("M" & i) <> "" Or WsVarliklarGlobal.Range("N" & i) <> "" Then 'mevcut, giriş, çıkış
                    If WsVarliklarGlobal.Range("J" & j + 1) = "" Then
                        WsVarliklarGlobal.Range("J" & j + 1) = WsVarliklarGlobal.Range("J" & i)
                        WsVarliklarGlobal.Range("K" & j + 1) = WsVarliklarGlobal.Range("K" & i)
                        WsVarliklarGlobal.Range("L" & j + 1) = WsVarliklarGlobal.Range("L" & i)
                        WsVarliklarGlobal.Range("M" & j + 1) = WsVarliklarGlobal.Range("M" & i)
                        WsVarliklarGlobal.Range("N" & j + 1) = WsVarliklarGlobal.Range("N" & i)
                        x = x + 1
                    End If
                End If
            End If
        Next j
jDonguSon1:
    End If
Next i

'Boş satırlara 0 yaz.
For j = SayEsasVarlik + 2 To x
    If WsVarliklarGlobal.Range("L" & j) = "" Then
        WsVarliklarGlobal.Range("L" & j).Value = 0
    End If
    If WsVarliklarGlobal.Range("M" & j) = "" Then
        WsVarliklarGlobal.Range("M" & j).Value = 0
    End If
    If WsVarliklarGlobal.Range("N" & j) = "" Then
        WsVarliklarGlobal.Range("N" & j).Value = 0
    End If
Next j

WsVarliklarGlobal.Range("J" & SayEsasVarlik & ":N" & SayEsasVarlik + 2).ClearContents
WsVarliklarGlobal.Range("C" & SayEsasVarlik + 1 & ":O" & SayEsasVarlik + 1).Merge
WsVarliklarGlobal.Range("C" & SayEsasVarlik & ":K" & SayEsasVarlik).Merge
WsVarliklarGlobal.Range("C" & SayEsasVarlik) = "Total"
WsVarliklarGlobal.Range("C" & SayEsasVarlik).Font.Bold = True
WsVarliklarGlobal.Range("C" & SayEsasVarlik).HorizontalAlignment = xlRight
WsVarliklarGlobal.Range("L" & SayEsasVarlik & ":O" & SayEsasVarlik).Font.Bold = True
'Toplamlar (Sol üstteki toplam)
For i = 7 To SayEsasVarlik - 1
    WsVarliklarGlobal.Range("O" & i).Value = WsVarliklarGlobal.Range("L" & i).Value + WsVarliklarGlobal.Range("M" & i).Value - WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("L" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("L" & i).Value
    WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("M" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("M" & i).Value
    WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("N" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("O" & SayEsasVarlik).Value = WsVarliklarGlobal.Range("O" & SayEsasVarlik).Value + WsVarliklarGlobal.Range("O" & i).Value
Next i

'ÖZET TABLO BAŞLIĞI
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2).Value = "Item-Based Totals (Qty)"
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).Merge
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 2 & ":O" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft

WsVarliklarGlobal.Range("J" & SayEsasVarlik + 3).Value = "Item Type"
WsVarliklarGlobal.Range("K" & SayEsasVarlik + 3).Value = "Item Value"
WsVarliklarGlobal.Range("L" & SayEsasVarlik + 3).Value = "Current Qty"
WsVarliklarGlobal.Range("M" & SayEsasVarlik + 3).Value = "Inbound Qty"
WsVarliklarGlobal.Range("N" & SayEsasVarlik + 3).Value = "Outbound Qty"
WsVarliklarGlobal.Range("O" & SayEsasVarlik + 3).Value = "Remaining Qty"
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 3 & ":O" & SayEsasVarlik + 3).Font.Bold = True
y = SayEsasVarlik + 3

'Toplamlar (öğelere göre, sağ alttaki)
For i = SayEsasVarlik + 4 To x
    WsVarliklarGlobal.Range("O" & i).Value = WsVarliklarGlobal.Range("L" & i).Value + WsVarliklarGlobal.Range("M" & i).Value - WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("L" & x + 1).Value = WsVarliklarGlobal.Range("L" & x + 1).Value + WsVarliklarGlobal.Range("L" & i).Value
    WsVarliklarGlobal.Range("M" & x + 1).Value = WsVarliklarGlobal.Range("M" & x + 1).Value + WsVarliklarGlobal.Range("M" & i).Value
    WsVarliklarGlobal.Range("N" & x + 1).Value = WsVarliklarGlobal.Range("N" & x + 1).Value + WsVarliklarGlobal.Range("N" & i).Value
    WsVarliklarGlobal.Range("O" & x + 1).Value = WsVarliklarGlobal.Range("O" & x + 1).Value + WsVarliklarGlobal.Range("O" & i).Value
Next i
'Sort Z to A
WsVarliklarGlobal.Range("J" & SayEsasVarlik + 4 & ":O" & x).Sort key1:=Range("K" & SayEsasVarlik + 4 & ":K" & x), order1:=xlDescending, Header:=xlNo


WsVarliklarGlobal.Range("L" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("M" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("N" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("O" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("J" & x + 1) = "Total"
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).Font.Bold = True
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).Merge
WsVarliklarGlobal.Range("J" & x + 1 & ":K" & x + 1).HorizontalAlignment = xlLeft


'Kenarlıklar.
SayEsasVarlik = x + 1
Set Kenarlar = WsVarliklarGlobal.Range("C" & 7 & ":O" & SayEsasVarlik)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
WsVarliklarGlobal.Range("C" & y - 2 & ":I" & x + 2).Borders.LineStyle = xlNone
WsVarliklarGlobal.Range("J" & y - 1 & ":J" & x + 1).Borders((xlLeft)).LineStyle = xlNone

Son:

End Sub

Sub Rapor1GelenTema()
Dim SiraBul As Range, IlkSira As Long

Set SiraBul = ThisWorkbook.Worksheets(3).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSira = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor1Devam1
End If

'Giden tema
GelenTemaGlobal = ""
If ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
    Else
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value
    End If
Else
    If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value)) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value))
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value
        End If
    End If
End If

Rapor1Devam1:

End Sub

Sub RaporGelenTema()
Dim SiraBul As Range, IlkSira As Long

Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSira = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo RaporDevam1
End If

'Giden tema
GelenTemaGlobal = ""
If ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
    End If
Else
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value))
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
        End If
    End If
End If

RaporDevam1:

End Sub

Sub Rapor2_2GelenTema()
Dim SiraBul As Range, IlkSira As Long

Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSira = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor2_2Devam1
End If

'Giden tema
GelenTemaGlobal = ""
If ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
    Else
        GelenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
    End If
Else
    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value))
        Else
            GelenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
        End If
    End If
End If

Rapor2_2Devam1:

End Sub

Sub Rapor1VarlikEsas()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim ItemBul As Range

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 1 – Primary Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If



On Error Resume Next 'Operation içinde Report 1 – Primary Asset Report.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 1 – Primary Asset Report.xlsx"
'Report 1 – Primary Asset Report açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 1 – Primary Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 1 – Primary Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 1 – Primary Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
'            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(3).Range("AO" & i & ":AO" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AL" & IlkSira & ":AL" & SonSira).Value
'            WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AO" & IlkSira & ":AO" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira).Value
            'ÖR yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "ÖR/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            
            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi(DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira))
            End If
        End If
    End If
Next ctl

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR1 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(3).Range("AR" & IlkSira & ":AR" & SonSira))
        End If
    End If
Next ctl

ToplamName = "Total (Qty)"
DevirName = "Carried-Over Balance from the Previous Day (Qty)"
GunSonuName = "End-of-Day Balance (Qty)"
WsVarliklar.Range("C" & SayEsasVarlik) = ToplamName
WsVarliklar.Range("C" & SayEsasVarlik + 2) = DevirName
WsVarliklar.Range("C" & SayEsasVarlik + 3) = GunSonuName
WsVarliklar.Range("C" & SayEsasVarlik & ":H" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":H" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":H" & SayEsasVarlik + 3).Font.Bold = True

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":H" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).Merge

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).HorizontalAlignment = xlLeft

If SayEsasVarlik - 1 >= 7 Then
    WsVarliklar.Range("G" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 7), WsVarliklar.Cells(SayEsasVarlik - 1, 7))) 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 8), WsVarliklar.Cells(SayEsasVarlik - 1, 8))) 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = WsVarliklar.Range("G" & SayEsasVarlik) + Devir - WsVarliklar.Range("H" & SayEsasVarlik) 'Gün sonu toplamı 'Gün sonu toplamı
Else
    WsVarliklar.Range("G" & SayEsasVarlik) = 0 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = 0 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = 0 + Devir - 0 'Gün sonu toplamı
End If

'Kontrol edilmiştir metni.
WsVarliklar.Range("C" & SayEsasVarlik + 5) = "The above records have been reviewed by us."
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).HorizontalAlignment = xlCenter 'xlLeft
'Anahtar sahipleri metni.
WsVarliklar.Range("C" & SayEsasVarlik + 7) = "Relevant Officers"
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).HorizontalAlignment = xlCenter 'xlLeft



'imzalar
WsVarliklar.Range("C" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza1.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza1Atla
End If
WsVarliklar.Range("C" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza1.Value <> "" Then
    WsVarliklar.Range("C" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza1Atla:

WsVarliklar.Range("E" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza2.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza2Atla
End If
WsVarliklar.Range("E" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza2.Value <> "" Then
    WsVarliklar.Range("E" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza2Atla:

WsVarliklar.Range("G" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza3.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza3Atla
End If
WsVarliklar.Range("G" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza3.Value <> "" Then
    WsVarliklar.Range("G" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza3Atla:


'İmza satırlarını ayarla
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter


'Kenarlıklar.
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":H" & SayEsasVarlik + 3)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub RaporVarlikEsas()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim ItemBul As Range

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 2.1 – Primary Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Report 2.1 – Primary Asset Report.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 2.1 – Primary Asset Report.xlsx"
'Report 2.1 – Primary Asset Report açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 2.1 – Primary Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 2.1 – Primary Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 2.1 – Primary Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
            'R yaz
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) <> "" Then
                    WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            
            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
            End If
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
        End If
    End If
Next ctl

ToplamName = "Total (Qty)"
DevirName = "Carried-Over Balance from the Previous Day (Qty)"
GunSonuName = "End-of-Day Balance (Qty)"
WsVarliklar.Range("C" & SayEsasVarlik) = ToplamName
WsVarliklar.Range("C" & SayEsasVarlik + 2) = DevirName
WsVarliklar.Range("C" & SayEsasVarlik + 3) = GunSonuName
WsVarliklar.Range("C" & SayEsasVarlik & ":H" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":H" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":H" & SayEsasVarlik + 3).Font.Bold = True

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":H" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).Merge

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).HorizontalAlignment = xlLeft

If SayEsasVarlik - 1 >= 7 Then
    WsVarliklar.Range("G" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 7), WsVarliklar.Cells(SayEsasVarlik - 1, 7))) 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 8), WsVarliklar.Cells(SayEsasVarlik - 1, 8))) 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = WsVarliklar.Range("G" & SayEsasVarlik) + Devir - WsVarliklar.Range("H" & SayEsasVarlik) 'Gün sonu toplamı 'Gün sonu toplamı
Else
    WsVarliklar.Range("G" & SayEsasVarlik) = 0 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = 0 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = 0 + Devir - 0 'Gün sonu toplamı
End If

'Kontrol edilmiştir metni.
WsVarliklar.Range("C" & SayEsasVarlik + 5) = "The above records have been reviewed by us."
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).HorizontalAlignment = xlCenter 'xlLeft
'Anahtar sahipleri metni.
WsVarliklar.Range("C" & SayEsasVarlik + 7) = "Relevant Officers"
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).HorizontalAlignment = xlCenter 'xlLeft



'imzalar
WsVarliklar.Range("C" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza1.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza1Atla
End If
WsVarliklar.Range("C" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza1.Value <> "" Then
    WsVarliklar.Range("C" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza1Atla:

WsVarliklar.Range("E" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza2.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza2Atla
End If
WsVarliklar.Range("E" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza2.Value <> "" Then
    WsVarliklar.Range("E" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza2Atla:

WsVarliklar.Range("G" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza3.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza3Atla
End If
WsVarliklar.Range("G" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza3.Value <> "" Then
    WsVarliklar.Range("G" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza3Atla:


'İmza satırlarını ayarla
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter


'Kenarlıklar.
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":H" & SayEsasVarlik + 3)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub Rapor2_2VarlikEsas()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim ItemBul As Range

Dim NomDeger As Variant, NomSay As Integer, RaporNoStr As String
Dim IlkSiraVarlik As Long, SonSiraVarlik As Long

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 2.2 – Primary Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If


On Error Resume Next 'Operation içinde Report 2.2 – Primary Asset Report.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 2.2 – Primary Asset Report.xlsx"
'Report 2.2 – Primary Asset Report açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 2.2 – Primary Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 2.2 – Primary Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 2.2 – Primary Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (GİRİŞ, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If

                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                Else
                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i

                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
    '                WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik

                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger

                        'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i

                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                Else
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value

                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger

                            'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                End If
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" Then
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklar.Range("D" & i) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (ÇIKIŞLAR, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                    End If
                    WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Merge
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    End If
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                Else
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    'Öğe Değeri string'ten long'a dönüştür ve yaz.
                    NomSay = 0
                    NomDeger = 0
                    For i = IlkSira To SonSira
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & i & ":AW" & i).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                        NomSay = NomSay + 1
                    Next i
                    
                    WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("K" & IlkSira & ":K" & SonSira).Value
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AT" & IlkSira & ":AT" & SonSira).Value
                    'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AW" & IlkSira & ":AW" & SonSira).Value
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    End If
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira).Value
                    SayEsasVarlikDongu = SayEsasVarlikDongu + (SonSira - IlkSira) + 1
    
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
                    End If

                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlik + (SonSira - IlkSira)).Value = "*"
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                SayEsasVarlikDongu = SayEsasVarlik
                j = IlkSira - 1
                WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                
                NomDeger = 0
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    j = j + 1
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 8) <> "Technique A" Then
                        WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                        WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
                        
                        'Öğe Değeri string'ten long'a dönüştür ve yaz.
                        NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        NomDeger = CLng(NomDeger)
                        WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                        
                        'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                        If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                            WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        End If
                        WsVarliklar.Range("H" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                        
                        'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                        If ctl.BackColor = &H80000003 Then
                            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & j))
                        End If
                        
                        SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                    End If
                Next i
                
                WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                SayEsasVarlikDongu = SayEsasVarlik
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") <> 0 Then
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1) 'Package A
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("F" & SayEsasVarlik).Value = "-"
                    
                    If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                            WsVarliklar.Range("G" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                        Else
                            WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                        End If
                    End If
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value <> "" Then
                        WsVarliklar.Range("H" & SayEsasVarlik).Value = ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    Else
                        WsVarliklar.Range("E" & SayEsasVarlik).Value = "Incomplete Data Entry!"
                    End If
                    WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlikDongu - 1).Merge
                    WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlikDongu - 1).Merge
                    'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                    If ctl.BackColor = &H80000003 Then
                        Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                    End If
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                Else
                    j = IlkSira - 1
                    WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
                    
                    NomDeger = 0
                    For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                        j = j + 1
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                            WsVarliklar.Range("D" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("K" & j).Value
                            WsVarliklar.Range("E" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AT" & j).Value
    
                            'Öğe Değeri string'ten long'a dönüştür ve yaz.
                            NomDeger = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            NomDeger = CLng(NomDeger)
                            WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = NomDeger
                            
                            'WsVarliklar.Range("F" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AW" & j).Value
                            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                                WsVarliklar.Range("G" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
                            End If
                            WsVarliklar.Range("H" & SayEsasVarlikDongu).Value = ThisWorkbook.Worksheets(4).Range("AZ" & j).Value
    
                            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
                            If ctl.BackColor = &H80000003 Then
                                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & j))
                            End If
                            
                            SayEsasVarlikDongu = SayEsasVarlikDongu + 1
                        End If
                    Next i
                    
                    WsVarliklar.Range("J" & SayEsasVarlik & ":J" & SayEsasVarlikDongu - 1).Value = "*"
                End If
            End If
            'R (B/xxx) yaz
            For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " Outgoing" Then
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & " (R/" & WsVarliklar.Range("D" & i) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                Else
                    For j = IlkSira To SonSira
                        If WsVarliklar.Range("D" & i) <> "" And ThisWorkbook.Worksheets(4).Range("K" & j) = WsVarliklar.Range("D" & i) Then
                            If Left(ThisWorkbook.Worksheets(4).Range("BK" & j).Value, 11) = "Technique A" Then
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i) & " (B/" & ThisWorkbook.Worksheets(4).Range("BM" & IlkSira) & ")"
                            Else
                                WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                            End If
                        End If
                    Next j
                End If
            Next i
            'Sıra ve kayıt noları dikeyde birleştir.
            If WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package A" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package B" Or _
                WsVarliklar.Range("E" & SayEsasVarlik).Value = "Package C" Then
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            Else
                WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlikDongu - 1).Merge
                For i = SayEsasVarlik To SayEsasVarlikDongu - 1 'SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("J" & i) <> "" Then
                        If i - 1 >= SayEsasVarlik Then
                            WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                        End If
                    End If
                Next i
            End If
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlikDongu  'SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl



'GoTo DuzenlemeyiAtla
'_____________________________________________________________BAŞLANGIÇ (Direkt kaldırsan yine de sorunsuz çalışır.)

Dim KontB As Integer, KontR As Integer, SayEsasVarlikx As Long
Dim SayEsasVarlikTakip As Long, a As Integer, b As Integer
Dim NumRaporNo As Variant

SayEsasVarlikTakip = 0
KontB = 0
KontR = 0
SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'General Primary Asset Report düzeltme faktörü (Technique A ve diğerlerinin sıra numarasını birleştir.)
If SayEsasVarlikx - 1 >= 7 Then
'    'RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox "2: " & StrRaporNo
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir1
                    End If
                Next j
DonguBitir1:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontB = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontB = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon1
                    ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon1
                    End If
                Next x
xDonguSon1:
                WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklar.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon2
                            ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon2
                            End If
                        Next x
xDonguSon2:
                        WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                        WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip
    
    'RAPOR NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'Tespit etme bölümü (1)
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo

                StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
                'MsgBox StrRaporNo
        
                For j = 1 To Len(StrRaporNo)
                    If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                        'MsgBox Mid(StrRaporNo, j, 1)
                        StrRaporNox = Left(StrRaporNo, j)
                    Else
                        GoTo DonguBitir2
                    End If
                Next j
DonguBitir2:
                StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
                'MsgBox StrRaporNo

                If KontR = 0 Then
                    SayEsasVarlikTakip = SayEsasVarlik + 1
                    KontR = 1
                End If

                IlkAltSatirNo = i
                SonAltSatirNo = i
                For x = i + 1 To SayEsasVarlik
                    If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                        SonAltSatirNo = x
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon3
                    ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                        GoTo xDonguSon3
                    End If
                Next x
xDonguSon3:
                WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                'Aratma bölümü (2)
                SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                For j = 7 To SayEsasVarlik 'Takip - 1
                    If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 Then
                    'If Left(WsVarliklar.Range("D" & j), 2 + Len(StrRaporNo)) = "R/" & StrRaporNo Then
                        IlkAltSatirNo = j
                        SonAltSatirNo = j
                        For x = j + 1 To SayEsasVarlik 'Takip
                            If WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) <> "" Then
                                SonAltSatirNo = x
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) <> "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklar.Range("D" & x) <> "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon4
                            ElseIf WsVarliklar.Range("D" & x) = "" And WsVarliklar.Range("J" & x) = "" Then
                                GoTo xDonguSon4
                            End If
                        Next x
xDonguSon4:

                        'WsVarliklar.Range("L" & j).Value = NumRaporNo

                        WsVarliklar.Range("C" & SayEsasVarlikTakip & ":H" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = WsVarliklar.Range("C" & IlkAltSatirNo & ":H" & SonAltSatirNo).Value
                        WsVarliklar.Range("I" & SayEsasVarlikTakip & ":I" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = StrRaporNo
                        WsVarliklar.Range("K" & SayEsasVarlikTakip & ":K" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = "*"
                        'WsVarliklar.Range("L" & SayEsasVarlikTakip & ":L" & SayEsasVarlikTakip + SonAltSatirNo - IlkAltSatirNo).Value = NumRaporNo
                        'WsVarliklar.Range("J" & IlkAltSatirNo & ":J" & SonAltSatirNo).Value = WsVarliklar.Range("D" & IlkAltSatirNo).Value '"Sil"
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).UnMerge
                        WsVarliklar.Range("C" & IlkAltSatirNo & ":J" & SonAltSatirNo).ClearContents
                        SayEsasVarlikTakip = SayEsasVarlikTakip + (SonAltSatirNo - IlkAltSatirNo + 1)
                    End If
                Next j
            End If
        End If
    Next i

    SayEsasVarlik = SayEsasVarlikTakip

    'Sıra no.lar
    SiraNoEsasVarlik = 0
    'MsgBox SayEsasVarlik
    'WsVarliklar.Range("C" & 7 & ":C" & SayEsasVarlik).ClearContents
    For i = 7 To SayEsasVarlik
        If WsVarliklar.Range("C" & i) <> "" And WsVarliklar.Range("I" & i) = "" Then
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
            WsVarliklar.Range("C" & i) = SiraNoEsasVarlik
        End If
        If WsVarliklar.Range("I" & i) <> "" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=WsVarliklar.Range("I" & i), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=WsVarliklar.Range("I" & i), SearchDirection:=xlPrevious, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'MsgBox IlkSira & " : " & SonSira
            If IlkSira = SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklar.Range("C" & i) = SiraNoEsasVarlik
                i = SonSira
            ElseIf IlkSira <> SonSira Then
                SiraNoEsasVarlik = SiraNoEsasVarlik + 1
                WsVarliklar.Range("C" & IlkSira) = SiraNoEsasVarlik
                WsVarliklar.Range("C" & IlkSira & ":C" & SonSira).Merge
                i = SonSira
            End If

        End If
    Next i
End If

SatirSilTekrar:
SayEsasVarlik = SayEsasVarlik - 1
If SayEsasVarlik < 7 Then
    GoTo SatirSilTekrarAtla
End If
For i = 7 To SayEsasVarlik
    If WsVarliklar.Range("K" & i) = "" Then
        WsVarliklar.Rows(i).EntireRow.Delete
        GoTo SatirSilTekrar
    End If
Next i
SatirSilTekrarAtla:

If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
Else
    SayEsasVarlik = SayEsasVarlik + 1
End If

'_____________________________________________________________BİTİŞ



'____________________İkinci düzenleme bölümü, BAŞLANGIÇ (Bu bölüm sadece XXXMud'den gelen raporları/öğelerin çıkışı içindir.)

SayEsasVarlikTakip = SayEsasVarlik
SayEsasVarlikx = SayEsasVarlik

'Raporların rapro no.'ları ile onun alt kırılımını sayısal bir değere dönüştür. Bu özellik raporları kendi içinde sıralamada kullanılacak.
If SayEsasVarlikx - 1 >= 7 Then
    'RAPOR ve RAPOR2_2 NO ESAS ALINIYOR
    For i = 7 To SayEsasVarlik - 1
        'R
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                'StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklar.Range("D" & i)
            End If
        End If
        'B
        If WsVarliklar.Range("D" & i) <> "" And Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
                'MsgBox "1: " & StrRaporNo
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo
        '_________________________
        For a = 1 To 50
            StrRaporNo = Replace(StrRaporNo, " ", "") 'Rapor no içinde varsa boşlukları kaldır
        Next a
        NumRaporNo = Replace(StrRaporNo, "-", "0000") '1-1'in 100001 gerçek rapor no ile çakışma ihtimali 0'a yakın.

        '_________________________

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir21
            End If
        Next j
DonguBitir21:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik1
        End If

        For j = IlkSiraVarlik To SonSiraVarlik
            If InStr(WsVarliklar.Range("D" & j).Value, "R/" & StrRaporNo) <> 0 And WsVarliklar.Range("L" & j).Value = "" Then
                WsVarliklar.Range("L" & j).Value = NumRaporNo
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklar.Range("D" & a + 1) = "" And WsVarliklar.Range("K" & a + 1) <> "" And WsVarliklar.Range("L" & a + 1).Value = "" Then
                            WsVarliklar.Range("L" & a + 1).Value = NumRaporNo
                        Else
                            GoTo aDonguSonu2
                        End If
                    End If
                Next a
            End If
        Next j
aDonguSonu2:
DonguSonuIlkSiraVarlik1:
    Next i
End If


'XXXMud'den gelen ve varlıkdaki raporlar. SIRALAMA ve BİRLEŞTİRME
For i = 7 To SayEsasVarlik - 1
    If WsVarliklar.Range("M" & i) = "" Then 'And WsVarliklar.Range("H" & i) <> "" Then
        If Left(WsVarliklar.Range("D" & i).Value, 1) = "R" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "B" Then
                StrRaporNo = Left(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") - 2) 'R/raporno şeklinde
                'MsgBox StrRaporNo
            ElseIf Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) <> "B" Then
                StrRaporNo = WsVarliklar.Range("D" & i)
            End If
        End If
        If Left(WsVarliklar.Range("D" & i).Value, 1) = "B" Then
            If Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, 1) = "R" Then
                StrRaporNo = Mid(WsVarliklar.Range("D" & i), InStr(WsVarliklar.Range("D" & i), "(") + 1, Len(WsVarliklar.Range("D" & i)) - InStr(WsVarliklar.Range("D" & i), "(") - 1) 'R/raporno şeklinde
            End If
        End If

        StrRaporNo = Mid(StrRaporNo, 3, Len(StrRaporNo)) 'Sadece rapor no
        'MsgBox StrRaporNo

        For j = 1 To Len(StrRaporNo)
            If IsNumeric(Mid(StrRaporNo, j, 1)) = True Then
                'MsgBox Mid(StrRaporNo, j, 1)
                StrRaporNox = Left(StrRaporNo, j)
            Else
                GoTo DonguBitir22
            End If
        Next j
DonguBitir22:
        StrRaporNo = StrRaporNox 'Alt rapor no kırılımı kaldırıldı
        'MsgBox StrRaporNo
        '_______________


        Set IlkSiraBul = Nothing
        Set SonSiraBul = Nothing
        IlkSiraVarlik = 0
        SonSiraVarlik = 0
        Set IlkSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraVarlik = IlkSiraBul.Row
        End If
        Set SonSiraBul = WsVarliklar.Range("I6:I100000").Find(What:=StrRaporNo, SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not SonSiraBul Is Nothing Then
            SonSiraVarlik = SonSiraBul.Row
        End If
        If IlkSiraVarlik = 0 Then
            GoTo DonguSonuIlkSiraVarlik
        End If

        WsVarliklar.Range("N" & IlkSiraVarlik) = IlkSiraVarlik
        WsVarliklar.Range("O" & SonSiraVarlik) = SonSiraVarlik
        WsVarliklar.Range("M" & IlkSiraVarlik & ":M" & SonSiraVarlik).Value = "*" 'Aynı satırlarda tekrar işlem yapmasın diye
        'Sorting Z to A by NumRaporNo
        WsVarliklar.Range("D" & IlkSiraVarlik & ":L" & SonSiraVarlik).UnMerge
        WsVarliklar.Range("D" & IlkSiraVarlik & ":L" & SonSiraVarlik).Sort key1:=Range("L" & IlkSiraVarlik & ":L" & SonSiraVarlik), order1:=xlAscending, Header:=xlNo

        'Rapor no.'ları ve Package A/Package B/Package C satırlarını birleştir.
        For j = IlkSiraVarlik To SonSiraVarlik
            If WsVarliklar.Range("L" & j).Value <> "" Then
                If WsVarliklar.Range("L" & j).Value = WsVarliklar.Range("L" & j + 1).Value Then
                    WsVarliklar.Range("D" & j & ":D" & j + 1).Merge
                End If
            End If
            If WsVarliklar.Range("E" & j).Value = "Package A" Or WsVarliklar.Range("E" & j).Value = "Package B" Or WsVarliklar.Range("E" & j).Value = "Package C" Then
                For a = j To SonSiraVarlik
                    If a + 1 <= SonSiraVarlik Then
                        If WsVarliklar.Range("E" & a + 1).Value = "" Then
                            WsVarliklar.Range("E" & a & ":E" & a + 1).Merge
                            WsVarliklar.Range("F" & a & ":F" & a + 1).Merge
                            WsVarliklar.Range("G" & a & ":G" & a + 1).Merge
                            WsVarliklar.Range("H" & a & ":H" & a + 1).Merge
                        Else
                            GoTo aDonguSonu
                        End If
                    End If
                Next a
            End If
aDonguSonu:
        Next j
    End If
DonguSonuIlkSiraVarlik:
Next i

'Artıkları temizle
If SayEsasVarlik < 7 Then
    SayEsasVarlik = 7
End If
WsVarliklar.Range("I" & 7 & ":O" & SayEsasVarlik).ClearContents


'____________________İkinci düzenleme bölümü, BİTİŞ


DuzenlemeyiAtla:

'GoTo Son

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR2_2 (DEVİRLER, KURUM ve XXXMud)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            If Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "T" Then 'Tümü
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                Else
                    Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & IlkSira & ":AZ" & SonSira))
                End If
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "R" Then 'Technique A olmayanlar
                For i = IlkSira To SonSira
                    If Left(ThisWorkbook.Worksheets(4).Range("BK" & i).Value, 8) <> "Technique A" Then
                        Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & i))
                    End If
                Next i
            ElseIf Mid(ctl.List(0), InStr(ctl.List(0), "<") + 1, 1) = "L" Then 'Technique A
                If Mid(ctl.List(0), InStr(ctl.List(0), ">") + 1, 8) = " (Incoming)" And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'Varlik paket
                    Devir = Devir + ThisWorkbook.Worksheets(4).Cells(IlkSira, 202).Value 'Package A adedi
                Else
                    For i = IlkSira To SonSira
                        If Left(ThisWorkbook.Worksheets(4).Range("BK" & i).Value, 11) = "Technique A" Then
                            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(4).Range("AZ" & i))
                        End If
                    Next i
                End If
            End If
        End If
    End If
Next ctl

ToplamName = "Total (Qty)"
DevirName = "Carried-Over Balance from the Previous Day (Qty)"
GunSonuName = "End-of-Day Balance (Qty)"
WsVarliklar.Range("C" & SayEsasVarlik) = ToplamName
WsVarliklar.Range("C" & SayEsasVarlik + 2) = DevirName
WsVarliklar.Range("C" & SayEsasVarlik + 3) = GunSonuName
WsVarliklar.Range("C" & SayEsasVarlik & ":H" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":H" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":H" & SayEsasVarlik + 3).Font.Bold = True

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":H" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).Merge

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).HorizontalAlignment = xlLeft

If SayEsasVarlik - 1 >= 7 Then
    WsVarliklar.Range("G" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 7), WsVarliklar.Cells(SayEsasVarlik - 1, 7))) 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 8), WsVarliklar.Cells(SayEsasVarlik - 1, 8))) 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = WsVarliklar.Range("G" & SayEsasVarlik) + Devir - WsVarliklar.Range("H" & SayEsasVarlik) 'Gün sonu toplamı 'Gün sonu toplamı
Else
    WsVarliklar.Range("G" & SayEsasVarlik) = 0 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = 0 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = 0 + Devir - 0 'Gün sonu toplamı
End If

'Kontrol edilmiştir metni.
WsVarliklar.Range("C" & SayEsasVarlik + 5) = "The above records have been reviewed by us."
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).HorizontalAlignment = xlCenter 'xlLeft
'Anahtar sahipleri metni.
WsVarliklar.Range("C" & SayEsasVarlik + 7) = "Relevant Officers"
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).HorizontalAlignment = xlCenter 'xlLeft



'imzalar
WsVarliklar.Range("C" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza1.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza1Atla
End If
WsVarliklar.Range("C" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza1.Value <> "" Then
    WsVarliklar.Range("C" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza1Atla:

WsVarliklar.Range("E" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza2.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza2Atla
End If
WsVarliklar.Range("E" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza2.Value <> "" Then
    WsVarliklar.Range("E" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza2Atla:

WsVarliklar.Range("G" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza3.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza3Atla
End If
WsVarliklar.Range("G" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza3.Value <> "" Then
    WsVarliklar.Range("G" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza3Atla:


'İmza satırlarını ayarla
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter


'Kenarlıklar.
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":H" & SayEsasVarlik + 3)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub

Sub Rapor3VarlikEsas()
Dim AutoPath, DestOperasyon, VarlikKlasoru, EsasVarlik, DestOpUserFolderName, DestOpUserFolder As String
Dim OperasyonEsasVarlik As String, OpenControl As String, KontrolFile As String
Dim ContSay As Long, SayEsasVarlik As Long, SiraNoEsasVarlik As Long, i As Long, j As Long, SayEsasVarlikDongu As Long
Dim fso As Object, WsVarliklar As Object
'Dim Ctl As MSForms.Control
Dim IlkSiraBul As Range, SonSiraBul As Range, Kenarlar As Range
Dim IlkSira As Long, SonSira As Long
Dim ToplamName As String, DevirName As String, GunSonuName As String
Dim Devir As Long ', GunSonu As Long
Dim StrRaporNo As String, StrRaporNox As String, x As Long, SonAltSatirNo As Long, IlkAltSatirNo As Long
Dim ItemBul As Range

Dim NomDeger As Variant, NomSay As Integer

'Application.DisplayAlerts = False 'Kodlar tamamlanınca iptal et

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
VarlikKlasoru = AutoPath & "\System Files\System Templates\Asset Templates\"
EsasVarlik = VarlikKlasoru & "Report 3 – Primary Asset Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"


'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Registry Reports klasör adını kontrol et.
If Not Dir(VarlikKlasoru, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & VarlikKlasoru & ". The folder named 'Asset Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(EsasVarlik, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & EsasVarlik & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

    


On Error Resume Next 'Operation içinde Report 3 – Primary Asset Report.xlsx dosyası yoksa oluşacak hata için
OperasyonEsasVarlik = DestOpUserFolder & "Report 3 – Primary Asset Report.xlsx"
'Report 3 – Primary Asset Report açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(OperasyonEsasVarlik)
If OpenControl = True Then
    Workbooks("Report 3 – Primary Asset Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (EsasVarlik), DestOpUserFolder & "Report 3 – Primary Asset Report" & ".xlsx", True
'open the file
Workbooks.Open (OperasyonEsasVarlik)

Set WsVarliklar = Workbooks("Report 3 – Primary Asset Report.xlsx").Worksheets(1)
WsVarliklar.Unprotect Password:="123"

On Error GoTo 0
'Varlik tarihi
WsVarliklar.Range("C4") = core_asset_manager_UI.VarlikTarihiText.Value

SayEsasVarlik = 7
SiraNoEsasVarlik = 1
Devir = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR3_2 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
            End If
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
        'RAPOR3_1 (GİRİŞ)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
            End If
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
            WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If
            
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR3_2 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("Z" & IlkSira & ":Z" & SonSira).Value
            End If
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("AZ" & IlkSira & ":AZ" & SonSira).Value

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("BC" & i & ":BC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BC" & IlkSira & ":BC" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira).Value
            
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 28).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If
            
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1

            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira))
            End If
        End If
        'RAPOR3_1 (ÇIKIŞLAR)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Aktarımı başlat
            WsVarliklar.Range("C" & SayEsasVarlik) = SiraNoEsasVarlik
            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("M" & IlkSira & ":M" & SonSira).Value
            Else
                WsVarliklar.Range("D" & SayEsasVarlik & ":D" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("CT" & IlkSira & ":CT" & SonSira).Value
            End If
            WsVarliklar.Range("E" & SayEsasVarlik & ":E" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("DZ" & IlkSira & ":DZ" & SonSira).Value

            'Öğe Değeri string'ten long'a dönüştür ve yaz.
            NomSay = 0
            NomDeger = 0
            For i = IlkSira To SonSira
                NomDeger = ThisWorkbook.Worksheets(5).Range("EC" & i & ":EC" & i).Value
                NomDeger = CLng(NomDeger)
                WsVarliklar.Range("F" & SayEsasVarlik + NomSay & ":F" & SayEsasVarlik + NomSay).Value = NomDeger
                NomSay = NomSay + 1
            Next i
            
            'WsVarliklar.Range("F" & SayEsasVarlik & ":F" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EC" & IlkSira & ":EC" & SonSira).Value
            If ctl.BackColor = &H80000000 Then '(GİRİŞİN AYNI GÜN ÇIKIŞI) 'Ctl.BackColor = &H80000003
                WsVarliklar.Range("G" & SayEsasVarlik & ":G" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value
            End If
            WsVarliklar.Range("H" & SayEsasVarlik & ":H" & SayEsasVarlik + (SonSira - IlkSira)).Value = ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira).Value

            If ThisWorkbook.Worksheets(5).Cells(IlkSira, 100).Value = "Type A" Then
                'R yaz
                For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                    If WsVarliklar.Range("D" & i) <> "" Then
                        WsVarliklar.Range("D" & i) = "R/" & WsVarliklar.Range("D" & i)
                    End If
                Next i
            End If
            
            'Sıra ve kayıt noları dikeyde birleştir.
            WsVarliklar.Range("C" & SayEsasVarlik & ":C" & SayEsasVarlik + (SonSira - IlkSira)).Merge
            For i = SayEsasVarlik To SayEsasVarlik + (SonSira - IlkSira)
                If WsVarliklar.Range("D" & i) = "" And WsVarliklar.Range("E" & i) <> "" Then
                    If i - 1 >= SayEsasVarlik Then
                        WsVarliklar.Range("D" & i & ":D" & i - 1).Merge
                    End If
                End If
            Next i
            'Satırları takip et.
            SayEsasVarlik = SayEsasVarlik + (SonSira - IlkSira) + 1
            SiraNoEsasVarlik = SiraNoEsasVarlik + 1

            'Devreden bakiyeden çıkışların da önceki günün devrine ilavesi (DEVİR)
            If ctl.BackColor = &H80000003 Then
                Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira))
            End If
        End If
    End If
Next ctl


For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        'RAPOR3_2 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("BF" & IlkSira & ":BF" & SonSira))
        End If
        'RAPOR3_1 (DEVİRLER)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            'Devirleri hesapla
            Devir = Devir + Application.Sum(ThisWorkbook.Worksheets(5).Range("EF" & IlkSira & ":EF" & SonSira))
        End If
    End If
Next ctl

ToplamName = "Total (Qty)"
DevirName = "Carried-Over Balance from the Previous Day (Qty)"
GunSonuName = "End-of-Day Balance (Qty)"
WsVarliklar.Range("C" & SayEsasVarlik) = ToplamName
WsVarliklar.Range("C" & SayEsasVarlik + 2) = DevirName
WsVarliklar.Range("C" & SayEsasVarlik + 3) = GunSonuName
WsVarliklar.Range("C" & SayEsasVarlik & ":H" & SayEsasVarlik).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":H" & SayEsasVarlik + 2).Font.Bold = True
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":H" & SayEsasVarlik + 3).Font.Bold = True

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 1 & ":H" & SayEsasVarlik + 1).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).Merge

WsVarliklar.Range("C" & SayEsasVarlik & ":F" & SayEsasVarlik).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 2 & ":G" & SayEsasVarlik + 2).HorizontalAlignment = xlLeft
WsVarliklar.Range("C" & SayEsasVarlik + 3 & ":G" & SayEsasVarlik + 3).HorizontalAlignment = xlLeft

If SayEsasVarlik - 1 >= 7 Then
    WsVarliklar.Range("G" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 7), WsVarliklar.Cells(SayEsasVarlik - 1, 7))) 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = Application.Sum(WsVarliklar.Range(Cells(7, 8), WsVarliklar.Cells(SayEsasVarlik - 1, 8))) 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = WsVarliklar.Range("G" & SayEsasVarlik) + Devir - WsVarliklar.Range("H" & SayEsasVarlik) 'Gün sonu toplamı 'Gün sonu toplamı
Else
    WsVarliklar.Range("G" & SayEsasVarlik) = 0 'Giriş toplamı
    WsVarliklar.Range("H" & SayEsasVarlik) = 0 'Çıkış toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 2) = Devir 'Devir toplamı
    WsVarliklar.Range("H" & SayEsasVarlik + 3) = 0 + Devir - 0 'Gün sonu toplamı
End If

'Kontrol edilmiştir metni.
WsVarliklar.Range("C" & SayEsasVarlik + 5) = "The above records have been reviewed by us."
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 5 & ":H" & SayEsasVarlik + 5).HorizontalAlignment = xlCenter 'xlLeft
'Anahtar sahipleri metni.
WsVarliklar.Range("C" & SayEsasVarlik + 7) = "Relevant Officers"
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 7 & ":H" & SayEsasVarlik + 7).HorizontalAlignment = xlCenter 'xlLeft



'imzalar
WsVarliklar.Range("C" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza1.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza1Atla
End If
WsVarliklar.Range("C" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza1.Value <> "" Then
    WsVarliklar.Range("C" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza1Atla:

WsVarliklar.Range("E" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza2.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza2Atla
End If
WsVarliklar.Range("E" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza2.Value <> "" Then
    WsVarliklar.Range("E" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza2Atla:

WsVarliklar.Range("G" & SayEsasVarlik + 11) = core_asset_manager_UI.VarlikImza3.Value
Set ItemBul = ThisWorkbook.Worksheets(2).Range("DY6:DY1000").Find(What:=core_asset_manager_UI.VarlikImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo VarlikImza3Atla
End If
WsVarliklar.Range("G" & SayEsasVarlik + 12) = ThisWorkbook.Worksheets(2).Range("DZ" & ItemBul.Row)
If core_asset_manager_UI.VarlikImza3.Value <> "" Then
    WsVarliklar.Range("G" & SayEsasVarlik + 10) = "Signature"
End If
VarlikImza3Atla:


'İmza satırlarını ayarla
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 10 & ":D" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 11 & ":D" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("C" & SayEsasVarlik + 12 & ":D" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 10 & ":F" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 11 & ":F" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("E" & SayEsasVarlik + 12 & ":F" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter

WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Font.Color = RGB(191, 191, 191)
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 10 & ":H" & SayEsasVarlik + 10).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 11 & ":H" & SayEsasVarlik + 11).HorizontalAlignment = xlCenter
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).Merge
WsVarliklar.Range("G" & SayEsasVarlik + 12 & ":H" & SayEsasVarlik + 12).HorizontalAlignment = xlCenter


'Kenarlıklar.
Set Kenarlar = WsVarliklar.Range("C" & 7 & ":H" & SayEsasVarlik + 3)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With


Son:

ThisWorkbook.Activate

'Application.DisplayAlerts = True 'Kodlar tamamlanınca iptal et

End Sub







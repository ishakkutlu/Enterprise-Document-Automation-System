VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_report1_entry_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   14460
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   21195
   OleObjectBlob   =   "core_report1_entry_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_report1_entry_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Abort As Boolean, KilitIptal As Boolean
Public OpenWordTakip As Boolean

Dim Tutanak1Kont As Integer, Rapor1Kont As Integer, Tutanak2Kont As Integer, UstYaziKont As Integer, MaxiR As Integer, Maxi As Integer
Dim TumKont As Integer

Dim WsFarkGiris As Worksheet, SayA As Long, SayD As Long, SayG As Long, SayJ As Long, SayM As Long, SayFarkGiris As Integer
Dim IlkSiraFarkGirisRapor1 As Long, SonSiraFarkGirisRapor1 As Long, WsFarkGirisRapor1 As Worksheet
Dim SayFarkGirisRapor1 As Long, Maxi1 As Integer

Dim StrRaporUnvan1 As String, StrRaporSicil1 As String, StrRaporUnvan2 As String, StrRaporSicil2 As String
Dim Threshold As Long

'Dim dHeight As Double

Sub Rapor1FormunuResetle()
Dim i As Integer
Dim ctl As MSForms.Control


ThisWorkbook.Activate

Il.Value = ""
Ilce.Value = ""
IlGiden.Value = ""
IlceGiden.Value = ""
BelgeTarihiText.Value = ""
BelgeNoText.Value = ""
TemaTipi.Value = ""
TemaNoText.Value = ""
OtomatikOption.Value = False
ManuelOption.Value = True
GelenMuhatapTemasi.Value = ""
GonderenBirim.Value = ""
GelisTarihiText.Value = ""
GelenPaketTipi.Value = ""
GelisSekli.Value = ""
Tutanak1TarihiText.Value = ""
Tutanak1Sonucu.Value = ""
GelenBelgeSayfa.Value = ""
DosyaNoText.Value = ""
DLEvetOption.Value = True
DLHayirOption.Value = False
DokumListesi.Value = ""
Tutanak1Imza1.Value = ""
Tutanak1Imza2.Value = ""
RaporImza1.Value = ""
RaporImza2.Value = ""
Tutanak2Imza1.Value = ""
Tutanak2Imza2.Value = ""
UstYaziImza1.Value = ""
UstYaziImza2.Value = ""
IlgiYaziFotokopisi.Value = ""

'Rapor1
'If Rapor1Frame.Visible = True Then
    Sonuc.Value = ""
    Rapor1No.Value = ""
    Rapor1TarihiText.Value = ""
'End If

OgeTuru.Value = ""
OgeDegeri.Value = ""
Adet.Value = ""
OgeIdNo.Value = ""
Aciklama.Value = ""

For i = 1 To 19
    Call EkleOge_Click
Next i
For i = 1 To 19
    Controls("OgeTuru" & i).Value = ""
    Controls("OgeDegeri" & i).Value = ""
    Controls("Adet" & i).Value = ""
    Controls("OgeIdNo" & i).Value = ""
    Controls("Aciklama" & i).Value = ""
    'If Rapor1Frame.Visible = True Then
        'Rapor1 için
        Controls("Sonuc" & i).Value = ""
        Controls("Rapor1No" & i).Value = ""
    'End If
Next i
For i = 1 To 19
    Call KaldirOge_Click
Next i

'Tutanak2
'If Tutanak2Frame.Visible = True Then
    Tutanak2TarihiText.Value = ""
    GidenMuhatapTemasi.Value = ""
    GonderilenBirim.Value = ""
    GidenPaketTipi.Value = ""
    GidenPaketAdedi.Value = ""
'End If

'Üst yazı
'If UstYaziFrame.Visible = True Then
    UstYaziTarihiText.Value = ""
    UstYaziNoText.Value = ""
'End If

'Taslak Renklerini resetle
For Each ctl In core_report1_entry_UI.Controls
    If TypeName(ctl) = "ComboBox" Then
        ctl.BackColor = RGB(255, 255, 255)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ctl) = "TextBox" Then
        ctl.BackColor = RGB(255, 255, 255)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
ComboGetir.BackColor = RGB(225, 235, 245)
ComboGetir.ForeColor = RGB(30, 30, 30)


'ComboGetir.SetFocus

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

End Sub

Sub Rapor1NoListClear()
Dim i As Integer

Rapor1No.Clear
For i = 1 To 19
    Controls("Rapor1No" & i).Clear
Next i

End Sub

Private Sub GonderenEkleKaldirLabel_Click()
support_contact_subunits_UI.Show
'support_contact_subunits_UI.Show vbModeless
End Sub

Private Sub GonderilenEkleKaldirLabel_Click()
support_contact_subunits_UI.Show
'support_contact_subunits_UI.Show vbModeless
End Sub

Private Sub LblDosyaNoGetir_Click()
Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


ThisWorkbook.Activate

Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\"
TaslakFile = "Report 1 Cover Letter.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & AutoPath & "\System Files\" & _
           ". The folder named 'System Files' may have been renamed or deleted.", _
           vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & DestTaslak & TaslakFile & _
           ". Folder or file names in this directory may have been changed.", _
           vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

'Oluşturulacak dosyayı aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
DosyaNoText.Value = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)


Son:

'ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub LblDuzeltme_Click()

'GetirLabelDuzeltme_Click

Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range
Dim Fark As Long
Dim i As Long, OgeFrame As Integer
Dim ctl As MSForms.Control, Resetle As Integer

'Columns("CE:CF").EntireColumn.Hidden = False

'Application.EnableEvents = False

'Application.ScreenUpdating = False

ThisWorkbook.Activate

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"
ThisWorkbook.Worksheets(8).Unprotect Password:="123"


'Rapor1 formunu resetle
Call UstYaziGirisi_Click
Call Rapor1FormunuResetle


If ComboGetir.Value = "" Then
    LblDuzeltme.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
    GoTo Son
End If

'Veri tabanını kontrol et
Say = Range("CE100000").End(xlUp).Row
If Say < 7 Or ComboGetir.Value = "" Then
    GoTo Son
End If

Set IlkSiraBul = Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkSiraBul Is Nothing Then
    IlkSira = IlkSiraBul.Row
Else
    MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not SonSiraBul Is Nothing Then
    SonSira = SonSiraBul.Row
Else
    MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Verileri sayfadan rapor1 formuna aktar.
'Tutanak1 bölümü
'Cells(IlkSira, 5)'Sıra numarası
Il.Value = Cells(IlkSira, 17).Value
Ilce.Value = Cells(IlkSira, 18).Value
BelgeTarihiText.Value = Cells(IlkSira, 20).Value
BelgeNoText.Value = Cells(IlkSira, 21).Value
TemaTipi.Value = Cells(IlkSira, 22).Value
TemaNoText.Value = Cells(IlkSira, 23).Value
If Cells(IlkSira, 24).Value = "Otomatik" Then
    OtomatikOption.Value = True
ElseIf Cells(IlkSira, 24).Value = "Manuel" Then
    ManuelOption.Value = True
Else
    OtomatikOption.Value = False
    ManuelOption.Value = False
End If
GelenMuhatapTemasi.Value = Cells(IlkSira, 25).Value
If Cells(IlkSira, 26).Value <> "" Then
    GonderenBirim.Value = Cells(IlkSira, 26).Value
Else
    GonderenBirim.Value = "Incoming Contact Theme"
End If
GelisTarihiText.Value = Cells(IlkSira, 28).Value
GelenPaketTipi.Value = Cells(IlkSira, 29).Value
GelisSekli.Value = Cells(IlkSira, 30).Value
Tutanak1TarihiText.Value = Cells(IlkSira, 31).Value
Tutanak1Sonucu.Value = Cells(IlkSira, 32).Value

'________________

'Fark girişleri için varsa esas kayıtları geçici kayıtlara aktar
If Tutanak1Sonucu.Value = "d. Discrepancy Detected" Then
    
    ThisWorkbook.Worksheets(7).Rows("3:30").EntireRow.Delete 'eski kaydı sil; düzenleden farklı bir sıra numarası çağrılırsa, önceki geçici kayıtlar silinsin.
                                                             'Varsa yenisi geçici kayıtlar sayfasına yüklensin.
    
    Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
    Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
    Else
        GoTo FarkGirisDuzeltmeAtla
    End If
    If Not SonSiraBul Is Nothing Then
        SonSiraFarkGirisRapor1 = SonSiraBul.Row
    Else
        GoTo FarkGirisDuzeltmeAtla
    End If
    
    Set WsFarkGiris = ThisWorkbook.Worksheets(7)
    
    WsFarkGiris.Range("A3:M" & (SonSiraFarkGirisRapor1 - IlkSiraFarkGirisRapor1 + 3)).Value = WsFarkGirisRapor1.Range("A" & IlkSiraFarkGirisRapor1 & ":M" & SonSiraFarkGirisRapor1).Value

FarkGirisDuzeltmeAtla:
Else
    ThisWorkbook.Worksheets(7).Rows("3:30").EntireRow.Delete
End If

'________________

GelenBelgeSayfa.Value = Cells(IlkSira, 33).Value
DosyaNoText.Value = Cells(IlkSira, 34).Value

If Cells(IlkSira, 35).Value = "Yes" Then
    DLEvetOption.Value = True
ElseIf Cells(IlkSira, 35).Value = "No" Then
    DLHayirOption.Value = True
Else
    DLEvetOption.Value = False
    DLHayirOption.Value = False
End If
DokumListesi.Value = Cells(IlkSira, 36).Value

Tutanak1Imza1.Value = Cells(IlkSira, 104).Value
Tutanak1Imza2.Value = Cells(IlkSira, 107).Value
RaporImza1.Value = Cells(IlkSira, 110).Value
RaporImza2.Value = Cells(IlkSira, 113).Value
Tutanak2Imza1.Value = Cells(IlkSira, 116).Value
Tutanak2Imza2.Value = Cells(IlkSira, 119).Value
UstYaziImza1.Value = Cells(IlkSira, 122).Value
UstYaziImza2.Value = Cells(IlkSira, 125).Value

OgeTuru.Value = Cells(IlkSira, 38).Value
OgeDegeri.Value = Cells(IlkSira, 41).Value
Adet.Value = Cells(IlkSira, 44).Value
OgeIdNo.Value = Cells(IlkSira, 47).Value
Aciklama.Value = Cells(IlkSira, 50).Value

Call Tutanak1Girisi_Click

'Rapor1
If Cells(IlkSira, 54).Value <> "" Or Cells(IlkSira, 59).Value <> "" Or Cells(IlkSira, 60).Value <> "" Then
    'Rapor1Frame.Visible = True
    Call RaporlamaGirisiPro
    Sonuc.Value = Cells(IlkSira, 54).Value
    'Rapor1No.Clear
    'Call Rapor1NoListClear
    Rapor1No.Value = Cells(IlkSira, 59).Value
    Rapor1TarihiText.Value = Cells(IlkSira, 60).Value
End If

Fark = SonSira - IlkSira + 1
If Fark > 1 And Fark < 21 Then
    For OgeFrame = 1 To Fark - 1
        'Controls("OgeTuruFrame" & OgeFrame).Visible = True
        Call EkleOge_Click
    Next OgeFrame
    For OgeFrame = 1 To Fark - 1
        Controls("OgeTuru" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 38).Value
        Controls("OgeDegeri" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 41).Value
        Controls("Adet" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 44).Value
        Controls("OgeIdNo" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 47).Value
        Controls("Aciklama" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 50).Value
        'Rapor1
        If Cells(IlkSira + OgeFrame, 54).Value <> "" Or Cells(IlkSira + OgeFrame, 59).Value <> "" Then
            'Rapor1Frame.Visible = True
            Call RaporlamaGirisiPro
            'Rapor1No.Clear
            Controls("Sonuc" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 54).Value
            'Controls("Rapor1No" & OgeFrame).Clear
            'Call Rapor1NoListClear
            Controls("Rapor1No" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 59).Value
        End If
    Next OgeFrame
End If

'Tutanak2
If Cells(IlkSira, 63).Value <> "" Or Cells(IlkSira, 64).Value <> "" Or Cells(IlkSira, 65).Value <> "" Or Cells(IlkSira, 67).Value <> "" Or Cells(IlkSira, 68).Value <> "" Then
    'Tutanak2Frame.Visible = True
    Call Tutanak2Girisi_Click
    'Rapor1No.Clear
    'Call Rapor1NoListClear
    Tutanak2TarihiText.Value = Cells(IlkSira, 63).Value
    GidenMuhatapTemasi.Value = Cells(IlkSira, 64).Value

    IlGiden.Value = Cells(IlkSira, 69).Value
    IlceGiden.Value = Cells(IlkSira, 70).Value
    
    If Cells(IlkSira, 65).Value <> "" Then
        GonderilenBirim.Value = Cells(IlkSira, 65).Value
    Else
        GonderilenBirim.Value = "Outgoing Contact Theme"
    End If
    GidenPaketTipi.Value = Cells(IlkSira, 67).Value
    GidenPaketAdedi.Value = Cells(IlkSira, 68).Value
End If

'Üst yazı
If Cells(IlkSira, 75).Value <> "" Or Cells(IlkSira, 76).Value <> "" Then
    'UstYaziFrame.Visible = True
    'Call UstYaziGirisi_Click
    Call UstYaziGirisi_Click
    'Rapor1No.Clear
    'Call Rapor1NoListClear
    UstYaziTarihiText.Value = Cells(IlkSira, 75).Value
    UstYaziNoText.Value = Cells(IlkSira, 76).Value
    IlgiYaziFotokopisi.Value = Cells(IlkSira, 74).Value
End If

LblDuzeltme.BackColor = RGB(180, 210, 240)
LblDuzeltme.ForeColor = RGB(30, 30, 30)



'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat


Son:

'Columns("CE:CF").EntireColumn.Hidden = True

ThisWorkbook.Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(7).Protect Password:="123"
ThisWorkbook.Worksheets(8).Protect Password:="123"
ThisWorkbook.Protect "123"

'Application.ScreenUpdating = True

'Application.EnableEvents = True


End Sub


Private Sub LblIl_Click()
MsgBox "Select the province where the request letter originated from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired province on your keyboard to quickly locate it. For example, pressing the letter A once may select Province 1, pressing it again may switch to Province 2." & vbNewLine & vbNewLine & _
"To update the details of a province or district, or to add a new one, click the ± icon and follow the instructions in the window that appears." & vbNewLine & vbNewLine & _
"The selected province is used for automatic TEMA code generation and appears in Statement 1, the Dispatch List, Report 1, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlce_Click()
MsgBox "Select the district where the request letter originated from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired district on your keyboard to locate it quickly. For example, to select 'District 2' under Province 1, press the letter A once to highlight District 1, and again to reach District 2." & vbNewLine & vbNewLine & _
"To update the information of a district or to add a new one, click the ± icon next to the Province label and follow the instructions in the window that appears." & vbNewLine & vbNewLine & _
"The selected district is used in automatic TEMA code generation and appears in Statement 1, the Dispatch List, Report 1, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblBelgeTarihi_Click()
MsgBox "Click the calendar icon on the right and select the date of the request letter from the calendar." & vbNewLine & vbNewLine & _
"The selected date in the Document Date field is used for automatic TEMA code generation and appears in Statement 1, the Dispatch List, Report 1, Statement 2, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblBelgeNo_Click()
MsgBox "Enter the request letter number into the box on the right." & vbNewLine & vbNewLine & _
"For example, if the number is 1234567, you may enter it as either '1234567' or '2018/1234567'. If entered as '1234567', the year required for the TEMA code will be taken from the Document Date field. If entered as '2018/1234567', the year will be extracted from the part before the slash." & vbNewLine & vbNewLine & _
"The value entered in the Document Number field is used in automatic TEMA code generation, as well as in Statement 1, the Dispatch List, Report 1, Statement 2, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTemaTipi_Click()
MsgBox "Select the unit that issued the TEMA code (or the requesting authority) from the dropdown list on the right. For example, if the TEMA code for the transaction in question was issued by Organization B, select Organization B. If it was issued by Organization C, select Organization C." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired item on your keyboard to quickly navigate to it." & vbNewLine & vbNewLine & _
"The selection made in the TEMA Type field is used for automatic TEMA code generation. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTemaNo_Click()
MsgBox "If all required fields—Province, District, Document Date, Document Number, and TEMA Type—are filled in, selecting the 'Automatic' option will generate the TEMA code automatically. While the Automatic option is selected, the TEMA code cannot be edited manually. If the Manual option is selected, the user can edit the TEMA code." & vbNewLine & vbNewLine & _
"When the Automatic option is active, and changes are made to any of the fields that form the TEMA code (Province, District, Document Date, Document Number, or TEMA Type), these changes will only be reflected in the TEMA code if the Manual option is selected first and then switched back to Automatic." & vbNewLine & vbNewLine & _
"The value in the TEMA Number field (i.e., the TEMA code) is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGelenMuhatapTemasi_Click()
MsgBox "Select the recipient mentioned in the request letter from the dropdown list on the right (e.g., a Directorate, Decision Board, or a theme such as Provincial Directorate B or District Directorate B)." & vbNewLine & vbNewLine & _
"For example, if the request letter was sent by the XXX Governorship Provincial Directorate B – XXX Unit Directorate, select 'Provincial Directorate B'. If it was sent by the X.X. X1 Process Monitoring Directorate – XXX Office, select 'X.X. X1 Process Monitoring Directorate'. The XXX Unit Directorate or XXX Office should be selected separately in the Sender Unit field." & vbNewLine & vbNewLine & _
"If the desired Directorate or Decision Board is not listed in the dropdown, click the ± icon on the right and follow the instructions in the window to add it to the system." & vbNewLine & vbNewLine & _
"The selection made in the Incoming Contact Theme field is used in Statement 1, the Dispatch List, Report 1, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGonderenBirim_Click()
MsgBox "Select the unit that sent the request letter from the dropdown list on the right." & vbNewLine & vbNewLine & _
"For example, if the request letter was sent by the XXX Governorship Provincial Directorate B – XXX Unit Directorate, select the 'XXX Unit Directorate'. If it was sent by the X.X. X1 Process Monitoring Directorate – XXX Office, select the 'XXX Office'. If the request letter does not clearly indicate the unit (e.g., whether it is a unit directorate or an office), select the Incoming Contact Theme instead." & vbNewLine & vbNewLine & _
"If the desired unit name is not available in the dropdown list, click the ± icon on the right and follow the instructions in the window to add it to the system." & vbNewLine & vbNewLine & _
"The selection made in the Sender Unit field is used in Statement 1, the Dispatch List, Report 1, and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGelisTarihi_Click()
MsgBox "Click the calendar icon on the right and select the date the request letter was received by your department." & vbNewLine & vbNewLine & _
"The selected date in the Received Date field is used in the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGelenPaketTipi_Click()
MsgBox "Select the package type of the incoming item from the dropdown list on the right." & vbNewLine & vbNewLine & _
"The selection made in the Incoming Package Type field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGelisSekli_Click()
MsgBox "Select the delivery method of the incoming item from the dropdown list on the right." & vbNewLine & vbNewLine & _
"The selection made in the Delivery Method field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak1Tarihi_Click()
MsgBox "Click the calendar icon on the right and select the Statement 1 date of the incoming item from the calendar." & vbNewLine & vbNewLine & _
"The selected date in the Statement 1 Date field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak1Sonucu_Click()
MsgBox "Select the outcome of Statement 1 from the dropdown list on the right." & vbNewLine & vbNewLine & _
"The selection made in the Statement 1 Outcome field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDokumListesi_Click()
MsgBox "If you want to include a Dispatch List in the document you are creating, select 'Yes' from the options on the right. If not, select 'No'. When 'Yes' is selected, you must also specify the number of Dispatch List pages using the dropdown list. When 'No' is selected, the dropdown becomes inactive." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first digit of the desired number on your keyboard to quickly locate it." & vbNewLine & vbNewLine & _
"The selection made in the Dispatch List field is used in Statement 1 and as an attachment to the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak1Imza1_Click()
MsgBox "Select the person to be displayed in the first signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the name of the relevant person is not available in the list, click the ± icon on the right and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to navigate quickly." & vbNewLine & vbNewLine & _
"The selection made in the Signature field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub
Private Sub LblTutanak1Imza2_Click()
MsgBox "Select the person to be displayed in the second signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the name of the relevant person is not available in the list, click the ± icon on the right and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to navigate quickly." & vbNewLine & vbNewLine & _
"The selection made in the Signature field is used in Statement 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGelenBelgeSayfa_Click()
MsgBox "Select the number of pages included in the incoming item from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first digit of the desired number on your keyboard to quickly locate it." & vbNewLine & vbNewLine & _
"The selection made in the Incoming Document Page Count field is used in the preparation of the receipt document submitted to XXX for entry into the XX System. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDosyaNo_Click()
MsgBox "Enter the file code to be used for processing the incoming document in the XXS system in the box on the right, or retrieve the relevant file code by clicking the Get button.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblOgeTuruUst_Click()
MsgBox "Select the item type related to the subject under review from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired item on your keyboard to quickly navigate to it." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"If the desired item type is not listed, click the ± icon to the left of the Item Type label and follow the instructions in the window to define it in the system." & vbNewLine & vbNewLine & _
"The selection made in the Item Type field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeDegeriUst_Click()
MsgBox "Select the item value related to the subject under review from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first digit of the desired value on your keyboard to quickly navigate to it." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"If the desired item value is not listed, click the ± icon to the left of the Nominal Value label and follow the instructions in the window to define it in the system." & vbNewLine & vbNewLine & _
"The selection made in the Nominal Value field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAdetUst_Click()
MsgBox "Enter the quantity related to the item under review in the box below." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"The value entered in the Quantity field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeIdNoUst_Click()
MsgBox "Enter the Item ID number of the item under review into the box below." & vbNewLine & vbNewLine & _
"If you previously selected 'Yes' for the Dispatch List option, the Item ID field will automatically display 'Dispatch List' when an item type is selected or changed. If you prefer, you may overwrite this text with the actual item ID number." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"The value entered in the Item ID field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAciklamaUst_Click()
MsgBox "You may enter a description related to the item under review in the box below." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"The content entered in the Description field is used in Statement 1 and Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblSonucUst_Click()
MsgBox "Select the evaluation result related to the item under review from the dropdown list below." & vbNewLine & vbNewLine & _
"To enter multiple item types/values/quantities, click the + icon at the far right of this row. To remove item type/value/quantity rows, click the - icon at the same location." & vbNewLine & vbNewLine & _
"The selection made in the Result field is used in Report 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRapor1NoUst_Click()
MsgBox "Enter the Report 1 number into the dropdown list below." & vbNewLine & vbNewLine & _
"If a Report 1 number is assigned to each row, the system will generate a separate Report 1 for each row. If Report 1 numbers are assigned at intervals, the system will group the rows from top to bottom until the next Report 1 number and generate one combined Report 1 for that group." & vbNewLine & vbNewLine & _
"For example, assume 5 different item types are entered — the first 3 are valid, and the last 2 are invalid. If all 5 rows have a Report 1 number (e.g., 180-1, 180-2, 180-3, 180-4, 180-5), the system will generate 5 separate Report 1 documents." & vbNewLine & vbNewLine & _
"If only the 1st and 4th rows are assigned a Report 1 number, then rows 1, 2, and 3 (valid items) will be combined into one Report 1, and rows 4 and 5 (invalid items) into another — resulting in a total of 2 Report 1 documents." & vbNewLine & vbNewLine & _
"The value entered in the Report 1 Number field is used in Report 1 and in the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblRapor1Tarihi_Click()
MsgBox "Click the calendar icon on the right and select the Report 1 date from the calendar." & vbNewLine & vbNewLine & _
"The selected date in the Report 1 Date field is used in Report 1 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporImza1_Click()
MsgBox "Select the person to be displayed in the first signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the name of the relevant person is not available in the list, click the ± icon on the right and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to navigate quickly." & vbNewLine & vbNewLine & _
"The selection made in the Signature field is used in Report 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporImza2_Click()
MsgBox "Select the person to be displayed in the second signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the name of the relevant person is not available in the list, click the ± icon on the right and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to navigate quickly." & vbNewLine & vbNewLine & _
"The selection made in the Signature field is used in Report 1. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblTutanak2Tarihi_Click()
MsgBox "Click the calendar icon on the right and select the Statement 2 date from the calendar." & vbNewLine & vbNewLine & _
"The selected date in the Statement 2 Date field is used in Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGidenMuhatapTemasi_Click()
MsgBox "Select the recipient of the response letter from the dropdown list on the right (e.g., a Directorate, Decision Board, or a theme such as Provincial Directorate B or District Directorate B)." & vbNewLine & vbNewLine & _
"For example, if the response letter will be sent to the XXX Governorship Provincial Directorate B – XXX Unit Directorate, select 'Provincial Directorate B'. If it will be sent to the X.X. X1 Process Monitoring Directorate – XXX Office, select 'X.X. X1 Process Monitoring Directorate'. The XXX Unit Directorate or XXX Office should be selected separately in the Recipient Unit field." & vbNewLine & vbNewLine & _
"If the desired Directorate or Decision Board name is not listed, click the ± icon on the right and follow the instructions in the window to define it in the system." & vbNewLine & vbNewLine & _
"The selection made in the Outgoing Contact Theme field is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGonderilenBirim_Click()
MsgBox "Select the unit to which the response letter will be sent from the dropdown list on the right." & vbNewLine & vbNewLine & _
"For example, if the response letter will be sent to the XXX Governorship Provincial Directorate B – XXX Unit Directorate, select 'XXX Unit Directorate'. If it will be sent to the X.X. X1 Process Monitoring Directorate – XXX Office, select 'XXX Office'. If the response letter is to be sent directly to the recipient defined in the Outgoing Contact Theme (without specifying a unit such as XXX Unit Directorate or XXX Office), select 'Outgoing Contact Theme'." & vbNewLine & vbNewLine & _
"If the name of the desired unit is not listed, click the ± icon on the right and follow the instructions in the window to define it in the system." & vbNewLine & vbNewLine & _
"The selection made in the Recipient Unit field is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlGiden_Click()
MsgBox "Select the province to which the response letter will be sent from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired province on your keyboard to navigate quickly. For example, pressing 'A' once may select Province 1, and pressing it again may move to Province 2." & vbNewLine & vbNewLine & _
"To update the details of a province or district, or to add a new one, click the ± icon next to the Province label in the Statement 1 Entry section and follow the instructions in the window." & vbNewLine & vbNewLine & _
"The selection made in the Province field in this section is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlceGiden_Click()
MsgBox "Select the district to which the response letter will be sent from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the desired district on your keyboard to navigate quickly. For example, pressing 'A' once may select District 1, and pressing it again may move to District 2." & vbNewLine & vbNewLine & _
"To update the details of a district or to add a new one, click the ± icon next to the Province label in the Statement 1 Entry section and follow the instructions in the window." & vbNewLine & vbNewLine & _
"The selection made in the District field in this section is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGidenPaketTipi_Click()
MsgBox "Select the package and delivery type of the outgoing dispatch from the dropdown list on the right." & vbNewLine & vbNewLine & _
"The selection made in the Outgoing Package Type field is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGidenPaketAdedi_Click()
MsgBox "Select the number of outgoing packages from the dropdown list on the right." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the initial digit of your desired selection on the keyboard until it appears." & vbNewLine & vbNewLine & _
"The selection made in the Outgoing Package Quantity field is used in Statement 2 and the Response Letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak2Imza1_Click()
MsgBox "Select the person to be displayed in the first signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the desired name is not available in the list, click the ± symbol and follow the instructions in the window that appears to define the person in the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first letter of the person's name on the keyboard until the correct name appears." & vbNewLine & vbNewLine & _
"The selection made in the signature field is used in Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak2Imza2_Click()
MsgBox "Select the person to be displayed in the second signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the desired name is not available in the list, click the ± icon and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person's name on your keyboard to quickly navigate." & vbNewLine & vbNewLine & _
"The selection made in the signature field is used in Statement 2. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblUstYaziTarihi_Click()
MsgBox "Click the calendar icon on the right and select the cover letter (response letter) date from the calendar." & vbNewLine & vbNewLine & _
"The selected date in the Cover Letter Date field is used in the response letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub UstYaziNoLabel_Click()
MsgBox "Enter the cover letter (response letter) number issued by XXX in the box below." & vbNewLine & vbNewLine & _
"The selection made in the Cover Letter Number field is used in the response letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlgiYaziFotokopisi_Click()
MsgBox "If you want to send a photocopy of the referenced document mentioned in the response letter, specify the number of pages of the photocopy in the box on the right. If you do not intend to send the photocopy, leave the box empty." & vbNewLine & vbNewLine & _
"If the response letter will be sent to a different recipient than the one in the request letter, it is recommended not to leave the Referenced Document Photocopy Pages field empty." & vbNewLine & vbNewLine & _
"The selection made in the Referenced Document Photocopy Pages field is used in the Type 1 statement and the response letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziImza1_Click()
MsgBox "Select the person to be displayed in the first signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the desired name is not available in the list, click the ± icon and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to quickly navigate." & vbNewLine & vbNewLine & _
"The selection made in the signature field is used in the response letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziImza2_Click()
MsgBox "Select the person to be displayed in the second signature field from the dropdown list on the right." & vbNewLine & vbNewLine & _
"If the desired name is not available in the list, click the ± icon and follow the instructions in the window to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can also press the first letter of the person’s name on your keyboard to quickly navigate." & vbNewLine & vbNewLine & _
"The selection made in the signature field is used in the response letter. For more details, click the Help button in the top-right corner.", _
vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Sub ComboGetirReset()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(3).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If
'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(3).Range("E" & i).Value <> "" Then
        With ComboGetir
            .AddItem (ThisWorkbook.Worksheets(3).Range("E" & i).Value)
        End With
    End If
Next i
GetirBos:

End Sub

Private Sub LblSil_Click()
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range
Dim Fark As Long, Sifre As Variant
Dim i As Long, j As Long, SiraNoSakla As Long
Dim OncekiSiraNo As Long
Dim AutoPath As String, IslemGunlukleriKlasor As String, IslemGunlugu As String, OpenControl As String

Dim WsRapor As Worksheet, WsIslemGunlugu As Worksheet, BulIslemGunlugu As Range, Kenarlar As Range

Dim IslemGunluguIlkSiraBul As Range, IslemGunluguSonSiraBul As Range, IslemGunluguIlkSira As Long, IslemGunluguSonSira As Long



ThisWorkbook.Activate

Application.ScreenUpdating = False
Application.DisplayAlerts = False
'Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"
ThisWorkbook.Worksheets(8).Unprotect Password:="123"
ThisWorkbook.Worksheets(11).Unprotect Password:="123"

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"

If ComboGetir.Value = "" Then
    MsgBox "To delete a record from the system, please select the serial number of the record you wish to delete from the dropdown list located between the Edit and Draft buttons, then click the Edit button. After confirming the correct record is loaded, click the Delete button and follow the instructions in the prompt.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'Veri tabanını kontrol et
Say = ThisWorkbook.Worksheets(3).Range("CE100000").End(xlUp).Row
If Say < 7 Or ComboGetir.Value = "" Then
    GoTo Out
End If

Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkSiraBul Is Nothing Then
    IlkSira = IlkSiraBul.Row
Else
    MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
If Not SonSiraBul Is Nothing Then
    SonSira = SonSiraBul.Row
Else
    MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"
'Check Registry Reports folder existence.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & IslemGunlukleriKlasor & " is not accessible. The folder named Registry Reports may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & IslemGunlugu & " is not accessible. The names of the folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Hazırlık
SiraNoSakla = ThisWorkbook.Worksheets(3).Cells(IlkSira, 5).Value
OncekiSiraNo = ThisWorkbook.Worksheets(3).Cells(IlkSira, 5).Value - 1
Set WsRapor = ThisWorkbook.Worksheets(3)

Sifre = InputBox(Prompt:="To delete the operation with serial number " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 5).Value & " from the system, please enter the password value '123'.", Title:="Enterprise Document Automation System")
If Sifre = "123" Then

    'İŞLEM GÜNLÜĞÜ
    'İşlem günlüğü açıksa kaydet ve kapat.
    OpenControl = IsWorkBookOpen(IslemGunlugu)
    If OpenControl = True Then
        Workbooks("System Registry Report 1.xlsx").Save
        Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
    End If
    Workbooks.Open (IslemGunlugu)
    Set WsIslemGunlugu = Workbooks("System Registry Report 1.xlsx").Worksheets(1)
    WsIslemGunlugu.Unprotect Password:="123"
    WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = False
    
    'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
    Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
    Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
    SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row
    
    Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IslemGunluguIlkSiraBul Is Nothing Then
        IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
        If Not IslemGunluguSonSiraBul Is Nothing Then
            IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
        End If

        'kayıt def. verileri sil, satırları işaretle
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 19)).ClearContents
        WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2).Value = "Sil" 'ilk satırı silmek üzere işaretle
        WsIslemGunlugu.Cells(IslemGunluguSonSira, 3).Value = "Sil" 'son satırı silmek üzere işaretle
        
        'Dönem sıra no.ları güncelle
        i = IslemGunluguSonSira
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'silinecek verinin dönemi en alt satırda değilse stop koşulu
            i = i + 1
            If i > Say2IslemGunlugu Then 'silinecek verinin dönemi en alt satırda ise stop koşulu
                GoTo SilDonemSiraNo
            End If
            If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then 'silinen veriden sonraki verileri dönem sıra no.ları 1 azalır
                WsIslemGunlugu.Cells(i, 6).Value = WsIslemGunlugu.Cells(i, 6).Value - 1
            End If
        Loop
SilDonemSiraNo:

        'Genel sıra no.ları güncelle
        SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
        i = IslemGunluguSonSira
        Do Until i > SayGenel
            i = i + 1
            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value - 1
            End If
        Loop
        
        Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
        If IslemGunluguIlkSira > 8 Then 'And IslemGunluguIlkSira < Say2IslemGunlugu Then
            Set Kenarlar = WsIslemGunlugu.Range("D" & IslemGunluguIlkSira - 1 & ":S" & IslemGunluguIlkSira - 1)
            With Kenarlar.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
        End If

        'Silinecek dönemde yer alan boş satır aralığını kaldır
        Set BulIslemGunlugu = WsIslemGunlugu.Range("B:B").Find(What:="Sil", SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not BulIslemGunlugu Is Nothing Then
            ilkrowx = BulIslemGunlugu.Row
            Set BulIslemGunlugu = WsIslemGunlugu.Range("C:C").Find(What:="Sil", SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not BulIslemGunlugu Is Nothing Then
                sonrowx = BulIslemGunlugu.Row
            End If
            WsIslemGunlugu.Rows(ilkrowx & ":" & sonrowx).EntireRow.Delete
        End If
        
    Else
        'Nothing
    End If
    
    'İşlem günlüğünde aşağı git
    Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
    On Error Resume Next
    ActiveWindow.ScrollRow = Say2IslemGunlugu - 10
    On Error GoTo 0

    WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = True

    WsIslemGunlugu.Protect Password:="123"
    
    'İşlem günlüğü açıksa kaydet ve kapat.
    OpenControl = IsWorkBookOpen(IslemGunlugu)
    If OpenControl = True Then
        Workbooks("System Registry Report 1.xlsx").Save
        Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
    End If
    

    'MODÜL işlemleri
    Set WsRapor = ThisWorkbook.Worksheets(3)
    'On Error Resume Next
    'Sira numaralarını düzelt
    If Say > IlkSira Then
        For i = IlkSira + 1 To Say
            If WsRapor.Cells(i, 5).Value <> "" Then
                OncekiSiraNo = OncekiSiraNo + 1
                WsRapor.Cells(i, 5).Value = OncekiSiraNo
                WsRapor.Cells(i, 83).Value = OncekiSiraNo 'başlangıç
                
                For j = i To i + 1000
                    If WsRapor.Cells(j, 84).Value <> "" Then
                        WsRapor.Cells(j, 84).Value = OncekiSiraNo 'bitiş
                        GoTo DonguJSon
                    End If
                Next j
DonguJSon:
            End If
        Next i

        'Geçici fark kayıtlarını sil
        ThisWorkbook.Worksheets(7).Rows("3:30").EntireRow.Delete
        'Kalıcı fark kayıtlarını sil
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
        Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=SiraNoSakla, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=SiraNoSakla, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
        Else
            GoTo FarkGirisSilAtla2
        End If
        If Not SonSiraBul Is Nothing Then
            SonSiraFarkGirisRapor1 = SonSiraBul.Row
        Else
            GoTo FarkGirisSilAtla2
        End If
        WsFarkGirisRapor1.Rows(IlkSiraFarkGirisRapor1 & ":" & SonSiraFarkGirisRapor1).EntireRow.Delete

        'Fark girişlerindeki sıra no.lar, modül içindeki sıra no.lar ile senkronize olarak değiştirilmesi gerekir.
        'Fark girişlerindeki sıra no.lar (modülün aksine) ardışık olarak birbirine takip etmeyebilir.
        For i = SiraNoSakla To OncekiSiraNo
            Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=i + 1, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=i + 1, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
            Else
                GoTo FarkGirisNoDuzeltAtla
            End If
            If Not SonSiraBul Is Nothing Then
                SonSiraFarkGirisRapor1 = SonSiraBul.Row
            Else
                GoTo FarkGirisNoDuzeltAtla
            End If
            WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1, 16).Value = i
            WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1, 17).Value = i
FarkGirisNoDuzeltAtla:
        Next i
FarkGirisSilAtla2:

    ElseIf Say = IlkSira Then
        'MsgBox " Modül: Güncellenecek no yok!"
        
        'Geçici fark kayıtlarını sil
        ThisWorkbook.Worksheets(7).Rows("3:30").EntireRow.Delete
        'Kalıcı fark kayıtlarını sil
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
        Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=SiraNoSakla, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=SiraNoSakla, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
        Else
            GoTo FarkGirisSilAtla1
        End If
        If Not SonSiraBul Is Nothing Then
            SonSiraFarkGirisRapor1 = SonSiraBul.Row
        Else
            GoTo FarkGirisSilAtla1
        End If
        WsFarkGirisRapor1.Rows(IlkSiraFarkGirisRapor1 & ":" & SonSiraFarkGirisRapor1).EntireRow.Delete
FarkGirisSilAtla1:

    End If


    'Rapor no sayfasında silme işlemini yap
    
    '__________Rapor No Senkronizasyon 30.11.2021

    Set WsRaporNo = ThisWorkbook.Worksheets(11)

    Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not RnoIlkSiraBul Is Nothing Then
        RnoIlkSira = RnoIlkSiraBul.Row
        If Not RnoSonSiraBul Is Nothing Then
            RnoSonSira = RnoSonSiraBul.Row
        End If
        WsRaporNo.Rows(RnoIlkSira & ":" & RnoSonSira).EntireRow.Delete
    End If
    
    '__________Rapor No Senkronizasyon 30.11.2021
        
        
    'Modülde silme işlemini gerçekleştir.
    WsRapor.Rows(IlkSira & ":" & SonSira).EntireRow.Delete


    '_______
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Call UstYaziGirisi_Click
    ComboGetir.Value = ""
    Call Rapor1FormunuResetle
    Call ComboGetirReset
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    'Ok
    MsgBox "The record with serial number " & SiraNoSakla & " has been successfully deleted from the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"


ElseIf Sifre = vbCancel Then
    'MsgBox "Şifre iptal"
    GoTo Out
ElseIf Sifre <> "" And Sifre <> "123" Then
    MsgBox "Deletion failed for record with serial number " & SiraNoSakla & " due to incorrect password.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If 'Şifre koşulu sonu


Out:

'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 1.xlsx").Save
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
End If
    
            
ThisWorkbook.Activate

ThisWorkbook.Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(7).Protect Password:="123"
ThisWorkbook.Worksheets(8).Protect Password:="123"
ThisWorkbook.Worksheets(11).Protect Password:="123"
ThisWorkbook.Protect "123"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True
    
End Sub

Private Sub LblTaslak_Click()
'GetirLabelTaslak_Click
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range
Dim Fark As Long
Dim i As Long, OgeFrame As Integer
Dim ctl As MSForms.Control, Resetle As Integer

'Columns("CE:CF").EntireColumn.Hidden = False

ThisWorkbook.Activate

KilitIptal = True

Call LblDuzeltme_Click
ComboGetir.Value = ""
'Rapor1 no değerlerini sıfırla
Call Son20RaporNo
Rapor1No.Value = ""
For i = 1 To 19
    Controls("Rapor1No" & i).Value = ""
Next i

LblDuzeltme.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
LblDuzeltme.ForeColor = RGB(30, 30, 30)

'Taslak Renkler
For Each ctl In core_report1_entry_UI.Controls
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            ctl.BackColor = RGB(60, 100, 180)
            ctl.ForeColor = RGB(255, 255, 255)
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            ctl.BackColor = RGB(60, 100, 180)
            ctl.ForeColor = RGB(255, 255, 255)
        End If
    End If
Next ctl
ComboGetir.BackColor = RGB(225, 235, 245)
ComboGetir.ForeColor = RGB(30, 30, 30)

Son:

KilitIptal = False

'Columns("CE:CF").EntireColumn.Hidden = True


End Sub

Private Sub IlIlceEkleKaldirLabel_Click()
support_provinces_districts_UI.Show
'support_provinces_districts_UI.Show vbModeless
End Sub
Private Sub IlIlceEkleKaldirLabel2_Click()
support_provinces_districts_UI.Show
'support_provinces_districts_UI.Show vbModeless
End Sub
Private Sub MuhatapEkleKaldirLabelGelen_Click()
support_contact_themes_UI.Show
'support_contact_themes_UI.Show vbModeless
End Sub

Private Sub MuhatapEkleKaldirLabelGiden_Click()
support_contact_themes_UI.Show
'support_contact_themes_UI.Show vbModeless
End Sub

Private Sub OgeDegeriEkleKaldirLabel_Click()
support_item_values_UI.Show
'support_item_values_UI.Show vbModeless
End Sub

Private Sub OgeEkleKaldirLabel_Click()
support_item_types_UI.Show
'support_item_types_UI.Show vbModeless
End Sub

Private Sub Tutanak1Imza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub Tutanak1Imza2EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub

Private Sub RaporImza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub RaporImza2EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub

Private Sub Tutanak2Imza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub Tutanak2Imza2EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub

Private Sub UstYaziImza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub UstYaziImza2EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub

Private Sub FarkGirisi_Click()
support_discrepancy_entry_UI.Show
'support_discrepancy_entry_UI.Show vbModeless
End Sub


Sub OpenWordControl()
Dim ObjWordx As Object
Dim objDocx As Object

'MsgBox "OpenWordControl prosedürü başlıyor."

    On Error GoTo NoOpenDoc
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    OpenWordTakip = True
    GoTo NoOpenDocAtla
NoOpenDoc:
    OpenWordTakip = False
NoOpenDocAtla:
    If OpenWordTakip = True Then
        'MsgBox objWordx.ActiveDocument.Name
        If ObjWordx.ActiveDocument.name <> "" Then
            ObjWordx.Quit SaveChanges:=True
            'MsgBox "Dosya OpenWordControl methodu ile kapatıldı."
        End If
    Else
        'MsgBox "Açık word dokümanı yok."
    End If

Son:
Set ObjWordx = Nothing

End Sub

Private Sub Kapat_Click()
Unload Me
End Sub

Private Sub Yardim_Click()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Taslak
SourceTaslak = AutoPath & "\System Files\Help Documents\Report 1 Entry – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'System Files folder check
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & AutoPath & "\System Files\" & _
           ". The folder named 'System Files' may have been renamed or deleted.", _
           vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation folder check
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & DestOperasyon & _
           ". The folder named 'Operation' may have been renamed or deleted.", _
           vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'RmDir DestOpUserFolder 'Sistem kapanırken DestOpUserFolder klasörünü temizle EKLENECEK!
'_______________

'Check folder names
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & SourceTaslak & _
           ". Folder or file names within this directory may have been changed.", _
           vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'On Error Resume Next
ReNameTaslak = "Help Documents"
'________________________________________

'Close the all Word application
Call OpenWordControl

'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
Do While OpenKontrolName <> ""
    OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
    If OpenControl = True Then 'Açıksa
        On Error Resume Next
        Set objWord = GetObject(, "Word.Application")
        Set objWord = GetObject(, "Word.Application")
        Set objWord = GetObject(, "Word.Application")
        Set objWord = GetObject(, "Word.Application")
        Set objWord = GetObject(, "Word.Application")
        objWord.Quit SaveChanges:=True
        'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

    End If
    OpenKontrolName = Dir()
Loop

Set objWord = Nothing
Set objDoc = Nothing
'________________________________________

On Error Resume Next
'    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If


'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (SourceTaslak), DestOpUserFolder & ReNameTaslak & ".docm", True

'________________________________________

'Oluşturulacak dosyayı aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTaslak & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
objWord.Application.WindowState = wdWindowStateMaximize

'Set objDoc = GetObject(DestOpUserFolder & ReNameTaslak & ".docm")

Son:

Set objWord = Nothing
Set objDoc = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Sub KaydetOnKontroller()
Dim ctl As MSForms.Control

TumKont = 0
For Each ctl In core_report1_entry_UI.EsasFrame.Controls 'EsasFrame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
Next ctl
For Each ctl In core_report1_entry_UI.ScrollFrame.Controls 'ScrollFrame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
Next ctl
For Each ctl In core_report1_entry_UI.Rapor1Frame.Controls 'Rapor1Frame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
Next ctl
For Each ctl In core_report1_entry_UI.Tutanak2Frame.Controls 'Tutanak2Frame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
Next ctl
For Each ctl In core_report1_entry_UI.UstYaziFrame.Controls 'UstYaziFrame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKont = 1
        End If
    End If
Next ctl


End Sub
Sub KontrolProseduru()
Dim YeniIslem As Long ', SiraBul As Range, SiraKontrol As Range
Dim i As Long, j As Long, OgeFrame As Integer, Kont As Integer
Dim ctl As MSForms.Control
'Dim TumKont As Integer
'Dim Tutanak1Kont As Integer, Rapor1Kont As Integer, Tutanak2Kont As Integer, UstYaziKont As Integer
Dim Bilgi As Variant
Dim OgeTuruKont As Integer, OgeDegeriKont As Integer, AdetKont As Integer
Dim OgeIdNoKont As Integer, AciklamaKont As Integer, SonucKont As Integer ', MaxiR As Integer, Maxi As Integer
Dim OgeTuruKontSatir As Integer, OgeDegeriKontSatir As Integer, AdetKontSatir As Integer
Dim OgeIdNoKontSatir As Integer, AciklamaKontSatir As Integer, SonucKontSatir As Integer
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range, Fark As Long
Dim FarkSay As Integer, SiraNoSakla As Long, SiraSay As Long
Dim Rapor1NoKont As Integer, Rapor1NoBul As Range, Rapor1NoBulIlk As Range
Dim Rapor1NoBulTire As Range, Rapor1NoBulTireKont As Integer, Rapor1NoBulKont As Integer
Dim Rapor1NoBulTirePart As Range, Rapor1NoBulTireKontPart As Integer
Dim Kenarlar As Range, DokumKontSatir As Integer, UserName As String
Dim Rapor1NoKontAyni As Integer, Rapor1NoKontAltNoHata As Integer, Rapor1NoKontUstNoHata As Integer

Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim OpenControl As String, Say1IslemGunlugu As Long, Say2IslemGunlugu As Long
Dim GelenTema As String, Sene As String, Ay As String
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long
Dim ItemBul As Range
'Dim StrRaporUnvan1 As String, StrRaporSicil1 As String, StrRaporUnvan2 As String, StrRaporSicil2 As String
Dim RefSatir As Long, Rapor1TarihBul As Range


'Tutanak1 validations
Tutanak1Kont = 0
If Il.Value = "" Then
    Bilgi = MsgBox("Province is not specified. To save anyway, click " & _
                   """Yes""; to correct, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet1
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet1:
If InStr(GelenMuhatapTemasi.Value, "District") <> 0 Then
    If Ilce.Value = "" Then
        Bilgi = MsgBox("District is referenced in the Incoming Contact Theme, but district is not specified. To save anyway, click " & _
                       """Yes""; to correct, click ""No"".", _
                       vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet2
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
End If

YinedeKaydet2:
If BelgeTarihiText.Value = "" Then
    Bilgi = MsgBox("Document date is not specified. To save anyway, click " & _
                   """Yes""; to correct, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet3
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet3:
If BelgeNoText.Value = "" Then
    Bilgi = MsgBox("Document number is not entered. To save anyway, click " & _
                   """Yes""; to correct, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet4
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet4:
'If TemaTipi.Value = "" Then
'    Bilgi = MsgBox("Theme type is not specified. To save anyway, click " & _
'                   """Yes""; to correct, click ""No"".", _
'                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
'    If Bilgi = vbYes Then
'        Tutanak1Kont = 1
'        GoTo YinedeKaydet5
'    ElseIf Bilgi = vbNo Then
'        Tutanak1Kont = 2
'        GoTo Son
'    End If
'End If
'YinedeKaydet5:

If TemaNoText.Value = "" Then
    '
End If
If OtomatikOption.Value = False And ManuelOption.Value = False Then
    Bilgi = MsgBox("The mode of theme number creation (Automatic/Manual) has not been selected. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet6
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
    End If
End If

YinedeKaydet6:
If GelenMuhatapTemasi.Value = "" Then
    Bilgi = MsgBox("Incoming Contact Theme is not specified. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet7Ek1
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet7Ek1:
If GonderenBirim.Value = "" Then
    Bilgi = MsgBox("Sending subunit is not specified. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet7
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet7:
If GelisTarihiText.Value = "" Then
    Bilgi = MsgBox("Received date is not entered. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet8
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet8:
If GelenPaketTipi.Value = "" Then
    Bilgi = MsgBox("Incoming package type is not specified. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet9
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet9:
If GelisSekli.Value = "" Then
    Bilgi = MsgBox("Receiving method is not specified. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet10
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet10:
If Tutanak1TarihiText.Value = "" Then
    Bilgi = MsgBox("Statement 1 date is not specified. To save anyway, click " & _
                   """Yes""; to make corrections, click ""No"".", _
                   vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet11
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet11:

If GelisTarihiText.Value <> "" And BelgeTarihiText.Value <> "" Then
    If Year(GelisTarihiText.Value) < Year(BelgeTarihiText.Value) Then
        Bilgi = MsgBox("The arrival date is earlier than the incoming document date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek1A
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek1A:
    If (Year(GelisTarihiText.Value) = Year(BelgeTarihiText.Value) And Month(GelisTarihiText.Value) < Month(BelgeTarihiText.Value)) Then
        Bilgi = MsgBox("The arrival date is earlier than the incoming document date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek2A
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek2A:
    If (Year(GelisTarihiText.Value) = Year(BelgeTarihiText.Value) And Month(GelisTarihiText.Value) = Month(BelgeTarihiText.Value) And Day(GelisTarihiText.Value) < Day(BelgeTarihiText.Value)) Then
        Bilgi = MsgBox("The arrival date is earlier than the incoming document date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek3A
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek3A:
End If

If Tutanak1TarihiText.Value <> "" And GelisTarihiText.Value <> "" And BelgeTarihiText.Value <> "" Then
    If Year(Tutanak1TarihiText.Value) < Year(BelgeTarihiText.Value) Or _
        Year(Tutanak1TarihiText.Value) < Year(GelisTarihiText.Value) Then
        Bilgi = MsgBox("Statement 1 date is earlier than the incoming document date and/or received date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek1
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek1:
    If (Year(Tutanak1TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Tutanak1TarihiText.Value) < Month(BelgeTarihiText.Value)) Or _
        (Year(Tutanak1TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Tutanak1TarihiText.Value) < Month(GelisTarihiText.Value)) Then
        Bilgi = MsgBox("Statement 1 date is earlier than the incoming document date and/or received date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek2
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek2:
    If (Year(Tutanak1TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Tutanak1TarihiText.Value) = Month(BelgeTarihiText.Value) And Day(Tutanak1TarihiText.Value) < Day(BelgeTarihiText.Value)) Or _
        (Year(Tutanak1TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Tutanak1TarihiText.Value) = Month(GelisTarihiText.Value) And Day(Tutanak1TarihiText.Value) < Day(GelisTarihiText.Value)) Then
        Bilgi = MsgBox("Statement 1 date is earlier than the incoming document date and/or received date. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet11Ek3
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If
YinedeKaydet11Ek3:
End If

If Tutanak1Sonucu.Value = "" Then
    Bilgi = MsgBox("Statement 1 result is not specified. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet12
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet12:
If GelenBelgeSayfa.Value = "" Then
    Bilgi = MsgBox("Number of pages of the incoming document is not specified. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet12Ek1
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet12Ek1:
If DLEvetOption.Value = False And DLHayirOption.Value = False Then
    Bilgi = MsgBox("Dispatch List request status (Yes/No) has not been selected. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet13
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet13:
If DLEvetOption.Value = True And DokumListesi.Value = "" Then
    Bilgi = MsgBox("Dispatch List has been requested but the number of pages is not specified. Do you want to save anyway?", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet14
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet14:


If OgeTuru.Value = "" Then
    Bilgi = MsgBox("The item type is not specified. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet15
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet15:
If OgeDegeri.Value = "" Then
    Bilgi = MsgBox("The item value is not specified. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet16
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet16:
If Adet.Value = "" Then
    Bilgi = MsgBox("The quantity is not specified. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet17
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet17:
If OgeIdNo.Value = "" Then
    Bilgi = MsgBox("The item ID number is not specified. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet18
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet18:
If OgeIdNo.Value = "Dispatch List" And DLEvetOption.Value = False Then
    Bilgi = MsgBox("The item ID number indicates 'Dispatch List', but the Dispatch List option is not selected as 'Yes'. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet18Ek1
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If
YinedeKaydet18Ek1:
If Aciklama.Value = "" Then
    ' (Buraya ek kod gelebilir)
End If

'___________________________

'Farklı tutanak1 girişi kontrolü
SayFarkGiris = 1
If Tutanak1Sonucu.Value = "d. Discrepancy Detected" Then

    Set WsFarkGiris = ThisWorkbook.Worksheets(7)
    'Maksimum değerler.
    SayA = WsFarkGiris.Range("A100000").End(xlUp).Row
    SayD = WsFarkGiris.Range("D100000").End(xlUp).Row
    SayG = WsFarkGiris.Range("G100000").End(xlUp).Row
    SayJ = WsFarkGiris.Range("J100000").End(xlUp).Row
    SayM = WsFarkGiris.Range("M100000").End(xlUp).Row
    SayFarkGiris = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)
    
    If SayFarkGiris < 3 Then
        Bilgi = MsgBox("Although 'Discrepancy Detected' is selected in the Statement 1 result field, no data has been entered in the 'Discrepancy Entry' field. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak1Kont = 1
            GoTo YinedeKaydet12FarkGirisi
        ElseIf Bilgi = vbNo Then
            Tutanak1Kont = 2
            GoTo Son
        End If
    End If

End If
YinedeKaydet12FarkGirisi:

'___________________________


'Arada boş bırakılan satırların kontrolü; öğe türü, öğe değeri, adet, öğe ID no (ve açıklama)
Kont = 0
For OgeFrame = 1 To 19
    If Controls("OgeTuruFrame" & OgeFrame).Visible = True Then
        Kont = OgeFrame
    End If
Next OgeFrame
OgeTuruKont = 0
OgeDegeriKont = 0
AdetKont = 0
OgeIdNoKont = 0
AciklamaKont = 0
If Kont > 0 Then
    For OgeFrame = 1 To Kont
        If Controls("OgeTuru" & OgeFrame).Value <> "" Then
            OgeTuruKont = OgeFrame
        End If
        If Controls("OgeDegeri" & OgeFrame).Value <> "" Then
            OgeDegeriKont = OgeFrame
        End If
        If Controls("Adet" & OgeFrame).Value <> "" Then
            AdetKont = OgeFrame
        End If
        If Controls("OgeIdNo" & OgeFrame).Value <> "" Then
            OgeIdNoKont = OgeFrame
        End If
        If Controls("Aciklama" & OgeFrame).Value <> "" Then
            AciklamaKont = OgeFrame
        End If
    Next OgeFrame
End If
OgeTuruKontSatir = 0
OgeDegeriKontSatir = 0
AdetKontSatir = 0
OgeIdNoKontSatir = 0
AciklamaKontSatir = 0
DokumKontSatir = 0
Maxi = Application.Max(OgeTuruKont, OgeDegeriKont, AdetKont, OgeIdNoKont, AciklamaKont)
If Maxi > 0 Then
    For i = 1 To Maxi
        If Controls("OgeTuru" & i).Value = "" Then
            OgeTuruKontSatir = i
        End If
        If Controls("OgeDegeri" & i).Value = "" Then
            OgeDegeriKontSatir = i
        End If
        If Controls("Adet" & i).Value = "" Then
            AdetKontSatir = i
        End If
        If Controls("OgeIdNo" & i).Value = "" Then
            OgeIdNoKontSatir = i
        End If
        If Controls("OgeIdNo" & i).Value = "Dispatch List" And DLEvetOption.Value = False Then
            DokumKontSatir = i
        End If
'        If Controls("Aciklama" & i).Value = "" Then
'            AciklamaKontSatir = i
'        End If
    Next i
End If
'Yukarıdaki maxi değeri, (aşağıda bulunan kodlarda) verilerin rapor1 formundan
'sayfaya aktarılmasında kullanılıyor.
If OgeTuruKontSatir <> 0 And OgeDegeriKontSatir <> 0 And AdetKontSatir <> 0 And OgeIdNoKontSatir Then
    Bilgi = MsgBox("A skipped row has been detected. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet19
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
End If

YinedeKaydet19:
If OgeTuruKontSatir <> 0 Then
    Bilgi = MsgBox("Missing item type detected. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
ElseIf OgeDegeriKontSatir <> 0 Then
    Bilgi = MsgBox("Missing item value detected. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
ElseIf AdetKontSatir <> 0 Then
    Bilgi = MsgBox("Missing quantity detected. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
ElseIf OgeIdNoKontSatir <> 0 Then
    Bilgi = MsgBox("Missing item ID number detected. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
ElseIf DokumKontSatir <> 0 Then
    Bilgi = MsgBox("Although 'Dispatch List' is specified in the item ID number field, the 'Dispatch List' option is not marked as 'Yes'. Click 'Yes' to save anyway or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        Tutanak1Kont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        Tutanak1Kont = 2
        GoTo Son
    End If
ElseIf AciklamaKontSatir <> 0 Then
    '
End If

YinedeKaydet20:

'Report1 validations
'Rapor1Kont = 2
If Rapor1Frame.Visible = True Then
    Rapor1Kont = 0
    If Sonuc.Value = "" Then
        Bilgi = MsgBox("Result is missing. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet21
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet21:

    If Rapor1No.Value <> "" Then
        If InStr(Rapor1No.Value, "-") = 0 Then 'No dash in report number
            '
        Else 'Dash exists
            If Mid(Rapor1No.Value, InStr(Rapor1No.Value, "-") + 1, 1) <> 1 Then
                Bilgi = MsgBox("The sub-number for the first line of Report 1 does not start at 1 (e.g., it starts with 18-2 instead of 18-1). Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                If Bilgi = vbYes Then
                    Rapor1Kont = 1
                    GoTo YinedeKaydet21AltNo1Degil
                ElseIf Bilgi = vbNo Then
                    Rapor1Kont = 2
                    GoTo SonRapor
                End If
            End If
        End If
    End If
YinedeKaydet21AltNo1Degil:
    
    '__________Rapor No Senkronizasyon 30.11.2021
    
    RefSatir = 0
    Set WsRaporNo = ThisWorkbook.Worksheets(11)
    If Rapor1TarihiText.Value <> "" Then
        StrRaporTarihiGlobal = Right(CStr(Rapor1TarihiText.Value), 4)
    Else
        StrRaporTarihiGlobal = Right(CStr(Format(Date, "dd.mm.yyyy")), 4)
    End If
    Set Rapor1TarihBul = WsRaporNo.Range("B6:B100000").Find(What:=StrRaporTarihiGlobal, SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
    If Not Rapor1TarihBul Is Nothing Then
        RefSatir = Rapor1TarihBul.Row
    Else
        RefSatir = 7
    End If
    If RefSatir < 7 Then
        RefSatir = 7
    End If
    'MsgBox "RefSatir: " & RefSatir
    
    'MsgBox "IlkSiraGlobal: " & IlkSiraGlobal
    If ComboGetir.Value <> "" Then 'Düzenleme işlemi ise cari işlemin rapor no.su aramalara takılmasın; bu yüzden sağa al; sonra tekrar yerine koymayı unutma!
        Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not RnoIlkSiraBul Is Nothing Then
            RnoIlkSira = RnoIlkSiraBul.Row
            If Not RnoSonSiraBul Is Nothing Then
                RnoSonSira = RnoSonSiraBul.Row
            End If
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 7), WsRaporNo.Cells(RnoSonSira, 11)).Value = WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 1), WsRaporNo.Cells(RnoSonSira, 5)).Value
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 1), WsRaporNo.Cells(RnoSonSira, 5)).ClearContents
        End If
    End If
 
    'Rapor1 numarasının daha önce kullanılıp kullanılmadığını kontrol et.
    If Rapor1No.Value <> "" Then
        
        RaporTireTek = 0 '82-1 gibi tek değer girilemez
        
        If InStr(Rapor1No.Value, "-") = 0 Then 'rapor no girişinde tire yok
            
            'Tiresiz değerler içinde ara
            StrAramaGlobal = Rapor1No.Value
            Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
            Set Rapor1NoBulIlk = MyRngGlobal.Find(What:=StrAramaGlobal, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not Rapor1NoBulIlk Is Nothing Then
                Bilgi = MsgBox("It has been detected that the first Report 1 number has already been used. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                If Bilgi = vbYes Then
                    Rapor1Kont = 1
                    GoTo YinedeKaydet21Ek1
                ElseIf Bilgi = vbNo Then
                    Rapor1Kont = 2
                    GoTo DuzeltmeniYapDaGit1
                    'GoTo SonRapor
                End If
            End If
        
            'Tireli değerler içinde ara
            StrAramaGlobal = Rapor1No.Value & "-"
            Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
            Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                            SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
            If Not MyFinderGlobal Is Nothing Then
                IlkAdresGlobal = MyFinderGlobal.Address
                'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                    Bilgi = MsgBox("The first Report 1 number has been detected as previously used. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                    If Bilgi = vbYes Then
                        Rapor1Kont = 1
                        GoTo YinedeKaydet21Ek1
                    ElseIf Bilgi = vbNo Then
                        Rapor1Kont = 2
                        GoTo DuzeltmeniYapDaGit1
                        'GoTo SonRapor
                    End If
                End If
                'Sonraki satırlarda aramaya devam et
                Do
                    SonrakiAdresGlobal = MyFinderGlobal.Address
                    'MsgBox Replace(SonrakiAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                    Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
                    SonrakiAdresGlobal = MyFinderGlobal.Address
                    If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                        Bilgi = MsgBox("It has been detected that the first Report 1 number has already been used. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                        If Bilgi = vbYes Then
                            Rapor1Kont = 1
                            GoTo YinedeKaydet21Ek1
                        ElseIf Bilgi = vbNo Then
                            Rapor1Kont = 2
                            GoTo DuzeltmeniYapDaGit1
                            'GoTo SonRapor
                        End If
                    End If
                Loop While IlkAdresGlobal <> SonrakiAdresGlobal
            End If
        
        Else 'rapor no girişinde tire var
            RaporTireTek = 1
            
            'Tiresiz değerler içinde ara
            StrAramaGlobal = Left(Rapor1No.Value, InStr(Rapor1No.Value, "-") - 1)
            Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
            Set Rapor1NoBulIlk = MyRngGlobal.Find(What:=StrAramaGlobal, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not Rapor1NoBulIlk Is Nothing Then
                Bilgi = MsgBox("The first Report 1 number has been detected as previously used. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                If Bilgi = vbYes Then
                    Rapor1Kont = 1
                    GoTo YinedeKaydet21Ek1
                ElseIf Bilgi = vbNo Then
                    Rapor1Kont = 2
                    GoTo DuzeltmeniYapDaGit1
                    'GoTo SonRapor
                End If
            End If
        
            'Tireli değerler içinde ara
            StrAramaGlobal = Left(Rapor1No.Value, InStr(Rapor1No.Value, "-") - 1) & "-"
            Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
            Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                            SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
            If Not MyFinderGlobal Is Nothing Then
                IlkAdresGlobal = MyFinderGlobal.Address
                'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                    Bilgi = MsgBox("It has been detected that the first Report 1 number has been used before. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                    If Bilgi = vbYes Then
                        Rapor1Kont = 1
                        GoTo YinedeKaydet21Ek1
                    ElseIf Bilgi = vbNo Then
                        Rapor1Kont = 2
                        GoTo DuzeltmeniYapDaGit1
                        'GoTo SonRapor
                    End If
                End If
                'Sonraki satırlarda aramaya devam et
                Do
                    SonrakiAdresGlobal = MyFinderGlobal.Address
                    'MsgBox Replace(SonrakiAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                    Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
                    SonrakiAdresGlobal = MyFinderGlobal.Address
                    If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                        Bilgi = MsgBox("It has been detected that the first Report 1 number has been used previously. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                        If Bilgi = vbYes Then
                            Rapor1Kont = 1
                            GoTo YinedeKaydet21Ek1
                        ElseIf Bilgi = vbNo Then
                            Rapor1Kont = 2
                            GoTo DuzeltmeniYapDaGit1
                            'GoTo SonRapor
                        End If
                    End If
                Loop While IlkAdresGlobal <> SonrakiAdresGlobal
            End If

        End If
    End If

YinedeKaydet21Ek1:

    '__________Rapor No Senkronizasyon 30.11.2021


    'Arada boş bırakılan satırların kontrolü; öğe türü, öğe değeri, adet, öğe ID no, sonuç (ve açıklama)
    Kont = 0
    For OgeFrame = 1 To 19
        If Controls("OgeTuruFrame" & OgeFrame).Visible = True Then
            Kont = OgeFrame
        End If
    Next OgeFrame
    OgeTuruKont = 0
    OgeDegeriKont = 0
    AdetKont = 0
    OgeIdNoKont = 0
    AciklamaKont = 0
    SonucKont = 0
    Rapor1NoBulTireKont = 0
    Rapor1NoBulKont = 0
    Rapor1NoBulTireKontPart = 0
    Rapor1NoKont = 0
    If Kont > 0 Then
        For OgeFrame = 1 To Kont
            If Controls("OgeTuru" & OgeFrame).Value <> "" Then
                OgeTuruKont = OgeFrame
            End If
            If Controls("OgeDegeri" & OgeFrame).Value <> "" Then
                OgeDegeriKont = OgeFrame
            End If
            If Controls("Adet" & OgeFrame).Value <> "" Then
                AdetKont = OgeFrame
            End If
            If Controls("OgeIdNo" & OgeFrame).Value <> "" Then
                OgeIdNoKont = OgeFrame
            End If
            If Controls("Aciklama" & OgeFrame).Value <> "" Then
                AciklamaKont = OgeFrame
            End If
                If Controls("Sonuc" & OgeFrame).Value <> "" Then
                SonucKont = OgeFrame
            End If
        Next OgeFrame
    End If
    
    OgeTuruKontSatir = 0
    OgeDegeriKontSatir = 0
    AdetKontSatir = 0
    OgeIdNoKontSatir = 0
    AciklamaKontSatir = 0
    SonucKontSatir = 0
    Rapor1NoKontAyni = 0
    Rapor1NoKontAltNoHata = 0
    Rapor1NoKontUstNoHata = 0
    
    MaxiR = Application.Max(OgeTuruKont, OgeDegeriKont, AdetKont, OgeIdNoKont, AciklamaKont, SonucKont)
    If MaxiR > 0 Then
        'Combolara girilen rapor1 numaraları aynı olamaz.
        For j = 1 To MaxiR
            If Controls("Rapor1No" & j).Value <> "" And Controls("Rapor1No" & j).Value = Rapor1No.Value Then
                Rapor1NoKontAyni = 1
            End If
            
            If Rapor1No.Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Rapor1No.Value, "-") <> 0 And InStr(Controls("Rapor1No" & j).Value, "-") = 0 Then
                Rapor1NoKontAltNoHata = 1
            ElseIf Rapor1No.Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Rapor1No.Value, "-") = 0 And InStr(Controls("Rapor1No" & j).Value, "-") <> 0 Then
                Rapor1NoKontAltNoHata = 1
            End If
            
            If Rapor1No.Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Rapor1No.Value, "-") <> 0 And InStr(Controls("Rapor1No" & j).Value, "-") <> 0 Then
                If Left(Rapor1No.Value, InStr(Rapor1No.Value, "-") - 1) <> Left(Controls("Rapor1No" & j).Value, InStr(Controls("Rapor1No" & j).Value, "-") - 1) Then
                    Rapor1NoKontUstNoHata = 1
                End If
            ElseIf Rapor1No.Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Rapor1No.Value, "-") = 0 And InStr(Controls("Rapor1No" & j).Value, "-") = 0 Then
                Rapor1NoKontUstNoHata = 1
            End If

            If MaxiR >= j + 1 Then
                For i = j + 1 To MaxiR
                    If Controls("Rapor1No" & i).Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Controls("Rapor1No" & i).Value, "-") <> 0 And InStr(Controls("Rapor1No" & j).Value, "-") <> 0 Then
                        If Left(Controls("Rapor1No" & i).Value, InStr(Controls("Rapor1No" & i).Value, "-") - 1) <> Left(Controls("Rapor1No" & j).Value, InStr(Controls("Rapor1No" & j).Value, "-") - 1) Then
                            Rapor1NoKontUstNoHata = 1
                        End If
                    ElseIf Controls("Rapor1No" & i).Value <> "" And Controls("Rapor1No" & j).Value <> "" And InStr(Controls("Rapor1No" & i).Value, "-") = 0 And InStr(Controls("Rapor1No" & j).Value, "-") = 0 Then
                        Rapor1NoKontUstNoHata = 1
                    End If
                    If Controls("Rapor1No" & j).Value <> "" And Controls("Rapor1No" & j).Value = Controls("Rapor1No" & i).Value Then
                        Rapor1NoKontAyni = 1
                    End If
                Next i
            End If
        Next j
        
        For i = 1 To MaxiR
            If Controls("OgeTuru" & i).Value = "" Then
                OgeTuruKontSatir = i
            End If
            If Controls("OgeDegeri" & i).Value = "" Then
                OgeDegeriKontSatir = i
            End If
            If Controls("Adet" & i).Value = "" Then
                AdetKontSatir = i
            End If
            If Controls("OgeIdNo" & i).Value = "" Then
                OgeIdNoKontSatir = i
            End If
    '        If Controls("Aciklama" & i).Value = "" Then
    '            AciklamaKontSatir = i
    '        End If
            If Controls("Sonuc" & i).Value = "" Then
                SonucKontSatir = i
            End If
            'Rapor1 noyu valid/invalid durumuna göre kontrol et
            If i = 1 Then
                If Controls("Sonuc" & i).Value <> "" And Controls("Sonuc" & i).Value <> Sonuc.Value And Controls("Rapor1No" & i).Value = "" Then
                    Rapor1NoKont = i
                End If
            End If
            If i > 1 Then
                If Controls("Sonuc" & i).Value <> "" And Controls("Sonuc" & i).Value <> Controls("Sonuc" & i - 1) And Controls("Rapor1No" & i).Value = "" Then
                    Rapor1NoKont = i
                End If
            End If

            '__________Rapor No Senkronizasyon 30.11.2021
         
            'Rapor1 numarasının daha önce kullanılıp kullanılmadığını kontrol et.
            If i >= 1 And Controls("Rapor1No" & i).Value <> "" Then
            
                If InStr(Controls("Rapor1No" & i).Value, "-") = 0 Then 'rapor no girişinde tire yok
                    
                    'Tiresiz değerler içinde ara
                    StrAramaGlobal = Controls("Rapor1No" & i).Value
                    Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
                    Set Rapor1NoBul = MyRngGlobal.Find(What:=StrAramaGlobal, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not Rapor1NoBul Is Nothing Then
                        Rapor1NoBulKont = i
                    End If
                
                    'Tireli değerler içinde ara
                    StrAramaGlobal = Controls("Rapor1No" & i).Value & "-"
                    Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
                    Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                                    SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
                    If Not MyFinderGlobal Is Nothing Then
                        IlkAdresGlobal = MyFinderGlobal.Address
                        'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                        If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                            Rapor1NoBulKont = i
                        End If
                        'Sonraki satırlarda aramaya devam et
                        Do
                            SonrakiAdresGlobal = MyFinderGlobal.Address
                            'MsgBox Replace(SonrakiAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                            Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
                            SonrakiAdresGlobal = MyFinderGlobal.Address
                            If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                                Rapor1NoBulKont = i
                            End If
                        Loop While IlkAdresGlobal <> SonrakiAdresGlobal
                    End If
                
                Else 'rapor no girişinde tire var
                    
                    RaporTireTek = RaporTireTek + 1 '82-1 gibi tek değer girilemez
        
                    'Tiresiz değerler içinde ara
                    StrAramaGlobal = Left(Controls("Rapor1No" & i).Value, InStr(Controls("Rapor1No" & i).Value, "-") - 1)
                    Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
                    Set Rapor1NoBul = MyRngGlobal.Find(What:=StrAramaGlobal, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not Rapor1NoBul Is Nothing Then
                        Rapor1NoBulTireKont = i
                    End If
                
                    'Tireli değerler içinde ara
                    StrAramaGlobal = Left(Controls("Rapor1No" & i).Value, InStr(Controls("Rapor1No" & i).Value, "-") - 1) & "-"
                    Set MyRngGlobal = WsRaporNo.Range("A" & RefSatir & ":A100000")
                    Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                                    SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
                    If Not MyFinderGlobal Is Nothing Then
                        IlkAdresGlobal = MyFinderGlobal.Address
                        'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                        If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                            Rapor1NoBulTireKont = i
                        End If
                        'Sonraki satırlarda aramaya devam et
                        Do
                            SonrakiAdresGlobal = MyFinderGlobal.Address
                            'MsgBox Replace(SonrakiAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
                            Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
                            SonrakiAdresGlobal = MyFinderGlobal.Address
                            If Left(MyFinderGlobal.Value, Len(StrAramaGlobal)) = StrAramaGlobal Then
                                Rapor1NoBulTireKont = i
                            End If
                        Loop While IlkAdresGlobal <> SonrakiAdresGlobal
                    End If
        
                End If
            End If
        
            '__________Rapor No Senkronizasyon 30.11.2021
       
        Next i
    End If
    

    '__________Rapor No Senkronizasyon 30.11.2021
    
    If ComboGetir.Value <> "" Then 'Düzenleme işlemi ise cari işlemin rapor no.su aramalara takılmasın diye yukarıda yapılan işlemin geri alınması
        Set RnoIlkSiraBul = WsRaporNo.Range("J6:J100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("K6:K100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not RnoIlkSiraBul Is Nothing Then
            RnoIlkSira = RnoIlkSiraBul.Row
            If Not RnoSonSiraBul Is Nothing Then
                RnoSonSira = RnoSonSiraBul.Row
            End If
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 1), WsRaporNo.Cells(RnoSonSira, 5)).Value = WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 7), WsRaporNo.Cells(RnoSonSira, 11)).Value
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 7), WsRaporNo.Cells(RnoSonSira, 11)).ClearContents
        End If
    End If

    GoTo DuzeltmeniYapDaGit1Atla
DuzeltmeniYapDaGit1:
    If ComboGetir.Value <> "" Then 'Düzenleme işlemi ise cari işlemin rapor no.su aramalara takılmasın diye yukarıda yapılan işlemin geri alınması
        Set RnoIlkSiraBul = WsRaporNo.Range("J6:J100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("K6:K100000").Find(What:=Cells(IlkSiraGlobal, 85).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not RnoIlkSiraBul Is Nothing Then
            RnoIlkSira = RnoIlkSiraBul.Row
            If Not RnoSonSiraBul Is Nothing Then
                RnoSonSira = RnoSonSiraBul.Row
            End If
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 1), WsRaporNo.Cells(RnoSonSira, 5)).Value = WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 7), WsRaporNo.Cells(RnoSonSira, 11)).Value
            WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 7), WsRaporNo.Cells(RnoSonSira, 11)).ClearContents
        End If
    End If
    GoTo SonRapor
DuzeltmeniYapDaGit1Atla:

    '__________Rapor No Senkronizasyon 30.11.2021



    If SonucKontSatir <> 0 Then
        Bilgi = MsgBox("At least one result row is missing or incomplete. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet22
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet22:
    If Rapor1No.Value = "" Then
        Bilgi = MsgBox("Report 1 number has not been entered. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23:
    ' Check for duplicate usage of Report1 numbers in combo boxes
    If Rapor1NoKontAyni <> 0 Then
        Bilgi = MsgBox("A duplicate number has been detected in at least one of the Report 1 number fields. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1A
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1A:
    If Rapor1NoKontAltNoHata <> 0 Then
        Bilgi = MsgBox("At least one Report 1 number field is missing a sub-number (e.g., '318' and '318-1'). Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1B
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1B:
    If Rapor1NoKontUstNoHata <> 0 Then
        Bilgi = MsgBox("At least one Report 1 number field has mismatched main numbers (e.g., '318-1' and '319-2' or '318' and '319'). Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1C
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1C:
    ' Validate Report1 number based on valid/invalid status
    If Rapor1NoKont <> 0 Then
        Bilgi = MsgBox("The Report 1 number has been entered incorrectly or is missing. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1:
    
    If RaporTireTek = 1 Then
        Bilgi = MsgBox("Only a single sub-number (e.g., '82-1') has been specified in the Report 1 number field. Click 'Yes' to proceed with saving, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1TireTek
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1TireTek:

    If Rapor1NoBulKont <> 0 Then
        Bilgi = MsgBox("At least one of the Report 1 numbers has been detected as previously used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek2
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
    If Rapor1NoBulTireKont <> 0 Then
        Bilgi = MsgBox("At least one of the Report 1 numbers has been detected as previously used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek2
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
    If Rapor1NoBulTireKontPart <> 0 Then
        Bilgi = MsgBox("At least one of the Report 1 numbers has been detected as previously used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek2
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
    
YinedeKaydet23Ek2:
    If Rapor1TarihiText.Value = "" Then
        Bilgi = MsgBox("Report 1 date has not been specified. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet24
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
    
YinedeKaydet24:
    If MaxiR > Maxi Then
        Bilgi = MsgBox("It has been detected that the number of result rows exceeds the number of item type rows. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet25
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
    
YinedeKaydet25:
    'Report1 date chronological check
    If Rapor1TarihiText.Value <> "" And _
       Tutanak1TarihiText.Value <> "" And GelisTarihiText.Value <> "" And BelgeTarihiText.Value <> "" Then
        If Year(Rapor1TarihiText.Value) < Year(BelgeTarihiText.Value) Or _
           Year(Rapor1TarihiText.Value) < Year(GelisTarihiText.Value) Or _
           Year(Rapor1TarihiText.Value) < Year(Tutanak1TarihiText.Value) Then
           
            Bilgi = MsgBox("It has been detected that the Report 1 date is earlier than the incoming document date, the received date, and/or Statement 1 date. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet25Ek1
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
    
YinedeKaydet25Ek1:
        If (Year(Rapor1TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Rapor1TarihiText.Value) < Month(BelgeTarihiText.Value)) Or _
           (Year(Rapor1TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Rapor1TarihiText.Value) < Month(GelisTarihiText.Value)) Or _
           (Year(Rapor1TarihiText.Value) = Year(Tutanak1TarihiText.Value) And Month(Rapor1TarihiText.Value) < Month(Tutanak1TarihiText.Value)) Then
           
            Bilgi = MsgBox("It has been detected that the Report1 date is earlier than the incoming document date, the received date, and/or Statement 1 date. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet25Ek2
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
    
YinedeKaydet25Ek2:
        If (Year(Rapor1TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Rapor1TarihiText.Value) = Month(BelgeTarihiText.Value) And Day(Rapor1TarihiText.Value) < Day(BelgeTarihiText.Value)) Or _
           (Year(Rapor1TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Rapor1TarihiText.Value) = Month(GelisTarihiText.Value) And Day(Rapor1TarihiText.Value) < Day(GelisTarihiText.Value)) Or _
           (Year(Rapor1TarihiText.Value) = Year(Tutanak1TarihiText.Value) And Month(Rapor1TarihiText.Value) = Month(Tutanak1TarihiText.Value) And Day(Rapor1TarihiText.Value) < Day(Tutanak1TarihiText.Value)) Then
           
            Bilgi = MsgBox("It has been detected that the Report1 date is earlier than the incoming document date, the received date, and/or Statement 1 date. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet25Ek3
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
YinedeKaydet25Ek3:
    End If
End If

'Tutanak2 controls
If Tutanak2Frame.Visible = True Then
    Tutanak2Kont = 0
    If Tutanak2TarihiText.Value = "" Then
        Bilgi = MsgBox("Statement 2 date has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet26
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet26:
    If GidenMuhatapTemasi.Value = "" Then
        Bilgi = MsgBox("Outgoing Contact Theme has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet27Ek1
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet27Ek1:
    If IlGiden.Value = "" Then
        Bilgi = MsgBox("The province to which the response letter will be sent has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet27Ek3
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet27Ek3:
    If InStr(GidenMuhatapTemasi.Value, "İlçe") <> 0 Then
        If IlceGiden.Value = "" Then
            Bilgi = MsgBox("Although the Outgoing Contact Theme includes a district, the district to which the response letter will be sent has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet27Ek2
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
    End If
YinedeKaydet27Ek2:
    If GonderilenBirim.Value = "" Then
        Bilgi = MsgBox("The recipient Subunit has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet27
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet27:
    If GidenPaketTipi.Value = "" Then
        Bilgi = MsgBox("Outgoing package type has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet28
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet28:
    If GidenPaketAdedi.Value = "" Then
        Bilgi = MsgBox("Outgoing package quantity has not been specified. To save anyway, please click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet29
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet29:

    ' Statement 2 date precedence check
    If Tutanak2TarihiText.Value <> "" And Rapor1TarihiText.Value <> "" And _
       Tutanak1TarihiText.Value <> "" And GelisTarihiText.Value <> "" And BelgeTarihiText.Value <> "" Then
    
        If Year(Tutanak2TarihiText.Value) < Year(BelgeTarihiText.Value) Or _
           Year(Tutanak2TarihiText.Value) < Year(GelisTarihiText.Value) Or _
           Year(Tutanak2TarihiText.Value) < Year(Tutanak1TarihiText.Value) Or _
           Year(Tutanak2TarihiText.Value) < Year(Rapor1TarihiText.Value) Then
    
            Bilgi = MsgBox("It has been detected that Statement 2 date is earlier than the incoming document date, received date, Statement 1 date, and/or Report 1 date. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet29Ek1
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
    
YinedeKaydet29Ek1:
        If (Year(Tutanak2TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(BelgeTarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(GelisTarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(Tutanak1TarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(Tutanak1TarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(Rapor1TarihiText.Value)) Then
    
            Bilgi = MsgBox("It has been detected that Statement 2 date is earlier than the incoming document date, received date, Statement 1 date, and/or Report 1 date. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet29Ek2
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
    
YinedeKaydet29Ek2:
        If (Year(Tutanak2TarihiText.Value) = Year(BelgeTarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(BelgeTarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(BelgeTarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(GelisTarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(GelisTarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(GelisTarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(Tutanak1TarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(Tutanak1TarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(Tutanak1TarihiText.Value)) Or _
           (Year(Tutanak2TarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(Rapor1TarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(Rapor1TarihiText.Value)) Then
    
            Bilgi = MsgBox("It has been detected that Statement 2 date is earlier than the incoming document date, received date, Statement 1 date, and/or Report 1 date. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet29Ek3
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
    
YinedeKaydet29Ek3:
    End If
End If

' cover letter checks
' UstYaziKont = 2
If UstYaziFrame.Visible = True Then
    UstYaziKont = 0
    If UstYaziTarihiText.Value = "" Then
        Bilgi = MsgBox("It has been detected that the cover letter date is not specified. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            UstYaziKont = 1
            GoTo YinedeKaydet30
        ElseIf Bilgi = vbNo Then
            UstYaziKont = 2
            GoTo SonUstYazi
        End If
    End If
YinedeKaydet30:
    If UstYaziNoText.Value = "" Then
        Bilgi = MsgBox("It has been detected that the cover letter number is not specified. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            UstYaziKont = 1
            GoTo YinedeKaydet31
        ElseIf Bilgi = vbNo Then
            UstYaziKont = 2
            GoTo SonUstYazi
        End If
    End If
YinedeKaydet31:
    ' cover letter date precedence check
    If UstYaziTarihiText.Value <> "" And Tutanak2TarihiText.Value <> "" And Rapor1TarihiText.Value <> "" And _
       Tutanak1TarihiText.Value <> "" And GelisTarihiText.Value <> "" And BelgeTarihiText.Value <> "" Then

        If CDate(UstYaziTarihiText.Value) < CDate(BelgeTarihiText.Value) Or _
           CDate(UstYaziTarihiText.Value) < CDate(GelisTarihiText.Value) Or _
           CDate(UstYaziTarihiText.Value) < CDate(Tutanak1TarihiText.Value) Or _
           CDate(UstYaziTarihiText.Value) < CDate(Rapor1TarihiText.Value) Or _
           CDate(UstYaziTarihiText.Value) < CDate(Tutanak2TarihiText.Value) Then

            Bilgi = MsgBox("It has been detected that the cover letter date is earlier than the incoming document date, received date, Statement 1, Report 1, and/or Statement 2 dates. To proceed with saving, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                UstYaziKont = 1
                GoTo YinedeKaydet32
            ElseIf Bilgi = vbNo Then
                UstYaziKont = 2
                GoTo SonUstYazi
            End If
        End If
YinedeKaydet32:

        If GelenMuhatapTemasi.Value = GidenMuhatapTemasi.Value And _
        ((GonderenBirim.Value = "Incoming Contact Theme" And GonderilenBirim.Value = "Outgoing Contact Theme") Or GonderenBirim.Value = GonderilenBirim.Value) And _
        Il.Value = IlGiden.Value And _
        Ilce.Value = IlceGiden.Value Then 'Gelen ve Giden yazı aynı yere ise
            '
        Else
            If IlgiYaziFotokopisi.Value = "" Then
                Bilgi = MsgBox("At least one of the province, district, sending subunit, or Incoming Contact Theme fields in the incoming document does not match any of the corresponding fields in the cover letter. Therefore, the cover letter will be prepared for a different recipient than specified in the incoming document." & vbNewLine & vbNewLine & _
                        "Please specify the number of pages of the referenced document to be attached with the cover letter in the 'Referenced Document Page Count' field. Click 'Yes' to save with this information or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
                If Bilgi = vbYes Then
                    UstYaziKont = 1
                    GoTo YinedeKaydet31x
                ElseIf Bilgi = vbNo Then
                    UstYaziKont = 2
                    GoTo SonUstYazi
                End If
            End If
        End If
YinedeKaydet31x:
    End If
End If


Son:
SonRapor:
SonTutanak2:
SonUstYazi:

End Sub

Private Sub Kaydet_Click()
Dim YeniIslem As Long
Dim i As Long, j As Long, OgeFrame As Integer, Kont As Integer
Dim ctl As MSForms.Control
Dim Bilgi As Variant
Dim OgeTuruKont As Integer, OgeDegeriKont As Integer, AdetKont As Integer
Dim OgeIdNoKont As Integer, AciklamaKont As Integer, SonucKont As Integer
Dim OgeTuruKontSatir As Integer, OgeDegeriKontSatir As Integer, AdetKontSatir As Integer
Dim OgeIdNoKontSatir As Integer, AciklamaKontSatir As Integer, SonucKontSatir As Integer
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range, Fark As Long
Dim FarkSay As Integer, SiraNoSakla As Long, SiraSay As Long
Dim Rapor1NoKont As Integer, Rapor1NoBul As Range, Rapor1NoBulIlk As Range
Dim Rapor1NoBulTire As Range, Rapor1NoBulTireKont As Integer, Rapor1NoBulKont As Integer
Dim Rapor1NoBulTirePart As Range, Rapor1NoBulTireKontPart As Integer
Dim Kenarlar As Range, DokumKontSatir As Integer, UserName As String
Dim Rapor1NoKontAyni As Integer, Rapor1NoKontAltNoHata As Integer, Rapor1NoKontUstNoHata As Integer

Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim OpenControl As String, Say1IslemGunlugu As Long, Say2IslemGunlugu As Long
Dim GelenTema As String, Sene As String, Ay As String
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long
Dim ItemBul As Range
Dim RefSatir As Long, Rapor1TarihBul As Range

Dim FarkSay1 As Integer, Fark1 As Long


ThisWorkbook.Activate

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"
ThisWorkbook.Worksheets(8).Unprotect Password:="123"
ThisWorkbook.Worksheets(11).Unprotect Password:="123"

UserName = Environ("UserProfile")
UserName = UCase(Right(UserName, 7)) 'UCase(Replace(Replace(Mid(Right(UserName, 7), 4, 2), "i", "I"), "ı", "I"))

Tutanak1Kont = 3
Rapor1Kont = 3
Tutanak2Kont = 3
UstYaziKont = 3
YeniIslem = 0
'___________________

'Sıra numarası bulunamazsa prosedürden çık (Bu kısım zorunlu değildir. Esas bölüm düzeltme ksımındadır.)
'Kullanıcının sıra numarası vermesi engellenniş olur.


'__________Rapor No Senkronizasyon 30.11.2021

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ResetAtla
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ResetAtla
    End If
End If
'__________Rapor No Senkronizasyon 30.11.2021


'__________________

'Tüm bölümler için ön kontrol

Call KaydetOnKontroller

If TumKont = 0 Then
    'MsgBox "Tümü boş."
    'Tutanak1Kont = 2
    'GoTo Son
    GoTo Out
End If
'MsgBox "En az biri dolu."

'Düzeltme kaydı ve yeni işlem bilgilendirme mesajı
If ComboGetir.Value <> "" Then
    Bilgi = MsgBox("The operation you are about to perform is a EDIT record for the transaction with serial number " & ComboGetir.Value & "." & vbNewLine & vbNewLine & _
                   "Click " & """" & "Yes" & """" & " to proceed with the edit, or " & """" & "No" & """" & " to cancel.", vbYesNo + vbInformation, "Enterprise Document Automation System")
    If Bilgi = vbNo Then
        GoTo Out:
    End If
Else
    Bilgi = MsgBox("You are about to create a NEW transaction." & vbNewLine & vbNewLine & _
                   "Click " & """" & "Yes" & """" & " to proceed with the new entry, or " & """" & "No" & """" & " to cancel.", vbYesNo + vbInformation, "Enterprise Document Automation System")
    If Bilgi = vbNo Then
        GoTo Out:
    End If
End If


'______________

Call KontrolProseduru

If Tutanak1Kont = 2 Then
    GoTo Son
End If
If Rapor1Kont = 2 Then
    GoTo SonRapor
End If
If Tutanak2Kont = 2 Then
    GoTo SonTutanak2
End If
If UstYaziKont = 2 Then
    GoTo SonUstYazi
End If

'______________



'DÜZELTME KAYDI
If ComboGetir.Value <> "" Then
    'Veri tabanını kontrol et
    Say = Range("CE100000").End(xlUp).Row
    If Say < 7 Then
        GoTo ResetAtla
    End If
    
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ResetAtla
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ResetAtla
    End If
    
'    IlkSiraAktar = IlkSira
    'SonSiraAktar = SonSira 'Bunu kasıtlı kapatıyorum. Çünkü son sira aşağıda değişebilir. İşlem günlüğü için bunu yaptım.
    '25.11.2021, 20:49, İLAVE (işlem günlüğünde yapılan güncelleme için)
    YeniIslemAktar = IlkSira

'    'Önceki veriyi sil (04.07.2019, 23:40)
'    'Kullanıcı bir işlemi düzenlemek için çağırır ve
'    '(visible frame kombinasyonu değişirse) veriler tamamen yeni karara göre kaydedilir.
'Başlangıç ve Bitiş numaraları, sayfalar, varlık takipleri silinmeyecek (CM ve DG arası silinmeyecek)
    Range("F" & IlkSira & ":BX" & SonSira).ClearContents
    Range("CZ" & IlkSira & ":DW" & SonSira).ClearContents 'En sondaki sayfa sayıları da hariç
    
    'Rapor1 formundan  sayfaya aktar.
    'Tutanak1 bölümü
    'Cells(IlkSira, 5)'Sıra numarası
    Cells(IlkSira, 17).Value = Il.Value
    Cells(IlkSira, 18).Value = Ilce.Value
    Cells(IlkSira, 20).Value = BelgeTarihiText.Value
    Cells(IlkSira, 21).Value = BelgeNoText.Value
    Cells(IlkSira, 22).Value = TemaTipi.Value
    Cells(IlkSira, 23).Value = TemaNoText.Value
    If OtomatikOption.Value = True Then
        Cells(IlkSira, 24).Value = "Otomatik"
    ElseIf ManuelOption.Value = True Then
        Cells(IlkSira, 24).Value = "Manuel"
    End If
    Cells(IlkSira, 25).Value = GelenMuhatapTemasi.Value
    If GonderenBirim.Value = "Incoming Contact Theme" Or GonderenBirim.Value = "Outgoing Contact Theme" Then
        Cells(IlkSira, 26).Value = ""
    Else
        Cells(IlkSira, 26).Value = GonderenBirim.Value
    End If
    Cells(IlkSira, 28).Value = GelisTarihiText.Value
    Cells(IlkSira, 29).Value = GelenPaketTipi.Value
    Cells(IlkSira, 30).Value = GelisSekli.Value
    Cells(IlkSira, 31).Value = Tutanak1TarihiText.Value
    Cells(IlkSira, 32).Value = Tutanak1Sonucu.Value
    Cells(IlkSira, 33).Value = GelenBelgeSayfa.Value
    Cells(IlkSira, 34).Value = DosyaNoText.Value
    If DLEvetOption.Value = True Then
        Cells(IlkSira, 35).Value = "Yes"
    ElseIf DLHayirOption.Value = True Then
        Cells(IlkSira, 35).Value = "No"
    End If
    Cells(IlkSira, 36).Value = DokumListesi.Value

    'Tutanak1 imzaları
    Cells(IlkSira, 104).Value = Tutanak1Imza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak1Imza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak1Imza1DuzeltmeIslemAtla
    End If
    Cells(IlkSira, 105).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(IlkSira, 106).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak1Imza1DuzeltmeIslemAtla:
    
    Cells(IlkSira, 107).Value = Tutanak1Imza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak1Imza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak1Imza2DuzeltmeIslemAtla
    End If
    Cells(IlkSira, 108).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(IlkSira, 109).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak1Imza2DuzeltmeIslemAtla:
''''''''''''tutanak1 imza sonu

    Cells(IlkSira, 38).Value = OgeTuru.Value
    Cells(IlkSira, 41).Value = OgeDegeri.Value
    Cells(IlkSira, 44).Value = Adet.Value
    Cells(IlkSira, 47).Value = OgeIdNo.Value
    Cells(IlkSira, 50).Value = Aciklama.Value
    
    'Rapor1
    If Rapor1Frame.Visible = True Then
        Cells(IlkSira, 54).Value = Sonuc.Value
        Cells(IlkSira, 59).Value = Rapor1No.Value
        Cells(IlkSira, 11).Value = Rapor1No.Value
        Cells(IlkSira, 60).Value = Rapor1TarihiText.Value
        'İmzalar (hazırlık)
        StrRaporUnvan1 = ""
        StrRaporSicil1 = ""
        StrRaporUnvan2 = ""
        StrRaporSicil2 = ""
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo RaporImza1DuzeltmeIslemAtla
        End If
        StrRaporUnvan1 = Worksheets(2).Range("DZ" & ItemBul.Row)
        StrRaporSicil1 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza1DuzeltmeIslemAtla:
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo RaporImza2DuzeltmeIslemAtla
        End If
        StrRaporUnvan2 = Worksheets(2).Range("DZ" & ItemBul.Row)
        StrRaporSicil2 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza2DuzeltmeIslemAtla:
    '''''''''''İmzalar (hazırlık) sonu
    End If
    
    'Fark girişleri
    If SayFarkGiris > 2 Then  'Kalıcı kaydı yap, geçici kaydı sil
    
'        Set WsFarkGiris = ThisWorkbook.Worksheets(7)
'        'Maksimum değerler.
'        SayA = WsFarkGiris.Range("A100000").End(xlUp).Row
'        SayD = WsFarkGiris.Range("D100000").End(xlUp).Row
'        SayG = WsFarkGiris.Range("G100000").End(xlUp).Row
'        SayJ = WsFarkGiris.Range("J100000").End(xlUp).Row
'        SayM = WsFarkGiris.Range("M100000").End(xlUp).Row
'        SayFarkGiris = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)
        
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
        Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
        Else
            'Maksimum değerler.
            SayA = WsFarkGirisRapor1.Range("A100000").End(xlUp).Row
            SayD = WsFarkGirisRapor1.Range("D100000").End(xlUp).Row
            SayG = WsFarkGirisRapor1.Range("G100000").End(xlUp).Row
            SayJ = WsFarkGirisRapor1.Range("J100000").End(xlUp).Row
            SayM = WsFarkGirisRapor1.Range("M100000").End(xlUp).Row
            SayFarkGirisRapor1 = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)
            If SayFarkGirisRapor1 < 3 Then
                SayFarkGirisRapor1 = 2
            End If
            WsFarkGirisRapor1.Range("A" & SayFarkGirisRapor1 + 1 & ":M" & SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3)).Value = WsFarkGiris.Range("A3:M" & SayFarkGiris).Value
            'Fark girişleri için başlangıç bitiş sıra no.ları işaretle
            WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1, 16).Value = ComboGetir.Value
            WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3), 17).Value = ComboGetir.Value
            'Geçici kaydı sil
            WsFarkGiris.Rows("3:30").EntireRow.Delete
    
            GoTo FarkGirisDuzeltmeAtla
        End If
        If Not SonSiraBul Is Nothing Then
            SonSiraFarkGirisRapor1 = SonSiraBul.Row
        Else
            'Maksimum değerler.
            SayA = WsFarkGirisRapor1.Range("A100000").End(xlUp).Row
            SayD = WsFarkGirisRapor1.Range("D100000").End(xlUp).Row
            SayG = WsFarkGirisRapor1.Range("G100000").End(xlUp).Row
            SayJ = WsFarkGirisRapor1.Range("J100000").End(xlUp).Row
            SayM = WsFarkGirisRapor1.Range("M100000").End(xlUp).Row
            SayFarkGirisRapor1 = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)
            If SayFarkGirisRapor1 < 3 Then
                SayFarkGirisRapor1 = 2
            End If
            WsFarkGirisRapor1.Range("A" & SayFarkGirisRapor1 + 1 & ":M" & SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3)).Value = WsFarkGiris.Range("A3:M" & SayFarkGiris).Value
            'Fark girişleri için başlangıç bitiş sıra no.ları işaretle
            WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1, 16).Value = ComboGetir.Value
            WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3), 17).Value = ComboGetir.Value
            'Geçici kaydı sil
            WsFarkGiris.Rows("3:30").EntireRow.Delete
            
            GoTo FarkGirisDuzeltmeAtla
        End If
    
        Maxi1 = SayFarkGiris - 3 + 1
        Fark1 = SonSiraFarkGirisRapor1 - IlkSiraFarkGirisRapor1 + 1
        
        If Maxi1 = Fark1 Then 'Sayfadaki satır sayısını değiştirme
            If Maxi1 > 0 And Maxi1 < 21 Then
                For OgeFrame = 0 To Maxi1 - 1
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 1).Value = WsFarkGiris.Cells(3 + OgeFrame, 1).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 4).Value = WsFarkGiris.Cells(3 + OgeFrame, 4).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 7).Value = WsFarkGiris.Cells(3 + OgeFrame, 7).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 10).Value = WsFarkGiris.Cells(3 + OgeFrame, 10).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 13).Value = WsFarkGiris.Cells(3 + OgeFrame, 13).Value
                Next OgeFrame
                'Geçici kaydı sil
                WsFarkGiris.Rows(3 & ":" & SayFarkGiris).EntireRow.Delete
            End If
        ElseIf Maxi1 > Fark1 Then 'Sayfaya satır ekle
            If Maxi1 > 0 And Maxi1 < 21 Then
                FarkSay1 = 0
                For i = 1 To Maxi1 - Fark1
                    WsFarkGirisRapor1.Rows(SonSiraFarkGirisRapor1 + 1).EntireRow.Insert Shift:=xlDown
                    FarkSay1 = FarkSay1 + 1
                Next i
                Application.CutCopyMode = False
                Application.CutCopyMode = True
                WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1 + FarkSay1, 17).Value = WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1, 17).Value
                WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1, 17).Value = ""
                For OgeFrame = 0 To Maxi1 - 1
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 1).Value = WsFarkGiris.Cells(3 + OgeFrame, 1).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 4).Value = WsFarkGiris.Cells(3 + OgeFrame, 4).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 7).Value = WsFarkGiris.Cells(3 + OgeFrame, 7).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 10).Value = WsFarkGiris.Cells(3 + OgeFrame, 10).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 13).Value = WsFarkGiris.Cells(3 + OgeFrame, 13).Value
                Next OgeFrame
                'Geçici kaydı sil
                WsFarkGiris.Rows(3 & ":" & SayFarkGiris).EntireRow.Delete
            End If
        ElseIf Maxi1 < Fark1 Then 'Sayfadan satır sil
            FarkSay1 = 0
            SiraNoSakla = WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1, 17).Value
            For i = 1 To Fark1 - Maxi1
                FarkSay1 = FarkSay1 + 1
                WsFarkGirisRapor1.Rows(SonSiraFarkGirisRapor1 - (FarkSay1 - 1)).EntireRow.Delete 'Shift:=xlDown
            Next i
            WsFarkGirisRapor1.Cells(SonSiraFarkGirisRapor1 - FarkSay1, 17).Value = SiraNoSakla
            If Maxi1 > 0 Then
                For OgeFrame = 0 To Maxi1 - 1
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 1).Value = WsFarkGiris.Cells(3 + OgeFrame, 1).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 4).Value = WsFarkGiris.Cells(3 + OgeFrame, 4).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 7).Value = WsFarkGiris.Cells(3 + OgeFrame, 7).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 10).Value = WsFarkGiris.Cells(3 + OgeFrame, 10).Value
                    WsFarkGirisRapor1.Cells(IlkSiraFarkGirisRapor1 + OgeFrame, 13).Value = WsFarkGiris.Cells(3 + OgeFrame, 13).Value
                Next OgeFrame
                'Geçici kaydı sil
                WsFarkGiris.Rows(3 & ":" & SayFarkGiris).EntireRow.Delete
            End If
        End If
FarkGirisDuzeltmeAtla:
    Else 'SayFarkGiris 3'ten küçükse (varsayılan değeri 1)
        
        'Asıl kaydı bul ve sil (İlk kayıtta fark girişi var; düzeltme kaydında ise farklar kaldırılarak geçici kayıt boşaltılmış. Bu durumda kalıcı kayıt da silinmeli.)
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
        Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBul Is Nothing Then
            IlkSiraFarkGirisRapor1 = IlkSiraBul.Row
        Else
            GoTo FarkGirisDuzeltmeAtla1
        End If
        If Not SonSiraBul Is Nothing Then
            SonSiraFarkGirisRapor1 = SonSiraBul.Row
        Else
            GoTo FarkGirisDuzeltmeAtla1
        End If
        'Buraya asıl kayıtların silinmesi kodları gelecek
        WsFarkGirisRapor1.Rows(IlkSiraFarkGirisRapor1 & ":" & SonSiraFarkGirisRapor1).EntireRow.Delete
        
FarkGirisDuzeltmeAtla1:
    End If

    'Tutanak1 ve rapor1
    Maxi = Application.Max(Maxi, MaxiR)
    Fark = SonSira - IlkSira '+ 1
    MaxiAktar = Maxi
    FarkAktar = Fark
    
    If Maxi = Fark Then 'Sayfadaki satır sayısını değiştirme
        If Maxi > 0 And Maxi < 20 Then
            For OgeFrame = 1 To Maxi
                Cells(IlkSira + OgeFrame, 38).Value = Controls("OgeTuru" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 41).Value = Controls("OgeDegeri" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 44).Value = Controls("Adet" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 47).Value = Controls("OgeIdNo" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 50).Value = Controls("Aciklama" & OgeFrame).Value
                'Rapor1
                If Rapor1Frame.Visible = True Then
                    Cells(IlkSira + OgeFrame, 54).Value = Controls("Sonuc" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 59).Value = Controls("Rapor1No" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 11).Value = Controls("Rapor1No" & OgeFrame).Value
                    'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                    If Cells(IlkSira + OgeFrame, 11).Value = "" Then
                        Cells(IlkSira + OgeFrame, 91).Value = ""
                    End If
                End If
            Next OgeFrame
        End If
    ElseIf Maxi > Fark Then 'Sayfaya satır ekle
        If Maxi > 0 And Maxi < 20 Then
            FarkSay = 0
            For i = 1 To Maxi - Fark
                Rows(SonSira + 1).EntireRow.Insert Shift:=xlDown
                'Rows(SonSira + 1).EntireRow.Copy
                'Rows(SonSira + 1 + i).EntireRow.PasteSpecial xlPasteFormats
                FarkSay = FarkSay + 1
            Next i
            Application.CutCopyMode = False
            Application.CutCopyMode = True
            Cells(SonSira + FarkSay, 84).Value = Cells(SonSira, 84).Value
            Cells(SonSira, 84).Value = ""
            For OgeFrame = 1 To Maxi
                Cells(IlkSira + OgeFrame, 38).Value = Controls("OgeTuru" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 41).Value = Controls("OgeDegeri" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 44).Value = Controls("Adet" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 47).Value = Controls("OgeIdNo" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 50).Value = Controls("Aciklama" & OgeFrame).Value
                'Rapor1
                If Rapor1Frame.Visible = True Then
                    Cells(IlkSira + OgeFrame, 54).Value = Controls("Sonuc" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 59).Value = Controls("Rapor1No" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 11).Value = Controls("Rapor1No" & OgeFrame).Value
                    'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                    If Cells(IlkSira + OgeFrame, 11).Value = "" Then
                        Cells(IlkSira + OgeFrame, 91).Value = ""
                    End If
                End If
            Next OgeFrame
        End If
    ElseIf Maxi < Fark Then 'Sayfadan satır sil
            FarkSay = 0
            SiraNoSakla = Cells(SonSira, 84).Value
            For i = 1 To Fark - Maxi
                FarkSay = FarkSay + 1
                Rows(SonSira - (FarkSay - 1)).EntireRow.Delete 'Shift:=xlDown
            Next i
            Cells(SonSira - FarkSay, 84).Value = SiraNoSakla 'Cells(SonSira, 84).Value
            If Maxi > 0 Then
                For OgeFrame = 1 To Maxi
                    Cells(IlkSira + OgeFrame, 38).Value = Controls("OgeTuru" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 41).Value = Controls("OgeDegeri" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 44).Value = Controls("Adet" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 47).Value = Controls("OgeIdNo" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 50).Value = Controls("Aciklama" & OgeFrame).Value
                    'Rapor1
                    If Rapor1Frame.Visible = True Then
                        Cells(IlkSira + OgeFrame, 54).Value = Controls("Sonuc" & OgeFrame).Value
                        Cells(IlkSira + OgeFrame, 59).Value = Controls("Rapor1No" & OgeFrame).Value
                        Cells(IlkSira + OgeFrame, 11).Value = Controls("Rapor1No" & OgeFrame).Value
                        'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                        If Cells(IlkSira + OgeFrame, 11).Value = "" Then
                            Cells(IlkSira + OgeFrame, 91).Value = ""
                        End If
                    End If
                Next OgeFrame
            End If
    End If

    'Tutanak2
    If Tutanak2Frame.Visible = True Then
        Cells(IlkSira, 63).Value = Tutanak2TarihiText.Value
        Cells(IlkSira, 64).Value = GidenMuhatapTemasi.Value

        Cells(IlkSira, 69).Value = IlGiden.Value
        Cells(IlkSira, 70).Value = IlceGiden.Value
        
        If GonderilenBirim.Value = "Outgoing Contact Theme" Or GonderilenBirim.Value = "Incoming Contact Theme" Then
            Cells(IlkSira, 65).Value = ""
        Else
            Cells(IlkSira, 65).Value = GonderilenBirim.Value
        End If
        Cells(IlkSira, 67).Value = GidenPaketTipi.Value
        Cells(IlkSira, 68).Value = GidenPaketAdedi.Value
        'Tutanak2 imzaları
        Cells(IlkSira, 116).Value = Tutanak2Imza1.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo Tutanak2Imza1DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 117).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 118).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza1DuzeltmeIslemAtla:
        
        Cells(IlkSira, 119).Value = Tutanak2Imza2.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo Tutanak2Imza2DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 120).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 121).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza2DuzeltmeIslemAtla:
    ''''''''''''tutanak2 imza sonu
    End If
    
    'Üst yazı
    If UstYaziFrame.Visible = True Then
        Cells(IlkSira, 75).Value = UstYaziTarihiText.Value
        Cells(IlkSira, 76).Value = UstYaziNoText.Value
        Cells(IlkSira, 74).Value = IlgiYaziFotokopisi.Value
        'Üst yazı imzaları
        Cells(IlkSira, 122).Value = UstYaziImza1.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo UstYaziImza1DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 123).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 124).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza1DuzeltmeIslemAtla:
        
        Cells(IlkSira, 125).Value = UstYaziImza2.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo UstYaziImza2DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 126).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 127).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza2DuzeltmeIslemAtla:
    ''''''''''''üst yazı imza sonu
    End If

    'işlem günlüğü için zaman damgası 'ESKİ VERİLER İÇİN ZAMAN DAMGASI OLUŞTUR
    If Len(Cells(IlkSira, 85).Value) < 12 Then
        StrTime = Format(Now, "ddmmyyyyhhmmss")
        Cells(IlkSira, 85).Value = StrTime
    End If


    '__________Rapor No Senkronizasyon 30.11.2021

    Set WsRaporNo = ThisWorkbook.Worksheets(11)

    Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not RnoIlkSiraBul Is Nothing Then
        RnoIlkSira = RnoIlkSiraBul.Row
        If Not RnoSonSiraBul Is Nothing Then
            RnoSonSira = RnoSonSiraBul.Row
        End If

        'Satırları düzenle
        WsRaporNo.Range(WsRaporNo.Cells(RnoIlkSira, 1), WsRaporNo.Cells(RnoSonSira, 5)).ClearContents
        Fark = (RnoSonSira - RnoIlkSira) - (IlkSira + Maxi - IlkSira)
        'MsgBox "Fark: " & Fark
        If Fark > 0 Then 'satır silinecek
            'MsgBox "Fark: " & Fark & " satır kaldır"
            WsRaporNo.Rows(RnoSonSira - (Fark - 1) & ":" & RnoSonSira).EntireRow.Delete
            ilkrow = RnoIlkSira
            sonrow = RnoSonSira - Fark
        ElseIf Fark < 0 Then 'satır eklenecek
            'MsgBox "Fark: " & Fark & " satır ekle"
            Fark = -1 * Fark
            For i = 1 To Fark
                WsRaporNo.Rows(RnoSonSira + 1).EntireRow.Insert Shift:=xlUp
            Next i
            ilkrow = RnoIlkSira
            sonrow = RnoSonSira + Fark
        ElseIf Fark = 0 Then 'satırlarda değişiklik olmayacak
            'MsgBox "Fark: " & Fark & " değişiklik yok"
            ilkrow = RnoIlkSira
            sonrow = RnoSonSira
        End If

        'Verileri aktar
        WsRaporNo.Range(WsRaporNo.Cells(ilkrow, 1), WsRaporNo.Cells(sonrow, 1)).Value = Range(Cells(IlkSira, 11), Cells(IlkSira + Maxi, 11)).Value 'Rapor no
        WsRaporNo.Cells(ilkrow, 2).Value = Cells(IlkSira, 60).Value
        WsRaporNo.Cells(ilkrow, 3).Value = "Request"
        WsRaporNo.Cells(ilkrow, 4).Value = Cells(IlkSira, 85).Value 'İlk zaman damgası
        WsRaporNo.Cells(sonrow, 5).Value = Cells(IlkSira, 85).Value 'Son zaman damgası

    End If

    '__________Rapor No Senkronizasyon 30.11.2021




'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
'
'    'Call ModuleReport1.IslemGunluguRapor1Duzeltme
'
'    Call ModuleReport1.IslemGunluguRapor1
'
'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
    
    ThisWorkbook.Activate
    
    'Prosedür sonu düzeltmeleri
    YeniIslem = IlkSira
    GoTo YeniIslemiAtla
End If

'YinedeKaydet:

'YENİ İŞLEM
YeniIslem = Range("CF100000").End(xlUp).Row
If YeniIslem < 7 Then
    YeniIslem = 7
    GoTo IlkIslem
End If
YeniIslem = YeniIslem + 1
IlkIslem:


'__________Rapor No Senkronizasyon 30.11.2021

Set WsRaporNo = ThisWorkbook.Worksheets(11)
islemNew = WsRaporNo.Range("E100000").End(xlUp).Row
If islemNew < 7 Then
    islemNew = 7
Else
    islemNew = islemNew + 1
End If

'__________Rapor No Senkronizasyon 30.11.2021



Maxi = Application.Max(Maxi, MaxiR)

MaxiAktar = Maxi
YeniIslemAktar = YeniIslem

'Verileri rapor1 formundan sayfaya aktar.
'Tutanak1 bölümü
If YeniIslem = 7 Then
    Cells(YeniIslem, 5).Value = 1 'İlk sıra numarasını ver
Else
    Cells(YeniIslem, 5).Value = Cells(YeniIslem - 1, 84).Value + 1 'Sıra numarası ver
End If
Cells(YeniIslem, 17).Value = Il.Value
Cells(YeniIslem, 18).Value = Ilce.Value
Cells(YeniIslem, 20).Value = BelgeTarihiText.Value
Cells(YeniIslem, 21).Value = BelgeNoText.Value
Cells(YeniIslem, 22).Value = TemaTipi.Value
Cells(YeniIslem, 23).Value = TemaNoText.Value
If OtomatikOption.Value = True Then
    Cells(YeniIslem, 24).Value = "Otomatik"
ElseIf ManuelOption.Value = True Then
    Cells(YeniIslem, 24).Value = "Manuel"
End If
Cells(YeniIslem, 25).Value = GelenMuhatapTemasi.Value
If GonderenBirim.Value = "Outgoing Contact Theme" Or GonderenBirim.Value = "Incoming Contact Theme" Then
    Cells(YeniIslem, 26).Value = ""
Else
    Cells(YeniIslem, 26).Value = GonderenBirim.Value
End If
Cells(YeniIslem, 28).Value = GelisTarihiText.Value
Cells(YeniIslem, 29).Value = GelenPaketTipi.Value
Cells(YeniIslem, 30).Value = GelisSekli.Value
Cells(YeniIslem, 31).Value = Tutanak1TarihiText.Value
Cells(YeniIslem, 32).Value = Tutanak1Sonucu.Value
Cells(YeniIslem, 33).Value = GelenBelgeSayfa.Value
Cells(YeniIslem, 34).Value = DosyaNoText.Value

If DLEvetOption.Value = True Then
    Cells(YeniIslem, 35).Value = "Yes"
ElseIf DLHayirOption.Value = True Then
    Cells(YeniIslem, 35).Value = "No"
End If
Cells(YeniIslem, 36).Value = DokumListesi.Value

'Tutanak1 imzaları
Cells(YeniIslem, 104).Value = Tutanak1Imza1.Value
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak1Imza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo Tutanak1Imza1YeniIslemAtla
End If
Cells(YeniIslem, 105).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
Cells(YeniIslem, 106).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak1Imza1YeniIslemAtla:

Cells(YeniIslem, 107).Value = Tutanak1Imza2.Value
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak1Imza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo Tutanak1Imza2YeniIslemAtla
End If
Cells(YeniIslem, 108).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
Cells(YeniIslem, 109).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak1Imza2YeniIslemAtla:
''''''''''''''

'Fark girişleri
If SayFarkGiris > 2 Then  'Kalıcı kaydı yap, geçici kaydı sil
    Set WsFarkGiris = ThisWorkbook.Worksheets(7)
    Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)

    'Maksimum değerler.
    SayA = WsFarkGirisRapor1.Range("A100000").End(xlUp).Row
    SayD = WsFarkGirisRapor1.Range("D100000").End(xlUp).Row
    SayG = WsFarkGirisRapor1.Range("G100000").End(xlUp).Row
    SayJ = WsFarkGirisRapor1.Range("J100000").End(xlUp).Row
    SayM = WsFarkGirisRapor1.Range("M100000").End(xlUp).Row
    SayFarkGirisRapor1 = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)
    If SayFarkGirisRapor1 < 3 Then
        SayFarkGirisRapor1 = 2
    End If
    
    WsFarkGirisRapor1.Range("A" & SayFarkGirisRapor1 + 1 & ":M" & SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3)).Value = WsFarkGiris.Range("A3:M" & SayFarkGiris).Value
    'Geçici kaydı sil
    WsFarkGiris.Rows(3 & ":" & SayFarkGiris).EntireRow.Delete
End If


Cells(YeniIslem, 38).Value = OgeTuru.Value
Cells(YeniIslem, 41).Value = OgeDegeri.Value
Cells(YeniIslem, 44).Value = Adet.Value
Cells(YeniIslem, 47).Value = OgeIdNo.Value
Cells(YeniIslem, 50).Value = Aciklama.Value
If Maxi > 0 Then
    For OgeFrame = 1 To Maxi
        Cells(YeniIslem + OgeFrame, 38).Value = Controls("OgeTuru" & OgeFrame).Value
        Cells(YeniIslem + OgeFrame, 41).Value = Controls("OgeDegeri" & OgeFrame).Value
        Cells(YeniIslem + OgeFrame, 44).Value = Controls("Adet" & OgeFrame).Value
        Cells(YeniIslem + OgeFrame, 47).Value = Controls("OgeIdNo" & OgeFrame).Value
        Cells(YeniIslem + OgeFrame, 50).Value = Controls("Aciklama" & OgeFrame).Value
    Next OgeFrame
End If

'Rapor bölümü
If Rapor1Frame.Visible = True Then
    Cells(YeniIslem, 54).Value = Sonuc.Value
    Cells(YeniIslem, 59).Value = Rapor1No.Value
    Cells(YeniIslem, 11).Value = Rapor1No.Value
    Cells(YeniIslem, 60).Value = Rapor1TarihiText.Value
    'İmzalar (hazırlık)
    StrRaporUnvan1 = ""
    StrRaporSicil1 = ""
    StrRaporUnvan2 = ""
    StrRaporSicil2 = ""
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo RaporImza1YeniIslemAtla
    End If
    StrRaporUnvan1 = Worksheets(2).Range("DZ" & ItemBul.Row)
    StrRaporSicil1 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza1YeniIslemAtla:
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo RaporImza2YeniIslemAtla
    End If
    StrRaporUnvan2 = Worksheets(2).Range("DZ" & ItemBul.Row)
    StrRaporSicil2 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza2YeniIslemAtla:
'''''''''''İmzalar (hazırlık) sonu

    If Maxi > 0 Then
        For OgeFrame = 1 To Maxi
            Cells(YeniIslem + OgeFrame, 59).Value = Controls("Rapor1No" & OgeFrame).Value
            Cells(YeniIslem + OgeFrame, 11).Value = Controls("Rapor1No" & OgeFrame).Value
            Cells(YeniIslem + OgeFrame, 54).Value = Controls("Sonuc" & OgeFrame).Value
        Next OgeFrame
    End If
End If

'Tutanak2 bölümü
If Tutanak2Frame.Visible = True Then
    Cells(YeniIslem, 63).Value = Tutanak2TarihiText.Value
    Cells(YeniIslem, 64).Value = GidenMuhatapTemasi.Value
    
    Cells(YeniIslem, 69).Value = IlGiden.Value
    Cells(YeniIslem, 70).Value = IlceGiden.Value
    
    If GonderilenBirim.Value = "Outgoing Contact Theme" Or GonderilenBirim.Value = "Incoming Contact Theme" Then
        Cells(YeniIslem, 65).Value = ""
    Else
        Cells(YeniIslem, 65).Value = GonderilenBirim.Value
    End If
    Cells(YeniIslem, 67).Value = GidenPaketTipi.Value
    Cells(YeniIslem, 68).Value = GidenPaketAdedi.Value

    'Tutanak2 imzaları
    Cells(YeniIslem, 116).Value = Tutanak2Imza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak2Imza1YeniIslemAtla
    End If
    Cells(YeniIslem, 117).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 118).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza1YeniIslemAtla:
    
    Cells(YeniIslem, 119).Value = Tutanak2Imza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak2Imza2YeniIslemAtla
    End If
    Cells(YeniIslem, 120).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 121).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza2YeniIslemAtla:
    ''''''''''''''
End If

'Üst yazı bölümü
If UstYaziFrame.Visible = True Then
    Cells(YeniIslem, 75).Value = UstYaziTarihiText.Value
    Cells(YeniIslem, 76).Value = UstYaziNoText.Value
    Cells(YeniIslem, 74).Value = IlgiYaziFotokopisi.Value
    'Üst yazı imzaları
    Cells(YeniIslem, 122).Value = UstYaziImza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo UstYaziImza1YeniIslemAtla
    End If
    Cells(YeniIslem, 123).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 124).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza1YeniIslemAtla:
    
    Cells(YeniIslem, 125).Value = UstYaziImza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo UstYaziImza2YeniIslemAtla
    End If
    Cells(YeniIslem, 126).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 127).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza2YeniIslemAtla:
    ''''''''''''''
End If

'işlem günlüğü için zaman damgası
StrTime = Format(Now, "ddmmyyyyhhmmss")
Cells(YeniIslem, 85).Value = StrTime

'İlk ve son satırları işaretle
Cells(YeniIslem, 83).Value = Cells(YeniIslem, 5).Value
Cells(YeniIslem + Maxi, 84).Value = Cells(YeniIslem, 5).Value

'Fark girişleri
If SayFarkGiris > 2 Then  'Fark girişleri için başlangıç bitiş sıra no.ları işaretle
    WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1, 16).Value = Cells(YeniIslem, 5).Value
    WsFarkGirisRapor1.Cells(SayFarkGirisRapor1 + 1 + (SayFarkGiris - 3), 17).Value = Cells(YeniIslem, 5).Value
End If


'__________Rapor No Senkronizasyon 30.11.2021

WsRaporNo.Range(WsRaporNo.Cells(islemNew, 1), WsRaporNo.Cells(islemNew + Maxi, 1)).Value = Range(Cells(YeniIslem, 11), Cells(YeniIslem + Maxi, 11)).Value 'Rapor no
WsRaporNo.Cells(islemNew, 2).Value = Cells(YeniIslem, 60).Value
WsRaporNo.Cells(islemNew, 3).Value = "Request"
WsRaporNo.Cells(islemNew, 4).Value = Cells(YeniIslem, 85).Value 'İlk zaman damgası
WsRaporNo.Cells(islemNew + Maxi, 5).Value = Cells(YeniIslem, 85).Value 'Son zaman damgası

'__________Rapor No Senkronizasyon 30.11.2021



YeniIslemiAtla:

'________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT

Call ModuleReport1.IslemGunluguRapor1

'________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT


ThisWorkbook.Activate

'Rapor1 için imzalar (Ek bölüm) Hem Düzeltme hem Yeni İşlem için kodlar.
If Rapor1Frame.Visible = True Then
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
        For i = IlkSiraBul.Row To SonSiraBul.Row
            Cells(i, 110).Value = ""
            Cells(i, 111).Value = ""
            Cells(i, 112).Value = ""
            Cells(i, 113).Value = ""
            Cells(i, 114).Value = ""
            Cells(i, 115).Value = ""
            If Cells(i, 59).Value <> "" Then
                Cells(i, 110).Value = RaporImza1.Value
                Cells(i, 111).Value = StrRaporUnvan1
                Cells(i, 112).Value = StrRaporSicil1
                Cells(i, 113).Value = RaporImza2.Value
                Cells(i, 114).Value = StrRaporUnvan2
                Cells(i, 115).Value = StrRaporSicil2
            End If
        Next i
    End If
End If

LblDuzeltme.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
LblDuzeltme.ForeColor = RGB(30, 30, 30)

'Rapor1 no aralıklarını buradan çekiyor.
Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)

'Satır renklendirme ve kenarlıklar.
Set Kenarlar = Range("E" & IlkSiraBul.Row & ":DW" & SonSiraBul.Row)
If Cells(YeniIslem, 5).Value Mod 2 = 0 Then
    Range("E" & IlkSiraBul.Row & ":DW" & SonSiraBul.Row).Interior.Color = RGB(201, 216, 230)
    'Kenarlıklar.
    With Kenarlar
        .Borders(xlEdgeLeft).LineStyle = xlNone '.Color = RGB(174, 185, 194)
        .Borders(xlEdgeTop).Color = RGB(174, 185, 194)
        .Borders(xlEdgeBottom).Color = RGB(174, 185, 194)
        .Borders(xlEdgeRight).LineStyle = xlNone '.Color = RGB(174, 185, 194)
        .Borders(xlInsideVertical).Color = RGB(174, 185, 194)
        .Borders(xlInsideHorizontal).Color = RGB(174, 185, 194)
    End With
Else
    Range("E" & IlkSiraBul.Row & ":DW" & SonSiraBul.Row).Interior.Color = RGB(174, 185, 194) 'RGB(180, 210, 240)
    'Kenarlıklar.
    With Kenarlar
        .Borders(xlEdgeLeft).LineStyle = xlNone '.Color = RGB(254, 254, 254)
        .Borders(xlEdgeTop).Color = RGB(201, 216, 230)
        .Borders(xlEdgeBottom).Color = RGB(201, 216, 230)
        .Borders(xlEdgeRight).LineStyle = xlNone '.Color = RGB(254, 254, 254)
        .Borders(xlInsideVertical).Color = RGB(201, 216, 230)
        .Borders(xlInsideHorizontal).Color = RGB(201, 216, 230)
    End With
End If

Son:
If Tutanak1Kont = 0 Then
    'Normal kaydet
    Cells(YeniIslem, 6).Value = "ü"
    Range("F" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 7).Value = "" Then
        Cells(YeniIslem, 7).Value = "?"
        Range("G" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 8).Value = "" Then
        Cells(YeniIslem, 8).Value = "?"
        Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak1Kont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 6).Value = "x"
    Range("F" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 7).Value = "" Then
        Cells(YeniIslem, 7).Value = "?"
        Range("G" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 8).Value = "" Then
        Cells(YeniIslem, 8).Value = "?"
        Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak1Kont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf Tutanak1Kont = 3 Then
    GoTo ReseteGit
End If


SonRapor:
If Rapor1Kont = 0 Then
    If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
        For i = IlkSiraBul.Row To SonSiraBul.Row
            Cells(i, 7).Value = ""
            Range("G" & i).Font.Color = RGB(60, 100, 180)
            If Cells(i, 59).Value <> "" Then
                'Normal kaydet
                Cells(i, 7).Value = "ü"
                Range("G" & i).Font.Color = RGB(60, 100, 180)
            End If
        Next i
    End If
    If Cells(YeniIslem, 8).Value = "" Then
        Cells(YeniIslem, 8).Value = "?"
        Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Rapor1Kont = 1 Then
    If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
        For i = IlkSiraBul.Row To SonSiraBul.Row
            Cells(i, 7).Value = ""
            Range("G" & i).Font.Color = RGB(60, 100, 180)
            If Cells(i, 59).Value <> "" Then
                'Sorunlu kaydet
                Cells(i, 7).Value = "x"
                Range("G" & i).Font.Color = RGB(60, 100, 180)
            End If
        Next i
        If Cells(IlkSiraBul.Row, 7).Value = "" Then
            'Sorunlu kaydet
            Cells(IlkSiraBul.Row, 7).Value = "x"
            Range("G" & IlkSiraBul.Row).Font.Color = RGB(60, 100, 180)
        End If
    End If
    If Cells(YeniIslem, 8).Value = "" Then
        Cells(YeniIslem, 8).Value = "?"
        Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Rapor1Kont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf Rapor1Kont = 3 Then
    GoTo ReseteGit
End If

SonTutanak2:
If Tutanak2Kont = 0 Then
    'Normal kaydet
    Cells(YeniIslem, 8).Value = "ü"
    Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak2Kont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 8).Value = "x"
    Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak2Kont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf Tutanak2Kont = 3 Then
    GoTo ReseteGit
End If

SonUstYazi:
If UstYaziKont = 0 Then
    'Normal kaydet
    Cells(YeniIslem, 9).Value = "ü"
    Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
ElseIf UstYaziKont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 9).Value = "x"
    Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
ElseIf UstYaziKont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf UstYaziKont = 3 Then
    GoTo ReseteGit
End If

ReseteGit:
'Son 20 raporu güncelle
If Rapor1Frame.Visible = True Then
    Call Son20RaporNo
End If

Call Rapor1FormunuResetle

ComboGetir.Clear
FarkGirisi.Visible = False

ResetAtla:

Say = Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
Else
    SiraSay = Range("E100000").End(xlUp)
End If
'Getir liste değerleri
For i = SiraSay To 1 Step -1
    With ComboGetir
        .AddItem (i)
    End With
Next i
GetirBos:

'Tümünü oluşturu işaretle
If YeniIslem <> 0 Then
    If Cells(YeniIslem, 6).Value = "ü" And Cells(YeniIslem, 7).Value = "ü" And Cells(YeniIslem, 8).Value = "ü" And Cells(YeniIslem, 9).Value = "ü" Then
        'Normal kaydet
        Cells(YeniIslem, 10).Value = "ü"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
        Cells(YeniIslem, 12).Value = UserName
    Else
        'Sorunlu kaydet
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
        Cells(YeniIslem, 12).Value = UserName
    End If
End If

'ThisWorkbook.Save

Out:
'Columns("CE:CF").EntireColumn.Hidden = True

ThisWorkbook.Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(7).Protect Password:="123"
ThisWorkbook.Worksheets(8).Protect Password:="123"
ThisWorkbook.Worksheets(11).Protect Password:="123"
ThisWorkbook.Protect "123"

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

End Sub

Sub Son20RaporNo()
Dim i As Integer
Dim Say As Long, j As Long, Cont As Long, Rno As Variant
Dim RefSatir As Long, Rapor1TarihBul As Range
'Verilen son 20 rapor numarasını göster

ThisWorkbook.Activate


'__________Rapor No Senkronizasyon 30.11.2021
    
Set WsRaporNo = ThisWorkbook.Worksheets(11)

'İlk satırda bulunan rapor1 numarası
Say = WsRaporNo.Range("E100000").End(xlUp).Row
If Say < 7 Then
    Say = 7
End If


RefSatir = 0
Set WsRaporNo = ThisWorkbook.Worksheets(11)
If Rapor1TarihiText.Value <> "" Then
    StrRaporTarihiGlobal = Right(CStr(Rapor1TarihiText.Value), 4)
Else
    StrRaporTarihiGlobal = Right(CStr(Format(Date, "dd.mm.yyyy")), 4)
End If
'MsgBox StrRaporTarihiGlobal
Set Rapor1TarihBul = WsRaporNo.Range("B6:B100000").Find(What:=StrRaporTarihiGlobal, SearchDirection:=xlNext, _
    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlPart)
If Not Rapor1TarihBul Is Nothing Then
    RefSatir = Rapor1TarihBul.Row
Else
    RefSatir = 7
End If
If RefSatir < 7 Then
    RefSatir = 7
End If
'MsgBox "RefSatir: " & RefSatir

Cont = 0
Rapor1No.Clear
For j = Say To RefSatir Step -1
    If Cont < 8 Then
        If WsRaporNo.Cells(j, 1).Value <> "" Then
            Cont = Cont + 1
            Rno = WsRaporNo.Cells(j, 1).Value
            With Rapor1No
                .AddItem (Rno)
            End With
        End If
    End If
Next j

'Sonraki satırlarda bulunan rapor1 numaraları
For i = 1 To 19
    Controls("Sonuc" & i).Visible = True
    Controls("LblSonuc" & i).Visible = True
    Controls("Rapor1No" & i).Visible = True
    Controls("LblRapor1No" & i).Visible = True
    Controls("Rapor1No" & i).Clear
    'Verilen son 20 rapor numarasını göster
    Say = WsRaporNo.Range("E100000").End(xlUp).Row
    If Say < 7 Then
        Say = 7
    End If
    Cont = 0
    For j = Say To RefSatir Step -1
        If Cont < 8 Then
            If WsRaporNo.Cells(j, 1).Value <> "" Then
                Cont = Cont + 1
                Rno = WsRaporNo.Cells(j, 1).Value
                With Controls("Rapor1No" & i)
                    .AddItem (Rno)
                End With
            End If
        End If
    Next j
Next i

'__________Rapor No Senkronizasyon 30.11.2021


End Sub

Private Sub Tutanak1Girisi_Click()
Dim i As Integer
Dim ctl As MSForms.Control

ThisWorkbook.Activate

Sonuc.Visible = False
LblSonuc.Visible = False
Rapor1No.Visible = False
LblRapor1No.Visible = False
LblSonucUst.Visible = False
LblRapor1NoUst.Visible = False
LblRapor1NoUst.Visible = False

For i = 1 To 19
    Controls("Sonuc" & i).Visible = False
    Controls("LblSonuc" & i).Visible = False
    Controls("Rapor1No" & i).Visible = False
    Controls("LblRapor1No" & i).Visible = False
Next i

EkleOge.Left = 571 '660 '(+)
KaldirOge.Left = 590 '684 '(-)

Rapor1Frame.Visible = False
Tutanak2Frame.Visible = False
UstYaziFrame.Visible = False

For Each ctl In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

'Tutanak1Girisi.BackColor = RGB(180, 210, 240)
'Tutanak1Girisi.ForeColor = RGB(30, 30, 30)
'
'If ComboGetir.Value <> "" Then
'    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
'    LblDuzeltme.ForeColor = RGB(30, 30, 30)
'End If

'OgeTurleriFrameUst.Caption = "Tutanak1 Girişi"

'Ekrana göre formun ayarlanması
If EkranKontrol = True Then
    
    core_report1_entry_UI.ScrollTop = 0
    core_report1_entry_UI.ScrollHeight = 0
    core_report1_entry_UI.ScrollBars = fmScrollBarsNone

    'Formun görünümü
    AltMenuFrame.Top = 462 '444 '299
    TasiyiciFrame.Height = 486
    core_report1_entry_UI.Height = 546 '556 '497 '352
    core_report1_entry_UI.Width = 1024
    
Else
    'Formun görünümü
    AltMenuFrame.Top = 462 '444 '299
    TasiyiciFrame.Height = 486
    core_report1_entry_UI.Height = 556 '497 '352
End If

EsasFrame.ZOrder msoBringToFront
OgeTurleriFrameUst.ZOrder msoBringToFront
ScrollFrame.ZOrder msoBringToFront
AltMenuFrame.ZOrder msoBringToFront

'___________

Call RaporlamaGirisiPro

End Sub
Private Sub RaporlamaGirisiPro() '_Click()
Dim i As Integer
Dim Say As Long, j As Long, Cont As Long, Rno As Variant
Dim ctl As MSForms.Control

ThisWorkbook.Activate

EkleOge.Left = 741 '660 '(+)
KaldirOge.Left = 760 '684 '(-)

Rapor1Frame.Visible = True

LblSonucUst.Visible = True
LblSonuc.Visible = True
Sonuc.Visible = True

LblRapor1NoUst.Visible = True
LblRapor1No.Visible = True
Rapor1No.Visible = True

Call Son20RaporNo

Tutanak2Frame.Visible = False
UstYaziFrame.Visible = False

For Each ctl In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

'RaporlamaGirisi.BackColor = RGB(180, 210, 240)
'RaporlamaGirisi.ForeColor = RGB(30, 30, 30)
'
'If ComboGetir.Value <> "" Then
'    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
'    LblDuzeltme.ForeColor = RGB(30, 30, 30)
'End If

Tutanak1Girisi.BackColor = RGB(180, 210, 240)
Tutanak1Girisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If


'OgeTurleriFrameUst.Caption = "Tutanak1 & Rapor1 Girişi"

'Ekrana göre formun ayarlanması
If EkranKontrol = True Then

    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + 6
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + 6
    core_report1_entry_UI.Height = 485 '556 + Rapor1Frame.Height + 6

    core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + 6 - 30
    core_report1_entry_UI.ScrollTop = 0
    core_report1_entry_UI.Width = 1024 + 12
    
Else
    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + 6
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + 6
    core_report1_entry_UI.Height = 556 + Rapor1Frame.Height + 6
End If

Rapor1Frame.ZOrder msoBringToFront

End Sub

Private Sub Tutanak2Girisi_Click()
Dim ctl As MSForms.Control

Call RaporlamaGirisiPro

Tutanak2Frame.Visible = True



UstYaziFrame.Visible = False

For Each ctl In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

Tutanak2Girisi.BackColor = RGB(180, 210, 240)
Tutanak2Girisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If

'OgeTurleriFrameUst.Caption = "Tutanak1 & Rapor1 Girişi"

'Ekrana göre formun ayarlanması
If EkranKontrol = True Then

    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
    core_report1_entry_UI.Height = 485 '556 + Rapor1Frame.Height + Tutanak2Frame.Height + 12

    core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + Tutanak2Frame.Height + 12 - 30
    core_report1_entry_UI.ScrollTop = 0
    core_report1_entry_UI.Width = 1024 + 12
    
Else
    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
    core_report1_entry_UI.Height = 556 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
End If

Tutanak2Frame.ZOrder msoBringToFront

End Sub

Private Sub UstYaziGirisi_Click()
Dim ctl As MSForms.Control

Tutanak2Girisi_Click

UstYaziFrame.Visible = True


For Each ctl In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

UstYaziGirisi.BackColor = RGB(180, 210, 240)
UstYaziGirisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If

'OgeTurleriFrameUst.Caption = "Tutanak1 & Rapor1 Girişi"

'Ekrana göre formun ayarlanması
If EkranKontrol = True Then

    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
    core_report1_entry_UI.Height = 485 '556 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18

    core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18 - 30
    core_report1_entry_UI.ScrollTop = 0
    core_report1_entry_UI.Width = 1024 + 12
    
Else
    'Formun görünümü
    AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
    TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
    core_report1_entry_UI.Height = 556 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
End If

UstYaziFrame.ZOrder msoBringToFront

End Sub

Sub ColorChangerGenel()

'Düzeltme
If LblDuzeltme.BackColor <> RGB(180, 210, 240) Then
    If LblDuzeltme.BackColor <> RGB(225, 235, 245) Then
        LblDuzeltme.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
        LblDuzeltme.ForeColor = RGB(30, 30, 30)
    End If
End If
'Taslak
If LblTaslak.BackColor <> RGB(180, 210, 240) Then
    If LblTaslak.BackColor <> RGB(225, 235, 245) Then
        LblTaslak.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
        LblTaslak.ForeColor = RGB(30, 30, 30)
    End If
End If
'Sil
If LblSil.BackColor <> RGB(225, 235, 245) Then
    LblSil.BackColor = RGB(225, 235, 245)
    LblSil.ForeColor = RGB(30, 30, 30)
End If
'Tutanak1
If Tutanak1Girisi.BackColor <> RGB(180, 210, 240) Then
    If Tutanak1Girisi.BackColor <> RGB(225, 235, 245) Then
        Tutanak1Girisi.BackColor = RGB(225, 235, 245)
        Tutanak1Girisi.ForeColor = RGB(30, 30, 30)
    End If
End If

''Rapor
'If RaporlamaGirisi.BackColor <> RGB(180, 210, 240) Then
'    If RaporlamaGirisi.BackColor <> RGB(225, 235, 245) Then
'        RaporlamaGirisi.BackColor = RGB(225, 235, 245)
'        RaporlamaGirisi.ForeColor = RGB(30, 30, 30)
'    End If
'End If

'Tutanak2
If Tutanak2Girisi.BackColor <> RGB(180, 210, 240) Then
    If Tutanak2Girisi.BackColor <> RGB(225, 235, 245) Then
        Tutanak2Girisi.BackColor = RGB(225, 235, 245)
        Tutanak2Girisi.ForeColor = RGB(30, 30, 30)
    End If
End If
'Üst yazı
If UstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    If UstYaziGirisi.BackColor <> RGB(225, 235, 245) Then
        UstYaziGirisi.BackColor = RGB(225, 235, 245)
        UstYaziGirisi.ForeColor = RGB(30, 30, 30)
    End If
End If
'Help Documents
If Yardim.BackColor <> RGB(225, 235, 245) Then
    Yardim.BackColor = RGB(225, 235, 245)
    Yardim.ForeColor = RGB(30, 30, 30)
End If
'Kapat
If Kapat.BackColor <> RGB(225, 235, 245) Then
    Kapat.BackColor = RGB(225, 235, 245)
    Kapat.ForeColor = RGB(30, 30, 30)
End If
'MaxiMini
If MaxiMini.BackColor <> RGB(225, 235, 245) Then
    MaxiMini.BackColor = RGB(225, 235, 245)
    MaxiMini.ForeColor = RGB(30, 30, 30)
End If
'Kaydet
If Kaydet.BackColor <> RGB(225, 235, 245) Then
    Kaydet.BackColor = RGB(225, 235, 245)
    Kaydet.ForeColor = RGB(30, 30, 30)
End If



'Belge tarihi
If BelgeTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    BelgeTarihiLabel.BackColor = RGB(254, 254, 254)
    BelgeTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'OtomatikOption
If OtomatikOption.BackColor <> RGB(254, 254, 254) Then
    OtomatikOption.BackColor = RGB(254, 254, 254)
    OtomatikOption.ForeColor = RGB(70, 70, 70)
End If
'ManuelOption
If ManuelOption.BackColor <> RGB(254, 254, 254) Then
    ManuelOption.BackColor = RGB(254, 254, 254)
    ManuelOption.ForeColor = RGB(70, 70, 70)
End If
'IlIlceEkleKaldirLabel
If IlIlceEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    IlIlceEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    IlIlceEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'IlIlceEkleKaldirLabel2
If IlIlceEkleKaldirLabel2.BackColor <> RGB(254, 254, 254) Then
    IlIlceEkleKaldirLabel2.BackColor = RGB(254, 254, 254)
    IlIlceEkleKaldirLabel2.ForeColor = RGB(70, 70, 70)
End If
'MuhatapEkleKaldirLabelGelen
If MuhatapEkleKaldirLabelGelen.BackColor <> RGB(254, 254, 254) Then
    MuhatapEkleKaldirLabelGelen.BackColor = RGB(254, 254, 254)
    MuhatapEkleKaldirLabelGelen.ForeColor = RGB(70, 70, 70)
End If
'GonderenEkleKaldirLabel
If GonderenEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    GonderenEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    GonderenEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'GelisTarihiLabel
If GelisTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    GelisTarihiLabel.BackColor = RGB(254, 254, 254)
    GelisTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'Tutanak1TarihiLabel
If Tutanak1TarihiLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak1TarihiLabel.BackColor = RGB(254, 254, 254)
    Tutanak1TarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'DLEvetOption
If DLEvetOption.BackColor <> RGB(254, 254, 254) Then
    DLEvetOption.BackColor = RGB(254, 254, 254)
    DLEvetOption.ForeColor = RGB(70, 70, 70)
End If
'DLHayirOption
If DLHayirOption.BackColor <> RGB(254, 254, 254) Then
    DLHayirOption.BackColor = RGB(254, 254, 254)
    DLHayirOption.ForeColor = RGB(70, 70, 70)
End If
'Tutanak1Imza1EkleKaldirLabel
If Tutanak1Imza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak1Imza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    Tutanak1Imza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'Tutanak1Imza2EkleKaldirLabel
If Tutanak1Imza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak1Imza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    Tutanak1Imza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If

'LblDosyaNoGetir
If LblDosyaNoGetir.BackColor <> RGB(254, 254, 254) Then
    LblDosyaNoGetir.BackColor = RGB(254, 254, 254)
    LblDosyaNoGetir.ForeColor = RGB(70, 70, 70)
End If

'OgeEkleKaldirLabel
If OgeEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    OgeEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    OgeEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'OgeDegeriEkleKaldirLabel
If OgeDegeriEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    OgeDegeriEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    OgeDegeriEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'EkleOge
If EkleOge.BackColor <> RGB(254, 254, 254) Then
    EkleOge.BackColor = RGB(254, 254, 254)
    EkleOge.ForeColor = RGB(70, 70, 70)
End If
'KaldirOge
If KaldirOge.BackColor <> RGB(254, 254, 254) Then
    KaldirOge.BackColor = RGB(254, 254, 254)
    KaldirOge.ForeColor = RGB(70, 70, 70)
End If
'Rapor1TarihiLabel
If Rapor1TarihiLabel.BackColor <> RGB(254, 254, 254) Then
    Rapor1TarihiLabel.BackColor = RGB(254, 254, 254)
    Rapor1TarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'RaporImza1EkleKaldirLabel
If RaporImza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    RaporImza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    RaporImza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'RaporImza2EkleKaldirLabel
If RaporImza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    RaporImza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    RaporImza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'Tutanak2TarihiLabel
If Tutanak2TarihiLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak2TarihiLabel.BackColor = RGB(254, 254, 254)
    Tutanak2TarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'MuhatapEkleKaldirLabelGiden
If MuhatapEkleKaldirLabelGiden.BackColor <> RGB(254, 254, 254) Then
    MuhatapEkleKaldirLabelGiden.BackColor = RGB(254, 254, 254)
    MuhatapEkleKaldirLabelGiden.ForeColor = RGB(70, 70, 70)
End If
'GonderilenEkleKaldirLabel
If GonderilenEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    GonderilenEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    GonderilenEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'Tutanak2Imza1EkleKaldirLabel
If Tutanak2Imza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak2Imza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    Tutanak2Imza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'Tutanak2Imza2EkleKaldirLabel
If Tutanak2Imza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak2Imza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    Tutanak2Imza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'UstYaziTarihiLabel
If UstYaziTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    UstYaziTarihiLabel.BackColor = RGB(254, 254, 254)
    UstYaziTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'UstYaziImza1EkleKaldirLabel
If UstYaziImza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    UstYaziImza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    UstYaziImza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'UstYaziImza2EkleKaldirLabel
If UstYaziImza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    UstYaziImza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    UstYaziImza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'FarkGirisi
If FarkGirisi.BackColor <> RGB(254, 254, 254) Then
    FarkGirisi.BackColor = RGB(254, 254, 254)
    FarkGirisi.ForeColor = RGB(70, 70, 70)
End If

End Sub

Private Sub LblSil_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LblSil.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
LblSil.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LblDuzeltme_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblDuzeltme.BackColor <> RGB(180, 210, 240) Then
    LblDuzeltme.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(255, 255, 255)
End If
End Sub
Private Sub LblTaslak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblTaslak.BackColor <> RGB(180, 210, 240) Then
    LblTaslak.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblTaslak.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Tutanak1Girisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Tutanak1Girisi.BackColor <> RGB(180, 210, 240) Then
    Tutanak1Girisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Tutanak1Girisi.ForeColor = RGB(255, 255, 255)
End If
End Sub
'Private Sub RaporlamaGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'Call ColorChangerGenel
'If RaporlamaGirisi.BackColor <> RGB(180, 210, 240) Then
'    RaporlamaGirisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
'    RaporlamaGirisi.ForeColor = RGB(255, 255, 255)
'End If
'End Sub
Private Sub Tutanak2Girisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Tutanak2Girisi.BackColor <> RGB(180, 210, 240) Then
    Tutanak2Girisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Tutanak2Girisi.ForeColor = RGB(255, 255, 255)
End If
End Sub
Private Sub UstYaziGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If UstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    UstYaziGirisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    UstYaziGirisi.ForeColor = RGB(255, 255, 255)
End If
End Sub
Private Sub Kaydet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kaydet.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kaydet.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub MaxiMini_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
MaxiMini.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
MaxiMini.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub OtomatikOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
OtomatikOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
OtomatikOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub ManuelOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
ManuelOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
ManuelOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub DLEvetOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
DLEvetOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
DLEvetOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub DLHayirOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
DLHayirOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
DLHayirOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub BelgeTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
BelgeTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
BelgeTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub GelisTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
GelisTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
GelisTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak1TarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak1TarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Tutanak1TarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LblDosyaNoGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LblDosyaNoGetir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
LblDosyaNoGetir.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub DosyaNoText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GelenBelgeFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub Rapor1TarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Rapor1TarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Rapor1TarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak2TarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak2TarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Tutanak2TarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub UstYaziTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
UstYaziTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
UstYaziTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub EkleOge_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
EkleOge.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
EkleOge.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub KaldirOge_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
KaldirOge.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
KaldirOge.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LblRapor1NoUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblAciklamaUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub IlIlceEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
IlIlceEkleKaldirLabel.BackColor = RGB(60, 100, 180)
IlIlceEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub IlIlceEkleKaldirLabel2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
IlIlceEkleKaldirLabel2.BackColor = RGB(60, 100, 180)
IlIlceEkleKaldirLabel2.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub MuhatapEkleKaldirLabelGelen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
MuhatapEkleKaldirLabelGelen.BackColor = RGB(60, 100, 180)
MuhatapEkleKaldirLabelGelen.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub MuhatapEkleKaldirLabelGiden_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
MuhatapEkleKaldirLabelGiden.BackColor = RGB(60, 100, 180)
MuhatapEkleKaldirLabelGiden.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub GonderenEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
GonderenEkleKaldirLabel.BackColor = RGB(60, 100, 180)
GonderenEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub GonderilenEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
GonderilenEkleKaldirLabel.BackColor = RGB(60, 100, 180)
GonderilenEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub OgeEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
OgeEkleKaldirLabel.BackColor = RGB(60, 100, 180)
OgeEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub OgeDegeriEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
OgeDegeriEkleKaldirLabel.BackColor = RGB(60, 100, 180)
OgeDegeriEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub FarkGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FarkGirisi.BackColor = RGB(60, 100, 180)
FarkGirisi.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub LblOgeTuru_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub OgeTuru_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(OgeTuru) 'Open scrollable with mouse
End Sub
Private Sub LblOgeDegeri_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub OgeDegeri_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(OgeDegeri) 'Open scrollable with mouse
End Sub
Private Sub LblAdet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Adet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblOgeIdNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub OgeIdNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblAciklama_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Aciklama_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblSonuc_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Sonuc_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblRapor1No_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Rapor1No_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub



'SCROLABLE COMBOBOXES (Öğe Alanı)
Private Sub OgeTuru1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru1) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru2) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru3) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru4) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru5) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru6) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru7) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru8) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru9) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru10) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru11) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru12) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru13) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru14) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru15) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru16) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru17) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru18) 'Open scrollable with mouse
End Sub
Private Sub OgeTuru19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeTuru19) 'Open scrollable with mouse
End Sub

Private Sub OgeDegeri1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri1) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri2) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri3) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri4) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri5) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri6) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri7) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri8) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri9) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri10) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri11) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri12) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri13) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri14) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri15) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri16) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri17) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri18) 'Open scrollable with mouse
End Sub
Private Sub OgeDegeri19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(OgeDegeri19) 'Open scrollable with mouse
End Sub




Private Sub BaslikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub EsasFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub ScrollFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call Move_SetScrollHook(Me.ScrollFrame, Threshold, ScrollTakip)
End Sub
Private Sub Rapor1Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub Tutanak2Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TasiyiciFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub

Private Sub ComboGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboGetir) 'Open scrollable with mouse
End Sub
Private Sub BelgeTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TemaNoText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblTemaNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GelisTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak1TarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Rapor1TarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak2TarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GelenMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(GelenMuhatapTemasi) 'Open scrollable with mouse
End Sub
Private Sub LblGelenMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GidenMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(GidenMuhatapTemasi) 'Open scrollable with mouse
End Sub
Private Sub LblGidenMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GonderenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(GonderenBirim) 'Open scrollable with mouse
End Sub
Private Sub LblGonderenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GonderilenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(GonderilenBirim) 'Open scrollable with mouse
End Sub
Private Sub LblGonderilenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblOgeTuruUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblOgeDegeriUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak1Imza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak1Imza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tutanak1Imza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak1Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tutanak1Imza1) 'Open scrollable with mouse
End Sub
Private Sub LblTutanak1Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak1Imza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak1Imza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tutanak1Imza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak1Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tutanak1Imza2) 'Open scrollable with mouse
End Sub
Private Sub LblTutanak1Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub RaporImza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporImza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
RaporImza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub RaporImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(RaporImza1) 'Open scrollable with mouse
End Sub
Private Sub LblRaporImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub RaporImza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporImza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
RaporImza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub RaporImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(RaporImza2) 'Open scrollable with mouse
End Sub
Private Sub LblRaporImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak2Imza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak2Imza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tutanak2Imza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak2Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tutanak2Imza1) 'Open scrollable with mouse
End Sub
Private Sub LblTutanak2Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak2Imza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak2Imza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tutanak2Imza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tutanak2Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tutanak2Imza2) 'Open scrollable with mouse
End Sub
Private Sub LblTutanak2Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziImza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
UstYaziImza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
UstYaziImza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub UstYaziImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(UstYaziImza1) 'Open scrollable with mouse
End Sub
Private Sub LblUstYaziImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziImza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
UstYaziImza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
UstYaziImza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub UstYaziImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(UstYaziImza2) 'Open scrollable with mouse
End Sub
Private Sub LblUstYaziImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub OgeTurleriFrameUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub

Private Sub DokumListesi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(DokumListesi) 'Open scrollable with mouse
End Sub
Private Sub LblDokumListesi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIlce_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblBelgeNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


'SCROLABLE COMBOBOXES
'Il
Private Sub Il_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Il) 'Open scrollable with mouse
End Sub
'Ilce
Private Sub Ilce_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Ilce) 'Open scrollable with mouse
End Sub
'IlGiden
Private Sub IlGiden_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(IlGiden) 'Open scrollable with mouse
End Sub
'IlceGiden
Private Sub IlceGiden_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(IlceGiden) 'Open scrollable with mouse
End Sub

'İkinci bölüm
'TemaTipi
Private Sub TemaTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(TemaTipi) 'Open scrollable with mouse
End Sub
''GelenPaketTipi
'Private Sub GelenPaketTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Call SetComboBoxHook(GelenPaketTipi) 'Open scrollable with mouse
'End Sub
''GelisSekli
'Private Sub GelisSekli_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Call SetComboBoxHook(GelisSekli) 'Open scrollable with mouse
'End Sub
'GelenBelgeSayfa
Private Sub GelenBelgeSayfa_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(GelenBelgeSayfa) 'Open scrollable with mouse
End Sub
'GidenPaketTipi
Private Sub GidenPaketTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(GidenPaketTipi) 'Open scrollable with mouse
End Sub
'GidenPaketAdedi
Private Sub GidenPaketAdedi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(GidenPaketAdedi) 'Open scrollable with mouse
End Sub
'IlgiYaziFotokopisi
Private Sub IlgiYaziFotokopisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(IlgiYaziFotokopisi) 'Open scrollable with mouse
End Sub

Private Sub OtomatikOption_Click()
Dim IlBul As Range, IlceBul As Range, IlDegeri As String, IlceDegeri As String, IlEsleyicisi As Integer
Dim Makam As String, Yil As String, EvrakNo As String, SlashFinder As Integer, CharLen As Integer
Dim TemaYil As String, TemaSayi As String, TireFinder As Integer, i As Integer

    On Error GoTo Son
    
    ThisWorkbook.Activate
    
    If OtomatikOption.Value = True Then
        'Tema kodu
            'Il
            Set IlBul = ThisWorkbook.Worksheets(2).Columns("F").Find(What:=Il.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlBul Is Nothing Then
                IlDegeri = ThisWorkbook.Worksheets(2).Range("E" & IlBul.Row)
                If IlDegeri < 10 Then
                    IlDegeri = 0 & IlDegeri
                End If
                IlEsleyicisi = ThisWorkbook.Worksheets(2).Range("C" & IlBul.Row).Value
            Else
                IlDegeri = ""
            End If
            'Ilce
            Set IlceBul = ThisWorkbook.Worksheets(2).Columns(IlEsleyicisi + 6).Find(What:=Ilce.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlceBul Is Nothing Then
                IlceDegeri = ThisWorkbook.Worksheets(2).Range("D" & IlceBul.Row)
                If IlceDegeri < 10 Then
                    IlceDegeri = 0 & IlceDegeri
                End If
            Else
                IlceDegeri = ""
            End If
            'Muhatap teması
            Makam = ""
            If InStr(TemaTipi.Value, "Organization A") <> 0 Then
                Makam = "A"
            ElseIf InStr(TemaTipi.Value, "Organization B") <> 0 Then
                Makam = "B"
            ElseIf InStr(TemaTipi.Value, "Organization C") <> 0 Then
                Makam = "C"
            ElseIf InStr(TemaTipi.Value, "Organization D") <> 0 Then
                Makam = "D"
            ElseIf InStr(TemaTipi.Value, "Organization E") <> 0 Then
                Makam = "E"
            End If
            
            'Belge no
            For i = 1 To 50
                BelgeNoText.Value = Replace(BelgeNoText.Value, " ", "")
            Next i
            EvrakNo = ""
            If InStr(BelgeNoText.Value, "/") <> 0 And InStr(BelgeNoText.Value, "-") <> 0 Then
                MsgBox "The document number contains both a slash (/) and a hyphen (-), so the theme number cannot be generated automatically." & vbNewLine & _
                        "Please enter the incoming document number in one of the following formats: 2018/1234567, 2018-1234567, or 1234567.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                TemaNoText.Value = ""
                GoTo Son
            End If
            If InStr(BelgeNoText.Value, "/") <> 0 Or InStr(BelgeNoText.Value, "-") <> 0 Then 'Belge numarasında slash varsa
                SlashFinder = InStr(BelgeNoText.Value, "/")
                TireFinder = InStr(BelgeNoText.Value, "-")
                CharLen = Len(BelgeNoText.Value)
                If SlashFinder = 5 Then
                    TemaYil = Mid(BelgeNoText.Value, 1, SlashFinder - 1)
                    TemaSayi = Mid(BelgeNoText.Value, SlashFinder + 1, CharLen - SlashFinder + 1)
                    If IsNumeric(TemaYil) = False Or IsNumeric(TemaSayi) = False Then
                        MsgBox "The document number contains non-numeric characters other than a slash (/) or a hyphen (-), so the theme number cannot be generated automatically." & vbNewLine & _
                                "Please enter the incoming document number in one of the following formats: 2018/1234567, 2018-1234567, or 1234567.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                        TemaNoText.Value = ""
                        GoTo Son
                    End If
                    Yil = Right(TemaYil, 2)
                    EvrakNo = TemaSayi
                ElseIf TireFinder = 5 Then
                    TemaYil = Mid(BelgeNoText.Value, 1, TireFinder - 1)
                    TemaSayi = Mid(BelgeNoText.Value, TireFinder + 1, CharLen - TireFinder + 1)
                    If IsNumeric(TemaYil) = False Or IsNumeric(TemaSayi) = False Then
                        MsgBox "The document number contains non-numeric characters other than a slash (/) or a hyphen (-), so the theme number cannot be generated automatically." & vbNewLine & _
                                "Please enter the incoming document number in one of the following formats: 2018/1234567, 2018-1234567, or 1234567.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                        TemaNoText.Value = ""
                        GoTo Son
                    End If
                    Yil = Right(TemaYil, 2)
                    EvrakNo = TemaSayi
                Else
                    MsgBox "The document number contains non-numeric characters other than a slash (/) or a hyphen (-), so the theme number cannot be generated automatically." & vbNewLine & _
                            "Please enter the incoming document number in one of these formats: 2018/1234567, 2018-1234567, or 1234567.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    TemaNoText.Value = ""
                    GoTo Son
                End If
            Else
                'Yıl
                Yil = ""
                Yil = Right(BelgeTarihiText, 2)
                'Belge numarası
                EvrakNo = ""
                EvrakNo = Right(BelgeNoText.Value, 5)
            End If
            'Belge numarasının başına sıfır ekleme
            If Len(EvrakNo) = 1 Then
                EvrakNo = 0 & 0 & 0 & 0 & EvrakNo
            ElseIf Len(EvrakNo) = 2 Then
                EvrakNo = 0 & 0 & 0 & EvrakNo
            ElseIf Len(EvrakNo) = 3 Then
                EvrakNo = 0 & 0 & EvrakNo
            ElseIf Len(EvrakNo) = 4 Then
                EvrakNo = 0 & EvrakNo
            ElseIf Len(EvrakNo) >= 5 Then
                EvrakNo = Right(EvrakNo, 5)
            End If
            'Tema no oluştur
            If IlDegeri <> "" And IlceDegeri <> "" And Makam <> "" And Yil <> "" And EvrakNo <> "" Then
                If InStr(TemaTipi.Value, "Organization A") <> 0 Then
                    IlceDegeri = "00"
                End If
                TemaNoText.Value = Makam & Yil & IlDegeri & IlceDegeri & EvrakNo
            Else
                TemaNoText.Value = ""
            End If
Son:
        TemaNoText.Locked = True
    End If

End Sub
Private Sub ManuelOption_Click()
    If ManuelOption.Value = True Then
        TemaNoText.Locked = False
        TemaNoText.Value = ""
    End If
End Sub

Private Sub DLEvetOption_Click()
    If DLEvetOption.Value = True Then
        DokumListesi.Enabled = True
    End If
End Sub
Private Sub DLHayirOption_Click()
    If DLHayirOption.Value = True Then
        DokumListesi.Enabled = False
        Me.DokumListesi.Value = Null
    End If
End Sub

Private Sub ComboGetir_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ComboGetir.DropDown

End Sub

Private Sub ComboGetir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    'Enter
    If KeyCode = vbKeyReturn Then
        '
    End If
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        '
    End If
    If KeyCode = vbKeyDown Then
        Il.SetFocus
    End If
    'Sağa ve sola
    If KeyCode = vbKeyLeft Then
        '
    End If
    If KeyCode = vbKeyRight Then
        Il.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If ComboGetir.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboGetir.ListIndex = ComboGetir.ListIndex
            End If
        Case 40 'Down
            If ComboGetir.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboGetir.ListIndex = ComboGetir.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub ComboGetir_Change()

If ComboGetir.Value = "" Then
    LblDuzeltme.BackColor = RGB(225, 235, 245)  'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub IlGiden_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.IlGiden.DropDown
IlGiden.BackColor = RGB(255, 255, 255)
IlGiden.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub IlGiden_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If IlGiden.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlGiden.ListIndex = IlGiden.ListIndex - 1
            End If
            Me.IlGiden.DropDown
            
        Case 40 'Aşağı
            If IlGiden.ListIndex = IlGiden.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlGiden.ListIndex = IlGiden.ListIndex + 1
            End If
            Me.IlGiden.DropDown
    End Select
    Abort = False

    
End Sub

Private Sub IlGiden_Change()

If IlGiden.ListIndex = -1 And IlGiden.Value <> "" Then
   IlGiden.Value = ""
   GoTo Son
End If

'Ilçe seçimlerini İl seçimine göre göster.
On Error GoTo Bos
IlceGiden.RowSource = Replace(IlGiden.Value, " ", "_")

'IlGiden.DropDown
GoTo Son

Bos:
IlceGiden.RowSource = ""

Son:

If IlGiden.Value <> "" Then
    IlGiden.SelStart = 0
    IlGiden.SelLength = Len(IlGiden.Value)
End If

IlGiden.DropDown
If IlGiden.BackColor = RGB(60, 100, 180) Then
IlGiden.BackColor = RGB(255, 255, 255)
IlGiden.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub IlceGiden_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.IlceGiden.DropDown

IlceGiden.BackColor = RGB(255, 255, 255)
IlceGiden.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub IlceGiden_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If IlceGiden.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlceGiden.ListIndex = IlceGiden.ListIndex - 1
            End If
            Me.IlceGiden.DropDown
            
        Case 40 'Aşağı
            If IlceGiden.ListIndex = IlceGiden.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlceGiden.ListIndex = IlceGiden.ListIndex + 1
            End If
            Me.IlceGiden.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub IlceGiden_Change()

If IlceGiden.ListIndex = -1 And IlceGiden.Value <> "" Then
   IlceGiden.Value = ""
   GoTo Son
End If

If IlceGiden.Value <> "" Then
    IlceGiden.SelStart = 0
    IlceGiden.SelLength = Len(IlceGiden.Value)
End If

Son:

IlceGiden.DropDown
If IlceGiden.BackColor = RGB(60, 100, 180) Then
IlceGiden.BackColor = RGB(255, 255, 255)
IlceGiden.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub Il_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Il.DropDown
Il.BackColor = RGB(255, 255, 255)
Il.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Il_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Il.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Il.ListIndex = Il.ListIndex - 1
            End If
            Me.Il.DropDown
            
        Case 40 'Aşağı
            If Il.ListIndex = Il.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Il.ListIndex = Il.ListIndex + 1
            End If
            Me.Il.DropDown
    End Select
    Abort = False

    
End Sub

Private Sub Il_Change()

If Il.ListIndex = -1 And Il.Value <> "" Then
   Il.Value = ""
   GoTo Son
End If

'Ilçe seçimlerini İl seçimine göre göster.
On Error GoTo Bos
Ilce.RowSource = Replace(Il.Value, " ", "_")
'Il.DropDown
GoTo Son

Bos:
Ilce.RowSource = ""

Son:

If Il.Value <> "" Then
    Il.SelStart = 0
    Il.SelLength = Len(Il.Value)
End If

Il.DropDown
If Il.BackColor = RGB(60, 100, 180) Then
Il.BackColor = RGB(255, 255, 255)
Il.ForeColor = RGB(30, 30, 30)
End If

If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If

End Sub


Private Sub Ilce_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Ilce.DropDown

Ilce.BackColor = RGB(255, 255, 255)
Ilce.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Ilce_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Ilce.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Ilce.ListIndex = Ilce.ListIndex - 1
            End If
            Me.Ilce.DropDown
            
        Case 40 'Aşağı
            If Ilce.ListIndex = Ilce.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Ilce.ListIndex = Ilce.ListIndex + 1
            End If
            Me.Ilce.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Ilce_Change()

If Ilce.ListIndex = -1 And Ilce.Value <> "" Then
   Ilce.Value = ""
   GoTo Son
End If

If Ilce.Value <> "" Then
    Ilce.SelStart = 0
    Ilce.SelLength = Len(Ilce.Value)
End If

Son:

Ilce.DropDown
If Ilce.BackColor = RGB(60, 100, 180) Then
Ilce.BackColor = RGB(255, 255, 255)
Ilce.ForeColor = RGB(30, 30, 30)
End If

If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If

End Sub

Private Sub BelgeTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        BelgeTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        BelgeTarihiText.Value = ""
    End If

End Sub

Private Sub BelgeTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

BelgeTarihiText.BackColor = RGB(255, 255, 255)
BelgeTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub BelgeTarihiText_Change()
If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If
End Sub

Private Sub BelgeTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    BelgeTarihiText.Value = CalTarih
    BelgeTarihiText.Value = Format(BelgeTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

BelgeTarihiText.BackColor = RGB(255, 255, 255)
BelgeTarihiText.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub BelgeNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

BelgeNoText.BackColor = RGB(255, 255, 255)
BelgeNoText.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub BelgeNoText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    
End Sub
Private Sub BelgeNoText_Change()

BelgeNoText.BackColor = RGB(255, 255, 255)
BelgeNoText.ForeColor = RGB(30, 30, 30)

If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If

End Sub

Private Sub TemaTipi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TemaTipi.DropDown
TemaTipi.BackColor = RGB(255, 255, 255)
TemaTipi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TemaTipi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If TemaTipi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TemaTipi.ListIndex = TemaTipi.ListIndex - 1
            End If
            Me.TemaTipi.DropDown
            
        Case 40 'Aşağı
            If TemaTipi.ListIndex = TemaTipi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TemaTipi.ListIndex = TemaTipi.ListIndex + 1
            End If
            Me.TemaTipi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TemaTipi_Change()

If TemaTipi.ListIndex = -1 And TemaTipi.Value <> "" Then
   TemaTipi.Value = ""
   GoTo Son
End If

If TemaTipi.Value <> "" Then
    TemaTipi.SelStart = 0
    TemaTipi.SelLength = Len(TemaTipi.Value)
End If


Son:

TemaTipi.DropDown
If TemaTipi.BackColor = RGB(60, 100, 180) Then
TemaTipi.BackColor = RGB(255, 255, 255)
TemaTipi.ForeColor = RGB(30, 30, 30)
End If

If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If

End Sub

Private Sub TemaNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TemaNoText.BackColor = RGB(255, 255, 255)
TemaNoText.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub TemaNoText_Change()
TemaNoText.BackColor = RGB(255, 255, 255)
TemaNoText.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub TemaNoText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    
End Sub
Private Sub GelenMuhatapTemasi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GelenMuhatapTemasi.DropDown
GelenMuhatapTemasi.BackColor = RGB(255, 255, 255)
GelenMuhatapTemasi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelenMuhatapTemasi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GelenMuhatapTemasi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenMuhatapTemasi.ListIndex = GelenMuhatapTemasi.ListIndex - 1
            End If
            Me.GelenMuhatapTemasi.DropDown
            
        Case 40 'Aşağı
            If GelenMuhatapTemasi.ListIndex = GelenMuhatapTemasi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenMuhatapTemasi.ListIndex = GelenMuhatapTemasi.ListIndex + 1
            End If
            Me.GelenMuhatapTemasi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GelenMuhatapTemasi_Change()

If GelenMuhatapTemasi.ListIndex = -1 And GelenMuhatapTemasi.Value <> "" Then
   GelenMuhatapTemasi.Value = ""
   GoTo Son
End If

If GelenMuhatapTemasi.Value <> "" Then
    GelenMuhatapTemasi.SelStart = 0
    GelenMuhatapTemasi.SelLength = Len(GelenMuhatapTemasi.Value)
End If

Son:

GelenMuhatapTemasi.DropDown
If GelenMuhatapTemasi.BackColor = RGB(60, 100, 180) Then
GelenMuhatapTemasi.BackColor = RGB(255, 255, 255)
GelenMuhatapTemasi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub GonderenBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GonderenBirim.DropDown
GonderenBirim.BackColor = RGB(255, 255, 255)
GonderenBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GonderenBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GonderenBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GonderenBirim.ListIndex = GonderenBirim.ListIndex - 1
            End If
            Me.GonderenBirim.DropDown
            
        Case 40 'Aşağı
            If GonderenBirim.ListIndex = GonderenBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GonderenBirim.ListIndex = GonderenBirim.ListIndex + 1
            End If
            Me.GonderenBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GonderenBirim_Change()

If GonderenBirim.ListIndex = -1 And GonderenBirim.Value <> "" Then
   GonderenBirim.Value = ""
   GoTo Son
End If

If GonderenBirim.Value <> "" Then
    GonderenBirim.SelStart = 0
    GonderenBirim.SelLength = Len(GonderenBirim.Value)
End If

Son:

GonderenBirim.DropDown
If GonderenBirim.BackColor = RGB(60, 100, 180) Then
GonderenBirim.BackColor = RGB(255, 255, 255)
GonderenBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub
Private Sub GelisTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        GelisTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        GelisTarihiText.Value = ""
    End If
    
End Sub

Private Sub GelisTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

GelisTarihiText.BackColor = RGB(255, 255, 255)
GelisTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelisTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    GelisTarihiText.Value = CalTarih
    GelisTarihiText.Value = Format(GelisTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

GelisTarihiText.BackColor = RGB(255, 255, 255)
GelisTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelenPaketTipi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GelenPaketTipi.DropDown
GelenPaketTipi.BackColor = RGB(255, 255, 255)
GelenPaketTipi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelenPaketTipi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GelenPaketTipi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenPaketTipi.ListIndex = GelenPaketTipi.ListIndex - 1
            End If
            Me.GelenPaketTipi.DropDown
            
        Case 40 'Aşağı
            If GelenPaketTipi.ListIndex = GelenPaketTipi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenPaketTipi.ListIndex = GelenPaketTipi.ListIndex + 1
            End If
            Me.GelenPaketTipi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GelenPaketTipi_Change()

If GelenPaketTipi.ListIndex = -1 And GelenPaketTipi.Value <> "" Then
   GelenPaketTipi.Value = ""
   GoTo Son
End If

If GelenPaketTipi.Value <> "" Then
    GelenPaketTipi.SelStart = 0
    GelenPaketTipi.SelLength = Len(GelenPaketTipi.Value)
End If

Son:

GelenPaketTipi.DropDown
If GelenPaketTipi.BackColor = RGB(60, 100, 180) Then
GelenPaketTipi.BackColor = RGB(255, 255, 255)
GelenPaketTipi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub GelisSekli_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GelisSekli.DropDown
GelisSekli.BackColor = RGB(255, 255, 255)
GelisSekli.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelisSekli_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GelisSekli.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelisSekli.ListIndex = GelisSekli.ListIndex - 1
            End If
            Me.GelisSekli.DropDown
            
        Case 40 'Aşağı
            If GelisSekli.ListIndex = GelisSekli.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelisSekli.ListIndex = GelisSekli.ListIndex + 1
            End If
            Me.GelisSekli.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GelisSekli_Change()

If GelisSekli.ListIndex = -1 And GelisSekli.Value <> "" Then
   GelisSekli.Value = ""
   GoTo Son
End If

If GelisSekli.Value <> "" Then
    GelisSekli.SelStart = 0
    GelisSekli.SelLength = Len(GelisSekli.Value)
End If

Son:

GelisSekli.DropDown
If GelisSekli.BackColor = RGB(60, 100, 180) Then
GelisSekli.BackColor = RGB(255, 255, 255)
GelisSekli.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Tutanak1TarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        Tutanak1TarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        Tutanak1TarihiText.Value = ""
    End If

End Sub

Private Sub Tutanak1TarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Tutanak1TarihiText.BackColor = RGB(255, 255, 255)
Tutanak1TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak1TarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    Tutanak1TarihiText.Value = CalTarih
    Tutanak1TarihiText.Value = Format(Tutanak1TarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

Tutanak1TarihiText.BackColor = RGB(255, 255, 255)
Tutanak1TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak1Sonucu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Tutanak1Sonucu.DropDown
Tutanak1Sonucu.BackColor = RGB(255, 255, 255)
Tutanak1Sonucu.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak1Sonucu_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tutanak1Sonucu.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Sonucu.ListIndex = Tutanak1Sonucu.ListIndex - 1
            End If
            Me.Tutanak1Sonucu.DropDown
            
        Case 40 'Aşağı
            If Tutanak1Sonucu.ListIndex = Tutanak1Sonucu.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Sonucu.ListIndex = Tutanak1Sonucu.ListIndex + 1
            End If
            Me.Tutanak1Sonucu.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tutanak1Sonucu_Change()

If Tutanak1Sonucu.ListIndex = -1 And Tutanak1Sonucu.Value <> "" Then
   Tutanak1Sonucu.Value = ""
   GoTo Son
End If

If Tutanak1Sonucu.Value <> "" Then
    Tutanak1Sonucu.SelStart = 0
    Tutanak1Sonucu.SelLength = Len(Tutanak1Sonucu.Value)
End If


If Tutanak1Sonucu.Value = "d. Discrepancy Detected" Then
    FarkGirisi.Visible = True
Else
    FarkGirisi.Visible = False
End If

Son:

Tutanak1Sonucu.DropDown
If Tutanak1Sonucu.BackColor = RGB(60, 100, 180) Then
Tutanak1Sonucu.BackColor = RGB(255, 255, 255)
Tutanak1Sonucu.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub DosyaNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

DosyaNoText.BackColor = RGB(255, 255, 255)
DosyaNoText.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub DosyaNoText_Change()

DosyaNoText.BackColor = RGB(255, 255, 255)
DosyaNoText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelenBelgeSayfa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GelenBelgeSayfa.DropDown
GelenBelgeSayfa.BackColor = RGB(255, 255, 255)
GelenBelgeSayfa.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GelenBelgeSayfa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GelenBelgeSayfa.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenBelgeSayfa.ListIndex = GelenBelgeSayfa.ListIndex - 1
            End If
            Me.GelenBelgeSayfa.DropDown
            
        Case 40 'Aşağı
            If GelenBelgeSayfa.ListIndex = GelenBelgeSayfa.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GelenBelgeSayfa.ListIndex = GelenBelgeSayfa.ListIndex + 1
            End If
            Me.GelenBelgeSayfa.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GelenBelgeSayfa_Change()

If GelenBelgeSayfa.ListIndex = -1 And GelenBelgeSayfa.Value <> "" Then
   GelenBelgeSayfa.Value = ""
   GoTo Son
End If

If GelenBelgeSayfa.Value <> "" Then
    GelenBelgeSayfa.SelStart = 0
    GelenBelgeSayfa.SelLength = Len(GelenBelgeSayfa.Value)
End If

Son:

GelenBelgeSayfa.DropDown
If GelenBelgeSayfa.BackColor = RGB(60, 100, 180) Then
GelenBelgeSayfa.BackColor = RGB(255, 255, 255)
GelenBelgeSayfa.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub DokumListesi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.DokumListesi.DropDown
DokumListesi.BackColor = RGB(255, 255, 255)
DokumListesi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub DokumListesi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If DokumListesi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DokumListesi.ListIndex = DokumListesi.ListIndex - 1
            End If
            Me.DokumListesi.DropDown
            
        Case 40 'Aşağı
            If DokumListesi.ListIndex = DokumListesi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DokumListesi.ListIndex = DokumListesi.ListIndex + 1
            End If
            Me.DokumListesi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub DokumListesi_Change()

If DokumListesi.ListIndex = -1 And DokumListesi.Value <> "" Then
   DokumListesi.Value = ""
   GoTo Son
End If

If DokumListesi.Value <> "" Then
    DokumListesi.SelStart = 0
    DokumListesi.SelLength = Len(DokumListesi.Value)
End If

Son:

DokumListesi.DropDown
If DokumListesi.BackColor = RGB(60, 100, 180) Then
DokumListesi.BackColor = RGB(255, 255, 255)
DokumListesi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Tutanak1Imza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Tutanak1Imza1.DropDown
Tutanak1Imza1.BackColor = RGB(255, 255, 255)
Tutanak1Imza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak1Imza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tutanak1Imza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Imza1.ListIndex = Tutanak1Imza1.ListIndex - 1
            End If
            Me.Tutanak1Imza1.DropDown
            
        Case 40 'Aşağı
            If Tutanak1Imza1.ListIndex = Tutanak1Imza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Imza1.ListIndex = Tutanak1Imza1.ListIndex + 1
            End If
            Me.Tutanak1Imza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tutanak1Imza1_Change()

If Tutanak1Imza1.ListIndex = -1 And Tutanak1Imza1.Value <> "" Then
   Tutanak1Imza1.Value = ""
   GoTo Son
End If

If Tutanak1Imza1.Value <> "" Then
    Tutanak1Imza1.SelStart = 0
    Tutanak1Imza1.SelLength = Len(Tutanak1Imza1.Value)
End If


Son:

Tutanak1Imza1.DropDown
If Tutanak1Imza1.BackColor = RGB(60, 100, 180) Then
Tutanak1Imza1.BackColor = RGB(255, 255, 255)
Tutanak1Imza1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Tutanak1Imza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Tutanak1Imza2.DropDown
Tutanak1Imza2.BackColor = RGB(255, 255, 255)
Tutanak1Imza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak1Imza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tutanak1Imza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Imza2.ListIndex = Tutanak1Imza2.ListIndex - 1
            End If
            Me.Tutanak1Imza2.DropDown
            
        Case 40 'Aşağı
            If Tutanak1Imza2.ListIndex = Tutanak1Imza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak1Imza2.ListIndex = Tutanak1Imza2.ListIndex + 1
            End If
            Me.Tutanak1Imza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tutanak1Imza2_Change()

If Tutanak1Imza2.ListIndex = -1 And Tutanak1Imza2.Value <> "" Then
   Tutanak1Imza2.Value = ""
   GoTo Son
End If

If Tutanak1Imza2.Value <> "" Then
    Tutanak1Imza2.SelStart = 0
    Tutanak1Imza2.SelLength = Len(Tutanak1Imza2.Value)
End If


Son:

Tutanak1Imza2.DropDown
If Tutanak1Imza2.BackColor = RGB(60, 100, 180) Then
Tutanak1Imza2.BackColor = RGB(255, 255, 255)
Tutanak1Imza2.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub RaporImza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporImza1.DropDown
RaporImza1.BackColor = RGB(255, 255, 255)
RaporImza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporImza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If RaporImza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporImza1.ListIndex = RaporImza1.ListIndex - 1
            End If
            Me.RaporImza1.DropDown
            
        Case 40 'Aşağı
            If RaporImza1.ListIndex = RaporImza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporImza1.ListIndex = RaporImza1.ListIndex + 1
            End If
            Me.RaporImza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub RaporImza1_Change()

If RaporImza1.ListIndex = -1 And RaporImza1.Value <> "" Then
   RaporImza1.Value = ""
   GoTo Son
End If

If RaporImza1.Value <> "" Then
    RaporImza1.SelStart = 0
    RaporImza1.SelLength = Len(RaporImza1.Value)
End If


Son:

RaporImza1.DropDown
If RaporImza1.BackColor = RGB(60, 100, 180) Then
RaporImza1.BackColor = RGB(255, 255, 255)
RaporImza1.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub RaporImza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporImza2.DropDown
RaporImza2.BackColor = RGB(255, 255, 255)
RaporImza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporImza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If RaporImza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporImza2.ListIndex = RaporImza2.ListIndex - 1
            End If
            Me.RaporImza2.DropDown
            
        Case 40 'Aşağı
            If RaporImza2.ListIndex = RaporImza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporImza2.ListIndex = RaporImza2.ListIndex + 1
            End If
            Me.RaporImza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub RaporImza2_Change()

If RaporImza2.ListIndex = -1 And RaporImza2.Value <> "" Then
   RaporImza2.Value = ""
   GoTo Son
End If

If RaporImza2.Value <> "" Then
    RaporImza2.SelStart = 0
    RaporImza2.SelLength = Len(RaporImza2.Value)
End If


Son:

RaporImza2.DropDown
If RaporImza2.BackColor = RGB(60, 100, 180) Then
RaporImza2.BackColor = RGB(255, 255, 255)
RaporImza2.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub Tutanak2Imza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Tutanak2Imza1.DropDown
Tutanak2Imza1.BackColor = RGB(255, 255, 255)
Tutanak2Imza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak2Imza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tutanak2Imza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak2Imza1.ListIndex = Tutanak2Imza1.ListIndex - 1
            End If
            Me.Tutanak2Imza1.DropDown
            
        Case 40 'Aşağı
            If Tutanak2Imza1.ListIndex = Tutanak2Imza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak2Imza1.ListIndex = Tutanak2Imza1.ListIndex + 1
            End If
            Me.Tutanak2Imza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tutanak2Imza1_Change()

If Tutanak2Imza1.ListIndex = -1 And Tutanak2Imza1.Value <> "" Then
   Tutanak2Imza1.Value = ""
   GoTo Son
End If

If Tutanak2Imza1.Value <> "" Then
    Tutanak2Imza1.SelStart = 0
    Tutanak2Imza1.SelLength = Len(Tutanak2Imza1.Value)
End If


Son:

Tutanak2Imza1.DropDown
If Tutanak2Imza1.BackColor = RGB(60, 100, 180) Then
Tutanak2Imza1.BackColor = RGB(255, 255, 255)
Tutanak2Imza1.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub Tutanak2Imza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Tutanak2Imza2.DropDown
Tutanak2Imza2.BackColor = RGB(255, 255, 255)
Tutanak2Imza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak2Imza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tutanak2Imza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak2Imza2.ListIndex = Tutanak2Imza2.ListIndex - 1
            End If
            Me.Tutanak2Imza2.DropDown
            
        Case 40 'Aşağı
            If Tutanak2Imza2.ListIndex = Tutanak2Imza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tutanak2Imza2.ListIndex = Tutanak2Imza2.ListIndex + 1
            End If
            Me.Tutanak2Imza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tutanak2Imza2_Change()

If Tutanak2Imza2.ListIndex = -1 And Tutanak2Imza2.Value <> "" Then
   Tutanak2Imza2.Value = ""
   GoTo Son
End If

If Tutanak2Imza2.Value <> "" Then
    Tutanak2Imza2.SelStart = 0
    Tutanak2Imza2.SelLength = Len(Tutanak2Imza2.Value)
End If


Son:

Tutanak2Imza2.DropDown
If Tutanak2Imza2.BackColor = RGB(60, 100, 180) Then
Tutanak2Imza2.BackColor = RGB(255, 255, 255)
Tutanak2Imza2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UstYaziImza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UstYaziImza1.DropDown
UstYaziImza1.BackColor = RGB(255, 255, 255)
UstYaziImza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UstYaziImza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If UstYaziImza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UstYaziImza1.ListIndex = UstYaziImza1.ListIndex - 1
            End If
            Me.UstYaziImza1.DropDown
            
        Case 40 'Aşağı
            If UstYaziImza1.ListIndex = UstYaziImza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UstYaziImza1.ListIndex = UstYaziImza1.ListIndex + 1
            End If
            Me.UstYaziImza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub UstYaziImza1_Change()

If UstYaziImza1.ListIndex = -1 And UstYaziImza1.Value <> "" Then
   UstYaziImza1.Value = ""
   GoTo Son
End If

If UstYaziImza1.Value <> "" Then
    UstYaziImza1.SelStart = 0
    UstYaziImza1.SelLength = Len(UstYaziImza1.Value)
End If


Son:

UstYaziImza1.DropDown
If UstYaziImza1.BackColor = RGB(60, 100, 180) Then
UstYaziImza1.BackColor = RGB(255, 255, 255)
UstYaziImza1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UstYaziImza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UstYaziImza2.DropDown
UstYaziImza2.BackColor = RGB(255, 255, 255)
UstYaziImza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UstYaziImza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If UstYaziImza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UstYaziImza2.ListIndex = UstYaziImza2.ListIndex - 1
            End If
            Me.UstYaziImza2.DropDown
            
        Case 40 'Aşağı
            If UstYaziImza2.ListIndex = UstYaziImza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UstYaziImza2.ListIndex = UstYaziImza2.ListIndex + 1
            End If
            Me.UstYaziImza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub UstYaziImza2_Change()

If UstYaziImza2.ListIndex = -1 And UstYaziImza2.Value <> "" Then
   UstYaziImza2.Value = ""
   GoTo Son
End If

If UstYaziImza2.Value <> "" Then
    UstYaziImza2.SelStart = 0
    UstYaziImza2.SelLength = Len(UstYaziImza2.Value)
End If


Son:

UstYaziImza2.DropDown
If UstYaziImza2.BackColor = RGB(60, 100, 180) Then
UstYaziImza2.BackColor = RGB(255, 255, 255)
UstYaziImza2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru.DropDown
OgeTuru.BackColor = RGB(255, 255, 255)
OgeTuru.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru1.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru.ListIndex = OgeTuru.ListIndex
            End If
        Case 40 'Down
            If OgeTuru.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru.ListIndex = OgeTuru.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru_Change()

If OgeTuru.ListIndex = -1 And OgeTuru.Value <> "" Then
   OgeTuru.Value = ""
   GoTo Son
End If

If OgeTuru.Value <> "" Then
    OgeTuru.SelStart = 0
    OgeTuru.SelLength = Len(OgeTuru.Value)
End If

If OgeTuru.Value <> "" And OgeIdNo.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo.Value = "Dispatch List"
ElseIf OgeTuru.Value = "" And OgeIdNo.Value <> "" Then
    OgeIdNo.Value = ""
End If

Son:

OgeTuru.DropDown
If OgeTuru.BackColor = RGB(60, 100, 180) Then
OgeTuru.BackColor = RGB(255, 255, 255)
OgeTuru.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru1.DropDown
OgeTuru1.BackColor = RGB(255, 255, 255)
OgeTuru1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru2.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru1.ListIndex = OgeTuru1.ListIndex
            End If
        Case 40 'Down
            If OgeTuru1.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru1.ListIndex = OgeTuru1.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru1_Change()

If OgeTuru1.ListIndex = -1 And OgeTuru1.Value <> "" Then
   OgeTuru1.Value = ""
   GoTo Son
End If

If OgeTuru1.Value <> "" Then
    OgeTuru1.SelStart = 0
    OgeTuru1.SelLength = Len(OgeTuru1.Value)
End If

If OgeTuru1.Value <> "" And OgeIdNo1.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo1.Value = "Dispatch List"
ElseIf OgeTuru1.Value = "" And OgeIdNo1.Value <> "" Then
    OgeIdNo1.Value = ""
End If

Son:

OgeTuru1.DropDown
If OgeTuru1.BackColor = RGB(60, 100, 180) Then
OgeTuru1.BackColor = RGB(255, 255, 255)
OgeTuru1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru2.DropDown
OgeTuru2.BackColor = RGB(255, 255, 255)
OgeTuru2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru3.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru2.ListIndex = OgeTuru2.ListIndex
            End If
        Case 40 'Down
            If OgeTuru2.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru2.ListIndex = OgeTuru2.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru2_Change()

If OgeTuru2.ListIndex = -1 And OgeTuru2.Value <> "" Then
   OgeTuru2.Value = ""
   GoTo Son
End If

If OgeTuru2.Value <> "" Then
    OgeTuru2.SelStart = 0
    OgeTuru2.SelLength = Len(OgeTuru2.Value)
End If

If OgeTuru2.Value <> "" And OgeIdNo2.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo2.Value = "Dispatch List"
ElseIf OgeTuru2.Value = "" And OgeIdNo2.Value <> "" Then
    OgeIdNo2.Value = ""
End If

Son:

OgeTuru2.DropDown
If OgeTuru2.BackColor = RGB(60, 100, 180) Then
OgeTuru2.BackColor = RGB(255, 255, 255)
OgeTuru2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru3.DropDown
OgeTuru3.BackColor = RGB(255, 255, 255)
OgeTuru3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru4.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru3.ListIndex = OgeTuru3.ListIndex
            End If
        Case 40 'Down
            If OgeTuru3.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru3.ListIndex = OgeTuru3.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru3_Change()

If OgeTuru3.ListIndex = -1 And OgeTuru3.Value <> "" Then
   OgeTuru3.Value = ""
   GoTo Son
End If

If OgeTuru3.Value <> "" Then
    OgeTuru3.SelStart = 0
    OgeTuru3.SelLength = Len(OgeTuru3.Value)
End If

If OgeTuru3.Value <> "" And OgeIdNo3.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo3.Value = "Dispatch List"
ElseIf OgeTuru3.Value = "" And OgeIdNo3.Value <> "" Then
    OgeIdNo3.Value = ""
End If

Son:

OgeTuru3.DropDown
If OgeTuru3.BackColor = RGB(60, 100, 180) Then
OgeTuru3.BackColor = RGB(255, 255, 255)
OgeTuru3.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru4.DropDown
OgeTuru4.BackColor = RGB(255, 255, 255)
OgeTuru4.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru5.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru4.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru4.ListIndex = OgeTuru4.ListIndex
            End If
        Case 40 'Down
            If OgeTuru4.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru4.ListIndex = OgeTuru4.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru4_Change()

If OgeTuru4.ListIndex = -1 And OgeTuru4.Value <> "" Then
   OgeTuru4.Value = ""
   GoTo Son
End If

If OgeTuru4.Value <> "" Then
    OgeTuru4.SelStart = 0
    OgeTuru4.SelLength = Len(OgeTuru4.Value)
End If

If OgeTuru4.Value <> "" And OgeIdNo4.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo4.Value = "Dispatch List"
ElseIf OgeTuru4.Value = "" And OgeIdNo4.Value <> "" Then
    OgeIdNo4.Value = ""
End If

Son:

OgeTuru4.DropDown
If OgeTuru4.BackColor = RGB(60, 100, 180) Then
OgeTuru4.BackColor = RGB(255, 255, 255)
OgeTuru4.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru5.DropDown
OgeTuru5.BackColor = RGB(255, 255, 255)
OgeTuru5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru6.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru5.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru5.ListIndex = OgeTuru5.ListIndex
            End If
        Case 40 'Down
            If OgeTuru5.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru5.ListIndex = OgeTuru5.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru5_Change()

If OgeTuru5.ListIndex = -1 And OgeTuru5.Value <> "" Then
   OgeTuru5.Value = ""
   GoTo Son
End If

If OgeTuru5.Value <> "" Then
    OgeTuru5.SelStart = 0
    OgeTuru5.SelLength = Len(OgeTuru5.Value)
End If

If OgeTuru5.Value <> "" And OgeIdNo5.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo5.Value = "Dispatch List"
ElseIf OgeTuru5.Value = "" And OgeIdNo5.Value <> "" Then
    OgeIdNo5.Value = ""
End If

Son:

OgeTuru5.DropDown
If OgeTuru5.BackColor = RGB(60, 100, 180) Then
OgeTuru5.BackColor = RGB(255, 255, 255)
OgeTuru5.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru6.DropDown
OgeTuru6.BackColor = RGB(255, 255, 255)
OgeTuru6.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru7.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru6.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru6.ListIndex = OgeTuru6.ListIndex
            End If
        Case 40 'Down
            If OgeTuru6.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru6.ListIndex = OgeTuru6.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru6_Change()

If OgeTuru6.ListIndex = -1 And OgeTuru6.Value <> "" Then
   OgeTuru6.Value = ""
   GoTo Son
End If

If OgeTuru6.Value <> "" Then
    OgeTuru6.SelStart = 0
    OgeTuru6.SelLength = Len(OgeTuru6.Value)
End If

If OgeTuru6.Value <> "" And OgeIdNo6.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo6.Value = "Dispatch List"
ElseIf OgeTuru6.Value = "" And OgeIdNo6.Value <> "" Then
    OgeIdNo6.Value = ""
End If

Son:

OgeTuru6.DropDown
If OgeTuru6.BackColor = RGB(60, 100, 180) Then
OgeTuru6.BackColor = RGB(255, 255, 255)
OgeTuru6.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru7.DropDown
OgeTuru7.BackColor = RGB(255, 255, 255)
OgeTuru7.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru8.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru7.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru7.ListIndex = OgeTuru7.ListIndex
            End If
        Case 40 'Down
            If OgeTuru7.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru7.ListIndex = OgeTuru7.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru7_Change()

If OgeTuru7.ListIndex = -1 And OgeTuru7.Value <> "" Then
   OgeTuru7.Value = ""
   GoTo Son
End If

If OgeTuru7.Value <> "" Then
    OgeTuru7.SelStart = 0
    OgeTuru7.SelLength = Len(OgeTuru7.Value)
End If

If OgeTuru7.Value <> "" And OgeIdNo7.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo7.Value = "Dispatch List"
ElseIf OgeTuru7.Value = "" And OgeIdNo7.Value <> "" Then
    OgeIdNo7.Value = ""
End If

Son:

OgeTuru7.DropDown
If OgeTuru7.BackColor = RGB(60, 100, 180) Then
OgeTuru7.BackColor = RGB(255, 255, 255)
OgeTuru7.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru8.DropDown
OgeTuru8.BackColor = RGB(255, 255, 255)
OgeTuru8.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru9.SetFocus
    End If
    
    Select Case KeyCode
        Case 38  'Up
            If OgeTuru8.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru8.ListIndex = OgeTuru8.ListIndex
            End If
        Case 40 'Down
            If OgeTuru8.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru8.ListIndex = OgeTuru8.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru8_Change()

If OgeTuru8.ListIndex = -1 And OgeTuru8.Value <> "" Then
   OgeTuru8.Value = ""
   GoTo Son
End If

If OgeTuru8.Value <> "" Then
    OgeTuru8.SelStart = 0
    OgeTuru8.SelLength = Len(OgeTuru8.Value)
End If

If OgeTuru8.Value <> "" And OgeIdNo8.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo8.Value = "Dispatch List"
ElseIf OgeTuru8.Value = "" And OgeIdNo8.Value <> "" Then
    OgeIdNo8.Value = ""
End If

Son:

OgeTuru8.DropDown
If OgeTuru8.BackColor = RGB(60, 100, 180) Then
OgeTuru8.BackColor = RGB(255, 255, 255)
OgeTuru8.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru9.DropDown
OgeTuru9.BackColor = RGB(255, 255, 255)
OgeTuru9.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru10.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru9.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru9.ListIndex = OgeTuru9.ListIndex
            End If
        Case 40 'Down
            If OgeTuru9.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru9.ListIndex = OgeTuru9.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru9_Change()

If OgeTuru9.ListIndex = -1 And OgeTuru9.Value <> "" Then
   OgeTuru9.Value = ""
   GoTo Son
End If

If OgeTuru9.Value <> "" Then
    OgeTuru9.SelStart = 0
    OgeTuru9.SelLength = Len(OgeTuru9.Value)
End If

If OgeTuru9.Value <> "" And OgeIdNo9.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo9.Value = "Dispatch List"
ElseIf OgeTuru9.Value = "" And OgeIdNo9.Value <> "" Then
    OgeIdNo9.Value = ""
End If

Son:

OgeTuru9.DropDown
If OgeTuru9.BackColor = RGB(60, 100, 180) Then
OgeTuru9.BackColor = RGB(255, 255, 255)
OgeTuru9.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru10.DropDown
OgeTuru10.BackColor = RGB(255, 255, 255)
OgeTuru10.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru11.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru10.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru10.ListIndex = OgeTuru10.ListIndex
            End If
        Case 40 'Down
            If OgeTuru10.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru10.ListIndex = OgeTuru10.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru10_Change()

If OgeTuru10.ListIndex = -1 And OgeTuru10.Value <> "" Then
   OgeTuru10.Value = ""
   GoTo Son
End If

If OgeTuru10.Value <> "" Then
    OgeTuru10.SelStart = 0
    OgeTuru10.SelLength = Len(OgeTuru10.Value)
End If

If OgeTuru10.Value <> "" And OgeIdNo10.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo10.Value = "Dispatch List"
ElseIf OgeTuru10.Value = "" And OgeIdNo10.Value <> "" Then
    OgeIdNo10.Value = ""
End If

Son:

OgeTuru10.DropDown
If OgeTuru10.BackColor = RGB(60, 100, 180) Then
OgeTuru10.BackColor = RGB(255, 255, 255)
OgeTuru10.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru11.DropDown
OgeTuru11.BackColor = RGB(255, 255, 255)
OgeTuru11.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru12.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru11.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru11.ListIndex = OgeTuru11.ListIndex
            End If
        Case 40 'Down
            If OgeTuru11.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru11.ListIndex = OgeTuru11.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru11_Change()

If OgeTuru11.ListIndex = -1 And OgeTuru11.Value <> "" Then
   OgeTuru11.Value = ""
   GoTo Son
End If

If OgeTuru11.Value <> "" Then
    OgeTuru11.SelStart = 0
    OgeTuru11.SelLength = Len(OgeTuru11.Value)
End If

If OgeTuru11.Value <> "" And OgeIdNo11.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo11.Value = "Dispatch List"
ElseIf OgeTuru11.Value = "" And OgeIdNo11.Value <> "" Then
    OgeIdNo11.Value = ""
End If

Son:

OgeTuru11.DropDown
If OgeTuru11.BackColor = RGB(60, 100, 180) Then
OgeTuru11.BackColor = RGB(255, 255, 255)
OgeTuru11.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru12.DropDown
OgeTuru12.BackColor = RGB(255, 255, 255)
OgeTuru12.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru13.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru12.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru12.ListIndex = OgeTuru12.ListIndex
            End If
        Case 40 'Down
            If OgeTuru12.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru12.ListIndex = OgeTuru12.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru12_Change()

If OgeTuru12.ListIndex = -1 And OgeTuru12.Value <> "" Then
   OgeTuru12.Value = ""
   GoTo Son
End If

If OgeTuru12.Value <> "" Then
    OgeTuru12.SelStart = 0
    OgeTuru12.SelLength = Len(OgeTuru12.Value)
End If

If OgeTuru12.Value <> "" And OgeIdNo12.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo12.Value = "Dispatch List"
ElseIf OgeTuru12.Value = "" And OgeIdNo12.Value <> "" Then
    OgeIdNo12.Value = ""
End If

Son:

OgeTuru12.DropDown
If OgeTuru12.BackColor = RGB(60, 100, 180) Then
OgeTuru12.BackColor = RGB(255, 255, 255)
OgeTuru12.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru13.DropDown
OgeTuru13.BackColor = RGB(255, 255, 255)
OgeTuru13.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru14.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru13.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru13.ListIndex = OgeTuru13.ListIndex
            End If
        Case 40 'Down
            If OgeTuru13.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru13.ListIndex = OgeTuru13.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru13_Change()

If OgeTuru13.ListIndex = -1 And OgeTuru13.Value <> "" Then
   OgeTuru13.Value = ""
   GoTo Son
End If

If OgeTuru13.Value <> "" Then
    OgeTuru13.SelStart = 0
    OgeTuru13.SelLength = Len(OgeTuru13.Value)
End If

If OgeTuru13.Value <> "" And OgeIdNo13.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo13.Value = "Dispatch List"
ElseIf OgeTuru13.Value = "" And OgeIdNo13.Value <> "" Then
    OgeIdNo13.Value = ""
End If

Son:

OgeTuru13.DropDown
If OgeTuru13.BackColor = RGB(60, 100, 180) Then
OgeTuru13.BackColor = RGB(255, 255, 255)
OgeTuru13.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru14.DropDown
OgeTuru14.BackColor = RGB(255, 255, 255)
OgeTuru14.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru15.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru14.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru14.ListIndex = OgeTuru14.ListIndex
            End If
        Case 40 'Down
            If OgeTuru14.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru14.ListIndex = OgeTuru14.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru14_Change()

If OgeTuru14.ListIndex = -1 And OgeTuru14.Value <> "" Then
   OgeTuru14.Value = ""
   GoTo Son
End If

If OgeTuru14.Value <> "" Then
    OgeTuru14.SelStart = 0
    OgeTuru14.SelLength = Len(OgeTuru14.Value)
End If

If OgeTuru14.Value <> "" And OgeIdNo14.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo14.Value = "Dispatch List"
ElseIf OgeTuru14.Value = "" And OgeIdNo14.Value <> "" Then
    OgeIdNo14.Value = ""
End If

Son:

OgeTuru14.DropDown
If OgeTuru14.BackColor = RGB(60, 100, 180) Then
OgeTuru14.BackColor = RGB(255, 255, 255)
OgeTuru14.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru15.DropDown
OgeTuru15.BackColor = RGB(255, 255, 255)
OgeTuru15.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru16.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru15.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru15.ListIndex = OgeTuru15.ListIndex
            End If
        Case 40 'Down
            If OgeTuru15.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru15.ListIndex = OgeTuru15.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru15_Change()

If OgeTuru15.ListIndex = -1 And OgeTuru15.Value <> "" Then
   OgeTuru15.Value = ""
   GoTo Son
End If

If OgeTuru15.Value <> "" Then
    OgeTuru15.SelStart = 0
    OgeTuru15.SelLength = Len(OgeTuru15.Value)
End If

If OgeTuru15.Value <> "" And OgeIdNo15.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo15.Value = "Dispatch List"
ElseIf OgeTuru15.Value = "" And OgeIdNo15.Value <> "" Then
    OgeIdNo15.Value = ""
End If

Son:

OgeTuru15.DropDown
If OgeTuru15.BackColor = RGB(60, 100, 180) Then
OgeTuru15.BackColor = RGB(255, 255, 255)
OgeTuru15.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru16.DropDown
OgeTuru16.BackColor = RGB(255, 255, 255)
OgeTuru16.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru17.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru16.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru16.ListIndex = OgeTuru16.ListIndex
            End If
        Case 40 'Down
            If OgeTuru16.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru16.ListIndex = OgeTuru16.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru16_Change()

If OgeTuru16.ListIndex = -1 And OgeTuru16.Value <> "" Then
   OgeTuru16.Value = ""
   GoTo Son
End If

If OgeTuru16.Value <> "" Then
    OgeTuru16.SelStart = 0
    OgeTuru16.SelLength = Len(OgeTuru16.Value)
End If

If OgeTuru16.Value <> "" And OgeIdNo16.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo16.Value = "Dispatch List"
ElseIf OgeTuru16.Value = "" And OgeIdNo16.Value <> "" Then
    OgeIdNo16.Value = ""
End If

Son:

OgeTuru16.DropDown
If OgeTuru16.BackColor = RGB(60, 100, 180) Then
OgeTuru16.BackColor = RGB(255, 255, 255)
OgeTuru16.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru17.DropDown
OgeTuru17.BackColor = RGB(255, 255, 255)
OgeTuru17.ForeColor = RGB(30, 30, 30)


End Sub

Private Sub OgeTuru17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru18.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru17.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru17.ListIndex = OgeTuru17.ListIndex
            End If
        Case 40 'Down
            If OgeTuru17.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru17.ListIndex = OgeTuru17.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru17_Change()

If OgeTuru17.ListIndex = -1 And OgeTuru17.Value <> "" Then
   OgeTuru17.Value = ""
   GoTo Son
End If

If OgeTuru17.Value <> "" Then
    OgeTuru17.SelStart = 0
    OgeTuru17.SelLength = Len(OgeTuru17.Value)
End If

If OgeTuru17.Value <> "" And OgeIdNo17.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo17.Value = "Dispatch List"
ElseIf OgeTuru17.Value = "" And OgeIdNo17.Value <> "" Then
    OgeIdNo17.Value = ""
End If

Son:

OgeTuru17.DropDown
If OgeTuru17.BackColor = RGB(60, 100, 180) Then
OgeTuru17.BackColor = RGB(255, 255, 255)
OgeTuru17.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub OgeTuru18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru18.DropDown
OgeTuru18.BackColor = RGB(255, 255, 255)
OgeTuru18.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeTuru19.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru18.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru18.ListIndex = OgeTuru18.ListIndex
            End If
        Case 40 'Down
            If OgeTuru18.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru18.ListIndex = OgeTuru18.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru18_Change()

If OgeTuru18.ListIndex = -1 And OgeTuru18.Value <> "" Then
   OgeTuru18.Value = ""
   GoTo Son
End If

If OgeTuru18.Value <> "" Then
    OgeTuru18.SelStart = 0
    OgeTuru18.SelLength = Len(OgeTuru18.Value)
End If

If OgeTuru18.Value <> "" And OgeIdNo18.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo18.Value = "Dispatch List"
ElseIf OgeTuru18.Value = "" And OgeIdNo18.Value <> "" Then
    OgeIdNo18.Value = ""
End If

Son:

OgeTuru18.DropDown
If OgeTuru18.BackColor = RGB(60, 100, 180) Then
OgeTuru18.BackColor = RGB(255, 255, 255)
OgeTuru18.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeTuru19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeTuru19.DropDown
OgeTuru19.BackColor = RGB(255, 255, 255)
OgeTuru19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeTuru19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeTuru18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'OgeTuru20.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeTuru19.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru19.ListIndex = OgeTuru19.ListIndex
            End If
        Case 40 'Down
            If OgeTuru19.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeTuru19.ListIndex = OgeTuru19.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeTuru19_Change()

If OgeTuru19.ListIndex = -1 And OgeTuru19.Value <> "" Then
   OgeTuru19.Value = ""
   GoTo Son
End If

If OgeTuru19.Value <> "" Then
    OgeTuru19.SelStart = 0
    OgeTuru19.SelLength = Len(OgeTuru19.Value)
End If

If OgeTuru19.Value <> "" And OgeIdNo19.Value = "" And DLEvetOption.Value = True Then
    OgeIdNo19.Value = "Dispatch List"
ElseIf OgeTuru19.Value = "" And OgeIdNo19.Value <> "" Then
    OgeIdNo19.Value = ""
End If

Son:

OgeTuru19.DropDown
If OgeTuru19.BackColor = RGB(60, 100, 180) Then
OgeTuru19.BackColor = RGB(255, 255, 255)
OgeTuru19.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri.DropDown
OgeDegeri.BackColor = RGB(255, 255, 255)
OgeDegeri.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        'OgeDegeri0.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri1.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri.ListIndex = OgeDegeri.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri.ListIndex = OgeDegeri.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri_Change()

If OgeDegeri.ListIndex = -1 And OgeDegeri.Value <> "" Then
   OgeDegeri.Value = ""
   GoTo Son
End If

If OgeDegeri.Value <> "" Then
    OgeDegeri.SelStart = 0
    OgeDegeri.SelLength = Len(OgeDegeri.Value)
End If

Son:
OgeDegeri.DropDown
If OgeDegeri.BackColor = RGB(60, 100, 180) Then
OgeDegeri.BackColor = RGB(255, 255, 255)
OgeDegeri.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri1.DropDown
OgeDegeri1.BackColor = RGB(255, 255, 255)
OgeDegeri1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri2.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri1.ListIndex = OgeDegeri1.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri1.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri1.ListIndex = OgeDegeri1.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri1_Change()

If OgeDegeri1.ListIndex = -1 And OgeDegeri1.Value <> "" Then
   OgeDegeri1.Value = ""
   GoTo Son
End If

If OgeDegeri1.Value <> "" Then
    OgeDegeri1.SelStart = 0
    OgeDegeri1.SelLength = Len(OgeDegeri1.Value)
End If

Son:
OgeDegeri1.DropDown
If OgeDegeri1.BackColor = RGB(60, 100, 180) Then
OgeDegeri1.BackColor = RGB(255, 255, 255)
OgeDegeri1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri2.DropDown
OgeDegeri2.BackColor = RGB(255, 255, 255)
OgeDegeri2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri3.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri2.ListIndex = OgeDegeri2.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri2.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri2.ListIndex = OgeDegeri2.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri2_Change()

If OgeDegeri2.ListIndex = -1 And OgeDegeri2.Value <> "" Then
   OgeDegeri2.Value = ""
   GoTo Son
End If

If OgeDegeri2.Value <> "" Then
    OgeDegeri2.SelStart = 0
    OgeDegeri2.SelLength = Len(OgeDegeri2.Value)
End If

Son:
OgeDegeri2.DropDown
If OgeDegeri2.BackColor = RGB(60, 100, 180) Then
OgeDegeri2.BackColor = RGB(255, 255, 255)
OgeDegeri2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri3.DropDown
OgeDegeri3.BackColor = RGB(255, 255, 255)
OgeDegeri3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri4.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri3.ListIndex = OgeDegeri3.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri3.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri3.ListIndex = OgeDegeri3.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri3_Change()

If OgeDegeri3.ListIndex = -1 And OgeDegeri3.Value <> "" Then
   OgeDegeri3.Value = ""
   GoTo Son
End If

If OgeDegeri3.Value <> "" Then
    OgeDegeri3.SelStart = 0
    OgeDegeri3.SelLength = Len(OgeDegeri3.Value)
End If

Son:
OgeDegeri3.DropDown
If OgeDegeri3.BackColor = RGB(60, 100, 180) Then
OgeDegeri3.BackColor = RGB(255, 255, 255)
OgeDegeri3.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri4.DropDown
OgeDegeri4.BackColor = RGB(255, 255, 255)
OgeDegeri4.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri5.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri4.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri4.ListIndex = OgeDegeri4.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri4.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri4.ListIndex = OgeDegeri4.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri4_Change()

If OgeDegeri4.ListIndex = -1 And OgeDegeri4.Value <> "" Then
   OgeDegeri4.Value = ""
   GoTo Son
End If

If OgeDegeri4.Value <> "" Then
    OgeDegeri4.SelStart = 0
    OgeDegeri4.SelLength = Len(OgeDegeri4.Value)
End If

Son:
OgeDegeri4.DropDown
If OgeDegeri4.BackColor = RGB(60, 100, 180) Then
OgeDegeri4.BackColor = RGB(255, 255, 255)
OgeDegeri4.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri5.DropDown
OgeDegeri5.BackColor = RGB(255, 255, 255)
OgeDegeri5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri6.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri5.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri5.ListIndex = OgeDegeri5.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri5.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri5.ListIndex = OgeDegeri5.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri5_Change()

If OgeDegeri5.ListIndex = -1 And OgeDegeri5.Value <> "" Then
   OgeDegeri5.Value = ""
   GoTo Son
End If

If OgeDegeri5.Value <> "" Then
    OgeDegeri5.SelStart = 0
    OgeDegeri5.SelLength = Len(OgeDegeri5.Value)
End If

Son:
OgeDegeri5.DropDown
If OgeDegeri5.BackColor = RGB(60, 100, 180) Then
OgeDegeri5.BackColor = RGB(255, 255, 255)
OgeDegeri5.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri6.DropDown
OgeDegeri6.BackColor = RGB(255, 255, 255)
OgeDegeri6.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri7.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri6.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri6.ListIndex = OgeDegeri6.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri6.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri6.ListIndex = OgeDegeri6.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri6_Change()

If OgeDegeri6.ListIndex = -1 And OgeDegeri6.Value <> "" Then
   OgeDegeri6.Value = ""
   GoTo Son
End If

If OgeDegeri6.Value <> "" Then
    OgeDegeri6.SelStart = 0
    OgeDegeri6.SelLength = Len(OgeDegeri6.Value)
End If

Son:
OgeDegeri6.DropDown
If OgeDegeri6.BackColor = RGB(60, 100, 180) Then
OgeDegeri6.BackColor = RGB(255, 255, 255)
OgeDegeri6.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri7.DropDown
OgeDegeri7.BackColor = RGB(255, 255, 255)
OgeDegeri7.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri8.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri7.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri7.ListIndex = OgeDegeri7.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri7.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri7.ListIndex = OgeDegeri7.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri7_Change()

If OgeDegeri7.ListIndex = -1 And OgeDegeri7.Value <> "" Then
   OgeDegeri7.Value = ""
   GoTo Son
End If

If OgeDegeri7.Value <> "" Then
    OgeDegeri7.SelStart = 0
    OgeDegeri7.SelLength = Len(OgeDegeri7.Value)
End If

Son:
OgeDegeri7.DropDown
If OgeDegeri7.BackColor = RGB(60, 100, 180) Then
OgeDegeri7.BackColor = RGB(255, 255, 255)
OgeDegeri7.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri8.DropDown
OgeDegeri8.BackColor = RGB(255, 255, 255)
OgeDegeri8.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri9.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri8.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri8.ListIndex = OgeDegeri8.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri8.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri8.ListIndex = OgeDegeri8.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri8_Change()

If OgeDegeri8.ListIndex = -1 And OgeDegeri8.Value <> "" Then
   OgeDegeri8.Value = ""
   GoTo Son
End If

If OgeDegeri8.Value <> "" Then
    OgeDegeri8.SelStart = 0
    OgeDegeri8.SelLength = Len(OgeDegeri8.Value)
End If

Son:
OgeDegeri8.DropDown
If OgeDegeri8.BackColor = RGB(60, 100, 180) Then
OgeDegeri8.BackColor = RGB(255, 255, 255)
OgeDegeri8.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri9.DropDown
OgeDegeri9.BackColor = RGB(255, 255, 255)
OgeDegeri9.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri10.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri9.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri9.ListIndex = OgeDegeri9.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri9.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri9.ListIndex = OgeDegeri9.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri9_Change()

If OgeDegeri9.ListIndex = -1 And OgeDegeri9.Value <> "" Then
   OgeDegeri9.Value = ""
   GoTo Son
End If

If OgeDegeri9.Value <> "" Then
    OgeDegeri9.SelStart = 0
    OgeDegeri9.SelLength = Len(OgeDegeri9.Value)
End If

Son:
OgeDegeri9.DropDown
If OgeDegeri9.BackColor = RGB(60, 100, 180) Then
OgeDegeri9.BackColor = RGB(255, 255, 255)
OgeDegeri9.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri10.DropDown
OgeDegeri10.BackColor = RGB(255, 255, 255)
OgeDegeri10.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri11.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri10.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri10.ListIndex = OgeDegeri10.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri10.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri10.ListIndex = OgeDegeri10.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri10_Change()

If OgeDegeri10.ListIndex = -1 And OgeDegeri10.Value <> "" Then
   OgeDegeri10.Value = ""
   GoTo Son
End If

If OgeDegeri10.Value <> "" Then
    OgeDegeri10.SelStart = 0
    OgeDegeri10.SelLength = Len(OgeDegeri10.Value)
End If

Son:
OgeDegeri10.DropDown
If OgeDegeri10.BackColor = RGB(60, 100, 180) Then
OgeDegeri10.BackColor = RGB(255, 255, 255)
OgeDegeri10.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri11.DropDown
OgeDegeri11.BackColor = RGB(255, 255, 255)
OgeDegeri11.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri12.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri11.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri11.ListIndex = OgeDegeri11.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri11.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri11.ListIndex = OgeDegeri11.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri11_Change()

If OgeDegeri11.ListIndex = -1 And OgeDegeri11.Value <> "" Then
   OgeDegeri11.Value = ""
   GoTo Son
End If

If OgeDegeri11.Value <> "" Then
    OgeDegeri11.SelStart = 0
    OgeDegeri11.SelLength = Len(OgeDegeri11.Value)
End If

Son:
OgeDegeri11.DropDown
If OgeDegeri11.BackColor = RGB(60, 100, 180) Then
OgeDegeri11.BackColor = RGB(255, 255, 255)
OgeDegeri11.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri12.DropDown
OgeDegeri12.BackColor = RGB(255, 255, 255)
OgeDegeri12.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri13.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri12.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri12.ListIndex = OgeDegeri12.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri12.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri12.ListIndex = OgeDegeri12.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri12_Change()

If OgeDegeri12.ListIndex = -1 And OgeDegeri12.Value <> "" Then
   OgeDegeri12.Value = ""
   GoTo Son
End If

If OgeDegeri12.Value <> "" Then
    OgeDegeri12.SelStart = 0
    OgeDegeri12.SelLength = Len(OgeDegeri12.Value)
End If

Son:
OgeDegeri12.DropDown
If OgeDegeri12.BackColor = RGB(60, 100, 180) Then
OgeDegeri12.BackColor = RGB(255, 255, 255)
OgeDegeri12.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri13.DropDown
OgeDegeri13.BackColor = RGB(255, 255, 255)
OgeDegeri13.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri14.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri13.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri13.ListIndex = OgeDegeri13.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri13.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri13.ListIndex = OgeDegeri13.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri13_Change()

If OgeDegeri13.ListIndex = -1 And OgeDegeri13.Value <> "" Then
   OgeDegeri13.Value = ""
   GoTo Son
End If

If OgeDegeri13.Value <> "" Then
    OgeDegeri13.SelStart = 0
    OgeDegeri13.SelLength = Len(OgeDegeri13.Value)
End If

Son:
OgeDegeri13.DropDown
If OgeDegeri13.BackColor = RGB(60, 100, 180) Then
OgeDegeri13.BackColor = RGB(255, 255, 255)
OgeDegeri13.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri14.DropDown
OgeDegeri14.BackColor = RGB(255, 255, 255)
OgeDegeri14.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri15.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri14.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri14.ListIndex = OgeDegeri14.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri14.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri14.ListIndex = OgeDegeri14.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri14_Change()

If OgeDegeri14.ListIndex = -1 And OgeDegeri14.Value <> "" Then
   OgeDegeri14.Value = ""
   GoTo Son
End If

If OgeDegeri14.Value <> "" Then
    OgeDegeri14.SelStart = 0
    OgeDegeri14.SelLength = Len(OgeDegeri14.Value)
End If

Son:
OgeDegeri14.DropDown
If OgeDegeri14.BackColor = RGB(60, 100, 180) Then
OgeDegeri14.BackColor = RGB(255, 255, 255)
OgeDegeri14.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri15.DropDown
OgeDegeri15.BackColor = RGB(255, 255, 255)
OgeDegeri15.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri16.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri15.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri15.ListIndex = OgeDegeri15.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri15.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri15.ListIndex = OgeDegeri15.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri15_Change()

If OgeDegeri15.ListIndex = -1 And OgeDegeri15.Value <> "" Then
   OgeDegeri15.Value = ""
   GoTo Son
End If

If OgeDegeri15.Value <> "" Then
    OgeDegeri15.SelStart = 0
    OgeDegeri15.SelLength = Len(OgeDegeri15.Value)
End If

Son:
OgeDegeri15.DropDown
If OgeDegeri15.BackColor = RGB(60, 100, 180) Then
OgeDegeri15.BackColor = RGB(255, 255, 255)
OgeDegeri15.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri16.DropDown
OgeDegeri16.BackColor = RGB(255, 255, 255)
OgeDegeri16.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri17.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri16.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri16.ListIndex = OgeDegeri16.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri16.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri16.ListIndex = OgeDegeri16.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri16_Change()

If OgeDegeri16.ListIndex = -1 And OgeDegeri16.Value <> "" Then
   OgeDegeri16.Value = ""
   GoTo Son
End If

If OgeDegeri16.Value <> "" Then
    OgeDegeri16.SelStart = 0
    OgeDegeri16.SelLength = Len(OgeDegeri16.Value)
End If

Son:
OgeDegeri16.DropDown
If OgeDegeri16.BackColor = RGB(60, 100, 180) Then
OgeDegeri16.BackColor = RGB(255, 255, 255)
OgeDegeri16.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri17.DropDown
OgeDegeri17.BackColor = RGB(255, 255, 255)
OgeDegeri17.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri18.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri17.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri17.ListIndex = OgeDegeri17.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri17.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri17.ListIndex = OgeDegeri17.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri17_Change()

If OgeDegeri17.ListIndex = -1 And OgeDegeri17.Value <> "" Then
   OgeDegeri17.Value = ""
   GoTo Son
End If

If OgeDegeri17.Value <> "" Then
    OgeDegeri17.SelStart = 0
    OgeDegeri17.SelLength = Len(OgeDegeri17.Value)
End If

Son:
OgeDegeri17.DropDown
If OgeDegeri17.BackColor = RGB(60, 100, 180) Then
OgeDegeri17.BackColor = RGB(255, 255, 255)
OgeDegeri17.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri18.DropDown
OgeDegeri18.BackColor = RGB(255, 255, 255)
OgeDegeri18.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeDegeri19.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri18.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri18.ListIndex = OgeDegeri18.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri18.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri18.ListIndex = OgeDegeri18.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri18_Change()

If OgeDegeri18.ListIndex = -1 And OgeDegeri18.Value <> "" Then
   OgeDegeri18.Value = ""
   GoTo Son
End If

If OgeDegeri18.Value <> "" Then
    OgeDegeri18.SelStart = 0
    OgeDegeri18.SelLength = Len(OgeDegeri18.Value)
End If

Son:
OgeDegeri18.DropDown
If OgeDegeri18.BackColor = RGB(60, 100, 180) Then
OgeDegeri18.BackColor = RGB(255, 255, 255)
OgeDegeri18.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub OgeDegeri19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.OgeDegeri19.DropDown
OgeDegeri19.BackColor = RGB(255, 255, 255)
OgeDegeri19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub OgeDegeri19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeDegeri18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'OgeDegeri20.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If OgeDegeri19.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri19.ListIndex = OgeDegeri19.ListIndex
            End If
        Case 40 'Down
            If OgeDegeri19.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                OgeDegeri19.ListIndex = OgeDegeri19.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub OgeDegeri19_Change()

If OgeDegeri19.ListIndex = -1 And OgeDegeri19.Value <> "" Then
   OgeDegeri19.Value = ""
   GoTo Son
End If

If OgeDegeri19.Value <> "" Then
    OgeDegeri19.SelStart = 0
    OgeDegeri19.SelLength = Len(OgeDegeri19.Value)
End If

Son:
OgeDegeri19.DropDown
If OgeDegeri19.BackColor = RGB(60, 100, 180) Then
OgeDegeri19.BackColor = RGB(255, 255, 255)
OgeDegeri19.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Adet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet.BackColor = RGB(255, 255, 255)
Adet.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub Adet_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet1.SetFocus
    End If
    
End Sub

Private Sub Adet_Change()
    If Adet.Value <> "" And IsNumeric(Adet.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet.Value = ""
    End If
Adet.BackColor = RGB(255, 255, 255)
Adet.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet1.BackColor = RGB(255, 255, 255)
Adet1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet2.SetFocus
    End If
    
End Sub
Private Sub Adet1_Change()
    If Adet1.Value <> "" And IsNumeric(Adet1.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet1.Value = ""
    End If
Adet1.BackColor = RGB(255, 255, 255)
Adet1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet2.BackColor = RGB(255, 255, 255)
Adet2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet3.SetFocus
    End If
    
End Sub
Private Sub Adet2_Change()
    If Adet2.Value <> "" And IsNumeric(Adet2.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet2.Value = ""
    End If
Adet2.BackColor = RGB(255, 255, 255)
Adet2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet3.BackColor = RGB(255, 255, 255)
Adet3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet4.SetFocus
    End If
    
End Sub
Private Sub Adet3_Change()
    If Adet3.Value <> "" And IsNumeric(Adet3.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet3.Value = ""
    End If
Adet3.BackColor = RGB(255, 255, 255)
Adet3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet4.BackColor = RGB(255, 255, 255)
Adet4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet5.SetFocus
    End If
    
End Sub
Private Sub Adet4_Change()
    If Adet4.Value <> "" And IsNumeric(Adet4.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet4.Value = ""
    End If
Adet4.BackColor = RGB(255, 255, 255)
Adet4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet5.BackColor = RGB(255, 255, 255)
Adet5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet6.SetFocus
    End If
    
End Sub
Private Sub Adet5_Change()
    If Adet5.Value <> "" And IsNumeric(Adet5.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet5.Value = ""
    End If
Adet5.BackColor = RGB(255, 255, 255)
Adet5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet6.BackColor = RGB(255, 255, 255)
Adet6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet7.SetFocus
    End If
    
End Sub
Private Sub Adet6_Change()
    If Adet6.Value <> "" And IsNumeric(Adet6.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet6.Value = ""
    End If
Adet6.BackColor = RGB(255, 255, 255)
Adet6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet7.BackColor = RGB(255, 255, 255)
Adet7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet8.SetFocus
    End If
    
End Sub
Private Sub Adet7_Change()
    If Adet7.Value <> "" And IsNumeric(Adet7.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet7.Value = ""
    End If
Adet7.BackColor = RGB(255, 255, 255)
Adet7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet8.BackColor = RGB(255, 255, 255)
Adet8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet9.SetFocus
    End If
    
End Sub
Private Sub Adet8_Change()
    If Adet8.Value <> "" And IsNumeric(Adet8.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet8.Value = ""
    End If
Adet8.BackColor = RGB(255, 255, 255)
Adet8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet9.BackColor = RGB(255, 255, 255)
Adet9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet10.SetFocus
    End If
    
End Sub
Private Sub Adet9_Change()
    If Adet9.Value <> "" And IsNumeric(Adet9.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet9.Value = ""
    End If
Adet9.BackColor = RGB(255, 255, 255)
Adet9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet10.BackColor = RGB(255, 255, 255)
Adet10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet11.SetFocus
    End If
    
End Sub
Private Sub Adet10_Change()
    If Adet10.Value <> "" And IsNumeric(Adet10.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet10.Value = ""
    End If
Adet10.BackColor = RGB(255, 255, 255)
Adet10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet11.BackColor = RGB(255, 255, 255)
Adet11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet12.SetFocus
    End If
    
End Sub
Private Sub Adet11_Change()
    If Adet11.Value <> "" And IsNumeric(Adet11.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet11.Value = ""
    End If
Adet11.BackColor = RGB(255, 255, 255)
Adet11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet12.BackColor = RGB(255, 255, 255)
Adet12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet13.SetFocus
    End If
    
End Sub
Private Sub Adet12_Change()
    If Adet12.Value <> "" And IsNumeric(Adet12.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet12.Value = ""
    End If
Adet12.BackColor = RGB(255, 255, 255)
Adet12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet13.BackColor = RGB(255, 255, 255)
Adet13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet14.SetFocus
    End If
    
End Sub
Private Sub Adet13_Change()
    If Adet13.Value <> "" And IsNumeric(Adet13.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet13.Value = ""
    End If
Adet13.BackColor = RGB(255, 255, 255)
Adet13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet14.BackColor = RGB(255, 255, 255)
Adet14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet15.SetFocus
    End If
    
End Sub
Private Sub Adet14_Change()
    If Adet14.Value <> "" And IsNumeric(Adet14.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet14.Value = ""
    End If
Adet14.BackColor = RGB(255, 255, 255)
Adet14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet15.BackColor = RGB(255, 255, 255)
Adet15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet16.SetFocus
    End If
    
End Sub
Private Sub Adet15_Change()
    If Adet15.Value <> "" And IsNumeric(Adet15.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet15.Value = ""
    End If
Adet15.BackColor = RGB(255, 255, 255)
Adet15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet16.BackColor = RGB(255, 255, 255)
Adet16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet17.SetFocus
    End If
    
End Sub
Private Sub Adet16_Change()
    If Adet16.Value <> "" And IsNumeric(Adet16.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet16.Value = ""
    End If
Adet16.BackColor = RGB(255, 255, 255)
Adet16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet17.BackColor = RGB(255, 255, 255)
Adet17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet18.SetFocus
    End If
    
End Sub
Private Sub Adet17_Change()
    If Adet17.Value <> "" And IsNumeric(Adet17.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet17.Value = ""
    End If
Adet17.BackColor = RGB(255, 255, 255)
Adet17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet18.BackColor = RGB(255, 255, 255)
Adet18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Adet19.SetFocus
    End If
    
End Sub
Private Sub Adet18_Change()
    If Adet18.Value <> "" And IsNumeric(Adet18.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet18.Value = ""
    End If
Adet18.BackColor = RGB(255, 255, 255)
Adet18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adet19.BackColor = RGB(255, 255, 255)
Adet19.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adet19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Adet18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'Adet20.SetFocus
    End If
    
End Sub
Private Sub Adet19_Change()
    If Adet19.Value <> "" And IsNumeric(Adet19.Value) = False Then
        MsgBox "The quantity field cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        Adet19.Value = ""
    End If
Adet19.BackColor = RGB(255, 255, 255)
Adet19.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub OgeIdNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo.BackColor = RGB(255, 255, 255)
OgeIdNo.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo_Change()
OgeIdNo.BackColor = RGB(255, 255, 255)
OgeIdNo.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub OgeIdNo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo1.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo1.BackColor = RGB(255, 255, 255)
OgeIdNo1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo1_Change()
OgeIdNo1.BackColor = RGB(255, 255, 255)
OgeIdNo1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo2.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo2.BackColor = RGB(255, 255, 255)
OgeIdNo2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo2_Change()
OgeIdNo2.BackColor = RGB(255, 255, 255)
OgeIdNo2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo3.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo3.BackColor = RGB(255, 255, 255)
OgeIdNo3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo3_Change()
OgeIdNo3.BackColor = RGB(255, 255, 255)
OgeIdNo3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo4.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo4.BackColor = RGB(255, 255, 255)
OgeIdNo4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo4_Change()
OgeIdNo4.BackColor = RGB(255, 255, 255)
OgeIdNo4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo5.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo5.BackColor = RGB(255, 255, 255)
OgeIdNo5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo5_Change()
OgeIdNo5.BackColor = RGB(255, 255, 255)
OgeIdNo5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo6.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo6.BackColor = RGB(255, 255, 255)
OgeIdNo6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo6_Change()
OgeIdNo6.BackColor = RGB(255, 255, 255)
OgeIdNo6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo7.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo7.BackColor = RGB(255, 255, 255)
OgeIdNo7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo7_Change()
OgeIdNo7.BackColor = RGB(255, 255, 255)
OgeIdNo7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo8.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo8.BackColor = RGB(255, 255, 255)
OgeIdNo8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo8_Change()
OgeIdNo8.BackColor = RGB(255, 255, 255)
OgeIdNo8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo9.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo9.BackColor = RGB(255, 255, 255)
OgeIdNo9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo9_Change()
OgeIdNo9.BackColor = RGB(255, 255, 255)
OgeIdNo9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo10.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo10.BackColor = RGB(255, 255, 255)
OgeIdNo10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo10_Change()
OgeIdNo10.BackColor = RGB(255, 255, 255)
OgeIdNo10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo11.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo11.BackColor = RGB(255, 255, 255)
OgeIdNo11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo11_Change()
OgeIdNo11.BackColor = RGB(255, 255, 255)
OgeIdNo11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo12.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo12.BackColor = RGB(255, 255, 255)
OgeIdNo12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo12_Change()
OgeIdNo12.BackColor = RGB(255, 255, 255)
OgeIdNo12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo13.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo13.BackColor = RGB(255, 255, 255)
OgeIdNo13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo13_Change()
OgeIdNo13.BackColor = RGB(255, 255, 255)
OgeIdNo13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo14.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo14.BackColor = RGB(255, 255, 255)
OgeIdNo14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo14_Change()
OgeIdNo14.BackColor = RGB(255, 255, 255)
OgeIdNo14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo15.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo15.BackColor = RGB(255, 255, 255)
OgeIdNo15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo15_Change()
OgeIdNo15.BackColor = RGB(255, 255, 255)
OgeIdNo15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo16.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo16.BackColor = RGB(255, 255, 255)
OgeIdNo16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo16_Change()
OgeIdNo16.BackColor = RGB(255, 255, 255)
OgeIdNo16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo17.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo17.BackColor = RGB(255, 255, 255)
OgeIdNo17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo17_Change()
OgeIdNo17.BackColor = RGB(255, 255, 255)
OgeIdNo17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo18.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo18.BackColor = RGB(255, 255, 255)
OgeIdNo18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo18_Change()
OgeIdNo18.BackColor = RGB(255, 255, 255)
OgeIdNo18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        OgeIdNo19.SetFocus
    End If
    
End Sub
Private Sub OgeIdNo19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
OgeIdNo19.BackColor = RGB(255, 255, 255)
OgeIdNo19.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo19_Change()
OgeIdNo19.BackColor = RGB(255, 255, 255)
OgeIdNo19.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub OgeIdNo19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        OgeIdNo18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'OgeIdNo20.SetFocus
    End If
    
End Sub
Private Sub Aciklama_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama.BackColor = RGB(255, 255, 255)
Aciklama.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama_Change()
Aciklama.BackColor = RGB(255, 255, 255)
Aciklama.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama1.SetFocus
    End If
    
End Sub
Private Sub Aciklama1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama1.BackColor = RGB(255, 255, 255)
Aciklama1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama1_Change()
Aciklama1.BackColor = RGB(255, 255, 255)
Aciklama1.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama2.SetFocus
    End If
    
End Sub
Private Sub Aciklama2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama2.BackColor = RGB(255, 255, 255)
Aciklama2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama2_Change()
Aciklama2.BackColor = RGB(255, 255, 255)
Aciklama2.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama3.SetFocus
    End If
    
End Sub
Private Sub Aciklama3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama3.BackColor = RGB(255, 255, 255)
Aciklama3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama3_Change()
Aciklama3.BackColor = RGB(255, 255, 255)
Aciklama3.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama4.SetFocus
    End If
    
End Sub
Private Sub Aciklama4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama4.BackColor = RGB(255, 255, 255)
Aciklama4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama4_Change()
Aciklama4.BackColor = RGB(255, 255, 255)
Aciklama4.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama5.SetFocus
    End If
    
End Sub
Private Sub Aciklama5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama5.BackColor = RGB(255, 255, 255)
Aciklama5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama5_Change()
Aciklama5.BackColor = RGB(255, 255, 255)
Aciklama5.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama6.SetFocus
    End If
    
End Sub
Private Sub Aciklama6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama6.BackColor = RGB(255, 255, 255)
Aciklama6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama6_Change()
Aciklama6.BackColor = RGB(255, 255, 255)
Aciklama6.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama7.SetFocus
    End If
    
End Sub
Private Sub Aciklama7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama7.BackColor = RGB(255, 255, 255)
Aciklama7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama7_Change()
Aciklama7.BackColor = RGB(255, 255, 255)
Aciklama7.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama8.SetFocus
    End If
    
End Sub
Private Sub Aciklama8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama8.BackColor = RGB(255, 255, 255)
Aciklama8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama8_Change()
Aciklama8.BackColor = RGB(255, 255, 255)
Aciklama8.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama9.SetFocus
    End If
    
End Sub
Private Sub Aciklama9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama9.BackColor = RGB(255, 255, 255)
Aciklama9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama9_Change()
Aciklama9.BackColor = RGB(255, 255, 255)
Aciklama9.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama10.SetFocus
    End If
    
End Sub
Private Sub Aciklama10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama10.BackColor = RGB(255, 255, 255)
Aciklama10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama10_Change()
Aciklama10.BackColor = RGB(255, 255, 255)
Aciklama10.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama11.SetFocus
    End If
    
End Sub
Private Sub Aciklama11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama11.BackColor = RGB(255, 255, 255)
Aciklama11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama11_Change()
Aciklama11.BackColor = RGB(255, 255, 255)
Aciklama11.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama12.SetFocus
    End If
    
End Sub
Private Sub Aciklama12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama12.BackColor = RGB(255, 255, 255)
Aciklama12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama12_Change()
Aciklama12.BackColor = RGB(255, 255, 255)
Aciklama12.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama13.SetFocus
    End If
    
End Sub
Private Sub Aciklama13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama13.BackColor = RGB(255, 255, 255)
Aciklama13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama13_Change()
Aciklama13.BackColor = RGB(255, 255, 255)
Aciklama13.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama14.SetFocus
    End If
    
End Sub
Private Sub Aciklama14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama14.BackColor = RGB(255, 255, 255)
Aciklama14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama14_Change()
Aciklama14.BackColor = RGB(255, 255, 255)
Aciklama14.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama15.SetFocus
    End If
    
End Sub
Private Sub Aciklama15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama15.BackColor = RGB(255, 255, 255)
Aciklama15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama15_Change()
Aciklama15.BackColor = RGB(255, 255, 255)
Aciklama15.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama16.SetFocus
    End If
    
End Sub
Private Sub Aciklama16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama16.BackColor = RGB(255, 255, 255)
Aciklama16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama16_Change()
Aciklama16.BackColor = RGB(255, 255, 255)
Aciklama16.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama17.SetFocus
    End If
    
End Sub
Private Sub Aciklama17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama17.BackColor = RGB(255, 255, 255)
Aciklama17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama17_Change()
Aciklama17.BackColor = RGB(255, 255, 255)
Aciklama17.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama18.SetFocus
    End If
    
End Sub
Private Sub Aciklama18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama18.BackColor = RGB(255, 255, 255)
Aciklama18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama18_Change()
Aciklama18.BackColor = RGB(255, 255, 255)
Aciklama18.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Aciklama19.SetFocus
    End If
    
End Sub
Private Sub Aciklama19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Aciklama19.BackColor = RGB(255, 255, 255)
Aciklama19.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama19_Change()
Aciklama19.BackColor = RGB(255, 255, 255)
Aciklama19.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Aciklama19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Aciklama18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'Aciklama20.SetFocus
    End If
    
End Sub

Private Sub Sonuc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc.DropDown
Sonuc.BackColor = RGB(255, 255, 255)
Sonuc.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc1.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc.ListIndex = Sonuc.ListIndex
            End If
        Case 40 'Down
            If Sonuc.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc.ListIndex = Sonuc.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc_Change()

If Sonuc.ListIndex = -1 And Sonuc.Value <> "" Then
   Sonuc.Value = ""
   GoTo Son
End If

If Sonuc.Value <> "" Then
    Sonuc.SelStart = 0
    Sonuc.SelLength = Len(Sonuc.Value)
End If

Son:

Sonuc.DropDown
If Sonuc.BackColor = RGB(60, 100, 180) Then
Sonuc.BackColor = RGB(255, 255, 255)
Sonuc.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc1.DropDown
Sonuc1.BackColor = RGB(255, 255, 255)
Sonuc1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc2.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc1.ListIndex = Sonuc1.ListIndex
            End If
        Case 40 'Down
            If Sonuc1.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc1.ListIndex = Sonuc1.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc1_Change()

If Sonuc1.ListIndex = -1 And Sonuc1.Value <> "" Then
   Sonuc1.Value = ""
   GoTo Son
End If

If Sonuc1.Value <> "" Then
    Sonuc1.SelStart = 0
    Sonuc1.SelLength = Len(Sonuc1.Value)
End If

Son:

Sonuc1.DropDown
If Sonuc1.BackColor = RGB(60, 100, 180) Then
Sonuc1.BackColor = RGB(255, 255, 255)
Sonuc1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc2.DropDown
Sonuc2.BackColor = RGB(255, 255, 255)
Sonuc2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc3.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc2.ListIndex = Sonuc2.ListIndex
            End If
        Case 40 'Down
            If Sonuc2.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc2.ListIndex = Sonuc2.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc2_Change()

If Sonuc2.ListIndex = -1 And Sonuc2.Value <> "" Then
   Sonuc2.Value = ""
   GoTo Son
End If

If Sonuc2.Value <> "" Then
    Sonuc2.SelStart = 0
    Sonuc2.SelLength = Len(Sonuc2.Value)
End If

Son:

Sonuc2.DropDown
If Sonuc2.BackColor = RGB(60, 100, 180) Then
Sonuc2.BackColor = RGB(255, 255, 255)
Sonuc2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc3.DropDown
Sonuc3.BackColor = RGB(255, 255, 255)
Sonuc3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc4.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc3.ListIndex = Sonuc3.ListIndex
            End If
        Case 40 'Down
            If Sonuc3.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc3.ListIndex = Sonuc3.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc3_Change()

If Sonuc3.ListIndex = -1 And Sonuc3.Value <> "" Then
   Sonuc3.Value = ""
   GoTo Son
End If

If Sonuc3.Value <> "" Then
    Sonuc3.SelStart = 0
    Sonuc3.SelLength = Len(Sonuc3.Value)
End If

Son:

Sonuc3.DropDown
If Sonuc3.BackColor = RGB(60, 100, 180) Then
Sonuc3.BackColor = RGB(255, 255, 255)
Sonuc3.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc4.DropDown
Sonuc4.BackColor = RGB(255, 255, 255)
Sonuc4.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc5.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc4.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc4.ListIndex = Sonuc4.ListIndex
            End If
        Case 40 'Down
            If Sonuc4.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc4.ListIndex = Sonuc4.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc4_Change()

If Sonuc4.ListIndex = -1 And Sonuc4.Value <> "" Then
   Sonuc4.Value = ""
   GoTo Son
End If

If Sonuc4.Value <> "" Then
    Sonuc4.SelStart = 0
    Sonuc4.SelLength = Len(Sonuc4.Value)
End If

Son:

Sonuc4.DropDown
If Sonuc4.BackColor = RGB(60, 100, 180) Then
Sonuc4.BackColor = RGB(255, 255, 255)
Sonuc4.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc5.DropDown
Sonuc5.BackColor = RGB(255, 255, 255)
Sonuc5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc6.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc5.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc5.ListIndex = Sonuc5.ListIndex
            End If
        Case 40 'Down
            If Sonuc5.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc5.ListIndex = Sonuc5.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc5_Change()

If Sonuc5.ListIndex = -1 And Sonuc5.Value <> "" Then
   Sonuc5.Value = ""
   GoTo Son
End If

If Sonuc5.Value <> "" Then
    Sonuc5.SelStart = 0
    Sonuc5.SelLength = Len(Sonuc5.Value)
End If

Son:

Sonuc5.DropDown
If Sonuc5.BackColor = RGB(60, 100, 180) Then
Sonuc5.BackColor = RGB(255, 255, 255)
Sonuc5.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc6.DropDown
Sonuc6.BackColor = RGB(255, 255, 255)
Sonuc6.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc7.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc6.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc6.ListIndex = Sonuc6.ListIndex
            End If
        Case 40 'Down
            If Sonuc6.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc6.ListIndex = Sonuc6.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc6_Change()

If Sonuc6.ListIndex = -1 And Sonuc6.Value <> "" Then
   Sonuc6.Value = ""
   GoTo Son
End If

If Sonuc6.Value <> "" Then
    Sonuc6.SelStart = 0
    Sonuc6.SelLength = Len(Sonuc6.Value)
End If

Son:

Sonuc6.DropDown
If Sonuc6.BackColor = RGB(60, 100, 180) Then
Sonuc6.BackColor = RGB(255, 255, 255)
Sonuc6.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc7.DropDown
Sonuc7.BackColor = RGB(255, 255, 255)
Sonuc7.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc8.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc7.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc7.ListIndex = Sonuc7.ListIndex
            End If
        Case 40 'Down
            If Sonuc7.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc7.ListIndex = Sonuc7.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc7_Change()

If Sonuc7.ListIndex = -1 And Sonuc7.Value <> "" Then
   Sonuc7.Value = ""
   GoTo Son
End If

If Sonuc7.Value <> "" Then
    Sonuc7.SelStart = 0
    Sonuc7.SelLength = Len(Sonuc7.Value)
End If

Son:

Sonuc7.DropDown
If Sonuc7.BackColor = RGB(60, 100, 180) Then
Sonuc7.BackColor = RGB(255, 255, 255)
Sonuc7.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc8.DropDown
Sonuc8.BackColor = RGB(255, 255, 255)
Sonuc8.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc9.SetFocus
    End If
    
    Select Case KeyCode
        Case 38  'Up
            If Sonuc8.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc8.ListIndex = Sonuc8.ListIndex
            End If
        Case 40 'Down
            If Sonuc8.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc8.ListIndex = Sonuc8.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc8_Change()

If Sonuc8.ListIndex = -1 And Sonuc8.Value <> "" Then
   Sonuc8.Value = ""
   GoTo Son
End If

If Sonuc8.Value <> "" Then
    Sonuc8.SelStart = 0
    Sonuc8.SelLength = Len(Sonuc8.Value)
End If

Son:

Sonuc8.DropDown
If Sonuc8.BackColor = RGB(60, 100, 180) Then
Sonuc8.BackColor = RGB(255, 255, 255)
Sonuc8.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc9.DropDown
Sonuc9.BackColor = RGB(255, 255, 255)
Sonuc9.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc10.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc9.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc9.ListIndex = Sonuc9.ListIndex
            End If
        Case 40 'Down
            If Sonuc9.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc9.ListIndex = Sonuc9.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc9_Change()

If Sonuc9.ListIndex = -1 And Sonuc9.Value <> "" Then
   Sonuc9.Value = ""
   GoTo Son
End If

If Sonuc9.Value <> "" Then
    Sonuc9.SelStart = 0
    Sonuc9.SelLength = Len(Sonuc9.Value)
End If

Son:

Sonuc9.DropDown
If Sonuc9.BackColor = RGB(60, 100, 180) Then
Sonuc9.BackColor = RGB(255, 255, 255)
Sonuc9.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc10.DropDown
Sonuc10.BackColor = RGB(255, 255, 255)
Sonuc10.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc11.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc10.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc10.ListIndex = Sonuc10.ListIndex
            End If
        Case 40 'Down
            If Sonuc10.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc10.ListIndex = Sonuc10.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc10_Change()

If Sonuc10.ListIndex = -1 And Sonuc10.Value <> "" Then
   Sonuc10.Value = ""
   GoTo Son
End If

If Sonuc10.Value <> "" Then
    Sonuc10.SelStart = 0
    Sonuc10.SelLength = Len(Sonuc10.Value)
End If

Son:

Sonuc10.DropDown
If Sonuc10.BackColor = RGB(60, 100, 180) Then
Sonuc10.BackColor = RGB(255, 255, 255)
Sonuc10.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc11.DropDown
Sonuc11.BackColor = RGB(255, 255, 255)
Sonuc11.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc12.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc11.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc11.ListIndex = Sonuc11.ListIndex
            End If
        Case 40 'Down
            If Sonuc11.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc11.ListIndex = Sonuc11.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc11_Change()

If Sonuc11.ListIndex = -1 And Sonuc11.Value <> "" Then
   Sonuc11.Value = ""
   GoTo Son
End If

If Sonuc11.Value <> "" Then
    Sonuc11.SelStart = 0
    Sonuc11.SelLength = Len(Sonuc11.Value)
End If

Son:

Sonuc11.DropDown
If Sonuc11.BackColor = RGB(60, 100, 180) Then
Sonuc11.BackColor = RGB(255, 255, 255)
Sonuc11.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc12.DropDown
Sonuc12.BackColor = RGB(255, 255, 255)
Sonuc12.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc13.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc12.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc12.ListIndex = Sonuc12.ListIndex
            End If
        Case 40 'Down
            If Sonuc12.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc12.ListIndex = Sonuc12.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc12_Change()

If Sonuc12.ListIndex = -1 And Sonuc12.Value <> "" Then
   Sonuc12.Value = ""
   GoTo Son
End If

If Sonuc12.Value <> "" Then
    Sonuc12.SelStart = 0
    Sonuc12.SelLength = Len(Sonuc12.Value)
End If

Son:

Sonuc12.DropDown
If Sonuc12.BackColor = RGB(60, 100, 180) Then
Sonuc12.BackColor = RGB(255, 255, 255)
Sonuc12.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc13.DropDown
Sonuc13.BackColor = RGB(255, 255, 255)
Sonuc13.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc14.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc13.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc13.ListIndex = Sonuc13.ListIndex
            End If
        Case 40 'Down
            If Sonuc13.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc13.ListIndex = Sonuc13.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc13_Change()

If Sonuc13.ListIndex = -1 And Sonuc13.Value <> "" Then
   Sonuc13.Value = ""
   GoTo Son
End If

If Sonuc13.Value <> "" Then
    Sonuc13.SelStart = 0
    Sonuc13.SelLength = Len(Sonuc13.Value)
End If

Son:

Sonuc13.DropDown
If Sonuc13.BackColor = RGB(60, 100, 180) Then
Sonuc13.BackColor = RGB(255, 255, 255)
Sonuc13.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc14.DropDown
Sonuc14.BackColor = RGB(255, 255, 255)
Sonuc14.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc15.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc14.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc14.ListIndex = Sonuc14.ListIndex
            End If
        Case 40 'Down
            If Sonuc14.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc14.ListIndex = Sonuc14.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc14_Change()

If Sonuc14.ListIndex = -1 And Sonuc14.Value <> "" Then
   Sonuc14.Value = ""
   GoTo Son
End If

If Sonuc14.Value <> "" Then
    Sonuc14.SelStart = 0
    Sonuc14.SelLength = Len(Sonuc14.Value)
End If

Son:

Sonuc14.DropDown
If Sonuc14.BackColor = RGB(60, 100, 180) Then
Sonuc14.BackColor = RGB(255, 255, 255)
Sonuc14.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc15.DropDown
Sonuc15.BackColor = RGB(255, 255, 255)
Sonuc15.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc16.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc15.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc15.ListIndex = Sonuc15.ListIndex
            End If
        Case 40 'Down
            If Sonuc15.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc15.ListIndex = Sonuc15.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc15_Change()

If Sonuc15.ListIndex = -1 And Sonuc15.Value <> "" Then
   Sonuc15.Value = ""
   GoTo Son
End If

If Sonuc15.Value <> "" Then
    Sonuc15.SelStart = 0
    Sonuc15.SelLength = Len(Sonuc15.Value)
End If

Son:

Sonuc15.DropDown
If Sonuc15.BackColor = RGB(60, 100, 180) Then
Sonuc15.BackColor = RGB(255, 255, 255)
Sonuc15.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc16.DropDown
Sonuc16.BackColor = RGB(255, 255, 255)
Sonuc16.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc17.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc16.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc16.ListIndex = Sonuc16.ListIndex
            End If
        Case 40 'Down
            If Sonuc16.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc16.ListIndex = Sonuc16.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc16_Change()

If Sonuc16.ListIndex = -1 And Sonuc16.Value <> "" Then
   Sonuc16.Value = ""
   GoTo Son
End If

If Sonuc16.Value <> "" Then
    Sonuc16.SelStart = 0
    Sonuc16.SelLength = Len(Sonuc16.Value)
End If

Son:

Sonuc16.DropDown
If Sonuc16.BackColor = RGB(60, 100, 180) Then
Sonuc16.BackColor = RGB(255, 255, 255)
Sonuc16.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc17.DropDown
Sonuc17.BackColor = RGB(255, 255, 255)
Sonuc17.ForeColor = RGB(30, 30, 30)


End Sub

Private Sub Sonuc17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc18.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc17.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc17.ListIndex = Sonuc17.ListIndex
            End If
        Case 40 'Down
            If Sonuc17.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc17.ListIndex = Sonuc17.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc17_Change()

If Sonuc17.ListIndex = -1 And Sonuc17.Value <> "" Then
   Sonuc17.Value = ""
   GoTo Son
End If

If Sonuc17.Value <> "" Then
    Sonuc17.SelStart = 0
    Sonuc17.SelLength = Len(Sonuc17.Value)
End If

Son:

Sonuc17.DropDown
If Sonuc17.BackColor = RGB(60, 100, 180) Then
Sonuc17.BackColor = RGB(255, 255, 255)
Sonuc17.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub Sonuc18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc18.DropDown
Sonuc18.BackColor = RGB(255, 255, 255)
Sonuc18.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Sonuc19.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc18.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc18.ListIndex = Sonuc18.ListIndex
            End If
        Case 40 'Down
            If Sonuc18.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc18.ListIndex = Sonuc18.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc18_Change()

If Sonuc18.ListIndex = -1 And Sonuc18.Value <> "" Then
   Sonuc18.Value = ""
   GoTo Son
End If

If Sonuc18.Value <> "" Then
    Sonuc18.SelStart = 0
    Sonuc18.SelLength = Len(Sonuc18.Value)
End If

Son:

Sonuc18.DropDown
If Sonuc18.BackColor = RGB(60, 100, 180) Then
Sonuc18.BackColor = RGB(255, 255, 255)
Sonuc18.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Sonuc19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Sonuc19.DropDown
Sonuc19.BackColor = RGB(255, 255, 255)
Sonuc19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Sonuc19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Sonuc18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'Sonuc20.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If Sonuc19.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc19.ListIndex = Sonuc19.ListIndex
            End If
        Case 40 'Down
            If Sonuc19.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Sonuc19.ListIndex = Sonuc19.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub Sonuc19_Change()

If Sonuc19.ListIndex = -1 And Sonuc19.Value <> "" Then
   Sonuc19.Value = ""
   GoTo Son
End If

If Sonuc19.Value <> "" Then
    Sonuc19.SelStart = 0
    Sonuc19.SelLength = Len(Sonuc19.Value)
End If

Son:

Sonuc19.DropDown
If Sonuc19.BackColor = RGB(60, 100, 180) Then
Sonuc19.BackColor = RGB(255, 255, 255)
Sonuc19.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub Rapor1No_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No.DropDown
Rapor1No.BackColor = RGB(255, 255, 255)
Rapor1No.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1No_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No1.SetFocus
    End If
    
End Sub

Private Sub Rapor1No_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No.Value Then
'        Rapor1No.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No.Value, "/") > 0 Or InStr(Rapor1No.Value, "\") > 0 Or InStr(Rapor1No.Value, "<") > 0 Or InStr(Rapor1No.Value, ">") > 0 Or InStr(Rapor1No.Value, ":") > 0 Or InStr(Rapor1No.Value, "*") > 0 Or InStr(Rapor1No.Value, "?") > 0 Or InStr(Rapor1No.Value, "|") > 0 Or InStr(Rapor1No.Value, """") > 0 Or InStr(Rapor1No.Value, "[") > 0 Or InStr(Rapor1No.Value, "]") > 0 Or InStr(Rapor1No.Value, "_") > 0 Or InStr(Rapor1No.Value, "(") > 0 Or InStr(Rapor1No.Value, ")") > 0 Or InStr(Rapor1No.Value, ".") > 0 Or InStr(Rapor1No.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No.Value = ""
End If
    
    
'Boşluklara izin verme
For j = 1 To 20
Rapor1No.Value = Replace(Rapor1No.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No.Value = UCase(Replace(Replace(Rapor1No.Value, "ı", "I"), "i", "I"))

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No.Value, i, 1)) = False And Mid(Rapor1No.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No.DropDown
Rapor1No.BackColor = RGB(255, 255, 255)
Rapor1No.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No1.DropDown
Rapor1No1.BackColor = RGB(255, 255, 255)
Rapor1No1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1No1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No2.SetFocus
    End If
    
End Sub

Private Sub Rapor1No1_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No1.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No1.Value Then
'        Rapor1No1.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No1.Value, "/") > 0 Or InStr(Rapor1No1.Value, "\") > 0 Or InStr(Rapor1No1.Value, "<") > 0 Or InStr(Rapor1No1.Value, ">") > 0 Or InStr(Rapor1No1.Value, ":") > 0 Or InStr(Rapor1No1.Value, "*") > 0 Or InStr(Rapor1No1.Value, "?") > 0 Or InStr(Rapor1No1.Value, "|") > 0 Or InStr(Rapor1No1.Value, """") > 0 Or InStr(Rapor1No1.Value, "[") > 0 Or InStr(Rapor1No1.Value, "]") > 0 Or InStr(Rapor1No1.Value, "_") > 0 Or InStr(Rapor1No1.Value, "(") > 0 Or InStr(Rapor1No1.Value, ")") > 0 Or InStr(Rapor1No1.Value, ".") > 0 Or InStr(Rapor1No1.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No1.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No1.Value = Replace(Rapor1No1.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No1.Value = UCase(Replace(Replace(Rapor1No1.Value, "ı", "I"), "i", "I"))

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No1.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No1.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No1.Value, i, 1)) = False And Mid(Rapor1No1.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No1.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No1.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No1.DropDown
Rapor1No1.BackColor = RGB(255, 255, 255)
Rapor1No1.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No2.DropDown
Rapor1No2.BackColor = RGB(255, 255, 255)
Rapor1No2.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No3.SetFocus
    End If
    
End Sub

Private Sub Rapor1No2_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No2.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No2.Value Then
'        Rapor1No2.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No2.Value, "/") > 0 Or InStr(Rapor1No2.Value, "\") > 0 Or InStr(Rapor1No2.Value, "<") > 0 Or InStr(Rapor1No2.Value, ">") > 0 Or InStr(Rapor1No2.Value, ":") > 0 Or InStr(Rapor1No2.Value, "*") > 0 Or InStr(Rapor1No2.Value, "?") > 0 Or InStr(Rapor1No2.Value, "|") > 0 Or InStr(Rapor1No2.Value, """") > 0 Or InStr(Rapor1No2.Value, "[") > 0 Or InStr(Rapor1No2.Value, "]") > 0 Or InStr(Rapor1No2.Value, "_") > 0 Or InStr(Rapor1No2.Value, "(") > 0 Or InStr(Rapor1No2.Value, ")") > 0 Or InStr(Rapor1No2.Value, ".") > 0 Or InStr(Rapor1No2.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No2.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No2.Value = Replace(Rapor1No2.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No2.Value = UCase(Replace(Replace(Rapor1No2.Value, "ı", "I"), "i", "I"))

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No2.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No2.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No2.Value, i, 1)) = False And Mid(Rapor1No2.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No2.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No2.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No2.DropDown
Rapor1No2.BackColor = RGB(255, 255, 255)
Rapor1No2.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No3.DropDown
Rapor1No3.BackColor = RGB(255, 255, 255)
Rapor1No3.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No4.SetFocus
    End If
    
End Sub

Private Sub Rapor1No3_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No3.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No3.Value Then
'        Rapor1No3.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No3.Value, "/") > 0 Or InStr(Rapor1No3.Value, "\") > 0 Or InStr(Rapor1No3.Value, "<") > 0 Or InStr(Rapor1No3.Value, ">") > 0 Or InStr(Rapor1No3.Value, ":") > 0 Or InStr(Rapor1No3.Value, "*") > 0 Or InStr(Rapor1No3.Value, "?") > 0 Or InStr(Rapor1No3.Value, "|") > 0 Or InStr(Rapor1No3.Value, """") > 0 Or InStr(Rapor1No3.Value, "[") > 0 Or InStr(Rapor1No3.Value, "]") > 0 Or InStr(Rapor1No3.Value, "_") > 0 Or InStr(Rapor1No3.Value, "(") > 0 Or InStr(Rapor1No3.Value, ")") > 0 Or InStr(Rapor1No3.Value, ".") > 0 Or InStr(Rapor1No3.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No3.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No3.Value = Replace(Rapor1No3.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No3.Value = UCase(Replace(Replace(Rapor1No3.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No3.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No3.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No3.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No3.Value, i, 1)) = False And Mid(Rapor1No3.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No3.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No3.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No3.BackColor = RGB(255, 255, 255)
Rapor1No3.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No4.DropDown
Rapor1No4.BackColor = RGB(255, 255, 255)
Rapor1No4.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No5.SetFocus
    End If
    
End Sub

Private Sub Rapor1No4_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No4.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No4.Value Then
'        Rapor1No4.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No4.Value, "/") > 0 Or InStr(Rapor1No4.Value, "\") > 0 Or InStr(Rapor1No4.Value, "<") > 0 Or InStr(Rapor1No4.Value, ">") > 0 Or InStr(Rapor1No4.Value, ":") > 0 Or InStr(Rapor1No4.Value, "*") > 0 Or InStr(Rapor1No4.Value, "?") > 0 Or InStr(Rapor1No4.Value, "|") > 0 Or InStr(Rapor1No4.Value, """") > 0 Or InStr(Rapor1No4.Value, "[") > 0 Or InStr(Rapor1No4.Value, "]") > 0 Or InStr(Rapor1No4.Value, "_") > 0 Or InStr(Rapor1No4.Value, "(") > 0 Or InStr(Rapor1No4.Value, ")") > 0 Or InStr(Rapor1No4.Value, ".") > 0 Or InStr(Rapor1No4.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No4.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No4.Value = Replace(Rapor1No4.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No4.Value = UCase(Replace(Replace(Rapor1No4.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No4.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No4.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No4.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No4.Value, i, 1)) = False And Mid(Rapor1No4.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No4.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No4.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No4.BackColor = RGB(255, 255, 255)
Rapor1No4.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No5.DropDown
Rapor1No5.BackColor = RGB(255, 255, 255)
Rapor1No5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1No5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No6.SetFocus
    End If
    
End Sub

Private Sub Rapor1No5_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No5.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No5.Value Then
'        Rapor1No5.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No5.Value, "/") > 0 Or InStr(Rapor1No5.Value, "\") > 0 Or InStr(Rapor1No5.Value, "<") > 0 Or InStr(Rapor1No5.Value, ">") > 0 Or InStr(Rapor1No5.Value, ":") > 0 Or InStr(Rapor1No5.Value, "*") > 0 Or InStr(Rapor1No5.Value, "?") > 0 Or InStr(Rapor1No5.Value, "|") > 0 Or InStr(Rapor1No5.Value, """") > 0 Or InStr(Rapor1No5.Value, "[") > 0 Or InStr(Rapor1No5.Value, "]") > 0 Or InStr(Rapor1No5.Value, "_") > 0 Or InStr(Rapor1No5.Value, "(") > 0 Or InStr(Rapor1No5.Value, ")") > 0 Or InStr(Rapor1No5.Value, ".") > 0 Or InStr(Rapor1No5.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No5.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No5.Value = Replace(Rapor1No5.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No5.Value = UCase(Replace(Replace(Rapor1No5.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No5.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No5.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No5.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No5.Value, i, 1)) = False And Mid(Rapor1No5.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No5.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No5.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No5.BackColor = RGB(255, 255, 255)
Rapor1No5.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No6.DropDown
Rapor1No6.BackColor = RGB(255, 255, 255)
Rapor1No6.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No7.SetFocus
    End If
    
End Sub

Private Sub Rapor1No6_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No6.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No6.Value Then
'        Rapor1No6.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No6.Value, "/") > 0 Or InStr(Rapor1No6.Value, "\") > 0 Or InStr(Rapor1No6.Value, "<") > 0 Or InStr(Rapor1No6.Value, ">") > 0 Or InStr(Rapor1No6.Value, ":") > 0 Or InStr(Rapor1No6.Value, "*") > 0 Or InStr(Rapor1No6.Value, "?") > 0 Or InStr(Rapor1No6.Value, "|") > 0 Or InStr(Rapor1No6.Value, """") > 0 Or InStr(Rapor1No6.Value, "[") > 0 Or InStr(Rapor1No6.Value, "]") > 0 Or InStr(Rapor1No6.Value, "_") > 0 Or InStr(Rapor1No6.Value, "(") > 0 Or InStr(Rapor1No6.Value, ")") > 0 Or InStr(Rapor1No6.Value, ".") > 0 Or InStr(Rapor1No6.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No6.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No6.Value = Replace(Rapor1No6.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No6.Value = UCase(Replace(Replace(Rapor1No6.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No6.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No6.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No6.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No6.Value, i, 1)) = False And Mid(Rapor1No6.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No6.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No6.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No6.BackColor = RGB(255, 255, 255)
Rapor1No6.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No7.DropDown
Rapor1No7.BackColor = RGB(255, 255, 255)
Rapor1No7.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No8.SetFocus
    End If
    
End Sub

Private Sub Rapor1No7_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No7.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No7.Value Then
'        Rapor1No7.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No7.Value, "/") > 0 Or InStr(Rapor1No7.Value, "\") > 0 Or InStr(Rapor1No7.Value, "<") > 0 Or InStr(Rapor1No7.Value, ">") > 0 Or InStr(Rapor1No7.Value, ":") > 0 Or InStr(Rapor1No7.Value, "*") > 0 Or InStr(Rapor1No7.Value, "?") > 0 Or InStr(Rapor1No7.Value, "|") > 0 Or InStr(Rapor1No7.Value, """") > 0 Or InStr(Rapor1No7.Value, "[") > 0 Or InStr(Rapor1No7.Value, "]") > 0 Or InStr(Rapor1No7.Value, "_") > 0 Or InStr(Rapor1No7.Value, "(") > 0 Or InStr(Rapor1No7.Value, ")") > 0 Or InStr(Rapor1No7.Value, ".") > 0 Or InStr(Rapor1No7.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No7.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No7.Value = Replace(Rapor1No7.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No7.Value = UCase(Replace(Replace(Rapor1No7.Value, "ı", "I"), "i", "I"))

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No7.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No7.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No7.Value, i, 1)) = False And Mid(Rapor1No7.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No7.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No7.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No7.DropDown
Rapor1No7.BackColor = RGB(255, 255, 255)
Rapor1No7.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No8.DropDown
Rapor1No8.BackColor = RGB(255, 255, 255)
Rapor1No8.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No9.SetFocus
    End If
    
End Sub

Private Sub Rapor1No8_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No8.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No8.Value Then
'        Rapor1No8.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No8.Value, "/") > 0 Or InStr(Rapor1No8.Value, "\") > 0 Or InStr(Rapor1No8.Value, "<") > 0 Or InStr(Rapor1No8.Value, ">") > 0 Or InStr(Rapor1No8.Value, ":") > 0 Or InStr(Rapor1No8.Value, "*") > 0 Or InStr(Rapor1No8.Value, "?") > 0 Or InStr(Rapor1No8.Value, "|") > 0 Or InStr(Rapor1No8.Value, """") > 0 Or InStr(Rapor1No8.Value, "[") > 0 Or InStr(Rapor1No8.Value, "]") > 0 Or InStr(Rapor1No8.Value, "_") > 0 Or InStr(Rapor1No8.Value, "(") > 0 Or InStr(Rapor1No8.Value, ")") > 0 Or InStr(Rapor1No8.Value, ".") > 0 Or InStr(Rapor1No8.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No8.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No8.Value = Replace(Rapor1No8.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No8.Value = UCase(Replace(Replace(Rapor1No8.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No8.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No8.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No8.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No8.Value, i, 1)) = False And Mid(Rapor1No8.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No8.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No8.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No8.BackColor = RGB(255, 255, 255)
Rapor1No8.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No9.DropDown
Rapor1No9.BackColor = RGB(255, 255, 255)
Rapor1No9.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No10.SetFocus
    End If
    
End Sub

Private Sub Rapor1No9_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No9.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No9.Value Then
'        Rapor1No9.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No9.Value, "/") > 0 Or InStr(Rapor1No9.Value, "\") > 0 Or InStr(Rapor1No9.Value, "<") > 0 Or InStr(Rapor1No9.Value, ">") > 0 Or InStr(Rapor1No9.Value, ":") > 0 Or InStr(Rapor1No9.Value, "*") > 0 Or InStr(Rapor1No9.Value, "?") > 0 Or InStr(Rapor1No9.Value, "|") > 0 Or InStr(Rapor1No9.Value, """") > 0 Or InStr(Rapor1No9.Value, "[") > 0 Or InStr(Rapor1No9.Value, "]") > 0 Or InStr(Rapor1No9.Value, "_") > 0 Or InStr(Rapor1No9.Value, "(") > 0 Or InStr(Rapor1No9.Value, ")") > 0 Or InStr(Rapor1No9.Value, ".") > 0 Or InStr(Rapor1No9.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No9.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No9.Value = Replace(Rapor1No9.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No9.Value = UCase(Replace(Replace(Rapor1No9.Value, "ı", "I"), "i", "I"))

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No9.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No9.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No9.Value, i, 1)) = False And Mid(Rapor1No9.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No9.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No9.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No9.DropDown
Rapor1No9.BackColor = RGB(255, 255, 255)
Rapor1No9.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No10.DropDown
Rapor1No10.BackColor = RGB(255, 255, 255)
Rapor1No10.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No11.SetFocus
    End If
    
End Sub

Private Sub Rapor1No10_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No10.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No10.Value Then
'        Rapor1No10.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No10.Value, "/") > 0 Or InStr(Rapor1No10.Value, "\") > 0 Or InStr(Rapor1No10.Value, "<") > 0 Or InStr(Rapor1No10.Value, ">") > 0 Or InStr(Rapor1No10.Value, ":") > 0 Or InStr(Rapor1No10.Value, "*") > 0 Or InStr(Rapor1No10.Value, "?") > 0 Or InStr(Rapor1No10.Value, "|") > 0 Or InStr(Rapor1No10.Value, """") > 0 Or InStr(Rapor1No10.Value, "[") > 0 Or InStr(Rapor1No10.Value, "]") > 0 Or InStr(Rapor1No10.Value, "_") > 0 Or InStr(Rapor1No10.Value, "(") > 0 Or InStr(Rapor1No10.Value, ")") > 0 Or InStr(Rapor1No10.Value, ".") > 0 Or InStr(Rapor1No10.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No10.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No10.Value = Replace(Rapor1No10.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No10.Value = UCase(Replace(Replace(Rapor1No10.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No10.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No10.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No10.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No10.Value, i, 1)) = False And Mid(Rapor1No10.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No10.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No10.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No10.BackColor = RGB(255, 255, 255)
Rapor1No10.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No11.DropDown
Rapor1No11.BackColor = RGB(255, 255, 255)
Rapor1No11.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No12.SetFocus
    End If
    
End Sub

Private Sub Rapor1No11_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No11.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No11.Value Then
'        Rapor1No11.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No11.Value, "/") > 0 Or InStr(Rapor1No11.Value, "\") > 0 Or InStr(Rapor1No11.Value, "<") > 0 Or InStr(Rapor1No11.Value, ">") > 0 Or InStr(Rapor1No11.Value, ":") > 0 Or InStr(Rapor1No11.Value, "*") > 0 Or InStr(Rapor1No11.Value, "?") > 0 Or InStr(Rapor1No11.Value, "|") > 0 Or InStr(Rapor1No11.Value, """") > 0 Or InStr(Rapor1No11.Value, "[") > 0 Or InStr(Rapor1No11.Value, "]") > 0 Or InStr(Rapor1No11.Value, "_") > 0 Or InStr(Rapor1No11.Value, "(") > 0 Or InStr(Rapor1No11.Value, ")") > 0 Or InStr(Rapor1No11.Value, ".") > 0 Or InStr(Rapor1No11.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No11.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No11.Value = Replace(Rapor1No11.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No11.Value = UCase(Replace(Replace(Rapor1No11.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No11.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No11.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No11.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No11.Value, i, 1)) = False And Mid(Rapor1No11.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No11.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No11.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No11.BackColor = RGB(255, 255, 255)
Rapor1No11.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No12.DropDown
Rapor1No12.BackColor = RGB(255, 255, 255)
Rapor1No12.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No13.SetFocus
    End If
    
End Sub

Private Sub Rapor1No12_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No12.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No12.Value Then
'        Rapor1No12.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No12.Value, "/") > 0 Or InStr(Rapor1No12.Value, "\") > 0 Or InStr(Rapor1No12.Value, "<") > 0 Or InStr(Rapor1No12.Value, ">") > 0 Or InStr(Rapor1No12.Value, ":") > 0 Or InStr(Rapor1No12.Value, "*") > 0 Or InStr(Rapor1No12.Value, "?") > 0 Or InStr(Rapor1No12.Value, "|") > 0 Or InStr(Rapor1No12.Value, """") > 0 Or InStr(Rapor1No12.Value, "[") > 0 Or InStr(Rapor1No12.Value, "]") > 0 Or InStr(Rapor1No12.Value, "_") > 0 Or InStr(Rapor1No12.Value, "(") > 0 Or InStr(Rapor1No12.Value, ")") > 0 Or InStr(Rapor1No12.Value, ".") > 0 Or InStr(Rapor1No12.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No12.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No12.Value = Replace(Rapor1No12.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No12.Value = UCase(Replace(Replace(Rapor1No12.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No12.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No12.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No12.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No12.Value, i, 1)) = False And Mid(Rapor1No12.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No12.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No12.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No12.BackColor = RGB(255, 255, 255)
Rapor1No12.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No13.DropDown
Rapor1No13.BackColor = RGB(255, 255, 255)
Rapor1No13.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No14.SetFocus
    End If
    
End Sub

Private Sub Rapor1No13_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No13.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No13.Value Then
'        Rapor1No13.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No13.Value, "/") > 0 Or InStr(Rapor1No13.Value, "\") > 0 Or InStr(Rapor1No13.Value, "<") > 0 Or InStr(Rapor1No13.Value, ">") > 0 Or InStr(Rapor1No13.Value, ":") > 0 Or InStr(Rapor1No13.Value, "*") > 0 Or InStr(Rapor1No13.Value, "?") > 0 Or InStr(Rapor1No13.Value, "|") > 0 Or InStr(Rapor1No13.Value, """") > 0 Or InStr(Rapor1No13.Value, "[") > 0 Or InStr(Rapor1No13.Value, "]") > 0 Or InStr(Rapor1No13.Value, "_") > 0 Or InStr(Rapor1No13.Value, "(") > 0 Or InStr(Rapor1No13.Value, ")") > 0 Or InStr(Rapor1No13.Value, ".") > 0 Or InStr(Rapor1No13.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No13.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No13.Value = Replace(Rapor1No13.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No13.Value = UCase(Replace(Replace(Rapor1No13.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No13.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No13.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No13.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No13.Value, i, 1)) = False And Mid(Rapor1No13.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No13.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No13.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No13.BackColor = RGB(255, 255, 255)
Rapor1No13.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No14.DropDown
Rapor1No14.BackColor = RGB(255, 255, 255)
Rapor1No14.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No15.SetFocus
    End If
    
End Sub

Private Sub Rapor1No14_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No14.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No14.Value Then
'        Rapor1No14.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No14.Value, "/") > 0 Or InStr(Rapor1No14.Value, "\") > 0 Or InStr(Rapor1No14.Value, "<") > 0 Or InStr(Rapor1No14.Value, ">") > 0 Or InStr(Rapor1No14.Value, ":") > 0 Or InStr(Rapor1No14.Value, "*") > 0 Or InStr(Rapor1No14.Value, "?") > 0 Or InStr(Rapor1No14.Value, "|") > 0 Or InStr(Rapor1No14.Value, """") > 0 Or InStr(Rapor1No14.Value, "[") > 0 Or InStr(Rapor1No14.Value, "]") > 0 Or InStr(Rapor1No14.Value, "_") > 0 Or InStr(Rapor1No14.Value, "(") > 0 Or InStr(Rapor1No14.Value, ")") > 0 Or InStr(Rapor1No14.Value, ".") > 0 Or InStr(Rapor1No14.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No14.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No14.Value = Replace(Rapor1No14.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No14.Value = UCase(Replace(Replace(Rapor1No14.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No14.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No14.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No14.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No14.Value, i, 1)) = False And Mid(Rapor1No14.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No14.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No14.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No14.BackColor = RGB(255, 255, 255)
Rapor1No14.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No15.DropDown
Rapor1No15.BackColor = RGB(255, 255, 255)
Rapor1No15.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No16.SetFocus
    End If
    
End Sub

Private Sub Rapor1No15_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No15.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No15.Value Then
'        Rapor1No15.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No15.Value, "/") > 0 Or InStr(Rapor1No15.Value, "\") > 0 Or InStr(Rapor1No15.Value, "<") > 0 Or InStr(Rapor1No15.Value, ">") > 0 Or InStr(Rapor1No15.Value, ":") > 0 Or InStr(Rapor1No15.Value, "*") > 0 Or InStr(Rapor1No15.Value, "?") > 0 Or InStr(Rapor1No15.Value, "|") > 0 Or InStr(Rapor1No15.Value, """") > 0 Or InStr(Rapor1No15.Value, "[") > 0 Or InStr(Rapor1No15.Value, "]") > 0 Or InStr(Rapor1No15.Value, "_") > 0 Or InStr(Rapor1No15.Value, "(") > 0 Or InStr(Rapor1No15.Value, ")") > 0 Or InStr(Rapor1No15.Value, ".") > 0 Or InStr(Rapor1No15.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No15.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No15.Value = Replace(Rapor1No15.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No15.Value = UCase(Replace(Replace(Rapor1No15.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No15.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No15.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No15.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No15.Value, i, 1)) = False And Mid(Rapor1No15.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No15.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No15.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No15.BackColor = RGB(255, 255, 255)
Rapor1No15.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No16.DropDown
Rapor1No16.BackColor = RGB(255, 255, 255)
Rapor1No16.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No17.SetFocus
    End If
    
End Sub

Private Sub Rapor1No16_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No16.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No16.Value Then
'        Rapor1No16.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No16.Value, "/") > 0 Or InStr(Rapor1No16.Value, "\") > 0 Or InStr(Rapor1No16.Value, "<") > 0 Or InStr(Rapor1No16.Value, ">") > 0 Or InStr(Rapor1No16.Value, ":") > 0 Or InStr(Rapor1No16.Value, "*") > 0 Or InStr(Rapor1No16.Value, "?") > 0 Or InStr(Rapor1No16.Value, "|") > 0 Or InStr(Rapor1No16.Value, """") > 0 Or InStr(Rapor1No16.Value, "[") > 0 Or InStr(Rapor1No16.Value, "]") > 0 Or InStr(Rapor1No16.Value, "_") > 0 Or InStr(Rapor1No16.Value, "(") > 0 Or InStr(Rapor1No16.Value, ")") > 0 Or InStr(Rapor1No16.Value, ".") > 0 Or InStr(Rapor1No16.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No16.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No16.Value = Replace(Rapor1No16.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No16.Value = UCase(Replace(Replace(Rapor1No16.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No16.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No16.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No16.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No16.Value, i, 1)) = False And Mid(Rapor1No16.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No16.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No16.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No16.BackColor = RGB(255, 255, 255)
Rapor1No16.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No17.DropDown
Rapor1No17.BackColor = RGB(255, 255, 255)
Rapor1No17.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No18.SetFocus
    End If
    
End Sub

Private Sub Rapor1No17_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No17.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No17.Value Then
'        Rapor1No17.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No17.Value, "/") > 0 Or InStr(Rapor1No17.Value, "\") > 0 Or InStr(Rapor1No17.Value, "<") > 0 Or InStr(Rapor1No17.Value, ">") > 0 Or InStr(Rapor1No17.Value, ":") > 0 Or InStr(Rapor1No17.Value, "*") > 0 Or InStr(Rapor1No17.Value, "?") > 0 Or InStr(Rapor1No17.Value, "|") > 0 Or InStr(Rapor1No17.Value, """") > 0 Or InStr(Rapor1No17.Value, "[") > 0 Or InStr(Rapor1No17.Value, "]") > 0 Or InStr(Rapor1No17.Value, "_") > 0 Or InStr(Rapor1No17.Value, "(") > 0 Or InStr(Rapor1No17.Value, ")") > 0 Or InStr(Rapor1No17.Value, ".") > 0 Or InStr(Rapor1No17.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No17.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No17.Value = Replace(Rapor1No17.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No17.Value = UCase(Replace(Replace(Rapor1No17.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No17.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No17.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No17.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No17.Value, i, 1)) = False And Mid(Rapor1No17.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No17.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No17.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No17.BackColor = RGB(255, 255, 255)
Rapor1No17.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1No18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No18.DropDown
Rapor1No18.BackColor = RGB(255, 255, 255)
Rapor1No18.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No19.SetFocus
    End If
    
End Sub

Private Sub Rapor1No18_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No18.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No18.Value Then
'        Rapor1No18.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No18.Value, "/") > 0 Or InStr(Rapor1No18.Value, "\") > 0 Or InStr(Rapor1No18.Value, "<") > 0 Or InStr(Rapor1No18.Value, ">") > 0 Or InStr(Rapor1No18.Value, ":") > 0 Or InStr(Rapor1No18.Value, "*") > 0 Or InStr(Rapor1No18.Value, "?") > 0 Or InStr(Rapor1No18.Value, "|") > 0 Or InStr(Rapor1No18.Value, """") > 0 Or InStr(Rapor1No18.Value, "[") > 0 Or InStr(Rapor1No18.Value, "]") > 0 Or InStr(Rapor1No18.Value, "_") > 0 Or InStr(Rapor1No18.Value, "(") > 0 Or InStr(Rapor1No18.Value, ")") > 0 Or InStr(Rapor1No18.Value, ".") > 0 Or InStr(Rapor1No18.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No18.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No18.Value = Replace(Rapor1No18.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No18.Value = UCase(Replace(Replace(Rapor1No18.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No18.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No18.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No18.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No18.Value, i, 1)) = False And Mid(Rapor1No18.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No18.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No18.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No18.BackColor = RGB(255, 255, 255)
Rapor1No18.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.Rapor1No19.DropDown
Rapor1No19.BackColor = RGB(255, 255, 255)
Rapor1No19.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub Rapor1No19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        Rapor1No18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        Rapor1No19.SetFocus
    End If
    
End Sub

Private Sub Rapor1No19_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

'If ComboGetir.Value = "" Then
''Comboda tanımlı değer girilemez.
'a() = Rapor1No19.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = Rapor1No19.Value Then
'        Rapor1No19.Value = ""
'    End If
'Next i
'End If

'Kullanılamaz karakterler...
If InStr(Rapor1No19.Value, "/") > 0 Or InStr(Rapor1No19.Value, "\") > 0 Or InStr(Rapor1No19.Value, "<") > 0 Or InStr(Rapor1No19.Value, ">") > 0 Or InStr(Rapor1No19.Value, ":") > 0 Or InStr(Rapor1No19.Value, "*") > 0 Or InStr(Rapor1No19.Value, "?") > 0 Or InStr(Rapor1No19.Value, "|") > 0 Or InStr(Rapor1No19.Value, """") > 0 Or InStr(Rapor1No19.Value, "[") > 0 Or InStr(Rapor1No19.Value, "]") > 0 Or InStr(Rapor1No19.Value, "_") > 0 Or InStr(Rapor1No19.Value, "(") > 0 Or InStr(Rapor1No19.Value, ")") > 0 Or InStr(Rapor1No19.Value, ".") > 0 Or InStr(Rapor1No19.Value, ",") > 0 Then
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 1 number. Please avoid using any of these characters in the Report 1 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No19.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No19.Value = Replace(Rapor1No19.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No19.Value = UCase(Replace(Replace(Rapor1No19.Value, "ı", "I"), "i", "I"))
'Me.Rapor1No19.DropDown

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No19.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No19.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No19.Value, i, 1)) = False And Mid(Rapor1No19.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No19.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No19.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 1 number, except for the dash (-). The required prefix for the Report 1 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

Rapor1No19.BackColor = RGB(255, 255, 255)
Rapor1No19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1TarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        Rapor1TarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        Rapor1TarihiText.Value = ""
    End If
    
End Sub

Private Sub Rapor1TarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Rapor1TarihiText.BackColor = RGB(255, 255, 255)
Rapor1TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Rapor1TarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    Rapor1TarihiText.Value = CalTarih
    Rapor1TarihiText.Value = Format(Rapor1TarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

Rapor1TarihiText.BackColor = RGB(255, 255, 255)
Rapor1TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak2TarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        Tutanak2TarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        Tutanak2TarihiText.Value = ""
    End If
    
End Sub

Private Sub Tutanak2TarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Tutanak2TarihiText.BackColor = RGB(255, 255, 255)
Tutanak2TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub Tutanak2TarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    Tutanak2TarihiText.Value = CalTarih
    Tutanak2TarihiText.Value = Format(Tutanak2TarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

Tutanak2TarihiText.BackColor = RGB(255, 255, 255)
Tutanak2TarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GidenMuhatapTemasi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GidenMuhatapTemasi.DropDown
GidenMuhatapTemasi.BackColor = RGB(255, 255, 255)
GidenMuhatapTemasi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GidenMuhatapTemasi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GidenMuhatapTemasi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenMuhatapTemasi.ListIndex = GidenMuhatapTemasi.ListIndex - 1
            End If
            Me.GidenMuhatapTemasi.DropDown
            
        Case 40 'Aşağı
            If GidenMuhatapTemasi.ListIndex = GidenMuhatapTemasi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenMuhatapTemasi.ListIndex = GidenMuhatapTemasi.ListIndex + 1
            End If
            Me.GidenMuhatapTemasi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GidenMuhatapTemasi_Change()

If GidenMuhatapTemasi.ListIndex = -1 And GidenMuhatapTemasi.Value <> "" Then
   GidenMuhatapTemasi.Value = ""
   GoTo Son
End If

If GidenMuhatapTemasi.Value <> "" Then
    GidenMuhatapTemasi.SelStart = 0
    GidenMuhatapTemasi.SelLength = Len(GidenMuhatapTemasi.Value)
End If


Son:

GidenMuhatapTemasi.DropDown
If GidenMuhatapTemasi.BackColor = RGB(60, 100, 180) Then
GidenMuhatapTemasi.BackColor = RGB(255, 255, 255)
GidenMuhatapTemasi.ForeColor = RGB(30, 30, 30)
End If

End Sub
Private Sub GonderilenBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GonderilenBirim.DropDown
GonderilenBirim.BackColor = RGB(255, 255, 255)
GonderilenBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GonderilenBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GonderilenBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GonderilenBirim.ListIndex = GonderilenBirim.ListIndex - 1
            End If
            Me.GonderilenBirim.DropDown
            
        Case 40 'Aşağı
            If GonderilenBirim.ListIndex = GonderilenBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GonderilenBirim.ListIndex = GonderilenBirim.ListIndex + 1
            End If
            Me.GonderilenBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GonderilenBirim_Change()

If GonderilenBirim.ListIndex = -1 And GonderilenBirim.Value <> "" Then
   GonderilenBirim.Value = ""
   GoTo Son
End If

If GonderilenBirim.Value <> "" Then
    GonderilenBirim.SelStart = 0
    GonderilenBirim.SelLength = Len(GonderilenBirim.Value)
End If

Son:

GonderilenBirim.DropDown
If GonderilenBirim.BackColor = RGB(60, 100, 180) Then
GonderilenBirim.BackColor = RGB(255, 255, 255)
GonderilenBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub GidenPaketTipi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GidenPaketTipi.DropDown

GidenPaketTipi.BackColor = RGB(255, 255, 255)
GidenPaketTipi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GidenPaketTipi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GidenPaketTipi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenPaketTipi.ListIndex = GidenPaketTipi.ListIndex - 1
            End If
            Me.GidenPaketTipi.DropDown
            
        Case 40 'Aşağı
            If GidenPaketTipi.ListIndex = GidenPaketTipi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenPaketTipi.ListIndex = GidenPaketTipi.ListIndex + 1
            End If
            Me.GidenPaketTipi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GidenPaketTipi_Change()

If GidenPaketTipi.ListIndex = -1 And GidenPaketTipi.Value <> "" Then
   GidenPaketTipi.Value = ""
   GoTo Son
End If

If GidenPaketTipi.Value <> "" Then
    GidenPaketTipi.SelStart = 0
    GidenPaketTipi.SelLength = Len(GidenPaketTipi.Value)
End If

Son:

GidenPaketTipi.DropDown
If GidenPaketTipi.BackColor = RGB(60, 100, 180) Then
GidenPaketTipi.BackColor = RGB(255, 255, 255)
GidenPaketTipi.ForeColor = RGB(30, 30, 30)
End If

End Sub
Private Sub GidenPaketAdedi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.GidenPaketAdedi.DropDown
GidenPaketAdedi.BackColor = RGB(255, 255, 255)
GidenPaketAdedi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub GidenPaketAdedi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If GidenPaketAdedi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenPaketAdedi.ListIndex = GidenPaketAdedi.ListIndex - 1
            End If
            Me.GidenPaketAdedi.DropDown
            
        Case 40 'Aşağı
            If GidenPaketAdedi.ListIndex = GidenPaketAdedi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                GidenPaketAdedi.ListIndex = GidenPaketAdedi.ListIndex + 1
            End If
            Me.GidenPaketAdedi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub GidenPaketAdedi_Change()

If GidenPaketAdedi.ListIndex = -1 And GidenPaketAdedi.Value <> "" Then
   GidenPaketAdedi.Value = ""
   GoTo Son
End If

If GidenPaketAdedi.Value <> "" Then
    GidenPaketAdedi.SelStart = 0
    GidenPaketAdedi.SelLength = Len(GidenPaketAdedi.Value)
End If

Son:

GidenPaketAdedi.DropDown
If GidenPaketAdedi.BackColor = RGB(60, 100, 180) Then
GidenPaketAdedi.BackColor = RGB(255, 255, 255)
GidenPaketAdedi.ForeColor = RGB(30, 30, 30)
End If

End Sub
Private Sub UstYaziTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        UstYaziTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        UstYaziTarihiText.Value = ""
    End If
    
End Sub

Private Sub UstYaziTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

UstYaziTarihiText.BackColor = RGB(255, 255, 255)
UstYaziTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UstYaziTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    UstYaziTarihiText.Value = CalTarih
    UstYaziTarihiText.Value = Format(UstYaziTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

UstYaziTarihiText.BackColor = RGB(255, 255, 255)
UstYaziTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UstYaziNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UstYaziNoText.BackColor = RGB(255, 255, 255)
UstYaziNoText.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub UstYaziNoText_Change()
UstYaziNoText.BackColor = RGB(255, 255, 255)
UstYaziNoText.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub IlgiYaziFotokopisi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.IlgiYaziFotokopisi.DropDown
IlgiYaziFotokopisi.BackColor = RGB(255, 255, 255)
IlgiYaziFotokopisi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub IlgiYaziFotokopisi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If IlgiYaziFotokopisi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlgiYaziFotokopisi.ListIndex = IlgiYaziFotokopisi.ListIndex - 1
            End If
            Me.IlgiYaziFotokopisi.DropDown
            
        Case 40 'Aşağı
            If IlgiYaziFotokopisi.ListIndex = IlgiYaziFotokopisi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlgiYaziFotokopisi.ListIndex = IlgiYaziFotokopisi.ListIndex + 1
            End If
            Me.IlgiYaziFotokopisi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub IlgiYaziFotokopisi_Change()

If IlgiYaziFotokopisi.ListIndex = -1 And IlgiYaziFotokopisi.Value <> "" Then
   IlgiYaziFotokopisi.Value = ""
   GoTo Son
End If

If IlgiYaziFotokopisi.Value <> "" Then
    IlgiYaziFotokopisi.SelStart = 0
    IlgiYaziFotokopisi.SelLength = Len(IlgiYaziFotokopisi.Value)
End If

Son:

IlgiYaziFotokopisi.DropDown
If IlgiYaziFotokopisi.BackColor = RGB(60, 100, 180) Then
IlgiYaziFotokopisi.BackColor = RGB(255, 255, 255)
IlgiYaziFotokopisi.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub EkleOge_Click()

If OgeTuruFrame1.Visible = False Then
    OgeTuruFrame1.Visible = True
    GoTo Son
ElseIf OgeTuruFrame2.Visible = False Then
    OgeTuruFrame2.Visible = True
    GoTo Son
ElseIf OgeTuruFrame3.Visible = False Then
    OgeTuruFrame3.Visible = True
    GoTo Son
ElseIf OgeTuruFrame4.Visible = False Then
    OgeTuruFrame4.Visible = True
    GoTo Son
ElseIf OgeTuruFrame5.Visible = False Then
    OgeTuruFrame5.Visible = True
    GoTo Son
ElseIf OgeTuruFrame6.Visible = False Then
    OgeTuruFrame6.Visible = True
    GoTo Son
ElseIf OgeTuruFrame7.Visible = False Then
    OgeTuruFrame7.Visible = True
    GoTo Son
ElseIf OgeTuruFrame8.Visible = False Then
    OgeTuruFrame8.Visible = True
    GoTo Son
ElseIf OgeTuruFrame9.Visible = False Then
    OgeTuruFrame9.Visible = True
    GoTo Son
ElseIf OgeTuruFrame10.Visible = False Then
    OgeTuruFrame10.Visible = True
    GoTo Son
ElseIf OgeTuruFrame11.Visible = False Then
    OgeTuruFrame11.Visible = True
    GoTo Son
ElseIf OgeTuruFrame12.Visible = False Then
    OgeTuruFrame12.Visible = True
    GoTo Son
ElseIf OgeTuruFrame13.Visible = False Then
    OgeTuruFrame13.Visible = True
    GoTo Son
ElseIf OgeTuruFrame14.Visible = False Then
    OgeTuruFrame14.Visible = True
    GoTo Son
ElseIf OgeTuruFrame15.Visible = False Then
    OgeTuruFrame15.Visible = True
    GoTo Son
ElseIf OgeTuruFrame16.Visible = False Then
    OgeTuruFrame16.Visible = True
    GoTo Son
ElseIf OgeTuruFrame17.Visible = False Then
    OgeTuruFrame17.Visible = True
    GoTo Son
ElseIf OgeTuruFrame18.Visible = False Then
    OgeTuruFrame18.Visible = True
    GoTo Son
ElseIf OgeTuruFrame19.Visible = False Then
    OgeTuruFrame19.Visible = True
    GoTo Son
End If

Son:

If ScrollTakip < 456 Then
    If ScrollTakip = 120 Then
        ScrollTakip = ScrollTakip + 24 + 6
    ElseIf ScrollTakip > 120 Then
        ScrollTakip = ScrollTakip + 24
    Else
        ScrollTakip = ScrollTakip + 24
    End If
End If
If ScrollTakip > 120 + 6 Then
    Call SetScrollHook(Me.ScrollFrame, ScrollTakip, 24)
    ScrollFrame.ScrollTop = ScrollTakip
End If

End Sub

Private Sub KaldirOge_Click()
'Çoğaltılan diğer satırlar için de verilerin silinmesi eklenecek.

'_____________________Güncelleme 18112019 1238

Dim OgeFrame As Integer, SonDoluSatir As Integer, IlkBosSatir As Integer


SonDoluSatir = 0
For OgeFrame = 19 To 1 Step -1
    If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
        Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
        Controls("Adet" & OgeFrame).Value <> "" Or _
        Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
        Controls("Aciklama" & OgeFrame).Value <> "" Or _
        Controls("Sonuc" & OgeFrame).Value <> "" Or _
        Controls("Rapor1No" & OgeFrame).Value <> "" Then

        SonDoluSatir = OgeFrame
        GoTo SonDoluSatirBulundu

    End If
Next OgeFrame
SonDoluSatirBulundu:

If SonDoluSatir = 0 Then
 GoTo NormalProsedureGit
End If

'Başlangıç satırı boşsa
If OgeTuru.Value = "" And OgeDegeri.Value = "" And Adet.Value = "" And OgeIdNo.Value = "" And Aciklama.Value = "" And Sonuc.Value = "" And Rapor1No.Value = "" Then

    For OgeFrame = 1 To SonDoluSatir

        If OgeFrame = 1 Then
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Or _
                Controls("Sonuc" & OgeFrame).Value <> "" Or _
                Controls("Rapor1No" & OgeFrame).Value <> "" Then

                OgeTuru.Value = Controls("OgeTuru" & OgeFrame).Value
                OgeDegeri.Value = Controls("OgeDegeri" & OgeFrame).Value
                Adet.Value = Controls("Adet" & OgeFrame).Value
                OgeIdNo.Value = Controls("OgeIdNo" & OgeFrame).Value
                Aciklama.Value = Controls("Aciklama" & OgeFrame).Value
                Sonuc.Value = Controls("Sonuc" & OgeFrame).Value
                Rapor1No.Value = Controls("Rapor1No" & OgeFrame).Value

                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""
                Controls("Sonuc" & OgeFrame).Value = ""
                Controls("Rapor1No" & OgeFrame).Value = ""

            End If
        Else
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Or _
                Controls("Sonuc" & OgeFrame).Value <> "" Or _
                Controls("Rapor1No" & OgeFrame).Value <> "" Then

                Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
                Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
                Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
                Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
                Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value
                Controls("Sonuc" & OgeFrame - 1).Value = Controls("Sonuc" & OgeFrame).Value
                Controls("Rapor1No" & OgeFrame - 1).Value = Controls("Rapor1No" & OgeFrame).Value

                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""
                Controls("Sonuc" & OgeFrame).Value = ""
                Controls("Rapor1No" & OgeFrame).Value = ""

            End If
        End If
    Next OgeFrame

End If

'Başlangıç satırından sonraki işlemler boşsa
IlkBosSatir = 0
For OgeFrame = 1 To SonDoluSatir
    If Controls("OgeTuru" & OgeFrame).Value = "" And _
        Controls("OgeDegeri" & OgeFrame).Value = "" And _
        Controls("Adet" & OgeFrame).Value = "" And _
        Controls("OgeIdNo" & OgeFrame).Value = "" And _
        Controls("Aciklama" & OgeFrame).Value = "" And _
        Controls("Sonuc" & OgeFrame).Value = "" And _
        Controls("Rapor1No" & OgeFrame).Value = "" Then

        IlkBosSatir = OgeFrame
        GoTo IlkBosSatirBulundu

    End If
Next OgeFrame
IlkBosSatirBulundu:

If IlkBosSatir = 0 Then
 GoTo NormalProsedureGit
End If

'Başlangıç satırından sonraki işlemler boşsa
For OgeFrame = IlkBosSatir + 1 To SonDoluSatir

    If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
        Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
        Controls("Adet" & OgeFrame).Value <> "" Or _
        Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
        Controls("Aciklama" & OgeFrame).Value <> "" Or _
        Controls("Sonuc" & OgeFrame).Value <> "" Or _
        Controls("Rapor1No" & OgeFrame).Value <> "" Then

        Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
        Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
        Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
        Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
        Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value
        Controls("Sonuc" & OgeFrame - 1).Value = Controls("Sonuc" & OgeFrame).Value
        Controls("Rapor1No" & OgeFrame - 1).Value = Controls("Rapor1No" & OgeFrame).Value

        Controls("OgeTuru" & OgeFrame).Value = ""
        Controls("OgeDegeri" & OgeFrame).Value = ""
        Controls("Adet" & OgeFrame).Value = ""
        Controls("OgeIdNo" & OgeFrame).Value = ""
        Controls("Aciklama" & OgeFrame).Value = ""
        Controls("Sonuc" & OgeFrame).Value = ""
        Controls("Rapor1No" & OgeFrame).Value = ""

    End If
        
Next OgeFrame

NormalProsedureGit:

'_____________________Güncelleme 18112019 1238


If OgeTuruFrame19.Visible = True Then
    OgeTuru19.Value = ""
    OgeDegeri19.Value = ""
    Adet19.Value = ""
    OgeIdNo19.Value = ""
    Aciklama19.Value = ""
    Sonuc19.Value = ""
    Rapor1No19.Value = ""
    OgeTuruFrame19.Visible = False
    GoTo Son
ElseIf OgeTuruFrame18.Visible = True Then
    OgeTuru18.Value = ""
    OgeDegeri18.Value = ""
    Adet18.Value = ""
    OgeIdNo18.Value = ""
    Aciklama18.Value = ""
    Sonuc18.Value = ""
    Rapor1No18.Value = ""
    OgeTuruFrame18.Visible = False
    GoTo Son
ElseIf OgeTuruFrame17.Visible = True Then
    OgeTuru17.Value = ""
    OgeDegeri17.Value = ""
    Adet17.Value = ""
    OgeIdNo17.Value = ""
    Aciklama17.Value = ""
    Sonuc17.Value = ""
    Rapor1No17.Value = ""
    OgeTuruFrame17.Visible = False
    GoTo Son
ElseIf OgeTuruFrame16.Visible = True Then
    OgeTuru16.Value = ""
    OgeDegeri16.Value = ""
    Adet16.Value = ""
    OgeIdNo16.Value = ""
    Aciklama16.Value = ""
    Sonuc16.Value = ""
    Rapor1No16.Value = ""
    OgeTuruFrame16.Visible = False
    GoTo Son
ElseIf OgeTuruFrame15.Visible = True Then
    OgeTuru15.Value = ""
    OgeDegeri15.Value = ""
    Adet15.Value = ""
    OgeIdNo15.Value = ""
    Aciklama15.Value = ""
    Sonuc15.Value = ""
    Rapor1No15.Value = ""
    OgeTuruFrame15.Visible = False
    GoTo Son
ElseIf OgeTuruFrame14.Visible = True Then
    OgeTuru14.Value = ""
    OgeDegeri14.Value = ""
    Adet14.Value = ""
    OgeIdNo14.Value = ""
    Aciklama14.Value = ""
    Sonuc14.Value = ""
    Rapor1No14.Value = ""
    OgeTuruFrame14.Visible = False
    GoTo Son
ElseIf OgeTuruFrame13.Visible = True Then
    OgeTuru13.Value = ""
    OgeDegeri13.Value = ""
    Adet13.Value = ""
    OgeIdNo13.Value = ""
    Aciklama13.Value = ""
    Sonuc13.Value = ""
    Rapor1No13.Value = ""
    OgeTuruFrame13.Visible = False
    GoTo Son
ElseIf OgeTuruFrame12.Visible = True Then
    OgeTuru12.Value = ""
    OgeDegeri12.Value = ""
    Adet12.Value = ""
    OgeIdNo12.Value = ""
    Aciklama12.Value = ""
    Sonuc12.Value = ""
    Rapor1No12.Value = ""
    OgeTuruFrame12.Visible = False
    GoTo Son
ElseIf OgeTuruFrame11.Visible = True Then
    OgeTuru11.Value = ""
    OgeDegeri11.Value = ""
    Adet11.Value = ""
    OgeIdNo11.Value = ""
    Aciklama11.Value = ""
    Sonuc11.Value = ""
    Rapor1No11.Value = ""
    OgeTuruFrame11.Visible = False
    GoTo Son
ElseIf OgeTuruFrame10.Visible = True Then
    OgeTuru10.Value = ""
    OgeDegeri10.Value = ""
    Adet10.Value = ""
    OgeIdNo10.Value = ""
    Aciklama10.Value = ""
    Sonuc10.Value = ""
    Rapor1No10.Value = ""
    OgeTuruFrame10.Visible = False
    GoTo Son
ElseIf OgeTuruFrame9.Visible = True Then
    OgeTuru9.Value = ""
    OgeDegeri9.Value = ""
    Adet9.Value = ""
    OgeIdNo9.Value = ""
    Aciklama9.Value = ""
    Sonuc9.Value = ""
    Rapor1No9.Value = ""
    OgeTuruFrame9.Visible = False
    GoTo Son
ElseIf OgeTuruFrame8.Visible = True Then
    OgeTuru8.Value = ""
    OgeDegeri8.Value = ""
    Adet8.Value = ""
    OgeIdNo8.Value = ""
    Aciklama8.Value = ""
    Sonuc8.Value = ""
    Rapor1No8.Value = ""
    OgeTuruFrame8.Visible = False
    GoTo Son
ElseIf OgeTuruFrame7.Visible = True Then
    OgeTuru7.Value = ""
    OgeDegeri7.Value = ""
    Adet7.Value = ""
    OgeIdNo7.Value = ""
    Aciklama7.Value = ""
    Sonuc7.Value = ""
    Rapor1No7.Value = ""
    OgeTuruFrame7.Visible = False
    GoTo Son
ElseIf OgeTuruFrame6.Visible = True Then
    OgeTuru6.Value = ""
    OgeDegeri6.Value = ""
    Adet6.Value = ""
    OgeIdNo6.Value = ""
    Aciklama6.Value = ""
    Sonuc6.Value = ""
    Rapor1No6.Value = ""
    OgeTuruFrame6.Visible = False
    GoTo Son
ElseIf OgeTuruFrame5.Visible = True Then
    OgeTuru5.Value = ""
    OgeDegeri5.Value = ""
    Adet5.Value = ""
    OgeIdNo5.Value = ""
    Aciklama5.Value = ""
    Sonuc5.Value = ""
    Rapor1No5.Value = ""
    OgeTuruFrame5.Visible = False
    GoTo Son
ElseIf OgeTuruFrame4.Visible = True Then
    OgeTuru4.Value = ""
    OgeDegeri4.Value = ""
    Adet4.Value = ""
    OgeIdNo4.Value = ""
    Aciklama4.Value = ""
    Sonuc4.Value = ""
    Rapor1No4.Value = ""
    OgeTuruFrame4.Visible = False
    GoTo Son
ElseIf OgeTuruFrame3.Visible = True Then
    OgeTuru3.Value = ""
    OgeDegeri3.Value = ""
    Adet3.Value = ""
    OgeIdNo3.Value = ""
    Aciklama3.Value = ""
    Sonuc3.Value = ""
    Rapor1No3.Value = ""
    OgeTuruFrame3.Visible = False
    GoTo Son
ElseIf OgeTuruFrame2.Visible = True Then
    OgeTuru2.Value = ""
    OgeDegeri2.Value = ""
    Adet2.Value = ""
    OgeIdNo2.Value = ""
    Aciklama2.Value = ""
    Sonuc2.Value = ""
    Rapor1No2.Value = ""
    OgeTuruFrame2.Visible = False
    GoTo Son
ElseIf OgeTuruFrame1.Visible = True Then
    OgeTuru1.Value = ""
    OgeDegeri1.Value = ""
    Adet1.Value = ""
    OgeIdNo1.Value = ""
    Aciklama1.Value = ""
    Sonuc1.Value = ""
    Rapor1No1.Value = ""
    OgeTuruFrame1.Visible = False
    GoTo Son
End If

Son:

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

If ScrollTakip > 0 Then
    If ScrollTakip = 120 Then
        ScrollTakip = ScrollTakip - 24 - 6
    ElseIf ScrollTakip > 120 Then
        ScrollTakip = ScrollTakip - 24
    Else
        ScrollTakip = ScrollTakip - 24
    End If
End If

If ScrollTakip > 120 + 6 Then
    Call SetScrollHook(Me.ScrollFrame, ScrollTakip, 24)
    ScrollFrame.ScrollTop = 0
    ScrollFrame.ScrollTop = ScrollFrame.ScrollHeight
ElseIf ScrollTakip > 0 And ScrollTakip <= 120 + 6 Then
    ScrollFrame.ScrollTop = 0
    RemoveScrollHook
    ScrollFrame.ScrollBars = fmScrollBarsNone
End If


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant

RemoveScrollHook 'Userform Frame


yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 100
Call timeout(0.01)
    If Rep > 50 Then
        core_report1_entry_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_report1_entry_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_report1_entry_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_report1_entry_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_report1_entry_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_report1_entry_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 60

Unload Me

End Sub

Private Sub UserForm_Initialize()
Dim ctl As MSForms.Control
Dim lCount As Long
Dim InputLblEvt As clLabelClassCalendar
Dim ClrLab As MSForms.Control
Dim i As Long, SiraSay As Long, Say As Long

ScrollTakip = 0
Threshold = 126

'Muhatap Temasını uyarla.
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"

ThisWorkbook.Worksheets(2).Range("CW6").Value = "Incoming Contact Theme"
ThisWorkbook.Worksheets(2).Range("CW7").Value = "Outgoing Contact Theme"

'Geçici fark kayıtlarını sil
ThisWorkbook.Worksheets(7).Rows("3:30").EntireRow.Delete

ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(7).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"


Call ComboGetirReset

'Nesne renkleri
For Each ClrLab In core_report1_entry_UI.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "CheckBox" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "OptionButton" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "TextBox" Then
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ClrLab) = "ComboBox" Then
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    
    'YENİ
    If TypeName(ClrLab) = "Frame" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
        ClrLab.BorderColor = RGB(180, 180, 180)
    End If
Next ClrLab

Call UstYaziGirisi_Click

core_report1_entry_UI.UstMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
core_report1_entry_UI.AltMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
ComboGetir.BackColor = RGB(225, 235, 245)


MaxiMini.BackColor = RGB(225, 235, 245)
MaxiMini.ForeColor = RGB(30, 30, 30)

Kaydet.BackColor = RGB(225, 235, 245)
Kaydet.ForeColor = RGB(30, 30, 30)

core_report1_entry_UI.BackColor = RGB(230, 230, 230) 'YENİ

FarkGirisi.Visible = False

Exit Sub
ErrorHandle:
MsgBox Err.Description

End Sub


Private Sub MaxiMini_Click()
'ThisWorkbook.Worksheets(3).Visible = True
'ThisWorkbook.Worksheets(3).Activate
If MaxiMini.Caption = "ÇÊ" Then
    MaxiMini.Caption = "ÉÈ"
    ThisWorkbook.Activate
    'ThisWorkbook.Worksheets(3).Range("E6").Select
    Call FormPositionMini
Else
    MaxiMini.Caption = "ÉÈ"
    MaxiMini.Caption = "ÇÊ"
    ThisWorkbook.Activate
    'ThisWorkbook.Worksheets(3).Range("E6").Select
    Call FormPositionMaxi
End If
End Sub

Sub FormPositionMaxi()
Dim AppXCenter, AppYCenter As Long
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant
Dim ClrLab As MSForms.Control

For Each ClrLab In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.Visible = True
        MaxiMini.Visible = True
    End If
    If TypeName(ClrLab) = "ComboBox" Then
        ClrLab.Visible = True
    End If
    If TypeName(ClrLab) = "OptionButton" Then
        ClrLab.Visible = True
    End If
    If TypeName(ClrLab) = "CommandButton" Then
        ClrLab.Visible = True
    End If
    If TypeName(ClrLab) = "TextBox" Then
        ClrLab.Visible = True
    End If
Next ClrLab


'Sağa doğru genişlet
BaslikFrame.Visible = True
UstMenuFrame.Visible = True
MaxiMini.Left = 926
MaxiMini.Top = 18
MaxiMini.Width = 50
MaxiMini.Height = 18
'Ekrana göre formun ayarlanması
If EkranKontrol = True Then
    TasiyiciFrame.Left = 12
    TasiyiciFrame.Top = 12
Else
    TasiyiciFrame.Left = 36
    TasiyiciFrame.Top = 12
End If


'Ekrana göre formun ayarlanması
If EkranKontrol = True Then

    AppXCenter = Application.Left + (Application.Width / 2)
    AppYCenter = Application.Top + (Application.Height / 2)
    
    'Formu önce ekrana ortala
    With core_report1_entry_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * (1024 + 12))
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 485)
    End With
        
    'Formun görünümü
    Do
    DoEvents
    genislik = genislik + 140
        Me.Width = genislik
        Call timeout(0.01)
        If genislik > 1024 + 12 Then
            genislik = 1024 + 12
            Me.Width = genislik
        End If
    Loop Until genislik = 1024 + 12
    
    'Formun görünümü (DİKEY FARKLILAŞMA)
    If Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And UstYaziFrame.Visible = True Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        Rep = 485 '556 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        'Rep = core_report1_entry_UI.Height
        
        core_report1_entry_UI.Width = 1024 + 12

        core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18 - 30
        core_report1_entry_UI.ScrollTop = 0
    ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And UstYaziFrame.Visible = False Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        Rep = 485 '556 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        'Rep = core_report1_entry_UI.Height
        
        core_report1_entry_UI.Width = 1024 + 12
        
        core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + Tutanak2Frame.Height + 12 - 30
        core_report1_entry_UI.ScrollTop = 0
    ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = False And UstYaziFrame.Visible = False Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + 6
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + 6
        Rep = 485 '556 + Rapor1Frame.Height + 6
        'Rep = core_report1_entry_UI.Height
        
        core_report1_entry_UI.Width = 1024 + 12

        core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report1_entry_UI.ScrollHeight = 546 + Rapor1Frame.Height + 6 - 30
        core_report1_entry_UI.ScrollTop = 0
    ElseIf Rapor1Frame.Visible = False And Tutanak2Frame.Visible = False And UstYaziFrame.Visible = False Then
        'Formun görünümü
        AltMenuFrame.Top = 462 '444 '299
        'TasiyiciFrame.Height = 486
        Rep = 546 '556 '497 '352
        
        core_report1_entry_UI.Width = 1024
        
        core_report1_entry_UI.ScrollTop = 0
        core_report1_entry_UI.ScrollHeight = 0
        core_report1_entry_UI.ScrollBars = fmScrollBarsNone
    End If
    
    'Aşağı doğru genişlet
    yukseklik = 70
    Do
    DoEvents
    yukseklik = yukseklik + 50
        Me.Height = yukseklik
        Call timeout(0.01)
        If yukseklik > Rep Then
            yukseklik = Rep
            Me.Height = yukseklik
        End If
    Loop Until yukseklik = Rep

Else

    AppXCenter = Application.Left + (Application.Width / 2)
    AppYCenter = Application.Top + (Application.Height / 2)
    
    'Formu önce ekrana ortala
    With core_report1_entry_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1072)
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 560)
    End With

    Do
    DoEvents
    genislik = genislik + 140
        Me.Width = genislik
        Call timeout(0.01)
        If genislik > 1072 Then
            genislik = 1072
            Me.Width = genislik
        End If
    Loop Until genislik = 1072
    
    'Formun görünümü (DİKEY FARKLILAŞMA)
    If Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And UstYaziFrame.Visible = True Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        Rep = 556 + Rapor1Frame.Height + Tutanak2Frame.Height + UstYaziFrame.Height + 18
        'Rep = core_report1_entry_UI.Height
    ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And UstYaziFrame.Visible = False Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        Rep = 556 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
        'Rep = core_report1_entry_UI.Height
    ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = False And UstYaziFrame.Visible = False Then
        AltMenuFrame.Top = 462 + Rapor1Frame.Height + 6
        'TasiyiciFrame.Height = 486 + Rapor1Frame.Height + 6
        Rep = 556 + Rapor1Frame.Height + 6
        'Rep = core_report1_entry_UI.Height
    ElseIf Rapor1Frame.Visible = False And Tutanak2Frame.Visible = False And UstYaziFrame.Visible = False Then
        AltMenuFrame.Top = 462
        'TasiyiciFrame.Height = 486
        Rep = 556
        'Rep = core_report1_entry_UI.Height
    End If
    
    'Aşağı doğru genişlet
    yukseklik = 70
    Do
    DoEvents
    yukseklik = yukseklik + 50
        Me.Height = yukseklik
        Call timeout(0.01)
        If yukseklik > Rep Then
            yukseklik = Rep
            Me.Height = yukseklik
        End If
    Loop Until yukseklik = Rep

End If

'Modeless modunda userformun mouseover seçeneği yavaşlıyor. Sorun bu şekilde çözüldü.
core_report1_entry_UI.Hide
core_report1_entry_UI.Show vbModal

End Sub

Sub FormPositionMini()
Dim AppXCenter, AppYCenter As Long
Dim yukseklik As Variant, genislik As Variant
Dim ClrLab As MSForms.Control


For Each ClrLab In core_report1_entry_UI.UstMenuFrame.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.Visible = False
        MaxiMini.Visible = True
    End If
    If TypeName(ClrLab) = "ComboBox" Then
        ClrLab.Visible = False
    End If
    If TypeName(ClrLab) = "OptionButton" Then
        ClrLab.Visible = False
    End If
    If TypeName(ClrLab) = "CommandButton" Then
        ClrLab.Visible = False
    End If
    If TypeName(ClrLab) = "TextBox" Then
        ClrLab.Visible = False
    End If
Next ClrLab

AppXCenter = Application.Left + (Application.Width / 2)
AppYCenter = Application.Top + (Application.Height / 2)

'Sağ üst köşeye çek
With core_report1_entry_UI
    .StartUpPosition = 0
    .Left = Application.Left '+ (0.5 * Application.Width) - (0.5 * 1034)
    .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 560)
End With

'If EkranKontrol = True Then
If core_report1_entry_UI.ScrollHeight > 0 Then
    core_report1_entry_UI.ScrollTop = 0
    core_report1_entry_UI.ScrollHeight = 0
    core_report1_entry_UI.ScrollBars = fmScrollBarsNone
End If

'Yukarı doğru daralt
yukseklik = Me.Height
Do
DoEvents
yukseklik = yukseklik - 100
    Me.Height = yukseklik
    Call timeout(0.01)
    If yukseklik < 52 Then
        yukseklik = 52
        Me.Height = yukseklik
    End If
Loop Until yukseklik = 52


'__________Formu sağa taşı ve genişliğini daralt.

TasiyiciFrame.Left = 0
TasiyiciFrame.Top = 0
genislik = 0
Do
DoEvents
genislik = genislik + 2
    Call timeout(0.01)
    core_report1_entry_UI.Left = core_report1_entry_UI.Left + Application.Left + (0.2 * Application.Width)
    If core_report1_entry_UI.Left > Application.Left + (0.9 * Application.Width) Then
        core_report1_entry_UI.Left = Application.Left + (0.9 * Application.Width)
    End If
    If genislik > 10 Then
        genislik = 10
    End If
Loop Until genislik = 10

Me.Width = 50
MaxiMini.Width = 85
MaxiMini.Left = 0
MaxiMini.Top = 0

BaslikFrame.Visible = False
UstMenuFrame.Visible = False


'Modeless modunda userformun mouseover seçeneği yavaşlıyor. Sorun bu şekilde çözüldü.
core_report1_entry_UI.Hide
core_report1_entry_UI.Show vbModeless

End Sub

Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub





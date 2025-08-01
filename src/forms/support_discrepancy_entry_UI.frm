VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_discrepancy_entry_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   7410
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   15195
   OleObjectBlob   =   "support_discrepancy_entry_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_discrepancy_entry_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Abort As Boolean
Dim Threshold As Long

Sub ColorChangerGenel()

'Kapat
If Kapat.BackColor <> RGB(225, 235, 245) Then
    Kapat.BackColor = RGB(225, 235, 245)
    Kapat.ForeColor = RGB(30, 30, 30)
End If
'Kaydet
If Kaydet.BackColor <> RGB(225, 235, 245) Then
    Kaydet.BackColor = RGB(225, 235, 245)
    Kaydet.ForeColor = RGB(30, 30, 30)
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

End Sub

Private Sub Kaydet_Click()
Dim OgeFrame As Integer, Kont As Integer, TumKontx As Integer, i As Integer
Dim ctl As MSForms.Control, Bilgi As Variant
Dim OgeTuruKont As Integer, OgeDegeriKont As Integer, AdetKont As Integer
Dim OgeIdNoKont As Integer, AciklamaKont As Integer, Maxi As Integer
Dim OgeTuruKontSatir As Integer, OgeDegeriKontSatir As Integer, AdetKontSatir As Integer
Dim OgeIdNoKontSatir As Integer, AciklamaKontSatir As Integer
Dim WsFarkGiris As Worksheet, SayA As Integer, SayD As Integer, SayG As Integer, SayJ As Integer, SayM As Integer, SayFarkGiris As Integer


'Fark Girişleri sayfasında bulunan verileri temizle.
Set WsFarkGiris = ThisWorkbook.Worksheets(7)
'Potect/Unprotect kodları gelecek
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"

'Maksimum değerler.
SayA = WsFarkGiris.Range("A100000").End(xlUp).Row
SayD = WsFarkGiris.Range("D100000").End(xlUp).Row
SayG = WsFarkGiris.Range("G100000").End(xlUp).Row
SayJ = WsFarkGiris.Range("J100000").End(xlUp).Row
SayM = WsFarkGiris.Range("M100000").End(xlUp).Row
SayFarkGiris = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)

If SayFarkGiris >= 3 Then
    WsFarkGiris.Rows("3:" & SayFarkGiris).EntireRow.Delete
End If

'______________________

'Tüm bölümler için ön kontrol
TumKontx = 0
For Each ctl In support_discrepancy_entry_UI.ScrollFrame.Controls 'ScrollFrame
    If TypeName(ctl) = "ComboBox" Then
        If ctl.Value <> "" Then
            TumKontx = 1
        End If
    End If
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value <> "" Then
            TumKontx = 1
        End If
    End If
Next ctl
If TumKontx = 0 Then
    GoTo Son
End If
'______________________


'Arada boş bırakılan satırların kontrolü; öğe türü, öğe değeri, adet, öğe ID no (ve açıklama)
If OgeTuru.Value = "" Then
    Bilgi = MsgBox("Item type is not specified. Please check the rows and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
End If

If OgeDegeri.Value = "" Then
    Bilgi = MsgBox("Item value is not specified. Please check the relevant fields and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
End If

If Adet.Value = "" Then
    Bilgi = MsgBox("Quantity is not specified. Please check the rows and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
End If

If OgeIdNo.Value = "" Then
    Bilgi = MsgBox("Item ID number is not specified. Please check the rows and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
End If

If Aciklama.Value = "" Then
    ' (İlgili eylemi buraya ekleyebilirsiniz.)
End If


'Aradaki satırları kontrol et.
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
'        If Controls("Aciklama" & i).Value = "" Then
'            AciklamaKontSatir = i
'        End If
    Next i
End If
'Yukarıdaki maxi değeri, (aşağıda bulunan kodlarda) verilerin rapor1 formundan
'sayfaya aktarılmasında kullanılıyor.
If OgeTuruKontSatir <> 0 And OgeDegeriKontSatir <> 0 And AdetKontSatir <> 0 And OgeIdNoKontSatir Then
    Bilgi = MsgBox("It has been detected that a row was skipped. Please check the rows and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
End If

YinedeKaydet19:
If OgeTuruKontSatir <> 0 Then
    Bilgi = MsgBox("Item type is missing. Please check the relevant fields and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
ElseIf OgeDegeriKontSatir <> 0 Then
    Bilgi = MsgBox("Item value is missing. Please check the relevant fields and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
ElseIf AdetKontSatir <> 0 Then
    Bilgi = MsgBox("Quantity is missing. Please check the relevant fields and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
ElseIf OgeIdNoKontSatir <> 0 Then
    Bilgi = MsgBox("Item ID number is missing. Please check the relevant fields and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System")
    GoTo Son
ElseIf AciklamaKontSatir <> 0 Then
    ' (You can insert the relevant message or action here.)
End If


'Verileri geçici olarak kaydet
WsFarkGiris.Cells(3, 1).Value = OgeTuru.Value
WsFarkGiris.Cells(3, 4).Value = OgeDegeri.Value
WsFarkGiris.Cells(3, 7).Value = Adet.Value
WsFarkGiris.Cells(3, 10).Value = OgeIdNo.Value
WsFarkGiris.Cells(3, 13).Value = Aciklama.Value
If Maxi > 0 Then
    For OgeFrame = 1 To Maxi
        WsFarkGiris.Cells(3 + OgeFrame, 1).Value = Controls("OgeTuru" & OgeFrame).Value
        WsFarkGiris.Cells(3 + OgeFrame, 4).Value = Controls("OgeDegeri" & OgeFrame).Value
        WsFarkGiris.Cells(3 + OgeFrame, 7).Value = Controls("Adet" & OgeFrame).Value
        WsFarkGiris.Cells(3 + OgeFrame, 10).Value = Controls("OgeIdNo" & OgeFrame).Value
        WsFarkGiris.Cells(3 + OgeFrame, 13).Value = Controls("Aciklama" & OgeFrame).Value
    Next OgeFrame
End If


ThisWorkbook.Protect "123"
ThisWorkbook.Worksheets(7).Protect Password:="123"
Unload Me
GoTo Out


Son:

ThisWorkbook.Protect "123"
ThisWorkbook.Worksheets(7).Protect Password:="123"

Out:

End Sub

Private Sub Kapat_Click()
    Unload Me
End Sub

Private Sub OgeEkleKaldirLabel_Click()
support_item_types_UI.Show
'support_item_types_UI.Show vbModeless
End Sub

Private Sub OgeDegeriEkleKaldirLabel_Click()
support_item_values_UI.Show
'support_item_values_UI.Show vbModeless
End Sub


Private Sub LblOgeTuruUst_Click()
MsgBox "Select the item type related to the inspected item from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list below, you can also press the first letter of your desired selection on the keyboard until it appears." & vbNewLine & vbNewLine & _
"To enter multiple item types/item values/quantities, click the + sign at the far right of this row. To remove item type/item value/quantity rows, click the - sign at the same location." & vbNewLine & vbNewLine & _
"If the relevant item type is not listed in the dropdown, click the ± sign to the left of the Item Type label and follow the instructions in the opened window to define the item type in the system." & vbNewLine & vbNewLine & _
"The selection made in the Item Type field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeDegeriUst_Click()
MsgBox "Select the item value related to the inspected item from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list below, you can also press the first digit of your desired selection on the keyboard until it appears." & vbNewLine & vbNewLine & _
"To enter multiple item types/item values/quantities, click the + sign at the far right of this row. To remove item type/item value/quantity rows, click the - sign at the same location." & vbNewLine & vbNewLine & _
"If the relevant item value is not listed in the dropdown, click the ± sign to the left of the Item Value label and follow the instructions in the opened window to define the item value in the system." & vbNewLine & vbNewLine & _
"The selection made in the Item Value field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAdetUst_Click()
MsgBox "Enter the quantity related to the inspected item in the box below." & vbNewLine & vbNewLine & _
"To enter multiple item types/item values/quantities, click the + sign at the far right of this row. To remove item type/item value/quantity rows, click the - sign at the same location." & vbNewLine & vbNewLine & _
"The selection made in the Quantity field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeIdNoUst_Click()
MsgBox "Enter the item ID number related to the inspected item in the box below." & vbNewLine & vbNewLine & _
"If you have selected 'Yes' for the Dispatch List option above, when an item type is selected (or changed), the term 'Dispatch List' will automatically appear in the item ID field. You may replace it with the actual item ID number if you prefer." & vbNewLine & vbNewLine & _
"To enter multiple item types/item values/quantities, click the + sign at the far right of this row. To remove item type/item value/quantity rows, click the - sign at the same location." & vbNewLine & vbNewLine & _
"The selection made in the Item ID No field is used in Statement 1, Report 1, and Statement 2. For more details, click the Help button in the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAciklamaUst_Click()
MsgBox "You may add a description related to the inspected item in the box below." & vbNewLine & vbNewLine & _
"To enter multiple item types/item values/quantities, click the + sign at the far right of this row. To remove item type/item value/quantity rows, click the - sign at the same location." & vbNewLine & vbNewLine & _
"The selection made in the Description field is used in Statement 1 and Statement 2. For more details, click the Help button in the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
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

If ScrollTakip2 < 456 Then
    If ScrollTakip2 = 120 Then
        ScrollTakip2 = ScrollTakip2 + 24 + 6
    ElseIf ScrollTakip2 > 120 Then
        ScrollTakip2 = ScrollTakip2 + 24
    Else
        ScrollTakip2 = ScrollTakip2 + 24
    End If
End If
If ScrollTakip2 > 120 + 6 Then
    Call SetScrollHook(Me.ScrollFrame, ScrollTakip2, 24)
    ScrollFrame.ScrollTop = ScrollTakip2
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
        Controls("Aciklama" & OgeFrame).Value <> "" Then

        SonDoluSatir = OgeFrame
        GoTo SonDoluSatirBulundu

    End If
Next OgeFrame
SonDoluSatirBulundu:

If SonDoluSatir = 0 Then
 GoTo NormalProsedureGit
End If

'Başlangıç satırı boşsa
If OgeTuru.Value = "" And OgeDegeri.Value = "" And Adet.Value = "" And OgeIdNo.Value = "" And Aciklama.Value = "" Then

    For OgeFrame = 1 To SonDoluSatir

        If OgeFrame = 1 Then
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Then

                OgeTuru.Value = Controls("OgeTuru" & OgeFrame).Value
                OgeDegeri.Value = Controls("OgeDegeri" & OgeFrame).Value
                Adet.Value = Controls("Adet" & OgeFrame).Value
                OgeIdNo.Value = Controls("OgeIdNo" & OgeFrame).Value
                Aciklama.Value = Controls("Aciklama" & OgeFrame).Value

                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""

            End If
        Else
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Then

                Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
                Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
                Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
                Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
                Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value

                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""

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
        Controls("Aciklama" & OgeFrame).Value = "" Then

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
        Controls("Aciklama" & OgeFrame).Value <> "" Then

        Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
        Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
        Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
        Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
        Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value

        Controls("OgeTuru" & OgeFrame).Value = ""
        Controls("OgeDegeri" & OgeFrame).Value = ""
        Controls("Adet" & OgeFrame).Value = ""
        Controls("OgeIdNo" & OgeFrame).Value = ""
        Controls("Aciklama" & OgeFrame).Value = ""

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
    OgeTuruFrame19.Visible = False
    GoTo Son
ElseIf OgeTuruFrame18.Visible = True Then
    OgeTuru18.Value = ""
    OgeDegeri18.Value = ""
    Adet18.Value = ""
    OgeIdNo18.Value = ""
    Aciklama18.Value = ""
    OgeTuruFrame18.Visible = False
    GoTo Son
ElseIf OgeTuruFrame17.Visible = True Then
    OgeTuru17.Value = ""
    OgeDegeri17.Value = ""
    Adet17.Value = ""
    OgeIdNo17.Value = ""
    Aciklama17.Value = ""
    OgeTuruFrame17.Visible = False
    GoTo Son
ElseIf OgeTuruFrame16.Visible = True Then
    OgeTuru16.Value = ""
    OgeDegeri16.Value = ""
    Adet16.Value = ""
    OgeIdNo16.Value = ""
    Aciklama16.Value = ""
    OgeTuruFrame16.Visible = False
    GoTo Son
ElseIf OgeTuruFrame15.Visible = True Then
    OgeTuru15.Value = ""
    OgeDegeri15.Value = ""
    Adet15.Value = ""
    OgeIdNo15.Value = ""
    Aciklama15.Value = ""
    OgeTuruFrame15.Visible = False
    GoTo Son
ElseIf OgeTuruFrame14.Visible = True Then
    OgeTuru14.Value = ""
    OgeDegeri14.Value = ""
    Adet14.Value = ""
    OgeIdNo14.Value = ""
    Aciklama14.Value = ""
    OgeTuruFrame14.Visible = False
    GoTo Son
ElseIf OgeTuruFrame13.Visible = True Then
    OgeTuru13.Value = ""
    OgeDegeri13.Value = ""
    Adet13.Value = ""
    OgeIdNo13.Value = ""
    Aciklama13.Value = ""
    OgeTuruFrame13.Visible = False
    GoTo Son
ElseIf OgeTuruFrame12.Visible = True Then
    OgeTuru12.Value = ""
    OgeDegeri12.Value = ""
    Adet12.Value = ""
    OgeIdNo12.Value = ""
    Aciklama12.Value = ""
    OgeTuruFrame12.Visible = False
    GoTo Son
ElseIf OgeTuruFrame11.Visible = True Then
    OgeTuru11.Value = ""
    OgeDegeri11.Value = ""
    Adet11.Value = ""
    OgeIdNo11.Value = ""
    Aciklama11.Value = ""
    OgeTuruFrame11.Visible = False
    GoTo Son
ElseIf OgeTuruFrame10.Visible = True Then
    OgeTuru10.Value = ""
    OgeDegeri10.Value = ""
    Adet10.Value = ""
    OgeIdNo10.Value = ""
    Aciklama10.Value = ""
    OgeTuruFrame10.Visible = False
    GoTo Son
ElseIf OgeTuruFrame9.Visible = True Then
    OgeTuru9.Value = ""
    OgeDegeri9.Value = ""
    Adet9.Value = ""
    OgeIdNo9.Value = ""
    Aciklama9.Value = ""
    OgeTuruFrame9.Visible = False
    GoTo Son
ElseIf OgeTuruFrame8.Visible = True Then
    OgeTuru8.Value = ""
    OgeDegeri8.Value = ""
    Adet8.Value = ""
    OgeIdNo8.Value = ""
    Aciklama8.Value = ""
    OgeTuruFrame8.Visible = False
    GoTo Son
ElseIf OgeTuruFrame7.Visible = True Then
    OgeTuru7.Value = ""
    OgeDegeri7.Value = ""
    Adet7.Value = ""
    OgeIdNo7.Value = ""
    Aciklama7.Value = ""
    OgeTuruFrame7.Visible = False
    GoTo Son
ElseIf OgeTuruFrame6.Visible = True Then
    OgeTuru6.Value = ""
    OgeDegeri6.Value = ""
    Adet6.Value = ""
    OgeIdNo6.Value = ""
    Aciklama6.Value = ""
    OgeTuruFrame6.Visible = False
    GoTo Son
ElseIf OgeTuruFrame5.Visible = True Then
    OgeTuru5.Value = ""
    OgeDegeri5.Value = ""
    Adet5.Value = ""
    OgeIdNo5.Value = ""
    Aciklama5.Value = ""
    OgeTuruFrame5.Visible = False
    GoTo Son
ElseIf OgeTuruFrame4.Visible = True Then
    OgeTuru4.Value = ""
    OgeDegeri4.Value = ""
    Adet4.Value = ""
    OgeIdNo4.Value = ""
    Aciklama4.Value = ""
    OgeTuruFrame4.Visible = False
    GoTo Son
ElseIf OgeTuruFrame3.Visible = True Then
    OgeTuru3.Value = ""
    OgeDegeri3.Value = ""
    Adet3.Value = ""
    OgeIdNo3.Value = ""
    Aciklama3.Value = ""
    OgeTuruFrame3.Visible = False
    GoTo Son
ElseIf OgeTuruFrame2.Visible = True Then
    OgeTuru2.Value = ""
    OgeDegeri2.Value = ""
    Adet2.Value = ""
    OgeIdNo2.Value = ""
    Aciklama2.Value = ""
    OgeTuruFrame2.Visible = False
    GoTo Son
ElseIf OgeTuruFrame1.Visible = True Then
    OgeTuru1.Value = ""
    OgeDegeri1.Value = ""
    Adet1.Value = ""
    OgeIdNo1.Value = ""
    Aciklama1.Value = ""
    OgeTuruFrame1.Visible = False
    GoTo Son
End If

Son:

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

If ScrollTakip2 > 0 Then
    If ScrollTakip2 = 120 Then
        ScrollTakip2 = ScrollTakip2 - 24 - 6
    ElseIf ScrollTakip2 > 120 Then
        ScrollTakip2 = ScrollTakip2 - 24
    Else
        ScrollTakip2 = ScrollTakip2 - 24
    End If
End If

If ScrollTakip2 > 120 + 6 Then
    Call SetScrollHook(Me.ScrollFrame, ScrollTakip2, 24)
    ScrollFrame.ScrollTop = 0
    ScrollFrame.ScrollTop = ScrollFrame.ScrollHeight
ElseIf ScrollTakip2 > 0 And ScrollTakip2 <= 120 + 6 Then
    ScrollFrame.ScrollTop = 0
    RemoveScrollHook
    ScrollFrame.ScrollBars = fmScrollBarsNone
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


Private Sub Kaydet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kaydet.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kaydet.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
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

Private Sub LblOgeTuruUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblOgeDegeriUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
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


Private Sub BaslikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub OgeTurleriFrameUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub ScrollFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call Move_SetScrollHook(Me.ScrollFrame, Threshold, ScrollTakip2)
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub TasiyiciFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub

Sub GeciciKayitlariCagir()
Dim OgeFrame As Integer, Fark As Long
Dim WsFarkGiris As Worksheet, SayA As Integer, SayD As Integer, SayG As Integer, SayJ As Integer, SayM As Integer, SayFarkGiris As Integer

SayFarkGiris = 0
'Fark Girişleri sayfasında bulunan verileri temizle.
Set WsFarkGiris = ThisWorkbook.Worksheets(7)
'Maksimum değerler.
SayA = WsFarkGiris.Range("A100000").End(xlUp).Row
SayD = WsFarkGiris.Range("D100000").End(xlUp).Row
SayG = WsFarkGiris.Range("G100000").End(xlUp).Row
SayJ = WsFarkGiris.Range("J100000").End(xlUp).Row
SayM = WsFarkGiris.Range("M100000").End(xlUp).Row
SayFarkGiris = WorksheetFunction.Max(SayA, SayD, SayG, SayJ, SayM)

If SayFarkGiris < 3 Then
    GoTo Son
End If

'İlk satırı aktar
OgeTuru.Value = WsFarkGiris.Cells(3, 1).Value
OgeDegeri.Value = WsFarkGiris.Cells(3, 4).Value
Adet.Value = WsFarkGiris.Cells(3, 7).Value
OgeIdNo.Value = WsFarkGiris.Cells(3, 10).Value
Aciklama.Value = WsFarkGiris.Cells(3, 13).Value

'Sonraki satırlarda varsa aktar
Fark = SayFarkGiris - 3 + 1
If Fark > 1 And Fark < 21 Then
    For OgeFrame = 1 To Fark - 1
        'Controls("OgeTuruFrame" & OgeFrame).Visible = True
        Call EkleOge_Click
    Next OgeFrame
    For OgeFrame = 1 To Fark - 1
        Controls("OgeTuru" & OgeFrame).Value = WsFarkGiris.Cells(3 + OgeFrame, 1).Value
        Controls("OgeDegeri" & OgeFrame).Value = WsFarkGiris.Cells(3 + OgeFrame, 4).Value
        Controls("Adet" & OgeFrame).Value = WsFarkGiris.Cells(3 + OgeFrame, 7).Value
        Controls("OgeIdNo" & OgeFrame).Value = WsFarkGiris.Cells(3 + OgeFrame, 10).Value
        Controls("Aciklama" & OgeFrame).Value = WsFarkGiris.Cells(3 + OgeFrame, 13).Value
    Next OgeFrame
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

Son:

End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

'ThisWorkbook.Activate

ScrollTakip2 = 0
Threshold = 126

For Each ClrLab In support_discrepancy_entry_UI.Controls
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

UstMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
AltMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
TasiyiciFrame.Height = 328

Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
Kaydet.BackColor = RGB(225, 235, 245)
Kaydet.ForeColor = RGB(30, 30, 30)

support_discrepancy_entry_UI.BackColor = RGB(230, 230, 230) 'YENİ

'Geçici kayıtları, frak arayüzüne yükle
Call GeciciKayitlariCagir


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant

RemoveScrollHook 'Userform Frame

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 60
Call timeout(0.01)
    If Rep > 60 Then
        support_discrepancy_entry_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_discrepancy_entry_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_discrepancy_entry_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_discrepancy_entry_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_discrepancy_entry_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_discrepancy_entry_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 50

Unload Me

End Sub

Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_report3_2_entry_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   14940
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   21195
   OleObjectBlob   =   "core_report3_2_entry_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_report3_2_entry_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Abort As Boolean, KilitIptal As Boolean
Public OpenWordTakip As Boolean
'Dim dHeight As Double
Dim TutKont As Integer, Rapor1Kont As Integer, Tutanak2Kont As Integer, FinansalBirimUstYaziKont As Integer, UstYaziKont As Integer, MaxiR As Integer, Maxi As Integer
Dim TumKont As Integer

Dim StrRaporUnvan1 As String, StrRaporSicil1 As String, StrRaporUnvan2 As String, StrRaporSicil2 As String
Dim StrRaporUnvan3 As String, StrRaporSicil3 As String
Dim Threshold As Long

Private Sub LblIl_Click()
MsgBox "Please select the province where the transaction took place from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired province repeatedly to navigate until your selection appears. For example, after clicking once, pressing 'A' once will select Ankara; pressing it a second time will select Adana." & vbNewLine & vbNewLine & _
"To update information or add a new province or district, please click the ± sign on the side and follow the instructions in the window that opens to save changes to the system." & vbNewLine & vbNewLine & _
"The selection in the Province field is used in the automatic generation of the THEME code, Report 3 statement, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlce_Click()
MsgBox "Please select the district where the transaction took place from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired district repeatedly to navigate until your selection appears. For example, for Altındağ district of Ankara, after clicking once, pressing 'A' once will select Akyurt; pressing it a second time will select Altındağ." & vbNewLine & vbNewLine & _
"To update or add a new district, please click the ± sign next to the Province label and follow the instructions in the window that opens to save changes to the system." & vbNewLine & vbNewLine & _
"The selection in the District field is used in Report 3 statement and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanakTarihi_Click()
MsgBox "Please click the calendar icon on the side and select the statement date from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Statement Date field is used in the automatic generation of the THEME code and Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblKayitNo_Click()
MsgBox "Please enter the last five digits of the code number obtained from the THEME screen into the box on the side." & vbNewLine & vbNewLine & _
"For example, if the last five digits of the code number from THEME screen are 00012, you can also enter '12' directly; the system will automatically convert it to 00012 and generate the full 12-digit THEME code." & vbNewLine & vbNewLine & _
"Alternatively, you may skip this step and manually enter the 12-digit THEME number by selecting Manual in the THEME No field on the side." & vbNewLine & vbNewLine & _
"The selection in the Record No field is used in the automatic generation of the THEME code. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTemaTipi_Click()
MsgBox "Please select the organization issuing the THEME code (choose Organization A) from the dropdown list on the side." & vbNewLine & vbNewLine & _
"The selection in the THEME Type field is used in the automatic generation of the THEME code. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTemaNo_Click()
MsgBox "Provided entries are made in Province, Statement Date, Record No, and THEME Type fields, if you select Automatic on the side, the THEME code will be generated automatically. When Automatic is selected, user edits to the THEME code are disabled; when Manual is selected, user edits are allowed." & vbNewLine & vbNewLine & _
"When Automatic is checked in the THEME No field, to reflect changes in Province, Statement Date, Record No, and THEME Type fields in the THEME code, after making changes, first select Manual, then reselect Automatic." & vbNewLine & vbNewLine & _
"The THEME No field data (i.e., THEME code) is used in Report 3 statement, Statement 2, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirim_Click()
MsgBox "Please select the financial unit delivering the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of your desired selection repeatedly to navigate the list." & vbNewLine & vbNewLine & _
"If the financial unit name does not appear, click the ± sign on the side and follow instructions in the opened window to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Financial Unit field is used in Report 3 statement, Financial Unit, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTeslimatBirim_Click()
MsgBox "Please select the unit delivering the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of your desired selection repeatedly to navigate the list." & vbNewLine & vbNewLine & _
"If the delivery unit name does not appear, click the ± sign on the side and follow instructions in the opened window to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Delivery Unit field is used in Report 3 statement, Financial Unit, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblBildirilenBirim_Click()
MsgBox "Please select the invalid item unit from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of your desired selection repeatedly to navigate the list." & vbNewLine & vbNewLine & _
"If the invalid item unit name does not appear, click the ± sign on the side and follow instructions in the opened window to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Invalid Item Unit field is used in Report 3 statement, Financial Unit, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTeslimTarihi_Click()
MsgBox "Please click the calendar icon on the side and select the delivery date of the invalid item from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Delivery Date field is used in Report 3 statement, Financial Unit, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAdSoyad_Click()
MsgBox "Please enter the first and last name of the financial unit officer delivering the invalid item into the box on the side." & vbNewLine & vbNewLine & _
"The data entered in the Name and Surname field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTCNo_Click()
MsgBox "Please enter the X.X. identification number of the financial unit officer delivering the invalid item into the box on the side." & vbNewLine & vbNewLine & _
"The data entered in the ID Number field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblBabaAdi_Click()
MsgBox "Please enter the father's name of the financial unit officer delivering the invalid item into the box on the side." & vbNewLine & vbNewLine & _
"The data entered in the Father's Name field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDogumYeri_Click()
MsgBox "Please enter the place of birth of the financial unit officer delivering the invalid item into the box on the side." & vbNewLine & vbNewLine & _
"The data entered in the Place of Birth field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDogumTarihi_Click()
MsgBox "Please click the calendar icon on the side and select the date of birth of the financial unit officer delivering the invalid item from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Date of Birth field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTelNo_Click()
MsgBox "Please enter the phone number of the organization to which the financial unit officer delivering the invalid item belongs into the box on the side." & vbNewLine & vbNewLine & _
"The data entered in the Phone Number field is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblMuhatapTemasi_Click()
MsgBox "Please select the recipient of the Report 3 transaction from the dropdown list on the side (such as Directorate or Provincial/District Directorate B)." & vbNewLine & vbNewLine & _
"If the Report 3 transaction will be sent to X.X. XXX Governorship Provincial Directorate B XXX Unit Directorate, please select Provincial Directorate B. If it will be sent to X.X. X1 Process Monitoring Directorate XXX Office, please select X.X. X1 Process Monitoring Directorate. The XXX Unit Directorate or XXX Office will be selected from the Sent Unit field on the side." & vbNewLine & vbNewLine & _
"If the relevant Directorate or Decision Board name does not appear in the dropdown list, please click the ± sign on the side and follow the instructions in the window that opens to add the relevant Directorate or Decision Board to the system." & vbNewLine & vbNewLine & _
"The selection in the Recipient Theme field is used in Report 3 statement, Statement 2 report, and cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGonderilenBirim_Click()
MsgBox "Please select the unit to which the Report 3 transaction will be sent from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the Report 3 transaction will be sent to X.X. XXX Governorship Provincial Directorate B XXX Unit Directorate, please select that Unit Directorate. If it will be sent to X.X. X1 Process Monitoring Directorate XXX Office, please select that Office. If the response letter will be sent directly to the recipient specified in the Recipient Theme (without specifying a unit such as XXX Unit Directorate or XXX Office), please select Recipient Theme." & vbNewLine & vbNewLine & _
"If the relevant unit name does not appear in the dropdown list, please click the ± sign on the side and follow the instructions in the window that opens to add the relevant unit to the system." & vbNewLine & vbNewLine & _
"The selection in the Sent Unit field is used in Report 3 statement, Statement 2 report, and cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDesteBandiFotokopisi_Click()
MsgBox "Please select the number of pages of the photocopy of the label(s) (bands) related to the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first digit of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"The selection in the Label Photocopy Page Count field is used in the Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTeslimTutFotokopisi_Click()
MsgBox "Please select the number of pages of the delivery statement photocopy related to the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first digit of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"The selection in the Delivery Statement Photocopy Page Count field is used in Report 3 statement (as Delivery Receipt). For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblBarkodNo_Click()
MsgBox "Please enter the barcode number of the bag in which the invalid item was detected into the box on the side." & vbNewLine & vbNewLine & _
"The barcode number is used in Report 3 statement. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeTuruUst_Click()
MsgBox "Please select the item type related to the item under inspection from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first letter of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"If the relevant item type does not appear in the dropdown list, please click the ± sign to the left of the Item Type label and follow the instructions to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Item Type field is used in Report 3 and Statement 2 reports. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeDegeriUst_Click()
MsgBox "Please select the item value related to the item under inspection from the dropdown list below." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first digit of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"If the relevant item value does not appear in the dropdown list, please click the ± sign to the left of the Nominal Value label and follow the instructions to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Nominal Value field is used in Report 3 and Statement 2 reports. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAdetUst_Click()
MsgBox "Please enter the quantity related to the item under inspection into the box below." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"The quantity entered is used in Report 3, Statement 2, and Directorate cover letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblOgeIdNoUst_Click()
MsgBox "Please enter the item ID number related to the item under inspection into the box below." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"The item ID entered is used in Report 3 and Statement 2 reports. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblAciklamaUst_Click()
MsgBox "You may add a description related to the item under inspection into the box below." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"The description entered is used in Report 3 and Statement 2 reports. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblSonucUst_Click()
MsgBox "Please select the evaluation result related to the item under inspection from the dropdown list below." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"The selection in the Result field is used in the report. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUretimOzelligiUst_Click()
MsgBox "If the evaluation result of the item is invalid, please select the production type (Technique A, Technique B, Technique C, etc.) from the dropdown list below. If the evaluation result is valid, the system will not allow selection in the production field." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first letter of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"The selection in the Production Feature field is used in the report and report 2.2 operation log. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRapor1NoUst_Click()
MsgBox "Please enter the report number into the dropdown list below." & vbNewLine & vbNewLine & _
"If a report number is assigned to each row, the system will generate a separate report for each row. If report numbers are assigned at intervals, the system will merge rows from top to bottom until the next report number and generate a single report." & vbNewLine & vbNewLine & _
"For example, if five item types are entered, with the first three valid and the last two invalid, and if report numbers (e.g., 180-1, 180-2, 180-3, 180-4, 180-5) are assigned to all five rows, the system will generate five separate reports." & vbNewLine & vbNewLine & _
"If report numbers are assigned only to rows 1, 4, and 5, the system will generate three reports in total: one for the valid items in rows 1, 2, and 3, and separate reports for the invalid items in rows 4 and 5." & vbNewLine & vbNewLine & _
"The selection in the Report No field is used in the report and response letter. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporOzelligiUst_Click()
MsgBox "If a report number has been assigned to the item, please select the report feature (normal, feature 1, feature 2, etc.) from the dropdown list below. If no report number is assigned after the first row, the system will not allow selection in the report feature field." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown list, you can also press the first letter of your desired selection on the keyboard repeatedly until it appears." & vbNewLine & vbNewLine & _
"To add multiple item types/item values/quantities, click the + sign on the far right of this row. To remove rows, click the - sign on the far right." & vbNewLine & vbNewLine & _
"If the relevant report feature does not appear in the dropdown list, please click the ± sign to the left of this label and follow the instructions to add it to the system." & vbNewLine & vbNewLine & _
"The selection in the Report Feature field is used in the report. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblNotUst_Click()
MsgBox "If a report number has been assigned to the item, the Add Note option will appear. To add a note to the relevant report (provided a note has been previously defined for the relevant item type in the system), please check the Add Note option. If no report number is assigned after the first row, the Add Note option will not be displayed." & vbNewLine & vbNewLine & _
"If the system does not allow you to add a note for the relevant item type, it means no note has been previously defined for that item type. Please click the ± sign to the left of this label and follow the instructions to define a note for the relevant item type in the system." & vbNewLine & vbNewLine & _
"The selection in the Add Note field is used in the report footnote. For more details, please click the Help button at the top right.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblRapor1Tarihi_Click()
MsgBox "Please click the calendar icon on the side and select the report date from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection made in the Report Date field is used in the report and response letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporImza1_Click()
MsgBox "Please select the person to be displayed in the second signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person's name does not appear in the list, please click the ± sign on the side and follow the instructions in the window that opens to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired person's name repeatedly on the keyboard to find the correct person." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporImza2_Click()
MsgBox "Please select the person to be displayed in the third signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person's name does not appear in the list, please click the ± sign on the side and follow the instructions in the window that opens to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired person's name repeatedly on the keyboard to find the correct person." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblRaporImza3_Click()
MsgBox "Please select the person to be displayed in the first signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person's name does not appear in the list, please click the ± sign on the side and follow the instructions in the window that opens to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired person's name repeatedly on the keyboard to find the correct person." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak2Tarihi_Click()
MsgBox "Please click the calendar icon on the side and select the Statement 2 date from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Statement 2 Date field is used in the Statement 2 report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGidenPaketTipi_Click()
MsgBox "Please select the package and delivery type of the outgoing shipment from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of your desired selection repeatedly on the keyboard to find it." & vbNewLine & vbNewLine & _
"The selection in the Outgoing Package Type field is used in the Statement 2 report and Directorate cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblGidenPaketAdedi_Click()
MsgBox "Please select the quantity of outgoing packages from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first digit of your desired selection repeatedly on the keyboard to find it." & vbNewLine & vbNewLine & _
"The selection in the Outgoing Package Quantity field is used in the Statement 2 report and Directorate cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LbFinansalBirimUstYaziTarihi_Click()
MsgBox "Please click the calendar icon on the side and select the date of the cover letter to be sent to the financial unit from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Financial Unit Cover Letter Date field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirimUstYaziNo_Click()
MsgBox "Please enter the number of the cover letter to be sent to the financial unit, provided by XXX, into the box on the side." & vbNewLine & vbNewLine & _
"The data in the Financial Unit Cover Letter Number field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlFinansalBirim_Click()
MsgBox "Please select the province to which the cover letter for the financial unit will be sent from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired province repeatedly on the keyboard to find it. For example, after clicking once, pressing 'A' once will select Ankara; pressing it a second time will select Adana." & vbNewLine & vbNewLine & _
"To update or add a new province or district, please click the ± sign next to the Province label in the Statement Entry section and follow the instructions in the window that opens to save changes to the system." & vbNewLine & vbNewLine & _
"The selection in the Province field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblIlceFinansalBirim_Click()
MsgBox "Please select the district to which the cover letter for the financial unit will be sent from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking the dropdown once, you can press the first letter of the desired district repeatedly on the keyboard to find it. For example, after clicking once, pressing 'A' once will select Akyurt; pressing it a second time will select Altındağ." & vbNewLine & vbNewLine & _
"To update or add a new district, please click the ± sign next to the Province label in the Statement Entry section and follow the instructions in the window that opens to save changes to the system." & vbNewLine & vbNewLine & _
"The selection in the District field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirimGonderilen_Click()
MsgBox "Please select the financial unit to which the financial unit cover letter will be sent from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the financial unit cover letter will be sent to an entity such as xxxxxxx or the General Directorate, please select xxxxxx or General Directorate; if sent directly to the financial unit, please select the Financial Unit Theme." & vbNewLine & vbNewLine & _
"If the relevant unit does not appear in the list, please click the ± sign on the side and follow the instructions in the window that opens to add the unit to the system." & vbNewLine & vbNewLine & _
"The selection in the Sent Unit field is used in the financial unit cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub LblAdres_Click()
MsgBox "Please enter the address to which the cover letter for the financial unit will be sent in the box on the side." & vbNewLine & vbNewLine & _
"Since province and district selections are specified separately in the two fields on the side, please do not include province or district information when entering the address." & vbNewLine & vbNewLine & _
"The information entered in the Address field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirimGonderiTipi_Click()
MsgBox "Please select the method of dispatch for the response letter addressed to the financial unit from the dropdown list on the side." & vbNewLine & vbNewLine & _
"The selection in the Dispatch Type field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDekontSayfa_Click()
MsgBox "Please select the number of pages of the delivery irregularity receipt related to the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can also press the first digit of your choice repeatedly to find your selection." & vbNewLine & vbNewLine & _
"The selection in the Delivery Irregularity Receipt Page Count field is used in the financial unit's cover letter and Annex-17 statement. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblDesteBandiAdedi_Click()
MsgBox "Please select or enter the number of labels related to the invalid item from the dropdown list on the side." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first digit of your choice repeatedly to find your selection. If the desired quantity is not listed, you may enter it manually." & vbNewLine & vbNewLine & _
"The selection in the Label Quantity field is used in the financial unit's cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziTarihi_Click()
MsgBox "Please click the calendar icon on the side and select the date of the cover letter to be sent to the Directorate from the calendar that appears." & vbNewLine & vbNewLine & _
"The selection in the Cover Letter Date field is used in the Directorate’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub UstYaziNoLabel_Click()
MsgBox "Please enter the number of the cover letter to be sent to the Directorate, as provided by XXX, into the box on the side." & vbNewLine & vbNewLine & _
"The information in the Cover Letter Number field is used in the Directorate’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanakImza1_Click()
MsgBox "Please select the person to be displayed in the first signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in Report 3 statement. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanakImza2_Click()
MsgBox "Please select the person to be displayed in the second signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in Report 3 statement. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanakImza3_Click()
MsgBox "Please select the person to be displayed in the third signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in Report 3 statement. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak2Imza1_Click()
MsgBox "Please select the person to be displayed in the first signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in Statement 2 report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblTutanak2Imza2_Click()
MsgBox "Please select the person to be displayed in the second signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in Statement 2 report. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirimUstYaziImza1_Click()
MsgBox "Please select the person to be displayed in the first signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the financial unit’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblFinansalBirimUstYaziImza2_Click()
MsgBox "Please select the person to be displayed in the second signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the financial unit’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziImza1_Click()
MsgBox "Please select the person to be displayed in the first signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the Directorate’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziImza2_Click()
MsgBox "Please select the person to be displayed in the second signature field from the dropdown list on the side." & vbNewLine & vbNewLine & _
"If the person’s name does not appear in the list, please click the ± sign and follow the instructions to add the person to the system." & vbNewLine & vbNewLine & _
"After clicking once on the dropdown, you can press the first letter of the desired person’s name repeatedly to locate them." & vbNewLine & vbNewLine & _
"The selection in the Signature field is used in the Directorate’s cover letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub

Private Sub LblUstYaziNotu_Click()
MsgBox "To add a note related to XXX Directorate/Decision Board in the Directorate/Decision Board cover letter for invalid types, please check the option on the side." & vbNewLine & vbNewLine & _
"The selection in the XXX Directorate Note field is used in the response letter. For more details, please click the Help button at the top right corner.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End Sub


Private Sub TipAOption_Click()
Dim i As Integer

If TipAOption.Value = True Then
    KayitNoText.Enabled = True
    LblKayitNo.Enabled = True
    TemaTipi.Enabled = True
    LblTemaTipi.Enabled = True
    LblTemaNo.Caption = "Theme 1 No."
    OtomatikOption.Enabled = True
    ManuelOption.Enabled = True
    OtomatikOption.Value = False
    ManuelOption.Value = True
    
    LblBildirilenBirim.Enabled = True
    BildirilenBirimEkleKaldirLabel.Enabled = True
    BildirilenBirim.Enabled = True
    LblDesteBandiFotokopisi.Enabled = True
    DesteBandiFotokopisi.Enabled = True
    LblDesteBandiAdedi.Enabled = True
    DesteBandiAdedi.Enabled = True
    LblBarkodNo.Enabled = True
    BarkodNo.Enabled = True
    
    LblOgeIdNoUst.Enabled = True
    OgeIdNo.Enabled = True
    For i = 1 To 19
        Controls("OgeIdNo" & i).Enabled = True
    Next i


    'RaporlamaGirisi.Enabled = True
    If TutanakGirisi.BackColor = RGB(180, 210, 240) Then
        Call TutanakGirisi_Click
    End If
'    If RaporlamaGirisi.BackColor = RGB(180, 210, 240) Then
'        Call RaporlamaGirisi  '_Click
'    End If
    If Tutanak2Girisi.BackColor = RGB(180, 210, 240) Then
        Call Tutanak2Girisi_Click
    End If
    If FinansalBirimUstYaziGirisi.BackColor = RGB(180, 210, 240) Then
        Call FinansalBirimUstYaziGirisi_Click
    End If
    If UstYaziGirisi.BackColor = RGB(180, 210, 240) Then
        Call UstYaziGirisi_Click
    End If

    UstYaziNotuCheck.Value = False
    UstYaziNotuFrame.Visible = True
    
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

End Sub

Private Sub TipBOption_Click()
Dim i As Integer

If TipBOption.Value = True Then
    KayitNoText.Value = ""
    TemaTipi.Value = ""
    BildirilenBirim.Value = ""
    DesteBandiFotokopisi.Value = ""
    BarkodNo.Value = ""
    DesteBandiAdedi.Value = ""
    
    KayitNoText.Enabled = False
    LblKayitNo.Enabled = False
    TemaTipi.Enabled = False
    LblTemaTipi.Enabled = False
    
    LblTemaNo.Caption = "Theme 2 No."
    OtomatikOption.Enabled = False
    ManuelOption.Enabled = True
    OtomatikOption.Value = False
    ManuelOption.Value = True
    
    LblBildirilenBirim.Enabled = False
    BildirilenBirimEkleKaldirLabel.Enabled = False
    BildirilenBirim.Enabled = False
    LblDesteBandiFotokopisi.Enabled = False
    DesteBandiFotokopisi.Enabled = False
    LblDesteBandiAdedi.Enabled = False
    DesteBandiAdedi.Enabled = False
    LblBarkodNo.Enabled = False
    BarkodNo.Enabled = False


    'RaporlamaGirisi.Enabled = False
    'If RaporlamaGirisi.BackColor = RGB(180, 210, 240) Then
        Call TutanakGirisi_Click
    'End If
    If Tutanak2Girisi.BackColor = RGB(180, 210, 240) Then
        Call Tutanak2Girisi_Click
    End If
    If FinansalBirimUstYaziGirisi.BackColor = RGB(180, 210, 240) Then
        Call FinansalBirimUstYaziGirisi_Click
    End If
    If UstYaziGirisi.BackColor = RGB(180, 210, 240) Then
        Call UstYaziGirisi_Click
    End If

'    If TutanakGirisi.BackColor = RGB(180, 210, 240) Then
'        GoTo Son
'    End If
  
  
    Sonuc.Visible = False
    LblSonuc.Visible = False
    
    Rapor1No.Visible = False
    LblRapor1No.Visible = False
    RaporOzelligi.Visible = False
    LblRaporOzelligi.Visible = False
    UretimOzelligi.Visible = False
    LblUretimOzelligi.Visible = False
    'Rapor2_2No.Visible = False
    'LblRapor2_2No.Visible = False
    
    LblSonucUst.Visible = False
    LblRapor1NoUst.Visible = False
    LblRapor1NoUst.Visible = False
    LblRaporOzelligiUst.Visible = False
    RaporOzelligiEkleKaldirLabel.Visible = False
    LblUretimOzelligiUst.Visible = False
    'LblRapor2_2NoUst.Visible = False
    NotEkleKaldirLabel.Visible = False
    LblNotUst.Visible = False
    NotCheck.Visible = False
    
    LblOgeIdNoUst.Enabled = False
    OgeIdNo.Value = ""
    OgeIdNo.Enabled = False
    For i = 1 To 19
        Controls("Sonuc" & i).Visible = False
        Controls("LblSonuc" & i).Visible = False
        Controls("Rapor1No" & i).Visible = False
        Controls("LblRapor1No" & i).Visible = False
        Controls("RaporOzelligi" & i).Visible = False
        Controls("LblRaporOzelligi" & i).Visible = False
        Controls("UretimOzelligi" & i).Visible = False
        Controls("LblUretimOzelligi" & i).Visible = False
    '    Controls("Rapor2_2No" & i).Visible = False
    '    Controls("LblRapor2_2No" & i).Visible = False
        Controls("NotCheck" & i).Visible = False
        
        Controls("OgeIdNo" & i).Value = ""
        Controls("OgeIdNo" & i).Enabled = False
    Next i
    
    EkleOge.Left = 518
    KaldirOge.Left = 538
    
    Rapor1Frame.Visible = False
    UstYaziNotuCheck.Value = False
    UstYaziNotuFrame.Visible = False
    
End If

Son:

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

End Sub

Sub Rapor3FinansalBirimFormunuResetle()
Dim i As Integer
Dim ctl As MSForms.Control

ThisWorkbook.Activate

'TipAOption.Value = False
'TipBOption.Value = False

Nokta2Option.Value = True
Nokta3Option.Value = False

Il.Value = ""
Ilce.Value = ""
TutanakTarihiText.Value = ""
KayitNoText.Value = ""
TemaTipi.Value = ""
TemaNoText.Value = ""
OtomatikOption.Value = False
ManuelOption.Value = True
FinansalBirim.Value = ""
TeslimatBirim.Value = ""
BildirilenBirim.Value = ""
TeslimTarihiText.Value = ""
AdSoyad.Value = ""
TCNo.Value = ""
BabaAdi.Value = ""
DogumYeri.Value = ""
DogumTarihiText.Value = ""
TelNo.Value = ""
MuhatapTemasi.Value = ""
GonderilenBirim.Value = ""
DesteBandiFotokopisi.Value = ""
TeslimTutFotokopisi.Value = ""
BarkodNo.Value = ""

TutanakImza1.Value = ""
TutanakImza2.Value = ""
TutanakImza3.Value = ""
RaporImza1.Value = ""
RaporImza2.Value = ""
RaporImza3.Value = ""
Tutanak2Imza1.Value = ""
Tutanak2Imza2.Value = ""
FinansalBirimUstYaziImza1.Value = ""
FinansalBirimUstYaziImza2.Value = ""
UstYaziImza1.Value = ""
UstYaziImza2.Value = ""

OgeTuru.Value = ""
OgeDegeri.Value = ""
Adet.Value = ""
OgeIdNo.Value = ""
Aciklama.Value = ""

Sonuc.Value = ""
Rapor1No.Value = ""
NotCheck.Value = False
UretimOzelligi.Value = ""
RaporOzelligi.Value = ""
Rapor1TarihiText.Value = ""


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
        Controls("UretimOzelligi" & i).Value = ""
        Controls("NotCheck" & i).Value = False
        Controls("RaporOzelligi" & i).Value = ""
        Controls("Rapor1No" & i).Value = ""
    'End If
Next i
For i = 1 To 19
    Call KaldirOge_Click
Next i

'Tutanak2
Tutanak2TarihiText.Value = ""
GidenPaketTipi.Value = ""
GidenPaketAdedi.Value = ""

'FinansalBirim üst yazı
FinansalBirimUstYaziTarihiText.Value = ""
FinansalBirimUstYaziNoText.Value = ""
IlFinansalBirim.Value = ""
IlceFinansalBirim.Value = ""
FinansalBirimGonderilen.Value = ""
Adres.Value = ""
DekontSayfa.Value = ""
FinansalBirimGonderiTipi.Value = ""
DesteBandiAdedi.Value = ""

'Üst yazı
UstYaziTarihiText.Value = ""
UstYaziNoText.Value = ""
UstYaziNotuCheck.Value = False

'Taslak Renklerini resetle
For Each ctl In core_report3_2_entry_UI.Controls
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

Private Sub LblDuzeltme_Click()
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range
Dim Fark As Long
Dim i As Long, OgeFrame As Integer
Dim ctl As MSForms.Control, Resetle As Integer

'Columns("FG:FH").EntireColumn.Hidden = False

'Application.EnableEvents = False

'Application.ScreenUpdating = False

ThisWorkbook.Activate

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(5).Unprotect Password:="123"

'Rapor1 formunu resetle
Call UstYaziGirisi_Click
Call Rapor3FinansalBirimFormunuResetle

If ComboGetir.Value = "" Then
    LblDuzeltme.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
    GoTo Son
End If

'Veri tabanını kontrol et
Say = Range("FG100000").End(xlUp).Row
If Say < 7 Or ComboGetir.Value = "" Then
    GoTo Son
End If

Set IlkSiraBul = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
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

If ThisWorkbook.Worksheets(5).Range("L" & IlkSira) = "Point2" Or ThisWorkbook.Worksheets(5).Range("L" & IlkSira) = "Point3" Then
    '
Else
    MsgBox "The entered serial number does not belong to a Point 2/Point 3 transaction, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Verileri sayfadan finansal birim formuna aktar.
'Tutanak bölümü
If Cells(IlkSira, 28).Value = "Type A" Then
    TipAOption.Value = True
ElseIf Cells(IlkSira, 28).Value = "Type B" Then
    TipBOption.Value = True
Else
    TipAOption.Value = False
    TipBOption.Value = False
End If

If Cells(IlkSira, 12).Value = "Point2" Then
    Nokta2Option.Value = True
ElseIf Cells(IlkSira, 12).Value = "Point3" Then
    Nokta3Option.Value = True
Else
    Nokta2Option.Value = True
    Nokta3Option.Value = False
End If

Il.Value = Cells(IlkSira, 19).Value
Ilce.Value = Cells(IlkSira, 20).Value
TutanakTarihiText.Value = Cells(IlkSira, 23).Value
KayitNoText.Value = Cells(IlkSira, 24).Value
TemaTipi.Value = Cells(IlkSira, 25).Value
TemaNoText.Value = Cells(IlkSira, 26).Value
If Cells(IlkSira, 27).Value = "Otomatik" Then
    OtomatikOption.Value = True
ElseIf Cells(IlkSira, 27).Value = "Manuel" Then
    ManuelOption.Value = True
Else
    OtomatikOption.Value = False
    ManuelOption.Value = False
End If
FinansalBirim.Value = Cells(IlkSira, 30).Value
TeslimatBirim.Value = Cells(IlkSira, 31).Value
BildirilenBirim.Value = Cells(IlkSira, 32).Value
TeslimTarihiText.Value = Cells(IlkSira, 35).Value
AdSoyad.Value = Cells(IlkSira, 36).Value
TCNo.Value = Cells(IlkSira, 38).Value
BabaAdi.Value = Cells(IlkSira, 39).Value
DogumYeri.Value = Cells(IlkSira, 40).Value
DogumTarihiText.Value = Cells(IlkSira, 41).Value
TelNo.Value = Cells(IlkSira, 44).Value
MuhatapTemasi.Value = Cells(IlkSira, 47).Value
If Cells(IlkSira, 48).Value <> "" Then
    GonderilenBirim.Value = Cells(IlkSira, 48).Value
Else
    GonderilenBirim.Value = "Contact Theme"
End If
DesteBandiFotokopisi.Value = Cells(IlkSira, 49).Value
TeslimTutFotokopisi.Value = Cells(IlkSira, 50).Value
BarkodNo.Value = Cells(IlkSira, 51).Value

TutanakImza1.Value = Cells(IlkSira, 184).Value
TutanakImza2.Value = Cells(IlkSira, 187).Value
TutanakImza3.Value = Cells(IlkSira, 190).Value
RaporImza1.Value = Cells(IlkSira, 220).Value
RaporImza2.Value = Cells(IlkSira, 223).Value
RaporImza3.Value = Cells(IlkSira, 226).Value
Tutanak2Imza1.Value = Cells(IlkSira, 193).Value
Tutanak2Imza2.Value = Cells(IlkSira, 196).Value
FinansalBirimUstYaziImza1.Value = Cells(IlkSira, 199).Value
FinansalBirimUstYaziImza2.Value = Cells(IlkSira, 202).Value
UstYaziImza1.Value = Cells(IlkSira, 205).Value
UstYaziImza2.Value = Cells(IlkSira, 208).Value

OgeTuru.Value = Cells(IlkSira, 52).Value
OgeDegeri.Value = Cells(IlkSira, 55).Value
Adet.Value = Cells(IlkSira, 58).Value
OgeIdNo.Value = Cells(IlkSira, 61).Value
Aciklama.Value = Cells(IlkSira, 64).Value

Call TutanakGirisi_Click

'Rapor1
If Cells(IlkSira, 212).Value <> "" Or Cells(IlkSira, 217).Value <> "" Or Cells(IlkSira, 218).Value <> "" _
Or Cells(IlkSira, 213).Value <> "" Or Cells(IlkSira, 214).Value <> "" Then
    'Rapor1Frame.Visible = True
    Call RaporlamaGirisiPro  '_Click
    Sonuc.Value = Cells(IlkSira, 212).Value
    UretimOzelligi.Value = Cells(IlkSira, 213).Value
    RaporOzelligi.Value = Cells(IlkSira, 214).Value
    'Rapor1No.Clear
    'Call Rapor1NoListClear
    If Cells(IlkSira, 216).Value = "Yes" Then
        NotCheck.Value = True
    Else
        NotCheck.Value = False
    End If
    Rapor1No.Value = Cells(IlkSira, 217).Value
    Rapor1TarihiText.Value = Cells(IlkSira, 218).Value
End If

Fark = SonSira - IlkSira + 1
If Fark > 1 And Fark < 21 Then
    For OgeFrame = 1 To Fark - 1
        'Controls("OgeTuruFrame" & OgeFrame).Visible = True
        Call EkleOge_Click
    Next OgeFrame
    For OgeFrame = 1 To Fark - 1
        Controls("OgeTuru" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 52).Value
        Controls("OgeDegeri" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 55).Value
        Controls("Adet" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 58).Value
        Controls("OgeIdNo" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 61).Value
        Controls("Aciklama" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 64).Value
        'Rapor1
        If Cells(IlkSira + OgeFrame, 212).Value <> "" Or Cells(IlkSira + OgeFrame, 217).Value <> "" Then
            'Rapor1Frame.Visible = True
            Call RaporlamaGirisiPro  '_Click
            'Rapor1No.Clear
            Controls("Sonuc" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 212).Value
            Controls("UretimOzelligi" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 213).Value
            Controls("RaporOzelligi" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 214).Value
            'Controls("Rapor1No" & OgeFrame).Clear
            'Call Rapor1NoListClear
            If Cells(IlkSira + OgeFrame, 217).Value <> "" And Cells(IlkSira + OgeFrame, 216).Value = "Yes" Then
                Controls("NotCheck" & OgeFrame).Value = True
            ElseIf Cells(IlkSira + OgeFrame, 217).Value <> "" And (Cells(IlkSira + OgeFrame, 216).Value = "No" Or Cells(IlkSira + OgeFrame, 216).Value = "") Then
                Controls("NotCheck" & OgeFrame).Value = False
            End If
            Controls("Rapor1No" & OgeFrame).Value = Cells(IlkSira + OgeFrame, 217).Value
        End If
    Next OgeFrame
End If

'Tutanak2
If Cells(IlkSira, 69).Value <> "" Or Cells(IlkSira, 71).Value <> "" Or Cells(IlkSira, 72).Value <> "" Then
    Call Tutanak2Girisi_Click
    'Call Rapor1NoListClear
    Tutanak2TarihiText.Value = Cells(IlkSira, 69).Value
    GidenPaketTipi.Value = Cells(IlkSira, 71).Value
    GidenPaketAdedi.Value = Cells(IlkSira, 72).Value
End If

'FinansalBirim üst yazı
If Cells(IlkSira, 75).Value <> "" Or Cells(IlkSira, 76).Value <> "" Or Cells(IlkSira, 77).Value <> "" _
Or Cells(IlkSira, 78).Value <> "" Or Cells(IlkSira, 79).Value <> "" Or Cells(IlkSira, 80).Value <> "" Or Cells(IlkSira, 81).Value <> "" Then
    Call FinansalBirimUstYaziGirisi_Click
    'Call Rapor1NoListClear
    FinansalBirimUstYaziTarihiText.Value = Cells(IlkSira, 75).Value
    FinansalBirimUstYaziNoText.Value = Cells(IlkSira, 76).Value
    IlFinansalBirim.Value = Cells(IlkSira, 77).Value
    IlceFinansalBirim.Value = Cells(IlkSira, 78).Value
    If Cells(IlkSira, 82).Value <> "" Then
        FinansalBirimGonderilen.Value = Cells(IlkSira, 82).Value
    Else
        FinansalBirimGonderilen.Value = "Financial Unit Theme"
    End If
    Adres.Value = Cells(IlkSira, 79).Value
    DekontSayfa.Value = Cells(IlkSira, 80).Value
    FinansalBirimGonderiTipi.Value = Cells(IlkSira, 85).Value
    DesteBandiAdedi.Value = Cells(IlkSira, 81).Value
End If

'Üst yazı
If Cells(IlkSira, 83).Value <> "" Or Cells(IlkSira, 84).Value <> "" Then
    Call UstYaziGirisi_Click
    'Call Rapor1NoListClear
    UstYaziTarihiText.Value = Cells(IlkSira, 83).Value
    UstYaziNoText.Value = Cells(IlkSira, 84).Value
    If Cells(IlkSira, 215).Value = "Yes" Then
        UstYaziNotuCheck.Value = True
    Else
        UstYaziNotuCheck.Value = False
    End If
End If

LblDuzeltme.BackColor = RGB(180, 210, 240)
LblDuzeltme.ForeColor = RGB(30, 30, 30)

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat


Son:

'Columns("FG:FH").EntireColumn.Hidden= True

ThisWorkbook.Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

'Application.ScreenUpdating = True

'Application.EnableEvents = True

End Sub

Private Sub LblTaslak_Click()
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range
Dim Fark As Long
Dim i As Long, OgeFrame As Integer
Dim ctl As MSForms.Control, Resetle As Integer

'Columns("FG:FH").EntireColumn.Hidden = False

ThisWorkbook.Activate

KilitIptal = True

Call LblDuzeltme_Click
ComboGetir.Value = ""

If TipBOption.Value = False Then
    'Rapor1 no değerlerini sıfırla
    Call Son20RaporNo
    Rapor1No.Value = ""
    NotCheck.Value = False
    For i = 1 To 19
        Controls("Rapor1No" & i).Value = ""
        Controls("NotCheck" & i).Value = False
    Next i
End If

LblDuzeltme.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
LblDuzeltme.ForeColor = RGB(30, 30, 30)

'Taslak Renkler
For Each ctl In core_report3_2_entry_UI.Controls
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

'Columns("FG:FH").EntireColumn.Hidden = True

End Sub

Sub ComboGetirReset()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(5).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If
'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(5).Range("L" & i) = "Point2" Or ThisWorkbook.Worksheets(5).Range("L" & i) = "Point3" Then
        With ComboGetir
            .AddItem (ThisWorkbook.Worksheets(5).Range("E" & i).Value)
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
ThisWorkbook.Worksheets(5).Unprotect Password:="123"
ThisWorkbook.Worksheets(10).Unprotect Password:="123"

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 2.1.xlsx"

If ComboGetir.Value = "" Then
    MsgBox "To delete a record from the system, select the serial number of the transaction you want to delete from the dropdown menu located between the Edit and Draft buttons, then click the Edit button. After confirming that the correct record has been loaded, click the Delete button and follow the instructions in the pop-up window.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Veri tabanını kontrol et
Say = ThisWorkbook.Worksheets(5).Range("FG100000").End(xlUp).Row
If Say < 7 Or ComboGetir.Value = "" Then
    GoTo Out
End If

Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
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

If ThisWorkbook.Worksheets(5).Range("L" & IlkSira) = "Point2" Or ThisWorkbook.Worksheets(5).Range("L" & IlkSira) = "Point3" Then
    '
Else
    MsgBox "The entered serial number does not belong to a Point 2/Point 3 transaction, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'Registry Reports klasör adını kontrol et.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox IslemGunlukleriKlasor & " directory cannot be accessed. The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox IslemGunlugu & " directory cannot be accessed. The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'Hazırlık
SiraNoSakla = ThisWorkbook.Worksheets(5).Cells(IlkSira, 5).Value
OncekiSiraNo = ThisWorkbook.Worksheets(5).Cells(IlkSira, 5).Value - 1
Set WsRapor = ThisWorkbook.Worksheets(5)

Sifre = InputBox(Prompt:="To delete the operation with serial number " & ThisWorkbook.Worksheets(5).Cells(IlkSira, 5).Value & " from the system, please enter the password value '123'.", Title:="Enterprise Document Automation System")
If Sifre = "123" Then

    'RAPOR İŞLEM GÜNLÜĞÜ
    'İşlem günlüğü açıksa kaydet ve kapat.
    OpenControl = IsWorkBookOpen(IslemGunlugu)
    If OpenControl = True Then
        Workbooks("System Registry Report 2.1.xlsx").Save
        Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
    End If
    Workbooks.Open (IslemGunlugu)
    Set WsIslemGunlugu = Workbooks("System Registry Report 2.1.xlsx").Worksheets(1)
    WsIslemGunlugu.Unprotect Password:="123"
    WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = False
    
    'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
    Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
    Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
    SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row
    
    Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IslemGunluguIlkSiraBul Is Nothing Then
        IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
        If Not IslemGunluguSonSiraBul Is Nothing Then
            IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
        End If
    
        'kayıt def. verileri sil, satırları işaretle
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
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
            Set Kenarlar = WsIslemGunlugu.Range("D" & IslemGunluguIlkSira - 1 & ":T" & IslemGunluguIlkSira - 1)
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
        Workbooks("System Registry Report 2.1.xlsx").Save
        Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
    End If
    
    'MODÜL işlemleri
    Set WsRapor = ThisWorkbook.Worksheets(5)
    'On Error Resume Next
    'Sira numaralarını düzelt
    If Say > IlkSira Then
        For i = IlkSira + 1 To Say
            If WsRapor.Cells(i, 5).Value <> "" Then
                OncekiSiraNo = OncekiSiraNo + 1
                WsRapor.Cells(i, 5).Value = OncekiSiraNo
                WsRapor.Cells(i, 163).Value = OncekiSiraNo 'başlangıç
                
                For j = i To i + 1000
                    If WsRapor.Cells(j, 164).Value <> "" Then
                        WsRapor.Cells(j, 164).Value = OncekiSiraNo 'bitiş
                        GoTo DonguJSon
                    End If
                Next j
DonguJSon:
            End If
        Next i
    
    ElseIf Say = IlkSira Then
        'MsgBox " Modül: Güncellenecek no yok!"
    End If

    'Rapor no sayfasında silme işlemini yap
    
    
    '__________Rapor No Senkronizasyon 30.11.2021

    Set WsRaporNo = ThisWorkbook.Worksheets(10)

    Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
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
    Call Rapor3FinansalBirimFormunuResetle
    Call ComboGetirReset
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    MsgBox "The operation with serial number " & SiraNoSakla & " has been successfully deleted from the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

ElseIf Sifre = vbCancel Then
    'MsgBox "Şifre iptal"
    GoTo Out
ElseIf Sifre <> "" And Sifre <> "123" Then
    MsgBox "The operation with serial number " & SiraNoSakla & " could not be deleted from the system due to an incorrect password.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If 'Şifre koşulu sonu


Out:

'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If

ThisWorkbook.Activate
            

ThisWorkbook.Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(10).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

End Sub


Private Sub MuhatapTemasiEkleKaldirLabel_Click()
support_contact_themes_UI.Show
'support_contact_themes_UI.Show vbModeless
End Sub
Private Sub GonderilenBirimEkleKaldirLabel_Click()
support_contact_subunits_UI.Show
'support_contact_subunits_UI.Show vbModeless
End Sub
Private Sub OgeEkleKaldirLabel_Click()
support_item_types_UI.Show
'support_item_types_UI.Show vbModeless
End Sub
Private Sub OgeDegeriEkleKaldirLabel_Click()
support_item_values_UI.Show
'support_item_values_UI.Show vbModeless
End Sub

Private Sub NotEkleKaldirLabel_Click()
support_item_type_notes_UI.Show
'support_item_type_notes_UI.Show vbModeless
End Sub
Private Sub RaporOzelligiEkleKaldirLabel_Click()
support_report_templates_UI.Show
'support_report_templates_UI.Show vbModeless
End Sub

Private Sub TeslimatBirimEkleKaldirLabel_Click()
support_delivery_units_UI.Show
'support_delivery_units_UI.Show vbModeless
End Sub
Private Sub BildirilenBirimEkleKaldirLabel_Click()
support_notification_units_UI.Show
'support_notification_units_UI.Show vbModeless
End Sub
Private Sub FinansalBirimEkleKaldirLabel_Click()
support_financial_themes_UI.Show
'support_financial_themes_UI.Show vbModeless
End Sub
Private Sub IlIlceEkleKaldirLabel_Click()
support_provinces_districts_UI.Show
'support_provinces_districts_UI.Show vbModeless
End Sub
Private Sub IlIlceEkleKaldirLabel2_Click()
support_provinces_districts_UI.Show
'support_provinces_districts_UI.Show vbModeless
End Sub

Private Sub TutanakImza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub TutanakImza2EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub TutanakImza3EkleKaldirLabel_Click()
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
Private Sub RaporImza3EkleKaldirLabel_Click()
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
Private Sub FinansalBirimGonderilenEkleKaldirLabel_Click()
support_delivery_units_UI.Show
'support_delivery_units_UI.Show vbModeless
End Sub
Private Sub FinansalBirimUstYaziImza1EkleKaldirLabel_Click()
support_signatures_UI.Show
'support_signatures_UI.Show vbModeless
End Sub
Private Sub FinansalBirimUstYaziImza2EkleKaldirLabel_Click()
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
SourceTaslak = AutoPath & "\System Files\Help Documents\Report 3.2 Entry – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory is not accessible. The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the name of the Operation folder.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory is not accessible. The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'RmDir DestOpUserFolder 'Sistem kapanırken DestOpUserFolder klasörünü temizle EKLENECEK!
'_______________

'Klasör isimlerini kontrol et.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox SourceTaslak & " directory is not accessible. The names of the folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Sub KontrolProseduru()
Dim YeniIslem As Long ', SiraBul As Range, SiraKontrol As Range
Dim i As Long, j As Long, OgeFrame As Integer, Kont As Integer
Dim ctl As MSForms.Control
Dim Bilgi As Variant
Dim OgeTuruKont As Integer, OgeDegeriKont As Integer, AdetKont As Integer
Dim OgeIdNoKont As Integer, AciklamaKont As Integer, SonucKont As Integer ', MaxiR As Integer, Maxi As Integer
Dim OgeTuruKontSatir As Integer, OgeDegeriKontSatir As Integer, AdetKontSatir As Integer
Dim OgeIdNoKontSatir As Integer, AciklamaKontSatir As Integer, SonucKontSatir As Integer
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range, Fark As Long
Dim FarkSay As Integer, SiraNoSakla As Long, SiraSay As Long
Dim Kenarlar As Range, DokumKontSatir As Integer, UserName As String

Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim OpenControl As String, Say1IslemGunlugu As Long, Say2IslemGunlugu As Long
Dim GelenTema As String, Sene As String, Ay As String
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long
Dim Olay1 As String, Olay2 As String, Olay3 As String, Olay4 As String, Olay5 As String
Dim ItemBul As Range

Dim Rapor1TarihBul As Range, Rapor1NoBulIlk As Range, RefSatir As Long
Dim UretimOzelligiKont As Integer, RaporOzelligiKont As Integer, Rapor1NoBulTireKont As Integer, Rapor1NoBulKont As Integer, Rapor1NoBulTireKontPart As Integer, Rapor1NoKont As Integer
Dim UretimOzelligiKontSatir As Integer, RaporOzelligiKontSatir As Integer, Rapor1NoKontAyni As Integer, Rapor1NoKontAltNoHata As Integer, Rapor1NoKontUstNoHata As Integer
Dim Rapor1NoBul As Range, Rapor1NoBulTire As Range, Rapor1NoBulTirePart As Range

Dim SonucKontrol As Boolean


'Tutanak kontrolleri
TutKont = 0
If TipAOption.Value = False And TipBOption.Value = False Then
    Bilgi = MsgBox("The item type has not been specified as either Type A or Type B. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet1Ek2
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet1Ek2:
If Nokta2Option.Value = False And Nokta3Option.Value = False Then
    Bilgi = MsgBox("It has been detected that the transaction type has not been specified as either Point 2 or Point 3. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet1Ek1
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet1Ek1:
If Il.Value = "" Then
    Bilgi = MsgBox("It has been detected that the province has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet1
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet1:
If InStr(MuhatapTemasi.Value, "İlçe") <> 0 Then
    If Ilce.Value = "" Then
        Bilgi = MsgBox("Although the recipient theme includes a district, no district has been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            TutKont = 1
            GoTo YinedeKaydet2
        ElseIf Bilgi = vbNo Then
            TutKont = 2
            GoTo Son
        End If
    End If
End If

YinedeKaydet2:
If TutanakTarihiText.Value = "" Then
    Bilgi = MsgBox("It has been detected that the statement date has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet3
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet3:


If OtomatikOption.Value = False And ManuelOption.Value = False Then
    Bilgi = MsgBox("The creation mode of the theme number (Automatic/Manual) has not been selected. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet6
    ElseIf Bilgi = vbNo Then
        TutKont = 2
    End If
End If
YinedeKaydet6:

If FinansalBirim.Value = "" Then
    Bilgi = MsgBox("The financial unit has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7Ek2
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7Ek2:

If TeslimatBirim.Value = "" Then
    Bilgi = MsgBox("The delivery unit has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7Ek3
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7Ek3:

If TipAOption.Value = True And BildirilenBirim.Value = "" Then
    Bilgi = MsgBox("The unit related to the invalid item has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7Ek4
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7Ek4:

If TeslimTarihiText.Value = "" Then
    Bilgi = MsgBox("The delivery date has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7Ek5
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7Ek5:

If MuhatapTemasi.Value = "" Then
    Bilgi = MsgBox("The contact theme has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7Ek1
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7Ek1:

If GonderilenBirim.Value = "" Then
    Bilgi = MsgBox("The receiving unit has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet7
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet7:

If TipAOption.Value = True And DesteBandiFotokopisi.Value = "" Then
    Bilgi = MsgBox("The number of pages of the band label photocopy has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet8
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet8:

If TeslimTutFotokopisi.Value = "" Then
    Bilgi = MsgBox("The number of pages of the delivery statement photocopy has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet9
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet9:

If TipAOption.Value = True And BarkodNo.Value = "" Then
    Bilgi = MsgBox("The barcode number of the package containing the invalid item has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet9BarkodEki
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If
YinedeKaydet9BarkodEki:

If OgeTuru.Value = "" Then
    Bilgi = MsgBox("The item type has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet15
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet15:
If OgeDegeri.Value = "" Then
    Bilgi = MsgBox("The item value has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet16
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet16:
If Adet.Value = "" Then
    Bilgi = MsgBox("The quantity has not been entered. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet17
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet17:
If TipAOption.Value = True Then
    If OgeIdNo.Value = "" Then
        Bilgi = MsgBox("The item ID number has not been entered. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            TutKont = 1
            GoTo YinedeKaydet18
        ElseIf Bilgi = vbNo Then
            TutKont = 2
            GoTo Son
        End If
    End If
End If

YinedeKaydet18:

If Aciklama.Value = "" Then
    ' You can optionally add:
    ' MsgBox "The description has not been entered.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
End If


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
        If TipAOption.Value = True Then
            If Controls("OgeIdNo" & i).Value = "" Then
                OgeIdNoKontSatir = i
            End If
        End If
'        If Controls("Aciklama" & i).Value = "" Then
'            AciklamaKontSatir = i
'        End If
    Next i
End If
'Yukarıdaki maxi değeri, (aşağıda bulunan kodlarda) verilerin Rapor3 finansal birim formundan
'sayfaya aktarılmasında kullanılıyor.
If OgeTuruKontSatir <> 0 And OgeDegeriKontSatir <> 0 And AdetKontSatir <> 0 And OgeIdNoKontSatir Then
    Bilgi = MsgBox("A missing row has been detected. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet19
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
End If

YinedeKaydet19:
If OgeTuruKontSatir <> 0 Then
    Bilgi = MsgBox("The item type was found to be incomplete. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
ElseIf OgeDegeriKontSatir <> 0 Then
    Bilgi = MsgBox("The item value was found to be incomplete. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
ElseIf AdetKontSatir <> 0 Then
    Bilgi = MsgBox("The quantity was found to be incomplete. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
ElseIf OgeIdNoKontSatir <> 0 Then
    Bilgi = MsgBox("The item ID number was found to be incomplete. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
    If Bilgi = vbYes Then
        TutKont = 1
        GoTo YinedeKaydet20
    ElseIf Bilgi = vbNo Then
        TutKont = 2
        GoTo Son
    End If
ElseIf AciklamaKontSatir <> 0 Then
    ' (Optional English message can be added here)
End If

YinedeKaydet20:


'Rapor1 kontrolleri
'Rapor1Kont = 2
If Rapor1Frame.Visible = True Then
    Rapor1Kont = 0
    
    If Sonuc.Value = "" Then
        Bilgi = MsgBox("The result field has not been specified. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet21
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If

YinedeKaydet21:
    If RaporOzelligi.Value = "" Then
        Bilgi = MsgBox("The report feature has not been specified. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet21A
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If

YinedeKaydet21A:
    If UretimOzelligi.Enabled = True And UretimOzelligi.Value = "" Then
        Bilgi = MsgBox("The production feature has not been specified. To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet21B
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If

YinedeKaydet21B:

    If Rapor1No.Value <> "" Then
        If InStr(Rapor1No.Value, "-") = 0 Then
            ' No hyphen found – no validation
        Else
            If Mid(Rapor1No.Value, InStr(Rapor1No.Value, "-") + 1, 1) <> 1 Then
                Bilgi = MsgBox("It was detected that the sub-number of the first row report does not start with 1 (e.g., it starts with 18-2 instead of 18-1). To save anyway, click ""Yes""; to make corrections, click ""No"".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
    Set WsRaporNo = ThisWorkbook.Worksheets(10)
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
        Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
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
                Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
                    Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
                        Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
                Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
                    Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
                        Bilgi = MsgBox("It has been detected that the first report number has already been used. To save anyway, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
    UretimOzelligiKont = 0
    RaporOzelligiKont = 0
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
            If Controls("UretimOzelligi" & OgeFrame).Value <> "" Then
                UretimOzelligiKont = OgeFrame
            End If
            If Controls("RaporOzelligi" & OgeFrame).Value <> "" Then
                RaporOzelligiKont = OgeFrame
            End If
        Next OgeFrame
    End If
    
    OgeTuruKontSatir = 0
    OgeDegeriKontSatir = 0
    AdetKontSatir = 0
    OgeIdNoKontSatir = 0
    AciklamaKontSatir = 0
    SonucKontSatir = 0
    UretimOzelligiKontSatir = 0
    RaporOzelligiKontSatir = 0
    Rapor1NoKontAyni = 0
    Rapor1NoKontAltNoHata = 0
    Rapor1NoKontUstNoHata = 0
    
    MaxiR = Application.Max(OgeTuruKont, OgeDegeriKont, AdetKont, OgeIdNoKont, AciklamaKont, SonucKont, UretimOzelligiKont, RaporOzelligiKont)
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
            If Controls("UretimOzelligi" & i).Enabled = True And Controls("UretimOzelligi" & i).Value = "" Then
                UretimOzelligiKontSatir = i
            End If
            If Controls("Rapor1No" & i).Value <> "" And Controls("RaporOzelligi" & i).Value = "" Then
                RaporOzelligiKontSatir = i
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
        Set RnoIlkSiraBul = WsRaporNo.Range("J6:J100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("K6:K100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
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
        Set RnoIlkSiraBul = WsRaporNo.Range("J6:J100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
        Set RnoSonSiraBul = WsRaporNo.Range("K6:K100000").Find(What:=Cells(IlkSiraGlobal, 165).Value, SearchDirection:=xlNext, _
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
        Bilgi = MsgBox("It has been detected that at least one of the result rows is either missing or incomplete. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet22
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet22:
    If UretimOzelligiKontSatir <> 0 Then
        Bilgi = MsgBox("It has been detected that at least one of the production feature rows is either missing or incomplete. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet22A
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet22A:
    If RaporOzelligiKontSatir <> 0 Then
        Bilgi = MsgBox("It has been detected that at least one of the report feature rows is either missing or incomplete. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet22B
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet22B:
    If Rapor1No.Value = "" Then
        Bilgi = MsgBox("It has been detected that the report number has not been entered. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23:

    'Check for duplicate report numbers in comboboxes
    If Rapor1NoKontAyni <> 0 Then
        Bilgi = MsgBox("It has been detected that at least one of the report number fields contains a duplicate entry. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("It has been detected that at least one of the report number fields is missing a sub-number (e.g., like 318 or 318-1). To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("It has been detected that the main number (e.g., 318-1 and 319-2 or 318 and 319) differs between at least one of the report number fields. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1C
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1C:
    'Check validity of report number
    If Rapor1NoKont <> 0 Then
        Bilgi = MsgBox("It has been detected that the report number is missing or incorrectly entered. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("It has been detected that only a single sub-number (e.g., only 82-1) was entered in the report number field. To proceed with saving, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet23Ek1TireTek
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet23Ek1TireTek:

    'If ComboGetir.Value = "" Then 'Do not start the procedure if it's a correction entry
        'Check if the Rapor1 number has already been used (numbers before the dash included both in the form and in the system)
        If Rapor1NoBulKont <> 0 Then
            Bilgi = MsgBox("It has been detected that at least one of the report numbers has already been used. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet23Ek2
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
        If Rapor1NoBulTireKont <> 0 Then
            Bilgi = MsgBox("It has been detected that at least one of the report numbers has already been used. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet23Ek2
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
        If Rapor1NoBulTireKontPart <> 0 Then
            Bilgi = MsgBox("It has been detected that at least one of the report numbers has already been used. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet23Ek2
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
    'End If
YinedeKaydet23Ek2:
    If Rapor1TarihiText.Value = "" Then
        Bilgi = MsgBox("It has been detected that the report date was not specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("It has been detected that at least one of the result, production feature and/or report feature rows exceeds the number of item type rows. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Rapor1Kont = 1
            GoTo YinedeKaydet25
        ElseIf Bilgi = vbNo Then
            Rapor1Kont = 2
            GoTo SonRapor
        End If
    End If
YinedeKaydet25:
    'Check report date for chronological correctness
    If Rapor1TarihiText.Value <> "" And TutanakTarihiText.Value <> "" Then
        If Year(Rapor1TarihiText.Value) < Year(TeslimTarihiText.Value) Or _
            Year(Rapor1TarihiText.Value) < Year(TutanakTarihiText.Value) Then
            Bilgi = MsgBox("It has been detected that the report date is earlier than the statement date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet25Ek1
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
YinedeKaydet25Ek1:
        If (Year(Rapor1TarihiText.Value) = Year(TutanakTarihiText.Value) And Month(Rapor1TarihiText.Value) < Month(TutanakTarihiText.Value)) Then
            Bilgi = MsgBox("It has been detected that the report date is earlier than the statement date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Rapor1Kont = 1
                GoTo YinedeKaydet25Ek2
            ElseIf Bilgi = vbNo Then
                Rapor1Kont = 2
                GoTo SonRapor
            End If
        End If
YinedeKaydet25Ek2:
        If (Year(Rapor1TarihiText.Value) = Year(TutanakTarihiText.Value) And Month(Rapor1TarihiText.Value) = Month(TutanakTarihiText.Value) And Day(Rapor1TarihiText.Value) < Day(TutanakTarihiText.Value)) Then
            Bilgi = MsgBox("It has been detected that the report date is earlier than the statement date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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


'Tutanak2 kontrolleri
'Tutanak2Kont = 2
If Tutanak2Frame.Visible = True Then
    Tutanak2Kont = 0
    If TemaNoText.Value = "" Then
        Bilgi = MsgBox("It has been detected that the theme number was not entered. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet26Ek1
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet26Ek1:
    If Tutanak2TarihiText.Value = "" Then
        Bilgi = MsgBox("It has been detected that the Statement 2 date was not specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet26
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet26:
    If GidenPaketTipi.Value = "" Then
        Bilgi = MsgBox("It has been detected that the type of the outgoing package was not specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("It has been detected that the quantity of the outgoing package was not specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            Tutanak2Kont = 1
            GoTo YinedeKaydet29
        ElseIf Bilgi = vbNo Then
            Tutanak2Kont = 2
            GoTo SonTutanak2
        End If
    End If
YinedeKaydet29:
    'Chronological validation for Statement 2 date
    If Tutanak2TarihiText.Value <> "" And TutanakTarihiText.Value <> "" And Rapor1TarihiText.Value <> "" Then
        If Year(Tutanak2TarihiText.Value) < Year(TutanakTarihiText.Value) Or Year(Tutanak2TarihiText.Value) < Year(Rapor1TarihiText.Value) Then
            Bilgi = MsgBox("It has been detected that the Statement 2 date is earlier than the statement and/or report date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet29Ek1
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
YinedeKaydet29Ek1:
        If (Year(Tutanak2TarihiText.Value) = Year(TutanakTarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(TutanakTarihiText.Value)) Or _
            (Year(Tutanak2TarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(Tutanak2TarihiText.Value) < Month(Rapor1TarihiText.Value)) Then
            Bilgi = MsgBox("It has been detected that the Statement 2 date is earlier than the statement and/or report date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                Tutanak2Kont = 1
                GoTo YinedeKaydet29Ek2
            ElseIf Bilgi = vbNo Then
                Tutanak2Kont = 2
                GoTo SonTutanak2
            End If
        End If
YinedeKaydet29Ek2:
        If (Year(Tutanak2TarihiText.Value) = Year(TutanakTarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(TutanakTarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(TutanakTarihiText.Value)) Or _
            (Year(Tutanak2TarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(Tutanak2TarihiText.Value) = Month(Rapor1TarihiText.Value) And Day(Tutanak2TarihiText.Value) < Day(Rapor1TarihiText.Value)) Then
            Bilgi = MsgBox("It has been detected that the Statement 2 date is earlier than the statement and/or report date. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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

'FinansalBirim üst yazı kontrolleri
'FinansalBirimUstYaziKont = 2
If FinansalBirimUstYaziFrame.Visible = True Then
    FinansalBirimUstYaziKont = 0
    If FinansalBirimUstYaziTarihiText.Value = "" Then
        Bilgi = MsgBox("The cover letter date for the Financial Unit has not been specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek1
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet30Ek1:
    If FinansalBirimUstYaziNoText.Value = "" Then
        Bilgi = MsgBox("The cover letter number for the Financial Unit has not been specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet31Ek1
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet31Ek1:
    If IlFinansalBirim.Value = "" Then
        Bilgi = MsgBox("It has been detected that the province was not specified in the Financial Unit cover letter section. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek2
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet30Ek2:
    If FinansalBirimGonderilen.Value = "" Then
        Bilgi = MsgBox("It has been detected that the Receiving Unit for the Financial Unit was not specified. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek21
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet30Ek21:
    If Adres.Value = "" Then
        Bilgi = MsgBox("It has been detected that the address was not specified in the Financial Unit cover letter section. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek3
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet30Ek3:
    If FinansalBirimGonderiTipi.Value = "" Then
        Bilgi = MsgBox("It has been detected that the delivery type was not specified in the Financial Unit cover letter section. To save it this way, click " & """" & "Yes" & """" & "; to make corrections, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek4
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
YinedeKaydet30Ek4:

    If TipAOption.Value = True And DesteBandiAdedi.Value = "" Then
        Bilgi = MsgBox("It has been detected that the number of labels is not specified in the Financial Unit cover letter section. To save the record anyway, click " & """" & "Yes" & """" & "; to correct the input, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet30Ek5
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
    
YinedeKaydet30Ek5:
    If FinansalBirimUstYaziTarihiText.Value <> "" And Tutanak2TarihiText.Value <> "" And TutanakTarihiText.Value <> "" And Rapor1TarihiText.Value <> "" Then
        If Year(FinansalBirimUstYaziTarihiText.Value) < Year(TutanakTarihiText.Value) Or _
           Year(FinansalBirimUstYaziTarihiText.Value) < Year(Rapor1TarihiText.Value) Or _
           Year(FinansalBirimUstYaziTarihiText.Value) < Year(Tutanak2TarihiText.Value) Then
            Bilgi = MsgBox("The date of the Financial Unit cover letter appears to be earlier than the dates of Statement, Report and/or Statement 2. To save the record anyway, click " & """" & "Yes" & """" & "; to correct the input, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                FinansalBirimUstYaziKont = 1
                GoTo YinedeKaydet32Ek1
            ElseIf Bilgi = vbNo Then
                FinansalBirimUstYaziKont = 2
                GoTo SonFinansalBirimUstYazi
            End If
        End If
    
YinedeKaydet32Ek1:
    If (Year(FinansalBirimUstYaziTarihiText.Value) = Year(TutanakTarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) < Month(TutanakTarihiText.Value)) Or _
       (Year(FinansalBirimUstYaziTarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) < Month(Rapor1TarihiText.Value)) Or _
       (Year(FinansalBirimUstYaziTarihiText.Value) = Year(Tutanak2TarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) < Month(Tutanak2TarihiText.Value)) Then
        Bilgi = MsgBox("The date of the Financial Unit cover letter appears to be earlier than the dates of Statement 1, Report 1 and/or Statement 2. To save the record anyway, click " & """" & "Yes" & """" & "; to correct the input, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet33Ek1
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
    
YinedeKaydet33Ek1:
    If (Year(FinansalBirimUstYaziTarihiText.Value) = Year(TutanakTarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) = Month(TutanakTarihiText.Value) And Day(FinansalBirimUstYaziTarihiText.Value) < Day(TutanakTarihiText.Value)) Or _
       (Year(FinansalBirimUstYaziTarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) = Month(Rapor1TarihiText.Value) And Day(FinansalBirimUstYaziTarihiText.Value) < Day(Rapor1TarihiText.Value)) Or _
       (Year(FinansalBirimUstYaziTarihiText.Value) = Year(Tutanak2TarihiText.Value) And Month(FinansalBirimUstYaziTarihiText.Value) = Month(Tutanak2TarihiText.Value) And Day(FinansalBirimUstYaziTarihiText.Value) < Day(Tutanak2TarihiText.Value)) Then
        Bilgi = MsgBox("The date of the Financial Unit cover letter appears to be earlier than the dates of Statement, Report and/or Statement 2. To save the record anyway, click " & """" & "Yes" & """" & "; to correct the input, click " & """" & "No" & """" & ".", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            FinansalBirimUstYaziKont = 1
            GoTo YinedeKaydet34Ek1
        ElseIf Bilgi = vbNo Then
            FinansalBirimUstYaziKont = 2
            GoTo SonFinansalBirimUstYazi
        End If
    End If
    
YinedeKaydet34Ek1:
    End If
End If

'Üst yazı kontrolleri
'UstYaziKont = 2
If UstYaziFrame.Visible = True Then
    UstYaziKont = 0
    If UstYaziTarihiText.Value = "" Then
        Bilgi = MsgBox("The date of the cover letter has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
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
        Bilgi = MsgBox("The cover letter number has not been specified. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
        If Bilgi = vbYes Then
            UstYaziKont = 1
            GoTo YinedeKaydet31
        ElseIf Bilgi = vbNo Then
            UstYaziKont = 2
            GoTo SonUstYazi
        End If
    End If
YinedeKaydet31:
    ' Check whether the cover letter date is earlier than other dates
    If UstYaziTarihiText.Value <> "" And FinansalBirimUstYaziTarihiText.Value <> "" And Tutanak2TarihiText.Value <> "" And TutanakTarihiText.Value <> "" And Rapor1TarihiText.Value <> "" Then
        If Year(UstYaziTarihiText.Value) < Year(TutanakTarihiText.Value) Or _
            Year(UstYaziTarihiText.Value) < Year(Rapor1TarihiText.Value) Or _
            Year(UstYaziTarihiText.Value) < Year(Tutanak2TarihiText.Value) Or _
            Year(UstYaziTarihiText.Value) < Year(FinansalBirimUstYaziTarihiText.Value) Then
            Bilgi = MsgBox("The date of the cover letter appears to be earlier than the dates of Statement, Report, Statement 2 and/or Financial Unit cover letter. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                UstYaziKont = 1
                GoTo YinedeKaydet32
            ElseIf Bilgi = vbNo Then
                UstYaziKont = 2
                GoTo SonUstYazi
            End If
        End If
YinedeKaydet32:
        If (Year(UstYaziTarihiText.Value) = Year(TutanakTarihiText.Value) And Month(UstYaziTarihiText.Value) < Month(TutanakTarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(UstYaziTarihiText.Value) < Month(Rapor1TarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(Tutanak2TarihiText.Value) And Month(UstYaziTarihiText.Value) < Month(Tutanak2TarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(FinansalBirimUstYaziTarihiText.Value) And Month(UstYaziTarihiText.Value) < Month(FinansalBirimUstYaziTarihiText.Value)) Then
            Bilgi = MsgBox("The date of the cover letter appears to be earlier than the dates of Statement, Report, Statement 2 and/or Financial Unit cover letter. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                UstYaziKont = 1
                GoTo YinedeKaydet33
            ElseIf Bilgi = vbNo Then
                UstYaziKont = 2
                GoTo SonUstYazi
            End If
        End If
YinedeKaydet33:
        If (Year(UstYaziTarihiText.Value) = Year(TutanakTarihiText.Value) And Month(UstYaziTarihiText.Value) = Month(TutanakTarihiText.Value) And Day(UstYaziTarihiText.Value) < Day(TutanakTarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(Rapor1TarihiText.Value) And Month(UstYaziTarihiText.Value) = Month(Rapor1TarihiText.Value) And Day(UstYaziTarihiText.Value) < Day(Rapor1TarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(Tutanak2TarihiText.Value) And Month(UstYaziTarihiText.Value) = Month(Tutanak2TarihiText.Value) And Day(UstYaziTarihiText.Value) < Day(Tutanak2TarihiText.Value)) Or _
            (Year(UstYaziTarihiText.Value) = Year(FinansalBirimUstYaziTarihiText.Value) And Month(UstYaziTarihiText.Value) = Month(FinansalBirimUstYaziTarihiText.Value) And Day(UstYaziTarihiText.Value) < Day(FinansalBirimUstYaziTarihiText.Value)) Then
            Bilgi = MsgBox("The date of the cover letter appears to be earlier than the dates of Statement, Report, Statement 2 and/or Financial Unit cover letter. Click 'Yes' to save anyway, or 'No' to make corrections.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                UstYaziKont = 1
                GoTo YinedeKaydet34
            ElseIf Bilgi = vbNo Then
                UstYaziKont = 2
                GoTo SonUstYazi
            End If
        End If
YinedeKaydet34:
    End If

'Bu kontrol kullanıcı inisiyatifine bırakılmamıştır. Otomatik olarak düzeltilip kullanıcı bilgilendirilir.
    If UstYaziNotuCheck.Value = True Then
        Kont = 0
        SonucKontrol = False
        If Controls("Sonuc").Value = "invalid" Then
            SonucKontrol = True
            GoTo Git
        End If
        For OgeFrame = 1 To 19
            If Controls("OgeTuruFrame" & OgeFrame).Visible = True Then
                Kont = OgeFrame
            End If
        Next OgeFrame
        If Kont > 0 Then
            For OgeFrame = 1 To Kont
                If Controls("Sonuc" & OgeFrame).Value = "invalid" Then
                    SonucKontrol = True
                    GoTo Git
                End If
            Next OgeFrame
        End If
Git:
        If SonucKontrol = False Then
            UstYaziNotuCheck.Value = False
            MsgBox "Since no invalid Type A has been detected in the operation (within the Result field(s)), the note added for the Directorate/Decision Board cover letter will be removed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        End If
    End If
End If

'Sağdaki ve soldaki tek boşluğu kaldır
'AdSoyad
Do While Left(AdSoyad.Value, 1) = " "
    AdSoyad.Value = Right(AdSoyad.Value, Len(AdSoyad.Value) - 1)
Loop
Do While Right(AdSoyad.Value, 1) = " "
    AdSoyad.Value = Left(AdSoyad.Value, Len(AdSoyad.Value) - 1)
Loop
'Birden fazla boşluk varsa kaldır
For i = 1 To 50
    AdSoyad.Value = Replace(AdSoyad.Value, "  ", " ")
Next i

'BabaAdi
Do While Left(BabaAdi.Value, 1) = " "
    BabaAdi.Value = Right(BabaAdi.Value, Len(BabaAdi.Value) - 1)
Loop
Do While Right(BabaAdi.Value, 1) = " "
    BabaAdi.Value = Left(BabaAdi.Value, Len(BabaAdi.Value) - 1)
Loop
'Birden fazla boşluk varsa kaldır
For i = 1 To 50
    BabaAdi.Value = Replace(BabaAdi.Value, "  ", " ")
Next i

'DogumYeri
Do While Left(DogumYeri.Value, 1) = " "
    DogumYeri.Value = Right(DogumYeri.Value, Len(DogumYeri.Value) - 1)
Loop
Do While Right(DogumYeri.Value, 1) = " "
    DogumYeri.Value = Left(DogumYeri.Value, Len(DogumYeri.Value) - 1)
Loop
'Birden fazla boşluk varsa kaldır
For i = 1 To 50
    DogumYeri.Value = Replace(DogumYeri.Value, "  ", " ")
Next i

'Adres
If FinansalBirimUstYaziFrame.Visible = True Then
    Do While Left(Adres.Value, 1) = " "
        Adres.Value = Right(Adres.Value, Len(Adres.Value) - 1)
    Loop
    Do While Right(Adres.Value, 1) = " "
        Adres.Value = Left(Adres.Value, Len(Adres.Value) - 1)
    Loop
    'Birden fazla boşluk varsa kaldır
    For i = 1 To 50
        Adres.Value = Replace(Adres.Value, "  ", " ")
    Next i
End If

'Karakterlerin ilk harfi büyük
AdSoyad.Value = WorksheetFunction.Proper(AdSoyad.Value)
BabaAdi.Value = WorksheetFunction.Proper(BabaAdi.Value)
DogumYeri.Value = WorksheetFunction.Proper(DogumYeri.Value)
If FinansalBirimUstYaziFrame.Visible = True Then
    Adres.Value = WorksheetFunction.Proper(Adres.Value)
End If


Son:
SonRapor:
SonTutanak2:
SonFinansalBirimUstYazi:
SonUstYazi:

End Sub


Private Sub Kaydet_Click()
Dim YeniIslem As Long ', SiraBul As Range, SiraKontrol As Range
Dim i As Long, j As Long, OgeFrame As Integer, Kont As Integer
Dim ctl As MSForms.Control
'Dim TumKont As Integer, TutKont As Integer, Rapor1Kont As Integer, Tutanak2Kont As Integer, FinansalBirimUstYaziKont As Integer, UstYaziKont As Integer
Dim Bilgi As Variant
Dim OgeTuruKont As Integer, OgeDegeriKont As Integer, AdetKont As Integer
Dim OgeIdNoKont As Integer, AciklamaKont As Integer, SonucKont As Integer ', MaxiR As Integer, Maxi As Integer
Dim OgeTuruKontSatir As Integer, OgeDegeriKontSatir As Integer, AdetKontSatir As Integer
Dim OgeIdNoKontSatir As Integer, AciklamaKontSatir As Integer, SonucKontSatir As Integer
Dim Say As Long, IlkSira As Long, SonSira As Long, IlkSiraBul As Range, SonSiraBul As Range, Fark As Long
Dim FarkSay As Integer, SiraNoSakla As Long, SiraSay As Long
Dim Kenarlar As Range, DokumKontSatir As Integer, UserName As String

Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim OpenControl As String, Say1IslemGunlugu As Long, Say2IslemGunlugu As Long
Dim GelenTema As String, Sene As String, Ay As String
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long
Dim Olay1 As String, Olay2 As String, Olay3 As String, Olay4 As String, Olay5 As String
Dim ItemBul As Range

Dim Rapor1TarihBul As Range, Rapor1NoBulIlk As Range, RefSatir As Long
Dim UretimOzelligiKont As Integer, RaporOzelligiKont As Integer, Rapor1NoBulTireKont As Integer, Rapor1NoBulKont As Integer, Rapor1NoBulTireKontPart As Integer, Rapor1NoKont As Integer
Dim UretimOzelligiKontSatir As Integer, RaporOzelligiKontSatir As Integer, Rapor1NoKontAyni As Integer, Rapor1NoKontAltNoHata As Integer, Rapor1NoKontUstNoHata As Integer
Dim Rapor1NoBul As Range, Rapor1NoBulTire As Range, Rapor1NoBulTirePart As Range
'Dim StrRaporUnvan1 As String, StrRaporSicil1 As String, StrRaporUnvan2 As String, StrRaporSicil2 As String, StrRaporUnvan3 As String, StrRaporSicil3 As String


'Columns("FG:FH").EntireColumn.Hidden = False

ThisWorkbook.Activate
ThisWorkbook.Worksheets(5).Range("E6").Select

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(5).Unprotect Password:="123"
ThisWorkbook.Worksheets(10).Unprotect Password:="123"

UserName = Environ("UserProfile")
UserName = UCase(Right(UserName, 7)) 'WorksheetFunction.Proper(Right(UserName, 7)) 'UCase(Replace(Replace(Mid(Right(UserName, 7), 4, 2), "i", "I"), "ı", "I"))

TutKont = 3
Rapor1Kont = 3
Tutanak2Kont = 3
FinansalBirimUstYaziKont = 3
UstYaziKont = 3
YeniIslem = 0
'___________________

'Sıra numarası bulunamazsa prosedürden çık (Bu kısım zorunlu değildir. Esas bölüm düzeltme ksımındadır.)
'Kullanıcının sıra numarası vermesi engellenniş olur.

'__________Rapor No Senkronizasyon 30.11.2021

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
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



'Tüm bölümler için ön kontrol
TumKont = 0
For Each ctl In core_report3_2_entry_UI.TutanakFrame.Controls 'TutanakFrame
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
For Each ctl In core_report3_2_entry_UI.ScrollFrame.Controls 'ScrollFrame
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
For Each ctl In core_report3_2_entry_UI.Rapor1Frame.Controls 'Rapor1Frame
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
For Each ctl In core_report3_2_entry_UI.Tutanak2Frame.Controls 'Tutanak2Frame
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
For Each ctl In core_report3_2_entry_UI.FinansalBirimUstYaziFrame.Controls 'FinansalBirimUstYaziFrame
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
For Each ctl In core_report3_2_entry_UI.UstYaziFrame.Controls 'UstYaziFrame
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

If TumKont = 0 Then
    'MsgBox "All boş."
    'TutKont = 2
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

If TutKont = 2 Then
    GoTo Son
End If
If Rapor1Kont = 2 Then
    GoTo SonRapor
End If
If Tutanak2Kont = 2 Then
    GoTo SonTutanak2
End If
If FinansalBirimUstYaziKont = 2 Then
    GoTo SonFinansalBirimUstYazi
End If
If UstYaziKont = 2 Then
    GoTo SonUstYazi
End If

'______________



'DÜZELTME KAYDI
If ComboGetir.Value <> "" Then
    'Veri tabanını kontrol et
    Say = Range("FG100000").End(xlUp).Row
    If Say < 7 Then
        GoTo ResetAtla
    End If

    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
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
    If Range("L" & IlkSira).Value = "Point2" Or Range("L" & IlkSira).Value = "Point3" Then
        Range("F" & IlkSira & ":EZ" & SonSira).ClearContents
        Range("GB" & IlkSira & ":HT" & SonSira).ClearContents 'En sondaki sayfa sayıları da hariç
    End If

    'Tutanak bölümü
    If TipAOption.Value = True Then
        Cells(IlkSira, 28).Value = "Type A"
    ElseIf TipBOption.Value = True Then
        Cells(IlkSira, 28).Value = "Type B"
    End If
    Cells(IlkSira, 19).Value = Il.Value
    Cells(IlkSira, 20).Value = Ilce.Value
    Cells(IlkSira, 23).Value = TutanakTarihiText.Value
    Cells(IlkSira, 24).Value = KayitNoText.Value
    Cells(IlkSira, 25).Value = TemaTipi.Value
    Cells(IlkSira, 26).Value = TemaNoText.Value
    If OtomatikOption.Value = True Then
        Cells(IlkSira, 27).Value = "Otomatik"
    ElseIf ManuelOption.Value = True Then
        Cells(IlkSira, 27).Value = "Manuel"
    End If
    Cells(IlkSira, 30).Value = FinansalBirim.Value
    Cells(IlkSira, 31).Value = TeslimatBirim.Value
    Cells(IlkSira, 32).Value = BildirilenBirim.Value
    Cells(IlkSira, 35).Value = TeslimTarihiText.Value
    Cells(IlkSira, 36).Value = AdSoyad.Value
    Cells(IlkSira, 38).Value = TCNo.Value
    Cells(IlkSira, 39).Value = BabaAdi.Value
    Cells(IlkSira, 40).Value = DogumYeri.Value
    Cells(IlkSira, 41).Value = DogumTarihiText.Value
    Cells(IlkSira, 44).Value = TelNo.Value
    Cells(IlkSira, 47).Value = MuhatapTemasi.Value
    If GonderilenBirim.Value = "Contact Theme" Then
        Cells(IlkSira, 48).Value = ""
    Else
        Cells(IlkSira, 48).Value = GonderilenBirim.Value
    End If
    Cells(IlkSira, 49).Value = DesteBandiFotokopisi.Value
    Cells(IlkSira, 50).Value = TeslimTutFotokopisi.Value
    Cells(IlkSira, 51).Value = BarkodNo.Value

    'Tutanak imzaları
    Cells(IlkSira, 184).Value = TutanakImza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo TutanakImza1DuzeltmeIslemAtla
    End If
    Cells(IlkSira, 185).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(IlkSira, 186).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza1DuzeltmeIslemAtla:
    
    Cells(IlkSira, 187).Value = TutanakImza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo TutanakImza2DuzeltmeIslemAtla
    End If
    Cells(IlkSira, 188).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(IlkSira, 189).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza2DuzeltmeIslemAtla:

    Cells(IlkSira, 190).Value = TutanakImza3.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza3.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo TutanakImza3DuzeltmeIslemAtla
    End If
    Cells(IlkSira, 191).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(IlkSira, 192).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza3DuzeltmeIslemAtla:
''''''''''''tutanak imza sonu

    Cells(IlkSira, 52).Value = OgeTuru.Value
    Cells(IlkSira, 55).Value = OgeDegeri.Value
    Cells(IlkSira, 58).Value = Adet.Value
    Cells(IlkSira, 61).Value = OgeIdNo.Value
    Cells(IlkSira, 64).Value = Aciklama.Value
    
    'Call TutanakGirisi_Click

    'Rapor1
    If Rapor1Frame.Visible = True Then
        Cells(IlkSira, 212).Value = Sonuc.Value
        Cells(IlkSira, 213).Value = UretimOzelligi.Value
        Cells(IlkSira, 214).Value = RaporOzelligi.Value
        If NotCheck.Value = True Then
            Cells(IlkSira, 216).Value = "Yes"
        Else
            Cells(IlkSira, 216).Value = "No"
        End If
        Cells(IlkSira, 217).Value = Rapor1No.Value
        Cells(IlkSira, 13).Value = Rapor1No.Value
        Cells(IlkSira, 218).Value = Rapor1TarihiText.Value
        'İmzalar (hazırlık)
        StrRaporUnvan1 = ""
        StrRaporSicil1 = ""
        StrRaporUnvan2 = ""
        StrRaporSicil2 = ""
        StrRaporUnvan3 = ""
        StrRaporSicil3 = ""
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
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza3.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo RaporImza3DuzeltmeIslemAtla
        End If
        StrRaporUnvan3 = Worksheets(2).Range("DZ" & ItemBul.Row)
        StrRaporSicil3 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza3DuzeltmeIslemAtla:
    '''''''''''İmzalar (hazırlık) sonu
    End If
 
    'Tutanak1 ve rapor1
    Maxi = Application.Max(Maxi, MaxiR)
    Fark = SonSira - IlkSira '+ 1
    MaxiAktar = Maxi
    FarkAktar = Fark
    If Maxi = Fark Then 'Sayfadaki satır sayısını değiştirme
        If Maxi > 0 And Maxi < 20 Then
            For OgeFrame = 1 To Maxi
                Cells(IlkSira + OgeFrame, 52).Value = Controls("OgeTuru" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 55).Value = Controls("OgeDegeri" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 58).Value = Controls("Adet" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 61).Value = Controls("OgeIdNo" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 64).Value = Controls("Aciklama" & OgeFrame).Value
                
                'Rapor1
                If Rapor1Frame.Visible = True Then
                    Cells(IlkSira + OgeFrame, 212).Value = Controls("Sonuc" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 213).Value = Controls("UretimOzelligi" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 214).Value = Controls("RaporOzelligi" & OgeFrame).Value
                    If Controls("Rapor1No" & OgeFrame).Value <> "" And Controls("NotCheck" & OgeFrame).Value = True Then
                        Cells(IlkSira + OgeFrame, 216).Value = "Yes"
                    ElseIf Controls("Rapor1No" & OgeFrame).Value <> "" And Controls("NotCheck" & OgeFrame).Value = False Then
                        Cells(IlkSira + OgeFrame, 216).Value = "No"
                    End If
                    Cells(IlkSira + OgeFrame, 217).Value = Controls("Rapor1No" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 13).Value = Controls("Rapor1No" & OgeFrame).Value
                    'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                    If Cells(IlkSira + OgeFrame, 13).Value = "" Then
                        Cells(IlkSira + OgeFrame, 174).Value = ""
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
            Cells(SonSira + FarkSay, 164).Value = Cells(SonSira, 164).Value
            Cells(SonSira, 164).Value = ""
            For OgeFrame = 1 To Maxi
                Cells(IlkSira + OgeFrame, 52).Value = Controls("OgeTuru" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 55).Value = Controls("OgeDegeri" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 58).Value = Controls("Adet" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 61).Value = Controls("OgeIdNo" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 64).Value = Controls("Aciklama" & OgeFrame).Value
                'Rapor1
                If Rapor1Frame.Visible = True Then
                    Cells(IlkSira + OgeFrame, 212).Value = Controls("Sonuc" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 213).Value = Controls("UretimOzelligi" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 214).Value = Controls("RaporOzelligi" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 217).Value = Controls("Rapor1No" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 13).Value = Controls("Rapor1No" & OgeFrame).Value
                    'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                    If Cells(IlkSira + OgeFrame, 13).Value = "" Then
                        Cells(IlkSira + OgeFrame, 174).Value = ""
                    End If
                End If
            Next OgeFrame
        End If
    ElseIf Maxi < Fark Then 'Sayfadan satır sil
        FarkSay = 0
        SiraNoSakla = Cells(SonSira, 164).Value
        For i = 1 To Fark - Maxi
            FarkSay = FarkSay + 1
            Rows(SonSira - (FarkSay - 1)).EntireRow.Delete 'Shift:=xlDown
        Next i
        Cells(SonSira - FarkSay, 164).Value = SiraNoSakla 'Cells(SonSira, 164).Value
        If Maxi > 0 Then
            For OgeFrame = 1 To Maxi
                Cells(IlkSira + OgeFrame, 52).Value = Controls("OgeTuru" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 55).Value = Controls("OgeDegeri" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 58).Value = Controls("Adet" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 61).Value = Controls("OgeIdNo" & OgeFrame).Value
                Cells(IlkSira + OgeFrame, 64).Value = Controls("Aciklama" & OgeFrame).Value
                'Rapor1
                If Rapor1Frame.Visible = True Then
                    Cells(IlkSira + OgeFrame, 212).Value = Controls("Sonuc" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 213).Value = Controls("UretimOzelligi" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 214).Value = Controls("RaporOzelligi" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 217).Value = Controls("Rapor1No" & OgeFrame).Value
                    Cells(IlkSira + OgeFrame, 13).Value = Controls("Rapor1No" & OgeFrame).Value
                    'Raporda düzeltme yapıldığında sayfa sayılarını tekrar düzenle
                    If Cells(IlkSira + OgeFrame, 13).Value = "" Then
                        Cells(IlkSira + OgeFrame, 174).Value = ""
                    End If
                End If
            Next OgeFrame
        End If
    End If
    
    
'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
'
'    Call ModuleReport3.IslemGunluguRapor3FinansalBirimDuzeltme
'
'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
'
'    ThisWorkbook.Activate
    
    'Tutanak2
    If Tutanak2Frame.Visible = True Then
        'Call Tutanak2Girisi_Click
        Cells(IlkSira, 69).Value = Tutanak2TarihiText.Value
        Cells(IlkSira, 71).Value = GidenPaketTipi.Value
        Cells(IlkSira, 72).Value = GidenPaketAdedi.Value
        
        'Tutanak2 imzaları
        Cells(IlkSira, 193).Value = Tutanak2Imza1.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo Tutanak2Imza1DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 194).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 195).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza1DuzeltmeIslemAtla:
        
        Cells(IlkSira, 196).Value = Tutanak2Imza2.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo Tutanak2Imza2DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 197).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 198).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza2DuzeltmeIslemAtla:
    ''''''''''''tutanak2 imza sonu
    End If
    'FinansalBirim üst yazı
    If FinansalBirimUstYaziFrame.Visible = True Then
        'Call FinansalBirimUstYaziGirisi_Click
        Cells(IlkSira, 75).Value = FinansalBirimUstYaziTarihiText.Value
        Cells(IlkSira, 76).Value = FinansalBirimUstYaziNoText.Value
        Cells(IlkSira, 77).Value = IlFinansalBirim.Value
        Cells(IlkSira, 78).Value = IlceFinansalBirim.Value
        If FinansalBirimGonderilen.Value = "FinansalBirim Teması" Then
            Cells(IlkSira, 82).Value = ""
        Else
            Cells(IlkSira, 82).Value = FinansalBirimGonderilen.Value
        End If
        Cells(IlkSira, 79).Value = Adres.Value
        Cells(IlkSira, 80).Value = DekontSayfa.Value
        Cells(IlkSira, 85).Value = FinansalBirimGonderiTipi.Value
        Cells(IlkSira, 81).Value = DesteBandiAdedi.Value
        'Bank üst yazı imzaları
        Cells(IlkSira, 199).Value = FinansalBirimUstYaziImza1.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=FinansalBirimUstYaziImza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo FinansalBirimUstYaziImza1DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 200).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 201).Value = Worksheets(2).Range("EA" & ItemBul.Row)
FinansalBirimUstYaziImza1DuzeltmeIslemAtla:
        
        Cells(IlkSira, 202).Value = FinansalBirimUstYaziImza2.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=FinansalBirimUstYaziImza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo FinansalBirimUstYaziImza2DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 203).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 204).Value = Worksheets(2).Range("EA" & ItemBul.Row)
FinansalBirimUstYaziImza2DuzeltmeIslemAtla:
    ''''''''''''FinansalBirim üst yazı imza sonu
    End If
    'Üst yazı
    If UstYaziFrame.Visible = True Then
        'Call UstYaziGirisi_Click
        Cells(IlkSira, 83).Value = UstYaziTarihiText.Value
        Cells(IlkSira, 84).Value = UstYaziNoText.Value

        If UstYaziNotuCheck.Value = True Then
            Cells(IlkSira, 215).Value = "Yes"
        Else
            Cells(IlkSira, 215).Value = "No"
        End If
            
        'Üst yazı imzaları
        Cells(IlkSira, 205).Value = UstYaziImza1.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza1.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo UstYaziImza1DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 206).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 207).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza1DuzeltmeIslemAtla:
        
        Cells(IlkSira, 208).Value = UstYaziImza2.Value
        Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza2.Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not ItemBul Is Nothing Then
            '
        Else
            GoTo UstYaziImza2DuzeltmeIslemAtla
        End If
        Cells(IlkSira, 209).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
        Cells(IlkSira, 210).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza2DuzeltmeIslemAtla:
    ''''''''''''üst yazı imza sonu
    End If

    'işlem günlüğü için zaman damgası 'ESKİ VERİLER İÇİN ZAMAN DAMGASI OLUŞTUR
    If Len(Cells(IlkSira, 165).Value) < 12 Then
        StrTime = Format(Now, "ddmmyyyyhhmmss")
        Cells(IlkSira, 165).Value = StrTime
    End If

    '__________Rapor No Senkronizasyon 30.11.2021

    Set WsRaporNo = ThisWorkbook.Worksheets(10)

    Set RnoIlkSiraBul = WsRaporNo.Range("D6:D100000").Find(What:=Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
    Set RnoSonSiraBul = WsRaporNo.Range("E6:E100000").Find(What:=Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
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
        WsRaporNo.Range(WsRaporNo.Cells(ilkrow, 1), WsRaporNo.Cells(sonrow, 1)).Value = Range(Cells(IlkSira, 13), Cells(IlkSira + Maxi, 13)).Value 'Rapor no
        WsRaporNo.Cells(ilkrow, 2).Value = Cells(IlkSira, 218).Value
        WsRaporNo.Cells(ilkrow, 3).Value = "Notification"
        WsRaporNo.Cells(ilkrow, 4).Value = Cells(IlkSira, 165).Value 'İlk zaman damgası
        WsRaporNo.Cells(sonrow, 5).Value = Cells(IlkSira, 165).Value 'Son zaman damgası

    End If

    '__________Rapor No Senkronizasyon 30.11.2021
    
    
    
'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
'
'    Call ModuleReport3.IslemGunluguRapor3FinansalBirimDuzeltme
'
'    '________________________İŞLEM GÜNLÜĞÜ DÜZELTME KAYDI
    
    ThisWorkbook.Activate
    
    'Prosedür sonu düzeltmeleri
    YeniIslem = IlkSira
    GoTo YeniIslemiAtla

End If

YinedeKaydet:


'YENİ İŞLEM
YeniIslem = Range("FH100000").End(xlUp).Row
If YeniIslem < 7 Then
    YeniIslem = 7
    GoTo IlkIslem
End If
YeniIslem = YeniIslem + 1
IlkIslem:

'__________Rapor No Senkronizasyon 30.11.2021

Set WsRaporNo = ThisWorkbook.Worksheets(10)
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

'Verileri Rapor3 Rapor3_1 formundan sayfaya aktar.
'Tutanak1 bölümü
If YeniIslem = 7 Then
    Cells(YeniIslem, 5).Value = 1 'İlk sıra numarasını ver
Else
    Cells(YeniIslem, 5).Value = Cells(YeniIslem - 1, 164).Value + 1 'Sıra numarası ver
End If

'Tutanak bölümü
If TipAOption.Value = True Then
    Cells(YeniIslem, 28).Value = "Type A"
ElseIf TipBOption.Value = True Then
    Cells(YeniIslem, 28).Value = "Type B"
End If
Cells(YeniIslem, 19).Value = Il.Value
Cells(YeniIslem, 20).Value = Ilce.Value
Cells(YeniIslem, 23).Value = TutanakTarihiText.Value
Cells(YeniIslem, 24).Value = KayitNoText.Value
Cells(YeniIslem, 25).Value = TemaTipi.Value
Cells(YeniIslem, 26).Value = TemaNoText.Value
If OtomatikOption.Value = True Then
    Cells(YeniIslem, 27).Value = "Otomatik"
ElseIf ManuelOption.Value = True Then
    Cells(YeniIslem, 27).Value = "Manuel"
End If
Cells(YeniIslem, 30).Value = FinansalBirim.Value
Cells(YeniIslem, 31).Value = TeslimatBirim.Value
Cells(YeniIslem, 32).Value = BildirilenBirim.Value
Cells(YeniIslem, 35).Value = TeslimTarihiText.Value
Cells(YeniIslem, 36).Value = AdSoyad.Value
Cells(YeniIslem, 38).Value = TCNo.Value
Cells(YeniIslem, 39).Value = BabaAdi.Value
Cells(YeniIslem, 40).Value = DogumYeri.Value
Cells(YeniIslem, 41).Value = DogumTarihiText.Value
Cells(YeniIslem, 44).Value = TelNo.Value
Cells(YeniIslem, 47).Value = MuhatapTemasi.Value
If GonderilenBirim.Value = "Contact Theme" Then
    Cells(YeniIslem, 48).Value = ""
Else
    Cells(YeniIslem, 48).Value = GonderilenBirim.Value
End If
Cells(YeniIslem, 49).Value = DesteBandiFotokopisi.Value
Cells(YeniIslem, 50).Value = TeslimTutFotokopisi.Value
Cells(YeniIslem, 51).Value = BarkodNo.Value


'Tutanak imzaları
Cells(YeniIslem, 184).Value = TutanakImza1.Value
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo TutanakImza1YeniIslemAtla
End If
Cells(YeniIslem, 185).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
Cells(YeniIslem, 186).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza1YeniIslemAtla:

Cells(YeniIslem, 187).Value = TutanakImza2.Value
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo TutanakImza2YeniIslemAtla
End If
Cells(YeniIslem, 188).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
Cells(YeniIslem, 189).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza2YeniIslemAtla:

Cells(YeniIslem, 190).Value = TutanakImza3.Value
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=TutanakImza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo TutanakImza3YeniIslemAtla
End If
Cells(YeniIslem, 191).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
Cells(YeniIslem, 192).Value = Worksheets(2).Range("EA" & ItemBul.Row)
TutanakImza3YeniIslemAtla:
''''''''''''tutanak imza sonu


Cells(YeniIslem, 52).Value = OgeTuru.Value
Cells(YeniIslem, 55).Value = OgeDegeri.Value
Cells(YeniIslem, 58).Value = Adet.Value
Cells(YeniIslem, 61).Value = OgeIdNo.Value
Cells(YeniIslem, 64).Value = Aciklama.Value

If Maxi > 0 Then
For OgeFrame = 1 To Maxi
    Cells(YeniIslem + OgeFrame, 52).Value = Controls("OgeTuru" & OgeFrame).Value
    Cells(YeniIslem + OgeFrame, 55).Value = Controls("OgeDegeri" & OgeFrame).Value
    Cells(YeniIslem + OgeFrame, 58).Value = Controls("Adet" & OgeFrame).Value
    Cells(YeniIslem + OgeFrame, 61).Value = Controls("OgeIdNo" & OgeFrame).Value
    Cells(YeniIslem + OgeFrame, 64).Value = Controls("Aciklama" & OgeFrame).Value
Next OgeFrame
End If


'Rapor bölümü
If Rapor1Frame.Visible = True Then
    Cells(YeniIslem, 212).Value = Sonuc.Value
    Cells(YeniIslem, 213).Value = UretimOzelligi.Value
    Cells(YeniIslem, 214).Value = RaporOzelligi.Value
    If NotCheck.Value = True Then
        Cells(YeniIslem, 216).Value = "Yes"
    Else
        Cells(YeniIslem, 216).Value = "No"
    End If
    Cells(YeniIslem, 217).Value = Rapor1No.Value
    Cells(YeniIslem, 13).Value = Rapor1No.Value
    Cells(YeniIslem, 218).Value = Rapor1TarihiText.Value
    'İmzalar (hazırlık)
    StrRaporUnvan1 = ""
    StrRaporSicil1 = ""
    StrRaporUnvan2 = ""
    StrRaporSicil2 = ""
    StrRaporUnvan3 = ""
    StrRaporSicil3 = ""
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
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=RaporImza3.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo RaporImza3YeniIslemAtla
    End If
    StrRaporUnvan3 = Worksheets(2).Range("DZ" & ItemBul.Row)
    StrRaporSicil3 = Worksheets(2).Range("EA" & ItemBul.Row)
RaporImza3YeniIslemAtla:
'''''''''''İmzalar (hazırlık) sonu

    If Maxi > 0 Then
        For OgeFrame = 1 To Maxi
            Cells(YeniIslem + OgeFrame, 217).Value = Controls("Rapor1No" & OgeFrame).Value
            Cells(YeniIslem + OgeFrame, 13).Value = Controls("Rapor1No" & OgeFrame).Value
            Cells(YeniIslem + OgeFrame, 212).Value = Controls("Sonuc" & OgeFrame).Value
            Cells(YeniIslem + OgeFrame, 213).Value = Controls("UretimOzelligi" & OgeFrame).Value
            If Controls("Rapor1No" & OgeFrame).Value <> "" And Controls("NotCheck" & OgeFrame).Value = True Then
                Cells(YeniIslem + OgeFrame, 216).Value = "Yes"
            ElseIf Controls("Rapor1No" & OgeFrame).Value <> "" And Controls("NotCheck" & OgeFrame).Value = False Then
                Cells(YeniIslem + OgeFrame, 216).Value = "No"
            End If
            Cells(YeniIslem + OgeFrame, 214).Value = Controls("RaporOzelligi" & OgeFrame).Value
        Next OgeFrame
    End If
End If


'''________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT
''
'Call ModuleReport3.IslemGunluguRapor3FinansalBirimYeni
''
'''________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT


'Tutanak2
If Tutanak2Frame.Visible = True Then
    Cells(YeniIslem, 69).Value = Tutanak2TarihiText.Value
    Cells(YeniIslem, 71).Value = GidenPaketTipi.Value
    Cells(YeniIslem, 72).Value = GidenPaketAdedi.Value

    'Tutanak2 imzaları
    Cells(YeniIslem, 193).Value = Tutanak2Imza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak2Imza1YeniIslemAtla
    End If
    Cells(YeniIslem, 194).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 195).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza1YeniIslemAtla:
    
    Cells(YeniIslem, 196).Value = Tutanak2Imza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tutanak2Imza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Tutanak2Imza2YeniIslemAtla
    End If
    Cells(YeniIslem, 197).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 198).Value = Worksheets(2).Range("EA" & ItemBul.Row)
Tutanak2Imza2YeniIslemAtla:
''''''''''''tutanak2 imza sonu
End If
'FinansalBirim üst yazı
If FinansalBirimUstYaziFrame.Visible = True Then
    Cells(YeniIslem, 75).Value = FinansalBirimUstYaziTarihiText.Value
    Cells(YeniIslem, 76).Value = FinansalBirimUstYaziNoText.Value
    Cells(YeniIslem, 77).Value = IlFinansalBirim.Value
    Cells(YeniIslem, 78).Value = IlceFinansalBirim.Value
    If FinansalBirimGonderilen.Value = "FinansalBirim Teması" Then
        Cells(YeniIslem, 82).Value = ""
    Else
        Cells(YeniIslem, 82).Value = FinansalBirimGonderilen.Value
    End If
    Cells(YeniIslem, 79).Value = Adres.Value
    Cells(YeniIslem, 80).Value = DekontSayfa.Value
    Cells(YeniIslem, 85).Value = FinansalBirimGonderiTipi.Value
    Cells(YeniIslem, 81).Value = DesteBandiAdedi.Value

    'Bank üst yazı imzaları
    Cells(YeniIslem, 199).Value = FinansalBirimUstYaziImza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=FinansalBirimUstYaziImza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo FinansalBirimUstYaziImza1YeniIslemAtla
    End If
    Cells(YeniIslem, 200).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 201).Value = Worksheets(2).Range("EA" & ItemBul.Row)
FinansalBirimUstYaziImza1YeniIslemAtla:
    
    Cells(YeniIslem, 202).Value = FinansalBirimUstYaziImza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=FinansalBirimUstYaziImza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo FinansalBirimUstYaziImza2YeniIslemAtla
    End If
    Cells(YeniIslem, 203).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 204).Value = Worksheets(2).Range("EA" & ItemBul.Row)
FinansalBirimUstYaziImza2YeniIslemAtla:
''''''''''''FinansalBirim üst yazı imza sonu
End If
'Üst yazı
If UstYaziFrame.Visible = True Then
    Cells(YeniIslem, 83).Value = UstYaziTarihiText.Value
    Cells(YeniIslem, 84).Value = UstYaziNoText.Value

    If UstYaziNotuCheck.Value = True Then
        Cells(YeniIslem, 215).Value = "Yes"
    Else
        Cells(YeniIslem, 215).Value = "No"
    End If
        
    'Üst yazı imzaları
    Cells(YeniIslem, 205).Value = UstYaziImza1.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza1.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo UstYaziImza1YeniIslemAtla
    End If
    Cells(YeniIslem, 206).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 207).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza1YeniIslemAtla:
    
    Cells(YeniIslem, 208).Value = UstYaziImza2.Value
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=UstYaziImza2.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo UstYaziImza2YeniIslemAtla
    End If
    Cells(YeniIslem, 209).Value = Worksheets(2).Range("DZ" & ItemBul.Row)
    Cells(YeniIslem, 210).Value = Worksheets(2).Range("EA" & ItemBul.Row)
UstYaziImza2YeniIslemAtla:
''''''''''''üst yazı imza sonu
End If
 
'işlem günlüğü için zaman damgası
StrTime = Format(Now, "ddmmyyyyhhmmss")
Cells(YeniIslem, 165).Value = StrTime

'İlk ve son satırları işaretle
Cells(YeniIslem, 163).Value = Cells(YeniIslem, 5).Value
Cells(YeniIslem + Maxi, 164).Value = Cells(YeniIslem, 5).Value


'__________Rapor No Senkronizasyon 30.11.2021

WsRaporNo.Range(WsRaporNo.Cells(islemNew, 1), WsRaporNo.Cells(islemNew + Maxi, 1)).Value = Range(Cells(YeniIslem, 13), Cells(YeniIslem + Maxi, 13)).Value 'Rapor no
WsRaporNo.Cells(islemNew, 2).Value = Cells(YeniIslem, 218).Value
WsRaporNo.Cells(islemNew, 3).Value = "Notification"
WsRaporNo.Cells(islemNew, 4).Value = Cells(YeniIslem, 165).Value 'İlk zaman damgası
WsRaporNo.Cells(islemNew + Maxi, 5).Value = Cells(YeniIslem, 165).Value 'Son zaman damgası

'__________Rapor No Senkronizasyon 30.11.2021


ThisWorkbook.Activate

'MsgBox Kont & ". satır görünür."

'TesteGit:

YeniIslemiAtla:

''________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT
'
Call ModuleReport3.IslemGunluguRapor3_2
'
''________________________İŞLEM GÜNLÜĞÜ YENİ KAYIT

ThisWorkbook.Activate

'Rapor1 için imzalar (Ek bölüm) Hem Düzeltme hem Yeni İşlem için kodlar.
If Rapor1Frame.Visible = True Then
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
        For i = IlkSiraBul.Row To SonSiraBul.Row
            Cells(i, 220).Value = ""
            Cells(i, 221).Value = ""
            Cells(i, 222).Value = ""
            Cells(i, 223).Value = ""
            Cells(i, 224).Value = ""
            Cells(i, 225).Value = ""
            Cells(i, 226).Value = ""
            Cells(i, 227).Value = ""
            Cells(i, 228).Value = ""
            If Cells(i, 217).Value <> "" Then
                Cells(i, 220).Value = RaporImza1.Value
                Cells(i, 221).Value = StrRaporUnvan1
                Cells(i, 222).Value = StrRaporSicil1
                Cells(i, 223).Value = RaporImza2.Value
                Cells(i, 224).Value = StrRaporUnvan2
                Cells(i, 225).Value = StrRaporSicil2
                Cells(i, 226).Value = RaporImza3.Value
                Cells(i, 227).Value = StrRaporUnvan3
                Cells(i, 228).Value = StrRaporSicil3
            End If
        Next i
    End If
End If

LblDuzeltme.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
LblDuzeltme.ForeColor = RGB(30, 30, 30)

Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(YeniIslem, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                
'Satır renklendirme ve kenarlıklar.
Set Kenarlar = Range("E" & IlkSiraBul.Row & ":HT" & SonSiraBul.Row)
If Cells(YeniIslem, 5).Value Mod 2 = 0 Then
    Range("E" & IlkSiraBul.Row & ":HT" & SonSiraBul.Row).Interior.Color = RGB(201, 216, 230)
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
    Range("E" & IlkSiraBul.Row & ":HT" & SonSiraBul.Row).Interior.Color = RGB(174, 185, 194) 'RGB(180, 210, 240)
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

If TutKont = 0 Then
    
    If Nokta2Option.Value = True Then
        Cells(YeniIslem, 12).Value = "Point2"
    ElseIf Nokta3Option.Value = True Then
        Cells(YeniIslem, 12).Value = "Point3"
    Else
        Cells(YeniIslem, 12).Value = "Point1"
    End If
    
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
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf TutKont = 1 Then
    If Nokta2Option.Value = True Then
        Cells(YeniIslem, 12).Value = "Point2"
    ElseIf Nokta3Option.Value = True Then
        Cells(YeniIslem, 12).Value = "Point3"
    Else
        Cells(YeniIslem, 12).Value = "Point1"
    End If
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
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf TutKont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf TutKont = 3 Then
    GoTo ReseteGit
End If


SonRapor:
'MsgBox Rapor1Kont
If TipBOption.Value = True Then
    If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
        For i = IlkSiraBul.Row To SonSiraBul.Row
            Cells(i, 7).Value = ""
            Range("G" & i).Font.Color = RGB(60, 100, 180)
        Next i
    End If
Else
    If Rapor1Kont = 0 Then
        If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
            For i = IlkSiraBul.Row To SonSiraBul.Row
                Cells(i, 7).Value = ""
                Range("G" & i).Font.Color = RGB(60, 100, 180)
                If Cells(i, 217).Value <> "" Then
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
        If Cells(YeniIslem, 10).Value = "" Then
            Cells(YeniIslem, 10).Value = "?"
            Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
        End If
    ElseIf Rapor1Kont = 1 Then
        If IlkSiraBul.Row <> 0 And SonSiraBul.Row <> 0 Then
            For i = IlkSiraBul.Row To SonSiraBul.Row
                Cells(i, 7).Value = ""
                Range("G" & i).Font.Color = RGB(60, 100, 180)
                If Cells(i, 217).Value <> "" Then
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
        If Cells(YeniIslem, 10).Value = "" Then
            Cells(YeniIslem, 10).Value = "?"
            Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
        End If
    ElseIf Rapor1Kont = 2 Then
        'Hiçbir şey yapma
        GoTo ResetAtla
    ElseIf Rapor1Kont = 3 Then
        GoTo ReseteGit
    End If
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
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak2Kont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 8).Value = "x"
    Range("H" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 9).Value = "" Then
        Cells(YeniIslem, 9).Value = "?"
        Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf Tutanak2Kont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf Tutanak2Kont = 3 Then
    GoTo ReseteGit
End If

SonFinansalBirimUstYazi:
If FinansalBirimUstYaziKont = 0 Then
    'Normal kaydet
    Cells(YeniIslem, 9).Value = "ü"
    Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf FinansalBirimUstYaziKont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 9).Value = "x"
    Range("I" & YeniIslem).Font.Color = RGB(60, 100, 180)
    If Cells(YeniIslem, 10).Value = "" Then
        Cells(YeniIslem, 10).Value = "?"
        Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
    End If
ElseIf FinansalBirimUstYaziKont = 2 Then
    'Hiçbir şey yapma
    GoTo ResetAtla
ElseIf FinansalBirimUstYaziKont = 3 Then
    GoTo ReseteGit
End If

SonUstYazi:
If UstYaziKont = 0 Then
    'Normal kaydet
    Cells(YeniIslem, 10).Value = "ü"
    Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
ElseIf UstYaziKont = 1 Then
    'Sorunlu kaydet
    Cells(YeniIslem, 10).Value = "x"
    Range("J" & YeniIslem).Font.Color = RGB(60, 100, 180)
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

Call Rapor3FinansalBirimFormunuResetle
ComboGetir.Value = ""
ComboGetir.Clear

ResetAtla:

'Say = Range("E100000").End(xlUp).Row
'If Say < 7 Then
'    GoTo GetirBos
'Else
'    SiraSay = Range("E100000").End(xlUp)
'End If
''Getir liste değerleri
'For i = SiraSay To 1 Step -1
'    With ComboGetir
'        .AddItem (i)
'    End With
'Next i
'GetirBos:

ComboGetirAktar = ComboGetir.Value
Call ComboGetirReset
'30.09.2021, 11:54 Güncelleme
'Çünkü hata ayıklama mesajına Hayır dendiğinde combogetir değeri siliniyor ve kullanıcı yanlışlıkla yeni kayıt oluşturmuş oluyor.
'Yukarıdaki komut ipatl edilirse de combogetir kendisini güncellemiyor.
'Çözüm combogetir değerini sakla ve combogetirreset prosedüründen sonra tekrar ekle.
ComboGetir.Value = ComboGetirAktar


If TipBOption.Value = True Then
    'Tümünü oluşturu işaretle
    If YeniIslem <> 0 Then
        If Cells(YeniIslem, 6).Value = "ü" And Cells(YeniIslem, 8).Value = "ü" And Cells(YeniIslem, 9).Value = "ü" And Cells(YeniIslem, 10).Value = "ü" Then
            'Normal kaydet
            Cells(YeniIslem, 11).Value = "ü"
            Range("K" & YeniIslem).Font.Color = RGB(60, 100, 180)
            Cells(YeniIslem, 14).Value = UserName
        Else
            'Sorunlu kaydet
            Cells(YeniIslem, 11).Value = "?"
            Range("K" & YeniIslem).Font.Color = RGB(60, 100, 180)
            Cells(YeniIslem, 14).Value = UserName
        End If
    End If
Else
    'Tümünü oluşturu işaretle
    If YeniIslem <> 0 Then
        If Cells(YeniIslem, 6).Value = "ü" And Cells(YeniIslem, 7).Value = "ü" And Cells(YeniIslem, 8).Value = "ü" And Cells(YeniIslem, 9).Value = "ü" And Cells(YeniIslem, 10).Value = "ü" Then
            'Normal kaydet
            Cells(YeniIslem, 11).Value = "ü"
            Range("K" & YeniIslem).Font.Color = RGB(60, 100, 180)
            Cells(YeniIslem, 14).Value = UserName
        Else
            'Sorunlu kaydet
            Cells(YeniIslem, 11).Value = "?"
            Range("K" & YeniIslem).Font.Color = RGB(60, 100, 180)
            Cells(YeniIslem, 14).Value = UserName
        End If
    End If
End If

'ThisWorkbook.Save

Out:

'Columns("FG:FH").EntireColumn.Hidden = True

ThisWorkbook.Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(10).Protect Password:="123"
ThisWorkbook.Protect "123"

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

End Sub

Private Sub TutanakGirisi_Click()
Dim i As Integer
Dim ctl As MSForms.Control

ThisWorkbook.Activate


Sonuc.Visible = False
LblSonuc.Visible = False

Rapor1No.Visible = False
LblRapor1No.Visible = False
RaporOzelligi.Visible = False
LblRaporOzelligi.Visible = False
UretimOzelligi.Visible = False
LblUretimOzelligi.Visible = False
'Rapor2_2No.Visible = False
'LblRapor2_2No.Visible = False

LblSonucUst.Visible = False
LblRapor1NoUst.Visible = False
LblRapor1NoUst.Visible = False
LblRaporOzelligiUst.Visible = False
RaporOzelligiEkleKaldirLabel.Visible = False
LblUretimOzelligiUst.Visible = False
'LblRapor2_2NoUst.Visible = False
NotEkleKaldirLabel.Visible = False
LblNotUst.Visible = False
NotCheck.Visible = False

For i = 1 To 19
    Controls("Sonuc" & i).Visible = False
    Controls("LblSonuc" & i).Visible = False
    Controls("Rapor1No" & i).Visible = False
    Controls("LblRapor1No" & i).Visible = False
    Controls("RaporOzelligi" & i).Visible = False
    Controls("LblRaporOzelligi" & i).Visible = False
    Controls("UretimOzelligi" & i).Visible = False
    Controls("LblUretimOzelligi" & i).Visible = False
'    Controls("Rapor2_2No" & i).Visible = False
'    Controls("LblRapor2_2No" & i).Visible = False
    Controls("NotCheck" & i).Visible = False
Next i


EkleOge.Left = 518
KaldirOge.Left = 538

Rapor1Frame.Visible = False
Tutanak2Frame.Visible = False
FinansalBirimUstYaziFrame.Visible = False
UstYaziFrame.Visible = False

For Each ctl In core_report3_2_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ctl) = "Frame" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

'TutanakGirisi.BackColor = RGB(180, 210, 240)
'TutanakGirisi.ForeColor = RGB(30, 30, 30)
'
'If ComboGetir.Value <> "" Then
'    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
'    LblDuzeltme.ForeColor = RGB(30, 30, 30)
'End If


If EkranKontrol = True Then
    
    core_report3_2_entry_UI.ScrollTop = 0
    core_report3_2_entry_UI.ScrollHeight = 0
    core_report3_2_entry_UI.ScrollBars = fmScrollBarsNone

    'Formun görünümü
    AltMenuFrame.Top = 426 '462 '444 '299
    TasiyiciFrame.Height = 448 '486
    core_report3_2_entry_UI.Height = 518 '546 '556 '497 '352
    core_report3_2_entry_UI.Width = 1024
    
Else
    'Formun görünümü
    AltMenuFrame.Top = 426 '462 '444 '299
    TasiyiciFrame.Height = 448 '560 '486
    core_report3_2_entry_UI.Height = 518 '584 '556 '497 '352
End If

TutanakFrame.ZOrder msoBringToFront
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

If TipBOption.Value = True Then

    'Nothing

Else

    EkleOge.Left = 916
    KaldirOge.Left = 936
    
    LblSonucUst.Visible = True
    LblSonuc.Visible = True
    Sonuc.Visible = True
    
    LblRapor1NoUst.Visible = True
    LblRapor1No.Visible = True
    Rapor1No.Visible = True
    
    LblRaporOzelligiUst.Visible = True
    LblRaporOzelligi.Visible = True
    RaporOzelligiEkleKaldirLabel.Visible = True
    RaporOzelligi.Visible = True
    
    LblUretimOzelligiUst.Visible = True
    LblUretimOzelligi.Visible = True
    UretimOzelligi.Visible = True
    
    'LblRapor2_2NoUst.Visible = True
    'LblRapor2_2No.Visible = True
    'Rapor2_2No.Visible = True
    
    NotEkleKaldirLabel.Visible = True
    LblNotUst.Visible = True
    NotCheck.Visible = True
    'Notları rapor nolara göre göster
    For i = 1 To 19
        If Controls("Rapor1No" & i).Value <> "" Then
            Controls("NotCheck" & i).Visible = True
        End If
    Next i
    
    Rapor1Frame.Visible = True
    
    Call Son20RaporNo

End If

Tutanak2Frame.Visible = False
FinansalBirimUstYaziFrame.Visible = False
UstYaziFrame.Visible = False

For Each ctl In core_report3_2_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

''RaporlamaGirisi.BackColor = RGB(180, 210, 240)
''RaporlamaGirisi.ForeColor = RGB(30, 30, 30)
'
'If ComboGetir.Value <> "" Then
'    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
'    LblDuzeltme.ForeColor = RGB(30, 30, 30)
'End If

TutanakGirisi.BackColor = RGB(180, 210, 240)
TutanakGirisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If


If TipBOption.Value = True Then
    '
Else

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 426 + Rapor1Frame.Height + 6 '462 '444 '299
        TasiyiciFrame.Height = 448 + Rapor1Frame.Height + 6 '486
        core_report3_2_entry_UI.Height = 518 '+ Rapor1Frame.Height + 6 '546 '556 '497 '352
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 518 + Rapor1Frame.Height + 6 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 426 + Rapor1Frame.Height + 6
        TasiyiciFrame.Height = 448 + Rapor1Frame.Height + 6
        core_report3_2_entry_UI.Height = 518 + Rapor1Frame.Height + 6
    End If
    
    
    Rapor1Frame.ZOrder msoBringToFront
End If

End Sub


Private Sub Tutanak2Girisi_Click()
Dim ctl As MSForms.Control

Call RaporlamaGirisiPro '()_Click

Tutanak2Frame.Visible = True

FinansalBirimUstYaziFrame.Visible = False
UstYaziFrame.Visible = False

For Each ctl In core_report3_2_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ctl) = "Frame" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

Tutanak2Girisi.BackColor = RGB(180, 210, 240)
Tutanak2Girisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If


If TipBOption.Value = True Then
    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + 6 '462 '444 '299
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + 6 '486
        core_report3_2_entry_UI.Height = 518 '+ Tutanak2Frame.Height + 6 '546 '556 '497 '352
        Tutanak2Frame.Top = 420
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + 6 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + 6
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + 6
        core_report3_2_entry_UI.Height = 518 + Tutanak2Frame.Height + 6
        Tutanak2Frame.Top = 420
    End If
    
    Tutanak2Frame.ZOrder msoBringToFront
Else

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + 6 '462 '444 '299
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + 6 '486
        core_report3_2_entry_UI.Height = 578 '+ Tutanak2Frame.Height + 6 '546 '556 '497 '352
        Tutanak2Frame.Top = 480
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 578 + Tutanak2Frame.Height + 6 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + 6
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + 6
        core_report3_2_entry_UI.Height = 578 + Tutanak2Frame.Height + 6
        Tutanak2Frame.Top = 480
    End If
    
    Tutanak2Frame.ZOrder msoBringToFront

End If


End Sub

Private Sub FinansalBirimUstYaziGirisi_Click()
Dim ctl As MSForms.Control

Tutanak2Girisi_Click

FinansalBirimUstYaziFrame.Visible = True

UstYaziFrame.Visible = False

For Each ctl In core_report3_2_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ctl) = "Frame" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

FinansalBirimUstYaziGirisi.BackColor = RGB(180, 210, 240)
FinansalBirimUstYaziGirisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If


If TipBOption.Value = True Then
    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        core_report3_2_entry_UI.Height = 518 '+ Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        FinansalBirimUstYaziFrame.Top = 480
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        core_report3_2_entry_UI.Height = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        FinansalBirimUstYaziFrame.Top = 480
    End If
    
    FinansalBirimUstYaziFrame.ZOrder msoBringToFront
Else

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        core_report3_2_entry_UI.Height = 578 '+ Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        FinansalBirimUstYaziFrame.Top = 540
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 578 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        core_report3_2_entry_UI.Height = 578 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
        FinansalBirimUstYaziFrame.Top = 540
    End If
    
    FinansalBirimUstYaziFrame.ZOrder msoBringToFront
    
End If

End Sub

Private Sub UstYaziGirisi_Click()
Dim ctl As MSForms.Control

FinansalBirimUstYaziGirisi_Click

UstYaziFrame.Visible = True

For Each ctl In core_report3_2_entry_UI.UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ctl) = "Frame" Then
        ctl.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl

UstYaziGirisi.BackColor = RGB(180, 210, 240)
UstYaziGirisi.ForeColor = RGB(30, 30, 30)

If ComboGetir.Value <> "" Then
    LblDuzeltme.BackColor = RGB(180, 210, 240) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
End If

If TipBOption.Value = True Then
    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        core_report3_2_entry_UI.Height = 518 '+ Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        UstYaziFrame.Top = 570
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        core_report3_2_entry_UI.Height = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        UstYaziFrame.Top = 570
    End If
    
    UstYaziFrame.ZOrder msoBringToFront
Else

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        core_report3_2_entry_UI.Height = 578 '+ Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        UstYaziFrame.Top = 630
        
        core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report3_2_entry_UI.ScrollHeight = 578 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18 - 30
        core_report3_2_entry_UI.ScrollTop = 0
        core_report3_2_entry_UI.Width = 1024 + 12
        
    Else
        'Formun görünümü
        AltMenuFrame.Top = 486 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        TasiyiciFrame.Height = 508 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        core_report3_2_entry_UI.Height = 578 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
        UstYaziFrame.Top = 630
    End If
    
    UstYaziFrame.ZOrder msoBringToFront
End If

End Sub


Sub ColorChangerGenel()

'Düzeltme
If LblDuzeltme.BackColor <> RGB(180, 210, 240) Then
    If LblDuzeltme.BackColor <> RGB(225, 235, 245) Then
        LblDuzeltme.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
        LblDuzeltme.ForeColor = RGB(30, 30, 30)
    End If
End If
'Taslak
If LblTaslak.BackColor <> RGB(180, 210, 240) Then
    If LblTaslak.BackColor <> RGB(225, 235, 245) Then
        LblTaslak.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
        LblTaslak.ForeColor = RGB(30, 30, 30)
    End If
End If
'Sil
If LblSil.BackColor <> RGB(225, 235, 245) Then
    LblSil.BackColor = RGB(225, 235, 245)
    LblSil.ForeColor = RGB(30, 30, 30)
End If
'TipAOption
If TipAOption.BackColor <> RGB(225, 235, 245) Then
    TipAOption.BackColor = RGB(225, 235, 245)
    TipAOption.ForeColor = RGB(30, 30, 30)
End If
'TipBOption
If TipBOption.BackColor <> RGB(225, 235, 245) Then
    TipBOption.BackColor = RGB(225, 235, 245)
    TipBOption.ForeColor = RGB(30, 30, 30)
End If
'Nokta2Option
If Nokta2Option.BackColor <> RGB(225, 235, 245) Then
    Nokta2Option.BackColor = RGB(225, 235, 245)
    Nokta2Option.ForeColor = RGB(30, 30, 30)
End If
'Nokta3Option
If Nokta3Option.BackColor <> RGB(225, 235, 245) Then
    Nokta3Option.BackColor = RGB(225, 235, 245)
    Nokta3Option.ForeColor = RGB(30, 30, 30)
End If

'TutanakGirisi
If TutanakGirisi.BackColor <> RGB(180, 210, 240) Then
    If TutanakGirisi.BackColor <> RGB(225, 235, 245) Then
        TutanakGirisi.BackColor = RGB(225, 235, 245)
        TutanakGirisi.ForeColor = RGB(30, 30, 30)
    End If
End If
''RaporlamaGirisi
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
'FinansalBirimUstYaziGirisi
If FinansalBirimUstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    If FinansalBirimUstYaziGirisi.BackColor <> RGB(225, 235, 245) Then
        FinansalBirimUstYaziGirisi.BackColor = RGB(225, 235, 245)
        FinansalBirimUstYaziGirisi.ForeColor = RGB(30, 30, 30)
    End If
End If
'Üst yazı
If UstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    If UstYaziGirisi.BackColor <> RGB(225, 235, 245) Then
        UstYaziGirisi.BackColor = RGB(225, 235, 245)
        UstYaziGirisi.ForeColor = RGB(30, 30, 30)
    End If
End If
'Kapat
If Kapat.BackColor <> RGB(225, 235, 245) Then
    Kapat.BackColor = RGB(225, 235, 245)
    Kapat.ForeColor = RGB(30, 30, 30)
End If
'Yardim
If Yardim.BackColor <> RGB(225, 235, 245) Then
    Yardim.BackColor = RGB(225, 235, 245)
    Yardim.ForeColor = RGB(30, 30, 30)
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
'TutanakTarihiLabel
If TutanakTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    TutanakTarihiLabel.BackColor = RGB(254, 254, 254)
    TutanakTarihiLabel.ForeColor = RGB(70, 70, 70)
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
'FinansalBirimEkleKaldirLabel
If FinansalBirimEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    FinansalBirimEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    FinansalBirimEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'TeslimatBirimEkleKaldirLabel
If TeslimatBirimEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    TeslimatBirimEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    TeslimatBirimEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'BildirilenBirimEkleKaldirLabel
If BildirilenBirimEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    BildirilenBirimEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    BildirilenBirimEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'TeslimTarihiLabel
If TeslimTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    TeslimTarihiLabel.BackColor = RGB(254, 254, 254)
    TeslimTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'DogumTarihiLabel
If DogumTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    DogumTarihiLabel.BackColor = RGB(254, 254, 254)
    DogumTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'MuhatapTemasiEkleKaldirLabel
If MuhatapTemasiEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    MuhatapTemasiEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    MuhatapTemasiEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'GonderilenBirimEkleKaldirLabel
If GonderilenBirimEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    GonderilenBirimEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    GonderilenBirimEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'TutanakImza1EkleKaldirLabel
If TutanakImza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    TutanakImza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    TutanakImza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'TutanakImza2EkleKaldirLabel
If TutanakImza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    TutanakImza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    TutanakImza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'TutanakImza3EkleKaldirLabel
If TutanakImza3EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    TutanakImza3EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    TutanakImza3EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
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

'RaporOzelligiEkleKaldirLabel
If RaporOzelligiEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    RaporOzelligiEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    RaporOzelligiEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'NotEkleKaldirLabel
If NotEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    NotEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    NotEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
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
'RaporImza3EkleKaldirLabel
If RaporImza3EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    RaporImza3EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    RaporImza3EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If

'Tutanak2TarihiLabel
If Tutanak2TarihiLabel.BackColor <> RGB(254, 254, 254) Then
    Tutanak2TarihiLabel.BackColor = RGB(254, 254, 254)
    Tutanak2TarihiLabel.ForeColor = RGB(70, 70, 70)
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
'FinansalBirimUstYaziTarihiLabel
If FinansalBirimUstYaziTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    FinansalBirimUstYaziTarihiLabel.BackColor = RGB(254, 254, 254)
    FinansalBirimUstYaziTarihiLabel.ForeColor = RGB(70, 70, 70)
End If
'FinansalBirimGonderilenEkleKaldirLabel
If FinansalBirimGonderilenEkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    FinansalBirimGonderilenEkleKaldirLabel.BackColor = RGB(254, 254, 254)
    FinansalBirimGonderilenEkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'FinansalBirimUstYaziImza1EkleKaldirLabel
If FinansalBirimUstYaziImza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    FinansalBirimUstYaziImza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    FinansalBirimUstYaziImza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
'FinansalBirimUstYaziImza2EkleKaldirLabel
If FinansalBirimUstYaziImza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
    FinansalBirimUstYaziImza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
    FinansalBirimUstYaziImza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
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

'UstYaziNotuCheck
If UstYaziNotuCheck.BackColor <> RGB(254, 254, 254) Then
    UstYaziNotuCheck.BackColor = RGB(254, 254, 254)
    UstYaziNotuCheck.ForeColor = RGB(70, 70, 70)
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
Private Sub TipAOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TipAOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
TipAOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TipBOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TipBOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
TipBOption.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Nokta2Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Nokta2Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Nokta2Option.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Nokta3Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Nokta3Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Nokta3Option.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub TutanakGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If TutanakGirisi.BackColor <> RGB(180, 210, 240) Then
    TutanakGirisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    TutanakGirisi.ForeColor = RGB(255, 255, 255)
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
Private Sub FinansalBirimUstYaziGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If FinansalBirimUstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    FinansalBirimUstYaziGirisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    FinansalBirimUstYaziGirisi.ForeColor = RGB(255, 255, 255)
End If
End Sub
Private Sub UstYaziGirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If UstYaziGirisi.BackColor <> RGB(180, 210, 240) Then
    UstYaziGirisi.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    UstYaziGirisi.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub TutanakTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
TutanakTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TeslimTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TeslimTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
TeslimTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub DogumTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
DogumTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
DogumTarihiLabel.ForeColor = RGB(255, 255, 255)
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
Private Sub FinansalBirimUstYaziTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FinansalBirimUstYaziTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
FinansalBirimUstYaziTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub UstYaziTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
UstYaziTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
UstYaziTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub UstYaziNotuCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
UstYaziNotuCheck.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
UstYaziNotuCheck.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LblUstYaziNotu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub FinansalBirimEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FinansalBirimEkleKaldirLabel.BackColor = RGB(60, 100, 180)
FinansalBirimEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TeslimatBirimEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TeslimatBirimEkleKaldirLabel.BackColor = RGB(60, 100, 180)
TeslimatBirimEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub BildirilenBirimEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
BildirilenBirimEkleKaldirLabel.BackColor = RGB(60, 100, 180)
BildirilenBirimEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub MuhatapTemasiEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
MuhatapTemasiEkleKaldirLabel.BackColor = RGB(60, 100, 180)
MuhatapTemasiEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub GonderilenBirimEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
GonderilenBirimEkleKaldirLabel.BackColor = RGB(60, 100, 180)
GonderilenBirimEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
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

Private Sub RaporOzelligiEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporOzelligiEkleKaldirLabel.BackColor = RGB(60, 100, 180)
RaporOzelligiEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub NotEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
NotEkleKaldirLabel.BackColor = RGB(60, 100, 180)
NotEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub FinansalBirimGonderilenEkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FinansalBirimGonderilenEkleKaldirLabel.BackColor = RGB(60, 100, 180)
FinansalBirimGonderilenEkleKaldirLabel.ForeColor = RGB(255, 255, 255)
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
Private Sub LblUretimOzelligi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UretimOzelligi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(UretimOzelligi) 'Open scrollable with mouse
End Sub
Private Sub LblRapor1No_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Rapor1No_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblRaporOzelligi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub RaporOzelligi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(RaporOzelligi) 'Open scrollable with mouse
End Sub
Private Sub NotCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub UretimOzelligi1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi1) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi2) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi3) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi4) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi5) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi6) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi7) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi8) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi9) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi10) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi11) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi12) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi13) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi14) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi15) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi16) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi17) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi18) 'Open scrollable with mouse
End Sub
Private Sub UretimOzelligi19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(UretimOzelligi19) 'Open scrollable with mouse
End Sub

Private Sub RaporOzelligi1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi1) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi2) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi3) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi4) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi5) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi6) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi7) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi8) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi9) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi10) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi11) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi12) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi13) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi14) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi15) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi16) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi17) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi18) 'Open scrollable with mouse
End Sub
Private Sub RaporOzelligi19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(RaporOzelligi19) 'Open scrollable with mouse
End Sub


Private Sub BaslikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TutanakFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub FinansalBirimUstYaziFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TasiyiciFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub ComboGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboGetir) 'Open scrollable with mouse
End Sub
Private Sub TemaNoText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblTemaNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TeslimTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub DogumTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TutanakTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Rapor1TarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tutanak2TarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FinansalBirimUstYaziTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstYaziTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FinansalBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(FinansalBirim) 'Open scrollable with mouse
End Sub
Private Sub LblFinansalBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TeslimatBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(TeslimatBirim) 'Open scrollable with mouse
End Sub
Private Sub LblTeslimatBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub BildirilenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(BildirilenBirim) 'Open scrollable with mouse
End Sub
Private Sub LblBildirilenBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub MuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(MuhatapTemasi) 'Open scrollable with mouse
End Sub
Private Sub LblMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub LblNotUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblRaporOzelligiUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblRapor1NoUst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub FinansalBirimGonderilen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(FinansalBirimGonderilen) 'Open scrollable with mouse
End Sub
Private Sub LblFinansalBirimGonderilen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub TutanakImza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakImza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
TutanakImza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(TutanakImza1) 'Open scrollable with mouse
End Sub
Private Sub LblTutanakImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TutanakImza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakImza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
TutanakImza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(TutanakImza2) 'Open scrollable with mouse
End Sub
Private Sub LblTutanakImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TutanakImza3EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakImza3EkleKaldirLabel.BackColor = RGB(60, 100, 180)
TutanakImza3EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(TutanakImza3) 'Open scrollable with mouse
End Sub
Private Sub LblTutanakImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub RaporImza3EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporImza3EkleKaldirLabel.BackColor = RGB(60, 100, 180)
RaporImza3EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub RaporImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(RaporImza3) 'Open scrollable with mouse
End Sub
Private Sub LblRaporImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub FinansalBirimUstYaziImza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FinansalBirimUstYaziImza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
FinansalBirimUstYaziImza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub FinansalBirimUstYaziImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(FinansalBirimUstYaziImza1) 'Open scrollable with mouse
End Sub
Private Sub LblFinansalBirimUstYaziImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FinansalBirimUstYaziImza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FinansalBirimUstYaziImza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
FinansalBirimUstYaziImza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub FinansalBirimUstYaziImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(FinansalBirimUstYaziImza2) 'Open scrollable with mouse
End Sub
Private Sub LblFinansalBirimUstYaziImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub LblIl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIlce_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblKayitNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub IlFinansalBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(IlFinansalBirim) 'Open scrollable with mouse
End Sub
'IlceGiden
Private Sub IlceFinansalBirim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(IlceFinansalBirim) 'Open scrollable with mouse
End Sub

'İkinci bölüm
'TemaTipi
Private Sub TemaTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(TemaTipi) 'Open scrollable with mouse
End Sub
'DesteBandiFotokopisi
Private Sub DesteBandiFotokopisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(DesteBandiFotokopisi) 'Open scrollable with mouse
End Sub
'TeslimTutFotokopisi
Private Sub TeslimTutFotokopisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(TeslimTutFotokopisi) 'Open scrollable with mouse
End Sub
'GidenPaketTipi
Private Sub GidenPaketTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(GidenPaketTipi) 'Open scrollable with mouse
End Sub
'GidenPaketAdedi
Private Sub GidenPaketAdedi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(GidenPaketAdedi) 'Open scrollable with mouse
End Sub
'DekontSayfa
Private Sub DekontSayfa_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(DekontSayfa) 'Open scrollable with mouse
End Sub
'DesteBandiAdedi
Private Sub DesteBandiAdedi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call SetComboBoxHook(DesteBandiAdedi) 'Open scrollable with mouse
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
'            'Ilce
'            Set IlceBul = ThisWorkbook.Worksheets(2).Columns(IlEsleyicisi + 6).Find(What:=Ilce.Value, SearchDirection:=xlNext, _
'                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
'            If Not IlceBul Is Nothing Then
'                IlceDegeri = ThisWorkbook.Worksheets(2).Range("D" & IlceBul.Row)
'                If IlceDegeri < 10 Then
'                    IlceDegeri = 0 & IlceDegeri
'                End If
'            Else
'                IlceDegeri = ""
'            End If
            'If TemaTipi.Value = "Organization A" Then
                IlceDegeri = "00"
            'End If
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
            
            'kayıt no
            For i = 1 To 50
                KayitNoText.Value = Replace(KayitNoText.Value, " ", "")
            Next i
            EvrakNo = ""
            'Yıl ve evrak no
            If KayitNoText.Value <> "" And IsNumeric(KayitNoText.Value) = False Then
                MsgBox "The registration number contains non-numeric characters, so the Theme number cannot be generated automatically.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                TemaNoText.Value = ""
                GoTo Son
            Else
                Yil = ""
                Yil = Right(TutanakTarihiText, 2)
                'Belge numarası
                EvrakNo = ""
                EvrakNo = Right(KayitNoText.Value, 5)
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
            If IlDegeri <> "" And Makam <> "" And Yil <> "" And EvrakNo <> "" Then 'And IlceDegeri <> "" And Makam <> "" And Yil <> "" And EvrakNo <> "" Then
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

Private Sub ComboGetir_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ComboGetir.DropDown

End Sub

Private Sub ComboGetir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    'Enter
    If KeyCode = vbKeyReturn Then
        'GetirLabelDuzeltme_Click
    End If
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        '
    End If
    If KeyCode = vbKeyDown Then
        'Il.SetFocus
    End If
    'Sağa ve sola
    If KeyCode = vbKeyLeft Then
        '
    End If
    If KeyCode = vbKeyRight Then
        'Il.SetFocus
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
    LblDuzeltme.BackColor = RGB(225, 235, 245) 'RGB(60, 100, 180)
    LblDuzeltme.ForeColor = RGB(30, 30, 30)
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

Private Sub TutanakTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        TutanakTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        TutanakTarihiText.Value = ""
    End If

End Sub

Private Sub TutanakTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TutanakTarihiText.BackColor = RGB(255, 255, 255)
TutanakTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TutanakTarihiText_Change()
If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If
End Sub

Private Sub TutanakTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    TutanakTarihiText.Value = CalTarih
    TutanakTarihiText.Value = Format(TutanakTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

TutanakTarihiText.BackColor = RGB(255, 255, 255)
TutanakTarihiText.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub KayitNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
KayitNoText.BackColor = RGB(255, 255, 255)
KayitNoText.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub KayitNoText_Change()
KayitNoText.BackColor = RGB(255, 255, 255)
KayitNoText.ForeColor = RGB(30, 30, 30)

If OtomatikOption.Value = True Then
    'Tema güncelle
    Call OtomatikOption_Click
End If

End Sub
Private Sub KayitNoText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'
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

'

End Sub

Private Sub FinansalBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.FinansalBirim.DropDown

FinansalBirim.BackColor = RGB(255, 255, 255)
FinansalBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If FinansalBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirim.ListIndex = FinansalBirim.ListIndex - 1
            End If
            Me.FinansalBirim.DropDown
            
        Case 40 'Aşağı
            If FinansalBirim.ListIndex = FinansalBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirim.ListIndex = FinansalBirim.ListIndex + 1
            End If
            Me.FinansalBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub FinansalBirim_Change()

If FinansalBirim.ListIndex = -1 And FinansalBirim.Value <> "" Then
   FinansalBirim.Value = ""
   GoTo Son
End If

If FinansalBirim.Value <> "" Then
    FinansalBirim.SelStart = 0
    FinansalBirim.SelLength = Len(FinansalBirim.Value)
End If

Son:

FinansalBirim.DropDown
If FinansalBirim.BackColor = RGB(60, 100, 180) Then
    FinansalBirim.BackColor = RGB(255, 255, 255)
    FinansalBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub TeslimatBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TeslimatBirim.DropDown
TeslimatBirim.BackColor = RGB(255, 255, 255)
TeslimatBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TeslimatBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If TeslimatBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TeslimatBirim.ListIndex = TeslimatBirim.ListIndex - 1
            End If
            Me.TeslimatBirim.DropDown
            
        Case 40 'Aşağı
            If TeslimatBirim.ListIndex = TeslimatBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TeslimatBirim.ListIndex = TeslimatBirim.ListIndex + 1
            End If
            Me.TeslimatBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TeslimatBirim_Change()

If TeslimatBirim.ListIndex = -1 And TeslimatBirim.Value <> "" Then
   TeslimatBirim.Value = ""
   GoTo Son
End If

If TeslimatBirim.Value <> "" Then
    TeslimatBirim.SelStart = 0
    TeslimatBirim.SelLength = Len(TeslimatBirim.Value)
End If

Son:

TeslimatBirim.DropDown
If TeslimatBirim.BackColor = RGB(60, 100, 180) Then
TeslimatBirim.BackColor = RGB(255, 255, 255)
TeslimatBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub BildirilenBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.BildirilenBirim.DropDown
BildirilenBirim.BackColor = RGB(255, 255, 255)
BildirilenBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub BildirilenBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If BildirilenBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                BildirilenBirim.ListIndex = BildirilenBirim.ListIndex - 1
            End If
            Me.BildirilenBirim.DropDown
            
        Case 40 'Aşağı
            If BildirilenBirim.ListIndex = BildirilenBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                BildirilenBirim.ListIndex = BildirilenBirim.ListIndex + 1
            End If
            Me.BildirilenBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub BildirilenBirim_Change()

If BildirilenBirim.ListIndex = -1 And BildirilenBirim.Value <> "" Then
   BildirilenBirim.Value = ""
   GoTo Son
End If

If BildirilenBirim.Value <> "" Then
    BildirilenBirim.SelStart = 0
    BildirilenBirim.SelLength = Len(BildirilenBirim.Value)
End If

Son:

BildirilenBirim.DropDown
If BildirilenBirim.BackColor = RGB(60, 100, 180) Then
BildirilenBirim.BackColor = RGB(255, 255, 255)
BildirilenBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub TeslimTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        TeslimTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        TeslimTarihiText.Value = ""
    End If

End Sub

Private Sub TeslimTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TeslimTarihiText.BackColor = RGB(255, 255, 255)
TeslimTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TeslimTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    TeslimTarihiText.Value = CalTarih
    TeslimTarihiText.Value = Format(TeslimTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

TeslimTarihiText.BackColor = RGB(255, 255, 255)
TeslimTarihiText.ForeColor = RGB(30, 30, 30)

End Sub
Private Sub AdSoyad_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
AdSoyad.BackColor = RGB(255, 255, 255)
AdSoyad.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub AdSoyad_Change()
AdSoyad.BackColor = RGB(255, 255, 255)
AdSoyad.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub AdSoyad_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'
    
End Sub

Private Sub TCNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TCNo.BackColor = RGB(255, 255, 255)
TCNo.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub TCNo_Change()
TCNo.BackColor = RGB(255, 255, 255)
TCNo.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub TCNo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'
    
End Sub

Private Sub BabaAdi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
BabaAdi.BackColor = RGB(255, 255, 255)
BabaAdi.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub BabaAdi_Change()
BabaAdi.BackColor = RGB(255, 255, 255)
BabaAdi.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub BabaAdi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'
    
End Sub

Private Sub DogumYeri_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
DogumYeri.BackColor = RGB(255, 255, 255)
DogumYeri.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub DogumYeri_Change()
DogumYeri.BackColor = RGB(255, 255, 255)
DogumYeri.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub DogumYeri_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'
    
End Sub

Private Sub DogumTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        DogumTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        DogumTarihiText.Value = ""
    End If

End Sub

Private Sub DogumTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

DogumTarihiText.BackColor = RGB(255, 255, 255)
DogumTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub DogumTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    DogumTarihiText.Value = CalTarih
    DogumTarihiText.Value = Format(DogumTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

DogumTarihiText.BackColor = RGB(255, 255, 255)
DogumTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TelNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TelNo.BackColor = RGB(255, 255, 255)
TelNo.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub TelNo_Change()
TelNo.BackColor = RGB(255, 255, 255)
TelNo.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub MuhatapTemasi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.MuhatapTemasi.DropDown

MuhatapTemasi.BackColor = RGB(255, 255, 255)
MuhatapTemasi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub MuhatapTemasi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If MuhatapTemasi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                MuhatapTemasi.ListIndex = MuhatapTemasi.ListIndex - 1
            End If
            Me.MuhatapTemasi.DropDown
            
        Case 40 'Aşağı
            If MuhatapTemasi.ListIndex = MuhatapTemasi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                MuhatapTemasi.ListIndex = MuhatapTemasi.ListIndex + 1
            End If
            Me.MuhatapTemasi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub MuhatapTemasi_Change()

If MuhatapTemasi.ListIndex = -1 And MuhatapTemasi.Value <> "" Then
   MuhatapTemasi.Value = ""
   GoTo Son
End If

If MuhatapTemasi.Value <> "" Then
    MuhatapTemasi.SelStart = 0
    MuhatapTemasi.SelLength = Len(MuhatapTemasi.Value)
End If

Son:

MuhatapTemasi.DropDown
If MuhatapTemasi.BackColor = RGB(60, 100, 180) Then
MuhatapTemasi.BackColor = RGB(255, 255, 255)
MuhatapTemasi.ForeColor = RGB(30, 30, 30)
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

Private Sub TeslimTutFotokopisi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TeslimTutFotokopisi.DropDown
TeslimTutFotokopisi.BackColor = RGB(255, 255, 255)
TeslimTutFotokopisi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TeslimTutFotokopisi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If TeslimTutFotokopisi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TeslimTutFotokopisi.ListIndex = TeslimTutFotokopisi.ListIndex - 1
            End If
            Me.TeslimTutFotokopisi.DropDown
            
        Case 40 'Aşağı
            If TeslimTutFotokopisi.ListIndex = TeslimTutFotokopisi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TeslimTutFotokopisi.ListIndex = TeslimTutFotokopisi.ListIndex + 1
            End If
            Me.TeslimTutFotokopisi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TeslimTutFotokopisi_Change()

If TeslimTutFotokopisi.ListIndex = -1 And TeslimTutFotokopisi.Value <> "" Then
   TeslimTutFotokopisi.Value = ""
   GoTo Son
End If

If TeslimTutFotokopisi.Value <> "" Then
    TeslimTutFotokopisi.SelStart = 0
    TeslimTutFotokopisi.SelLength = Len(TeslimTutFotokopisi.Value)
End If

Son:

TeslimTutFotokopisi.DropDown
If TeslimTutFotokopisi.BackColor = RGB(60, 100, 180) Then
TeslimTutFotokopisi.BackColor = RGB(255, 255, 255)
TeslimTutFotokopisi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub DesteBandiFotokopisi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.DesteBandiFotokopisi.DropDown
DesteBandiFotokopisi.BackColor = RGB(255, 255, 255)
DesteBandiFotokopisi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub DesteBandiFotokopisi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If DesteBandiFotokopisi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DesteBandiFotokopisi.ListIndex = DesteBandiFotokopisi.ListIndex - 1
            End If
            Me.DesteBandiFotokopisi.DropDown
            
        Case 40 'Aşağı
            If DesteBandiFotokopisi.ListIndex = DesteBandiFotokopisi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DesteBandiFotokopisi.ListIndex = DesteBandiFotokopisi.ListIndex + 1
            End If
            Me.DesteBandiFotokopisi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub DesteBandiFotokopisi_Change()

If DesteBandiFotokopisi.ListIndex = -1 And DesteBandiFotokopisi.Value <> "" Then
   DesteBandiFotokopisi.Value = ""
   GoTo Son
End If

If DesteBandiFotokopisi.Value <> "" Then
    DesteBandiFotokopisi.SelStart = 0
    DesteBandiFotokopisi.SelLength = Len(DesteBandiFotokopisi.Value)
End If

Son:

DesteBandiFotokopisi.DropDown
If DesteBandiFotokopisi.BackColor = RGB(60, 100, 180) Then
DesteBandiFotokopisi.BackColor = RGB(255, 255, 255)
DesteBandiFotokopisi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub BarkodNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
BarkodNo.BackColor = RGB(255, 255, 255)
BarkodNo.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub BarkodNo_Change()
BarkodNo.BackColor = RGB(255, 255, 255)
BarkodNo.ForeColor = RGB(30, 30, 30)
End Sub


Private Sub TutanakImza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TutanakImza1.DropDown
TutanakImza1.BackColor = RGB(255, 255, 255)
TutanakImza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TutanakImza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If TutanakImza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza1.ListIndex = TutanakImza1.ListIndex - 1
            End If
            Me.TutanakImza1.DropDown
            
        Case 40 'Aşağı
            If TutanakImza1.ListIndex = TutanakImza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza1.ListIndex = TutanakImza1.ListIndex + 1
            End If
            Me.TutanakImza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TutanakImza1_Change()

If TutanakImza1.ListIndex = -1 And TutanakImza1.Value <> "" Then
   TutanakImza1.Value = ""
   GoTo Son
End If

If TutanakImza1.Value <> "" Then
    TutanakImza1.SelStart = 0
    TutanakImza1.SelLength = Len(TutanakImza1.Value)
End If


Son:

TutanakImza1.DropDown
If TutanakImza1.BackColor = RGB(60, 100, 180) Then
TutanakImza1.BackColor = RGB(255, 255, 255)
TutanakImza1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub TutanakImza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TutanakImza2.DropDown
TutanakImza2.BackColor = RGB(255, 255, 255)
TutanakImza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TutanakImza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If TutanakImza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza2.ListIndex = TutanakImza2.ListIndex - 1
            End If
            Me.TutanakImza2.DropDown
            
        Case 40 'Aşağı
            If TutanakImza2.ListIndex = TutanakImza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza2.ListIndex = TutanakImza2.ListIndex + 1
            End If
            Me.TutanakImza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TutanakImza2_Change()

If TutanakImza2.ListIndex = -1 And TutanakImza2.Value <> "" Then
   TutanakImza2.Value = ""
   GoTo Son
End If

If TutanakImza2.Value <> "" Then
    TutanakImza2.SelStart = 0
    TutanakImza2.SelLength = Len(TutanakImza2.Value)
End If


Son:

TutanakImza2.DropDown
If TutanakImza2.BackColor = RGB(60, 100, 180) Then
TutanakImza2.BackColor = RGB(255, 255, 255)
TutanakImza2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub TutanakImza3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.TutanakImza3.DropDown
TutanakImza3.BackColor = RGB(255, 255, 255)
TutanakImza3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub TutanakImza3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If TutanakImza3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza3.ListIndex = TutanakImza3.ListIndex - 1
            End If
            Me.TutanakImza3.DropDown
            
        Case 40 'Aşağı
            If TutanakImza3.ListIndex = TutanakImza3.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                TutanakImza3.ListIndex = TutanakImza3.ListIndex + 1
            End If
            Me.TutanakImza3.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub TutanakImza3_Change()

If TutanakImza3.ListIndex = -1 And TutanakImza3.Value <> "" Then
   TutanakImza3.Value = ""
   GoTo Son
End If

If TutanakImza3.Value <> "" Then
    TutanakImza3.SelStart = 0
    TutanakImza3.SelLength = Len(TutanakImza3.Value)
End If


Son:

TutanakImza3.DropDown
If TutanakImza3.BackColor = RGB(60, 100, 180) Then
TutanakImza3.BackColor = RGB(255, 255, 255)
TutanakImza3.ForeColor = RGB(30, 30, 30)
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



Private Sub FinansalBirimUstYaziImza1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.FinansalBirimUstYaziImza1.DropDown
FinansalBirimUstYaziImza1.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziImza1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimUstYaziImza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If FinansalBirimUstYaziImza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimUstYaziImza1.ListIndex = FinansalBirimUstYaziImza1.ListIndex - 1
            End If
            Me.FinansalBirimUstYaziImza1.DropDown
            
        Case 40 'Aşağı
            If FinansalBirimUstYaziImza1.ListIndex = FinansalBirimUstYaziImza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimUstYaziImza1.ListIndex = FinansalBirimUstYaziImza1.ListIndex + 1
            End If
            Me.FinansalBirimUstYaziImza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub FinansalBirimUstYaziImza1_Change()

If FinansalBirimUstYaziImza1.ListIndex = -1 And FinansalBirimUstYaziImza1.Value <> "" Then
   FinansalBirimUstYaziImza1.Value = ""
   GoTo Son
End If

If FinansalBirimUstYaziImza1.Value <> "" Then
    FinansalBirimUstYaziImza1.SelStart = 0
    FinansalBirimUstYaziImza1.SelLength = Len(FinansalBirimUstYaziImza1.Value)
End If


Son:

FinansalBirimUstYaziImza1.DropDown
If FinansalBirimUstYaziImza1.BackColor = RGB(60, 100, 180) Then
FinansalBirimUstYaziImza1.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziImza1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub FinansalBirimUstYaziImza2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.FinansalBirimUstYaziImza2.DropDown
FinansalBirimUstYaziImza2.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziImza2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimUstYaziImza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If FinansalBirimUstYaziImza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimUstYaziImza2.ListIndex = FinansalBirimUstYaziImza2.ListIndex - 1
            End If
            Me.FinansalBirimUstYaziImza2.DropDown
            
        Case 40 'Aşağı
            If FinansalBirimUstYaziImza2.ListIndex = FinansalBirimUstYaziImza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimUstYaziImza2.ListIndex = FinansalBirimUstYaziImza2.ListIndex + 1
            End If
            Me.FinansalBirimUstYaziImza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub FinansalBirimUstYaziImza2_Change()

If FinansalBirimUstYaziImza2.ListIndex = -1 And FinansalBirimUstYaziImza2.Value <> "" Then
   FinansalBirimUstYaziImza2.Value = ""
   GoTo Son
End If

If FinansalBirimUstYaziImza2.Value <> "" Then
    FinansalBirimUstYaziImza2.SelStart = 0
    FinansalBirimUstYaziImza2.SelLength = Len(FinansalBirimUstYaziImza2.Value)
End If


Son:

FinansalBirimUstYaziImza2.DropDown
If FinansalBirimUstYaziImza2.BackColor = RGB(60, 100, 180) Then
FinansalBirimUstYaziImza2.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziImza2.ForeColor = RGB(30, 30, 30)
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

Private Sub UstYaziNotuCheck_Click()
Dim Kont As Integer, OgeFrame As Integer, SonucKontrol As Boolean

If UstYaziNotuCheck.Value = True Then
    Kont = 0
    SonucKontrol = False
    If Controls("Sonuc").Value = "invalid" Then
        SonucKontrol = True
        GoTo Git
    End If
    For OgeFrame = 1 To 19
        If Controls("OgeTuruFrame" & OgeFrame).Visible = True Then
            Kont = OgeFrame
        End If
    Next OgeFrame
    If Kont > 0 Then
        For OgeFrame = 1 To Kont
            If Controls("Sonuc" & OgeFrame).Value = "invalid" Then
                SonucKontrol = True
                GoTo Git
            End If
        Next OgeFrame
    End If
Git:
    If SonucKontrol = False Then
        UstYaziNotuCheck.Value = False
        MsgBox "Since no invalid Type A was detected in the Result field(s), the note cannot be added to the Directorate/Decision Board cover letter.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
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
        OgeDegeri.SetFocus
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

If Sonuc.Value = "invalid" Then
    UretimOzelligi.Enabled = True
    'RaporOzelligi.Enabled = True
Else
    UretimOzelligi.Enabled = False
    'RaporOzelligi.Enabled = False
    UretimOzelligi.Value = ""
    'RaporOzelligi.Value = ""
End If

'Sonuc.DropDown

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

If Sonuc1.Value = "invalid" Then
    UretimOzelligi1.Enabled = True
    'RaporOzelligi1.Enabled = True
Else
    UretimOzelligi1.Enabled = False
    'RaporOzelligi1.Enabled = False
    UretimOzelligi1.Value = ""
    'RaporOzelligi1.Value = ""
End If

'Sonuc1.DropDown

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

If Sonuc2.Value = "invalid" Then
    UretimOzelligi2.Enabled = True
    'RaporOzelligi2.Enabled = True
Else
    UretimOzelligi2.Enabled = False
    'RaporOzelligi2.Enabled = False
    UretimOzelligi2.Value = ""
    'RaporOzelligi2.Value = ""
End If

'Sonuc2.DropDown

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

If Sonuc3.Value = "invalid" Then
    UretimOzelligi3.Enabled = True
    'RaporOzelligi3.Enabled = True
Else
    UretimOzelligi3.Enabled = False
    'RaporOzelligi3.Enabled = False
    UretimOzelligi3.Value = ""
    'RaporOzelligi3.Value = ""
End If

'Sonuc3.DropDown

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

If Sonuc4.Value = "invalid" Then
    UretimOzelligi4.Enabled = True
    'RaporOzelligi4.Enabled = True
Else
    UretimOzelligi4.Enabled = False
    'RaporOzelligi4.Enabled = False
    UretimOzelligi4.Value = ""
    'RaporOzelligi4.Value = ""
End If

'Sonuc4.DropDown

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

If Sonuc5.Value = "invalid" Then
    UretimOzelligi5.Enabled = True
    'RaporOzelligi5.Enabled = True
Else
    UretimOzelligi5.Enabled = False
    'RaporOzelligi5.Enabled = False
    UretimOzelligi5.Value = ""
    'RaporOzelligi5.Value = ""
End If

'Sonuc5.DropDown

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

If Sonuc6.Value = "invalid" Then
    UretimOzelligi6.Enabled = True
    'RaporOzelligi6.Enabled = True
Else
    UretimOzelligi6.Enabled = False
    'RaporOzelligi6.Enabled = False
    UretimOzelligi6.Value = ""
    'RaporOzelligi6.Value = ""
End If

'Sonuc6.DropDown

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

If Sonuc7.Value = "invalid" Then
    UretimOzelligi7.Enabled = True
    'RaporOzelligi7.Enabled = True
Else
    UretimOzelligi7.Enabled = False
    'RaporOzelligi7.Enabled = False
    UretimOzelligi7.Value = ""
    'RaporOzelligi7.Value = ""
End If

'Sonuc7.DropDown

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

If Sonuc8.Value = "invalid" Then
    UretimOzelligi8.Enabled = True
    'RaporOzelligi8.Enabled = True
Else
    UretimOzelligi8.Enabled = False
    'RaporOzelligi8.Enabled = False
    UretimOzelligi8.Value = ""
    'RaporOzelligi8.Value = ""
End If

'Sonuc8.DropDown

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

If Sonuc9.Value = "invalid" Then
    UretimOzelligi9.Enabled = True
    'RaporOzelligi9.Enabled = True
Else
    UretimOzelligi9.Enabled = False
    'RaporOzelligi9.Enabled = False
    UretimOzelligi9.Value = ""
    'RaporOzelligi9.Value = ""
End If

'Sonuc9.DropDown

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

If Sonuc10.Value = "invalid" Then
    UretimOzelligi10.Enabled = True
    'RaporOzelligi10.Enabled = True
Else
    UretimOzelligi10.Enabled = False
    'RaporOzelligi10.Enabled = False
    UretimOzelligi10.Value = ""
    'RaporOzelligi10.Value = ""
End If

'Sonuc10.DropDown

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

If Sonuc11.Value = "invalid" Then
    UretimOzelligi11.Enabled = True
    'RaporOzelligi11.Enabled = True
Else
    UretimOzelligi11.Enabled = False
    'RaporOzelligi11.Enabled = False
    UretimOzelligi11.Value = ""
    'RaporOzelligi11.Value = ""
End If

'Sonuc11.DropDown

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

If Sonuc12.Value = "invalid" Then
    UretimOzelligi12.Enabled = True
    'RaporOzelligi12.Enabled = True
Else
    UretimOzelligi12.Enabled = False
    'RaporOzelligi12.Enabled = False
    UretimOzelligi12.Value = ""
    'RaporOzelligi12.Value = ""
End If

'Sonuc12.DropDown

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

If Sonuc13.Value = "invalid" Then
    UretimOzelligi13.Enabled = True
    'RaporOzelligi13.Enabled = True
Else
    UretimOzelligi13.Enabled = False
    'RaporOzelligi13.Enabled = False
    UretimOzelligi13.Value = ""
    'RaporOzelligi13.Value = ""
End If

'Sonuc13.DropDown

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

If Sonuc14.Value = "invalid" Then
    UretimOzelligi14.Enabled = True
    'RaporOzelligi14.Enabled = True
Else
    UretimOzelligi14.Enabled = False
    'RaporOzelligi14.Enabled = False
    UretimOzelligi14.Value = ""
    'RaporOzelligi14.Value = ""
End If

'Sonuc14.DropDown

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

If Sonuc15.Value = "invalid" Then
    UretimOzelligi15.Enabled = True
    'RaporOzelligi15.Enabled = True
Else
    UretimOzelligi15.Enabled = False
    'RaporOzelligi15.Enabled = False
    UretimOzelligi15.Value = ""
    'RaporOzelligi15.Value = ""
End If

'Sonuc15.DropDown

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

If Sonuc16.Value = "invalid" Then
    UretimOzelligi16.Enabled = True
    'RaporOzelligi16.Enabled = True
Else
    UretimOzelligi16.Enabled = False
    'RaporOzelligi16.Enabled = False
    UretimOzelligi16.Value = ""
    'RaporOzelligi16.Value = ""
End If

'Sonuc16.DropDown

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

If Sonuc17.Value = "invalid" Then
    UretimOzelligi17.Enabled = True
    'RaporOzelligi17.Enabled = True
Else
    UretimOzelligi17.Enabled = False
    'RaporOzelligi17.Enabled = False
    UretimOzelligi17.Value = ""
    'RaporOzelligi17.Value = ""
End If

'Sonuc17.DropDown

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

If Sonuc18.Value = "invalid" Then
    UretimOzelligi18.Enabled = True
    'RaporOzelligi18.Enabled = True
Else
    UretimOzelligi18.Enabled = False
    'RaporOzelligi18.Enabled = False
    UretimOzelligi18.Value = ""
    'RaporOzelligi18.Value = ""
End If

'Sonuc18.DropDown

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
        'Sonuc19.SetFocus
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

If Sonuc19.Value = "invalid" Then
    UretimOzelligi19.Enabled = True
    'RaporOzelligi19.Enabled = True
Else
    UretimOzelligi19.Enabled = False
    'RaporOzelligi19.Enabled = False
    UretimOzelligi19.Value = ""
    'RaporOzelligi19.Value = ""
End If

'Sonuc19.DropDown

Son:

Sonuc19.DropDown
If Sonuc19.BackColor = RGB(60, 100, 180) Then
Sonuc19.BackColor = RGB(255, 255, 255)
Sonuc19.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi.DropDown
UretimOzelligi.BackColor = RGB(255, 255, 255)
UretimOzelligi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi1.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi.ListIndex = UretimOzelligi.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi.ListIndex = UretimOzelligi.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi_Change()

If UretimOzelligi.ListIndex = -1 And UretimOzelligi.Value <> "" Then
   UretimOzelligi.Value = ""
   GoTo Son
End If

If UretimOzelligi.Value <> "" Then
    UretimOzelligi.SelStart = 0
    UretimOzelligi.SelLength = Len(UretimOzelligi.Value)
End If

Son:

UretimOzelligi.DropDown
If UretimOzelligi.BackColor = RGB(60, 100, 180) Then
UretimOzelligi.BackColor = RGB(255, 255, 255)
UretimOzelligi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi1.DropDown
UretimOzelligi1.BackColor = RGB(255, 255, 255)
UretimOzelligi1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi2.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi1.ListIndex = UretimOzelligi1.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi1.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi1.ListIndex = UretimOzelligi1.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi1_Change()

If UretimOzelligi1.ListIndex = -1 And UretimOzelligi1.Value <> "" Then
   UretimOzelligi1.Value = ""
   GoTo Son
End If

If UretimOzelligi1.Value <> "" Then
    UretimOzelligi1.SelStart = 0
    UretimOzelligi1.SelLength = Len(UretimOzelligi1.Value)
End If

Son:

UretimOzelligi1.DropDown
If UretimOzelligi1.BackColor = RGB(60, 100, 180) Then
UretimOzelligi1.BackColor = RGB(255, 255, 255)
UretimOzelligi1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi2.DropDown
UretimOzelligi2.BackColor = RGB(255, 255, 255)
UretimOzelligi2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi3.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi2.ListIndex = UretimOzelligi2.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi2.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi2.ListIndex = UretimOzelligi2.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi2_Change()

If UretimOzelligi2.ListIndex = -1 And UretimOzelligi2.Value <> "" Then
   UretimOzelligi2.Value = ""
   GoTo Son
End If

If UretimOzelligi2.Value <> "" Then
    UretimOzelligi2.SelStart = 0
    UretimOzelligi2.SelLength = Len(UretimOzelligi2.Value)
End If

Son:

UretimOzelligi2.DropDown
If UretimOzelligi2.BackColor = RGB(60, 100, 180) Then
UretimOzelligi2.BackColor = RGB(255, 255, 255)
UretimOzelligi2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi3.DropDown
UretimOzelligi3.BackColor = RGB(255, 255, 255)
UretimOzelligi3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi4.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi3.ListIndex = UretimOzelligi3.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi3.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi3.ListIndex = UretimOzelligi3.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi3_Change()

If UretimOzelligi3.ListIndex = -1 And UretimOzelligi3.Value <> "" Then
   UretimOzelligi3.Value = ""
   GoTo Son
End If

If UretimOzelligi3.Value <> "" Then
    UretimOzelligi3.SelStart = 0
    UretimOzelligi3.SelLength = Len(UretimOzelligi3.Value)
End If

Son:

UretimOzelligi3.DropDown
If UretimOzelligi3.BackColor = RGB(60, 100, 180) Then
UretimOzelligi3.BackColor = RGB(255, 255, 255)
UretimOzelligi3.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi4.DropDown
UretimOzelligi4.BackColor = RGB(255, 255, 255)
UretimOzelligi4.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi5.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi4.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi4.ListIndex = UretimOzelligi4.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi4.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi4.ListIndex = UretimOzelligi4.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi4_Change()

If UretimOzelligi4.ListIndex = -1 And UretimOzelligi4.Value <> "" Then
   UretimOzelligi4.Value = ""
   GoTo Son
End If

If UretimOzelligi4.Value <> "" Then
    UretimOzelligi4.SelStart = 0
    UretimOzelligi4.SelLength = Len(UretimOzelligi4.Value)
End If

Son:

UretimOzelligi4.DropDown
If UretimOzelligi4.BackColor = RGB(60, 100, 180) Then
UretimOzelligi4.BackColor = RGB(255, 255, 255)
UretimOzelligi4.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi5.DropDown
UretimOzelligi5.BackColor = RGB(255, 255, 255)
UretimOzelligi5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi6.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi5.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi5.ListIndex = UretimOzelligi5.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi5.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi5.ListIndex = UretimOzelligi5.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi5_Change()

If UretimOzelligi5.ListIndex = -1 And UretimOzelligi5.Value <> "" Then
   UretimOzelligi5.Value = ""
   GoTo Son
End If

If UretimOzelligi5.Value <> "" Then
    UretimOzelligi5.SelStart = 0
    UretimOzelligi5.SelLength = Len(UretimOzelligi5.Value)
End If

Son:

UretimOzelligi5.DropDown
If UretimOzelligi5.BackColor = RGB(60, 100, 180) Then
UretimOzelligi5.BackColor = RGB(255, 255, 255)
UretimOzelligi5.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi6.DropDown
UretimOzelligi6.BackColor = RGB(255, 255, 255)
UretimOzelligi6.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi7.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi6.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi6.ListIndex = UretimOzelligi6.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi6.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi6.ListIndex = UretimOzelligi6.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi6_Change()

If UretimOzelligi6.ListIndex = -1 And UretimOzelligi6.Value <> "" Then
   UretimOzelligi6.Value = ""
   GoTo Son
End If

If UretimOzelligi6.Value <> "" Then
    UretimOzelligi6.SelStart = 0
    UretimOzelligi6.SelLength = Len(UretimOzelligi6.Value)
End If

Son:

UretimOzelligi6.DropDown
If UretimOzelligi6.BackColor = RGB(60, 100, 180) Then
UretimOzelligi6.BackColor = RGB(255, 255, 255)
UretimOzelligi6.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi7.DropDown
UretimOzelligi7.BackColor = RGB(255, 255, 255)
UretimOzelligi7.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi8.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi7.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi7.ListIndex = UretimOzelligi7.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi7.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi7.ListIndex = UretimOzelligi7.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi7_Change()

If UretimOzelligi7.ListIndex = -1 And UretimOzelligi7.Value <> "" Then
   UretimOzelligi7.Value = ""
   GoTo Son
End If

If UretimOzelligi7.Value <> "" Then
    UretimOzelligi7.SelStart = 0
    UretimOzelligi7.SelLength = Len(UretimOzelligi7.Value)
End If

Son:

UretimOzelligi7.DropDown
If UretimOzelligi7.BackColor = RGB(60, 100, 180) Then
UretimOzelligi7.BackColor = RGB(255, 255, 255)
UretimOzelligi7.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi8.DropDown
UretimOzelligi8.BackColor = RGB(255, 255, 255)
UretimOzelligi8.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi9.SetFocus
    End If
    
    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi8.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi8.ListIndex = UretimOzelligi8.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi8.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi8.ListIndex = UretimOzelligi8.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi8_Change()

If UretimOzelligi8.ListIndex = -1 And UretimOzelligi8.Value <> "" Then
   UretimOzelligi8.Value = ""
   GoTo Son
End If

If UretimOzelligi8.Value <> "" Then
    UretimOzelligi8.SelStart = 0
    UretimOzelligi8.SelLength = Len(UretimOzelligi8.Value)
End If

Son:

UretimOzelligi8.DropDown
If UretimOzelligi8.BackColor = RGB(60, 100, 180) Then
UretimOzelligi8.BackColor = RGB(255, 255, 255)
UretimOzelligi8.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi9.DropDown
UretimOzelligi9.BackColor = RGB(255, 255, 255)
UretimOzelligi9.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi10.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi9.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi9.ListIndex = UretimOzelligi9.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi9.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi9.ListIndex = UretimOzelligi9.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi9_Change()

If UretimOzelligi9.ListIndex = -1 And UretimOzelligi9.Value <> "" Then
   UretimOzelligi9.Value = ""
   GoTo Son
End If

If UretimOzelligi9.Value <> "" Then
    UretimOzelligi9.SelStart = 0
    UretimOzelligi9.SelLength = Len(UretimOzelligi9.Value)
End If

Son:

UretimOzelligi9.DropDown
If UretimOzelligi9.BackColor = RGB(60, 100, 180) Then
UretimOzelligi9.BackColor = RGB(255, 255, 255)
UretimOzelligi9.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi10.DropDown
UretimOzelligi10.BackColor = RGB(255, 255, 255)
UretimOzelligi10.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi11.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi10.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi10.ListIndex = UretimOzelligi10.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi10.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi10.ListIndex = UretimOzelligi10.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi10_Change()

If UretimOzelligi10.ListIndex = -1 And UretimOzelligi10.Value <> "" Then
   UretimOzelligi10.Value = ""
   GoTo Son
End If

If UretimOzelligi10.Value <> "" Then
    UretimOzelligi10.SelStart = 0
    UretimOzelligi10.SelLength = Len(UretimOzelligi10.Value)
End If

Son:

UretimOzelligi10.DropDown
If UretimOzelligi10.BackColor = RGB(60, 100, 180) Then
UretimOzelligi10.BackColor = RGB(255, 255, 255)
UretimOzelligi10.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi11.DropDown
UretimOzelligi11.BackColor = RGB(255, 255, 255)
UretimOzelligi11.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi12.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi11.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi11.ListIndex = UretimOzelligi11.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi11.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi11.ListIndex = UretimOzelligi11.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi11_Change()

If UretimOzelligi11.ListIndex = -1 And UretimOzelligi11.Value <> "" Then
   UretimOzelligi11.Value = ""
   GoTo Son
End If

If UretimOzelligi11.Value <> "" Then
    UretimOzelligi11.SelStart = 0
    UretimOzelligi11.SelLength = Len(UretimOzelligi11.Value)
End If

Son:

UretimOzelligi11.DropDown
If UretimOzelligi11.BackColor = RGB(60, 100, 180) Then
UretimOzelligi11.BackColor = RGB(255, 255, 255)
UretimOzelligi11.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi12.DropDown
UretimOzelligi12.BackColor = RGB(255, 255, 255)
UretimOzelligi12.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi13.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi12.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi12.ListIndex = UretimOzelligi12.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi12.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi12.ListIndex = UretimOzelligi12.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi12_Change()

If UretimOzelligi12.ListIndex = -1 And UretimOzelligi12.Value <> "" Then
   UretimOzelligi12.Value = ""
   GoTo Son
End If

If UretimOzelligi12.Value <> "" Then
    UretimOzelligi12.SelStart = 0
    UretimOzelligi12.SelLength = Len(UretimOzelligi12.Value)
End If

Son:

UretimOzelligi12.DropDown
If UretimOzelligi12.BackColor = RGB(60, 100, 180) Then
UretimOzelligi12.BackColor = RGB(255, 255, 255)
UretimOzelligi12.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi13.DropDown
UretimOzelligi13.BackColor = RGB(255, 255, 255)
UretimOzelligi13.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi14.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi13.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi13.ListIndex = UretimOzelligi13.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi13.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi13.ListIndex = UretimOzelligi13.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi13_Change()

If UretimOzelligi13.ListIndex = -1 And UretimOzelligi13.Value <> "" Then
   UretimOzelligi13.Value = ""
   GoTo Son
End If

If UretimOzelligi13.Value <> "" Then
    UretimOzelligi13.SelStart = 0
    UretimOzelligi13.SelLength = Len(UretimOzelligi13.Value)
End If

Son:

UretimOzelligi13.DropDown
If UretimOzelligi13.BackColor = RGB(60, 100, 180) Then
UretimOzelligi13.BackColor = RGB(255, 255, 255)
UretimOzelligi13.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi14.DropDown
UretimOzelligi14.BackColor = RGB(255, 255, 255)
UretimOzelligi14.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi15.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi14.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi14.ListIndex = UretimOzelligi14.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi14.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi14.ListIndex = UretimOzelligi14.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi14_Change()

If UretimOzelligi14.ListIndex = -1 And UretimOzelligi14.Value <> "" Then
   UretimOzelligi14.Value = ""
   GoTo Son
End If

If UretimOzelligi14.Value <> "" Then
    UretimOzelligi14.SelStart = 0
    UretimOzelligi14.SelLength = Len(UretimOzelligi14.Value)
End If

Son:

UretimOzelligi14.DropDown
If UretimOzelligi14.BackColor = RGB(60, 100, 180) Then
UretimOzelligi14.BackColor = RGB(255, 255, 255)
UretimOzelligi14.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi15.DropDown
UretimOzelligi15.BackColor = RGB(255, 255, 255)
UretimOzelligi15.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi16.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi15.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi15.ListIndex = UretimOzelligi15.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi15.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi15.ListIndex = UretimOzelligi15.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi15_Change()

If UretimOzelligi15.ListIndex = -1 And UretimOzelligi15.Value <> "" Then
   UretimOzelligi15.Value = ""
   GoTo Son
End If

If UretimOzelligi15.Value <> "" Then
    UretimOzelligi15.SelStart = 0
    UretimOzelligi15.SelLength = Len(UretimOzelligi15.Value)
End If

Son:

UretimOzelligi15.DropDown
If UretimOzelligi15.BackColor = RGB(60, 100, 180) Then
UretimOzelligi15.BackColor = RGB(255, 255, 255)
UretimOzelligi15.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi16.DropDown
UretimOzelligi16.BackColor = RGB(255, 255, 255)
UretimOzelligi16.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi17.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi16.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi16.ListIndex = UretimOzelligi16.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi16.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi16.ListIndex = UretimOzelligi16.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi16_Change()

If UretimOzelligi16.ListIndex = -1 And UretimOzelligi16.Value <> "" Then
   UretimOzelligi16.Value = ""
   GoTo Son
End If

If UretimOzelligi16.Value <> "" Then
    UretimOzelligi16.SelStart = 0
    UretimOzelligi16.SelLength = Len(UretimOzelligi16.Value)
End If

Son:

UretimOzelligi16.DropDown
If UretimOzelligi16.BackColor = RGB(60, 100, 180) Then
UretimOzelligi16.BackColor = RGB(255, 255, 255)
UretimOzelligi16.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi17.DropDown
UretimOzelligi17.BackColor = RGB(255, 255, 255)
UretimOzelligi17.ForeColor = RGB(30, 30, 30)


End Sub

Private Sub UretimOzelligi17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi18.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi17.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi17.ListIndex = UretimOzelligi17.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi17.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi17.ListIndex = UretimOzelligi17.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi17_Change()

If UretimOzelligi17.ListIndex = -1 And UretimOzelligi17.Value <> "" Then
   UretimOzelligi17.Value = ""
   GoTo Son
End If

If UretimOzelligi17.Value <> "" Then
    UretimOzelligi17.SelStart = 0
    UretimOzelligi17.SelLength = Len(UretimOzelligi17.Value)
End If

Son:

UretimOzelligi17.DropDown
If UretimOzelligi17.BackColor = RGB(60, 100, 180) Then
UretimOzelligi17.BackColor = RGB(255, 255, 255)
UretimOzelligi17.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub UretimOzelligi18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi18.DropDown
UretimOzelligi18.BackColor = RGB(255, 255, 255)
UretimOzelligi18.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        UretimOzelligi19.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi18.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi18.ListIndex = UretimOzelligi18.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi18.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi18.ListIndex = UretimOzelligi18.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi18_Change()

If UretimOzelligi18.ListIndex = -1 And UretimOzelligi18.Value <> "" Then
   UretimOzelligi18.Value = ""
   GoTo Son
End If

If UretimOzelligi18.Value <> "" Then
    UretimOzelligi18.SelStart = 0
    UretimOzelligi18.SelLength = Len(UretimOzelligi18.Value)
End If

Son:

UretimOzelligi18.DropDown
If UretimOzelligi18.BackColor = RGB(60, 100, 180) Then
UretimOzelligi18.BackColor = RGB(255, 255, 255)
UretimOzelligi18.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub UretimOzelligi19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.UretimOzelligi19.DropDown
UretimOzelligi19.BackColor = RGB(255, 255, 255)
UretimOzelligi19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub UretimOzelligi19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        UretimOzelligi18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'UretimOzelligi20.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If UretimOzelligi19.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi19.ListIndex = UretimOzelligi19.ListIndex
            End If
        Case 40 'Down
            If UretimOzelligi19.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                UretimOzelligi19.ListIndex = UretimOzelligi19.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub UretimOzelligi19_Change()

If UretimOzelligi19.ListIndex = -1 And UretimOzelligi19.Value <> "" Then
   UretimOzelligi19.Value = ""
   GoTo Son
End If

If UretimOzelligi19.Value <> "" Then
    UretimOzelligi19.SelStart = 0
    UretimOzelligi19.SelLength = Len(UretimOzelligi19.Value)
End If

Son:

UretimOzelligi19.DropDown
If UretimOzelligi19.BackColor = RGB(60, 100, 180) Then
UretimOzelligi19.BackColor = RGB(255, 255, 255)
UretimOzelligi19.ForeColor = RGB(30, 30, 30)
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No1.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No1.Value = Replace(Rapor1No1.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No1.Value = UCase(Replace(Replace(Rapor1No1.Value, "ı", "I"), "i", "I"))

If Rapor1No1 <> "" Then
    NotCheck1.Visible = True
    RaporOzelligi1.Enabled = True
Else
    NotCheck1.Visible = False
    NotCheck1.Value = False
    RaporOzelligi1.Value = ""
    RaporOzelligi1.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No1.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No1.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No1.Value, i, 1)) = False And Mid(Rapor1No1.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No1.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No1.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No2.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No2.Value = Replace(Rapor1No2.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No2.Value = UCase(Replace(Replace(Rapor1No2.Value, "ı", "I"), "i", "I"))

If Rapor1No2 <> "" Then
    NotCheck2.Visible = True
    RaporOzelligi2.Enabled = True
Else
    NotCheck2.Visible = False
    NotCheck2.Value = False
    RaporOzelligi2.Value = ""
    RaporOzelligi2.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No2.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No2.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No2.Value, i, 1)) = False And Mid(Rapor1No2.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No2.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No2.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No3.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No3.Value = Replace(Rapor1No3.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No3.Value = UCase(Replace(Replace(Rapor1No3.Value, "ı", "I"), "i", "I"))

If Rapor1No3 <> "" Then
    NotCheck3.Visible = True
    RaporOzelligi3.Enabled = True
Else
    NotCheck3.Visible = False
    NotCheck3.Value = False
    RaporOzelligi3.Value = ""
    RaporOzelligi3.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No3.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No3.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No3.Value, i, 1)) = False And Mid(Rapor1No3.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No3.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No3.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No3.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No4.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No4.Value = Replace(Rapor1No4.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No4.Value = UCase(Replace(Replace(Rapor1No4.Value, "ı", "I"), "i", "I"))

If Rapor1No4 <> "" Then
    NotCheck4.Visible = True
    RaporOzelligi4.Enabled = True
Else
    NotCheck4.Visible = False
    NotCheck4.Value = False
    RaporOzelligi4.Value = ""
    RaporOzelligi4.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No4.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No4.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No4.Value, i, 1)) = False And Mid(Rapor1No4.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No4.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No4.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No4.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No5.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No5.Value = Replace(Rapor1No5.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No5.Value = UCase(Replace(Replace(Rapor1No5.Value, "ı", "I"), "i", "I"))

If Rapor1No5 <> "" Then
    NotCheck5.Visible = True
    RaporOzelligi5.Enabled = True
Else
    NotCheck5.Visible = False
    NotCheck5.Value = False
    RaporOzelligi5.Value = ""
    RaporOzelligi5.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No5.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No5.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No5.Value, i, 1)) = False And Mid(Rapor1No5.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No5.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No5.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No5.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No6.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No6.Value = Replace(Rapor1No6.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No6.Value = UCase(Replace(Replace(Rapor1No6.Value, "ı", "I"), "i", "I"))

If Rapor1No6 <> "" Then
    NotCheck6.Visible = True
    RaporOzelligi6.Enabled = True
Else
    NotCheck6.Visible = False
    NotCheck6.Value = False
    RaporOzelligi6.Value = ""
    RaporOzelligi6.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No6.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No6.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No6.Value, i, 1)) = False And Mid(Rapor1No6.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No6.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No6.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No6.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No7.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No7.Value = Replace(Rapor1No7.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No7.Value = UCase(Replace(Replace(Rapor1No7.Value, "ı", "I"), "i", "I"))

If Rapor1No7 <> "" Then
    NotCheck7.Visible = True
    RaporOzelligi7.Enabled = True
Else
    NotCheck7.Visible = False
    NotCheck7.Value = False
    RaporOzelligi7.Value = ""
    RaporOzelligi7.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No7.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No7.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No7.Value, i, 1)) = False And Mid(Rapor1No7.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No7.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No7.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No8.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No8.Value = Replace(Rapor1No8.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No8.Value = UCase(Replace(Replace(Rapor1No8.Value, "ı", "I"), "i", "I"))

If Rapor1No8 <> "" Then
    NotCheck8.Visible = True
    RaporOzelligi8.Enabled = True
Else
    NotCheck8.Visible = False
    NotCheck8.Value = False
    RaporOzelligi8.Value = ""
    RaporOzelligi8.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No8.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No8.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No8.Value, i, 1)) = False And Mid(Rapor1No8.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No8.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No8.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No8.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No9.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No9.Value = Replace(Rapor1No9.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No9.Value = UCase(Replace(Replace(Rapor1No9.Value, "ı", "I"), "i", "I"))

If Rapor1No9 <> "" Then
    NotCheck9.Visible = True
    RaporOzelligi9.Enabled = True
Else
    NotCheck9.Visible = False
    NotCheck9.Value = False
    RaporOzelligi9.Value = ""
    RaporOzelligi9.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No9.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No9.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No9.Value, i, 1)) = False And Mid(Rapor1No9.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No9.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No9.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No10.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No10.Value = Replace(Rapor1No10.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No10.Value = UCase(Replace(Replace(Rapor1No10.Value, "ı", "I"), "i", "I"))

If Rapor1No10 <> "" Then
    NotCheck10.Visible = True
    RaporOzelligi10.Enabled = True
Else
    NotCheck10.Visible = False
    NotCheck10.Value = False
    RaporOzelligi10.Value = ""
    RaporOzelligi10.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No10.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No10.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No10.Value, i, 1)) = False And Mid(Rapor1No10.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No10.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No10.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No10.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No11.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No11.Value = Replace(Rapor1No11.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No11.Value = UCase(Replace(Replace(Rapor1No11.Value, "ı", "I"), "i", "I"))

If Rapor1No11 <> "" Then
    NotCheck11.Visible = True
    RaporOzelligi11.Enabled = True
Else
    NotCheck11.Visible = False
    NotCheck11.Value = False
    RaporOzelligi11.Value = ""
    RaporOzelligi11.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No11.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No11.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No11.Value, i, 1)) = False And Mid(Rapor1No11.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No11.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No11.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No11.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No12.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No12.Value = Replace(Rapor1No12.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No12.Value = UCase(Replace(Replace(Rapor1No12.Value, "ı", "I"), "i", "I"))

If Rapor1No12 <> "" Then
    NotCheck12.Visible = True
    RaporOzelligi12.Enabled = True
Else
    NotCheck12.Visible = False
    NotCheck12.Value = False
    RaporOzelligi12.Value = ""
    RaporOzelligi12.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No12.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No12.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No12.Value, i, 1)) = False And Mid(Rapor1No12.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No12.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No12.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No12.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No13.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No13.Value = Replace(Rapor1No13.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No13.Value = UCase(Replace(Replace(Rapor1No13.Value, "ı", "I"), "i", "I"))

If Rapor1No13 <> "" Then
    NotCheck13.Visible = True
    RaporOzelligi13.Enabled = True
Else
    NotCheck13.Visible = False
    NotCheck13.Value = False
    RaporOzelligi13.Value = ""
    RaporOzelligi13.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No13.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No13.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No13.Value, i, 1)) = False And Mid(Rapor1No13.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No13.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No13.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No13.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No14.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No14.Value = Replace(Rapor1No14.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No14.Value = UCase(Replace(Replace(Rapor1No14.Value, "ı", "I"), "i", "I"))

If Rapor1No14 <> "" Then
    NotCheck14.Visible = True
    RaporOzelligi14.Enabled = True
Else
    NotCheck14.Visible = False
    NotCheck14.Value = False
    RaporOzelligi14.Value = ""
    RaporOzelligi14.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No14.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No14.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No14.Value, i, 1)) = False And Mid(Rapor1No14.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No14.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No14.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No14.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No15.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No15.Value = Replace(Rapor1No15.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No15.Value = UCase(Replace(Replace(Rapor1No15.Value, "ı", "I"), "i", "I"))

If Rapor1No15 <> "" Then
    NotCheck15.Visible = True
    RaporOzelligi15.Enabled = True
Else
    NotCheck15.Visible = False
    NotCheck15.Value = False
    RaporOzelligi15.Value = ""
    RaporOzelligi15.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No15.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No15.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No15.Value, i, 1)) = False And Mid(Rapor1No15.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No15.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No15.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No15.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No16.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No16.Value = Replace(Rapor1No16.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No16.Value = UCase(Replace(Replace(Rapor1No16.Value, "ı", "I"), "i", "I"))

If Rapor1No16 <> "" Then
    NotCheck16.Visible = True
    RaporOzelligi16.Enabled = True
Else
    NotCheck16.Visible = False
    NotCheck16.Value = False
    RaporOzelligi16.Value = ""
    RaporOzelligi16.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No16.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No16.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No16.Value, i, 1)) = False And Mid(Rapor1No16.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No16.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No16.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No16.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No17.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No17.Value = Replace(Rapor1No17.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No17.Value = UCase(Replace(Replace(Rapor1No17.Value, "ı", "I"), "i", "I"))

If Rapor1No17 <> "" Then
    NotCheck17.Visible = True
    RaporOzelligi17.Enabled = True
Else
    NotCheck17.Visible = False
    NotCheck17.Value = False
    RaporOzelligi17.Value = ""
    RaporOzelligi17.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No17.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No17.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No17.Value, i, 1)) = False And Mid(Rapor1No17.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No17.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No17.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No17.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No18.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No18.Value = Replace(Rapor1No18.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No18.Value = UCase(Replace(Replace(Rapor1No18.Value, "ı", "I"), "i", "I"))

If Rapor1No18 <> "" Then
    NotCheck18.Visible = True
    RaporOzelligi18.Enabled = True
Else
    NotCheck18.Visible = False
    NotCheck18.Value = False
    RaporOzelligi18.Value = ""
    RaporOzelligi18.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No18.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No18.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No18.Value, i, 1)) = False And Mid(Rapor1No18.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No18.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No18.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No18.DropDown
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
    MsgBox """" & "/, \, <, >, ], [, :, " & """" & " , *, |, ?, _, (, ), ., ," & """" & " characters are reserved by the system and cannot be used when creating the Report 3.2 number. Please avoid using any of these characters in the Report 3.2 number. You may use the dash (-) character instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Rapor1No19.Value = ""
End If
'Boşluklara izin verme
For j = 1 To 20
Rapor1No19.Value = Replace(Rapor1No19.Value, " ", "")
Next j
'Daima büyük harf
Rapor1No19.Value = UCase(Replace(Replace(Rapor1No19.Value, "ı", "I"), "i", "I"))

If Rapor1No19 <> "" Then
    NotCheck19.Visible = True
    RaporOzelligi19.Enabled = True
Else
    NotCheck19.Visible = False
    NotCheck19.Value = False
    RaporOzelligi19.Value = ""
    RaporOzelligi19.Enabled = False
End If

'Tire hariç alfabetik karaktere izin verme
For i = 1 To 50
    If Mid(Rapor1No19.Value, i, 1) = "-" Then
        'MsgBox Mid(Rapor1No19.Value, i, 1)
    ElseIf IsNumeric(Mid(Rapor1No19.Value, i, 1)) = False And Mid(Rapor1No19.Value, i, 1) <> "" Then
        'MsgBox Mid(Rapor1No19.Value, i, 1)
        'MsgBox "Sayısal olmayan karakter var."
        Rapor1No19.Value = ""
        MsgBox "Please avoid using alphabetic characters when specifying the Report 3.2 number, except for the dash (-). The required prefix for the Report 3.2 number will be automatically added to the relevant documents.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
Next i

'Me.Rapor1No19.DropDown
Rapor1No19.BackColor = RGB(255, 255, 255)
Rapor1No19.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub RaporOzelligi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi.DropDown
RaporOzelligi.BackColor = RGB(255, 255, 255)
RaporOzelligi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi1.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi.ListIndex = RaporOzelligi.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi.ListIndex = RaporOzelligi.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi_Change()

If RaporOzelligi.ListIndex = -1 And RaporOzelligi.Value <> "" Then
   RaporOzelligi.Value = ""
   GoTo Son
End If

If RaporOzelligi.Value <> "" Then
    RaporOzelligi.SelStart = 0
    RaporOzelligi.SelLength = Len(RaporOzelligi.Value)
End If

Son:

RaporOzelligi.DropDown
If RaporOzelligi.BackColor = RGB(60, 100, 180) Then
RaporOzelligi.BackColor = RGB(255, 255, 255)
RaporOzelligi.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi1.DropDown
RaporOzelligi1.BackColor = RGB(255, 255, 255)
RaporOzelligi1.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi2.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi1.ListIndex = RaporOzelligi1.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi1.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi1.ListIndex = RaporOzelligi1.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi1_Change()

If RaporOzelligi1.ListIndex = -1 And RaporOzelligi1.Value <> "" Then
   RaporOzelligi1.Value = ""
   GoTo Son
End If

If RaporOzelligi1.Value <> "" Then
    RaporOzelligi1.SelStart = 0
    RaporOzelligi1.SelLength = Len(RaporOzelligi1.Value)
End If

Son:

RaporOzelligi1.DropDown
If RaporOzelligi1.BackColor = RGB(60, 100, 180) Then
RaporOzelligi1.BackColor = RGB(255, 255, 255)
RaporOzelligi1.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi2.DropDown
RaporOzelligi2.BackColor = RGB(255, 255, 255)
RaporOzelligi2.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi1.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi3.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi2.ListIndex = RaporOzelligi2.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi2.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi2.ListIndex = RaporOzelligi2.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi2_Change()

If RaporOzelligi2.ListIndex = -1 And RaporOzelligi2.Value <> "" Then
   RaporOzelligi2.Value = ""
   GoTo Son
End If

If RaporOzelligi2.Value <> "" Then
    RaporOzelligi2.SelStart = 0
    RaporOzelligi2.SelLength = Len(RaporOzelligi2.Value)
End If

Son:

RaporOzelligi2.DropDown
If RaporOzelligi2.BackColor = RGB(60, 100, 180) Then
RaporOzelligi2.BackColor = RGB(255, 255, 255)
RaporOzelligi2.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi3.DropDown
RaporOzelligi3.BackColor = RGB(255, 255, 255)
RaporOzelligi3.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi2.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi4.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi3.ListIndex = RaporOzelligi3.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi3.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi3.ListIndex = RaporOzelligi3.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi3_Change()

If RaporOzelligi3.ListIndex = -1 And RaporOzelligi3.Value <> "" Then
   RaporOzelligi3.Value = ""
   GoTo Son
End If

If RaporOzelligi3.Value <> "" Then
    RaporOzelligi3.SelStart = 0
    RaporOzelligi3.SelLength = Len(RaporOzelligi3.Value)
End If

Son:

RaporOzelligi3.DropDown
If RaporOzelligi3.BackColor = RGB(60, 100, 180) Then
RaporOzelligi3.BackColor = RGB(255, 255, 255)
RaporOzelligi3.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi4.DropDown
RaporOzelligi4.BackColor = RGB(255, 255, 255)
RaporOzelligi4.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi3.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi5.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi4.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi4.ListIndex = RaporOzelligi4.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi4.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi4.ListIndex = RaporOzelligi4.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi4_Change()

If RaporOzelligi4.ListIndex = -1 And RaporOzelligi4.Value <> "" Then
   RaporOzelligi4.Value = ""
   GoTo Son
End If

If RaporOzelligi4.Value <> "" Then
    RaporOzelligi4.SelStart = 0
    RaporOzelligi4.SelLength = Len(RaporOzelligi4.Value)
End If

Son:

RaporOzelligi4.DropDown
If RaporOzelligi4.BackColor = RGB(60, 100, 180) Then
RaporOzelligi4.BackColor = RGB(255, 255, 255)
RaporOzelligi4.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi5.DropDown
RaporOzelligi5.BackColor = RGB(255, 255, 255)
RaporOzelligi5.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi4.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi6.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi5.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi5.ListIndex = RaporOzelligi5.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi5.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi5.ListIndex = RaporOzelligi5.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi5_Change()

If RaporOzelligi5.ListIndex = -1 And RaporOzelligi5.Value <> "" Then
   RaporOzelligi5.Value = ""
   GoTo Son
End If

If RaporOzelligi5.Value <> "" Then
    RaporOzelligi5.SelStart = 0
    RaporOzelligi5.SelLength = Len(RaporOzelligi5.Value)
End If

Son:

RaporOzelligi5.DropDown
If RaporOzelligi5.BackColor = RGB(60, 100, 180) Then
RaporOzelligi5.BackColor = RGB(255, 255, 255)
RaporOzelligi5.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi6.DropDown
RaporOzelligi6.BackColor = RGB(255, 255, 255)
RaporOzelligi6.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi5.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi7.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi6.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi6.ListIndex = RaporOzelligi6.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi6.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi6.ListIndex = RaporOzelligi6.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi6_Change()

If RaporOzelligi6.ListIndex = -1 And RaporOzelligi6.Value <> "" Then
   RaporOzelligi6.Value = ""
   GoTo Son
End If

If RaporOzelligi6.Value <> "" Then
    RaporOzelligi6.SelStart = 0
    RaporOzelligi6.SelLength = Len(RaporOzelligi6.Value)
End If

Son:

RaporOzelligi6.DropDown
If RaporOzelligi6.BackColor = RGB(60, 100, 180) Then
RaporOzelligi6.BackColor = RGB(255, 255, 255)
RaporOzelligi6.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi7.DropDown
RaporOzelligi7.BackColor = RGB(255, 255, 255)
RaporOzelligi7.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi6.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi8.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi7.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi7.ListIndex = RaporOzelligi7.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi7.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi7.ListIndex = RaporOzelligi7.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi7_Change()

If RaporOzelligi7.ListIndex = -1 And RaporOzelligi7.Value <> "" Then
   RaporOzelligi7.Value = ""
   GoTo Son
End If

If RaporOzelligi7.Value <> "" Then
    RaporOzelligi7.SelStart = 0
    RaporOzelligi7.SelLength = Len(RaporOzelligi7.Value)
End If

Son:

RaporOzelligi7.DropDown
If RaporOzelligi7.BackColor = RGB(60, 100, 180) Then
RaporOzelligi7.BackColor = RGB(255, 255, 255)
RaporOzelligi7.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi8.DropDown
RaporOzelligi8.BackColor = RGB(255, 255, 255)
RaporOzelligi8.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi7.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi9.SetFocus
    End If
    
    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi8.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi8.ListIndex = RaporOzelligi8.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi8.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi8.ListIndex = RaporOzelligi8.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi8_Change()

If RaporOzelligi8.ListIndex = -1 And RaporOzelligi8.Value <> "" Then
   RaporOzelligi8.Value = ""
   GoTo Son
End If

If RaporOzelligi8.Value <> "" Then
    RaporOzelligi8.SelStart = 0
    RaporOzelligi8.SelLength = Len(RaporOzelligi8.Value)
End If

Son:

RaporOzelligi8.DropDown
If RaporOzelligi8.BackColor = RGB(60, 100, 180) Then
RaporOzelligi8.BackColor = RGB(255, 255, 255)
RaporOzelligi8.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi9.DropDown
RaporOzelligi9.BackColor = RGB(255, 255, 255)
RaporOzelligi9.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi8.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi10.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi9.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi9.ListIndex = RaporOzelligi9.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi9.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi9.ListIndex = RaporOzelligi9.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi9_Change()

If RaporOzelligi9.ListIndex = -1 And RaporOzelligi9.Value <> "" Then
   RaporOzelligi9.Value = ""
   GoTo Son
End If

If RaporOzelligi9.Value <> "" Then
    RaporOzelligi9.SelStart = 0
    RaporOzelligi9.SelLength = Len(RaporOzelligi9.Value)
End If

Son:

RaporOzelligi9.DropDown
If RaporOzelligi9.BackColor = RGB(60, 100, 180) Then
RaporOzelligi9.BackColor = RGB(255, 255, 255)
RaporOzelligi9.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi10.DropDown
RaporOzelligi10.BackColor = RGB(255, 255, 255)
RaporOzelligi10.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi9.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi11.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi10.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi10.ListIndex = RaporOzelligi10.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi10.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi10.ListIndex = RaporOzelligi10.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi10_Change()

If RaporOzelligi10.ListIndex = -1 And RaporOzelligi10.Value <> "" Then
   RaporOzelligi10.Value = ""
   GoTo Son
End If

If RaporOzelligi10.Value <> "" Then
    RaporOzelligi10.SelStart = 0
    RaporOzelligi10.SelLength = Len(RaporOzelligi10.Value)
End If

Son:

RaporOzelligi10.DropDown
If RaporOzelligi10.BackColor = RGB(60, 100, 180) Then
RaporOzelligi10.BackColor = RGB(255, 255, 255)
RaporOzelligi10.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi11.DropDown
RaporOzelligi11.BackColor = RGB(255, 255, 255)
RaporOzelligi11.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi10.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi12.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi11.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi11.ListIndex = RaporOzelligi11.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi11.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi11.ListIndex = RaporOzelligi11.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi11_Change()

If RaporOzelligi11.ListIndex = -1 And RaporOzelligi11.Value <> "" Then
   RaporOzelligi11.Value = ""
   GoTo Son
End If

If RaporOzelligi11.Value <> "" Then
    RaporOzelligi11.SelStart = 0
    RaporOzelligi11.SelLength = Len(RaporOzelligi11.Value)
End If

Son:

RaporOzelligi11.DropDown
If RaporOzelligi11.BackColor = RGB(60, 100, 180) Then
RaporOzelligi11.BackColor = RGB(255, 255, 255)
RaporOzelligi11.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi12.DropDown
RaporOzelligi12.BackColor = RGB(255, 255, 255)
RaporOzelligi12.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi11.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi13.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi12.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi12.ListIndex = RaporOzelligi12.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi12.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi12.ListIndex = RaporOzelligi12.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi12_Change()

If RaporOzelligi12.ListIndex = -1 And RaporOzelligi12.Value <> "" Then
   RaporOzelligi12.Value = ""
   GoTo Son
End If

If RaporOzelligi12.Value <> "" Then
    RaporOzelligi12.SelStart = 0
    RaporOzelligi12.SelLength = Len(RaporOzelligi12.Value)
End If

Son:

RaporOzelligi12.DropDown
If RaporOzelligi12.BackColor = RGB(60, 100, 180) Then
RaporOzelligi12.BackColor = RGB(255, 255, 255)
RaporOzelligi12.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi13.DropDown
RaporOzelligi13.BackColor = RGB(255, 255, 255)
RaporOzelligi13.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi12.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi14.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi13.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi13.ListIndex = RaporOzelligi13.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi13.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi13.ListIndex = RaporOzelligi13.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi13_Change()

If RaporOzelligi13.ListIndex = -1 And RaporOzelligi13.Value <> "" Then
   RaporOzelligi13.Value = ""
   GoTo Son
End If

If RaporOzelligi13.Value <> "" Then
    RaporOzelligi13.SelStart = 0
    RaporOzelligi13.SelLength = Len(RaporOzelligi13.Value)
End If

Son:

RaporOzelligi13.DropDown
If RaporOzelligi13.BackColor = RGB(60, 100, 180) Then
RaporOzelligi13.BackColor = RGB(255, 255, 255)
RaporOzelligi13.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi14.DropDown
RaporOzelligi14.BackColor = RGB(255, 255, 255)
RaporOzelligi14.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi13.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi15.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi14.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi14.ListIndex = RaporOzelligi14.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi14.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi14.ListIndex = RaporOzelligi14.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi14_Change()

If RaporOzelligi14.ListIndex = -1 And RaporOzelligi14.Value <> "" Then
   RaporOzelligi14.Value = ""
   GoTo Son
End If

If RaporOzelligi14.Value <> "" Then
    RaporOzelligi14.SelStart = 0
    RaporOzelligi14.SelLength = Len(RaporOzelligi14.Value)
End If

Son:

RaporOzelligi14.DropDown
If RaporOzelligi14.BackColor = RGB(60, 100, 180) Then
RaporOzelligi14.BackColor = RGB(255, 255, 255)
RaporOzelligi14.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi15.DropDown
RaporOzelligi15.BackColor = RGB(255, 255, 255)
RaporOzelligi15.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi14.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi16.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi15.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi15.ListIndex = RaporOzelligi15.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi15.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi15.ListIndex = RaporOzelligi15.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi15_Change()

If RaporOzelligi15.ListIndex = -1 And RaporOzelligi15.Value <> "" Then
   RaporOzelligi15.Value = ""
   GoTo Son
End If

If RaporOzelligi15.Value <> "" Then
    RaporOzelligi15.SelStart = 0
    RaporOzelligi15.SelLength = Len(RaporOzelligi15.Value)
End If

Son:

RaporOzelligi15.DropDown
If RaporOzelligi15.BackColor = RGB(60, 100, 180) Then
RaporOzelligi15.BackColor = RGB(255, 255, 255)
RaporOzelligi15.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi16.DropDown
RaporOzelligi16.BackColor = RGB(255, 255, 255)
RaporOzelligi16.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi15.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi17.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi16.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi16.ListIndex = RaporOzelligi16.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi16.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi16.ListIndex = RaporOzelligi16.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi16_Change()

If RaporOzelligi16.ListIndex = -1 And RaporOzelligi16.Value <> "" Then
   RaporOzelligi16.Value = ""
   GoTo Son
End If

If RaporOzelligi16.Value <> "" Then
    RaporOzelligi16.SelStart = 0
    RaporOzelligi16.SelLength = Len(RaporOzelligi16.Value)
End If

Son:

RaporOzelligi16.DropDown
If RaporOzelligi16.BackColor = RGB(60, 100, 180) Then
RaporOzelligi16.BackColor = RGB(255, 255, 255)
RaporOzelligi16.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi17.DropDown
RaporOzelligi17.BackColor = RGB(255, 255, 255)
RaporOzelligi17.ForeColor = RGB(30, 30, 30)


End Sub

Private Sub RaporOzelligi17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi16.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi18.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi17.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi17.ListIndex = RaporOzelligi17.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi17.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi17.ListIndex = RaporOzelligi17.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi17_Change()

If RaporOzelligi17.ListIndex = -1 And RaporOzelligi17.Value <> "" Then
   RaporOzelligi17.Value = ""
   GoTo Son
End If

If RaporOzelligi17.Value <> "" Then
    RaporOzelligi17.SelStart = 0
    RaporOzelligi17.SelLength = Len(RaporOzelligi17.Value)
End If

Son:

RaporOzelligi17.DropDown
If RaporOzelligi17.BackColor = RGB(60, 100, 180) Then
RaporOzelligi17.BackColor = RGB(255, 255, 255)
RaporOzelligi17.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub RaporOzelligi18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi18.DropDown
RaporOzelligi18.BackColor = RGB(255, 255, 255)
RaporOzelligi18.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi17.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        RaporOzelligi19.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi18.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi18.ListIndex = RaporOzelligi18.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi18.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi18.ListIndex = RaporOzelligi18.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi18_Change()

If RaporOzelligi18.ListIndex = -1 And RaporOzelligi18.Value <> "" Then
   RaporOzelligi18.Value = ""
   GoTo Son
End If

If RaporOzelligi18.Value <> "" Then
    RaporOzelligi18.SelStart = 0
    RaporOzelligi18.SelLength = Len(RaporOzelligi18.Value)
End If

Son:

RaporOzelligi18.DropDown
If RaporOzelligi18.BackColor = RGB(60, 100, 180) Then
RaporOzelligi18.BackColor = RGB(255, 255, 255)
RaporOzelligi18.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub RaporOzelligi19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.RaporOzelligi19.DropDown
RaporOzelligi19.BackColor = RGB(255, 255, 255)
RaporOzelligi19.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub RaporOzelligi19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    'Yukarı ve aşağıya
    If KeyCode = vbKeyUp Then
        RaporOzelligi18.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        'RaporOzelligi20.SetFocus
    End If

    Select Case KeyCode
        Case 38  'Up
            If RaporOzelligi19.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi19.ListIndex = RaporOzelligi19.ListIndex
            End If
        Case 40 'Down
            If RaporOzelligi19.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                RaporOzelligi19.ListIndex = RaporOzelligi19.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub RaporOzelligi19_Change()

If RaporOzelligi19.ListIndex = -1 And RaporOzelligi19.Value <> "" Then
   RaporOzelligi19.Value = ""
   GoTo Son
End If

If RaporOzelligi19.Value <> "" Then
    RaporOzelligi19.SelStart = 0
    RaporOzelligi19.SelLength = Len(RaporOzelligi19.Value)
End If

Son:

RaporOzelligi19.DropDown
If RaporOzelligi19.BackColor = RGB(60, 100, 180) Then
RaporOzelligi19.BackColor = RGB(255, 255, 255)
RaporOzelligi19.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub NotCheck_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck.Value = True And OgeTuru.Value = "" Then
    NotCheck.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:

End Sub
Private Sub NotCheck1_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck1.Value = True And OgeTuru1.Value = "" Then
    NotCheck1.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru1.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck1.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck1.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck2_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck2.Value = True And OgeTuru2.Value = "" Then
    NotCheck2.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru2.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck2.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck2.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck3_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck3.Value = True And OgeTuru3.Value = "" Then
    NotCheck3.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru3.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck3.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck3.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck4_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck4.Value = True And OgeTuru4.Value = "" Then
    NotCheck4.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru4.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck4.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck4.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck5_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck5.Value = True And OgeTuru5.Value = "" Then
    NotCheck5.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru5.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck5.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck5.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck6_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck6.Value = True And OgeTuru6.Value = "" Then
    NotCheck6.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru6.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck6.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck6.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck7_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck7.Value = True And OgeTuru7.Value = "" Then
    NotCheck7.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru7.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck7.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck7.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck8_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck8.Value = True And OgeTuru8.Value = "" Then
    NotCheck8.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru8.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck8.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck8.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck9_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck9.Value = True And OgeTuru9.Value = "" Then
    NotCheck9.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru9.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck9.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck9.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck10_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck10.Value = True And OgeTuru10.Value = "" Then
    NotCheck10.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru10.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck10.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck10.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck11_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck11.Value = True And OgeTuru11.Value = "" Then
    NotCheck11.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru11.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck11.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck11.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck12_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck12.Value = True And OgeTuru12.Value = "" Then
    NotCheck12.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru12.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck12.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck12.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck13_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck13.Value = True And OgeTuru13.Value = "" Then
    NotCheck13.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru13.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck13.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck13.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck14_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck14.Value = True And OgeTuru14.Value = "" Then
    NotCheck14.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru14.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck14.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck14.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck15_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck15.Value = True And OgeTuru15.Value = "" Then
    NotCheck15.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru15.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck15.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck15.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck16_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck16.Value = True And OgeTuru16.Value = "" Then
    NotCheck16.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru16.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck16.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck16.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck17_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck17.Value = True And OgeTuru17.Value = "" Then
    NotCheck17.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru17.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck17.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck17.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck18_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck18.Value = True And OgeTuru18.Value = "" Then
    NotCheck18.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru18.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck18.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck18.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
End Sub
Private Sub NotCheck19_Click()
Dim AutoPath As String, DestTarget As String, FileName As String
Dim HedefFile As String

If NotCheck19.Value = True And OgeTuru19.Value = "" Then
    NotCheck19.Value = False
    MsgBox "A note cannot be added without selecting an item type. Please select the item type before adding a note.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"
FileName = OgeTuru19.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If NotCheck19.Value = True And Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    NotCheck19.Value = False
    MsgBox "Since no note has been created for the " & FileName & ", the note cannot be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
  
Son:
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

Private Sub FinansalBirimUstYaziTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        FinansalBirimUstYaziTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        FinansalBirimUstYaziTarihiText.Value = ""
    End If
    
End Sub

Private Sub FinansalBirimUstYaziTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

FinansalBirimUstYaziTarihiText.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimUstYaziTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    FinansalBirimUstYaziTarihiText.Value = CalTarih
    FinansalBirimUstYaziTarihiText.Value = Format(FinansalBirimUstYaziTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

FinansalBirimUstYaziTarihiText.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziTarihiText.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimUstYaziNoText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
FinansalBirimUstYaziNoText.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziNoText.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub FinansalBirimUstYaziNoText_Change()
FinansalBirimUstYaziNoText.BackColor = RGB(255, 255, 255)
FinansalBirimUstYaziNoText.ForeColor = RGB(30, 30, 30)
End Sub

Private Sub IlFinansalBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.IlFinansalBirim.DropDown
IlFinansalBirim.BackColor = RGB(255, 255, 255)
IlFinansalBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub IlFinansalBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If IlFinansalBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlFinansalBirim.ListIndex = IlFinansalBirim.ListIndex - 1
            End If
            Me.IlFinansalBirim.DropDown
            
        Case 40 'Aşağı
            If IlFinansalBirim.ListIndex = IlFinansalBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlFinansalBirim.ListIndex = IlFinansalBirim.ListIndex + 1
            End If
            Me.IlFinansalBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub IlFinansalBirim_Change()

If IlFinansalBirim.ListIndex = -1 And IlFinansalBirim.Value <> "" Then
   IlFinansalBirim.Value = ""
   GoTo Son
End If

'Ilçe seçimlerini İl seçimine göre göster.
On Error GoTo Bos
IlceFinansalBirim.RowSource = Replace(IlFinansalBirim.Value, " ", "_")
'IlceFinansalBirim.DropDown
GoTo Son

Bos:
IlceFinansalBirim.RowSource = ""

Son:

If IlFinansalBirim.Value <> "" Then
    IlFinansalBirim.SelStart = 0
    IlFinansalBirim.SelLength = Len(IlFinansalBirim.Value)
End If

IlFinansalBirim.DropDown
If IlFinansalBirim.BackColor = RGB(60, 100, 180) Then
IlFinansalBirim.BackColor = RGB(255, 255, 255)
IlFinansalBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub IlceFinansalBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.IlceFinansalBirim.DropDown
IlceFinansalBirim.BackColor = RGB(255, 255, 255)
IlceFinansalBirim.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub IlceFinansalBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If IlceFinansalBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlceFinansalBirim.ListIndex = IlceFinansalBirim.ListIndex - 1
            End If
            Me.IlceFinansalBirim.DropDown
            
        Case 40 'Aşağı
            If IlceFinansalBirim.ListIndex = IlceFinansalBirim.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                IlceFinansalBirim.ListIndex = IlceFinansalBirim.ListIndex + 1
            End If
            Me.IlceFinansalBirim.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub IlceFinansalBirim_Change()

If IlceFinansalBirim.ListIndex = -1 And IlceFinansalBirim.Value <> "" Then
   IlceFinansalBirim.Value = ""
   GoTo Son
End If

If IlceFinansalBirim.Value <> "" Then
    IlceFinansalBirim.SelStart = 0
    IlceFinansalBirim.SelLength = Len(IlceFinansalBirim.Value)
End If

Son:

IlceFinansalBirim.DropDown
If IlceFinansalBirim.BackColor = RGB(60, 100, 180) Then
IlceFinansalBirim.BackColor = RGB(255, 255, 255)
IlceFinansalBirim.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub FinansalBirimGonderilen_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.FinansalBirimGonderilen.DropDown
FinansalBirimGonderilen.BackColor = RGB(255, 255, 255)
FinansalBirimGonderilen.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimGonderilen_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If FinansalBirimGonderilen.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimGonderilen.ListIndex = FinansalBirimGonderilen.ListIndex - 1
            End If
            Me.FinansalBirimGonderilen.DropDown
            
        Case 40 'Aşağı
            If FinansalBirimGonderilen.ListIndex = FinansalBirimGonderilen.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimGonderilen.ListIndex = FinansalBirimGonderilen.ListIndex + 1
            End If
            Me.FinansalBirimGonderilen.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub FinansalBirimGonderilen_Change()

If FinansalBirimGonderilen.ListIndex = -1 And FinansalBirimGonderilen.Value <> "" Then
   FinansalBirimGonderilen.Value = ""
   GoTo Son
End If

If FinansalBirimGonderilen.Value <> "" Then
    FinansalBirimGonderilen.SelStart = 0
    FinansalBirimGonderilen.SelLength = Len(FinansalBirimGonderilen.Value)
End If

Son:

FinansalBirimGonderilen.DropDown
If FinansalBirimGonderilen.BackColor = RGB(60, 100, 180) Then
FinansalBirimGonderilen.BackColor = RGB(255, 255, 255)
FinansalBirimGonderilen.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub Adres_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Adres.BackColor = RGB(255, 255, 255)
Adres.ForeColor = RGB(30, 30, 30)
End Sub
Private Sub Adres_Change()
Adres.BackColor = RGB(255, 255, 255)
Adres.ForeColor = RGB(30, 30, 30)
End Sub


Private Sub FinansalBirimGonderiTipi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.FinansalBirimGonderiTipi.DropDown

FinansalBirimGonderiTipi.BackColor = RGB(255, 255, 255)
FinansalBirimGonderiTipi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub FinansalBirimGonderiTipi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If FinansalBirimGonderiTipi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimGonderiTipi.ListIndex = FinansalBirimGonderiTipi.ListIndex - 1
            End If
            Me.FinansalBirimGonderiTipi.DropDown
            
        Case 40 'Aşağı
            If FinansalBirimGonderiTipi.ListIndex = FinansalBirimGonderiTipi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                FinansalBirimGonderiTipi.ListIndex = FinansalBirimGonderiTipi.ListIndex + 1
            End If
            Me.FinansalBirimGonderiTipi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub FinansalBirimGonderiTipi_Change()

If FinansalBirimGonderiTipi.ListIndex = -1 And FinansalBirimGonderiTipi.Value <> "" Then
   FinansalBirimGonderiTipi.Value = ""
   GoTo Son
End If

If FinansalBirimGonderiTipi.Value <> "" Then
    FinansalBirimGonderiTipi.SelStart = 0
    FinansalBirimGonderiTipi.SelLength = Len(FinansalBirimGonderiTipi.Value)
End If

Son:

FinansalBirimGonderiTipi.DropDown
If FinansalBirimGonderiTipi.BackColor = RGB(60, 100, 180) Then
FinansalBirimGonderiTipi.BackColor = RGB(255, 255, 255)
FinansalBirimGonderiTipi.ForeColor = RGB(30, 30, 30)
End If

End Sub


Private Sub DekontSayfa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'Me.DekontSayfa.DropDown
DekontSayfa.BackColor = RGB(255, 255, 255)
DekontSayfa.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub DekontSayfa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If DekontSayfa.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DekontSayfa.ListIndex = DekontSayfa.ListIndex - 1
            End If
            Me.DekontSayfa.DropDown
            
        Case 40 'Aşağı
            If DekontSayfa.ListIndex = DekontSayfa.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DekontSayfa.ListIndex = DekontSayfa.ListIndex + 1
            End If
            Me.DekontSayfa.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub DekontSayfa_Change()

If DekontSayfa.ListIndex = -1 And DekontSayfa.Value <> "" Then
   DekontSayfa.Value = ""
   GoTo Son
End If

If DekontSayfa.Value <> "" Then
    DekontSayfa.SelStart = 0
    DekontSayfa.SelLength = Len(DekontSayfa.Value)
End If

Son:

DekontSayfa.DropDown
If DekontSayfa.BackColor = RGB(60, 100, 180) Then
DekontSayfa.BackColor = RGB(255, 255, 255)
DekontSayfa.ForeColor = RGB(30, 30, 30)
End If

End Sub
Private Sub DesteBandiAdedi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Me.DesteBandiAdedi.DropDown
DesteBandiAdedi.BackColor = RGB(255, 255, 255)
DesteBandiAdedi.ForeColor = RGB(30, 30, 30)

End Sub

Private Sub DesteBandiAdedi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
   Select Case KeyCode
        Case 38  'Yukarı
            If DesteBandiAdedi.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DesteBandiAdedi.ListIndex = DesteBandiAdedi.ListIndex - 1
            End If
            Me.DesteBandiAdedi.DropDown
            
        Case 40 'Aşağı
            If DesteBandiAdedi.ListIndex = DesteBandiAdedi.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                DesteBandiAdedi.ListIndex = DesteBandiAdedi.ListIndex + 1
            End If
            Me.DesteBandiAdedi.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub DesteBandiAdedi_Change()

'If DesteBandiAdedi.ListIndex = -1 And DesteBandiAdedi.Value <> "" Then
'   DesteBandiAdedi.Value = ""
'   GoTo Son
'End If

'If DesteBandiAdedi.Value <> "" Then
'    DesteBandiAdedi.SelStart = 0
'    DesteBandiAdedi.SelLength = Len(DesteBandiAdedi.Value)
'End If

'DesteBandiAdedi.DropDown

Son:

DesteBandiAdedi.DropDown
If DesteBandiAdedi.BackColor = RGB(60, 100, 180) Then
DesteBandiAdedi.BackColor = RGB(255, 255, 255)
DesteBandiAdedi.ForeColor = RGB(30, 30, 30)
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

Sub Son20RaporNo()
Dim i As Integer
Dim Say As Long, j As Long, Cont As Long, Rno As Variant
Dim RefSatir As Long, Rapor1TarihBul As Range

'Verilen son 20 rapor numarasını göster

ThisWorkbook.Activate

'__________Rapor No Senkronizasyon 30.11.2021
    
Set WsRaporNo = ThisWorkbook.Worksheets(10)

'İlk satırda bulunan rapor1 numarası
Say = WsRaporNo.Range("E100000").End(xlUp).Row
If Say < 7 Then
    Say = 7
End If

RefSatir = 0
Set WsRaporNo = ThisWorkbook.Worksheets(10)
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

    Controls("RaporOzelligi" & i).Visible = True
    Controls("LblRaporOzelligi" & i).Visible = True
    Controls("UretimOzelligi" & i).Visible = True
    Controls("LblUretimOzelligi" & i).Visible = True
    
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
    If ScrollTakip = 48 Then
        ScrollTakip = ScrollTakip + 24 + 6
    ElseIf ScrollTakip > 48 Then
        ScrollTakip = ScrollTakip + 24
    Else
        ScrollTakip = ScrollTakip + 24
    End If
End If
If ScrollTakip > 48 + 6 Then
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
        Controls("UretimOzelligi" & OgeFrame).Value <> "" Or _
        Controls("RaporOzelligi" & OgeFrame).Value <> "" Or _
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
If OgeTuru.Value = "" And OgeDegeri.Value = "" And Adet.Value = "" And OgeIdNo.Value = "" And Aciklama.Value = "" And Sonuc.Value = "" And _
    UretimOzelligi.Value = "" And RaporOzelligi.Value = "" And Rapor1No.Value = "" Then

    For OgeFrame = 1 To SonDoluSatir

        If OgeFrame = 1 Then
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Or _
                Controls("Sonuc" & OgeFrame).Value <> "" Or _
                Controls("UretimOzelligi" & OgeFrame).Value <> "" Or _
                Controls("RaporOzelligi" & OgeFrame).Value <> "" Or _
                Controls("Rapor1No" & OgeFrame).Value <> "" Then

                OgeTuru.Value = Controls("OgeTuru" & OgeFrame).Value
                OgeDegeri.Value = Controls("OgeDegeri" & OgeFrame).Value
                Adet.Value = Controls("Adet" & OgeFrame).Value
                OgeIdNo.Value = Controls("OgeIdNo" & OgeFrame).Value
                Aciklama.Value = Controls("Aciklama" & OgeFrame).Value
                Sonuc.Value = Controls("Sonuc" & OgeFrame).Value
                UretimOzelligi.Value = Controls("UretimOzelligi" & OgeFrame).Value
                RaporOzelligi.Value = Controls("RaporOzelligi" & OgeFrame).Value
                Rapor1No.Value = Controls("Rapor1No" & OgeFrame).Value
                NotCheck.Value = Controls("NotCheck" & OgeFrame).Value
                
                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""
                Controls("Sonuc" & OgeFrame).Value = ""
                Controls("UretimOzelligi" & OgeFrame).Value = ""
                Controls("RaporOzelligi" & OgeFrame).Value = ""
                Controls("Rapor1No" & OgeFrame).Value = ""
                Controls("NotCheck" & OgeFrame).Value = False

            End If
        Else
            If Controls("OgeTuru" & OgeFrame).Value <> "" Or _
                Controls("OgeDegeri" & OgeFrame).Value <> "" Or _
                Controls("Adet" & OgeFrame).Value <> "" Or _
                Controls("OgeIdNo" & OgeFrame).Value <> "" Or _
                Controls("Aciklama" & OgeFrame).Value <> "" Or _
                Controls("Sonuc" & OgeFrame).Value <> "" Or _
                Controls("UretimOzelligi" & OgeFrame).Value <> "" Or _
                Controls("RaporOzelligi" & OgeFrame).Value <> "" Or _
                Controls("Rapor1No" & OgeFrame).Value <> "" Then

                Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
                Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
                Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
                Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
                Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value
                Controls("Sonuc" & OgeFrame - 1).Value = Controls("Sonuc" & OgeFrame).Value
                Controls("UretimOzelligi" & OgeFrame - 1).Value = Controls("UretimOzelligi" & OgeFrame).Value
                Controls("RaporOzelligi" & OgeFrame - 1).Value = Controls("RaporOzelligi" & OgeFrame).Value
                Controls("Rapor1No" & OgeFrame - 1).Value = Controls("Rapor1No" & OgeFrame).Value
                Controls("NotCheck" & OgeFrame - 1).Value = Controls("NotCheck" & OgeFrame).Value

                Controls("OgeTuru" & OgeFrame).Value = ""
                Controls("OgeDegeri" & OgeFrame).Value = ""
                Controls("Adet" & OgeFrame).Value = ""
                Controls("OgeIdNo" & OgeFrame).Value = ""
                Controls("Aciklama" & OgeFrame).Value = ""
                Controls("Sonuc" & OgeFrame).Value = ""
                Controls("UretimOzelligi" & OgeFrame).Value = ""
                Controls("RaporOzelligi" & OgeFrame).Value = ""
                Controls("Rapor1No" & OgeFrame).Value = ""
                Controls("NotCheck" & OgeFrame).Value = False

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
        Controls("UretimOzelligi" & OgeFrame).Value = "" And _
        Controls("RaporOzelligi" & OgeFrame).Value = "" And _
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
        Controls("UretimOzelligi" & OgeFrame).Value <> "" Or _
        Controls("RaporOzelligi" & OgeFrame).Value <> "" Or _
        Controls("Rapor1No" & OgeFrame).Value <> "" Then

        Controls("OgeTuru" & OgeFrame - 1).Value = Controls("OgeTuru" & OgeFrame).Value
        Controls("OgeDegeri" & OgeFrame - 1).Value = Controls("OgeDegeri" & OgeFrame).Value
        Controls("Adet" & OgeFrame - 1).Value = Controls("Adet" & OgeFrame).Value
        Controls("OgeIdNo" & OgeFrame - 1).Value = Controls("OgeIdNo" & OgeFrame).Value
        Controls("Aciklama" & OgeFrame - 1).Value = Controls("Aciklama" & OgeFrame).Value
        Controls("Sonuc" & OgeFrame - 1).Value = Controls("Sonuc" & OgeFrame).Value
        Controls("UretimOzelligi" & OgeFrame - 1).Value = Controls("UretimOzelligi" & OgeFrame).Value
        Controls("RaporOzelligi" & OgeFrame - 1).Value = Controls("RaporOzelligi" & OgeFrame).Value
        Controls("Rapor1No" & OgeFrame - 1).Value = Controls("Rapor1No" & OgeFrame).Value
        Controls("NotCheck" & OgeFrame - 1).Value = Controls("NotCheck" & OgeFrame).Value
        
        Controls("OgeTuru" & OgeFrame).Value = ""
        Controls("OgeDegeri" & OgeFrame).Value = ""
        Controls("Adet" & OgeFrame).Value = ""
        Controls("OgeIdNo" & OgeFrame).Value = ""
        Controls("Aciklama" & OgeFrame).Value = ""
        Controls("Sonuc" & OgeFrame).Value = ""
        Controls("UretimOzelligi" & OgeFrame).Value = ""
        Controls("RaporOzelligi" & OgeFrame).Value = ""
        Controls("Rapor1No" & OgeFrame).Value = ""
        Controls("NotCheck" & OgeFrame).Value = False

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
    UretimOzelligi19.Value = ""
    RaporOzelligi19.Enabled = False
    RaporOzelligi19.Value = ""
    Rapor1No19.Value = ""
    OgeTuruFrame19.Visible = False
    NotCheck19.Value = False
    NotCheck19.Visible = False
    GoTo Son
ElseIf OgeTuruFrame18.Visible = True Then
    OgeTuru18.Value = ""
    OgeDegeri18.Value = ""
    Adet18.Value = ""
    OgeIdNo18.Value = ""
    Aciklama18.Value = ""
    Sonuc18.Value = ""
    UretimOzelligi18.Value = ""
    RaporOzelligi18.Enabled = False
    RaporOzelligi18.Value = ""
    Rapor1No18.Value = ""
    OgeTuruFrame18.Visible = False
    NotCheck18.Value = False
    NotCheck18.Visible = False
    GoTo Son
ElseIf OgeTuruFrame17.Visible = True Then
    OgeTuru17.Value = ""
    OgeDegeri17.Value = ""
    Adet17.Value = ""
    OgeIdNo17.Value = ""
    Aciklama17.Value = ""
    Sonuc17.Value = ""
    UretimOzelligi17.Value = ""
    RaporOzelligi17.Enabled = False
    RaporOzelligi17.Value = ""
    Rapor1No17.Value = ""
    OgeTuruFrame17.Visible = False
    NotCheck17.Value = False
    NotCheck17.Visible = False
    GoTo Son
ElseIf OgeTuruFrame16.Visible = True Then
    OgeTuru16.Value = ""
    OgeDegeri16.Value = ""
    Adet16.Value = ""
    OgeIdNo16.Value = ""
    Aciklama16.Value = ""
    Sonuc16.Value = ""
    UretimOzelligi16.Value = ""
    RaporOzelligi16.Enabled = False
    RaporOzelligi16.Value = ""
    Rapor1No16.Value = ""
    OgeTuruFrame16.Visible = False
    NotCheck16.Value = False
    NotCheck16.Visible = False
    GoTo Son
ElseIf OgeTuruFrame15.Visible = True Then
    OgeTuru15.Value = ""
    OgeDegeri15.Value = ""
    Adet15.Value = ""
    OgeIdNo15.Value = ""
    Aciklama15.Value = ""
    Sonuc15.Value = ""
    UretimOzelligi15.Value = ""
    RaporOzelligi15.Enabled = False
    RaporOzelligi15.Value = ""
    Rapor1No15.Value = ""
    OgeTuruFrame15.Visible = False
    NotCheck15.Value = False
    NotCheck15.Visible = False
    GoTo Son
ElseIf OgeTuruFrame14.Visible = True Then
    OgeTuru14.Value = ""
    OgeDegeri14.Value = ""
    Adet14.Value = ""
    OgeIdNo14.Value = ""
    Aciklama14.Value = ""
    Sonuc14.Value = ""
    UretimOzelligi14.Value = ""
    RaporOzelligi14.Enabled = False
    RaporOzelligi14.Value = ""
    Rapor1No14.Value = ""
    OgeTuruFrame14.Visible = False
    NotCheck14.Value = False
    NotCheck14.Visible = False
    GoTo Son
ElseIf OgeTuruFrame13.Visible = True Then
    OgeTuru13.Value = ""
    OgeDegeri13.Value = ""
    Adet13.Value = ""
    OgeIdNo13.Value = ""
    Aciklama13.Value = ""
    Sonuc13.Value = ""
    UretimOzelligi13.Value = ""
    RaporOzelligi13.Enabled = False
    RaporOzelligi13.Value = ""
    Rapor1No13.Value = ""
    OgeTuruFrame13.Visible = False
    NotCheck13.Value = False
    NotCheck13.Visible = False
    GoTo Son
ElseIf OgeTuruFrame12.Visible = True Then
    OgeTuru12.Value = ""
    OgeDegeri12.Value = ""
    Adet12.Value = ""
    OgeIdNo12.Value = ""
    Aciklama12.Value = ""
    Sonuc12.Value = ""
    UretimOzelligi12.Value = ""
    RaporOzelligi12.Enabled = False
    RaporOzelligi12.Value = ""
    Rapor1No12.Value = ""
    OgeTuruFrame12.Visible = False
    NotCheck12.Value = False
    NotCheck12.Visible = False
    GoTo Son
ElseIf OgeTuruFrame11.Visible = True Then
    OgeTuru11.Value = ""
    OgeDegeri11.Value = ""
    Adet11.Value = ""
    OgeIdNo11.Value = ""
    Aciklama11.Value = ""
    Sonuc11.Value = ""
    UretimOzelligi11.Value = ""
    RaporOzelligi11.Enabled = False
    RaporOzelligi11.Value = ""
    Rapor1No11.Value = ""
    OgeTuruFrame11.Visible = False
    NotCheck11.Value = False
    NotCheck11.Visible = False
    GoTo Son
ElseIf OgeTuruFrame10.Visible = True Then
    OgeTuru10.Value = ""
    OgeDegeri10.Value = ""
    Adet10.Value = ""
    OgeIdNo10.Value = ""
    Aciklama10.Value = ""
    Sonuc10.Value = ""
    UretimOzelligi10.Value = ""
    RaporOzelligi10.Enabled = False
    RaporOzelligi10.Value = ""
    Rapor1No10.Value = ""
    OgeTuruFrame10.Visible = False
    NotCheck10.Value = False
    NotCheck10.Visible = False
    GoTo Son
ElseIf OgeTuruFrame9.Visible = True Then
    OgeTuru9.Value = ""
    OgeDegeri9.Value = ""
    Adet9.Value = ""
    OgeIdNo9.Value = ""
    Aciklama9.Value = ""
    Sonuc9.Value = ""
    UretimOzelligi9.Value = ""
    RaporOzelligi9.Enabled = False
    RaporOzelligi9.Value = ""
    Rapor1No9.Value = ""
    OgeTuruFrame9.Visible = False
    NotCheck9.Value = False
    NotCheck9.Visible = False
    GoTo Son
ElseIf OgeTuruFrame8.Visible = True Then
    OgeTuru8.Value = ""
    OgeDegeri8.Value = ""
    Adet8.Value = ""
    OgeIdNo8.Value = ""
    Aciklama8.Value = ""
    Sonuc8.Value = ""
    UretimOzelligi8.Value = ""
    RaporOzelligi8.Enabled = False
    RaporOzelligi8.Value = ""
    Rapor1No8.Value = ""
    OgeTuruFrame8.Visible = False
    NotCheck8.Value = False
    NotCheck8.Visible = False
    GoTo Son
ElseIf OgeTuruFrame7.Visible = True Then
    OgeTuru7.Value = ""
    OgeDegeri7.Value = ""
    Adet7.Value = ""
    OgeIdNo7.Value = ""
    Aciklama7.Value = ""
    Sonuc7.Value = ""
    UretimOzelligi7.Value = ""
    RaporOzelligi7.Enabled = False
    RaporOzelligi7.Value = ""
    Rapor1No7.Value = ""
    OgeTuruFrame7.Visible = False
    NotCheck7.Value = False
    NotCheck7.Visible = False
    GoTo Son
ElseIf OgeTuruFrame6.Visible = True Then
    OgeTuru6.Value = ""
    OgeDegeri6.Value = ""
    Adet6.Value = ""
    OgeIdNo6.Value = ""
    Aciklama6.Value = ""
    Sonuc6.Value = ""
    UretimOzelligi6.Value = ""
    RaporOzelligi6.Enabled = False
    RaporOzelligi6.Value = ""
    Rapor1No6.Value = ""
    OgeTuruFrame6.Visible = False
    NotCheck6.Value = False
    NotCheck6.Visible = False
    GoTo Son
ElseIf OgeTuruFrame5.Visible = True Then
    OgeTuru5.Value = ""
    OgeDegeri5.Value = ""
    Adet5.Value = ""
    OgeIdNo5.Value = ""
    Aciklama5.Value = ""
    Sonuc5.Value = ""
    UretimOzelligi5.Value = ""
    RaporOzelligi5.Enabled = False
    RaporOzelligi5.Value = ""
    Rapor1No5.Value = ""
    OgeTuruFrame5.Visible = False
    NotCheck5.Value = False
    NotCheck5.Visible = False
    GoTo Son
ElseIf OgeTuruFrame4.Visible = True Then
    OgeTuru4.Value = ""
    OgeDegeri4.Value = ""
    Adet4.Value = ""
    OgeIdNo4.Value = ""
    Aciklama4.Value = ""
    Sonuc4.Value = ""
    UretimOzelligi4.Value = ""
    RaporOzelligi4.Enabled = False
    RaporOzelligi4.Value = ""
    Rapor1No4.Value = ""
    OgeTuruFrame4.Visible = False
    NotCheck4.Value = False
    NotCheck4.Visible = False
    GoTo Son
ElseIf OgeTuruFrame3.Visible = True Then
    OgeTuru3.Value = ""
    OgeDegeri3.Value = ""
    Adet3.Value = ""
    OgeIdNo3.Value = ""
    Aciklama3.Value = ""
    Sonuc3.Value = ""
    UretimOzelligi3.Value = ""
    RaporOzelligi3.Enabled = False
    RaporOzelligi3.Value = ""
    Rapor1No3.Value = ""
    OgeTuruFrame3.Visible = False
    NotCheck3.Value = False
    NotCheck3.Visible = False
    GoTo Son
ElseIf OgeTuruFrame2.Visible = True Then
    OgeTuru2.Value = ""
    OgeDegeri2.Value = ""
    Adet2.Value = ""
    OgeIdNo2.Value = ""
    Aciklama2.Value = ""
    Sonuc2.Value = ""
    UretimOzelligi2.Value = ""
    RaporOzelligi2.Enabled = False
    RaporOzelligi2.Value = ""
    Rapor1No2.Value = ""
    OgeTuruFrame2.Visible = False
    NotCheck2.Value = False
    NotCheck2.Visible = False
    GoTo Son
ElseIf OgeTuruFrame1.Visible = True Then
    OgeTuru1.Value = ""
    OgeDegeri1.Value = ""
    Adet1.Value = ""
    OgeIdNo1.Value = ""
    Aciklama1.Value = ""
    Sonuc1.Value = ""
    UretimOzelligi1.Value = ""
    RaporOzelligi1.Enabled = False
    RaporOzelligi1.Value = ""
    Rapor1No1.Value = ""
    OgeTuruFrame1.Visible = False
    NotCheck1.Value = False
    NotCheck1.Visible = False
    GoTo Son
End If


Son:

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

If ScrollTakip > 0 Then
    If ScrollTakip = 48 Then
        ScrollTakip = ScrollTakip - 24 - 6
    ElseIf ScrollTakip > 48 Then
        ScrollTakip = ScrollTakip - 24
    Else
        ScrollTakip = ScrollTakip - 24
    End If
End If

If ScrollTakip > 48 + 6 Then
    Call SetScrollHook(Me.ScrollFrame, ScrollTakip, 24)
    ScrollFrame.ScrollTop = 0
    ScrollFrame.ScrollTop = ScrollFrame.ScrollHeight
ElseIf ScrollTakip > 0 And ScrollTakip <= 48 + 6 Then
    ScrollFrame.ScrollTop = 0
    RemoveScrollHook
    ScrollFrame.ScrollBars = fmScrollBarsNone
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Dim yukseklik As Variant
'Dim Rep As Variant

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
        core_report3_2_entry_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_report3_2_entry_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_report3_2_entry_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_report3_2_entry_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_report3_2_entry_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_report3_2_entry_UI.Height = yukseklik
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
Dim a() As Variant, j As Variant

ScrollTakip = 0
Threshold = 54

'Muhatap Temasını uyarla.
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"
ThisWorkbook.Worksheets(2).Range("CW6").Value = ""
ThisWorkbook.Worksheets(2).Range("CW7").Value = "Contact Theme"
ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

Nokta2Option.Value = True
TipAOption.Value = True

Call ComboGetirReset

'FinansalBirim gönderi tipi combosuna öğeleri ekle
With FinansalBirimGonderiTipi
    .AddItem ("Hand Delivery")
    .AddItem ("Cargo")
    .AddItem ("Courier")
    .AddItem ("")
End With

'Nesne renkleri
For Each ClrLab In core_report3_2_entry_UI.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "CheckBox" Then
        ClrLab.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
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


core_report3_2_entry_UI.BackColor = RGB(230, 230, 230) 'YENİ
core_report3_2_entry_UI.UstMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
core_report3_2_entry_UI.AltMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
ComboGetir.BackColor = RGB(225, 235, 245)

FrameTipATipB.BackColor = RGB(225, 235, 245)
TipAOption.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
TipAOption.ForeColor = RGB(30, 30, 30)
TipBOption.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
TipBOption.ForeColor = RGB(30, 30, 30)

Nokta2Option.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Nokta2Option.ForeColor = RGB(30, 30, 30)
Nokta3Option.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Nokta3Option.ForeColor = RGB(30, 30, 30)

Kaydet.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Kaydet.ForeColor = RGB(30, 30, 30)

MaxiMini.BackColor = RGB(225, 235, 245)
MaxiMini.ForeColor = RGB(30, 30, 30)

'Rapor özelliğini ve üretim özelliğini açılışta kapat
'RaporOzelligi.Enabled = False
UretimOzelligi.Enabled = False
For i = 1 To 19
    Controls("RaporOzelligi" & i).Enabled = False
    Controls("UretimOzelligi" & i).Enabled = False
Next i

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

For Each ClrLab In core_report3_2_entry_UI.UstMenuFrame.Controls
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


If TipBOption.Value = True Then
    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        AppXCenter = Application.Left + (Application.Width / 2)
        AppYCenter = Application.Top + (Application.Height / 2)
    
        'Formu önce ekrana ortala
        With core_report3_2_entry_UI
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
        If Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = True Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = False Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + 6  '462 '444 '299
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + 6  '486
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Tutanak2Frame.Height + 6 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            core_report3_2_entry_UI.ScrollTop = 0
            core_report3_2_entry_UI.ScrollHeight = 0
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsNone
            
            core_report3_2_entry_UI.Width = 1024
        
            'Formun görünümü
            AltMenuFrame.Top = 426 '462 '444 '299
            TasiyiciFrame.Height = 448 '486
            Rep = 518 '546 '556 '497 '352
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
        With core_report3_2_entry_UI
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
        If Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = True Then
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
            Rep = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 18
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
            Rep = 518 + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 12
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 + Tutanak2Frame.Height + 6
            TasiyiciFrame.Height = 448 + Tutanak2Frame.Height + 6
            Rep = 518 + Tutanak2Frame.Height + 6
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426
            TasiyiciFrame.Height = 448
            Rep = 518
            'Rep = core_report3_2_entry_UI.Height
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

Else

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then
    
        AppXCenter = Application.Left + (Application.Width / 2)
        AppYCenter = Application.Top + (Application.Height / 2)
    
        'Formu önce ekrana ortala
        With core_report3_2_entry_UI
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
        If Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = True Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = False Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + 12 '462 '444 '299
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + 12 '486
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + 12 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            'Formun görünümü
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + 6  '462 '444 '299
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + 6  '486
            Rep = 518
            
            core_report3_2_entry_UI.Width = 1024 + 12
            
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report3_2_entry_UI.ScrollHeight = 518 + Rapor1Frame.Height + 6 - 30
            core_report3_2_entry_UI.ScrollTop = 0
        ElseIf Rapor1Frame.Visible = False And Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            core_report3_2_entry_UI.ScrollTop = 0
            core_report3_2_entry_UI.ScrollHeight = 0
            core_report3_2_entry_UI.ScrollBars = fmScrollBarsNone
            
            core_report3_2_entry_UI.Width = 1024
        
            'Formun görünümü
            AltMenuFrame.Top = 426 '462 '444 '299
            TasiyiciFrame.Height = 448 '486
            Rep = 518 '546 '556 '497 '352
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
        With core_report3_2_entry_UI
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
        If Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = True Then
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24
            Rep = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + UstYaziFrame.Height + 24
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = True And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18
            Rep = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + FinansalBirimUstYaziFrame.Height + 18
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = True And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
            Rep = 518 + Rapor1Frame.Height + Tutanak2Frame.Height + 12
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Rapor1Frame.Visible = True And Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 + Rapor1Frame.Height + 6
            TasiyiciFrame.Height = 448 + Rapor1Frame.Height + 6
            Rep = 518 + Rapor1Frame.Height + 6
            'Rep = core_report3_2_entry_UI.Height
        ElseIf Rapor1Frame.Visible = False And Tutanak2Frame.Visible = False And FinansalBirimUstYaziFrame.Visible = False And UstYaziFrame.Visible = False Then
            AltMenuFrame.Top = 426 '462 '444 '299
            TasiyiciFrame.Height = 448 '560 '486
            Rep = 518 '584 '556 '497 '352
            'Rep = core_report3_2_entry_UI.Height
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
    
End If

'Modeless modunda userformun mouseover seçeneği yavaşlıyor. Sorun bu şekilde çözüldü.
core_report3_2_entry_UI.Hide
core_report3_2_entry_UI.Show vbModal

End Sub

Sub FormPositionMini()
Dim AppXCenter, AppYCenter As Long
Dim yukseklik As Variant, genislik As Variant
Dim ClrLab As MSForms.Control


For Each ClrLab In core_report3_2_entry_UI.UstMenuFrame.Controls
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
With core_report3_2_entry_UI
    .StartUpPosition = 0
    .Left = Application.Left '+ (0.5 * Application.Width) - (0.5 * 1034)
    .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 560)
End With

'If EkranKontrol = True Then
If core_report3_2_entry_UI.ScrollHeight > 0 Then
    core_report3_2_entry_UI.ScrollTop = 0
    core_report3_2_entry_UI.ScrollHeight = 0
    core_report3_2_entry_UI.ScrollBars = fmScrollBarsNone
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
    core_report3_2_entry_UI.Left = core_report3_2_entry_UI.Left + Application.Left + (0.2 * Application.Width)
    If core_report3_2_entry_UI.Left > Application.Left + (0.9 * Application.Width) Then
        core_report3_2_entry_UI.Left = Application.Left + (0.9 * Application.Width)
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
core_report3_2_entry_UI.Hide
core_report3_2_entry_UI.Show vbModeless

End Sub


Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub





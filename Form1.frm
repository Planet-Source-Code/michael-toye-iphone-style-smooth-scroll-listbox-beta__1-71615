VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   1740
      Width           =   1095
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   1380
      Width           =   1095
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   1020
      Width           =   1095
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   285
      Left            =   5700
      TabIndex        =   2
      Top             =   480
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      Caption         =   "Delete"
      Enabled         =   -1  'True
      BackStyle       =   3
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin Project1.mtListBox mtListBox1 
      Height          =   4575
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   285
      Left            =   5700
      TabIndex        =   3
      Top             =   1740
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      Caption         =   "Add"
      Enabled         =   -1  'True
      BackStyle       =   3
   End
   Begin Project1.mtListBox mtListBox2 
      Height          =   4575
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
   End
   Begin Project1.UserControl1 UserControl13 
      Height          =   300
      Left            =   4020
      TabIndex        =   12
      Top             =   4395
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   529
      Caption         =   "Exit Listbox Demo"
      Enabled         =   -1  'True
      BackStyle       =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   10
      Top             =   540
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Val 2"
      Height          =   195
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Val 1"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()


Randomize Timer





mtListBox1.SortBy = Value1
mtListBox1.Caption = "Currency Table"
mtListBox1.OnlyOneColumn = True


mtListBox1.AddItem "AFA", "Afghanistan Afgani (AFA)"
mtListBox1.AddItem "TZS", "Tanzanian Shilling (TZS)"
mtListBox1.AddItem "THB", "Thai Baht (THB)"
mtListBox1.AddItem "TTD", "Trinidad/tobago Dollar (TTD)"
mtListBox1.AddItem "TND", "Tunisian Dinar (TND)"
mtListBox1.AddItem "TRL", "Turkish Lira (TRL)"
mtListBox1.AddItem "TMS", "Turkmenistan Manat (TMS)"
mtListBox1.AddItem "AED", "Uae Dirham (AED)"
mtListBox1.AddItem "UGS", "Ugandan Shilling (UGS)"
mtListBox1.AddItem "UAK", "Ukraine Hryvna (UAK)"
mtListBox1.AddItem "USD", "United States Dollar (USD)"
mtListBox1.AddItem "UYP", "Uruguay Peso (UYP)"
mtListBox1.AddItem "UZS", "Uzbekistan Sum (UZS)"
mtListBox1.AddItem "VUV", "Vanuatu Vatu (VUV)"
mtListBox1.AddItem "VEB", "Venezuelan Bolivar (VEB)"
mtListBox1.AddItem "VND", "Vietnam Dong (VND)"
mtListBox1.AddItem "YDD", "Yemeni Dinar (YDD)"
mtListBox1.AddItem "YER", "Yemeni Rial (YER)"
mtListBox1.AddItem "ZRZ", "Zaire Zaire (ZRZ)"
mtListBox1.AddItem "ZMK", "Zambian Kwacha (ZMK)"
mtListBox1.AddItem "ZWD", "Zimbabwe Dollar (ZWD)"
mtListBox1.AddItem "MXV", "Mexican Unidad de Inversion (MXV)"
mtListBox1.AddItem "PEN", "Peruvian Nuevo Peso (PEN)"
mtListBox1.AddItem "UYU", "Uruguayan Peso Uruguayano (UYU)"
mtListBox1.AddItem "TRY", "New Turkish Lira"
mtListBox1.AddItem "IDR", "Indonesian Rupiah (IDR)"
mtListBox1.AddItem "IRR", "Iranian Rial (IRR)"
mtListBox1.AddItem "IQD", "Iraqi Dinar (IQD)"
mtListBox1.AddItem "IEP", "Irish Punt (IEP)"
mtListBox1.AddItem "ISS", "Israeli Shekel (ISS)"
mtListBox1.AddItem "ITL", "Italian Lira (ITL)"
mtListBox1.AddItem "JMD", "Jamaica Dollar (JMD)"
mtListBox1.AddItem "JPY", "Japanese Yen (JPY)"
mtListBox1.AddItem "JOD", "Jordanian Dinar (JOD)"
mtListBox1.AddItem "KTS", "Kazakhstan Tenge (KTS)"
mtListBox1.AddItem "KES", "Kenyan Shilling (KES)"
mtListBox1.AddItem "KWD", "Kuwaiti Dinar (KWD)"
mtListBox1.AddItem "KYS", "Kyrgyzstan Som (KYS)"
mtListBox1.AddItem "LAK", "Laos New Kip (LAK)"
mtListBox1.AddItem "LVR", "Latvian Lat (LVR)"
mtListBox1.AddItem "LBP", "Lebanese Pound (LBP)"
mtListBox1.AddItem "LSM", "Lesotho Loti (LSM)"
mtListBox1.AddItem "LRD", "Liberian Dollar (LRD)"
mtListBox1.AddItem "LTT", "Lithuanian Lit (LTT)"
mtListBox1.AddItem "LUF", "Luxembourg Franc (LUF)"
mtListBox1.AddItem "MOP", "Macau Pataca (MOP)"
mtListBox1.AddItem "MWK", "Malawi Kwacha (MWK)"
mtListBox1.AddItem "MYR", "Malaysian Ringgit (MYR)"
mtListBox1.AddItem "MTL", "Maltese Lira (MTL)"
mtListBox1.AddItem "MRO", "Mauritania Ouguiya (MRO)"
mtListBox1.AddItem "MUR", "Mauritius Rupee (MUR)"
mtListBox1.AddItem "MXN", "Mexican Peso (MXN)"
mtListBox1.AddItem "MVS", "Moldova Lei (MVS)"
mtListBox1.AddItem "MNT", "Mongolia Tugrik (MNT)"
mtListBox1.AddItem "MAD", "Moroccan Dirham (MAD)"
mtListBox1.AddItem "MZM", "Mozambique Metical (MZM)"
mtListBox1.AddItem "MMK", "Myanmar Kyat (MMK)"
mtListBox1.AddItem "ANG", "Netherland Antilles Guilder (ANG)"
mtListBox1.AddItem "NZD", "New Zealand Dollar (NZD)"
mtListBox1.AddItem "NIC", "Nicaragua Cordoba (NIC)"
mtListBox1.AddItem "NGN", "Nigeria Naira (NGN)"
mtListBox1.AddItem "NOK", "Norwegian Krone (NOK)"
mtListBox1.AddItem "OMR", "Omani Rial (OMR)"
mtListBox1.AddItem "PKR", "Pakistani Rupee (PKR)"
mtListBox1.AddItem "PAB", "Panamanian Balboa (PAB)"
mtListBox1.AddItem "PGK", "Papua New Guinea Kina (PGK)"
mtListBox1.AddItem "PYG", "Paraguay Guarani (PYG)"
mtListBox1.AddItem "PSS", "Peruvian New Sol (PSS)"
mtListBox1.AddItem "PHP", "Philippines Peso (PHP)"
mtListBox1.AddItem "PLN", "Polish Zloty (PLN)"
mtListBox1.AddItem "PTE", "Portuguese Escudo (PTE)"
mtListBox1.AddItem "QAR", "Qatari Riyal (QAR)"
mtListBox1.AddItem "ROL", "Romanian Leu (ROL)"
mtListBox1.AddItem "RUB", "Russian Ruble (RUB)"
mtListBox1.AddItem "RWS", "Rwanda Franc (RWS)"
mtListBox1.AddItem "STD", "Sao Tome Dobra (STD)"
mtListBox1.AddItem "SAR", "Saudi Riyal (SAR)"
mtListBox1.AddItem "SCR", "Seychelles Rupee (SCR)"
mtListBox1.AddItem "SLL", "Sierra Leone Leone (SLL)"
mtListBox1.AddItem "SGD", "Singapore Dollar (SGD)"
mtListBox1.AddItem "SKK", "Slovakia Koruna (SKK)"
mtListBox1.AddItem "SIT", "Slovenia Tolar (SIT)"
mtListBox1.AddItem "SBD", "Solomon Island Dollar (SBD)"
mtListBox1.AddItem "SOS", "Somali Schilling (SOS)"
mtListBox1.AddItem "ZAR", "South African Rand (ZAR)"
mtListBox1.AddItem "KRW", "South Korean Won (KRW)"
mtListBox1.AddItem "ESP", "Spanish Peseta (ESP)"
mtListBox1.AddItem "LKR", "Sri Lankan Rupee (LKR)"
mtListBox1.AddItem "SHP", "St. Helena Pound (SHP)"
mtListBox1.AddItem "SDD", "Sudanese Pound (SDD)"
mtListBox1.AddItem "SRG", "Surinam Guilder (SRG)"
mtListBox1.AddItem "SZL", "Swaziland Lilangeni (SZL)"
mtListBox1.AddItem "SEK", "Swedish Krona (SEK)"
mtListBox1.AddItem "CHF", "Swiss Franc (CHF)"
mtListBox1.AddItem "SYP", "Syrian Pound (SYP)"
mtListBox1.AddItem "TWD", "Taiwan Dollar (TWD)"
mtListBox1.AddItem "TJS", "Tajikistan Ruble (TJS)"
mtListBox1.AddItem "ALL", "Albanian Lek (ALL)"
mtListBox1.AddItem "DZD", "Algerian Dinar (DZD)"
mtListBox1.AddItem "ADP", "Andorran Peseta (ADP)"
mtListBox1.AddItem "AOK", "Angolan Kwanza (AOK)"
mtListBox1.AddItem "XCD", "Antigua Dollar (XCD)"
mtListBox1.AddItem "ARS", "Argentine Peso (ARS)"
mtListBox1.AddItem "AMD", "Armenia Dram (AMD)"
mtListBox1.AddItem "AUD", "Australian Dollar (AUD)"
mtListBox1.AddItem "ATS", "Austrian Schilling (ATS)"
mtListBox1.AddItem "AZS", "Azerbaijan Manat (AZS)"
mtListBox1.AddItem "BSD", "Bahamas Dollar (BSD)"
mtListBox1.AddItem "BHD", "Bahraini Dinar (BHD)"
mtListBox1.AddItem "BDT", "Bangladesh Taka (BDT)"
mtListBox1.AddItem "BBD", "Barbados Dollar (BBD)"
mtListBox1.AddItem "BES", "Belarus Rouble (BES)"
mtListBox1.AddItem "BEF", "Belgian Franc (BEF)"
mtListBox1.AddItem "BZD", "Belize Dollar (BZD)"
mtListBox1.AddItem "XOF", "Benin Franc (XOF)"
mtListBox1.AddItem "BMD", "Bermudian Dollar (BMD)"
mtListBox1.AddItem "BTN", "Bhutan Ngultrum (BTN)"
mtListBox1.AddItem "BOP", "Bolivian Boliviano (BOP)"
mtListBox1.AddItem "BWP", "Botswana Pula (BWP)"
mtListBox1.AddItem "BRL", "Brazil Real (BRL)"
mtListBox1.AddItem "GBP", "British Pound (GBP)"
mtListBox1.AddItem "BND", "Brunei Dollar (BND)"
mtListBox1.AddItem "BGL", "Bulgarian Lev (BGL)"
mtListBox1.AddItem "XAF", "Cameroon Franc (XAF)"
mtListBox1.AddItem "CAD", "Canadian Dollar (CAD)"
mtListBox1.AddItem "BPS", "Canton & Enderbury Island Pound (BPS)"
mtListBox1.AddItem "CVE", "Cape Verde Escudo (CVE)"
mtListBox1.AddItem "KYD", "Cayman Islands (KYD)"
mtListBox1.AddItem "CLP", "Chilean Peso (CLP)"
mtListBox1.AddItem "CNY", "China Renminbi (CNY)"
mtListBox1.AddItem "COP", "Colombian Peso (COP)"
mtListBox1.AddItem "KMF", "Comoros Franc (KMF)"
mtListBox1.AddItem "CRC", "Costa Rica Colon (CRC)"
mtListBox1.AddItem "HRK", "Croatian Kuna (HRK)"
mtListBox1.AddItem "CUP", "Cuban Peso (CUP)"
mtListBox1.AddItem "CYP", "Cypriot Pound (CYP)"
mtListBox1.AddItem "CZK", "Czech Koruna (CZK)"
mtListBox1.AddItem "DKK", "Danish Krone (DKK)"
mtListBox1.AddItem "DJF", "Djibouti Franc (DJF)"
mtListBox1.AddItem "DOP", "Dominican Republic (DOP)"
mtListBox1.AddItem "NLG", "Dutch Guilder (NLG)"
mtListBox1.AddItem "ESS", "Ecuadoran Sucre (ESS)"
mtListBox1.AddItem "EGP", "Egyptian Pound (EGP)"
mtListBox1.AddItem "SVC", "El Salvador Colon (SVC)"
mtListBox1.AddItem "EEK", "Estonian Kroon (EEK)"
mtListBox1.AddItem "ETB", "Ethiopian Birr (ETB)"
mtListBox1.AddItem "EUR", "Euro (EUR)"
mtListBox1.AddItem "XEU", "European Currency Unit (*XEU)"
mtListBox1.AddItem "FKP", "Falkland Island Pound (FKP)"
mtListBox1.AddItem "FJD", "Fiji Dollar (FJD)"
mtListBox1.AddItem "FIM", "Finnish Markka (FIM)"
mtListBox1.AddItem "FRF", "French Franc (FRF)"
mtListBox1.AddItem "CFP", "French Pacific Island Franc (CFP)"
mtListBox1.AddItem "GMD", "Gambian Dalasi (GMD)"
mtListBox1.AddItem "GEL", "Georgian Lari (GEL)"
mtListBox1.AddItem "DEM", "German Mark (DEM)"
mtListBox1.AddItem "GHC", "Ghana Cedi (GHC)"
mtListBox1.AddItem "GIP", "Gibraltar Pound (GIP)"
mtListBox1.AddItem "GRD", "Greek Drachma (GRD)"
mtListBox1.AddItem "GTQ", "Guatemala Quetzal (GTQ)"
mtListBox1.AddItem "GWP", "Guinea Bissau Peso (GWP)"
mtListBox1.AddItem "GNS", "Guinea Franc (GNS)"
mtListBox1.AddItem "GYD", "Guyana Dollar (GYD)"
mtListBox1.AddItem "HTG", "Haiti Gourde (HTG)"
mtListBox1.AddItem "HNL", "Honduras Lempira (HNL)"
mtListBox1.AddItem "HKD", "Hong Kong Dollar (HKD)"
mtListBox1.AddItem "HUF", "Hungarian Forint (HUF)"
mtListBox1.AddItem "ISK", "Iceland Krona (ISK)"
mtListBox1.AddItem "INR", "Indian Rupee (INR)"


mtListBox2.SortBy = Value1
mtListBox2.Caption = "Currency Table - 2 columns"
mtListBox2.OnlyOneColumn = False
mtListBox2.AddItem "AFA", "Afghanistan Afgani (AFA)", "AFA"
mtListBox2.AddItem "TZS", "Tanzanian Shilling (TZS)", "TZS"
mtListBox2.AddItem "THB", "Thai Baht (THB)", "THB"
mtListBox2.AddItem "TTD", "Trinidad/tobago Dollar (TTD)", "TTD"
mtListBox2.AddItem "TND", "Tunisian Dinar (TND)", "TND"
mtListBox2.AddItem "TRL", "Turkish Lira (TRL)", "TRL"
mtListBox2.AddItem "TMS", "Turkmenistan Manat (TMS)", "TMS"
mtListBox2.AddItem "AED", "Uae Dirham (AED)", "AED"
mtListBox2.AddItem "UGS", "Ugandan Shilling (UGS)", "UGS"
mtListBox2.AddItem "UAK", "Ukraine Hryvna (UAK)", "UAK"
mtListBox2.AddItem "USD", "United States Dollar (USD)", "USD"
mtListBox2.AddItem "UYP", "Uruguay Peso (UYP)", "UYP"
mtListBox2.AddItem "UZS", "Uzbekistan Sum (UZS)", "UZS"

End Sub

Private Sub UserControl11_Click()
    mtListBox1.DeleteItemByKey txtKey.Text
    mtListBox2.DeleteItemByKey txtKey.Text
End Sub

Private Sub UserControl12_Click()
    mtListBox1.AddItem txtAdd(0), txtAdd(1), txtAdd(2)
    mtListBox2.AddItem txtAdd(0), txtAdd(1), txtAdd(2)
End Sub

Private Sub UserControl13_Click()
    Unload Me
    
End Sub

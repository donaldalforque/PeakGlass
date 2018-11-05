VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form INV_ExtraSellingInfoFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAddProduct 
      Caption         =   "Add Product"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pricing Reference"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   7695
      Begin VB.CheckBox chkClipping 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Track Clipping"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkLength 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sold in lengths only"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin VB.Frame FRE_PrimaryProductInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Primary Product Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   7215
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   23
            Text            =   "1"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtWidth 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   22
            Text            =   "1"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtLength 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   21
            Text            =   "1"
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cmbExtraUomInfo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   720
            Width           =   3495
         End
         Begin VB.ComboBox cmbStandardUom 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3000
            TabIndex        =   13
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label lblPrice 
            BackStyle       =   0  'Transparent
            Caption         =   "Price per "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   25
            Top             =   1920
            Width           =   2400
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Unit of Measure"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   20
            Top             =   720
            Width           =   2160
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Length:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3000
            TabIndex        =   19
            Top             =   1200
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5040
            TabIndex        =   18
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4680
            TabIndex        =   17
            Top             =   1200
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Unit of Measure"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   16
            Top             =   360
            Width           =   2355
         End
         Begin VB.Label lblData 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3000
            TabIndex        =   15
            Top             =   1560
            Width           =   3480
         End
      End
      Begin VB.CheckBox chkCutSizePricing 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refer pricing to ""Cut Size Pricing Reference"""
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1050
         Width           =   4455
      End
      Begin VB.CheckBox chkCutSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item is sold with indefinite sizes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.CheckBox chkIrregularCutSizePricing 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prioritize ""Irregular Cut Size Pricing Reference"""
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut Size Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   29
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Reference"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   5400
         MouseIcon       =   "INV_ExtraSellingInfoFrm.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   1875
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblAddExtraSellingInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Reference"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   5400
         MouseIcon       =   "INV_ExtraSellingInfoFrm.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1125
         Width           =   1185
      End
   End
   Begin VB.CheckBox chkTrack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Track Inventory based on other material."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   6720
      Width           =   4455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ExtraSellingInfoFrm.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ExtraSellingInfoFrm.frx":6B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ExtraSellingInfoFrm.frx":D368
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ExtraSellingInfoFrm.frx":13BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ExtraSellingInfoFrm.frx":1A42C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   2805
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save && Close"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Copy"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvInfo 
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   7200
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reference Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "InventoryMaterialId"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Only 1 product material can be added."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4680
      TabIndex        =   24
      Top             =   8400
      Width           =   3300
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   $"INV_ExtraSellingInfoFrm.frx":20C8E
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   7680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Selling Information"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   960
      TabIndex        =   9
      Top             =   765
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "INV_ExtraSellingInfoFrm.frx":20D7E
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   120
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "INV_ExtraSellingInfoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProdExtraInfoId As String
Private DeleteOtherProduct As String

Private Sub btnAddProduct_Click()
    INV_ProductSelectionFrm.Show (1)
    DisplayProducts
End Sub

Private Sub chkCutSize_Click()
    If chkCutSize.value = Checked Then
        chkCutSizePricing.Enabled = True
        chkIrregularCutSizePricing.Enabled = True
        chkLength.Enabled = True
        FRE_PrimaryProductInfo.Enabled = True
    Else
        chkCutSizePricing.Enabled = False
        chkCutSizePricing.value = Unchecked
        chkIrregularCutSizePricing.Enabled = False
        chkIrregularCutSizePricing.value = Unchecked
        chkLength.Enabled = False
        chkLength.value = Unchecked
        FRE_PrimaryProductInfo.Enabled = False
    End If
End Sub

Private Sub chkCutSizePricing_Click()
    If chkCutSizePricing.value = 1 Then
        cmbName.Enabled = True
    Else
        cmbName.Enabled = False
    End If
End Sub

Private Sub chkTrack_Click()
    If chkTrack.value = Checked Then
        btnAddProduct.Enabled = True
        lvInfo.Enabled = True
        lblData.Visible = False
    Else
        btnAddProduct.Enabled = False
        lvInfo.Enabled = False
        lblData.Visible = True
    End If
End Sub

Private Sub cmbExtraUomInfo_Click()
    DisplayData
End Sub

Private Sub Form_Load()
    lvInfo.ColumnHeaders(2).width = lvInfo.width * 0.1875
    lvInfo.ColumnHeaders(3).width = lvInfo.width * 0.5375
    lvInfo.ColumnHeaders(4).width = lvInfo.width * 0.225
    lvInfo.ColumnHeaders(5).width = lvInfo.width * 0
    
    btnAddProduct.Enabled = False
    lvInfo.Enabled = False
    chkCutSizePricing.Enabled = False
    ProdExtraInfoId = 0
    
    cmbStandardUom.text = INV_NewProductFrm.cmbStorageInfo_Uom.text
    
    Populate "MetricUnit"
    Populate "CutSizePricing"
    Populate "ExtraInfo"
    
    
    DeleteOtherProduct = "0"
End Sub
Private Sub DisplayData()
    lblData.Caption = "1 " & cmbStandardUom.text & " = " & FormatNumber((NVAL(txtLength.text) * NVAL(txtWidth.text)), 2, vbTrue, vbFalse) & " sq." & cmbExtraUomInfo.text
    lblPrice.Caption = "Price per " & cmbExtraUomInfo.text
End Sub
Private Sub Populate(ByVal data As String)
    Select Case data
        Case "MetricUnit"
            Set rec = Global_Data("MetricUnit")
            cmbExtraUomInfo.Clear
            cmbExtraUomInfo.AddItem ""
            cmbExtraUomInfo.ItemData(cmbExtraUomInfo.NewIndex) = 0
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    'If rec!isActive = "True" Then
                        cmbExtraUomInfo.AddItem rec!Name
                        cmbExtraUomInfo.ItemData(cmbExtraUomInfo.NewIndex) = rec!MetricId
                    'End If
                    rec.MoveNext
                Loop
            End If
            cmbExtraUomInfo.ListIndex = 0
        Case "ExtraInfo"
            Dim con As New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandText = "INV_ProductExtraSellingInfo_Get"
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(INV_NewProductFrm.ProductId))
            Set rec = cmd.Execute
            If Not rec.EOF Then
                ProdExtraInfoId = rec!ProdExtraInfoId
                If rec!isIndefiniteSize = "False" Then
                    chkCutSize.value = Unchecked
                Else
                    chkCutSize.value = Checked
                End If
                
                If rec!isLengthOnly = "False" Then
                    chkLength.value = Unchecked
                Else
                    chkLength.value = Checked
                End If
                
                If rec!isPricingReference = "False" Then
                    chkCutSizePricing.value = Unchecked
                Else
                    chkCutSizePricing.value = Checked
                End If
                
                If rec!isIrregularPricingReference = "False" Then
                    chkIrregularCutSizePricing.value = Unchecked
                Else
                    chkIrregularCutSizePricing.value = Checked
                End If
                
                If rec!TrackClipping = "False" Then
                    chkClipping.value = Unchecked
                Else
                    chkClipping.value = Checked
                End If
                
                On Error Resume Next
                cmbName.text = rec!cutsizename
                txtLength.text = rec!length
                txtWidth.text = rec!width
                txtPrice.text = FormatNumber(rec!extrainfoprice, 2, vbTrue, vbFalse)
                If rec!isTrackOnOtherMaterial = "False" Then
                    chkTrack.value = Unchecked
                Else
                    chkTrack.value = Checked
                End If
                
                Dim item As MSComctlLib.ListItem
                lvInfo.ListItems.Clear
                If IsNull(rec!OtherMaterialProductId) = False Then
                    Set item = lvInfo.ListItems.add(, , rec!OtherMaterialProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = rec!OtherProductMetricName
                End If
                
                On Error Resume Next
                cmbExtraUomInfo.text = rec!MetricName
            Else
                chkCutSize.value = Unchecked
                chkCutSizePricing.value = Unchecked
                chkIrregularCutSizePricing.value = Unchecked
                txtLength.text = "1"
                txtWidth.text = "1"
                txtPrice.text = "1"
                chkTrack.value = Unchecked
                lvInfo.ListItems.Clear
                cmbExtraUomInfo.ListIndex = 0
            End If
            con.Close
        Case "CutSizePricing"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("CutSizePricing")
            cmbName.Clear
            cmbName.AddItem ""
            cmbName.ItemData(cmbName.NewIndex) = 0
            cmbName.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbName.AddItem rec!Name
                        cmbName.ItemData(cmbName.NewIndex) = rec!CutSizePricingId
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub


Private Sub Label5_Click()
    BASE_IrregularCutSizePricingReferenceFrm.Show (1)
End Sub

Private Sub lblAddExtraSellingInfo_Click()
    BASE_CutSizePricingReferenceFrm.Show (1)
End Sub

Private Sub lvInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If lvInfo.ListItems.Count > 0 Then
                lvInfo.ListItems.Clear
                chkTrack.value = Unchecked
            End If
    End Select
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'save
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@ProdExtraInfoId", adInteger, adParamInputOutput, , Val(ProdExtraInfoId))
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(INV_NewProductFrm.ProductId))
            cmd.Parameters.Append cmd.CreateParameter("@isIndefiniteSize", adBoolean, adParamInput, , Val(chkCutSize.value))
            cmd.Parameters.Append cmd.CreateParameter("@isLengthOnly", adBoolean, adParamInput, , Val(chkLength.value))
            cmd.Parameters.Append cmd.CreateParameter("@isPricingReference", adBoolean, adParamInput, , Val(chkCutSizePricing.value))
            cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingId", adInteger, adParamInput, , cmbName.ItemData(cmbName.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@isIrregularPricingReference", adBoolean, adParamInput, , Val(chkIrregularCutSizePricing.value))
            cmd.Parameters.Append cmd.CreateParameter("@SellingUomId", adInteger, adParamInput, , cmbExtraUomInfo.ItemData(cmbExtraUomInfo.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@Length", adDecimal, adParamInput, , NVAL(txtLength.text))
                                  cmd.Parameters("@Length").NumericScale = 2
                                  cmd.Parameters("@Length").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Width", adDecimal, adParamInput, , NVAL(txtWidth.text))
                                  cmd.Parameters("@Width").NumericScale = 2
                                  cmd.Parameters("@Width").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@ExtraInfoPrice", adDecimal, adParamInput, , NVAL(txtPrice.text))
                                  cmd.Parameters("@ExtraInfoPrice").NumericScale = 2
                                  cmd.Parameters("@ExtraInfoPrice").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@isTrackOnOtherMaterial", adBoolean, adParamInput, , Val(chkTrack.value))
            If lvInfo.ListItems.Count > 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(lvInfo.ListItems(1).text))
            Else
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Null)
            End If
            cmd.Parameters.Append cmd.CreateParameter("@TrackClipping", adBoolean, adParamInput, , Val(chkClipping.value))
                
            If Val(ProdExtraInfoId) = 0 Then
                cmd.CommandText = "INV_ProductExtraSellingInfo_Insert"
                cmd.Execute
                ProdExtraInfoId = cmd.Parameters("@ProdExtraInfoId")
            Else
                cmd.CommandText = "INV_ProductExtraSellingInfo_Update"
                cmd.Execute
            End If
            
            con.Close
            Unload Me
    End Select
End Sub

Private Sub txtLength_Change()
    If IsNumeric(txtLength.text) = False Then
        txtLength.text = "1"
        DisplayData
    Else
        DisplayData
    End If
End Sub

Private Sub txtLength_GotFocus()
    selectText txtLength
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.text) = False Then
        txtPrice.text = "1.00"
        selectText txtPrice
    End If
End Sub

Private Sub txtPrice_GotFocus()
    selectText txtPrice
End Sub

Private Sub txtWidth_Change()
    If IsNumeric(txtWidth.text) = False Then
        txtWidth.text = "1"
        DisplayData
    Else
        DisplayData
    End If
End Sub

Private Sub txtWidth_GotFocus()
    selectText txtWidth
End Sub

Public Sub DisplayProducts()
    'On Error Resume Next
    
    Dim exists As Boolean
    Dim item As MSComctlLib.ListItem
    
    If ProductSet.RecordCount <= 0 Then Exit Sub

    'Dim item As MSComctlLib.ListItem
    If Not ProductSet.EOF Then
        ProductSet.MoveFirst
        Do Until ProductSet.EOF
            For Each item In lvInfo.ListItems
                If item.text = ProductSet!ProductId Then
                    exists = True
                    Exit Sub
                End If
            Next

            If exists = False Then
                Set item = lvInfo.ListItems.add(, , ProductSet!ProductId)
                item.SubItems(1) = ProductSet!itemcode
                item.SubItems(2) = ProductSet!Name
                Exit Sub
            End If
            ProductSet.MoveNext
        Loop
    End If
End Sub

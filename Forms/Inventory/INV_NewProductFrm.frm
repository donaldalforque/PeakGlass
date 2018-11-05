VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form INV_NewProductFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Product"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14400
      Top             =   0
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
            Picture         =   "INV_NewProductFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":1A188
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   30
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtSortField 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Frame FRAME_ProductDetails3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   6120
         TabIndex        =   57
         Top             =   4440
         Width           =   4215
         Begin VB.ComboBox cmbTaxInfo_Tax 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax"
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
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   58
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.TextBox txtSalesInfoSRPMarkUp 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         TabIndex        =   10
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoDPMarkUp 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   12
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoMSMarkUp 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoBCMarkUp 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   16
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBasicInfo_Barcode 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtSalesInfo_BCP 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   15
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSalesInfo_SP 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   13
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSalesInfo_DP 
         Alignment       =   1  'Right Justify
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
         Left            =   -9999
         TabIndex        =   11
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame_ProductDetails2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3735
         Left            =   6120
         TabIndex        =   43
         Top             =   720
         Width           =   4215
         Begin VB.TextBox txtReorderQuantity 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            TabIndex        =   20
            Top             =   3000
            Width           =   2535
         End
         Begin VB.TextBox txtReorderPoint 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            TabIndex        =   19
            Top             =   2640
            Width           =   2535
         End
         Begin VB.ComboBox cmbVendor 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   840
            Width           =   2535
         End
         Begin VB.ComboBox cmbStorageInfo_Uom 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtCostingInfo_AverageCost 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            TabIndex        =   7
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reorder Qty"
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
            Left            =   120
            TabIndex        =   63
            Top             =   3000
            Width           =   1125
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reorder Point"
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
            Left            =   120
            TabIndex        =   62
            Top             =   2640
            Width           =   1290
         End
         Begin VB.Label lblShowConversion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show Product Conversion"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "INV_NewProductFrm.frx":209EA
            MousePointer    =   99  'Custom
            TabIndex        =   61
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
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
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lblStorageInfoTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Storage Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   50
            Top             =   1800
            Width           =   1425
         End
         Begin VB.Label lblUom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit of Measure"
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
            Left            =   120
            TabIndex        =   49
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costing Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Cost"
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
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1155
         End
      End
      Begin VB.TextBox txtSalesInfo_Price 
         Alignment       =   1  'Right Justify
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
         Left            =   3050
         TabIndex        =   9
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Frame Frame_ProductDetails1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   37
         Top             =   5640
         Width           =   10095
         Begin MSComctlLib.ListView lvInventory 
            Height          =   1815
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3201
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
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "InventoryId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Location"
               Object.Width           =   6115
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Sublocation"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Quantity"
               Object.Width           =   2806
            EndProperty
         End
         Begin VB.Label lblInventory_QtyOnHand 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   4320
            TabIndex        =   67
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   $"INV_NewProductFrm.frx":20B3C
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1575
            Left            =   5880
            TabIndex        =   66
            Top             =   720
            Width           =   4080
         End
         Begin VB.Label lblAddExtraSellingInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add Extra Selling Info"
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
            Left            =   8340
            MouseIcon       =   "INV_NewProductFrm.frx":20BF7
            MousePointer    =   99  'Custom
            TabIndex        =   65
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extra Selling Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   5880
            TabIndex        =   64
            Top             =   0
            Width           =   1920
         End
         Begin VB.Label lblInventory_MoreLocations 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specify Locations"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   4320
            MouseIcon       =   "INV_NewProductFrm.frx":20D49
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity on Hand:"
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
            Left            =   240
            TabIndex        =   39
            Top             =   2400
            Width           =   1680
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.ComboBox cmbBasicInfo_Type 
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
         ItemData        =   "INV_NewProductFrm.frx":20E9B
         Left            =   2280
         List            =   "INV_NewProductFrm.frx":20E9D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtBasicInfo_ItemCode 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtBasicInfo_Name 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox cmbBasicInfo_Category 
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
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   582
         ButtonWidth     =   1588
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Field"
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
         TabIndex        =   68
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label lblShowMorePrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show more Prices"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4500
         MouseIcon       =   "INV_NewProductFrm.frx":20E9F
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   4320
         Width           =   1245
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark-Up (%)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4740
         TabIndex        =   56
         Top             =   3600
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
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
         TabIndex        =   54
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblDiscount3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mega Discount"
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
         Left            =   -9999
         TabIndex        =   53
         Top             =   5040
         Width           =   1365
      End
      Begin VB.Label lblDiscount2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Special Discount"
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
         Left            =   -9999
         TabIndex        =   52
         Top             =   4680
         Width           =   1515
      End
      Begin VB.Label lblDiscount1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Discount"
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
         Left            =   -9999
         TabIndex        =   51
         Top             =   4320
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   42
         Top             =   3360
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Retail Price"
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
         TabIndex        =   41
         Top             =   3960
         Width           =   1920
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         TabIndex        =   36
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         TabIndex        =   34
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   33
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   32
         Top             =   2280
         Width           =   825
      End
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox cmbSearch_Status 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1580
         Width           =   3015
      End
      Begin VB.TextBox txtSearch_ItemCode 
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
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmbSearch_Category 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtSearch_Name 
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
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Refresh"
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
         Left            =   3240
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6375
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   11245
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Category"
            Object.Width           =   2893
         EndProperty
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   240
         TabIndex        =   48
         Top             =   1580
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item code"
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
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "INV_NewProductFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ProductId As Long
Public isService, isInsert, isActive As Boolean
Dim deleteCtr(10000) As Integer
Dim ctr As Integer
Private Sub Initialize()
    Dim txtControl As Control
    isService = False
    ProductId = 0
    ctr = 0
    isActive = True
    isActivated (True)
    
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Name" Then
            txtControl.text = ""
        End If
    Next
    
    txtBasicInfo_ItemCode.BackColor = &HC0C0FF
    txtBasicInfo_Name.BackColor = &HC0C0FF
    
    cmbBasicInfo_Category.ListIndex = 0
    cmbSearch_Category.ListIndex = 0
    cmbBasicInfo_Type.ListIndex = 0
    cmbStorageInfo_Uom.ListIndex = 0
    cmbVendor.ListIndex = 0
    
    tb_Standard.Buttons(4).Image = 3
    tb_Standard.Buttons(4).Caption = "Delete"
    
    lvInventory.ListItems.Clear
    Populate ("Location")
    
    On Error Resume Next
    txtBasicInfo_ItemCode.SetFocus
    CountQuantity
End Sub
Public Sub CountQuantity()
    Dim item As MSComctlLib.ListItem
    Dim Total As Double
    For Each item In lvInventory.ListItems
        If item.SubItems(1) < 4 Then 'SHOULD HAVE LIKE <...
            Total = Total + Val(Replace(item.SubItems(5), ",", ""))
        End If
    Next
    lblInventory_QtyOnHand.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub
Private Sub isActivated(value As Boolean)
    'BASIC INFO
    txtBasicInfo_ItemCode.Enabled = value
    txtBasicInfo_Name.Enabled = value
    txtBasicInfo_Barcode.Enabled = value
    cmbBasicInfo_Category.Enabled = value
    cmbBasicInfo_Type.Enabled = value
    
    'SALES INFO
    txtSalesInfo_Price.Enabled = value
    txtSalesInfoSRPMarkUp.Enabled = value
    txtSalesInfo_DP.Enabled = value
    txtSalesInfoDPMarkUp.Enabled = value
    txtSalesInfo_SP.Enabled = value
    txtSalesInfoMSMarkUp.Enabled = value
    txtSalesInfo_BCP.Enabled = value
    txtSalesInfoBCMarkUp.Enabled = value
    
    'INVENTORY INFO
    lvInventory.Enabled = value
    
    'COSTING INFO
    txtCostingInfo_AverageCost.Enabled = value
    cmbVendor.Enabled = value
    
    'STORAGE INFO
    cmbStorageInfo_Uom.Enabled = value
End Sub

Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Category"
            Set rec = Global_Data("Category")
            cmbBasicInfo_Category.Clear
            cmbSearch_Category.Clear
            cmbSearch_Category.AddItem ""
            cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbSearch_Category.AddItem rec!Category
                        cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = rec!CategoryId
                        
                        cmbBasicInfo_Category.AddItem rec!Category
                        cmbBasicInfo_Category.ItemData(cmbBasicInfo_Category.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "Vendor"
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            cmbVendor.ItemData(cmbVendor.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbVendor.AddItem rec!Name
                    cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
        Case "Status"
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = -1
            cmbSearch_Status.AddItem "Active"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 1
            cmbSearch_Status.AddItem "Deactivated"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 1
        Case "Type"
            Set rec = Global_Data("Type")
            cmbBasicInfo_Type.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbBasicInfo_Type.AddItem rec!Type
                    cmbBasicInfo_Type.ItemData(cmbBasicInfo_Type.NewIndex) = rec!TypeId
                    rec.MoveNext
                Loop
            End If
        Case "Uom"
            Set rec = Global_Data("Uom")
            cmbStorageInfo_Uom.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbStorageInfo_Uom.AddItem rec!Uom
                        cmbStorageInfo_Uom.ItemData(cmbStorageInfo_Uom.NewIndex) = rec!UomId
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "Tax"
            Set rec = Global_Data("Tax")
            cmbTaxInfo_Tax.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbTaxInfo_Tax.AddItem rec!TaxName
                        cmbTaxInfo_Tax.ItemData(cmbTaxInfo_Tax.NewIndex) = rec!TaxId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbTaxInfo_Tax.ListIndex = 1
        Case "Location"
            Set rec = Global_Data("Location")
            Dim item As MSComctlLib.ListItem
            lvInventory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!LocationId = 1 Then
                        Set item = lvInventory.ListItems.add(, , "")
                            item.SubItems(1) = rec!LocationId 'LocationId
                            item.SubItems(3) = rec!Location 'Location
                            item.SubItems(5) = "0.00" 'Quantity
                        Exit Do
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "Product"
            Set rec = Global_Data("Product")
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvSearch.ListItems.add(, , rec!ProductId)
                            item.SubItems(1) = rec!itemcode
                            item.SubItems(2) = rec!Name
                            item.SubItems(3) = rec!Category
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "ProductSelect"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Product_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                txtBasicInfo_ItemCode.text = rec!itemcode
                txtBasicInfo_Name.text = rec!Name
                txtBasicInfo_Barcode.text = rec!Barcode
                cmbBasicInfo_Category.text = rec!Category
                cmbBasicInfo_Type.text = rec!Type
                cmbTaxInfo_Tax.text = rec!TaxName
                txtSortField.text = rec!SortField
                On Error Resume Next
                cmbStorageInfo_Uom.text = rec!Uom
                txtReorderPoint.text = FormatNumber(rec!reorderpoint, 2, vbTrue, vbFalse)
                txtReorderQuantity.text = FormatNumber(rec!ReorderQuantity, 2, vbTrue, vbFalse)
                If IsNull(rec!VendorId) Then
                    cmbVendor.ListIndex = 0
                Else
                    On Error Resume Next
                    cmbVendor.text = rec!Vendor
                End If
                If IsNull(rec!UnitPriceMarkUp) = True Then txtSalesInfoSRPMarkUp.text = "" Else txtSalesInfoSRPMarkUp.text = rec!UnitPriceMarkUp
                If IsNull(rec!Price1MarkUp) = True Then txtSalesInfoDPMarkUp.text = "" Else txtSalesInfoDPMarkUp.text = rec!Price1MarkUp
                If IsNull(rec!Price2MarkUp) = True Then txtSalesInfoMSMarkUp.text = "" Else txtSalesInfoMSMarkUp.text = rec!Price2MarkUp
                If IsNull(rec!Price3MarkUp) = True Then txtSalesInfoBCMarkUp.text = "" Else txtSalesInfoBCMarkUp.text = rec!Price3MarkUp
                txtSalesInfo_Price.text = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                txtSalesInfo_DP.text = FormatNumber(rec!price1, 2, vbTrue, vbFalse)
                txtSalesInfo_SP.text = FormatNumber(rec!price2, 2, vbTrue, vbFalse)
                txtSalesInfo_BCP.text = FormatNumber(rec!price3, 2, vbTrue, vbFalse)
                txtCostingInfo_AverageCost.text = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
                
            End If
            isActive = rec!isActive
            If rec!isActive = "False" Then
                tb_Standard.Buttons(4).Caption = "Activate"
                tb_Standard.Buttons(4).Image = 4
                isActivated (False)
            Else
                tb_Standard.Buttons(4).Caption = "Delete"
                tb_Standard.Buttons(4).Image = 3
                isActivated (True)
            End If
            con.Close
        Case "InventoryLoad"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Inventory_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
            Set rec = cmd.Execute
            lvInventory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvInventory.ListItems.add(, , rec!inventoryId)
                        item.SubItems(1) = rec!LocationId
                        item.SubItems(2) = rec!ProductId
                        item.SubItems(3) = rec!Location
                        item.SubItems(5) = FormatNumber(rec!Quantity, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Private Function ValidateData() As Boolean
    ValidateData = False
    
    'CHECK EMPTY FIELDS
    If Trim(txtSalesInfo_Price.text) = "" Then txtSalesInfo_Price.text = "0.00"
    If Trim(txtSalesInfo_DP.text) = "" Then txtSalesInfo_DP.text = "0.00"
    If Trim(txtSalesInfo_SP.text) = "" Then txtSalesInfo_SP.text = "0.00"
    If Trim(txtSalesInfo_BCP.text) = "" Then txtSalesInfo_BCP.text = "0.00"
    
    If Trim(txtCostingInfo_AverageCost.text) = "" Then txtCostingInfo_AverageCost.text = "0.00"
    
    If Trim(txtBasicInfo_ItemCode.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(1)
        GLOBAL_MessageFrm.Show (1)
        txtBasicInfo_ItemCode.SetFocus
    ElseIf Trim(txtBasicInfo_Name.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(2)
        GLOBAL_MessageFrm.Show (1)
        txtBasicInfo_Name.SetFocus
    ElseIf cmbBasicInfo_Category.ListIndex = -1 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(6)
        GLOBAL_MessageFrm.Show (1)
        cmbBasicInfo_Category.SetFocus
    ElseIf IsNumeric(txtSalesInfo_Price.text) = False And Trim(txtSalesInfo_Price.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_Price.SetFocus
    ElseIf IsNumeric(txtSalesInfoSRPMarkUp.text) = False And Trim(txtSalesInfoSRPMarkUp.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoSRPMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_DP.text) = False And Trim(txtSalesInfo_DP.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_DP.SetFocus
    ElseIf IsNumeric(txtSalesInfoDPMarkUp.text) = False And Trim(txtSalesInfoDPMarkUp.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoDPMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_SP.text) = False And Trim(txtSalesInfo_SP.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_SP.SetFocus
    ElseIf IsNumeric(txtSalesInfoMSMarkUp.text) = False And Trim(txtSalesInfoMSMarkUp.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoMSMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_BCP.text) = False And Trim(txtSalesInfo_BCP.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_BCP.SetFocus
    ElseIf IsNumeric(txtSalesInfoBCMarkUp.text) = False And Trim(txtSalesInfoBCMarkUp.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoBCMarkUp.SetFocus
    ElseIf Trim(cmbStorageInfo_Uom.text) = "" And isService = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(10)
        GLOBAL_MessageFrm.Show (1)
        cmbStorageInfo_Uom.SetFocus
    ElseIf IsNumeric(txtCostingInfo_AverageCost.text) = False And Trim(txtCostingInfo_AverageCost.text) <> "" And isService = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(9)
        GLOBAL_MessageFrm.Show (1)
        txtCostingInfo_AverageCost.SetFocus
    ElseIf IsNumeric(txtReorderPoint.text) = False And Trim(txtReorderPoint.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(72)
        GLOBAL_MessageFrm.Show (1)
        txtReorderPoint.SetFocus
    ElseIf IsNumeric(txtReorderQuantity.text) = False And Trim(txtReorderQuantity.text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(73)
        GLOBAL_MessageFrm.Show (1)
        txtReorderQuantity.SetFocus
    ElseIf Trim(cmbTaxInfo_Tax.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(50)
        GLOBAL_MessageFrm.Show (1)
        cmbTaxInfo_Tax.SetFocus
    Else
        ValidateData = True
    End If
End Function

Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_Name.text)
    'If Trim(txtSearch_ItemCode.text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtSearch_ItemCode.text)
'    Else
'        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, Null)
'    End If
    If cmbSearch_Category.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbSearch_Category.ItemData(cmbSearch_Category.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , Null)
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If

    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Category
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
    BASE_ContainerFrm.statusBar_Main.Panels(1).text = "Total Items: " & lvSearch.ListItems.Count
End Sub

Private Sub cmbBasicInfo_Type_Click()
    If cmbBasicInfo_Type.ListIndex <> 0 Then
        Frame_ProductDetails1.Visible = False
        'Frame_ProductDetails2.Visible = False
        isService = True
    Else
        Frame_ProductDetails1.Visible = True
        Frame_ProductDetails2.Visible = True
        isService = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyS
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(2)
            End If
        Case vbKeyN
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(1)
            End If
        Case vbKeyD
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(4)
            End If
        Case vbKeyC
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(6)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Populate ("Category")
    Populate ("Type")
    Populate ("Uom")
    Populate ("MetricUnit")
    Populate ("Location")
    Populate ("Product")
    Populate ("Status")
    Populate ("Vendor")
    Populate ("Tax")
    Initialize
    btnSearch_Click
    'StatusBarWidth Me, statusBar_Main
End Sub

Private Sub Label26_Click()
    
End Sub

Private Sub lblAddExtraSellingInfo_Click()
    If ProductId = 0 Then Exit Sub
    INV_ExtraSellingInfoFrm.Show (1)
End Sub

Private Sub lblInventory_MoreLocations_Click()
    CenterChildForm INV_LocationFrm
    INV_LocationFrm.Show
End Sub

Private Sub lblShowConversion_Click()
    If ProductId <> 0 Then
        INV_UomConversionFrm.Show (1)
    End If
End Sub

Private Sub lblShowMorePrice_Click()
    If ProductId <> 0 Then
        INV_UomPricingFrm.Show (1)
    End If
End Sub

Private Sub lvInventory_DblClick()
    With lvInventory
        If .ListItems.Count > 0 Then
            Dim i As String
            i = InputBox("Input quantity.", "Quantity", lvInventory.SelectedItem.SubItems(5))
            If i = "" Then
                Exit Sub
            ElseIf IsNumeric(i) = False Then
                Exit Sub
            Else
                .SelectedItem.SubItems(5) = FormatNumber(i, 2, vbFalse, vbFalse)
                .SetFocus
                CountQuantity
            End If
        End If
    End With
End Sub

Private Sub lvInventory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        If lvInventory.ListItems.Count > 0 Then
            If lvInventory.SelectedItem.SubItems(1) <> "1" Then 'NOT Default Location
                If lvInventory.SelectedItem.text <> "" Then 'Existing data
                        deleteCtr(ctr) = Val(lvInventory.SelectedItem.text)
                        ctr = ctr + 1
                        lvInventory.ListItems.Remove (lvInventory.SelectedItem.Index)
                Else
                    lvInventory.ListItems.Remove (lvInventory.SelectedItem.Index)
                End If
            End If
        End If
    Case 13
        Call lvInventory_DblClick
    End Select
End Sub

Public Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    With lvSearch
        If .ListItems.Count > 0 Then
            ProductId = .SelectedItem.text
            Populate "ProductSelect"
            Populate "InventoryLoad"
            CountQuantity
            isInsert = False
        End If
    End With
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim item As MSComctlLib.ListItem
    Select Case Button.Index
        Case 1 'New
            Initialize
            BASE_ContainerFrm.statusBar_Main.Panels(1).text = MessageCodes(6) & " " & MessageCodes(1)
        Case 2 'Save
            If isActive = False Then Exit Sub
            If ValidateData = True Then
                If UserId <> 1 Then
                    POS_UserPinFrm.Show (1)
                Else
                    AllowAccess = True
                End If
                If AllowAccess = False Then Exit Sub
                On Error GoTo ErrHandler
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                
                'SAVE MAIN PRODUCT DETAILS
                con.ConnectionString = ConnString
                con.Open
                con.BeginTrans
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                'Price Mark Up variables
                Dim UnitPriceMarkUp, Price1MarkUp, Price2MarkUp, Price3MarkUp As Variant
                If Trim(txtSalesInfoSRPMarkUp.text) = "" Then UnitPriceMarkUp = Null Else UnitPriceMarkUp = txtSalesInfoSRPMarkUp.text
                If Trim(txtSalesInfoDPMarkUp.text) = "" Then Price1MarkUp = Null Else Price1MarkUp = txtSalesInfoDPMarkUp.text
                If Trim(txtSalesInfoMSMarkUp.text) = "" Then Price2MarkUp = Null Else Price2MarkUp = txtSalesInfoMSMarkUp.text
                If Trim(txtSalesInfoBCMarkUp.text) = "" Then Price3MarkUp = Null Else Price3MarkUp = txtSalesInfoBCMarkUp.text
                
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInputOutput, , ProductId)
                cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtBasicInfo_ItemCode.text)
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtBasicInfo_Name.text)
                cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 50, txtBasicInfo_Barcode.text)
                cmd.Parameters.Append cmd.CreateParameter("@SortField", adVarChar, adParamInput, 50, txtSortField.text)
                cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbBasicInfo_Category.ItemData(cmbBasicInfo_Category.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , cmbBasicInfo_Type.ItemData(cmbBasicInfo_Type.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@TaxId", adInteger, adParamInput, , cmbTaxInfo_Tax.ItemData(cmbTaxInfo_Tax.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@UnitPrice", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_Price.text, ",", "")))
                                      cmd.Parameters("@UnitPrice").Precision = 18
                                      cmd.Parameters("@UnitPrice").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price1", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_DP.text, ",", "")))
                                      cmd.Parameters("@Price1").Precision = 18
                                      cmd.Parameters("@Price1").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price2", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_SP.text, ",", "")))
                                      cmd.Parameters("@Price2").Precision = 18
                                      cmd.Parameters("@Price2").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price3", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_BCP.text, ",", "")))
                                      cmd.Parameters("@Price3").Precision = 18
                                      cmd.Parameters("@Price3").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitPriceMarkUp", adDecimal, adParamInput, , UnitPriceMarkUp)
                                      cmd.Parameters("@UnitPriceMarkUp").Precision = 18
                                      cmd.Parameters("@UnitPriceMarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price1MarkUp", adDecimal, adParamInput, , Price1MarkUp)
                                      cmd.Parameters("@Price1MarkUp").Precision = 18
                                      cmd.Parameters("@Price1MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price2MarkUp", adDecimal, adParamInput, , Price2MarkUp)
                                      cmd.Parameters("@Price2MarkUp").Precision = 18
                                      cmd.Parameters("@Price2MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price3MarkUp", adDecimal, adParamInput, , Price3MarkUp)
                                      cmd.Parameters("@Price3MarkUp").Precision = 18
                                      cmd.Parameters("@Price3MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , Val(Replace(txtCostingInfo_AverageCost.text, ",", "")))
                                      cmd.Parameters("@UnitCost").Precision = 18
                                      cmd.Parameters("@UnitCost").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ReorderPoint", adDecimal, adParamInput, , Val(Replace(txtReorderPoint.text, ",", "")))
                                      cmd.Parameters("@ReorderPoint").Precision = 18
                                      cmd.Parameters("@ReorderPoint").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ReorderQuantity", adDecimal, adParamInput, , Val(Replace(txtReorderQuantity.text, ",", "")))
                                      cmd.Parameters("@ReorderQuantity").Precision = 18
                                      cmd.Parameters("@ReorderQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, cmbStorageInfo_Uom.text)
                cmd.Parameters.Append cmd.CreateParameter("@MetricId", adVarChar, adParamInput, 50, 0)
                If cmbVendor.ListIndex = 0 Then
                    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , Null)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbVendor.ItemData(cmbVendor.ListIndex))
                End If
                         
                If ProductId = 0 Then
                    cmd.CommandText = "BASE_Product_Insert"
                    cmd.Execute
                    ProductId = cmd.Parameters("@ProductId")
                    isInsert = True
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                    cmd.CommandText = "BASE_Product_Update"
                    cmd.Execute
                    isInsert = False
                End If
                
                
                'UNIT OF MEASURE
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, cmbStorageInfo_Uom.text)
                cmd.CommandText = "BASE_Uom_Insert"
                cmd.Execute
                
                'INV_UomPricing/Conversion
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.Parameters.Append cmd.CreateParameter("@UomConversionId", adInteger, adParamInputOutput, , 0)
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                cmd.Parameters.Append cmd.CreateParameter("@UomId", adInteger, adParamInput, , cmbStorageInfo_Uom.ItemData(cmbStorageInfo_Uom.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@ToQty", adDecimal, adParamInput, , Null)
                                      cmd.Parameters("@ToQty").NumericScale = 2
                                      cmd.Parameters("@ToQty").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_Price.text, ",", "")))
                                      cmd.Parameters("@Price").NumericScale = 2
                                      cmd.Parameters("@Price").Precision = 18
                
                cmd.CommandText = "INV_UomConversion_Insert"
                cmd.Execute
                
                
                'INVENTORY
                For Each item In lvInventory.ListItems
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    
                    cmd.Parameters.Append cmd.CreateParameter("@InventoryId", adInteger, adParamInputOutput, , Val(item.text))
                    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Val(item.SubItems(1)))
                    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                                          cmd.Parameters("@Quantity").Precision = 18
                                          cmd.Parameters("@Quantity").NumericScale = 2
                    If Trim(item.text) = "" Then
                        cmd.CommandText = "BASE_Inventory_Insert"
                        cmd.Execute
                        item.text = cmd.Parameters("@InventoryId")
                        item.SubItems(2) = ProductId
                    Else
                        cmd.CommandText = "BASE_Inventory_Update"
                        cmd.Execute
                    End If
                Next
                
                'DELETE INVENTORY ITEMS
                Dim i As Long
                For i = 0 To 10000
                    If deleteCtr(i) = 0 Then Exit For
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.Parameters.Append cmd.CreateParameter("@InventoryId", adInteger, adParamInput, , deleteCtr(i))
                    cmd.CommandText = "BASE_Inventory_Delete"
                    cmd.Execute
                Next i
                                
                Dim isFound As Boolean
                isFound = False
                For Each item In lvSearch.ListItems
                    If ProductId = item.text Then
                        item.SubItems(1) = txtBasicInfo_ItemCode.text
                        item.SubItems(2) = txtBasicInfo_Name.text
                        isFound = True
                        item.Selected = True
                        item.EnsureVisible
                        Exit For
                    End If
                Next
                If isFound = False Then
                    Set item = lvSearch.ListItems.add(, , ProductId)
                        item.SubItems(1) = txtBasicInfo_ItemCode.text
                        item.SubItems(2) = txtBasicInfo_Name.text
                        item.Selected = True
                        item.EnsureVisible
                End If
                
                'AUDIT TRAIL
                'saveposaudittrail userid,workstationid,
                
                con.CommitTrans
                BASE_ContainerFrm.statusBar_Main.Panels(1).text = MessageCodes(1) & " " & MessageCodes(0)
                con.Close
            End If
        Case 4 ' Delete
            If ProductId <> 0 Then
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Product_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                If isActive = True Then
                    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "True")
                End If
                con.BeginTrans
                cmd.Execute
                con.CommitTrans
                con.Close
                If isActive = True Then
                    BASE_ContainerFrm.statusBar_Main.Panels(1).text = MessageCodes(1) & " " & MessageCodes(4)
                    isActive = False
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 4
                    isActivated (False)
                Else
                    BASE_ContainerFrm.statusBar_Main.Panels(1).text = MessageCodes(1) & " " & MessageCodes(5)
                    isActive = True
                    tb_Standard.Buttons(4).Caption = "Delete"
                    tb_Standard.Buttons(4).Image = 3
                    isActivated (True)
                End If
            End If
        Case 6 ' COPY
            If ProductId = 0 Then Exit Sub
            txtBasicInfo_ItemCode.SetFocus
            'txtBasicInfo_ItemCode.text = ""
            txtBasicInfo_ItemCode.BackColor = &HC0C0FF
            ProductId = 0
            
            For Each item In lvInventory.ListItems
                item.text = ""
            Next
    End Select
    Exit Sub
ErrHandler:
    con.RollbackTrans
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
        BASE_ContainerFrm.statusBar_Main.Panels(1).text = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
        BASE_ContainerFrm.statusBar_Main.Panels(1).text = ErrorCodes(0) & " " & Err.Description
    End If
        GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub txtBasicInfo_ItemCode_Change()
    If Trim(txtBasicInfo_ItemCode.text) = "" Then
        txtBasicInfo_ItemCode.BackColor = &HC0C0FF
    Else
        txtBasicInfo_ItemCode.BackColor = vbWhite
    End If
End Sub

Private Sub txtBasicInfo_Name_Change()
    If Trim(txtBasicInfo_Name.text) = "" Then
        txtBasicInfo_Name.BackColor = &HC0C0FF
    Else
        txtBasicInfo_Name.BackColor = vbWhite
    End If
End Sub

Private Sub txtCostingInfo_AverageCost_Change()
    If IsNumeric(txtCostingInfo_AverageCost.text) = False And Trim(txtCostingInfo_AverageCost.text) <> "" Then
        txtCostingInfo_AverageCost.BackColor = &HC0C0FF
    Else
        txtCostingInfo_AverageCost.BackColor = vbWhite
    End If
End Sub

Private Sub txtReorderPoint_GotFocus()
    selectText txtReorderPoint
End Sub

Private Sub txtReorderQuantity_GotFocus()
    selectText txtReorderQuantity
End Sub

Private Sub txtSalesInfo_BCP_Change()
    If IsNumeric(txtSalesInfo_BCP.text) = False And Trim(txtSalesInfo_BCP.text) <> "" Then
        txtSalesInfo_BCP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_BCP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_BCP_GotFocus()
    selectText txtSalesInfo_BCP
End Sub

Private Sub txtSalesInfo_DP_Change()
    If IsNumeric(txtSalesInfo_DP.text) = False And Trim(txtSalesInfo_DP.text) <> "" Then
        txtSalesInfo_DP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_DP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_DP_GotFocus()
    selectText txtSalesInfo_DP
End Sub

Private Sub txtSalesInfo_Price_Change()
    If IsNumeric(txtSalesInfo_Price.text) = False And Trim(txtSalesInfo_Price.text) <> "" Then
        txtSalesInfo_Price.BackColor = &HC0C0FF
    Else
        txtSalesInfo_Price.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_Price_GotFocus()
    selectText txtSalesInfo_Price
End Sub

Private Sub txtSalesInfo_SP_Change()
    If IsNumeric(txtSalesInfo_SP.text) = False And Trim(txtSalesInfo_SP.text) <> "" Then
        txtSalesInfo_SP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_SP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_SP_GotFocus()
    selectText txtSalesInfo_SP
End Sub

Private Sub txtSalesInfoBCMarkUp_Change()
    If IsNumeric(txtSalesInfoBCMarkUp.text) = False And Trim(txtSalesInfoBCMarkUp.text) <> "" Then
        txtSalesInfoBCMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoBCMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.text, ",", "")) * Val(Replace(txtSalesInfoBCMarkUp.text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.text, ",", ""))
        txtSalesInfo_BCP.text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoBCMarkUp_GotFocus()
    selectText txtSalesInfoBCMarkUp
End Sub

Private Sub txtSalesInfoDPMarkUp_Change()
    If IsNumeric(txtSalesInfoDPMarkUp.text) = False And Trim(txtSalesInfoDPMarkUp.text) <> "" Then
        txtSalesInfoDPMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoDPMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.text, ",", "")) * Val(Replace(txtSalesInfoDPMarkUp.text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.text, ",", ""))
        txtSalesInfo_DP.text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoDPMarkUp_GotFocus()
    selectText txtSalesInfoDPMarkUp
End Sub

Private Sub txtSalesInfoMSMarkUp_Change()
    If IsNumeric(txtSalesInfoMSMarkUp.text) = False And Trim(txtSalesInfoMSMarkUp.text) <> "" Then
        txtSalesInfoMSMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoMSMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.text, ",", "")) * Val(Replace(txtSalesInfoMSMarkUp.text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.text, ",", ""))
        txtSalesInfo_SP.text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoMSMarkUp_GotFocus()
    selectText txtSalesInfoMSMarkUp
End Sub

Private Sub txtSalesInfoSRPMarkUp_Change()
    If IsNumeric(txtSalesInfoSRPMarkUp.text) = False And Trim(txtSalesInfoSRPMarkUp.text) <> "" Then
        txtSalesInfoSRPMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoSRPMarkUp.BackColor = vbWhite
        'compute SRP Mark-up
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.text, ",", "")) * Val(Replace(txtSalesInfoSRPMarkUp.text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.text, ",", ""))
        txtSalesInfo_Price.text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoSRPMarkUp_GotFocus()
    selectText txtSalesInfoSRPMarkUp
End Sub

Private Sub txtSearch_ItemCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub

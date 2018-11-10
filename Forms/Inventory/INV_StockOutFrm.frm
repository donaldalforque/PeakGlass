VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_StockOutFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   7560
      Picture         =   "INV_StockOutFrm.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   3750
      TabIndex        =   49
      Top             =   1560
      Width           =   3755
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   4215
      Left            =   5040
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
      NumItems        =   7
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "On Hand"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Reserved"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Available"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   3615
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
         Left            =   2280
         TabIndex        =   37
         Top             =   2400
         Width           =   1215
      End
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
         TabIndex        =   36
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtSearch_OrderNumber 
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
         TabIndex        =   35
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbSearch_Customer 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6015
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   10610
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PurchaseOrderId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1440
         TabIndex        =   39
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94568449
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94568449
         CurrentDate     =   41686
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
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
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
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
         TabIndex        =   45
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label3 
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
         TabIndex        =   44
         Top             =   840
         Width           =   570
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
         TabIndex        =   43
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock-Out #"
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
         TabIndex        =   42
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         TabIndex        =   41
         Top             =   1200
         Width           =   900
      End
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   11535
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1755
         Left            =   7770
         TabIndex        =   6
         Top             =   360
         Width           =   3645
         Begin VB.TextBox txtOrderNumber 
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
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtStatus 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1080
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   960
            Width           =   2415
         End
         Begin VB.ComboBox cmbCustomer 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   1080
            TabIndex        =   10
            Top             =   600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   94568449
            CurrentDate     =   41509
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   900
         End
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   6735
         Begin VB.TextBox txtSalesman 
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
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salesman"
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
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Stock-out"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   2910
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   5700
         Left            =   120
         TabIndex        =   22
         Top             =   2085
         Width           =   11295
         Begin VB.TextBox txtItemSearch 
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
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   3495
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   4800
            Picture         =   "INV_StockOutFrm.frx":7D72
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   4935
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   8705
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
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "NewStockId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "NewStockLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Lot #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Expiry Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Cost"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "Sub-Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Location"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "ActualQuantity"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Name"
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
            TabIndex        =   26
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   7800
         Width           =   11295
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox txtCash 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   -9999
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   9120
            TabIndex        =   48
            Top             =   600
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Items:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   6720
            TabIndex        =   47
            Top             =   600
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   21
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CASH"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -9999
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   9120
            TabIndex        =   19
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   6720
            TabIndex        =   18
            Top             =   120
            Width           =   2235
         End
      End
      Begin VB.CommandButton btnComplete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Complete Order"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   582
         ButtonWidth     =   1667
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
               Caption         =   "Cancel"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   32
         Top             =   930
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CustomerId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Contact"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5760
      Picture         =   "INV_StockOutFrm.frx":7F96
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   2520
      Width           =   6195
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5880
      Picture         =   "INV_StockOutFrm.frx":10624
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   2520
      Width           =   6195
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":212C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":27B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":2E384
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":34BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":34E5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOutFrm.frx":354CC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INV_StockOutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StockOutId As Long
Dim StatusId As Integer
Dim item As MSComctlLib.ListItem
Private Sub Initialize()
    dtOrder.value = Format(Now, "MM/DD/YY hh:mm:ss")
    'picCompleted.Visible = False
    'pic_Cancelled.Visible = False
    
    StatusId = 0
    lvItems.ListItems.Clear
    txtStatus.text = ""
    txtRemarks.text = ""
    txtOrderNumber.text = ""
    lvItemList.Visible = False
    StockOutId = 0
    
    On Error Resume Next
    txtItemSearch.SetFocus
    cmbCustomer.ListIndex = 0
    
    LoadImageStatus picStatus, GetStatus(StatusId)
End Sub
Public Sub CountTotal()
    Dim item As MSComctlLib.ListItem
    Dim Total As Double
    
    For Each item In lvItems.ListItems
        item.SubItems(11) = FormatNumber(NVAL(item.SubItems(6)) * NVAL(item.SubItems(10)), 2, vbTrue, vbFalse)
        Total = Total + NVAL(item.SubItems(11))
    Next
    lblTotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub
Private Sub btnComplete_Click()
'    If EditAccessRights(28) = False Then
'        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
'        Exit Sub
'    End If
    If StatusId = 2 Then
        MsgBox "Save failed. No changes made. Order is already complete.", vbCritical
        Exit Sub
    End If
    Dim x As Variant
    x = MsgBox("This will complete the transaction. Product inventories will now be updated. Proceed?", vbExclamation + vbOKCancel)
    If x = vbOK Then
        Save (1)
        Save (2)
    End If
End Sub

Private Sub btnItemSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Name"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.text)
    
    Set rec = cmd.Execute
    lvItemList.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvItemList.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Uom
                    'Item.SubItems(4) = FormatNumber(rec!totalQty, 2, vbTrue, vbFalse)
                lvItemList.Visible = True
                lvItemList.Left = 5040
                lvItemList.Top = 2640
                LastProductId = rec!ProductId
                rec.MoveNext
            Else
                rec.MoveNext
            End If
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub btnSearch_Click()
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    cmd.ActiveConnection = con
    cmd.CommandText = "INV_StockOut_Search"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtOrderNumber.text)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    If cmbSearch_Status.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    If cmbSearch_Customer.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , cmbSearch_Customer.ItemData(cmbSearch_Customer.ListIndex))
    End If
    
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!StockOutId)
                item.SubItems(1) = rec!OrderNumber
                item.SubItems(2) = rec!Status
            rec.MoveNext
        Loop
    End If
    con.Close
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            If lvItemList.Visible = True Then
                lvItemList.Visible = False
                lvItemList.Left = -9999
                txtItemSearch.SetFocus
            End If
        Case vbKeyN
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(1)
            End If
        Case vbKeyS
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(2)
            End If
        Case vbKeyO
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(4)
            End If
        Case vbKeyP
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(6)
            End If
        Case vbKeyF4
            txtItemSearch.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.0871
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.4341
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.08
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.08
    lvItems.ColumnHeaders(9).width = lvItems.width * 0
    lvItems.ColumnHeaders(10).width = lvItems.width * 0
    lvItems.ColumnHeaders(11).width = lvItems.width * 0.0871
    lvItems.ColumnHeaders(12).width = lvItems.width * 0.093
    lvItems.ColumnHeaders(13).width = lvItems.width * 0.1071
    
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.48
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.45
    
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.285
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.685
    
    DateFrom.value = Format(Now, "mm/dd/yy")
    DateTo.value = Format(Now, "mm/dd/yy")
    
    Populate "Status"
    Populate "StockOutLoad"
    Populate "Customer"
    Initialize
End Sub

Private Sub lvItemList_DblClick()
    INV_StockOutOptFrm.Show (1)
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItemList_DblClick
    End Select
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.Count <= 0 Then Exit Sub
    lvItemList.Visible = False
    INV_StockOutOptFrm.Show (1)
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItems_DblClick
        Case vbKeyDelete
            Dim x As Variant
            x = MsgBox("This will permanently delete the selected record. Proceed?", vbExclamation + vbOKCancel)
            If x = vbOK Then
                Dim con As New ADODB.Connection
                con.ConnectionString = ConnString
                con.Open
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "INV_StockOutLine_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@StockOutLineId", adInteger, adParamInput, , NVAL(lvItems.SelectedItem.SubItems(3)))
                cmd.Execute
                con.Close
            End If
            lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
            MsgBox "Item removed.", vbInformation
    End Select
End Sub
Private Sub SelectOrders()
    Initialize
    
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    
    'Get StockOut
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_StockOut_Get"
    cmd.ActiveConnection = con
    cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , lvSearch.SelectedItem.text)
    
    Set rec = cmd.Execute
    If Not rec.EOF Then
        txtOrderNumber.text = rec!OrderNumber
        dtOrder.value = rec!Date
        txtRemarks.text = rec!Remarks
        StatusId = rec!StatusId
        txtStatus.text = rec!Status
        StockOutId = rec!StockOutId
        If rec!Customer <> "" Then
            cmbCustomer.text = rec!Customer
        Else
            cmbCustomer.ListIndex = 0
        End If
    End If
    
    'Get Line
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandText = "INV_StockOutLine_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , lvSearch.SelectedItem.text)
    Set rec = cmd.Execute
    
    lvItems.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvItems.ListItems.add(, , rec!ProductId)
                item.SubItems(1) = rec!StockOutId
                item.SubItems(2) = rec!LocationId
                item.SubItems(3) = rec!StockOutlineid
                item.SubItems(4) = rec!itemcode
                item.SubItems(5) = rec!Name
                item.SubItems(6) = FormatNumber(rec!Quantity, 2, vbTrue, vbFalse)
                item.SubItems(7) = rec!Unit
                item.SubItems(8) = rec!lotnumber
                If Not IsNull(rec!expirydate) Then
                    item.SubItems(9) = Format(rec!expirydate, "MM/DD/YY")
                End If
                item.SubItems(10) = FormatNumber(rec!cost, 2, vbTrue, vbFalse) 'Cost
                item.SubItems(11) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse) 'subtotal
                item.SubItems(12) = rec!Location
                item.SubItems(13) = rec!actualquantity
            rec.MoveNext
        Loop
    End If
    
    If StatusId = 2 Then
        txtStatus.text = "Completed"
        'picCompleted.Left = 5760
        'picCompleted.Visible = True
        'pic_Cancelled.Visible = False
    End If
    
    If StatusId = 7 Then
        'pic_Cancelled.Left = 5760
        'pic_Cancelled.Visible = True
        'picCompleted.Visible = False
        txtStatus.text = "Cancelled"
    End If
    CountTotal
    con.Close
    
    LoadImageStatus picStatus, GetStatus(StatusId)
End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    SelectOrders
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If EditAccessRights(28) = False Then
'        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
'        Exit Sub
'    End If
    Select Case Button.Index
        Case 1 ' New
            Initialize
        Case 2 'Save
            If StatusId <= 1 Then
                Save (1) 'Status Open
            ElseIf StatusId = 2 Then
                MsgBox "Save failed. No changes made. Order is already complete.", vbCritical
            ElseIf StatusId = 7 Then
                MsgBox "Save failed. No changes made. Order is already cancelled.", vbCritical
            End If
        Case 4 'Cancel
            If StatusId = 1 Then
                MsgBox "Save failed. No changes made. Cannot cancel an open/draft order.", vbCritical
                Exit Sub
            ElseIf StatusId = 7 Then
                MsgBox "Save failed. No changes made. Order is already cancelled.", vbCritical
                Exit Sub
            End If
            If StockOutId <> 0 Then
                Dim x As Variant
                x = MsgBox("This will cancel the transaction and revert all product inventories. Proceed?", vbCritical + vbOKCancel)
                If x = vbOK Then
                    Save (7)
                End If
            End If
        Case 6 'PRINT PREVIEW
            If StockOutId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_StockOut.rpt")
                crxRpt.RecordSelectionFormula = "{INV_StockOut.StockOutId}= " & Val(StockOutId) & ""
                crxRpt.DiscardSavedData
                
                Call ResetRptDB(crxRpt)
                'crxRpt.PrintOut False

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
            End If
    End Select
    LoadImageStatus picStatus, GetStatus(StatusId)
End Sub

Private Sub txtItemSearch_Change()
    btnItemSearch_Click
End Sub


Private Sub Save(ByVal iStatusId As Integer, Optional isReopen As Variant)
    If lvItems.ListItems.Count > 0 Then
        
        'On Error GoTo ErrorHandler
        
        Dim item As MSComctlLib.ListItem
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInputOutput, , StockOutId)
        If Trim(txtOrderNumber.text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDBDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
        cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(lblTotal.Caption))
                                  cmd.Parameters("@Total").Precision = 18
                                  cmd.Parameters("@Total").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , cmbCustomer.ItemData(cmbCustomer.ListIndex))
        
        If StockOutId = 0 Then
            cmd.CommandText = "INV_StockOut_Insert"
        Else
            cmd.CommandText = "INV_StockOut_Update"
        End If
        
        cmd.Execute
        
        If StockOutId = 0 Then 'ADD TO SEARCH
            Set item = lvSearch.ListItems.add(, , cmd.Parameters("@StockOutId"))
                item.SubItems(1) = cmd.Parameters("@OrderNumber")
                If iStatusId = 1 Then item.SubItems(2) = "Open"
                If iStatusId = 2 Then item.SubItems(2) = "Completed"
                item.Selected = True
                item.EnsureVisible
        Else
            For Each item In lvSearch.ListItems
                If Val(item.text) = Val(cmd.Parameters("@StockOutId")) Then
                    item.SubItems(1) = cmd.Parameters("@OrderNumber")
                    If iStatusId = 1 Then item.SubItems(2) = "Open"
                    If iStatusId = 2 Then item.SubItems(2) = "Completed"
                    item.Selected = True
                    item.EnsureVisible
                    Exit For
                End If
            Next
        End If
        
        StockOutId = cmd.Parameters("@StockOutId")
        txtOrderNumber.text = cmd.Parameters("@OrderNumber")
        
        'SAVE ORDER LINE
        For Each item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            cmd.Parameters.Append cmd.CreateParameter("@StockOutLineId", adInteger, adParamInputOutput, , NVAL(item.SubItems(3)))
            cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , StockOutId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.text)
            cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , NVAL(item.SubItems(2)))
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(6)))
                                  cmd.Parameters("@Quantity").Precision = 18
                                  cmd.Parameters("@Quantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(7))
            cmd.Parameters.Append cmd.CreateParameter("@LotNumber", adVarChar, adParamInput, 250, item.SubItems(8))
            If item.SubItems(9) = "" Then
                cmd.Parameters.Append cmd.CreateParameter("@ExpiryDate", adDate, adParamInput, , Null)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@ExpiryDate", adDate, adParamInput, , item.SubItems(9))
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Cost", adDecimal, adParamInput, , NVAL(item.SubItems(10)))
                                  cmd.Parameters("@Cost").Precision = 18
                                  cmd.Parameters("@Cost").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(item.SubItems(11)))
                                  cmd.Parameters("@Subtotal").Precision = 18
                                  cmd.Parameters("@Subtotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
            cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , item.SubItems(13))
                                  cmd.Parameters("@ActualQuantity").Precision = 18
                                  cmd.Parameters("@ActualQuantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Customer", adVarChar, adParamInput, 250, cmbCustomer.text)
            If item.SubItems(3) = "" Then
                cmd.CommandText = "INV_StockOutLine_Insert"
                cmd.Execute
            Else
                cmd.CommandText = "INV_StockOutLine_Update"
                cmd.Execute
            End If
            'cmd.Execute
            item.SubItems(3) = cmd.Parameters("@StockOutLineId")
        Next
        
        con.CommitTrans
        con.Close
        
        txtStatus.text = "Open"
        
        If iStatusId = 2 Then
            txtStatus.text = "Completed"
            'picCompleted.Left = 5760
            'picCompleted.Visible = True
            'pic_Cancelled.Visible = False
            StatusId = 2
        End If
        
        If iStatusId = 7 Then
            'pic_Cancelled.Left = 5760
            'pic_Cancelled.Visible = True
            'picCompleted.Visible = False
            txtStatus.text = "Cancelled"
            StatusId = 7
        End If
''
'        Dim isFound As Boolean
'        isFound = False
'        For Each item In lvSearch.ListItems
'            If StockOutId = item.text Then
'                item.SubItems(1) = txtOrderNumber.text
'                item.SubItems(2) = txtStatus.text
'                isFound = True
'                item.Selected = True
'                item.EnsureVisible
'                Exit For
'            End If
'        Next
'        If isFound = False Then
'            Set item = lvSearch.ListItems.add(, , StockOutId)
'                item.SubItems(1) = txtOrderNumber.text
'                item.SubItems(2) = txtStatus.text
'                item.Selected = True
'                item.EnsureVisible
'        End If
        
'        'PRINT PREVIEW
'        If StatusId = 2 Then
'            'Screen.MousePointer = vbHourglass
'            'BASE_PrintPreviewFrm.Show
'            Dim crxApp As New CRAXDRT.Application
'            Dim crxRpt As New CRAXDRT.Report
'            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_StockOut.rpt")
'            crxRpt.RecordSelectionFormula = "{INV_StockOut.StockOutId}= " & StockOutId & ""
'            crxRpt.DiscardSavedData
'
'            Call ResetRptDB(crxRpt)
'            crxRpt.PrintOut False
'
''            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
''            BASE_PrintPreviewFrm.CRViewer.ViewReport
''            BASE_PrintPreviewFrm.CRViewer.Zoom 1
''            Screen.MousePointer = vbDefault
'        End If
    Else
        MsgBox "Save failed. No data found.", vbCritical, "PeakPOS"
    End If
    Exit Sub
'ErrorHandler:
'    con.RollbackTrans
'    con.Close
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
End Sub

Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Status"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Status")
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSearch_Status.AddItem rec!Status
                    cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = rec!StatusId
                    rec.MoveNext
                Loop
            End If
        Case "StockOutLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_StockOut_Get"
            cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , Null)
            Set rec = cmd.Execute
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvSearch.ListItems.add(, , rec!StockOutId)
                        item.SubItems(1) = rec!OrderNumber
                        item.SubItems(2) = rec!Status
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "StockOutGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_StockOut_Get"
            cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , StockOutId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                StatusId = rec!StatusId
                txtStatus.text = rec!Status
                txtOrderNumber.text = rec!OrderNumber
                dtOrder.value = Format(rec!Date, "MM/DD/YY")
                txtRemarks.text = rec!Remarks
            End If
            con.Close
        Case "StockOutLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_StockOutLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@StockOutId", adInteger, adParamInput, , StockOutId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!StockOutlineid)
                        
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "Customer"
            Set rec = Global_Data("Customer")
            cmbCustomer.Clear
            cmbCustomer.AddItem ""
            cmbCustomer.ItemData(cmbCustomer.NewIndex) = 0
            
            cmbSearch_Customer.Clear
            cmbSearch_Customer.AddItem ""
            cmbSearch_Customer.ItemData(cmbSearch_Customer.NewIndex) = 0
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbCustomer.AddItem rec!Name
                    cmbCustomer.ItemData(cmbCustomer.NewIndex) = rec!CustomerId
                    
                    cmbSearch_Customer.AddItem rec!Name
                    cmbSearch_Customer.ItemData(cmbSearch_Customer.NewIndex) = rec!CustomerId
                    rec.MoveNext
                Loop
            End If
            
            On Error Resume Next
            cmbCustomer.ListIndex = 0
            cmbSearch_Customer.ListIndex = 0
    End Select
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
    
End Sub

Private Sub txtSearch_OrderNumber_Change()
    btnSearch_Click
End Sub


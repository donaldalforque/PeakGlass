VERSION 5.00
Begin VB.Form POS_DiscountFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "POS_DiscountFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   2880
      Picture         =   "POS_DiscountFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton btnNoDiscount 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4: STANDARD RETAIL PRICE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton btnCenter 
      BackColor       =   &H008080FF&
      Caption         =   "3: MEGA DISCOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton btnSatellite 
      BackColor       =   &H00FF8080&
      Caption         =   "2: SPECIAL DISCOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton btnDistributor 
      BackColor       =   &H0080C0FF&
      Caption         =   "1: STANDARD DISCOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DISCOUNTS"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_DiscountFrm.frx":239B
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_DiscountFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Pricing(ByVal a As String)
    Dim item As MSComctlLib.ListItem
    Dim savings As Double
    For Each item In POS_CashierFrm.lvList.ListItems
        If item.Selected = True Then
            If UCase(a) = UCase("None") Then
                item.SubItems(3) = FormatNumber(item.SubItems(9), 2, vbTrue, vbFalse)
                savings = (Val(Replace(item.SubItems(9), ",", "")) - Val(Replace(item.SubItems(3), ",", "")))
                POS_CashierFrm.lblDiscount.Caption = "| DISCOUNT TYPE: NONE"
                item.SubItems(4) = ""
            ElseIf UCase(a) = UCase("Distributor's Price") Then
                item.SubItems(3) = FormatNumber(item.SubItems(9), 2, vbTrue, vbFalse)
                If Val(Replace(item.SubItems(10), ",", "")) = 0 Then
                    savings = 0
                Else
                    savings = (Val(Replace(item.SubItems(9), ",", "")) * Val(Replace(item.SubItems(2), ",", ""))) _
                    - ((Val(Replace(item.SubItems(10), ",", "")) * Val(Replace(item.SubItems(2), ",", ""))))
                End If
                POS_CashierFrm.lblDiscount.Caption = "| DISCOUNT TYPE: DP"
            ElseIf UCase(a) = UCase("Satellite Price") Then
                item.SubItems(3) = FormatNumber(item.SubItems(9), 2, vbTrue, vbFalse)
                If Val(Replace(item.SubItems(11), ",", "")) = 0 Then
                    savings = 0
                Else
                    savings = (Val(Replace(item.SubItems(9), ",", "")) * Val(Replace(item.SubItems(2), ",", ""))) _
                    - (Val(Replace(item.SubItems(11), ",", "")) * Val(Replace(item.SubItems(2), ",", "")) * -1)
                End If
                POS_CashierFrm.lblDiscount.Caption = "| DISCOUNT TYPE: MS"
            ElseIf UCase(a) = UCase("Business Center Price") Then
                item.SubItems(3) = FormatNumber(item.SubItems(9), 2, vbTrue, vbFalse)
                If Val(Replace(item.SubItems(12), ",", "")) = 0 Then
                    savings = 0
                Else
                    savings = (Val(Replace(item.SubItems(9), ",", "")) * Val(Replace(item.SubItems(2), ",", ""))) _
                    - (Val(Replace(item.SubItems(12), ",", "")) * Val(Replace(item.SubItems(2), ",", "")) * -1)
                End If
                POS_CashierFrm.lblDiscount.Caption = "| DISCOUNT TYPE: BC"
            End If
            item.SubItems(4) = FormatNumber(savings, 2, vbTrue, vbFalse)
            If POS_CashierFrm.lblDiscount.Caption = "| DISCOUNT TYPE: NONE" Then
                item.SubItems(4) = ""
            End If
        End If
    Next
    POS_CashierFrm.CountTax
    POS_CashierFrm.CountTotal
End Sub

Private Sub btnAccept_Click()
    
End Sub

Private Sub btnCancel_Click()
    Pricing ("None")
    Unload Me
End Sub

Private Sub btnCenter_Click()
    'business center's price
    Pricing ("Business Center Price")
    Unload Me
End Sub

Private Sub btnDistributor_Click()
    'distributor's price
    Pricing ("Distributor's Price")
    Unload Me
End Sub

Private Sub btnNoDiscount_Click()
    'no discount
    Pricing ("None")
    Unload Me
End Sub

Private Sub btnSatellite_Click()
    Pricing ("Satellite Price")
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey1
            btnDistributor_Click
        Case vbKey2
            btnSatellite_Click
        Case vbKey3
            btnCenter_Click
        Case vbKey4
            btnNoDiscount_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub


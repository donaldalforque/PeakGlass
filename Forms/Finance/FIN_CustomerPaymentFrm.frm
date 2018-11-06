VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_CustomerPaymentFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Payment"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
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
      Left            =   10140
      TabIndex        =   15
      Top             =   9240
      Width           =   1575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
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
      Left            =   8520
      TabIndex        =   14
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   11655
      Begin VB.TextBox txtCheckAmount 
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
         Left            =   7200
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1575
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbAccount 
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
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbBank 
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
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
      End
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
         Height          =   330
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCheckNumber 
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
         Left            =   7200
         TabIndex        =   8
         Top             =   1935
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbType 
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
         Left            =   -9999
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtCash 
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
         Left            =   1560
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtSalesReturn 
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
         Left            =   7200
         TabIndex        =   10
         Top             =   2655
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CheckBox chkOnline 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bank Payment"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker dtCheckDate 
         Height          =   345
         Left            =   7200
         TabIndex        =   9
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
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
         Format          =   138149889
         CurrentDate     =   41646
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   345
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
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
         Format          =   138149889
         CurrentDate     =   41646
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amt."
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
         Left            =   5760
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
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
         Left            =   5760
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Left            =   5760
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date"
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
         Left            =   5760
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   23
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check #"
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
         Left            =   5760
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
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
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
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
         TabIndex        =   20
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[For Bank Payment, Save to]"
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
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Discount"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvOrders 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SalesOrderId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Terms"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2716
      EndProperty
   End
   Begin VB.Label lblTotalBalance 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7200
      TabIndex        =   30
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7200
      TabIndex        =   29
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1080
      MouseIcon       =   "FIN_CustomerPaymentFrm.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblSelectAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      MouseIcon       =   "FIN_CustomerPaymentFrm.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5040
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Payments"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FIN_CustomerPaymentFrm.frx":02A4
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "FIN_CustomerPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CustomerId As Long

Private Sub btnCancel_Click()
    Unload Me
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbBank.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbBank.AddItem rec!bankname
                        cmbBank.ItemData(cmbBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbBank.ListIndex = 0
    End Select
End Sub
Private Sub btnSave_Click()
    Dim x As Variant
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo, "Verify")
    If x = vbNo Then
        Exit Sub
    End If
    Dim hasSelected As Boolean
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        If item.Checked = True Then
            hasSelected = True
            Exit For
        End If
    Next
    If hasSelected = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(37)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    Dim isOnline As Boolean
    If chkOnline.value = Checked Then isOnline = True
    If chkOnline.value = Unchecked Then isOnline = False
    
    'SAVE PAYMENT
    If isOnline = True Then
        If cmbAccount.text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
            GLOBAL_MessageFrm.Show (1)
            cmbAccount.SetFocus
            Exit Sub
        End If
    End If
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    con.BeginTrans
    
    'TRANSACTION ID
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_TransactionId_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInputOutput, , 0)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "Customer Payment")
    cmd.Execute
    Dim TransactionId As Long
    TransactionId = cmd.Parameters("@TransactionId")
    
    'SO_Payment
    Dim Payment As Double
    Payment = Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", ""))
    Dim Orders As String
    
    For Each item In lvOrders.ListItems
        If item.Checked = True Then
            If Payment <= 0 Then Exit For
            
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Payment_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , item.SubItems(1))
            If Payment > Val(Replace(item.SubItems(6), ",", "")) Then
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                                      cmd.Parameters("@Amount").NumericScale = 2
                                      cmd.Parameters("@Amount").Precision = 18
            Else
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Payment)
                                      cmd.Parameters("@Amount").NumericScale = 2
                                      cmd.Parameters("@Amount").Precision = 18
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@SalesReturn").NumericScale = 2
                                  cmd.Parameters("@SalesReturn").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, Null)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Null)
            If isOnline = True Then
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , isOnline)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
            cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            cmd.Execute
            
            Payment = Payment - Val(Replace(item.SubItems(6), ",", ""))
            Orders = Orders & "[" & item.SubItems(2) & "]"
            
            'CUSTOMER LEDGER
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CustomerLedger_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , item.SubItems(1))
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment " & txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@Debit", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@Debit").NumericScale = 2
                                  cmd.Parameters("@Debit").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Credit", adDecimal, adParamInput, , (Val(Replace(txtSalesReturn.text, ",", "")) + Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", ""))))
                                  cmd.Parameters("@Credit").NumericScale = 2
                                  cmd.Parameters("@Credit").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
            cmd.Execute
        End If
    Next
    
    'PAYMENT HISTORY
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_PaymentHistory_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                          cmd.Parameters("@Amount").NumericScale = 2
                          cmd.Parameters("@Amount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                          cmd.Parameters("@CheckAmount").NumericScale = 2
                          cmd.Parameters("@CheckAmount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 50, txtCheckNumber.text)
    cmd.Parameters.Append cmd.CreateParameter("@SalesDiscount", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                          cmd.Parameters("@SalesDiscount").NumericScale = 2
                          cmd.Parameters("@SalesDiscount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 4000, txtRemarks.text & " " & Orders)
    cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
    cmd.Execute
    
    'INFLOW AND OUTFLOW
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_InFlow_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                          cmd.Parameters("@Cash").Precision = 18
                          cmd.Parameters("@Cash").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                          cmd.Parameters("@CheckAmount").Precision = 18
                          cmd.Parameters("@CheckAmount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, txtRemarks.text & ":" & FIN_AccountsReceivable.lvSearch.SelectedItem.SubItems(2))
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
    cmd.Execute
        
    If chkOnline.value = Checked Then
        'UPDATE BANK ACCOUNT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_FundBank_Add"
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Now)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").Precision = 18
                              cmd.Parameters("@Amount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 400, "")
        cmd.Execute
    Else
        'UPDATE FUND ACCOUNT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_Fund_AddBalance"
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'DEFAULT
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").Precision = 18
                              cmd.Parameters("@Amount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Execute
    End If
    
    con.CommitTrans
    con.Close
    
    MsgBox "Payment Successful!", vbInformation, "Success!"
    
    Dim Y As Variant
    Y = MsgBox("Payment Successful! Do you want to print a receipt?", vbInformation + vbYesNo)
    If Y = vbYes Then
        'PRINT
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Account.rpt")
        'crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & Val(POS_SalesId) & ""
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
        crxRpt.ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(lvOrders.SelectedItem.SubItems(1))
        crxRpt.OpenSubreport("POS_ReceiptAccount_Payments.rpt").ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(lvOrders.SelectedItem.SubItems(1))
        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
    End If
    
    FIN_AccountsReceivable.btnSearch_Click
    Unload Me
End Sub

Private Sub chkOnline_Click()
  If chkOnline.value = Checked Then
        lblBank.Visible = True
        cmbBank.Visible = True
        lblAccount.Visible = True
        cmbAccount.Visible = True
        lblCaption.Visible = True
'        txtCheckAmount.Enabled = True
'        dtCheckDate.Enabled = True
'        txtCheckNumber.Enabled = True
    Else
        lblBank.Visible = False
        cmbBank.Visible = False
        lblAccount.Visible = False
        cmbAccount.Visible = False
        lblCaption.Visible = False
'        txtCheckAmount.Enabled = False
'        dtCheckDate.Enabled = False
'        txtCheckNumber.Enabled = False
    End If
End Sub

Private Sub cmbBank_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Load"
    
    cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbBank.ItemData(cmbBank.ListIndex))
    Set rec = cmd.Execute
    cmbAccount.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccount.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccount.ItemData(cmbAccount.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Load()
    lvOrders.ColumnHeaders(1).width = lvOrders.width * 0.025
    lvOrders.ColumnHeaders(3).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(4).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(5).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(6).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(7).width = lvOrders.width * 0.295
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandText = "SO_SalesOrder_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 4) 'INVOICED
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
    cmd.Parameters.Append cmd.CreateParameter("@Sort", adVarChar, adParamInput, 50, "Date")
    
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!balance > 0 Then
            Set item = lvOrders.ListItems.add(, , "")
                item.SubItems(1) = rec!SalesOrderId
                item.SubItems(2) = rec!OrderNumber
                item.SubItems(3) = Format(rec!Date, "MM/DD/YY")
                item.SubItems(4) = Format(rec!DueDate, "MM/DD/YY")
                If Not IsNull(rec!Terms) Then item.SubItems(5) = rec!Terms
                item.SubItems(6) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
            End If
            rec.MoveNext
        Loop
    End If
    
    con.Close
    
    dtDate.value = Format(Now, "MM/DD/YY hh:mm:ss")
    dtCheckDate.value = Format(Now, "MM/DD/YY")
    
    Populate "Bank"
    CountTotal
End Sub
Private Sub CountTotal()
    Dim item As MSComctlLib.ListItem
    Dim Total, balance As Double
    Total = 0
    For Each item In lvOrders.ListItems
        balance = balance + Val(Replace(item.SubItems(6), ",", ""))
        If item.Checked = True Then
            Total = Total + Val(Replace(item.SubItems(6), ",", ""))
        End If
    Next
    lblTotal.Caption = "Total Selected: " & FormatNumber(Total, 2, vbTrue, vbFalse)
    lblTotalBalance.Caption = "Total Balance: " & FormatNumber(balance, 2, vbTrue, vbFalse)
'    CountTotal
End Sub

Private Sub Label2_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        item.Checked = False
    Next
    CountTotal
End Sub

Private Sub lblSelectAll_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        item.Checked = True
    Next
    CountTotal
End Sub

Private Sub lvOrders_ItemCheck(ByVal item As MSComctlLib.ListItem)
    CountTotal
    Dim itemx As MSComctlLib.ListItem
    For Each itemx In lvOrders.ListItems
        If item.Index <> itemx.Index Then
            itemx.Checked = False
        End If
    Next
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.text) = False Then
        txtCash.text = "0.00"
        selectText txtCash
    ElseIf Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtCash.text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", ""))), 2, vbTrue)
    End If
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.text) = False Then
        txtCheckAmount.text = "0.00"
        selectText txtCheckAmount
     ElseIf Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtCheckAmount.text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", ""))), 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesReturn_Change()
    If IsNumeric(txtSalesReturn.text) = False Then
        txtSalesReturn.text = "0.00"
        selectText txtSalesReturn
     ElseIf Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtSalesReturn.text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtSalesReturn.text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", ""))), 2, vbTrue, vbFalse)
    End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form SO_PenaltyFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "SO_PenaltyFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtPenalty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
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
         Left            =   1800
         TabIndex        =   9
         Top             =   6960
         Width           =   4695
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   6600
         Width           =   4695
      End
      Begin VB.TextBox txtTruckingCharge 
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
         Left            =   1800
         TabIndex        =   7
         Top             =   6240
         Width           =   4695
      End
      Begin VB.TextBox txtScaleCharge 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   5880
         Width           =   4695
      End
      Begin VB.TextBox txtInterestCharge 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   5520
         Width           =   4695
      End
      Begin VB.TextBox txtParticulars 
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
         Left            =   1800
         TabIndex        =   0
         Text            =   "Charges"
         Top             =   3720
         Width           =   4695
      End
      Begin VB.TextBox txtInterestRate 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   5160
         Width           =   4695
      End
      Begin VB.TextBox txtDays 
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
         Left            =   1800
         TabIndex        =   3
         Top             =   4800
         Width           =   4695
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   2
         Top             =   4440
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dtPenalty 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   4080
         Width           =   4695
         _ExtentX        =   8281
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
         Format          =   71368705
         CurrentDate     =   41509
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         TabIndex        =   40
         Top             =   6600
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penalty 2"
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
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penalty 1"
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
         TabIndex        =   38
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Charge"
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
         TabIndex        =   37
         Top             =   5520
         Width           =   1425
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total"
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
         TabIndex        =   36
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Left            =   1800
         TabIndex        =   35
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label12 
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
         Left            =   240
         TabIndex        =   34
         Top             =   4080
         Width           =   435
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblCustomer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblOrderNumber 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Left            =   1800
         TabIndex        =   31
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Left            =   1800
         TabIndex        =   30
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1800
         TabIndex        =   29
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lblDueDate 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1800
         TabIndex        =   28
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label lblDaysOverdue 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
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
         Left            =   1800
         TabIndex        =   27
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1800
         TabIndex        =   26
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bal. Forwarded"
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
         TabIndex        =   25
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total with Int."
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   30
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   6255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days Overdue"
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
         TabIndex        =   21
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
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
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Late Payment Penalties"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order #"
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
         TabIndex        =   18
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   17
         Top             =   1695
         Width           =   435
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
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
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
         TabIndex        =   15
         Top             =   4800
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Rate"
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
         TabIndex        =   14
         Top             =   5160
         Width           =   1200
      End
      Begin VB.Label lblCashLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         TabIndex        =   13
         Top             =   4440
         Width           =   720
      End
   End
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
      Left            =   5400
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save && Close"
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
      Left            =   3960
      TabIndex        =   10
      Top             =   7560
      Width           =   1335
   End
End
Attribute VB_Name = "SO_PenaltyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SalesOrderId, CustomerId As Long
Private Sub CountTotal()
    Dim penalty, total, interestrate, days, balance, Interest, scalecharge, trucking, others As Double
    
    days = Val(Replace(txtDays.text, ",", ""))
    interestrate = Val(Replace(txtInterestRate.text, ",", ""))
    balance = Val(Replace(txtBalance.text, ",", ""))
    scalecharge = Val(Replace(txtScaleCharge.text, ",", ""))
    trucking = Val(Replace(txtTruckingCharge.text, ",", ""))
    others = Val(Replace(txtOthers.text, ",", ""))
    
    'COMPUTE INTEREST
    Interest = (days / 30) * (balance * (interestrate / 100))
    txtInterestCharge.text = FormatNumber(Interest, 2, vbTrue, vbFalse)
    txtPenalty.text = FormatNumber(scalecharge + trucking + others + Interest, 2, vbTrue)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_Penalty_Insert"
    
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtPenalty.value)
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 255, txtParticulars.text)
    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtPenalty.text, ",", "")))
                          cmd.Parameters("@Amount").NumericScale = 2
                          cmd.Parameters("@Amount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , Val(Replace(txtInterestRate.text, ",", "")))
                          cmd.Parameters("@InterestRate").NumericScale = 2
                          cmd.Parameters("@InterestRate").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@Days", adDecimal, adParamInput, , Val(Replace(txtDays.text, ",", "")))
                          cmd.Parameters("@Days").NumericScale = 2
                          cmd.Parameters("@Days").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@InterestCharge", adDecimal, adParamInput, , Val(Replace(txtInterestCharge.text, ",", "")))
                          cmd.Parameters("@InterestCharge").NumericScale = 2
                          cmd.Parameters("@InterestCharge").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@ScaleCharge", adDecimal, adParamInput, , Val(Replace(txtScaleCharge.text, ",", "")))
                          cmd.Parameters("@ScaleCharge").NumericScale = 2
                          cmd.Parameters("@ScaleCharge").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@TruckingCharge", adDecimal, adParamInput, , Val(Replace(txtTruckingCharge.text, ",", "")))
                          cmd.Parameters("@TruckingCharge").NumericScale = 2
                          cmd.Parameters("@TruckingCharge").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@OtherCharge", adDecimal, adParamInput, , Val(Replace(txtOthers.text, ",", "")))
                          cmd.Parameters("@OtherCharge").NumericScale = 2
                          cmd.Parameters("@OtherCharge").Precision = 18
    cmd.Execute
    con.Close
    MsgBox "Charges applied.", vbInformation, "Success"
    Unload Me
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
    
End Sub

Private Sub Form_Load()
    'txtBalance.text = lblBalance.Caption
    dtPenalty.value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SalesOrderId = 0
    CustomerId = 0
End Sub

Private Sub txtBalance_Change()
    If IsNumeric(txtBalance.text) = False Then
        txtBalance.text = "0.00"
    End If
    CountTotal
End Sub

Private Sub txtBalance_GotFocus()
    selectText txtBalance
End Sub

Private Sub txtDays_Change()
    If IsNumeric(txtDays.text) = False Then
        txtDays.text = "0"
    End If
    CountTotal
End Sub

Private Sub txtDays_GotFocus()
    selectText txtDays
End Sub

Private Sub txtInterestCharge_Change()
    'CountTotal
    If IsNumeric(Val(Replace(txtInterestCharge.text, ",", ""))) = False Then
        txtInterestCharge.text = "0.00"
    End If
End Sub

Private Sub txtInterestCharge_LostFocus()
    txtInterestCharge.text = FormatNumber(txtInterestCharge.text, 2, vbTrue)
End Sub

Private Sub txtInterestRate_Change()
    If IsNumeric(txtInterestRate.text) = False Then
        txtInterestRate.text = "0"
    End If
    CountTotal
End Sub

Private Sub txtInterestRate_GotFocus()
    selectText txtInterestRate
End Sub

Private Sub txtOthers_Change()
    CountTotal
End Sub

Private Sub txtOthers_LostFocus()
    txtOthers.text = FormatNumber(txtOthers.text, 2, vbTrue)
End Sub

Private Sub txtParticulars_GotFocus()
    selectText txtParticulars
End Sub

Private Sub txtPenalty_LostFocus()
    txtPenalty.text = FormatNumber(txtPenalty.text, 2, vbTrue)
End Sub

Private Sub txtScaleCharge_Change()
    CountTotal
End Sub

Private Sub txtScaleCharge_LostFocus()
    txtScaleCharge.text = FormatNumber(txtScaleCharge.text, 2, vbTrue)
End Sub

Private Sub txtTruckingCharge_Change()
    CountTotal
End Sub

Private Sub txtTruckingCharge_LostFocus()
    txtTruckingCharge.text = FormatNumber(txtTruckingCharge.text, 2, vbTrue)
End Sub

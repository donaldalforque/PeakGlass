VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form POS_EndOfShiftFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "POS_EndOfShiftFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEnd 
      Caption         =   "F2: End Shift and Log Out"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "POS_EndOfShiftFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print X-Reading"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "POS_EndOfShiftFrm.frx":0609
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
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
      Left            =   240
      Picture         =   "POS_EndOfShiftFrm.frx":2831
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   3975
   End
   Begin VB.ComboBox cmbCashier 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2280
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95485953
      CurrentDate     =   42297
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_EndOfShiftFrm.frx":4BC0
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End of Shift"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASHIER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5655
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POS_EndOfShiftFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "User"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("User")
            cmbCashier.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then  'Cashier
                        cmbCashier.AddItem rec!Name
                        cmbCashier.ItemData(cmbCashier.NewIndex) = rec!UserId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbCashier.ListIndex = 0
    End Select
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnEnd_Click()
    'Confirm End of Shift
    Dim X As Variant
    X = MsgBox("Are you sure you want to end your shift and log out?", vbQuestion + vbYesNo, "")
    If X = vbYes Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_UserAudit_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@Activity", adVarChar, adParamInput, 250, "END OF SHIFT")
        cmd.Execute
        con.Close
        Unload Me
        Unload POS_CashierFrm
        POS_UserLoginFrm.Show
    End If
End Sub

Private Sub btnPrint_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_XReading.rpt")
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
    
    'crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.text
    'crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue DateFrom.value & " " & TimeFrom.value
    'crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue DateTo.value & " " & TimeTo.value
    
    crxRpt.ParameterFields.GetItemByName("@Date").AddCurrentValue dtDate.value
    crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue cmbCashier.ItemData(cmbCashier.ListIndex)
    crxRpt.ParameterFields.GetItemByName("@WorkstationId").AddCurrentValue WorkstationId
    
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail UserId, WorkstationId, 0, "Generate X-Reading Report"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyF2
            btnEnd_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Populate "User"
    Dim X As Integer
    For X = 0 To cmbCashier.ListCount - 1
        If cmbCashier.ItemData(X) = UserId Then cmbCashier.ListIndex = X
    Next X
    dtDate.value = Format(Now, "MM/DD/YY")
End Sub

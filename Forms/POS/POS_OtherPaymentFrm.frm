VERSION 5.00
Begin VB.Form POS_OtherPaymentFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "POS_OtherPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Caption         =   "ALT+R: Remove"
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
      Left            =   2760
      Picture         =   "POS_OtherPaymentFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
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
      Left            =   6360
      Picture         =   "POS_OtherPaymentFrm.frx":065B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Left            =   4680
      Picture         =   "POS_OtherPaymentFrm.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtReferenceNumber 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label lblReferenceNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Number:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_OtherPaymentFrm.frx":4DBE
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Payment Option"
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
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "POS_OtherPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAccept_Click()
    If Val(Replace(txtAmount.text, ",", "")) < -1 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
        GLOBAL_MessageFrm.Show (1)
        txtAmount.SetFocus
    Else
        On Error GoTo ERRMSG
        Dim x As Variant
        x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo)
        If x = vbYes Then
            POS_ConfirmPaymentFrm.Show (1)
            If AllowAccess = False Then Exit Sub
        
            Dim totaldiscount As Double
            Dim item As MSComctlLib.ListItem
            For Each item In POS_CashierFrm.lvList.ListItems
                totaldiscount = totaldiscount + NVAL(item.SubItems(4))
            Next
        
            Dim SalesOrderId As Long
            
            'Save to Order
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            con.BeginTrans
            cmd.CommandType = adCmdStoredProc
            'cmd.CommandText = "POS_Sales_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInputOutput, , SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Now)
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , Now)
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2)
            cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , 1)
            cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId)
            cmd.Parameters.Append cmd.CreateParameter("@Days", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@Days").Precision = 18
                                  cmd.Parameters("@Days").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@InterestRate").Precision = 18
                                  cmd.Parameters("@InterestRate").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@Cash").Precision = 18
                                  cmd.Parameters("@Cash").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Interest", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@Interest").Precision = 18
                                  cmd.Parameters("@Interest").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                                  cmd.Parameters("@Subtotal").Precision = 18
                                  cmd.Parameters("@Subtotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                                  cmd.Parameters("@Total").Precision = 18
                                  cmd.Parameters("@Total").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 250, "")
            cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 250, txtReferenceNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@GatePass", adVarChar, adParamInput, 250, txtReferenceNumber.text)
            'cmd.Parameters.Append cmd.CreateParameter("@LastUser", adVarChar, adParamInput, 400, CurrentUser)
            cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , totaldiscount)
                                  cmd.Parameters("@Discount").Precision = 18
                                  cmd.Parameters("@Discount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'NOT SET!
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            cmd.CommandText = "SO_SalesOrder_Insert"
            cmd.Execute
            SalesOrderId = cmd.Parameters("@SalesOrderId")
            
            'Line
            'SAVE ORDER LINE
            'Dim item As MSComctlLib.ListItem
    
            For Each item In POS_CashierFrm.lvList.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "SO_SalesOrderLine_Insert"
                
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInputOutput, , 1)
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId)
                cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Now)
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(8))
                cmd.Parameters.Append cmd.CreateParameter("@ProductName", adVarChar, adParamInput, 250, item.text)
                cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(1)))
                                      cmd.Parameters("@Quantity").Precision = 18
                                      cmd.Parameters("@Quantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(item.SubItems(3)))
                                      cmd.Parameters("@Price").Precision = 18
                                      cmd.Parameters("@Price").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                                  cmd.Parameters("@UnitCost").NumericScale = 2
                                  cmd.Parameters("@UnitCost").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(item.SubItems(5)))
                                      cmd.Parameters("@Subtotal").Precision = 18
                                      cmd.Parameters("@Subtotal").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , 1)
                cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2)
                cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , 0)
                cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , (Val(Replace(item.SubItems(1), ",", "")) * Val(Replace(item.SubItems(16), ",", ""))))
                                      cmd.Parameters("@ActualQuantity").Precision = 18
                                      cmd.Parameters("@ActualQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ProductDescription", adVarChar, adParamInput, 250, item.text)
                cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@CutPurchase", adBoolean, adParamInput, , item.SubItems(19))
                cmd.Parameters.Append cmd.CreateParameter("@TrackClipping", adVarChar, adParamInput, 250, item.SubItems(21))
                cmd.Execute
            Next
            
            'UPDATE SO REMAINING BALANCE
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Balance_Update"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
            cmd.Execute

            'SAVE DOWNPAYMENT
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Payment_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtAmount.text))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Now)
            cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@SalesReturn").NumericScale = 2
                                  cmd.Parameters("@SalesReturn").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, Null)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , False)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
            cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , 0)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, txtReferenceNumber.text)
            cmd.Execute
            con.CommitTrans
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
                Call ResetRptDB(crxRpt)
                crxRpt.EnableParameterPrompting = False
                crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
                crxRpt.ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(SalesOrderId)
                crxRpt.OpenSubreport("POS_ReceiptAccount_Payments.rpt").ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(SalesOrderId)
                
                crxRpt.PrintOut False
            End If
            
            Unload Me
            Unload POS_PayFrm
            POS_CashierFrm.Initialize
        End If
    End If
    Exit Sub
ERRMSG:
    MsgBox Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    If btnRemove.Visible = False Then Exit Sub
    ClearClassData (3)
    POS_PayFrm.txtOthers.text = "0.00"
    POS_PayFrm.ComputeChange
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyR
            If btnRemove.Visible = False Then Exit Sub
            If Shift = vbAltMask Then
                btnRemove_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    With POS_OtherPaymentFrm
        .txtReferenceNumber.text = OtherInfo.ReferenceNumber
        .txtRemarks.text = OtherInfo.Remarks
        .txtAmount.text = FormatNumber(OtherInfo.Amount, 2, vbTrue, vbFalse)
    End With
    
    If Val(Replace(txtAmount.text, ",", "")) = 0 Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
    End If
End Sub

Private Sub txtAmount_Click()
    Set SYS_OSKFrm.txtControl = txtAmount
    SYS_OSKFrm.Caption = "Amount"
    SYS_OSKFrm.Show (1)
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub



Private Sub txtReferenceNumber_Click()
    Set SYS_OskAlphaFrm.txtControl = txtReferenceNumber
    SYS_OskAlphaFrm.Caption = lblReferenceNumber.Caption
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtReferenceNumber_GotFocus()
    selectText txtReferenceNumber
End Sub

Private Sub txtRemarks_Click()
    Set SYS_OskAlphaFrm.txtControl = txtRemarks
    SYS_OskAlphaFrm.Caption = "Remarks"
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtRemarks_GotFocus()
    selectText txtRemarks
End Sub

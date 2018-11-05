Attribute VB_Name = "UserFunctions"
Option Explicit
Global UserId, WorkstationId As Integer
Global AllowNegativeInventory As Boolean
Global ReorderPointCheckScheduled As String
Global ReorderPointCheckFrequency As Double
Global LastReorderPointDateCheck As Date
Global AllowAccess As Boolean
Public CardInfo As New CardPaymentInfo
Public CheckInfo As New CheckPaymentInfo
Public LoyaltyInfo As New LoyaltyPointsInfo
Public OtherInfo As New OtherPaymentInfo
Public CurrentUser As String
Public isModify As Boolean
Public ProductSet As ADODB.Recordset
Public isIndefiniteSize, isPricingReference, isLengthOnly, TrackClipping As Boolean
Public CurrentSetid As Integer
Public PriceDetails As String
Global NotificationTimer As Integer
Public Function Hostname() As String
    'Get Hostname from Text
    Open App.Path & "\Resources\Hostname.txt" For Input As #1
    Input #1, Hostname
    Close #1
End Function
Public Function ConnString() As String
    'Dim Hostname As String
    'Hostname = Environ("COMPUTERNAME") 'NOT SET
    'Hostname = "DSASERVER"
    ConnString = "Provider=SQLNCLI.1;Data Source = " & Hostname & "\PEAKSQL;User Id=sa; " & _
                 "Password=PeakPOS2015;Initial Catalog=Peak_GM"
End Function
Public Sub ResetRptDB(ByRef crxReport As CRAXDRT.Report)
    Dim DBProviderName As String ' i.e SQLOLEDB.1;
    Dim DBDataSource As String ' i.e brandon-pc\sqlexpress
    Dim DBName As String
    Dim DBUsername As String
    Dim DBPwd As String
    Dim ConnectionString As String
    Dim crxTable As DatabaseTable
    'Dim objDataAccess As DataAccess.clsDataAccess
    Dim i As Integer
    Dim crxSection As CRAXDRT.Section
    Dim ReportObject
    Dim crxSubReportObj
    Dim crxsubreport
    Dim crxdatatable
    'Dim Hostname As String
    
    DBProviderName = "SQLNCLI.1"
    'Hostname = Environ("COMPUTERNAME") 'NOT SET!!
    'Hostname = "DSASERVER"
    DBDataSource = Hostname & "\PEAKSQL"
    DBName = "Peak_GM"
    DBUsername = "sa"
    DBPwd = "PeakPOS2015"
    
    For Each crxTable In crxReport.Database.Tables
        Call crxTable.SetLogOnInfo(DBDataSource, DBName, DBUsername, DBPwd)
        Call crxTable.SetTableLocation(crxTable.Location, "", ConnString)
    Next
    For Each crxSection In crxReport.Sections
        For Each ReportObject In crxSection.ReportObjects
            If ReportObject.Kind = crSubreportObject Then
                Set crxSubReportObj = ReportObject
                Set crxsubreport = crxSubReportObj.OpenSubreport
                For Each crxdatatable In crxsubreport.Database.Tables
                    Call crxdatatable.SetLogOnInfo(DBDataSource, DBName, DBUsername, DBPwd)
                    Call crxdatatable.SetTableLocation(crxdatatable.Location, "", ConnectionString)
                Next
            End If
        Next
    Next
End Sub
Public Sub selectText(ByVal text As Control)
    text.SelStart = 0
    text.SelLength = Len(text.text)
End Sub
Public Sub CenterChildForm(ByVal Form As Form)
    Form.Left = (BASE_ContainerFrm.ScaleWidth - Form.width) / 2
    Form.Top = (BASE_ContainerFrm.ScaleHeight - Form.Height) / 2
End Sub
Public Sub CornerChildForm(ByVal Form As Form)
    On Error Resume Next
    Form.Left = 0
    Form.Top = 0
End Sub
Public Sub StatusBarWidth(ByVal Form As Form, ByVal Statusbar As Statusbar)
    On Error Resume Next
    Dim width As Double
    width = Form.ScaleWidth
    Statusbar.Panels(1).width = width * 0.3
    Statusbar.Panels(2).width = width * 0.2
    Statusbar.Panels(3).width = width * 0.2
    Statusbar.Panels(4).width = width * 0.3
End Sub
'Public Sub DistinctList(lv As MSComctlLib.ListView)
'    Dim i As Long
'    Dim j As Long
'    With lv
'        For i = 1 To .ListItems.Count
'            For j = .ListItems.Count To (i + 1) Step -1
'                If .ListItems(j) = .ListItems(i) Then
'                    .ListItems.Remove j
'                End If
'            Next
'        Next
'    End With
'End Sub

Public Function ErrorCodes(ByVal Code As Integer) As String
    Dim Errors(100) As String
    Errors(0) = "Save failed."
    Errors(1) = "Product code is required."
    Errors(2) = "Product name is required."
    Errors(3) = "Product name is already in use."
    Errors(4) = "Probably with an inactive one."
    Errors(5) = "Category is required."
    Errors(6) = "Invalid category."
    Errors(7) = "Invalid Unit Price."
    Errors(8) = "Price must be numeric."
    Errors(9) = "Invalid Unit Cost."
    Errors(10) = "Unit of Measure is required."
    Errors(11) = "Code is already in use."
    Errors(12) = "Numeric data is required."
    Errors(13) = "Customer is required."
    Errors(14) = "Terms is required."
    Errors(15) = "Order number is already in use."
    Errors(16) = "Bank account is required."
    Errors(17) = "No valid payment found."
    Errors(18) = "Name is required."
    Errors(19) = "Name is already in use."
    Errors(20) = "Fund account is required."
    Errors(21) = "Account number is required."
    Errors(22) = "Bank is required."
    Errors(23) = "Account number is already in use."
    Errors(24) = "Amount is required."
    Errors(25) = "Amount is invalid."
    Errors(26) = "Expense is required."
    Errors(27) = "There is already a forwarded balance in this date."
    Errors(28) = "Password did not match."
    Errors(29) = "Invalid username and/or password."
    Errors(30) = "User Name is required."
    Errors(31) = "User Name is already in use."
    Errors(32) = "Check # is required."
    Errors(33) = "Insufficient quantity."
    Errors(34) = "Payment is insufficient."
    Errors(35) = "Delete failed. No item selected."
    Errors(36) = "No items selected."
    Errors(37) = "Please select accounts to pay."
    Errors(38) = "Login failed."
    Errors(39) = "Username and/or password is invalid."
    Errors(40) = "Code is required."
    Errors(41) = "Mark-up is invalid."
    Errors(42) = "Field required."
    Errors(43) = "Invalid data."
    Errors(44) = "User number must be numeric."
    Errors(45) = "Pin must be numeric."
    Errors(46) = "User cannot be deactivated."
    Errors(47) = "User number already in use."
    Errors(48) = "Name already exists."
    Errors(49) = "Password is required."
    Errors(50) = "Tax is required."
    Errors(51) = "Card number is required."
    Errors(52) = "Reference is required."
    Errors(53) = "Card number does not exist."
    Errors(54) = "Card already in use."
    Errors(55) = "Login error. Machine is not registerd in the system."
    Errors(56) = "Invalid user number."
    Errors(57) = "Invalid pin."
    Errors(58) = "Login error. Machine is not activated in the system."
    Errors(59) = "Item does not exists in the purchase order list."
    Errors(60) = "Cannot receive inventory when order is already complete."
    Errors(61) = "Cannot pick inventory when order is already complete."
    Errors(62) = "Cannot pick inventory when order is already invoiced."
    Errors(63) = "Cannot pick inventory when order is already paid or cancelled."
    Errors(64) = "Order is cancelled. No changes made."
    Errors(65) = "User pin not set."
    Errors(66) = "User not allowed."
    Errors(67) = "No more records to display."
    Errors(68) = "Invalid O.R. number."
    Errors(69) = "Child is required."
    Errors(70) = "Attendant is required."
    Errors(71) = "Hours must be greater than 0."
    Errors(72) = "Reorder point must be numeric."
    Errors(73) = "Reorder quantity must be numeric."
    ErrorCodes = Errors(Code)
End Function

Public Function MessageCodes(ByVal Code As Integer) As String
    Dim Message(100) As String
    Message(0) = "saved."
    Message(1) = "Product"
    Message(2) = "deleted."
    Message(3) = "Payments"
    Message(4) = "deactivated."
    Message(5) = "activated."
    Message(6) = "New"
    Message(7) = "Record"
    MessageCodes = Message(Code)
End Function

Public Sub ClearClassData(ByVal info As Integer)
    Select Case info
        Case 0
            With CardInfo
                .Amount = 0
                .BankId = 0
                .CardNumber = ""
                .CardTypeId = 0
                .NameOnCard = ""
                .Reference = ""
            End With
        Case 1
            With CheckInfo
                .Amount = 0
                .BankId = 0
                .CheckDate = Format(Now, "MM/DD/YY")
                .CheckNumber = ""
            End With
        Case 2
            With LoyaltyInfo
                .CardNumber = ""
                .UsePoints = "0.00"
            End With
        Case 3
            With OtherInfo
                .ReferenceNumber = ""
                .Remarks = ""
                .Amount = "0.00"
            End With
    End Select
End Sub
Public Sub SavePOSAuditTrail(ByVal UserId As Integer, ByVal WorkstationId As Integer, _
                ByVal POS_SalesId As String, ByVal Activity As String)
    Dim newcon As ADODB.Connection
    Set newcon = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    newcon.ConnectionString = ConnString
    newcon.Open
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = newcon
    cmd.CommandText = "POS_UserAudit_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Val(POS_SalesId))
    cmd.Parameters.Append cmd.CreateParameter("@Activity", adVarChar, adParamInput, 250, Activity)
    cmd.Execute
    newcon.Close
End Sub



Public Sub GetInventorySettings()
    'Get Settings
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_Settings_Get"
    Set rec = cmd.Execute
    
    If Not rec.EOF Then
        AllowNegativeInventory = rec!AllowNegativeInventory
        ReorderPointCheckScheduled = rec!ReorderPointCheckScheduled
        ReorderPointCheckFrequency = rec!ReorderPointCheckFrequency
        LastReorderPointDateCheck = rec!LastReorderPointDateCheck
    End If
    
    con.Close
End Sub

Public Function CheckAvailableQuantity(ByVal ProductId As String) As Double
    Dim chk_con As New ADODB.Connection
    Dim chkrec As New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    chk_con.ConnectionString = ConnString
    chk_con.Open
    cmd.ActiveConnection = chk_con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_CheckAvailableQuantity"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    Set chkrec = cmd.Execute
    If Not chkrec.EOF Then
        CheckAvailableQuantity = chkrec!AvailableQuantity
    End If
    chk_con.Close
End Function

Public Function ReserveProduct(ByVal ReserveId As String, ByVal ProductId As String, _
    ByVal Quantity As Double, ByVal UserId As Integer, ByVal isPOS As Boolean, _
    Optional SalesOrderId As String = "0", Optional ByVal PurchaseReturnId As String = "0") As String
    
    Dim res_con As New ADODB.Connection
    
    Set cmd = New ADODB.Command
    
    res_con.ConnectionString = ConnString
    res_con.Open
    res_con.BeginTrans
    cmd.ActiveConnection = res_con
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInputOutput, , Val(ReserveId))
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Quantity)
                          cmd.Parameters("@Quantity").NumericScale = 2
                          cmd.Parameters("@Quantity").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@isPOS", adBoolean, adParamInput, , isPOS)
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Val(SalesOrderId))
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , 0)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , Val(PurchaseReturnId))
    If Val(ReserveId) = 0 Then
        cmd.CommandText = "INV_ProductReserve_Insert"
        cmd.Execute
        ReserveProduct = cmd.Parameters("@ReserveId")
    Else
        cmd.CommandText = "INV_ProductReserve_Update"
        cmd.Execute
        ReserveProduct = cmd.Parameters("@ReserveId")
    End If
    res_con.CommitTrans
    res_con.Close
    
End Function
    
Public Sub DeleteReserves(ByVal UserId As Integer, ByVal isPOS As Boolean, ByVal isSalesOrder As Boolean, ByVal isPurchaseReturn As Boolean)
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductReserve_DeleteByUser"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@isPOS", adBoolean, adParamInput, , isPOS)
    cmd.Parameters.Append cmd.CreateParameter("@isSalesOrder", adBoolean, adParamInput, , isSalesOrder)
    cmd.Parameters.Append cmd.CreateParameter("@isPurchaseReturn", adBoolean, adParamInput, , isPurchaseReturn)
    cmd.Execute
    con.Close
End Sub

Public Sub DeleteReserveLine(ByVal ReserveId As String)
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductReserves_Delete"
    cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(ReserveId))
    cmd.Execute
    con.Close
End Sub

Public Sub UpdateReserveQuantity(ByVal ReserveId As String, ByVal Quantity As Double, ByVal ProductId As String, _
        ByVal SalesOrderId As String)
        
    Dim newcon As New ADODB.Connection
    Dim newcmd As New ADODB.Command
    newcon.ConnectionString = ConnString
    
    newcon.Open
    newcmd.ActiveConnection = newcon
    newcmd.CommandType = adCmdStoredProc
    newcmd.CommandText = "INV_ProductReserve_QuantityUpdate"
    newcmd.Parameters.Append newcmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(ReserveId))
    newcmd.Parameters.Append newcmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    newcmd.Parameters.Append newcmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Val(SalesOrderId))
    newcmd.Parameters.Append newcmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Quantity)
                          newcmd.Parameters("@Quantity").NumericScale = 2
                          newcmd.Parameters("@Quantity").Precision = 18
    newcmd.Execute
    newcon.Close
End Sub

Public Function NVAL(ByVal expression As String) As Double
    NVAL = Val(Replace(expression, ",", ""))
End Function

Public Sub GetItemExtraSellingInfo(ByVal ProductId As String)
    isIndefiniteSize = False
    isPricingReference = False
    TrackClipping = False
    
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ProductExtraSellingInfo_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    Set rec = cmd.Execute
    If Not rec.EOF Then
        isIndefiniteSize = rec!isIndefiniteSize
        isPricingReference = rec!isPricingReference
        isLengthOnly = rec!isLengthOnly
        PriceDetails = "@ " & rec!extrainfoprice & "/" & rec!Uom
        TrackClipping = rec!TrackClipping
    Else
        isIndefiniteSize = False
        isPricingReference = False
        isLengthOnly = False
        TrackClipping = False
        PriceDetails = ""
    End If
    con.Close
End Sub
Public Function ProductItemCode(ByVal Name As String) As ADODB.Recordset
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Code"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    'con.Close
    
    Set ProductItemCode = rec
    Set con = Nothing
End Function
Public Function ProductName(ByVal Name As String) As ADODB.Recordset
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Name"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    'con.Close
    
    Set ProductName = rec
    Set con = Nothing
End Function



Public Function ProductBarcode(ByVal Barcode As String) As ADODB.Recordset
    On Error GoTo ErrMessage
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search_Barcode"
    cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 250, Barcode)
    
    Set rec = cmd.Execute
    'con.Close
    
    Set ProductBarcode = rec
    Set con = Nothing
    Exit Function
ErrMessage:
    MsgBox "There was a problem connecting to server. Please try again.", vbCritical
    
End Function

Public Sub LoadImageStatus(ByVal picturebox As picturebox, ByVal Status As String)
    Status = UCase(Status)
    picturebox.Visible = True
    Select Case Status
        Case UCase("open")
            picturebox.Visible = False
        Case ""
            picturebox.Visible = False
        Case Else
            picturebox.Picture = LoadPicture(App.Path & "\images\" & Status & ".jpg")
    End Select
    
End Sub

Public Function GetStatus(ByVal StatusId As Long) As String
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocStatus_Get"
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        GetStatus = rec!Status
    End If
    con.Close
End Function

Public Sub ShowNotification()
    'On Error Resume Next
    BASE_NotificationFrm.Left = (BASE_ContainerFrm.width - BASE_NotificationFrm.width) - 600
    BASE_NotificationFrm.Top = 0
    BASE_NotificationFrm.ZOrder 0
    BASE_NotificationFrm.Show
    BASE_NotificationFrm.ZOrder 0
End Sub

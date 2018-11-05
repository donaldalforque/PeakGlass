VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RPT_PO_PurchaseOrderSummary 
   Caption         =   "Purchase Order Summary"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtOR 
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
         TabIndex        =   17
         Top             =   1320
         Width           =   2415
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
         ItemData        =   "RPT_PO_PurchaseOrderSummary.frx":0000
         Left            =   1320
         List            =   "RPT_PO_PurchaseOrderSummary.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "RPT_PO_PurchaseOrderSummary.frx":002F
         Left            =   1320
         List            =   "RPT_PO_PurchaseOrderSummary.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtTitle 
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
         TabIndex        =   3
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate Report"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   4560
         Width           =   1815
      End
      Begin VB.ComboBox cmbSort 
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
         ItemData        =   "RPT_PO_PurchaseOrderSummary.frx":005E
         Left            =   1320
         List            =   "RPT_PO_PurchaseOrderSummary.frx":0074
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3480
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1320
         TabIndex        =   13
         Top             =   2400
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   94830593
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1320
         TabIndex        =   16
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   94830593
         CurrentDate     =   41686
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR #"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label8 
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
         TabIndex        =   15
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By"
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
         TabIndex        =   10
         Top             =   120
         Width           =   1005
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
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
         TabIndex        =   7
         Top             =   3000
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title"
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
         TabIndex        =   6
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By"
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
         TabIndex        =   5
         Top             =   3480
         Width           =   645
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   11295
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "RPT_PO_PurchaseOrderSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Public Sub Populate(ByVal Data As String)
    Dim item As MSComctlLib.ListItem
    Select Case Data
        Case "Status"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Status")
            cmbStatus.Clear
            cmbStatus.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbStatus.AddItem rec!Status
                    cmbStatus.ItemData(cmbStatus.NewIndex) = rec!StatusId
                    rec.MoveNext
                Loop
            End If
            cmbStatus.ListIndex = 0
'        Case "Terms"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("Terms")
'            cmbTerms.Clear
'            cmbTerms.AddItem ""
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    If rec!isActive = "True" Then
'                        cmbTerms.AddItem rec!Terms
'                        cmbTerms.ItemData(cmbTerms.NewIndex) = rec!TermId
'                    End If
'                    rec.MoveNext
'                Loop
'            End If
'            cmbTerms.ListIndex = 0
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbVendor.AddItem rec!Name
                        cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbVendor.ListIndex = 0
    End Select
End Sub

Private Sub btnGenerate_Click()
    Dim sql, OrderBy As String
    Dim Status, Vendor, Terms, DateRange As Variant
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\PurchaseOrderSummary.rpt")
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
    
    Select Case cmbStatus.ListIndex
        Case 0
            Status = ""
        Case Else
            Status = "AND GLOBAL_DocStatus.StatusId = '" & cmbStatus.ItemData(cmbStatus.ListIndex) & "' "
    End Select
    
    Select Case cmbVendor.ListIndex
        Case 0
            Vendor = ""
        Case Else
            Vendor = "AND BASE_Vendor.VendorId = '" & cmbVendor.ItemData(cmbVendor.ListIndex) & "' "
    End Select
    
'    Select Case cmbTerms.ListIndex
'        Case 0
'            Terms = ""
'        Case Else
'            Terms = "AND BASE_Terms.TermId = '" & cmbTerms.ItemData(cmbTerms.ListIndex) & "' "
'    End Select
    
    'If optOrderDate.value = True Then
        DateRange = "AND PO_PurchaseOrder.Date >= '" & DateFrom.value & " 00:00:00' " & _
                    "AND PO_PurchaseOrder.Date <= '" & DateTo.value & " 23:23:59' "
    'Else
'        Select Case cmbDueDate.ListIndex
'            Case 0
'                DateRange = ""
'            Case 1
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 15 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 2
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 30 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 3
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 45 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 4
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 60 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 5
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 90 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 6
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) <= 120 " & _
'                            "AND DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) > 0 "
'            Case 7
'                DateRange = "DATEDIFF(day,GETDATE(),PO_PurchaseOrder.DueDate) >= 120 "
'            Case 8
'                DateRange = "PO_PurchaseOrder.DueDate <= GETDATE() "
'        End Select
    'End If
    
    Select Case cmbSort.ListIndex
        Case 0
            OrderBy = "ORDER BY PO_PurchaseOrder.OrderNumber ASC"
        Case 1
            OrderBy = "ORDER BY PO_PurchaseOrder.OrderNumber ASC"
        Case 2
            OrderBy = "ORDER BY GLOBAL_DocStatus.Status ASC"
        Case 3
            OrderBy = "ORDER BY PO_PurchaseOrder.Date DESC"
        Case 4
            OrderBy = "ORDER BY PO_PurchaseOrder.Total ASC"
        Case 5
            OrderBy = "ORDER BY PO_PurchaseOrder.Balance DESC"
    End Select
    
    sql = "SELECT PO_PurchaseOrder.OrderNumber,PO_PurchaseOrder.Date,PO_PurchaseOrder.DueDate," & _
          "PO_PurchaseOrder.total , PO_PurchaseOrder.OutStandingbalance, GLOBAL_DocStatus.Status, BASE_Vendor.Name, BASE_Terms.Terms " & _
          "FROM ((Peak_GM.dbo.PO_PurchaseOrder PO_PurchaseOrder INNER JOIN Peak_GM.dbo.GLOBAL_DocStatus GLOBAL_DocStatus ON " & _
          "PO_PurchaseOrder.StatusId = GLOBAL_DocStatus.StatusId) INNER JOIN Peak_GM.dbo.BASE_Vendor BASE_Vendor ON " & _
          "PO_PurchaseOrder.VendorId = BASE_Vendor.VendorId) LEFT OUTER JOIN Peak_GM.dbo.BASE_Terms BASE_Terms ON " & _
          "PO_PurchaseOrder.TermId = BASE_Terms.TermId WHERE PO_PurchaseOrder.OrderNumber LIKE '" & txtOR.text & "%' " & _
          DateRange & Terms & Status & Vendor & OrderBy
    
    sql = Replace(sql, "WHERE AND", "WHERE ")
    sql = Replace(sql, "WHERE ORDER BY", "ORDER BY")
    
    crxRpt.ParameterFields(1).AddCurrentValue txtTitle.text
    'If optOrderDate.value = True Then
        crxRpt.ParameterFields(2).AddCurrentValue Str(DateFrom.value)
        crxRpt.ParameterFields(3).AddCurrentValue Str(DateTo.value)
    'End If
    
    crxRpt.SQLQueryString = sql
    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub

Private Sub Form_Load()
    cmbStatus.ListIndex = 0
    cmbSort.ListIndex = 0
    Populate "Status"
    Populate "Vendor"
    Populate "Terms"
    
    
    Me.Height = 9390
    Me.width = 15180
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    
    txtTitle.text = Me.Caption
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub







VERSION 5.00
Begin VB.Form frmADOX 
   Caption         =   "ADO Access 2000 Database Creation Utility"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatabase 
      Height          =   735
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   375
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label12 
         Caption         =   "Database Name"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   280
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   ".mdb"
         Height          =   255
         Left            =   3840
         TabIndex        =   38
         Top             =   280
         Width           =   375
      End
   End
   Begin VB.Frame fraTable 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   6975
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   5
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   5
         Left            =   4920
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   6
         Left            =   4920
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   7
         Left            =   4920
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   8
         Left            =   4920
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   9
         Left            =   4920
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtTable 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   5655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   21
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   20
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   19
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   16
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   15
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   14
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Field 1"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Field 2"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Field 3"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Field 4"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Field 5"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Field 6"
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Field 7"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Field 8"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Field 9"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Field 10"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Table Name"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Menu mExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmADOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim catDB As ADOX.Catalog
Dim tblNew As ADOX.Table
Dim CnnString As String
Option Explicit

Private Sub cmdClear_Click()
    x = 0
    While x <= 9
        txtTable.Text = ""
        Check1.Item(x).Value = 0
        txtFld.Item(x).Text = ""
        x = x + 1
    Wend
End Sub

Private Sub cmdAdd_Click()
    Set tblNew = New ADOX.Table
    tblNew.Name = txtTable.Text
    
    x = 0
    While x <= 9
        If Check1.Item(x).Value = 1 Then
            tblNew.Columns.Append txtFld.Item(x).Text, adVarWChar
        End If
        x = x + 1
    Wend
           
    catDB.Tables.Append tblNew
    
    MsgBox "Table " & txtTable.Text & " was successfully added into " & txtDatabase.Text & ".mdb", vbOKOnly, "ADO Extension"
End Sub

Private Sub cmdCreate_Click()
    CnnString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" _
                & App.Path & "\" & txtDatabase.Text & ".mdb"
    
    Set catDB = New ADOX.Catalog
    catDB.Create CnnString
    catDB.ActiveConnection = CnnString
    
    fraDatabase.Enabled = False
    fraTable.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set catDB = Nothing
End Sub

Private Sub mAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub

Private Sub mExit_Click()
    Set catDB = Nothing
    End
End Sub

Private Sub txtFld_Change(Index As Integer)
    If txtFld.Item(Index) = "" Then
        Check1.Item(Index).Value = 0
    Else
        Check1.Item(Index).Value = 1
    End If
End Sub

Private Sub txtFld_GotFocus(Index As Integer)
    Call SelectAll(txtFld.Item(Index))
End Sub

Private Sub txtTable_GotFocus()
    Call SelectAll(txtTable)
End Sub

Private Sub txtDatabase_GotFocus()
    Call SelectAll(txtDatabase)
End Sub

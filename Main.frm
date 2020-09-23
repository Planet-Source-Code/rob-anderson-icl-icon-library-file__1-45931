VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICL Reader"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export Icon"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1590
      TabIndex        =   18
      Top             =   4980
      Width           =   1245
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import Icon"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1620
      TabIndex        =   17
      Top             =   4530
      Width           =   1245
   End
   Begin VB.TextBox txtImport 
      Height          =   315
      Left            =   150
      TabIndex        =   14
      Top             =   1620
      Width           =   4065
   End
   Begin VB.TextBox txtExport 
      Height          =   315
      Left            =   150
      TabIndex        =   13
      Top             =   2220
      Width           =   4065
   End
   Begin VB.CommandButton cmdSaveLibrary 
      Caption         =   "&Save Library"
      Enabled         =   0   'False
      Height          =   405
      Left            =   210
      TabIndex        =   12
      Top             =   4980
      Width           =   1245
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Library"
      Height          =   405
      Left            =   210
      TabIndex        =   11
      Top             =   4530
      Width           =   1245
   End
   Begin VB.HScrollBar hsbImage 
      Height          =   255
      Left            =   1830
      TabIndex        =   8
      Top             =   3510
      Width           =   1005
   End
   Begin VB.HScrollBar hsbIcon 
      Height          =   255
      Left            =   1830
      TabIndex        =   6
      Top             =   2910
      Width           =   1005
   End
   Begin VB.TextBox txtSaveName 
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   990
      Width           =   4065
   End
   Begin VB.TextBox txtICLToOpen 
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   390
      Width           =   4065
   End
   Begin VB.PictureBox pctImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1475
      Left            =   210
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   2910
      Width           =   1475
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1860
      TabIndex        =   19
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon to Import..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   16
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon to Export..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   150
      TabIndex        =   15
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image in Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1860
      TabIndex        =   10
      Top             =   3870
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image in Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1830
      TabIndex        =   9
      Top             =   3270
      Width           =   975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon in Library"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1830
      TabIndex        =   7
      Top             =   2670
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save Icon Library As..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   750
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Library to Open..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   150
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   1
      Top             =   2640
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjICL As ICL

Private Sub Form_Load()

    txtICLToOpen.Text = App.Path & "\test.icl"
    txtSaveName.Text = App.Path & "\testsave.icl"
    txtImport.Text = App.Path & "\import.ico"
    txtExport.Text = App.Path & "\export.ico"
    
    lblDesc.Caption = vbNullString
    lblName.Caption = vbNullString
    
    hsbIcon.Max = 1
    hsbImage.Max = 1
    
    hsbIcon.Min = 1
    hsbImage.Min = 1
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim frmTest As Form

    If Not (mobjICL Is Nothing) Then
        Call mobjICL.Term
        Set mobjICL = Nothing
    End If

    End

End Sub

Private Sub cmdOpen_Click()

    If Not (mobjICL Is Nothing) Then
        Call mobjICL.Term
        Set mobjICL = Nothing
    End If
    
    Set mobjICL = New ICL
    Call mobjICL.Init

    If (Len(txtICLToOpen.Text) = 0) Then
        MsgBox "Library name not present", vbInformation, "Open Library"
        txtICLToOpen.SetFocus
        Exit Sub
    End If
    
    If (Len(Dir$(txtICLToOpen.Text)) = 0) Then
        MsgBox "Library file not present", vbInformation, "Open Library"
        txtICLToOpen.SetFocus
        Exit Sub
    End If
    
    mobjICL.Filename = txtICLToOpen.Text
    Call mobjICL.ReadLibrary

    cmdSaveLibrary.Enabled = True
    cmdImport.Enabled = True
    cmdExport.Enabled = True
    
    hsbIcon.Max = mobjICL.IconCount
    Call hsbIcon_Change
    
End Sub

Private Sub cmdImport_Click()

    If (Len(txtImport.Text) = 0) Then
        MsgBox "Icon name not present", vbInformation, "Import Icon"
        txtImport.SetFocus
        Exit Sub
    End If
    
    If (Len(Dir$(txtImport.Text)) = 0) Then
        MsgBox "Icon file not present", vbInformation, "Import Icon"
        txtImport.SetFocus
        Exit Sub
    End If
    
    Call mobjICL.AddIcon(txtImport.Text)

    hsbIcon.Max = mobjICL.IconCount
    Call hsbIcon_Change

End Sub

Private Sub cmdExport_Click()

    If (Len(txtExport.Text) = 0) Then
        MsgBox "Export name not present", vbInformation, "Export Icon"
        txtExport.SetFocus
        Exit Sub
    End If

    If (Len(Dir$(txtExport.Text)) > 0) Then
        If MsgBox("Save file exists. Overwrite?", vbYesNo, "Export Icon") = vbNo Then
            txtExport.SetFocus
            Exit Sub
        End If
    End If
    
    Call mobjICL.ExportIcon(hsbIcon.Value, txtExport.Text)

    MsgBox txtExport.Text & " export successfully!", vbExclamation, "Export Icon"
    
End Sub

Private Sub cmdSaveLibrary_Click()

    If (Len(txtSaveName.Text) = 0) Then
        MsgBox "Save name not present", vbInformation, "Save Library"
        txtSaveName.SetFocus
        Exit Sub
    End If

    If (Len(Dir$(txtSaveName.Text)) > 0) Then
        If MsgBox("Save file exists. Overwrite?", vbYesNo, "Save Library") = vbNo Then
            txtSaveName.SetFocus
            Exit Sub
        End If
    End If

    Call mobjICL.WriteLibrary(txtSaveName.Text)

    MsgBox txtSaveName.Text & " saved successfully!", vbExclamation, "Save Library"

End Sub

Private Sub hsbIcon_Change()

    If Not (mobjICL Is Nothing) Then
        hsbImage.Max = mobjICL.ImageCount(hsbIcon.Value)
        Call hsbImage_Change
        lblName.Caption = mobjICL.IconName(hsbIcon.Value)
    End If
    
End Sub

Private Sub hsbIcon_Scroll()
    Call hsbIcon_Change
End Sub

Private Sub hsbImage_Change()
    
    If Not (mobjICL Is Nothing) Then
        Set Me.pctImage = mobjICL.IconPicture(pctImage.hDC, hsbIcon.Value, hsbImage.Value)
        lblDesc.Caption = mobjICL.ImageWidth(hsbIcon.Value, hsbImage.Value) & "x" & _
                          mobjICL.ImageHeight(hsbIcon.Value, hsbImage.Value) & ", " & _
                          mobjICL.ColorDepth(hsbIcon.Value, hsbImage.Value) & "-bit"
    End If
    
End Sub

Private Sub hsbImage_Scroll()
    Call hsbImage_Change
End Sub



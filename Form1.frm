VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Licence maker - by Wouter vas Koppen (Please vote!!!)"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLicence 
      Height          =   3015
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore to no licence!"
         Height          =   735
         Left            =   1320
         TabIndex        =   5
         Top             =   1140
         Width           =   1815
      End
   End
   Begin VB.Frame fraNoLicence 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate a licence number for ONLY this computer"
         Height          =   735
         Left            =   1260
         TabIndex        =   3
         Top             =   780
         Width           =   1815
      End
      Begin VB.TextBox txtKeycode 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   4095
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register!"
         Height          =   495
         Left            =   2340
         TabIndex        =   1
         Top             =   2220
         Width           =   1335
      End
   End
   Begin VB.Label lblMail 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mail: Xbrain3000@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4620
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   3300
      Width           =   4275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As New clsConvert
'Author: Wouter van Koppen
'I'd like to thanx Fred T who helps me with writing
'this code
'Mail: Xbrain3000@hotmail.com

Private Sub cmdGenerate_Click()
    txtKeycode.Text = X.UsercodeToKeycode(X.GetMACAddress(0))
End Sub

Private Sub cmdRegister_Click()
    If X.UsercodeToKeycode(X.GetMACAddress(0)) = txtKeycode Then
        X.CreateKey "HKCU\Software\NerdApp\Licence", txtKeycode.Text
        MsgBox "Licenced!", vbInformation
        Form_Load
    Else
        MsgBox "Invalid Keycode!", vbExclamation
    End If
End Sub

Private Sub cmdRestore_Click()
    X.DeleteKey "HKCU\Software\NerdApp\"
    MsgBox "NOT licenced!", vbInformation
    Form_Load
End Sub

Private Sub Form_Load()
    If X.ReadKey("HKCU\Software\NerdApp\Licence") = X.UsercodeToKeycode(X.GetMACAddress(0)) Then
        fraLicence.Visible = True
        fraNoLicence.Visible = False
    Else
        fraLicence.Visible = False
        fraNoLicence.Visible = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = &H0&
End Sub

Private Sub lblMail_Click()
    Shell "start mailto:Xbrain3000@hotmail.com"
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = &HFF0000
End Sub

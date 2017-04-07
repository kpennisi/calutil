VERSION 5.00
Begin VB.Form frmWait 
   Caption         =   "Loading ..."
   ClientHeight    =   924
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5460
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   924
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   852
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5412
      Begin VB.Timer tmrOne 
         Interval        =   2000
         Left            =   4800
         Top             =   600
      End
      Begin VB.Label lblMain 
         Caption         =   "Please wait while the utility gathers information from Outlook ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5052
      End
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   'Unload Me
End Sub

Private Sub Form_Load()
   'Call frmMain.getFolders
   'frmMain.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmWait = Nothing
End Sub

Private Sub tmrOne_Timer()
   frmMain.Show
End Sub

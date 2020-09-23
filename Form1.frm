VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   4935
   ClientTop       =   3825
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   3390
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36873
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36873
   End
   Begin VB.Label Label2 
      Caption         =   "End:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Begin:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtpBegin_Change()
  If dtpEnd.Value <= dtpBegin.Value Then dtpEnd.Value = later(dtpBegin.Value)
End Sub

Private Sub dtpEnd_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
  If dtpEnd.Value <= dtpBegin.Value Then dtpEnd.Value = later(dtpBegin.Value)
End Sub

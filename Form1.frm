VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2895
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Jerimiah.Barsalou@NCR.com"
      Top             =   120
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   2143
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "RIGHT CLICK FORM TO EXIT"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim FormLoad As Boolean

Private Sub Form_Load()
    WebBrowser1.Navigate "about:<HTML><BODY scroll='no'><BODY TOPMARGIN='0' LEFTMARGIN='0' MARGINWIDTH='0' MARGINHEIGHT='0'><A href='mailto:Jerimiah.Barsalou@NCR.com'><img border='0' src='" & App.Path & "\Mail.gif'></A></BODY></HTML>"
    WebBrowser2.Navigate "about:<HTML><BODY scroll='no'><BODY TOPMARGIN='0' LEFTMARGIN='0' MARGINWIDTH='0' MARGINHEIGHT='0'><A href=''><img border='0' src='" & App.Path & "\Floppy.gif'></A></BODY></HTML>"
    FormLoad = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As POINTAPI
    If Button = 2 Then
        pt.X = ScaleX(X, vbTwips, vbPixels)
        pt.Y = ScaleY(Y, vbTwips, vbPixels)
        frmEnd.Show
        frmEnd.Move ScaleX(pt.X, vbPixels, vbTwips) + Form1.Left, ScaleX(pt.Y, vbPixels, vbTwips) + Form1.Top
    End If
End Sub

Private Sub WebBrowser2_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    If FormLoad = False Then MsgBox "CLICK EVENT"
    FormLoad = False
End Sub

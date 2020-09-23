VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmEnd 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   5115
   ClientTop       =   3480
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      ExtentX         =   4048
      ExtentY         =   1931
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
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WebBrowser1.Navigate "about:<HTML><HEAD><SCRIPT LANGUAGE=JAVASCRIPT>function closewindow() {self.opener = this; self.close()}</SCRIPT></HEAD><BODY scroll='no'><BODY TOPMARGIN='0' LEFTMARGIN='0' MARGINWIDTH='0' MARGINHEIGHT='0'><A href='javascript:closewindow();'><img border='0' src='" & App.Path & "\Exit.gif'></A></BODY></HTML>"
End Sub

Private Sub WebBrowser1_LostFocus()
    Unload Me
End Sub

Private Sub WebBrowser1_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    Dim X As Integer
    X = MsgBox("EXIT CONFIRMATION?", vbOKCancel, "EXIT")
    If X = 1 Then End
End Sub

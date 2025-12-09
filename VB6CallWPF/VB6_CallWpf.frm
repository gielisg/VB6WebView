VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Web Frame"
      Height          =   1095
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1.frm
Option Explicit

Private Sub Form_Load()
    Me.Caption = "VB6 Host - Open WPF Web Browser"
    cmdOpen.Caption = "Open Web Frame"
    cmdOpen.Width = 1500 ' optional to make button bigger
    cmdOpen.Height = 400
End Sub

Private Sub cmdOpen_Click()
    ' Late-binding: no references required
    Dim host As Object
    On Error GoTo ErrHandler

    ' Create the COM object registered by your .NET WPF DLL
    Set host = CreateObject("WpfBrowserLib.BrowserHost")
    If host Is Nothing Then
        MsgBox "Failed to create WpfBrowserLib.BrowserHost. Is the DLL registered?", vbCritical
        Exit Sub
    End If

    ' Show the browser modal (passes default URL if empty)
    ' This call blocks until the WPF window is closed.
    host.ShowBrowser

    ' Clean up
    Set host = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error creating or calling WpfBrowserLib.BrowserHost:" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    Set host = Nothing
End Sub



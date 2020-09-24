VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Command Line Args"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMsg 
      Caption         =   "     Click Me To Display Command     Line Arguments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do drop me a line if you liked this code, that will encourage me
' write more such stuff. Send me an email at udgama@rocketmail.com
' if you have any problems or question regarding this code.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' This routine is used for illustrating how to get the command line arguments
' and use them in you program. Using this method you can Command line Enable your
' visual basic application.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdMsg_Click()
On Error GoTo EH
Dim CmdLineStr As String
Dim lFirstArg  As String
Dim lSecondArg As String
Dim lPos      As Integer

   ' Get the whole command line argument.
   ' Command$ Returns the argument portion of the command line
   CmdLineStr = Command$
   If Trim(CmdLineStr) = "" Then
      MsgBox "No command line arguments for this program.", vbInformation, "udgama@rocketmail.com"
   End If
   
   ' Now parse the command line for the first argument
   lPos = InStr(1, CmdLineStr, " ")
   If lPos = 0 Then
      ' If ipos is zero then there was only one argument found
      lFirstArg = Trim(CmdLineStr)
   Else
      lFirstArg = Mid(CmdLineStr, 1, lPos)
      ' Now parse the command line for the second argument
      lSecondArg = Mid(CmdLineStr, lPos, Len(CmdLineStr))
   End If
   
   ' Now that you got both your command line arguments do what you
   ' like with them.
   MsgBox "The First Argument is :=" & lFirstArg & vbCrLf & _
      "The Second Argument is :=" & lSecondArg & vbCrLf, vbInformation, "udgama@rocketmail.com"
   
   Exit Sub
EH:
   MsgBox Err.Description, vbInformation, "udgama@rocketmail.com"
   Exit Sub
End Sub

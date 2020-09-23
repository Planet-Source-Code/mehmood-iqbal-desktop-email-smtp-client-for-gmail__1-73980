VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Dialog2 
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Desktop_Email_Client.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5535
      Caption         =   "Connecting to smtp.gmail.com on port 465. . . . ."
      Size            =   "9763;661"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Size            =   "10186;2566"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Timer1_Timer()

'Attached files counter veriables
Dim i As Integer
Dim Index

'Email Sending veriables
Dim iMsg
Dim iConf
Dim Flds
Dim Schema

'Setting Progressbar & Sending Status
Dialog2.XP_ProgressBar1.Value = 20
Dialog2.Label2.Caption = "Connecting to the server . . . . ."

On Error GoTo SendMail_Error:

Set iMsg = CreateObject("CDO.Message")

Set iConf = CreateObject("CDO.Configuration")

Set Flds = iConf.Fields

Schema = "http://schemas.microsoft.com/cdo/configuration/"

Flds.Item(Schema & "sendusing") = 2

'Server Address (Must be Smtp.gmail.com)
Flds.Item(Schema & "smtpserver") = Main_Form.Label10.Caption

'Server Port (465)
Flds.Item(Schema & "smtpserverport") = Main_Form.Label11.Caption

Dialog2.XP_ProgressBar1.Value = 30

'Athentication type
Flds.Item(Schema & "smtpauthenticate") = 1

'Gmail complete address as Username
Flds.Item(Schema & "sendusername") = Main_Form.TextBox9.Text

'Gmail ID password
Flds.Item(Schema & "sendpassword") = Main_Form.TextBox10.Text

'Connection timeout
Flds.Item(Schema & "smtpConnectionTimeout") = 40

'SSL setting
Flds.Item(Schema & "smtpusessl") = 1

Flds.Update

'Show progress of sending
Dialog2.XP_ProgressBar1.Value = 50
Dialog2.Label2.Caption = "Please wait while sending email . . . . ."

'Setting-up email perameters
With iMsg
   .To = Main_Form.TextBox2.Text & "<" & Main_Form.TextBox3.Text & ">"
   .From = Main_Form.TextBox4.Text & "<" & Main_Form.TextBox5.Text & ">"
   .CC = Main_Form.TextBox6.Text
   .Bcc = Main_Form.TextBox7.Text
   .Subject = Main_Form.TextBox8.Text
   
    'E-mail Text-body
    Dialog2.XP_ProgressBar1.Value = 60
   .TextBody = Main_Form.TextBox1.Text
   
'Check If Files attached then send them one by one
For Index = 0 To Main_Form.ListBox1.ListCount - 1

   If Main_Form.ListBox1.ListCount = 0 Then
   
            'If files not attached
            Dialog2.XP_ProgressBar1.Value = 80
            GoTo Leave_Attachents:
   
   Else
   
            'If files(s) attached then
            Dialog2.XP_ProgressBar1.Value = 75
            Dialog2.Label2.Caption = "Please wait while sending email with atttachments . . . . ."
           .AddAttachment (Main_Form.ListBox1.List(Index))
         
  
   End If
   
Next

Leave_Attachents:

'Send all
Set .Configuration = iConf
   .Send

End With

'Clear veriables if needed
Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing
Set Schema = Nothing

'Email sent, then progress
Dialog2.XP_ProgressBar1.Value = 90

'Hide progress-bar form & show sucess message
Dialog2.Hide

'Email sent with all attachments, then progress bar value
Dialog2.XP_ProgressBar1.Value = 100

'Show sucess message
MsgBox "Email(s) sent sucessfully.", vbInformation, "Sucess"

'Disable sending timer
Dialog2.Timer1.Enabled = False

'Call Clear fields function
Clear_Data

GoTo End_Sub:

'If an error takes, then show a error message
SendMail_Error:

'An error occured when sending email.
Dialog2.Hide
Dialog2.Timer1.Enabled = False
MsgBox "An error occoured when sending email. Please try again." & vbCrLf & "And please also check for right gmail address & password.", vbOKOnly + vbCritical, "Error"

End_Sub:
End Sub
Private Sub Clear_Data()

'Clear files list & disable delete button
Main_Form.ListBox1.Clear
Main_Form.CommandButton3.Enabled = False

'Clear Main form's data fields
Main_Form.TextBox1.Text = ""
Main_Form.TextBox2.Text = ""
Main_Form.TextBox3.Text = ""
Main_Form.TextBox6.Text = ""
Main_Form.TextBox7.Text = ""
Main_Form.TextBox8.Text = ""

'Update attached files counter label
Main_Form.Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."

End Sub

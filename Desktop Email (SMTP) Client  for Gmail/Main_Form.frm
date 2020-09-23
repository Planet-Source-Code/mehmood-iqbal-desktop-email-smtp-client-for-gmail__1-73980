VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Email (SMTP) Client For Gmail"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin Desktop_Email_Client.jcbutton CommandButton3 
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   13
      font            =   "Main_Form.frx":0000
      backcolor       =   0
      caption         =   "Delete"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin Desktop_Email_Client.jcbutton CommandButton2 
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   13
      font            =   "Main_Form.frx":0028
      backcolor       =   0
      caption         =   "Attach Files"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin Desktop_Email_Client.jcbutton CommandButton1 
      Height          =   615
      Left            =   3000
      TabIndex        =   24
      Top             =   6360
      Width           =   2175
      _extentx        =   3836
      _extenty        =   1085
      buttonstyle     =   13
      font            =   "Main_Form.frx":0050
      backcolor       =   0
      caption         =   "Send Email"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.Label Label13 
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   600
      Width           =   855
      Caption         =   "Password :"
      Size            =   "1508;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Width           =   1095
      Caption         =   "Email Address :"
      Size            =   "1931;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   8040
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   8040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   8040
      Y1              =   3240
      Y2              =   3240
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   6495
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "11456;1296"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox10 
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   480
      Width           =   2655
      VariousPropertyBits=   746604571
      Size            =   "4683;661"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox9 
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   480
      Width           =   2775
      VariousPropertyBits=   746604571
      Size            =   "4895;661"
      SpecialEffect   =   3
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label11 
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   120
      Width           =   375
      Caption         =   "465"
      Size            =   "661;450"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   1575
      Caption         =   "SMTP.Gmail.Com"
      Size            =   "2778;450"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   3360
      Width           =   5175
      Size            =   "9128;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
      Caption         =   "Attachments :"
      Size            =   "2143;450"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox8 
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   2760
      Width           =   7095
      VariousPropertyBits=   746604571
      Size            =   "12515;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   855
      Caption         =   "Subject :"
      Size            =   "1508;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox7 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   7095
      VariousPropertyBits=   746604571
      Size            =   "12515;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   615
      Caption         =   "BCC :"
      Size            =   "1085;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox6 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2040
      Width           =   7095
      VariousPropertyBits=   746604571
      Size            =   "12515;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   495
      Caption         =   "CC :"
      Size            =   "873;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   1680
      Width           =   495
      Caption         =   "Email :"
      Size            =   "873;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox5 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
      VariousPropertyBits=   746604573
      Size            =   "5106;661"
      SpecialEffect   =   3
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox4 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
      VariousPropertyBits=   746604571
      Size            =   "4895;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
      VariousPropertyBits=   746604571
      Size            =   "5106;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
      VariousPropertyBits=   746604571
      Size            =   "4895;661"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
      Caption         =   "Receiver's Name :"
      Size            =   "2566;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   495
      Caption         =   "Email :"
      Size            =   "873;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
      Caption         =   "Sender's Name :"
      Size            =   "2143;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   7815
      VariousPropertyBits=   -1399830501
      ScrollBars      =   2
      Size            =   "13785;2778"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'On some requests when same project uploaded on PSC with  '
'the name of 'Urdu Desktop (SMTP) Email Client for Gmail' '
'Some users required to get the english version of the    '
'same project then it is converted by me to english &     '
'being uploaded on Planet Source Code.                    '
'There is a desktop email client to send email using gmail'
'address. No need to go on Gmail's website & login to send'
'emails. Simply put you Gmail address & Password & send   '
'email with the (n) number of atachments. No size limit to'
'attach files with email. All files with large size will  '
'be easily sent by this desktop client. CC & BCC function '
'also support you to send emails with large attachments,  '
'to (n) number of receivers. Gmail's SMTP address & port  '
'fixed in this client. Keep in mind that this email client'
'only designed to work with Gmail. No other email service '
'provider checked with this email client and that may take'
'errors, if you try to do that. This project may need more'
'attention but at the start, i think it is enough to use. '
'                                                         '
' Waiting for your Feedbacks.Thank You.                   '
'                                                         '
'                                                         '
'                 Muhammad Mehmood Iqbal                  '
'                   ME_IQ_TM@yahoo.com                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CommandButton1_Click()

'If an email already sent then reset progressbar
If Dialog2.XP_ProgressBar1.Value = 100 Then

'Set progress value to zero
Dialog2.XP_ProgressBar1.Value = 0

End If

'Check if data is incomplete

'Gmail address not entered
If TextBox9.Text = "" Then
MsgBox "Please enter Gmail address.", vbExclamation, "Gmail Address"

'Password not entered
ElseIf TextBox10.Text = "" Then
MsgBox "Please enter Gmail address password.", vbExclamation, "Password"

'Receiver's email address not entered
ElseIf TextBox3.Text = "" Then
MsgBox "Please enter Receiver's email address.", vbExclamation, "Receiver's Email"

'Subject of email is also compulsory
ElseIf TextBox8.Text = "" Then
MsgBox "Please enter email subject.", vbExclamation, "Subject"

Else
    
    'Check for a valid gmail address
    If InStr(1, TextBox9.Text, "@gmail.com") < 1 Then
  
  
        MsgBox "Email address is not a valid gmail address. Please enter correct gmail address like sample@gmail.com.", vbOKOnly + vbCritical, App.Title
        TextBox9.SetFocus
        Exit Sub
        
    'Check for a valid receiver's email address
    ElseIf InStr(1, TextBox3.Text, "@") < 1 Then
  
        MsgBox "Receiver's email address is not a valid email address. Please enter correct address like sample@sampleserver.com.", vbOKOnly + vbCritical, App.Title
        TextBox3.SetFocus
        Exit Sub
        
    Else
    
    'If addresses are Ok, then trim spaces
     Trim_Functions
     
    'Start Timer to start sending email
    Dialog2.Timer1.Enabled = True

    'Show progressing form
     Dialog2.Show vbModal
     
   End If
   
End If


End Sub

Private Sub CommandButton2_Click()

  Dim File_Name
  
  'Set Dialogbox Title
  CommonDialog1.DialogTitle = "Select File to Attach"

  ' Set flags
  CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
  
  ' Set filters
  CommonDialog1.Filter = "All Files (*.*)"
  
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  
  If CommonDialog1.CancelError = True Or CommonDialog1.FileName = "" Then
  GoTo Exit_Sub
  Else
  
  'Count attached file show them
  ListBox1.AddItem CommonDialog1.FileName
  Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."
  
  'Enable delete button
  CommandButton3.Enabled = True
  
  CommonDialog1.FileName = ""
  
  End If
   
Exit_Sub:
End Sub

Private Sub CommandButton3_Click()

Dim Selected_Item

'Check if no-item seleted in listbox
ListBox1.SetFocus
If ListBox1.ListIndex = -1 Then
GoTo End_Sub:

ElseIf ListBox1.ListIndex >= 0 Then

'Delete selected item
Selected_Item = ListBox1.ListIndex
ListBox1.RemoveItem (Selected_Item)

'Count -1 from attached files
Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."

End If

'If no file in Listbox then disable Delete Button
If ListBox1.ListCount = 0 Then
CommandButton3.Enabled = False
ListBox1.SetFocus

End If

End_Sub:
End Sub

Private Sub Form_Load()


'Disable delete button of attached file
CommandButton3.Enabled = False

'Set attached file's status
Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Response1 As Integer

'If Form canceled with Close button
Cancel = 1

'Confirm if user wants or not
Response1 = MsgBox("Do you really wants to close program?", vbQuestion + vbYesNo, "Confirmation")


If Response1 = vbYes Then

'If user wants then
End

Else

'Else Close message
End If

End Sub

Private Sub TextBox9_Change()

'Automatically set sender email address
TextBox5.Text = TextBox9.Text

End Sub

Private Sub Trim_Functions()

'Setting veriables to trim text of textboxes
Dim Gmail_Address As String
Dim Receiver_email As String
Dim Subject As String
Dim Sender_name As String
Dim Receiver_name As String
Dim CC As String
Dim Bcc As String

'Trimming all
Gmail_Address = Trim$(TextBox9.Text)
Receiver_email = Trim$(TextBox3.Text)
Subject = Trim$(TextBox8.Text)
Sender_name = Trim$(TextBox4.Text)
Receiver_name = Trim$(TextBox2.Text)
CC = Trim$(TextBox6.Text)
Bcc = Trim$(TextBox7.Text)

'Putting trimed text back to textboxes
Main_Form.TextBox9.Text = Gmail_Address
Main_Form.TextBox3.Text = Receiver_email
Main_Form.TextBox8.Text = Subject
Main_Form.TextBox4.Text = Sender_name
Main_Form.TextBox2.Text = Receiver_name
Main_Form.TextBox6.Text = CC
Main_Form.TextBox7.Text = Bcc


End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secret Messenger"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox pass 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Width           =   4815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5760
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Put Message"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Message"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4935
      Left            =   6960
      TabIndex        =   3
      Top             =   0
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4935
         Left            =   120
         ScaleHeight     =   329
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   497
         TabIndex        =   1
         Top             =   240
         Width           =   7455
         Begin MSComDlg.CommonDialog Dialog2 
            Left            =   3120
            Top             =   2760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "Pictures (*.bmp)|*.bmp"
         End
         Begin MSComDlg.CommonDialog Dialog1 
            Left            =   1800
            Top             =   2760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "Pictures (*.bmp;*.jpg;*.jpeg)|*.bmp;*.jpg;*.jpeg"
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Input the password here :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program uses three pixels for a character. It uses eight components of the
'three pixels(last remains unused) to form a binary representation of the ascii
'code of the character, but instead of using zeroes and ones, it uses odd and even
'numbers. To be more clear let me give you a practical example. Let's say that I
'need to store the character "a" in the following three pixels: (154, 73, 211),
'(98, 110,39) and (16,255,85). The code for "a" is 97(in decimal) and in binary is
'11000001 .  The program transforms the first eight components in even number so
'the three pixels became: (154,72,210),(98,110,38),(16,254,85). Then, it makes a
'corespondence between the components and the binary reprezentation of the character.
'If a digit in the binary number is 1, program adds 1 to the corespondent component.
'The final values for the components are: (155,73,210),(98,110,38),(16,255,85).
'You can notice that there is not much difference between the initial and the
'final values.
'A little bit about the password protection. Consider the ascii caracter set as a
'deck of cards. The program creates a number of such "decks" which is equal to the
'number of the caracter in the password. Then it shuffles each "deck" in a way
'which depends on the the entire password and on the corespondent character in
'the password. For example, if the character "b" is the fifth in the password,
'the shuffle will depend on 5 * Asc("b") * a number which depend on the entire password.
'The total number of arangements of a sequence of 256 elements is 1*2*3* ...*255*256 ,
'which is a number with more than 600 digits ! You notice also that the longer is the
'password, the stroger is the encryption.
'If you don't input a password when you put the message, you don't have to input a password
'when you get the message.

Option Explicit
Private Sub start()
   Picture1.Move 0, 0
   Picture2.Move 0, 0
   If Picture2.Height < Picture1.Height Then Picture2.Top = (Picture1.Height - Picture2.Height) / 2
   If Picture2.Width < Picture1.Width Then Picture2.Left = (Picture1.Width - Picture2.Width) / 2
   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width

   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Height

   VScroll1.Visible = (Picture1.Height < Picture2.Height)
   HScroll1.Visible = (Picture1.Width < Picture2.Width)
   HScroll1.Value = 0
   VScroll1.Value = 0
End Sub

Private Sub Command1_Click()
PutMessage
End Sub

Private Sub Command2_Click()
Dim d As String
GetMessage
d = " "
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Dialog1.ShowOpen
If Dialog1.FileName <> "" Then
Picture2.Picture = LoadPicture(Dialog1.FileName)
Text1.Text = ""
start
End If
End Sub

Private Sub Command5_Click()
Dialog2.ShowSave
If Dialog2.FileName <> "" Then SavePicture Picture2.Image, Dialog2.FileName
End Sub

Private Sub Form_Load()
Picture2.Picture = LoadPicture(App.Path & "/USflag.bmp")
start
End Sub

Private Sub VScroll1_Change()
   Picture2.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Picture2.Left = -HScroll1.Value
End Sub
Private Sub GetMessage()
Dim i As Long, j As Long, k As Long, n As Long, pix(0 To 2) As Long
Dim tx As String, nmd As Long, start As Integer
Dim endmess As String, comp(1 To 8) As Long, ch As Long
work 1
pass.Text = Trim(pass.Text)
Shuffle (pass.Text)
For i = 0 To Picture2.ScaleWidth - 1
For j = 0 To Picture2.ScaleHeight - 1
nmd = n Mod 3
If nmd = 0 Then
        If start < 14 Then
            start = start + 1
            If start = 14 And tx <> "start message" Then
            Text1.Text = "THIS PICTURE HAS NO SECRET MESSAGE"
            work 0
            Exit Sub
            ElseIf start = 14 Then
            tx = ""
            End If
        End If
ch = 0
pix(nmd) = Picture2.Point(i, j)
comp(8) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
comp(7) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)
comp(6) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2)
        For k = 8 To 6 Step -1
        ch = ch + (2 ^ (k - 1)) * comp(k)
        Next k
End If
If nmd = 1 Then
pix(nmd) = Picture2.Point(i, j)
comp(5) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
comp(4) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)
comp(3) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2)
        For k = 5 To 3 Step -1
        ch = ch + (2 ^ (k - 1)) * comp(k)
        Next k
End If
If nmd = 2 Then
pix(nmd) = Picture2.Point(i, j)
comp(2) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
comp(1) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)
        For k = 2 To 1 Step -1
        ch = ch + (2 ^ (k - 1)) * comp(k)
        Next k
End If
        
n = n + 1
    If n = 3 Then
        n = 0
        tx = tx & getc(Chr(ch))
    End If
        endmess = Right(tx, 11)
        If endmess = "end message" Then
            Text1.Text = Left(tx, Len(tx) - 11)
            work 0
            Exit Sub
        End If
Next j
Next i
End Sub
Private Sub PutMessage()
Dim i As Long, j As Long, tx As String, ch As String, NrPix As Long
Dim pix(0 To 2) As Long, wid As Long, hig As Long
Dim r As Long, g As Long, b As Long, comp(1 To 8) As Long
Dim aa(0 To 2) As Long, bb(0 To 2) As Long
work 1
tx = "start message" & Text1.Text & "end message"
wid = Picture2.ScaleWidth
hig = Picture2.ScaleHeight
If Len(tx) * 3 > hig * wid Then
tx = MsgBox("Text is " & Len(tx) * 3 - wid * hig & " characters longer than this picture can store", vbCritical)
work 0
Exit Sub
End If
pass.Text = Trim(pass.Text)
Shuffle (pass.Text)
For i = 1 To Len(tx)
    ch = ByteToBin(Asc(putc(Mid(tx, i, 1))))
    NrPix = (CLng(i) - 1) * 3
    aa(0) = (NrPix Mod hig)
    bb(0) = (NrPix \ hig)
    pix(0) = Picture2.Point(bb(0), aa(0)) 'the first pixel in the group of three.
    r = (pix(0) And RGB(255, 0, 0)) - (pix(0) And RGB(255, 0, 0)) Mod 2: comp(1) = r
    g = ((pix(0) And RGB(0, 255, 0)) \ 256) - ((pix(0) And RGB(0, 255, 0)) \ 256) Mod 2: comp(2) = g
    b = ((pix(0) And RGB(0, 0, 255)) \ 65536) - ((pix(0) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(3) = b
    
    NrPix = NrPix + 1
    aa(1) = (NrPix Mod hig)
    bb(1) = (NrPix \ hig)
    pix(1) = Picture2.Point(bb(1), aa(1)) 'the second pixel in the group of three.
    r = (pix(1) And RGB(255, 0, 0)) - (pix(1) And RGB(255, 0, 0)) Mod 2: comp(4) = r
    g = ((pix(1) And RGB(0, 255, 0)) \ 256) - ((pix(1) And RGB(0, 255, 0)) \ 256) Mod 2: comp(5) = g
    b = ((pix(1) And RGB(0, 0, 255)) \ 65536) - ((pix(1) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(6) = b
    
    NrPix = NrPix + 1
    aa(2) = (NrPix Mod hig)
    bb(2) = (NrPix \ hig)
    pix(2) = Picture2.Point(bb(2), aa(2)) 'the third pixel in the group of three.
    r = (pix(2) And RGB(255, 0, 0)) - (pix(2) And RGB(255, 0, 0)) Mod 2: comp(7) = r
    g = ((pix(2) And RGB(0, 255, 0)) \ 256) - ((pix(2) And RGB(0, 255, 0)) \ 256) Mod 2: comp(8) = g
    b = ((pix(2) And RGB(0, 0, 255)) \ 65536) 'last component remains unchanged
    
    For j = 1 To 8
    comp(j) = comp(j) + CInt(Mid(ch, j, 1)) * 1
    Next j
    Picture2.PSet (bb(0), aa(0)), RGB(comp(1), comp(2), comp(3))
    Picture2.PSet (bb(1), aa(1)), RGB(comp(4), comp(5), comp(6))
    Picture2.PSet (bb(2), aa(2)), RGB(comp(7), comp(8), b)
    
Next i
work 0
End Sub
Private Sub work(i As Integer)
If i = 1 Then Form1.Caption = "Secret Messenger (Working...)"
If i = 0 Then Form1.Caption = "Secret Messenger"
End Sub


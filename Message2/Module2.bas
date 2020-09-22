Attribute VB_Name = "Module"
Dim arrayA() As Integer, arrayB() As Integer, ln As Integer
Public Function ByteToBin(n As Integer) As String   'This function transforms an integer (which is the
Dim j As String                                     'the ascii code of a character) into a string (which
Do While n >= 1                                     'is the binary representation of the ascii code)
j = n Mod 2 & j
n = n \ 2
Loop
If Len(j) < 8 Then j = String(8 - Len(j), "0") & j
ByteToBin = j
End Function
Public Function putc(c As String) As String 'For each character in the message the program picks randomly
Dim ps As String                            'a "deck" of characters, depending on the character itself and
ps = Form1.pass.Text                        'and on the length of the password.
If ps <> "" Then
Randomize Asc(Mid(ps, 1 + Int(Len(ps) * Rnd), 1)) * (1 + Int(Len(ps) * Rnd)) * 13
putc = Chr(arrayA(Asc(c), 1 + Int(Len(ps) * Rnd)))
Else
putc = c
End If
End Function
Public Function getc(c As String) As String
Dim ps As String
ps = Form1.pass.Text
If ps <> "" Then
Randomize Asc(Mid(ps, 1 + Int(Len(ps) * Rnd), 1)) * (1 + Int(Len(ps) * Rnd)) * 13
getc = Chr(arrayB(Asc(c), 1 + Int(Len(ps) * Rnd)))
Else
getc = c
End If
End Function

Public Sub Shuffle(pas As String)
Dim i As Integer, j As Integer, k As Double, x As Integer, y As Integer, t As Integer
ln = Len(pas)
If ln > 0 Then
k = 1
For j = 1 To ln
k = k + Asc(Mid(pas, j, 1)) * j
Next j
k = Sqr(k)
ReDim arrayA(0 To 255, 1 To ln) As Integer
ReDim arrayB(0 To 255, 1 To ln) As Integer
For i = 1 To Len(pas)
    For j = 0 To 255
     arrayA(j, i) = j
    Next j
Next i
For j = 1 To ln
f = Rnd(-1)
Randomize Asc(Mid(pas, j, 1)) * CDbl(j) * k
    For i = 1 To 10000
        y = Int(255 * Rnd)
        t = 255 - Int(255 * Rnd)
        x = arrayA(y, j)
        arrayA(y, j) = arrayA(t, j)
        arrayA(t, j) = x
    Next i
Next j
For i = 1 To ln
For j = 0 To 255
arrayB(arrayA(j, i), i) = j
Next j
Next i
End If
End Sub

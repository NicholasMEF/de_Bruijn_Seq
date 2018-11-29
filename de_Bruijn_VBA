Function de_Bruijn(alphabet As Double, seqLen As Double) As String
Dim totalLen As Double, letterLen As Double, w As Variant, w1 As Variant, i As Double, alphaChar As Double, sortChar As Double, sortCharLen As Double
Dim tracker As Variant, poniter As Double, transMap As Variant, sequence As String

totalLen = alphabet ^ seqLen
letterLen = totalLen / alphabet
ReDim w(totalLen - 1)
ReDim w1(totalLen - 1)
ReDim tracker(totalLen - 1)
ReDim transMap(totalLen - 1)

i = 0
alphaChar = 0
sortChar = 0
sortCharLen = 0
Do
    w(i) = alphaChar
    If alphaChar < alphabet - 1 Then
        alphaChar = alphaChar + 1
    Else
        alphaChar = 0
    End If
    w1(i) = sortChar
    If sortCharLen < letterLen - 1 Then
        sortCharLen = sortCharLen + 1
    Else
        sortCharLen = 0
        sortChar = sortChar + 1
    End If
    tracker(i) = 0
    i = i + 1
Loop While i < totalLen

i = 0
pointer = 0
Do
    i1 = 0
    Do
        If w1(i) = w(i1) And tracker(i1) <> 1 Then
            pointer = i1
            Exit Do
        End If
        i1 = i1 + 1
    Loop While i1 < totalLen
    transMap(i) = pointer
    tracker(pointer) = 1
    i = i + 1
Loop While i < totalLen

i = 0
pointer = 0
Do
    pointer = transMap(pointer)
    If tracker(pointer) = 1 Then
        sequence = sequence & w1(pointer)
        If i < totalLen - 1 Then
            sequence = sequence & ", "
        End If
    Else
        Do
            pointer = pointer + 1
            If tracker(pointer) = 1 Then
                Exit Do
            End If
        Loop While pointer < totalLen
        sequence = sequence & w1(pointer)
        If i < totalLen - 1 Then
            sequence = sequence & ", "
        End If
    End If
    tracker(pointer) = 2
    i = i + 1
Loop While i < totalLen

de_Bruijn = sequence

End Function



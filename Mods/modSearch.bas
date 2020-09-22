Attribute VB_Name = "modSearch"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sign(4096)     As String    'The Signatures will be loaded into this array

Public Sub BuildSigns()

  'This builds the Signature - Array
  
  Dim sIn      As String
  Dim swords() As String
  Dim X        As Long

    On Error GoTo err
    sIn = FileText(AV.Signature.SignatureFilename)
    swords = Split(sIn, vbCrLf)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        Sign(X) = swords(X)
    Next X
    AV.Signature.SignatureDate = Sign(0)
    AV.Signature.SignatureCount = UBound(swords) - 1

Exit Sub

err:
    MsgBox "An error has occured while loading the signature File!" & vbCrLf & "This could be caused by an empty or damaged file!" & vbCrLf & vbCrLf & "The error message was: " & err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Function FindTerm(ByVal File As String, _
                          ByVal s As String, _
                          ZZ() As String, _
                          ByVal tl As String, _
                          ByVal tr As String) As Boolean

'this scans the given file for the given signature
  Dim c    As Long
  Dim F    As Long
  Dim i    As Long
  Dim j    As Long
  Dim L    As Long
  Dim lc   As Long
  Dim p    As Long
  Dim v    As Long
  Dim w    As Long
  Dim a    As String
  Dim d    As String
  Dim n    As String
  Dim o    As String
  Const PS As Long = 1024&

    ReDim ZZ(0)
    If LenB(tl) = 0 Or LenB(tr) = 0 Or LenB(s) = 0 Or LenB(Dir(File, vbNormal)) = 0 Then
        Exit Function
    End If
    F = FreeFile
    Open File For Binary Shared As #F
    L = LOF(F)
    p = L \ PS
    If L Mod PS <> 0 Then
        p = p + 1
    End If
    For c = 1 To p
        n = Space$(PS)
        Get F, , n
        a = o & n
        i = InStr(1, a, s)
        If i <> 0 Then
            lc = 0
            Do
                i = InStr(i, a, s)
                If i <> 0 Then
                    v = 1
                    For j = i To 1 Step -1
                        d = Mid$(a, j, 1)
                        If InStr(1, tl, d) Then
                            v = j + 1
                            Exit For
                        End If
                    Next j
                    w = 0
                    For j = i To Len(a)
                        d = Mid$(a, j, 1)
                        If InStr(1, tr, d) Then
                            w = j - 1
                            Exit For
                        End If
                    Next j
                    If w <> 0 Then
                        ZZ(UBound(ZZ)) = Mid$(a, v, w - v + 1)
                        ReDim Preserve ZZ(0 To UBound(ZZ) + 1)
                        lc = w
                    End If
                    i = w
                End If
            Loop Until i = 0
            If lc = 0 Then
                o = a
             Else 'NOT LC...
                o = Mid$(a, lc)
            End If
         Else 'NOT I...
            o = n
        End If
    Next c
    Close F
    If UBound(ZZ) > 0 Then
        FindTerm = True
    End If
    DoEvents

End Function

Public Function Search(ByVal strFilename As String) As String

  Dim Zeilen() As String
  Dim FName    As String
  Dim Current  As Long

    FName = strFilename
    For Current = 1 To 4096
        If Sign(Current) = "#END#" Then
            GoTo Finish
        End If
        If FindTerm(FName, Sign(Current), Zeilen, vbLf, vbCr) Then
            DoEvents
            Search = Sign(Current)
            Exit Function
         Else 'NOT FINDTERM(FNAME,...
            Search = "NOTHING"
        End If
    Next Current
Finish:

End Function



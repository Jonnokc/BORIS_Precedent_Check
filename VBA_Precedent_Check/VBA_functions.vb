
Function HotFuzz(ByVal S1 As String, ByVal S2 As String, Optional ByVal N As Boolean = True, Optional ByVal x As String = "", Optional ByVal w As Single = 2) As Single
  'Using Like operator for filtering, with added code to allow special characters in the input strings, including hyphen and the right bracket - passed in the 'x' parameter.
  'Use x & Chr(34) if you need to allow double quotes (") in the input strings
  'Allowing numbers in the input strings is optional (the 'n' parameter)
  'The 'w' parameter is the weight of "order" over "frequency" scores in the final score. Feel free to experiment, to get best matching results with your data.
  Dim i       As Integer
  Dim d1      As Integer
  Dim d2      As Integer
  Dim y       As String
  Dim b       As Boolean
  Dim c       As String
  Dim a1      As String
  Dim a2      As String
  Dim k       As Integer
  Dim p       As Integer
  Dim f       As Single
  Dim o       As Single
  '
  '        ******* INPUT STRINGS CLEANSING *******
  '
  HotFuzz = 0
  b = False
  If N Then                                      'allow numbers in the input strings?
    y = "[A-Z0-9"
  Else
    y = "[A-Z"
  End If
  If Len(x) > 0 Then                             'we want to allow some special characters in the input strings, i.e. space, punctuation etc
    If InStr(1, x, "-", 0) Then
      y = Replace(x, "-", "") & "-"              'hyphen must be placed first or last inside a [..] group in a Like comparison
    End If
    If InStr(1, x, "]", 0) Then
      y = Replace(x, "]", "")                    'right bracket can't be part of a [..] group in a Like comparison - dedicated logic must be developed to treat this case
      b = True                                   'if we want to allow the right bracket in the input strings
    End If
  End If
  y = y & "]"                                    'closing the group
  S1 = UCase$(S1)                                'input strings are converted to uppercase
  d1 = Len(S1)
  a1 = ""
  For i = 1 To d1
    c = Mid$(S1, i, 1)
    If c Like y Then                             'filter the allowable characters
      a1 = a1 & c                                'a1 is what remains from s1 after filtering
    ElseIf b Then
      If c = "]" Then                            'special treatment for the right bracket
        a1 = a1 & c
      End If
    End If
  Next
  d1 = Len(a1)
  If d1 = 0 Then Exit Function
  S2 = UCase$(S2)
  d2 = Len(S2)
  a2 = ""
  For i = 1 To d2
    c = Mid$(S2, i, 1)
    If c Like y Then
      a2 = a2 & c
    End If
  Next
  d2 = Len(a2)
  If d2 = 0 Then Exit Function
  k = d1
  If d2 < d1 Then                                'to prevent doubling the code below s1 must be made the shortest string,
    'so we swap the variables
    k = d2
    d2 = d1
    d1 = k
    S1 = a2
    S2 = a1
    a1 = S1
    a2 = S2
  Else
    S1 = a1
    S2 = a2
  End If
  If k = 1 Then                                  'degenerate case, where the shortest string is just one character
    If InStr(1, S2, S1, 0) Then
      HotFuzz = 1 / d2
    Else
      HotFuzz = 0
    End If
  Else                                           '******* MAIN LOGIC HERE *******
    i = 1
    f = 0
    o = 0
    Do                                           'count the identical characters in s1 and s2 ("frequency analysis")
      p = InStr(1, S2, Mid$(S1, i, 1), 0)
      'search the character at position i from s1 in s2
      If p > 0 Then                              'found a matching character, at position p in s2
        f = f + 1                                'increment the frequency counter
        Mid$(S2, p, 1) = "~"
        'replace the found character with one outside the allowable list
        '(I used tilde here), to prevent re-finding
        Do                                       'check the order of characters
          If i >= k Then Exit Do                 'no more characters to search
          If Mid$(S2, p + 1, 1) = Mid$(S1, i + 1, 1) Then
            'test if the next character is the same in the two strings
            f = f + 1                            'increment the frequency counter
            o = o + 1                            'increment the order counter
            i = i + 1
            p = p + 1
          Else
            Exit Do
          End If
        Loop
      End If
      If i >= k Then Exit Do
      i = i + 1
    Loop
    If o > 0 Then o = o + 1                      'if we got at least one match, adjust the order counter because two characters are required to define "order"
    HotFuzz = (w * o + f) / (w + 1) / d2
  End If
End Function

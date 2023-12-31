Attribute VB_Name = "jsonParser"
Option Explicit

' https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a

' SEE USAGE AT THE BOTTOM

'-------------------------------------------------------------------
' VBA JSON Parser
'-------------------------------------------------------------------

Private p&, token, dic
Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function
Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function
Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function
'-------------------------------------------------------------------
' Support Functions
'-------------------------------------------------------------------
Function Tokenize(s$)
    Const pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, pattern, True)
End Function
Function RExtract(s$, pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .pattern = pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function
Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function
Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function
Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function
Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function

Private Sub UsageExample()
    Dim sample As String: sample = _
"{ " & _
"   ""data"" : { " & _
"     ""receipt_time"" : ""2018-09-28T10:00:00.000Z"", " & _
"     ""site"" : ""Los Angeles"", " & _
"     ""measures"" : [ { " & _
"        ""test_id"" : ""C23_PV"", " & _
"        ""metrics"" : [ { " & _
"            ""val1"" : [ 0.76, 0.75, 0.71 ], " & _
"            ""temp"" : [ 0, 2, 5 ], " & _
"            ""TS"" : [ 1538128801336, 1538128810408, 1538128818420 ] " & _
"          } ] " & _
"       }, " & _
"    { " & _
"            ""test_id"" : ""HBI2_XX"", " & _
"            ""metrics"" : [ { " & _
"            ""val1"" : [ 0.65, 0.71 ], " & _
"            ""temp"" : [ 1, -7], " & _
"            ""TS"" : [ 1538128828433, 1538128834541 ] " & _
"            } ] " & _
"       }] " & _
"    } " & _
"  } "

    Dim dic As Object: Set dic = ParseJSON(sample) ' PARSE A STRING
    Debug.Print ListPaths(dic) ' SHOW AVAILABLE PATHS
    MsgBox dic("obj.data.measures(0).metrics(0).temp(2)") ' ACCESS A PATH
    Dim v As Variant: v = GetFilteredValues(dic, "*.metrics*") ' FILTER ITEMS TO AN ARRAY
    Dim i As Variant
    For Each i In v
        Debug.Print i
    Next
    
End Sub

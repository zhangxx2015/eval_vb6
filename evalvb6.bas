'standard project
'all in module
Option Explicit

Private Function Eval(ByVal s As String) As Variant
  Dim Lb As Long, Rb As Long
  Lb = InStrRev(s, "(")
  While Lb <> 0
  Rb = InStr(Lb, s, ")")
  s = Replace(s, Mid$(s, Lb, Rb - Lb + 1), CStr(Eval(Mid$(s, Lb + 1, Rb - Lb - 1))))
  Lb = InStrRev(s, "(")
  Wend
  If IsNumeric(s) Then
  Eval = Val(s)
  Else
  Dim High As Long
  High = InStr(s, "+")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) + Eval(Right$(s, Len(s) - High))
  Exit Function
  End If
  High = InStrRev(s, "-")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) - Eval(Right$(s, Len(s) - High))
  Exit Function
  End If
  High = InStr(s, "*")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) * Eval(Right$(s, Len(s) - High))
  Exit Function
  End If
  High = InStrRev(s, "/")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) / Eval(Right$(s, Len(s) - High))
  Exit Function
  End If
  High = InStrRev(s, "%")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) Mod Eval(Right$(s, Len(s) - High))
  Exit Function
  End If
  High = InStrRev(s, "!=")
  If High <> 0 Then
  Eval = Eval(Left$(s, High - 1)) <> Eval(Right$(s, Len(s) - High - 1))
  Exit Function
  End If
  High = InStrRev(s, "==")
  If High <> 0 Then
  Eval = (Eval(Left$(s, High - 1)) = Eval(Right$(s, Len(s) - High - 1)))
  Exit Function
  End If
  End If
End Function

'主函数
Public Sub main()
  Dim Expression As String
  Expression = "149.5+((100+(6+(90-5*2*2)*4+(1-1))+202)%441)*2*2+0.88+150.5"
  MsgBox Eval(Expression)
End Sub
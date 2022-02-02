Private Sub Worksheet_Change(ByVal Target As Range)

Dim Oldvalue As String
Dim Newvalue As String
Application.EnableEvents = True
On Error GoTo Exitsub

If Target.SpecialCells(xlCellTypeAllValidation) Is Nothing Then
    GoTo Exitsub
Else:
    If Target.Value = "" Then
        GoTo Exitsub
    Else:
        Application.EnableEvents = False
        Newvalue = Target.Value
        Application.Undo
        Oldvalue = Target.Value
        If Oldvalue = "" Then
            Target.Value = Newvalue
        Else
            If InStr(1, Oldvalue, Newvalue) = 0 Then
                Target.Value = Oldvalue & ", " & Newvalue
            Else:
                Dim Position As Integer
                Position = InStr(1, Oldvalue, Newvalue)
                
                If Position > 1 Then
                    Newvalue = Left(Oldvalue, Position - 3)
                    Position = InStr(Position, Oldvalue, ",")
                
                    If Position > 0 Then
                        Newvalue = Newvalue & Right(Oldvalue, Len(Oldvalue) - Position + 1)
                    End If
                    
                Else:
                    Position = InStr(Position, Oldvalue, ",")
                    Newvalue = Right(Oldvalue, Len(Oldvalue) - Position)
                End If
                
                

                Target.Value = Trim(Newvalue)
            End If
        End If
    End If
End If

Application.EnableEvents = True
Exitsub:
Application.EnableEvents = True
End Sub

Option Explicit

' Clear buttons. Each one clears the part inputs of one tab only.
' Plant settings (rates, tooling costs, press list, placeholders) are kept.

Public Sub ClearBusbarStamping()
    If MsgBox("Clear the part inputs on Busbar Stamping (part size, terminals, tabs, bends, pierces and layout / press overrides)?" & vbCrLf & _
              "Material source, annual volume, rates and tooling costs are kept.", vbYesNo + vbQuestion, "Clear Busbar Stamping") <> vbYes Then Exit Sub
    With ThisWorkbook.Worksheets("Busbar Stamping")
        .Range("B7").ClearContents
        .Range("B23:B24").ClearContents
        .Range("B26").ClearContents
        .Range("B28:B35").Value = 0
        .Range("B46:B47").ClearContents
        .Range("B63:B65").Value = 0
        .Range("I61:I62").ClearContents
        .Activate
        .Range("B23").Select
    End With
End Sub

Public Sub ClearBOMPaste()
    Dim d As Variant, i As Long
    If MsgBox("Clear the part numbers on BOM Paste and reset the descriptions?", vbYesNo + vbQuestion, "Clear BOM Paste") <> vbYes Then Exit Sub
    d = Array("Stamped PCB busbar", "Fuse block plastic base", "Fuse block cover", "Fuse block busbar (copper)", _
              "Fuse bolts and nuts (purchased)", "B+ connection part (purchased)", "Keepout zone stud hardware (purchased)", _
              "Assembly process adders")
    With ThisWorkbook.Worksheets("BOM Paste")
        .Range("A5:A12").ClearContents
        For i = 0 To 7
            .Range("B" & (5 + i)).Value = d(i)
        Next i
        .Activate
        .Range("A5").Select
    End With
End Sub

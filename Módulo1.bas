Attribute VB_Name = "Módulo1"
Sub filtro_chave()
Dim LINHA As Double, LinhaList As Double

Planilha1.Activate
Planilha1.Range("A1:E5000").Select
UserForm1.ListBox1.Clear
    With UserForm1.ListBox1
    .AddItem
    .List(0, 0) = "ID"
    .List(0, 1) = "PALAVRA CHAVE"
    .List(0, 2) = "OBSERVAÇÃO"
    .List(0, 3) = "CÓDIGO"
    .List(0, 4) = "LINGUAGEM"
     .ColumnWidths = "30;250;500;0;90"
            'teste para git-hub
    End With
    
UserForm1.ListBox2.ColumnHeads = True
UserForm1.ListBox2.RowSource = "A2:E2"
UserForm1.ListBox2.ColumnWidths = "30;250;500;0;90"
 
 
 
LINHA = 1
LinhaList = 0


If UserForm1.TextBox1.Value = Empty Then
Selection.AutoFilter Field:=1
Selection.AutoFilter Field:=2
Selection.AutoFilter Field:=3
Selection.AutoFilter Field:=4
Selection.AutoFilter Field:=5

With Planilha1

Do
LINHA = LINHA + 1

If .Cells(LINHA, 1).Value <> Empty Then
    With UserForm1.ListBox1
    .AddItem
    .List(LinhaList, 0) = Planilha1.Cells(LINHA, 1).Value
    .List(LinhaList, 1) = Planilha1.Cells(LINHA, 2).Value
    .List(LinhaList, 2) = Planilha1.Cells(LINHA, 3).Value
    .List(LinhaList, 3) = Planilha1.Cells(LINHA, 4).Value
    .List(LinhaList, 4) = Planilha1.Cells(LINHA, 5).Value
        LinhaList = LinhaList + 1
    End With

End If
Loop Until .Cells(LINHA, 1).Value = Empty

End With


Exit Sub


End If
Selection.AutoFilter Field:=2, Criteria1:=CStr("*" + UserForm1.TextBox1.Text) + "*"
'Selection.AutoFilter Field:=5, Criteria1:=CStr("*" + TextBox3.Text) + "*"


With Planilha1
Do
LINHA = LINHA + 1

If .Rows(LINHA).EntireRow.Hidden = False And .Cells(LINHA, 1).Value <> Empty Then
 With UserForm1.ListBox1
    .AddItem
    .List(LinhaList, 0) = Planilha1.Cells(LINHA, 1).Value
    .List(LinhaList, 1) = Planilha1.Cells(LINHA, 2).Value
    .List(LinhaList, 2) = Planilha1.Cells(LINHA, 3).Value
    .List(LinhaList, 3) = Planilha1.Cells(LINHA, 4).Value
     .List(LinhaList, 4) = Planilha1.Cells(LINHA, 5).Value
        LinhaList = LinhaList + 1
End With
End If
    Loop Until .Cells(LINHA, 1).Value = Empty
End With



Exit Sub
Exit Sub

End Sub


Sub filtro_ling()
Dim LINHA As Double, LinhaList As Double

Planilha1.Activate
Planilha1.Range("A1:E5000").Select
UserForm1.ListBox1.Clear
    With UserForm1.ListBox1
    .AddItem
    .List(0, 0) = "ID"
    .List(0, 1) = "PALAVRA CHAVE"
    .List(0, 2) = "OBSERVAÇÃO"
    .List(0, 3) = "CÓDIGO"
    .List(0, 4) = "LINGUAGEM"
     .ColumnWidths = "30;250;500;0;90"

    End With
UserForm1.ListBox2.ColumnHeads = True
UserForm1.ListBox2.RowSource = "A2:E2"
UserForm1.ListBox2.ColumnWidths = "30;250;500;0;90"
 
 
 
LINHA = 1
LinhaList = 0


If UserForm1.TextBox3.Value = Empty Then
Selection.AutoFilter Field:=1
Selection.AutoFilter Field:=2
Selection.AutoFilter Field:=3
Selection.AutoFilter Field:=4
Selection.AutoFilter Field:=5

With Planilha1

Do
LINHA = LINHA + 1

If .Cells(LINHA, 1).Value <> Empty Then
    With UserForm1.ListBox1
    .AddItem
    .List(LinhaList, 0) = Planilha1.Cells(LINHA, 1).Value
    .List(LinhaList, 1) = Planilha1.Cells(LINHA, 2).Value
    .List(LinhaList, 2) = Planilha1.Cells(LINHA, 3).Value
    .List(LinhaList, 3) = Planilha1.Cells(LINHA, 4).Value
    .List(LinhaList, 4) = Planilha1.Cells(LINHA, 5).Value
        LinhaList = LinhaList + 1
    End With

End If
Loop Until .Cells(LINHA, 1).Value = Empty

End With


Exit Sub
                                ' sdg  testesda
End If
Selection.AutoFilter Field:=2, Criteria1:=CStr("*" + UserForm1.TextBox1.Text) + "*"
Selection.AutoFilter Field:=5, Criteria1:=CStr("*" + UserForm1.TextBox3.Text) + "*"


With Planilha1
Do
LINHA = LINHA + 1

If .Rows(LINHA).EntireRow.Hidden = False And .Cells(LINHA, 1).Value <> Empty Then
 With UserForm1.ListBox1
    .AddItem
    .List(LinhaList, 0) = Planilha1.Cells(LINHA, 1).Value
    .List(LinhaList, 1) = Planilha1.Cells(LINHA, 2).Value
    .List(LinhaList, 2) = Planilha1.Cells(LINHA, 3).Value
    .List(LinhaList, 3) = Planilha1.Cells(LINHA, 4).Value
     .List(LinhaList, 4) = Planilha1.Cells(LINHA, 5).Value
        LinhaList = LinhaList + 1
End With
End If
    Loop Until .Cells(LINHA, 1).Value = Empty
End With
Exit Sub
Exit Sub


End Sub

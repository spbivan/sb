# sb
 test_task

 Ниже реализовано 2 подхода

 #v1
 Просматриваем последовательно все ячейки матрицы.
 Если значение ячейки = 1, то выписываем адрес ячейки.

 #v2
 Сначала суммируем значение 2х соседних ячеек в матрице.
 Если сумма = 2, то выписываем адреса ячеек.
 Если сумма = 1, то просматриваем одну ячейку.
 Алгоритм действителен для четного m.

 #Резюме:

 v2 предпочтителенее, тк на больших значениях m и n можно выиграть в производительности алгоритма.
 Плюс алгоритм эффективнее, если распределение 0 и 1 будет не 50% на 50%, а например 25% на 75%


 Sub check_v1()

 Dim m, n, r, counter As Integer
 Dim list, ocell As String


 m = 20 ' rows i
 n = 20 ' columns j
 counter = 0
    
 ThisWorkbook.ActiveSheet.Cells(21, 1) = ""
 ThisWorkbook.ActiveSheet.Cells(22, 1) = ""
    
 list = ""
 list = "["
 For c = 1 To n
 For r = 1 To m
    ocell = ""
    If ThisWorkbook.ActiveSheet.Cells(r, c) = 1 Then
    ocell = "(" & "I" & CStr(r) & ";" & "J" & CStr(c) & "), "
    counter = counter + 1
 End If
 list = list + ocell
 Next r
 Next c
 list = list + "]"
 ThisWorkbook.ActiveSheet.Cells(21, 1) = list
 ThisWorkbook.ActiveSheet.Cells(22, 1) = counter
    
    MsgBox list
 End Sub

=====================================================================================================================================

 Sub check_v2()

 Dim m, n, r, v1, v2, counter As Integer
 Dim list, ocell As String


 m = 20 ' rows i
 n = 20 ' columns j
 counter = 0

 ThisWorkbook.ActiveSheet.Cells(21, 1) = ""
 ThisWorkbook.ActiveSheet.Cells(22, 1) = ""

    
 list = ""
 list = "["
 For c = 1 To n
 For r = 1 To m Step 2
    ocell = ""
    v = ThisWorkbook.ActiveSheet.Cells(r, c) + ThisWorkbook.ActiveSheet.Cells(r + 1, c)
    If v = 2 Then
    ocell = "(" & "I" & CStr(r) & ";" & "J" & CStr(c) & "), "
    ocell = "(" & "I" & CStr(r + 1) & ";" & "J" & CStr(c) & "), " + ocell
    counter = counter + 2
    ElseIf v = 1 Then
         If ThisWorkbook.ActiveSheet.Cells(r, c) = 1 Then
         ocell = "(" & "I" & CStr(r) & ";" & "J" & CStr(c) & "), "
         Else
         ocell = "(" & "I" & CStr(r + 1) & ";" & "J" & CStr(c) & "), "
         End If
     counter = counter + 1
    End If
 list = list + ocell
 Next r
 Next c
 list = list + "]"
 ThisWorkbook.ActiveSheet.Cells(21, 1) = list
 ThisWorkbook.ActiveSheet.Cells(22, 1) = counter
    
    MsgBox list
 End Sub

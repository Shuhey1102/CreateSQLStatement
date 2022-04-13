Attribute VB_Name = "Module1"
Const StartIndex As Integer = 10 'Start Index

Sub ExecuteProcess()

    'Initial Process
    Worksheets("Sheet2").Cells.Clear
    
    Dim TableName As String
    TableName = Sheets("Sheet1").Cells(3, 2).Value 'Get TableName

    Dim MandatoryCheckExeFLG As Boolean
    If (Sheets("Sheet1").Cells(2, 11).Value = "") Then
        MandatoryCheckExeFLG = False
    Else
        MandatoryCheckExeFLG = Sheets("Sheet1").Cells(2, 11).Value 'Get MadndatoryCheckFlg
    End If
    
    Dim Columnlength As Integer 'ColumnNumber
    Columnlength = GetColumnNumber 'Get ColumnNumber
    
    'Validation Check
    Dim StrInsert As String      'CREATE Statment
    Dim StrUpdate As String      'UPDATE Statment
    Dim StrDelete As String      'DELETE Statment
    
    'Create Template SQL Statement
    StrInsert = "INSERT INTO DBO." + TableName + "("
    StrUpdate = "UPDATE DBO." + TableName + " SET "
    StrDelete = "DELETE FROM DBO." + TableName + " WHERE ID = "
    Dim k As Integer
    For k = 1 To Columnlength
        If (k <> Columnlength) Then
            StrInsert = StrInsert + Sheets("Sheet1").Cells(6, k + 2).Value + ","
        Else
            StrInsert = StrInsert + Sheets("Sheet1").Cells(6, k + 2).Value + ") VALUES"
        End If
    Next
    Dim ExecuteFlg As Boolean 'EXE Create SQL FLG
    ExecuteFlg = True
    
    Dim i As Integer
    i = StartIndex
    
    'Target ID List
    Dim ADDNo As Variant
    Dim UPDNo As Variant
    Dim DELNo As Variant
    Dim tmp As Integer
    Dim ADDcnt As Integer
    Dim UPDcnt As Integer
    Dim DELcnt As Integer
    ADDcnt = 0
    UPDcnt = 0
    DELcnt = 0
    
    'Blank Line Count(if Greater than 5 blank lines, the record is End)
    Dim Blankcnt As Integer
    Blankcnt = 0
    
    'Validation Check & Append ID to each List
    Do While True
        If Sheets("Sheet1").Cells(i, 2) = "" Then
            If Sheets("Sheet1").Cells(i, 3) = "" Then
                Blankcnt = Blankcnt + 1
                If (Blankcnt > 5) Then    'Greater than 5 blank lines
                    Exit Do
                End If
            Else
                ExecuteFlg = False
                MsgBox "CMD for SQL(ADD,UPD,DEL) must be set (col:" & i & ")"
            End If
        Else
            If Sheets("Sheet1").Cells(i, 3) = "" Then
                 ExecuteFlg = False
                MsgBox "ID must be set (col:" & i & ")"
            End If
            Blankcnt = 0
            If Sheets("Sheet1").Cells(i, 2) = "ADD" Then
                If VarType(ADDNo) = 0 Then
                    tmp = 0
                    ReDim ADDNo(tmp)
                Else
                    tmp = UBound(ADDNo) + 1
                    ReDim Preserve ADDNo(tmp)
                End If
                ADDNo(tmp) = i
                ADDcnt = tmp + 1
            Else
                    If Sheets("Sheet1").Cells(i, 2) = "UPD" Then
                        If VarType(UPDNo) = 0 Then
                            tmp = 0
                            ReDim UPDNo(tmp)
                        Else
                            tmp = UBound(UPDNo) + 1
                            ReDim Preserve UPDNo(tmp)
                        End If
                            
                        UPDNo(tmp) = i
                        UPDcnt = tmp + 1
                    Else
                            If Sheets("Sheet1").Cells(i, 2) = "DEL" Then
                                If VarType(DELNo) = 0 Then
                                    tmp = 0
                                    ReDim DELNo(tmp)
                                Else
                                    tmp = UBound(DELNo) + 1
                                    ReDim Preserve DELNo(tmp)
                                End If
                
                                DELNo(tmp) = i
                                DELcnt = tmp + 1
                            Else
                                 MsgBox "CMD for SQL(ADD,UPD,DEL) not valid (col:" & i & ")"
                                 Exit Sub
                            End If
                    End If
            End If
        End If
        i = i + 1
    Loop
    
    If ExecuteFlg = False Then
        Exit Sub
    End If
    
    'Create SQL
    Dim OutPutCol As Integer
    OutPutCol = 2
    
    OutPutCol = CreateInsert(ADDNo, Columnlength, OutPutCol, StrInsert, MandatoryCheckExeFLG)
    If (OutPutCol <> -1) Then
        OutPutCol = CreateUpdate(UPDNo, Columnlength, OutPutCol, StrUpdate, MandatoryCheckExeFLG)
    End If
    If (OutPutCol <> -1) Then
        OutPutCol = CreateDelete(DELNo, Columnlength, OutPutCol, StrDelete, MandatoryCheckExeFLG)
    End If
    If (OutPutCol <> -1) Then
        MsgBox "Process is Success" & vbCrLf & "Insert : " & ADDcnt & vbCrLf & "Update : " & UPDcnt & vbCrLf & "Delete : " & DELcnt
    End If
End Sub
'Create Insert Statement
Function CreateInsert(ADDNo As Variant, Columnlength As Integer, OutPutCol As Integer, StrInsert As String, MandatoryCheckExeFLG As Boolean) As Integer
    'Initial Process
    If VarType(ADDNo) = 0 Then
        CreateInsert = OutPutCol
        Exit Function
    Else
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = "Insert"
        OutPutCol = OutPutCol + 1
        
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = StrInsert
        OutPutCol = OutPutCol + 1
    End If
    'Create Record
    For i = 0 To UBound(ADDNo)
        Dim Record As String
        Record = "("

        For j = 0 To Columnlength
            Dim columnType As String
            columnType = Sheets("Sheet1").Cells(7, j + 3).Value
            'case : VARCHAR,NVARCHAR,CHAR,DATETIME,DATE,TIME
            If columnType = "VARCHAR" Or columnType = "NVARCHAR" Or columnType = "CHAR" Or columnType = "DATETIME" Or columnType = "DATE" Or columnType = "TIME" Then
                If Sheets("Sheet1").Cells(ADDNo(i), j + 3).Value = "" Then
                    If Sheets("Sheet1").Cells(9, j + 3).Value = "NOT NULL" Then
                         If MandatoryCheckExeFLG Then
                            MsgBox (Sheets("Sheet1").Cells(6, j + 3).Value & " must be set. (col : " & ADDNo(i) & ")")
                            CreateInsert = -1
                            Exit Function
                         End If
                         Record = Record & "''"
                    Else
                        Record = Record & "NULL"
                    End If
                Else
                   Record = Record & "'" & Sheets("Sheet1").Cells(ADDNo(i), j + 3).Value & "'"
                End If
            'case : Except for VARCHAR,NVARCHAR,CHAR,DATETIME,DATE,TIME
            Else
                If Sheets("Sheet1").Cells(ADDNo(i), j + 3).Value = "" Then
                    If Sheets("Sheet1").Cells(9, j + 3).Value = "NOT NULL" Then
                        If MandatoryCheckExeFLG Then
                            MsgBox (Sheets("Sheet1").Cells(6, j + 3).Value & " must be set. (col : " & ADDNo(i) & ")")
                            CreateInsert = -1
                            Exit Function
                        Else
                             Record = Record & 0
                         End If
                    Else
                        Record = Record & 0
                    End If
                Else
                    Record = Record & Sheets("Sheet1").Cells(ADDNo(i), j + 3).Value
                End If
            End If
            If (j = Columnlength) Then
                Record = Record & ")"
            Else
                Record = Record & ","
            End If
        Next
        If (i <> UBound(ADDNo)) Then
            Record = Record & ","
        End If
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = Record
        OutPutCol = OutPutCol + 1
    Next
    'End Process
    OutPutCol = OutPutCol + 1
    CreateInsert = OutPutCol
End Function
'Create Update Statement
Function CreateUpdate(UPDNo As Variant, Columnlength As Integer, OutPutCol As Integer, StrUpdate As String, MandatoryCheckExeFLG As Boolean) As Integer
    'Initial Process
    If VarType(UPDNo) = 0 Then
        CreateUpdate = OutPutCol
        Exit Function
    Else
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = "Update"
        OutPutCol = OutPutCol + 1
        
    End If
    'Create Record
    For i = 0 To UBound(UPDNo)
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = StrUpdate
        OutPutCol = OutPutCol + 1
    
        Record = ""
        For j = 1 To Columnlength
            Dim columnType As String
            columnType = Sheets("Sheet1").Cells(7, j + 3).Value
            'case : VARCHAR,NVARCHAR,CHAR,DATETIME,DATE,TIME
            If columnType = "VARCHAR" Or columnType = "NVARCHAR" Or columnType = "CHAR" Or columnType = "DATETIME" Or columnType = "DATE" Or columnType = "TIME" Then
                If Sheets("Sheet1").Cells(UPDNo(i), j + 3).Value = "" Then
                    If Sheets("Sheet1").Cells(9, j + 3).Value = "NOT NULL" Then
                         If MandatoryCheckExeFLG Then
                            MsgBox (Sheets("Sheet1").Cells(6, j + 3).Value & " must be set. (col : " & UPDNo(i) & ")")
                            CreateUpdate = -1
                            Exit Function
                         End If
                         Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = ''"
                    Else
                        Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = NULL"
                    End If
                Else
                   Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = " & "'" & Sheets("Sheet1").Cells(UPDNo(i), j + 3).Value & "'"
                End If
            'case : Except for VARCHAR,NVARCHAR,CHAR,DATETIME,DATE,TIME
            Else
                If Sheets("Sheet1").Cells(UPDNo(i), j + 3).Value = "" Then
                    If Sheets("Sheet1").Cells(9, j + 3).Value = "NOT NULL" Then
                        If MandatoryCheckExeFLG Then
                            MsgBox (Sheets("Sheet1").Cells(6, j + 3).Value & " must be set. (col : " & UPDNo(i) & ")")
                            CreateUpdate = -1
                            Exit Function
                        Else
                             Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = " & 0
                         End If
                    Else
                        Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = " & 0
                    End If
                Else
                    Record = Record & Sheets("Sheet1").Cells(6, j + 3).Value & " = " & Sheets("Sheet1").Cells(UPDNo(i), j + 3).Value
                End If
            End If
            If (j = Columnlength) Then
                Record = Record & " WHERE ID = " & Sheets("Sheet1").Cells(UPDNo(i), 3).Value
            Else
                Record = Record & ","
            End If
        Next
        Record = Record & ";"
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = Record
        OutPutCol = OutPutCol + 1
    Next
    'End Process
    OutPutCol = OutPutCol + 1
    CreateUpdate = OutPutCol
End Function
'Create Delete Statement
Function CreateDelete(DELNo As Variant, Columnlength As Integer, OutPutCol As Integer, StrDelete As String, MandatoryCheckExeFLG As Boolean) As Integer
    'Initial Process
    If VarType(DELNo) = 0 Then
        CreateDelete = OutPutCol
        Exit Function
    Else
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = "Delete"
        OutPutCol = OutPutCol + 1
    End If
    'Create Record
    For i = 0 To UBound(DELNo)
        Sheets("Sheet2").Cells(OutPutCol, 2).Value = StrDelete & Sheets("Sheet1").Cells(DELNo(i), 3).Value & ";"
        OutPutCol = OutPutCol + 1
    Next
    'End Process
    OutPutCol = OutPutCol + 1
    CreateDelete = OutPutCol
End Function
'Check Column Number
Function GetColumnNumber()
    
    Dim j As Integer
    j = 3
    
    Do While True
        If Sheets("Sheet1").Cells(5, j) = "" Then
            Exit Do
        End If
        j = j + 1
    Loop
    
    GetColumnNumber = j - 3
End Function

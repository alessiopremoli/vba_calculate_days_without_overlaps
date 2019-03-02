Function calculateDaysWithoutOverlap(datesRange As Range) As Integer
    calculateDaysWithoutOverlap = 0
    Dim dat As Variant
    dat = datesRange
    
    Dim datesToCalculate As New Collection
    Dim inputData As New Collection
    
    'move input data to a collection:
    For Row = LBound(dat, 1) To UBound(dat, 1)
        inputData.Add Array(dat(Row, 1), dat(Row, 2), CStr(Row))
    Next Row
    
    'in place bubblesort:
    For i = 1 To inputData.Count - 1
        For j = i + 1 To inputData.Count
            If inputData(i)(0) > inputData(j)(0) Then
                'store the lesser item
                vTemp = inputData(j)
                'remove the lesser item
                inputData.Remove j
                're-add the lesser item before the
                'greater Item
                inputData.Add vTemp, vTemp(2), i
            End If
        Next j
    Next i
       
    calculateDaysWithoutOverlap = inputData(1)(1) - inputData(1)(0)
    intervalToCheck = Array(inputData(1)(0), inputData(1)(1))
    
    For i = 2 To inputData.Count
    
        startDate = inputData(i)(0)
        endDate = inputData(i)(1)
        
        If startDate <= intervalToCheck(1) And endDate > intervalToCheck(1) Then
            
            calculateDaysWithoutOverlap = calculateDaysWithoutOverlap + (endDate - intervalToCheck(1))
            intervalToCheck(1) = endDate
            
        ElseIf startDate > intervalToCheck(1) Then
        
            calculateDaysWithoutOverlap = calculateDaysWithoutOverlap + (endDate - startDate)
            intervalToCheck = Array(startDate, endDate)
        
        End If
    
    Next i

End Function

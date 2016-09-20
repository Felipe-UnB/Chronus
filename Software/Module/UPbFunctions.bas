Attribute VB_Name = "UPbFunctions"
Option Explicit


Function IsUDTvariableInitialized(AnyVariable As Variant) As Boolean
    
    Dim Counter As Integer
    
    On Error Resume Next
    Counter = LBound(AnyVariable)
    
        If Err.Number <> 0 Then
            IsUDTvariableInitialized = False
        Else
            IsUDTvariableInitialized = True
        End If
        
End Function

Function FindItemInArray(ByVal ItemToFind As Variant, ArrayToBeSearched As Variant) As Boolean

    Dim a As Integer
    Dim Type1 As String
    Dim Type2 As String
    
    If IsArrayEmpty(ArrayToBeSearched) = True Then
        MsgBox "ArrayToBeSearched is empty."
            End
    End If
    
    FindItemInArray = False
    
    Type1 = TypeName(ItemToFind)
    Type2 = TypeName(ArrayToBeSearched)
    
    'While ItemToFind must not be an array, ArrayToBeSearched must be an array.
    If IsArray(ItemToFind) = True Then
        MsgBox "ItemToFind must no be an array!"
            End
    ElseIf IsArray(ArrayToBeSearched) = False Then
        MsgBox "ArrayToBeSearched must be an array"
            End
    End If
    
    'The code below replaces any parentheses in Type1 and Type2 variables by nothing so we can compare them.
    Type1 = Replace(Type1, "(", "")
        Type1 = Replace(Type1, ")", "")
    Type2 = Replace(Type2, "(", "")
        Type2 = Replace(Type2, ")", "")
    
    If Type1 <> Type2 Then
        MsgBox "The type of ItemToFind(" & Type1 & ") is different to ArrayToBeSearched(" & Type2 & "), " & _
            "so it's not possible to compare both"
                End
    End If
    
    For a = LBound(ArrayToBeSearched) To UBound(ArrayToBeSearched)

        If ItemToFind = ArrayToBeSearched(a) Then
            FindItemInArray = True
                a = UBound(ArrayToBeSearched)
        End If
    Next


End Function


Function LineFitSlopeError(rng1 As Range, Rng2 As Range)
    
    'This function calculates the standard deviation (sometimes called standard error) of the intercept
    'of the line fit (errors of parameter A, from A + Bx; Taylor, 1997, Error Analysis). It supposes that
    'the uncertainty for each Yi is equal (Bevington and Robinson, 2002) and that Xi uncertanties are much
    'smaller than Yi.
    
    'Arguments
    'rng1 is the dependent (Y) variable range
    'rng2 is the independent (X) variable range
    
Dim cell As Integer
Dim rng1Count As Double 'Number of cells not empty in rng1
Dim rng2Count As Double 'Number of cells not empty in rng2
Dim E As Range 'Value of a specific cell in rng1
Dim f As Range 'Value of a specific cell in rng2
Dim SumXiSquared As Double 'sum of Xi^2, from i = 1 to n
Dim SumXi As Double 'sum of Xi, from i =1 to n

    
    rng1Count = rng1.Rows.count
    rng2Count = Rng2.Rows.count
    
    If Not rng1Count = rng2Count Then 'Both ranges must be of the same size!
        MsgBox ("rng1 and rng2 must be of equal size!")
            LineFitSlopeError = ""
                End
    End If
    
    
    cell = 1: SumXiSquared = 0: SumXi = 0
    
    While Not cell > rng1Count 'Loop through all values in the selected range
        
        Set E = rng1.Item(cell): Set f = Rng2.Item(cell)
        
        If IsEmpty(E) = False And IsEmpty(f) = False Then
            If WorksheetFunction.IsNumber(E) = True And WorksheetFunction.IsNumber(f) = True Then
                
                SumXiSquared = SumXiSquared + (f) ^ 2
                
                SumXi = SumXi + f
                            
            End If
        End If
        
        cell = cell + 1
    Wend
    
    LineFitSlopeError = Sqr(LineFitStdDev(rng1, Rng2) * rng1Count / (rng1Count * SumXiSquared - (SumXi) ^ 2))

End Function

Function LineFitYiPred(Y_Range As Range, X_Range As Range, Xi As Double) As Double

    'This function takes two ranges (dependent and independent variables) and, for
    'a Xi, calculates the predicted Yi based on a line fit.
    
    Dim lineSlope As Double
    Dim lineIntercept As Double
    Dim Y_RangeCount As Long
    Dim X_RangeCount As Long
    
    Y_RangeCount = Y_Range.Rows.count
    X_RangeCount = X_Range.Rows.count
    
    If Not Y_RangeCount = X_RangeCount Then 'Both ranges must be of the same size!
        MsgBox ("Y_RangeCount and X_RangeCount must be of equal size!")
            LineFitYiPred = ""
                End
    End If

    lineSlope = WorksheetFunction.Slope(Y_Range, X_Range)
    lineIntercept = WorksheetFunction.Intercept(Y_Range, X_Range)
    
    LineFitYiPred = lineIntercept + lineSlope * Xi

End Function
Function LineFitInterceptError(Y_Range As Range, X_Range As Range)

    'This function calculates the standard deviation (sometimes called standard error) of the intercept
    'of the line fit (errors of parameter A, from A + Bx; Taylor, 1997, Error Analysis). It supposes that
    'the uncertainty for each Yi is equal (Bevington and Robinson, 2002) and that Xi uncertanties are much
    'smaller than Yi.
    
    'Arguments
    'rng1 is the dependent (Y) variable range
    'rng2 is the independent (X) variable range
    
Dim cell As Integer
Dim Y_RangeCount As Double 'Number of cells not empty in Y_Range
Dim X_RangeCount As Double 'Number of cells not empty in X_Range
Dim E As Range 'Value of a specific cell in rng1
Dim f As Range 'Value of a specific cell in rng2
Dim SumXiSquared As Double 'sum of Xi^2, from i = 1 to n
Dim SumXi As Double 'sum of Xi, from i =1 to n

    
    Y_RangeCount = Y_Range.Rows.count
    X_RangeCount = X_Range.Rows.count
    
    If Not Y_RangeCount = X_RangeCount Then 'Both ranges must be of the same size!
        MsgBox ("Y_RangeCount and X_RangeCount must be of equal size!")
            LineFitInterceptError = ""
                End
    End If
    
    
    cell = 1: SumXiSquared = 0: SumXi = 0
    
    While Not cell > Y_RangeCount 'Loop through all values in the selected range
        
        Set E = Y_Range.Item(cell): Set f = X_Range.Item(cell)
        
        If IsEmpty(E) = False And IsEmpty(f) = False Then
            If WorksheetFunction.IsNumber(E) = True And WorksheetFunction.IsNumber(f) = True Then
                
                SumXiSquared = SumXiSquared + (f) ^ 2
                
                SumXi = SumXi + f
                            
            End If
        End If
        
        cell = cell + 1
    Wend
            
    LineFitInterceptError = Sqr(LineFitStdDev(Y_Range, X_Range) ^ 2 * (SumXiSquared / (Y_RangeCount * SumXiSquared - (SumXi) ^ 2)))

End Function

Function LineFitStdDev(rng1 As Range, Rng2 As Range)

    'This function calculates the sum of deviations between Yi (measured) and Yi,pred (Yi predicted
    'using the intercept and slope estimated by least squares fit), and then divide the sum by N - 2, where
    'N is the number of points. This is the STANDARD DEVIATION of the points in the line fit.
    
    'Expression taken from Bevington and Robinson, 2002

    
    'Arguments
    'rng1 is the dependent (Y) variable range
    'rng2 is the independent (X) variable range


Dim cell As Integer
Dim lineSlope As Double
Dim lineIntercept As Double
Dim rng1Count As Double 'Number of cells not empty in rng1
Dim rng2Count As Double 'Number of cells not empty in rng2
Dim E As Range 'Value of a specific cell in rng1
Dim f As Range 'Value of a specific cell in rng2
Dim a As Double

    
    rng1Count = rng1.Rows.count: rng2Count = Rng2.Rows.count

        If Not rng1Count = rng2Count Then 'Both ranges must be of the same size!
            MsgBox ("rng1 and rng2 must be of equal size!")
                LineFitStdDev = ""
                    End
        End If


    lineSlope = WorksheetFunction.Slope(rng1, Rng2)
    lineIntercept = WorksheetFunction.Intercept(rng1, Rng2)
        
    cell = 1: LineFitStdDev = 0: a = 0
    
    While Not cell > rng1Count 'Loop through all values in the selected range
        
        Set E = rng1.Item(cell): Set f = Rng2.Item(cell)
        
        If IsEmpty(E) = False And IsEmpty(f) = False Then
            If WorksheetFunction.IsNumber(E) = True And WorksheetFunction.IsNumber(f) = True Then
                a = a + (E - (lineIntercept + lineSlope * f)) ^ 2
            End If
        End If
        
        cell = cell + 1
    Wend
    
    ''debug.print A & " a"
    
    LineFitStdDev = Sqr(a / (rng1Count - 2))
        If LineFitStdDev = 0 Then
            MsgBox "It's not possible to calculate standard deviation for points of the line. All cells are empty or not populated with number"
                Application.GoTo rng1
                    End
        End If
    
End Function

Function SumSquaredDev(Rng As Range)
'Calculates the sum of squared deviations

Dim cell As Range
Dim a As Double

    a = WorksheetFunction.Average(Rng)


For Each cell In Rng 'Loop through all values in the selected range
    If Not cell = "" Then
        SumSquaredDev = SumSquaredDev + (cell.Value - a) ^ 2
    End If
Next cell

    If SumSquaredDev = "" Then
        MsgBox "It's not possible to calculate the sum of the squared deviations, all cells are empty."
            End
    End If
End Function

Function SumPrudDev(rng1 As Range, Rng2 As Range)

'Calculates the sum of the product of deviations between (Xi,Average X) and (Yi, Average Y).

Dim cell As Integer
Dim a As Integer 'Number of cells not empty in rng1
Dim B As Integer 'Number of cells not empty in rng2
Dim C As Double 'Average of rng1
Dim d As Double 'Average of rng2
Dim E As Range 'Value of a specific cell in rng1
Dim f As Range 'Value of a specific cell in rng2
    
    a = rng1.Rows.count: B = Rng2.Rows.count
        
    If Not a = B Then 'Both ranges must be of the same size!
        MsgBox ("rng1 and rng2 must be of equal size!")
        SumPrudDev = ""
        End
    End If
        
    
    C = WorksheetFunction.Average(rng1)
    d = WorksheetFunction.Average(Rng2)

    cell = 1
    
    While Not cell = a 'Loop through all values in the selected range
        
        Set E = rng1.Item(cell)
        Set f = Rng2.Item(cell)
        
        If IsEmpty(E) = False And IsEmpty(f) = False Then
            If WorksheetFunction.IsNumber(E) = True And WorksheetFunction.IsNumber(f) = True Then
                SumPrudDev = SumPrudDev + ((E - C) * (f - d))
            End If
        End If
        
        cell = cell + 1
    Wend

End Function

Function TimeCustomFormat(TimeCell As Range, Format As String)

    'I couldn't find way to make excel understand the Neptune time format (hh:mm:ss:ms),
    'so I had to find a way to deal with it. This program takes the ms, last 3 numbers of
    'cycle time, converts it to 1 day (time in excel can be expressed as decimal, 1=1 day)
    'and then add it to the hh:mm:ss. As you can see, I use an If structure, so I am able
    'to create different custom formats to other equipments if necessary
    
    'This function returns the time plus milliseconds in a way excel can understand it as time
    'TimeCell is where the time is and format is and indication of its format
    
    Dim ms As Double
    Dim msTOday As Long
    
    msTOday = 86400000
    
    Select Case Format
        
        Case "hh:mm:ss:ms(xxx)" 'Format used by Neptune software
    
            ms = Val(Right(TimeCell, 3)) / msTOday
                TimeCell = Left(TimeCell, Len(TimeCell) - 4) 'Pay attention to the "-4", it is there because time from neptune comes with ms hh:mm:ss:ms
                    TimeCell.NumberFormat = "h:mm:ss.000"
                        TimeCell = TimeCell + ms
                        
    End Select
    
    TimeCustomFormat = TimeCell
        
End Function

Function DateTimeCustomFormat(WB As Workbook, TimeCell As Range, DateCell As Range, formatTime As String, formatDate As String) As Double
    
    'I couldn't find way to make excel understand the Neptune time format (hh:mm:ss:ms),
    'so I had to find a way to deal with it.
    
    'Function created after a little discussion in the link below
    'https://stackoverflow.com/questions/26850609/how-to-convert-text-to-date-and-sum-to-time-by-vba
    
    'As you can see, I use an If structure, so I am able to create different custom formats
    'to other equipments if necessary.
    
    Dim ms As Double 'Means milliseconds
    Dim msTOday As Long 'Factor to convert ms to day (ms divided by all milliseconds in a day)
    Dim sTime As String, sDate As String 'String parts of the given parameters
    Dim dTime As Date, dDate As Date 'Calculated datetime values of the given parameters

    msTOday = 86400000

    On Error GoTo ErrHandler

    Select Case formatTime
        Case "hh:mm:ss:ms(xxx)"
            ms = Val(Right(TimeCell.Value, 3)) / msTOday
            sTime = Left(TimeCell.Value, Len(TimeCell.Value) - 4)
            'dTime = TimeValue(sTime) + ms 'please read help for TimeValue
            dTime = TimeSerial(Left(sTime, 2), Mid(sTime, 4, 2), Mid(sTime, 7, 2)) + ms
        Case Else
            dTime = 0
    End Select
        
    Select Case formatDate
        Case "Date: dd/mm/yyyy"
            sDate = Right(DateCell.Value, Len(DateCell.Value) - 6)
            'dDate = DateValue(sDate) 'please read help for DateValue
            dDate = DateSerial(Right(sDate, 4), Mid(sDate, 4, 2), Left(sDate, 2))
        Case Else
            dDate = 0
    End Select

    DateTimeCustomFormat = dTime + dDate
    
    Exit Function
    
ErrHandler:
    MsgBox "Date or time format (" & Chr(34) & formatDate & Chr(34) & " or " & Chr(34) & formatTime & Chr(34) & ") from sample is not correct. Please, check it."
        'Application.GoTo WB.Range("A1")
        Call UnloadAll
            End

End Function

Public Function SheetExists(SheetName As String, WB As Workbook) As Boolean

    'Modified from http://www.cpearson.com/excel/SheetNameFunctions.aspx
    'This function is better than SheetExists2 because it's not necessary
    'to check the name by name of each worksheet, something much faster.
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SheetExists
    ' This tests whether SheetName exists in a workbook.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim Ws As Worksheet
        
        On Error Resume Next
        Err.Clear
        Set Ws = WB.Worksheets(SheetName)
        
        If Err.Number = 0 Then
            SheetExists = True
        Else
            SheetExists = False
        End If
        
'original code
'Public Function SheetExists(SheetName As String, Optional R As Range) As Boolean
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' SheetExists
'    ' This tests whether SheetName exists in a workbook. If R is
'    ' present, the workbook containing R is used. If R is omitted,
'    ' Application.Caller.Worksheet.Parent is used.
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Dim WS As Worksheet
'        Dim WB As Workbook
'        If R Is Nothing Then
'            Set WB = Application.Caller.Worksheet.Parent
'        Else
'            Set WB = R.Worksheet.Parent
'        End If
'        on error resume Next
'        Err.Clear
'        Set WS = WB.Worksheets(SheetName)
'
'        If Err.Number = 0 Then
'            SheetExists = True
'        Else
'            SheetExists = False
'        End If
'    End Function

    End Function

Function SheetExists2(n As String, WB As Workbook) As Boolean
    'This function takes a name (n as string) and
    'and checks if this is the name of any of the
    'worksheets in a choosen workbook
  
  Dim Ws As Worksheet
  
  SheetExists2 = False
  
  For Each Ws In WB.Sheets
    
    If n = Ws.Name Then
      
      SheetExists2 = True
      
      Exit Function
    
    End If
  
  Next Ws
  
End Function

Function IntegerToStringArray(ByRef IntegerArray() As Integer)
    'This functions takes an array of integers and converts all of its items to string.
    'This was necessary in CreatSamListMap sub because to validate data entry in cells
    'of SamListMap I can only use strings.

    Dim a As Integer
    Dim B As Integer
    Dim NewArray() As String
    ReDim NewArray(UBound(IntegerArray)) As String
    B = LBound(IntegerArray)
    
    For a = B To UBound(IntegerArray)
        NewArray(a) = Str(IntegerArray(a))
    Next
    
    IntegerToStringArray = NewArray
        
        
End Function

Function IsUserFormLoaded(ByVal UFName As String) As Boolean

' gijsmo April 24th, 2011; http://www.ozgrid.com/forum/showthread.php?t=152892
    
    Dim UForm As Object
    Dim a As Variant
    
    a = UserForms.count
    
    IsUserFormLoaded = False
    For Each UForm In UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded

'Public Function IsArrayEmpty(Arr As Variant) As Boolean
'
''The code below was taken from http://www.cpearson.com/excel/vbaarrays.htm
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' IsArrayEmpty
'' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
''
'' The VBA IsArray function indicates whether a variable is an array, but it does not
'' distinguish between allocated and unallocated arrays. It will return TRUE for both
'' allocated and unallocated arrays. This function tests whether the array has actually
'' been allocated.
''
'' This function is really the reverse of IsArrayAllocated.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    Dim LB As Long
'    Dim UB As Long
'
'    Err.Clear
'    on error resume Next
'    If IsArray(Arr) = False Then
'        ' we weren't passed an array, return True
'        IsArrayEmpty = True
'    End If
'
'    ' Attempt to get the UBound of the array. If the array is
'    ' unallocated, an error will occur.
'    UB = UBound(Arr, 1)
'    If (Err.Number <> 0) Then
'        IsArrayEmpty = True
'    Else
'        ''''''''''''''''''''''''''''''''''''''''''
'        ' On rare occassion, under circumstances I
'        ' cannot reliably replictate, Err.Number
'        ' will be 0 for an unallocated, empty array.
'        ' On these occassions, LBound is 0 and
'        ' UBound is -1.
'        ' To accomodate the weird behavior, test to
'        ' see if LB > UB. If so, the array is not
'        ' allocated.
'        ''''''''''''''''''''''''''''''''''''''''''
'        Err.Clear
'        LB = LBound(Arr)
'        If LB > UB Then
'            IsArrayEmpty = True
'        Else
'            IsArrayEmpty = False
'        End If
'    End If
'
'End Function

Function FindStrings(WhatToFind As String, WhereToStart As Range, WhereToFinish As Range)

    'Function used to look for some strings in a range.
    'WhatWasFound is an array
    'This function returns an array with the cell addresses where the string chooseen was found

    Dim RangeStartFinish As Range
    Dim WhatWasFound() As Variant 'Array of BCO
    Dim FindString As Object
    Dim First As String
    Dim a As String

    
    WhatWasFound = Array()
    
    'These statements are necessary to
    ReDim Preserve WhatWasFound(0 To UBound(WhatWasFound) + 1) As Variant

    Set RangeStartFinish = Range(WhereToStart, WhereToFinish)
    
    If Len(WhereToStart.Value) = 0 Then
        FindStrings = Array() 'This is fundamental because even if the function does not find the "FindString", the function must return an empty array.
        'If this does not happen, an error of type mysmatch will raise
            Exit Function
        Else
          With RangeStartFinish
                Set FindString = .Find(WhatToFind)
                    
                    If FindString Is Nothing Then
                        FindStrings = Array()
                            Exit Function
                        
                    Else
                    
                        First = FindString.Address
                        Do
                            WhatWasFound(UBound(WhatWasFound)) = FindString.Address
                                Set FindString = .FindNext(FindString)
                                                                
                                    If FindString.Address = First Then
                                        Exit Do
                                    Else
                                        ReDim Preserve WhatWasFound(0 To UBound(WhatWasFound) + 1) As Variant
                                    End If
                                    
                        Loop While Not FindString Is Nothing And FindString.Address <> First
                    End If
            End With
        End If
        
    FindStrings = WhatWasFound
End Function

Function VarCovar(Rng As Range) As Variant

'Function copied with some adaptations from http://financeandnotes.blogspot.com.br/2010/08/vba-for-variance-covariance-matrix.html
'This function calculates the variance-covariance matrix of n variables, measured t times, returning a matrix N x N.
'It's necessary that both of the variables were copied to a worksheet.

    
    Dim i As Integer
    Dim J As Integer
    Dim colnum As Integer
    Dim matrix() As Double
        
    colnum = Rng.Columns.count 'Number of columns in CovarSheet (equal to the number of variables)
    ReDim matrix(1 To colnum, 1 To colnum)
        
    For i = 1 To colnum
        For J = 1 To colnum
            matrix(i, J) = Application.WorksheetFunction.Covariance_S(Rng.Columns(i), Rng.Columns(J))
        Next J
    Next i

    VarCovar = matrix

End Function

Function SumCovar(VarCovarArray As Variant)

    'This function sums all covariances between all variables, which are stored in
    'CovarArray, an array created by VarCovar function.

    Dim i As Integer
    Dim J As Integer
    Dim Summation As Double
    
    Summation = 0
    For i = 1 To UBound(VarCovarArray) - 1
        For J = i + 1 To UBound(VarCovarArray)
            Summation = Summation + VarCovarArray(i, J)
        Next
    Next
    
    SumCovar = Summation
    
End Function



Function ColumnsRowsNumber(RangesArray As Variant)
       
    'This function checks if all ranges have the same number of columns and rows
    
    Dim SameNumber As Boolean
    Dim Counter As Integer
    Dim RowNumber As Integer
    Dim ColumnNumber As Integer
    
    SameNumber = True
    
    RowNumber = RangesArray(1).Value
    ColumnNumber = RangesArray(1).Value
    
    For Counter = LBound(RangesArray) To UBound(RangesArray)
    
        Counter = Counter + 1
    
    Next
        
    
End Function


Function TetaFactor(a As Integer)

    'Based on the sample and external standards IDs, analysed before and after the sample,
    'this function calculates de teta factor, used to correct the samples by the external
    'standard.
    
    'Arguments - a is the index of the sample being calculated in AnalysesList

    Dim SlpTime As Double 'Time of sample first cycle
    Dim Std1Time As Double 'Time of standard (before) first cycle
    Dim Std2Time As Double 'Time of standard (after) first cycle
    Dim FindIDObj As Object

    SlpTime = PathsNamesIDsTimesCycles(4, AnalysesList(a).sample)
    Std1Time = PathsNamesIDsTimesCycles(4, AnalysesList(a).Std1)
    Std2Time = PathsNamesIDsTimesCycles(4, AnalysesList(a).Std2)
    
    If WorksheetFunction.IsNumber(SlpTime) = False Or WorksheetFunction.IsNumber(Std1Time) = False Or _
    WorksheetFunction.IsNumber(Std2Time) = False Then
        
        MsgBox "Please, check if the sample with ID equal to " & AnalysesList(a).sample & _
        " and its external standards are correct."
            
            'SamList_Sh.Activate
                
                With SamList_Sh.Columns(SamList_ID)
                    Set FindIDObj = .Find(AnalysesList(a).sample)
                    Application.GoTo .Range(FindIDObj.Address)
                End With
                    
                    End
    End If


        TetaFactor = (SlpTime - Std1Time) / (Std2Time - Std1Time)
        
'        Debug.Print TetaFactor
    
End Function

Function EntryIsValid(cell) As Variant

    'Code modified from http://www.java2s.com/Code/VBA-Excel-Access-Word/Excel/ValidatingdataentryinWorksheetchangeevent.htm
    
    If cell = "" Then
        EntryIsValid = True
        Exit Function
    End If

    If Not IsNumeric(cell) Then
        EntryIsValid = "Non-numeric entry."
        Exit Function
    End If

    If CInt(cell) <> cell Then
        EntryIsValid = "Integer required."
        Exit Function
    End If

    If cell < 1 Or cell > 12 Then
        EntryIsValid = "Valid values are between 1 and 12."
        Exit Function
    End If

    EntryIsValid = True
End Function

Function Array_Unique(arr As Variant) As Variant

    'Modified from http://www.jpsoftwaretech.com/useful-array-functions-for-vba-part-5/ 07012015
    
    'Scrub duplicates from an Array. If you have an array with one or more duplicates, this method
    'gets rid of them and returns the array sans duplicates. Currently the function only supports
    'one dimensional arrays.
    
    'This function only supports one dimensional arrays.
    
    
    Dim tempArray As Variant
    Dim i As Long
    
    If NumberOfArrayDimensions(arr) > 1 Then
        MsgBox "This function is only suitable for one dimensional arrays. This program will stop."
            End
    End If
    
    ' start the temp array with one element and
    ' populate with first value
    ReDim tempArray(0)
        tempArray(0) = arr(LBound(arr))
    
        For i = LBound(arr) To UBound(arr)
            If Not IsInArray(tempArray, arr(i)) Then  ' not in destination array
                ReDim Preserve tempArray(UBound(tempArray) + 1)
                    tempArray(UBound(tempArray)) = arr(i)
            End If
        Next i
        
    Array_Unique = tempArray
    End Function

Function IsInArray(arr As Variant, valueToFind As Variant) As Boolean

    'Taken from http://www.jpsoftwaretech.com/useful-array-functions-for-vba-part-5/ 07012015

    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        
        If StrComp(arr(i), valueToFind) = 0 Then
            IsInArray = True
                Exit For
        End If
    Next i

End Function

Function NonEmptyCellsRange(Rng As Range, rngFirstcell As Range, SH As Worksheet, Optional OnlyNumericCells As Boolean = False) As Range

    'This function takes the rng range, eliminates the empty cells and
    'returns a new range with only non empty cells. The optional argument
    'let the user choose if only cells with number will be added to the new
    'range.
    
    'ONLY RANGES WITH 1 AREA CAN BE PROCESSED
    
    'Arguments
    'Rng is the range where empty cells (and maybe non-numeric cells) should be ignored
    'Rngfirstcell is the range of the first cell of Rng
    'Sh is the worksheet where rng is set.
    'OnlyNumericCells is an option to the user choose if only cells with number should be copied.
    
    Dim ItemNumber As Integer 'Number of the range item
    Dim ItemsNewRange() As Double 'Array with the items number of cells that
    Dim CountCells As Long
    Dim NewItem As Double
    Dim ArrayItem As Variant
    Dim Counter As Integer
    Dim RedimCounter As Integer
    Dim IsThereEmptyElementArray As Boolean
    
    ReDim ItemsNewRange(1 To 1) As Double
    
    If Rng.Areas.count > 1 Then
        MsgBox "Only ranges with 1 area can be processed.", vbOKOnly
            End
    End If
    
    CountCells = Rng.count 'Number of cells in rng
    Counter = 1
    RedimCounter = 1
    
    For ItemNumber = 1 To CountCells
        
        NewItem = Rng.Item(ItemNumber)
        
        If IsEmpty(Rng.Item(ItemNumber)) = False And OnlyNumericCells = True Then 'Cell is not empty and user wants only cells with numbers
                    
                    If WorksheetFunction.IsNumber(Rng.Item(ItemNumber)) = True Then
                        
                        ItemsNewRange(RedimCounter) = NewItem
                            RedimCounter = RedimCounter + 1
                            
                    End If

                    If ItemNumber < CountCells Then 'This condition prevents ItemsNewRange from being redimensioned unnecessarily.
                        ReDim Preserve ItemsNewRange(1 To UBound(ItemsNewRange) + 1) As Double
                    End If

            ElseIf IsEmpty(Rng.Item(ItemNumber)) = False And OnlyNumericCells = False Then 'Cell is not empty and user wants all nonempty cells
            
                    ItemsNewRange(RedimCounter) = NewItem
                        RedimCounter = RedimCounter + 1
        
                    If ItemNumber < CountCells Then 'This condition prevents ItemsNewRange from being redimensioned unnecessarily.
                        ReDim Preserve ItemsNewRange(1 To UBound(ItemsNewRange) + 1) As Double
                    End If
        
        End If
            
10  Next

        If ItemsNewRange(UBound(ItemsNewRange)) = 0 Then 'The last array element will always be empty if the last item from Rng doesn't fail the previous test.
            IsThereEmptyElementArray = DeleteArrayElement(ItemsNewRange, UBound(ItemsNewRange), True)
        End If

        If IsArrayAllNumeric(ItemsNewRange) = True Then 'A new range will only be created if cells that pass the previous conditions were found
            
            Rng.Clear
            ItemNumber = 1
                
                For Each ArrayItem In ItemsNewRange
                
                    Rng.Item(ItemNumber) = ArrayItem
                    
                    ItemNumber = ItemNumber + 1
                    
                Next
                
            Set NonEmptyCellsRange = SH.Range(rngFirstcell, rngFirstcell.Offset(NumElements(ItemsNewRange, 1) - 1))
            
'            NonEmptyCellsRange.Select
        Else
                
'            If OnlyNumericCells = False Then
'                MsgBox "All cells are empty in the range.", vbOKOnly
'            Else
'                MsgBox "Cells with numbers were not found.", vbOKOnly
'            End If
                        
            Set NonEmptyCellsRange = Rng
                Rng = 0
            
        End If
            
End Function

'Function Age_Pb6U238(Ratio68 As Variant)
'
'    'Age based on 206Pb and 238U.
'
'    If WorksheetFunction.IsNumber(Ratio68) = False Then
'        Age_Pb6U238 = "Ratio must be a number."
'            Exit Function
'    End If
'
'    If Not Ratio68 < 0 Then
'        Age_Pb6U238 = (1 / Decay238U_yrs) * WorksheetFunction.Ln(Ratio68 + 1)
'    Else
'        Age_Pb6U238 = "Ratio68 must be a number > 0."
'    End If
'
'End Function
'
'Function Age_Pb7U235(Ratio75 As Variant)
'
'    'Age based on 207Pb and 235U.
'
'    If WorksheetFunction.IsNumber(Ratio75) = False Then
'        Age_Pb7U235 = "Ratio must be a number."
'            Exit Function
'    End If
'
'    If Not Ratio75 < 0 Then
'        Age_Pb7U235 = (1 / Decay235U_yrs) * WorksheetFunction.Ln(Ratio75 + 1)
'    Else
'        Age_Pb7U235 = "Ratio68 must be a number > 0."
'    End If
'
'End Function

Function InstalledIsoplot() As Boolean

    'Returns true if Isoplot4.15 is installed and loaded.

    Dim AddInInList As Boolean
    Dim IsoplotAddin As AddIn
    Dim Counter As Integer
    
    ScreenUpd = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    
    For Counter = 1 To AddIns.count
        If AddIns.Item(Counter).Name = "Isoplot4.15.xlam" Then
            AddInInList = True
                Counter = AddIns.count
        End If
    Next
    
    If AddInInList = True Then
    'On Error Resume Next
        Set IsoplotAddin = AddIns("Isoplot 4.15.11.10.15")
            If IsoplotAddin.Installed = True Then
                InstalledIsoplot = True
            Else
                Err.Clear
                On Error Resume Next
                    IsoplotAddin.Installed = True
                        If Err.Number <> 0 Then
                            InstalledIsoplot = False
                        End If
                On Error GoTo 0
            End If
        Else
            InstalledIsoplot = False
    End If
    
    Application.ScreenUpdating = ScreenUpd

End Function

Function Ratio68Concordant(age As Double, Decay238 As Double)
    
    'Calculates the ratio 68 for the indicated age. Decay constant must be in the same
    'unit as the age (years, millions of years, etc). Age must be 0 or any other positive
    'number.
    
    On Error GoTo BadEntry
    
    If age = 0 Then
        Ratio68Concordant = 0
    End If
    
    If age < 0 Then
        GoTo BadEntry
    End If
    
    Ratio68Concordant = Exp(Decay238 * age) - 1
    
    Exit Function
    
BadEntry:
    MsgBox "An error occurred." & vbNewLine & "Check if you entered a negative number"
    Ratio68Concordant = "Error"
    
    Exit Function
        
End Function

Function Ratio75Concordant(age As Double, Decay235 As Double)
    'Calculates the ratio 75 for the indicated age. Decay constant must be in the same
    'unit as the age (years, millions of years, etc). Age must be 0 or any other positive
    'number.
    
    On Error GoTo BadEntry
    
    If age = 0 Then
        Ratio75Concordant = 0
    End If
    
    If age < 0 Then
        GoTo BadEntry
    End If
    
    Ratio75Concordant = Exp(Decay235 * age) - 1
    
    Exit Function
    
BadEntry:
    MsgBox "An error occurred." & vbNewLine & "Check if you entered a negative number"
    Ratio75Concordant = "Error"
    
    Exit Function
        
End Function

Function Ratio76BasedOn6875Ratios(Ratio68 As Double, Ratio75 As Double, Ratio238235 As Double)
    'Calculates the ratio 75 for the indicated age. Decay constant must be in the same
    'unit as the age (years, millions of years, etc). Age must be 0 or any other positive
    'number.
    
    On Error GoTo BadEntry
    
    If Ratio68 = 0 Or Ratio75 = 0 Or Ratio238235 = 0 Then
        Ratio76BasedOn6875Ratios = 0
    End If
    
    If Ratio68 <= 0 Or Ratio75 <= 0 Then
        GoTo BadEntry
    End If
    
    Ratio76BasedOn6875Ratios = (Ratio75 / Ratio68) / Ratio238235
    
    Exit Function
    
BadEntry:
    MsgBox "An error occurred." & vbNewLine & "Check if you entered a negative number"
    Ratio76BasedOn6875Ratios = "Error"
    
    Exit Function
        
End Function

Function Ratio76fromInterations6875(age76 As Double, startRatio As Double, Optional Incre As Double = 0.00000001, Optional Delta As Double = 0.001, Optional Intercations As Long = 10000)

    'Calculates the 76 ratio based on the expected 76 age. startRatio is a guess that the user should give about the
    'the ratio. A good guess can reduce the amount of time necessary to approximate the ratio. Interactions are the
    'number of times that the program will increase the ratio by delta and compare the ages. Incre is the increment
    'added to startRatio every interaction. Delta is the acceptable difference between expected 76 age and the age
    'calculated by Isoplot considering the startRatio (increased or not).

    Dim Counter As Integer
    Dim age76Isoplot
    
    Incre = 0.00000001
    Delta = 0.001
    
    For Counter = 1 To Intercations
        age76Isoplot = agepb76(startRatio)
            
            If Sqr(WorksheetFunction.Power(age76Isoplot - age76, 2)) <= Delta Then
                Ratio76fromInterations6875 = startRatio
                    Exit Function
            End If
            
        startRatio = startRatio + Incre
    Next
    
    Ratio76fromInterations6875 = "Error"
    
    
End Function

Function IsStrike(rCell As Range)
' '******************************************************
'/------------------------------------------------------\
'|  Macro desenvolvida por: Felipe Valença de Oliveira  |
'|  Laboratório de Geocronologia - UnB                  |
'|  Primeira versão (v1): Agosto - 2012 (Felipe Valença)|
'\------------------------------------------------------/
'********************************************************

  IsStrike = rCell.Font.Strikethrough
End Function

Sub testStringsMatch()

    Dim TestString As String
    Dim ArrayString(1 To 7) As String
    Dim result1 As Boolean
    
    TestString = "AA"
    
    ArrayString(1) = "AA"
    ArrayString(2) = "As"
    ArrayString(3) = "BB"
    ArrayString(4) = "ASFFGDASAABFBF"
    ArrayString(5) = "REWREWGFAaHGFH"
    ArrayString(6) = "Aa"
    ArrayString(7) = "aA"
    
    Application.SendKeys "^g^a{DEL}"
    
    result1 = StringsMatch(TestString, ArrayString)

End Sub

Function StringsMatch(TargetString As String, StringsToCompare() As String)
    'This procedure will take an string and compare it with each element in an array of strings. The objective is
    'to know if the target string is equal to or is part of any of the array strings. This procedure is not case sentive.
    
    'RETURNS TRUE IF THE TARGETSTRING IS PRESENT IN ONE OF THE ARRAY ELEMENTS OR IS EQUAL TO THEM.
    
    'Created 27082015
    
    'UPDATED 02102015 - The procedure is no longer case sensitive.
    
    Dim FirstString As Long
    Dim LastString As Long
    Dim n As Long
    
    FirstString = LBound(StringsToCompare) 'Index of the first element in StringsToCompare array
    LastString = UBound(StringsToCompare) ''Index of the last element in StringsToCompare array
    
    StringsMatch = False
    
    For n = FirstString To LastString
'        Debug.Print TargetString & " - " & StringsToCompare(n)
'        Debug.Print TargetString = StringsToCompare(n)
'        Debug.Print "CONTAINS " & InStr(1, StringsToCompare(n), TargetString, vbBinaryCompare)
        
        If UCase(TargetString) = UCase(StringsToCompare(n)) Or _
            InStr(1, StringsToCompare(n), TargetString, vbBinaryCompare) <> 0 Then
            'First the program cheack if the strings are equal, then if the string in array contains the TargetString.
            
                StringsMatch = True
        End If

'        Debug.Print
        
    Next
'        Debug.Print StringsMatch
End Function

Function CompareAnalysisNames(ByRef TxtBox As MSForms.TextBox)

    'TxtBox is the textbox with the name of the analysis type that it will be checked.
    'If the name of any of the analyses types is duplicated, this function will return "ERROR", otherwise
    'it will return "OK"

    'Created 27082015
    
    Dim AnalysisNames() As String
    Dim SimilarName As Boolean
    Dim TargetName As String
    Dim ControlsArray() As MSForms.TextBox  'Array of all textboxes with analyses types names
    Dim Counter1 As Long
    Dim Counter2 As Long
    
    TargetName = TxtBox.Value
    CompareAnalysisNames = "OK"
    
    ReDim AnalysisNames(1 To 3) As String
    ReDim ControlsArray(1 To 4) As MSForms.TextBox
    
    Set ControlsArray(1) = Box1_Start.TextBox10_ExternalStandardName
    Set ControlsArray(2) = Box1_Start.TextBox9_SamplesNames
    Set ControlsArray(3) = Box1_Start.TextBox8_BlankName
    Set ControlsArray(4) = Box1_Start.TextBox5_InternalStandardName
    
    Counter1 = 1
    
        For Counter2 = 1 To 4
            If ControlsArray(Counter2).Name <> TxtBox.Name Then 'Compares the name of the textbox to avoid adding the TxtBox
                AnalysisNames(Counter1) = ControlsArray(Counter2).Value
                    Counter1 = Counter1 + 1
            End If
        Next
            
    SimilarName = StringsMatch(TargetName, AnalysisNames)
    
    If SimilarName = True Then
        CompareAnalysisNames = "ERROR"
            Exit Function
    End If

End Function

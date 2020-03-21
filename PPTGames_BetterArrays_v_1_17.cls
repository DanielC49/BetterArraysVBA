VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' ==============================================================
'
' ##############################################################
' #                                                            #
' #                   PPTGames Better Arrays                   #
' #                         List Class                         #
' #                                                            #
' ##############################################################
'
' » version 1.17
'
' » https://pptgamespt.wixsite.com/pptg-coding/better-arrays-vba
'
' ===============================================================



Option Explicit

Private ba_items
Private ba_length As Long
'

Property Get Items(Optional ByVal ItemIndex As Long = -1)
    If ba_length = 0 Then
        ThrowError 1, "List is empty."
    ElseIf ItemIndex < -1 Then
        ThrowError 2, "'" & ItemIndex & "' is not a valid ItemIndex. Minimum allowed is -1."
    ElseIf ItemIndex + 1 > ba_length Then
        ThrowError 8, "'" & ItemIndex & "' is not a valid ItemIndex. Maximum allowed is " & UBound(ba_items) & "."
    Else
        If ItemIndex = -1 Then
            Items = ba_items
        Else
            If IsObject(ba_items(ItemIndex)) Then
                Set Items = ba_items(ItemIndex)
            Else
                Items = ba_items(ItemIndex)
            End If
        End If
    End If
End Property

Property Let Items(Optional ByVal ItemIndex As Long = -1, NewItemValue)
Attribute Items.VB_UserMemId = 0
    If ba_length = 0 And ItemIndex > -1 Then
        ThrowError 1, "List is empty."
    ElseIf ItemIndex < -1 Then
        ThrowError 2, "'" & ItemIndex & "' is not a valid ItemIndex. Minimum allowed is -1."
    ElseIf ItemIndex + 1 > ba_length Then
        ThrowError 8, "'" & ItemIndex & "' is not a valid ItemIndex. Maximum allowed is " & UBound(ba_items) & "."
    Else
        If ItemIndex = -1 Then
            ba_items = NewItemValue
        Else
            If IsObject(NewItemValue) Then
                Set ba_items(ItemIndex) = NewItemValue
            Else
                ba_items(ItemIndex) = NewItemValue
            End If
        End If
        ba_length = UBound(ba_items) + 1
    End If
End Property

Sub SetItems(ParamArray Items())
    ba_items = Items
    ba_length = UBound(ba_items) + 1
End Sub

Sub AddItem(ByVal Item, Optional Index As Long = -1)
    If Index < -1 Then
        ThrowError 16, "'" & Index & "' is not a valid Index. Minimum allowed is -1."
    ElseIf Index + 1 > ba_length Then
        ThrowError 32, "'" & Index & "' is not a valid Index. Maximum allowed is " & UBound(ba_items) & "."
    Else
        If ba_length = 0 Then
            ba_items = Array(Item)
        Else
            If Index = -1 Then
                ReDim Preserve ba_items(UBound(ba_items) + 1)
                If IsObject(Item) Then
                    Set ba_items(UBound(ba_items)) = Item
                Else
                  ba_items(UBound(ba_items)) = Item
                End If
            Else
                Dim i As Long
                ReDim Preserve ba_items(0 To UBound(ba_items) + 1)
                For i = UBound(ba_items) To Index + 1 Step -1
                    If IsObject(ba_items(i - 1)) Then
                        Set ba_items(i) = ba_items(i - 1)
                    Else
                        ba_items(i) = ba_items(i - 1)
                    End If
                Next i
                If IsObject(Item) Then
                    Set ba_items(Index) = Item
                Else
                    ba_items(Index) = Item
                End If
            End If
        End If
        ba_length = UBound(ba_items) + 1
    End If
End Sub

Sub RemoveItem(ByVal ItemIndex As Long)
    If ba_length = 0 Then
        ThrowError 1, "List is empty."
    ElseIf ItemIndex < -1 Then
        ThrowError 4, "'" & ItemIndex & "' is not a valid ItemIndex. Minimum allowed is 0."
    ElseIf ItemIndex + 1 > ba_length Then
        ThrowError 8, "'" & ItemIndex & "' is not a valid ItemIndex. Maximum allowed is " & UBound(ba_items) & "."
    Else
        If UBound(ba_items) = 0 Then
            ba_items = Array(Empty)
            ba_length = 0
        Else
            Dim i As Long
            For i = ItemIndex To UBound(ba_items) - 1
                If IsObject(ba_items(i + 1)) Then
                    Set ba_items(i) = ba_items(i + 1)
                Else
                    ba_items(i) = ba_items(i + 1)
                End If
            Next
            ReDim Preserve ba_items(LBound(ba_items) To UBound(ba_items) - 1)
            ba_length = UBound(ba_items) + 1
        End If
    End If
End Sub

Sub Clear()
    ba_items = Array(Empty)
    ba_length = 0
End Sub

Function IndexOf(Item, Optional ByVal StartIndex As Long) As Long
    If ba_length = 0 Then
        IndexOf = -1
    ElseIf StartIndex < 0 Then
        ThrowError 64, "'" & StartIndex & "' is not a valid StartIndex. Minimum allowed is 0."
    Else
        Dim i As Long
        For i = StartIndex To UBound(ba_items)
            If ba_items(i) = Item Then IndexOf = i: Exit Function
        Next
        IndexOf = -1
    End If
End Function

Function Length() As Long
    Length = ba_length
End Function

Function Concat(OtherArray)
    If ba_length = 0 Then
        Concat = OtherArray
    Else
        Dim arr, i As Long
        arr = ba_items
        For i = 0 To UBound(OtherArray)
            ReDim Preserve arr(UBound(arr) + 1)
            If IsObject(OtherArray(i)) Then Set arr(UBound(arr)) = OtherArray(i) Else: arr(UBound(arr)) = OtherArray(i)
        Next
        Concat = arr
    End If
End Function

Function Reverse()
    If ba_length = 0 Then
        ThrowError 1, "List is empty."
    Else
        On Error GoTo handler
        If UBound(ba_items) = 1 Then
            Reverse = Array(ba_items(1), ba_items(0))
        Else
            Dim arr(), lng As Long, i As Long, tmp
            arr = ba_items
            lng = UBound(arr)
            For i = LBound(arr) To ((UBound(arr) - LBound(arr) + 1) \ 2)
                If IsObject(arr(i)) Then Set tmp = arr(i) Else: tmp = arr(i)
                If IsObject(arr(lng)) Then Set arr(i) = arr(lng) Else: arr(i) = arr(lng)
                If IsObject(tmp) Then Set arr(lng) = tmp Else: arr(lng) = tmp
                lng = lng - 1
            Next
            Reverse = arr
        End If
        Exit Function
handler: Err.Raise 90, "PPTGames Better Arrays - List", "Unexpected error. Array might be empty."
    End If
End Function

Function Sort()
    If ba_length = 0 Then
        ThrowError 1, "List is empty."
    Else
        Dim arr(), strTemp, i As Long, j As Long, lngMin As Long, lngMax As Long
        arr = ba_items
        lngMin = LBound(arr)
        lngMax = UBound(arr)
        For i = lngMin To lngMax - 1
            For j = i + 1 To lngMax
                If IsObject(arr(i)) Or IsObject(arr(j)) Then Err.Raise 114, "PPTGames Better Arrays - List", "Can't sort List with objects."
                If arr(i) > arr(j) Then
                    strTemp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = strTemp
                End If
            Next j
        Next i
        Sort = arr
    End If
End Function

Function Slice(ByVal StartIndex As Long, ByVal EndIndex As Long)
    If ba_length = 0 Then
        ThrowError 1, "List is empty."
    Else
        Dim i As Long, arr()
        ReDim Preserve arr(0)
        arr(0) = ""
        For i = StartIndex To EndIndex
        ReDim Preserve arr(UBound(arr) + 1)
            If IsObject(ba_items(i)) Then
                Set arr(UBound(arr)) = ba_items(i)
            Else
                arr(UBound(arr)) = ba_items(i)
            End If
        Next
        For i = 0 To UBound(arr) - 1
            If IsObject(arr(i + 1)) Then
                Set arr(i) = arr(i + 1)
            Else
                arr(i) = arr(i + 1)
            End If
        Next
        ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
        Slice = arr
    End If
End Function

Function Join(Optional ByVal Separator As String) As String
    If ba_length = 0 Then
        Join = ""
    Else
        Dim i As Long
        For i = 0 To UBound(ba_items)
            If i = UBound(ba_items) Then
                Join = Join & IIf(IsObject(ba_items(i)), "Object", ba_items(i))
            Else
                Join = Join & IIf(IsObject(ba_items(i)), "Object", ba_items(i)) & Separator
            End If
        Next
    End If
End Function


Private Sub ThrowError(Number As Long, Description As String)
    Err.Raise Number, "List", Description
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ba_items()

Property Get Items(Optional ItemIndex As Long = -1)
On Error GoTo handler
Attribute Items.VB_UserMemId = -1
If ItemIndex = -1 Then
Items = ba_items
Else
Items = ba_items(ItemIndex)
End If
Exit Property
handler: Items = ""
End Property

Property Let Items(Optional ItemIndex As Long = -1, NewItemValue)
Attribute Items.VB_UserMemId = 0
If ItemIndex = -1 Then
ba_items = NewItemValue
Else
ba_items(ItemIndex) = NewItemValue
End If
End Property

Sub AddItem(Item, Optional Index As Long = -1)
On Error GoTo handler
If Index = -1 Then
    ReDim Preserve ba_items(UBound(ba_items) + 1)
    ba_items(UBound(ba_items)) = Item
Else
    Dim i As Long
    ReDim Preserve ba_items(0 To UBound(ba_items) + 1)
    For i = UBound(ba_items) To Index + 1 Step -1
        ba_items(i) = ba_items(i - 1)
    Next i
    ba_items(Index) = Item
End If
Exit Sub
handler: If Index = -1 Then ba_items = Array(Item) Else: ba_items(UBound(ba_items)) = Item
End Sub

Sub RemoveItem(ItemIndex)
    If UBound(ba_items) = 0 Then
        ba_items = Array(Empty)
    Else
        Dim i As Long
        For i = ItemIndex To UBound(ba_items) - 1
            ba_items(i) = ba_items(i + 1)
        Next
        ReDim Preserve ba_items(LBound(ba_items) To UBound(ba_items) - 1)
    End If
End Sub

Function Length()
On Error GoTo handler
Length = UBound(ba_items) + 1
Exit Function
handler: Length = 0
End Function

Function IndexOf(Item, Optional StartIndex As Long)
    Dim i As Long
    For i = StartIndex To UBound(ba_items)
        If ba_items(i) = Item Then IndexOf = i: Exit Function
    Next
    IndexOf = -1
End Function

Function Concat(OtherArray)
    Dim arr(), i As Long
    arr = ba_items
    For i = 0 To UBound(OtherArray)
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = OtherArray(i)
    Next
    Concat = arr
End Function

Function Reverse()
    If UBound(ba_items) = 1 Then
        Reverse = Array(ba_items(1), ba_items(0))
    Else
        Dim arr(), lng As Long, i As Long, tmp
        arr = ba_items
        lng = UBound(arr)
        For i = LBound(arr) To ((UBound(arr) - LBound(arr) + 1) \ 2)
            tmp = arr(i)
            arr(i) = arr(lng)
            arr(lng) = tmp
            lng = lng - 1
        Next
        Reverse = arr
    End If
End Function

Function Sort()
    Dim arr(), strTemp, i As Long, j As Long, lngMin As Long, lngMax As Long
    arr = ba_items
    lngMin = LBound(arr)
    lngMax = UBound(arr)
    For i = lngMin To lngMax - 1
        For j = i + 1 To lngMax
            If arr(i) > arr(j) Then
                strTemp = arr(i)
                arr(i) = arr(j)
                arr(j) = strTemp
            End If
        Next j
    Next i
    Sort = arr
End Function

Function Slice(StartIndex As Long, EndIndex As Long)
    Dim i As Long, arr()
    ReDim Preserve arr(0)
    arr(0) = ""
    For i = StartIndex To EndIndex
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = ba_items(i)
    Next
    For i = 0 To UBound(arr) - 1
        arr(i) = arr(i + 1)
    Next
    ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
    Slice = arr
End Function

Function Join(Optional Separator As String)
    Dim i As Long
    For i = 0 To UBound(ba_items)
        If i = UBound(ba_items) Then
            Join = Join & ba_items(i)
        Else
            Join = Join & ba_items(i) & Separator
        End If
    Next
End Function

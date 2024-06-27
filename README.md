# About
Adds a List class that works just like an array but is much easier to use and features useful functions.

To learn more visit: https://pptgamespt.wixsite.com/pptg-coding/better-arrays


# Documentation

#### Better Arrays - version 1.12 - 1.20
> NOTE: This documentation only applies to version 1.12 - 1.17 and 1.20. To see the documentation of version 1.10, download the .zip file of the 1.10 version and open the Documentation.txt file.

# Get Started
Create a new Better Arrays List. Better Arrays' arrays are called Lists.

Example [VBA]:
```Dim myList As New List```

# Properties and Methods
Items
Sets or gets one or more items of a List. The items can be of any type. When trying to get an item of an invalid index, an error is returned.

NOTE: In version 1.20, the Items property has changed.
Syntax
# version 1.12 - 1.17

<ListName>.Items[(<ItemIndex>)]

\- or -

<ListName>[(<ItemIndex>)]

ItemIndex - index of the item to select. Default: -1. When is set to -1: selects all items from the List.

​
# version 1.20
​
<ListName>.Items
Example [VBA]
' version 1.12 - 1.17
​
```vb
Dim myList As New List
myList.Items = Array("item1", "item2") 'Sets the items to be equal to an array
​
Debug.Print myList.Items(0) 'Prints the first item of the List, in this case "item1"
Debug.Print myList(1) 'Prints the second item of the list, in this case "item2"
​```
​
' version 1.20
​
Dim myList As New List
myList.Items = Array("item1", "item2") 'Sets the items to be equal to an array
​
' myList.Items --> returns an array, so you can use ..Items(<index>) to return an item from the list/array.
Debug.Print myList.Items(0) 'Prints the first item of the returned array, in this case "item1"
Item (v1.20 only)
Sets or gets an item of a List. The item can be of any type. When trying to get an item of an invalid index, an error is returned.
Syntax
<ListName>.Item[(<ItemIndex>)]
- or -
<ListName>[(<ItemIndex>)]
​
ItemIndex - index of the item to select.
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2") 'Sets the items to be equal to an array
​
Debug.Print myList.Item(0) 'Prints the first item of the List, in this case "item1"
Debug.Print myList(1) 'Prints the second item of the list, in this case "item2"
SetItems
Sets the items of a List.
Syntax
<ListName>.SetItems(<Item>)
​
Items - ParamArray of items to set.
Example [VBA]
Dim myList As New List
myList.SetItems "item1", "item2" ' Sets the items of the list to "item1", "item2".
Join
Joins all items of the list and returns them as a string.
Syntax
<ListName>.Join[(<Separator>)]
​
Separator - Expression that separates each item. Default: "" (empty string).
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2", "item3")
​
Debug.Print myList.Join(", ") 'Prints the following string: "item1, item2, item3".
AddItem
Adds a new item to the List.
Syntax
<ListName>.AddItem <Item>[, Index]
​
Item - Item to add
Index - Index at which the item should be added. If it's -1 then the item is added at the end of the List. Default: -1.
Example [VBA]
Dim myList As New List
myList.Items = Array("item1")
​
myList.AddItem "item2" 'Adds the item "item2" to the List. The list is now "item1", "item2".
myList.AddItem "item3", 1 'Adds the item "item3" at index 1 to the List. The list is now "item1", "item3", "item2".
RemoveItem
Removes an item from the List.
Syntax
<ListName>.RemoveItem <ItemIndex>
​
ItemIndex - index of the item to remove.
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2")
​
myList.RemoveItem 0 'Removes the first item. Now the list only contains the item "item2".
Clear (v1.15 and above only)
Removes all items from the List. Only applies to versions 1.15 and 1.16.
Syntax
<ListName>.Clear
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2")
​
myList.Clear 'Removes all items. Now the list contains no items and the length is 0.
Length
Returns the number of items of the list.
Syntax
<ListName>.Length
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2")
​
Debug.Print myList.Length 'Prints 2.
IndexOf
Returns the first index at which a given element can be found in the array, or -1 if it is not present.
Syntax
<ListName>.IndexOf(<Item>[, StartIndex])
​
Item - item to locate in the list.
StartIndex - index to start looking for the item. Default: 0.
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2", "item1")
​
Debug.Print myList.IndexOf("item1") 'Prints 0.
Debug.Print myList.IndexOf("item1", 1) 'Prints 2.
Debug.Print myList.IndexOf("item2") 'Prints 1.
Debug.Print myList.IndexOf("item3") 'Prints -1, because the list does not contain "item3".
Reverse
Returns the list but reversed.
Syntax
<ListName>.Reverse
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2", "item3")
​
Debug.Print myList.Reverse 'Print the list but reversed.
myList.Items = myList.Reverse 'Sets the items of the list to the same list but reversed. Basically it reverses the list. The list is now "item3", "item2", "item1".
Sort
Returns the list but sorted.
Syntax
<ListName>.Sort
Example [VBA]
Dim myList As New List
myList.Items = Array("item3", "item1", "item2")
​
Debug.Print myList.Sort 'Print the list but sorted.
myList.Items = myList.Reverse 'Sets the items of the list to the same list but sorted. Basically it sorts the list. The list is now "item1", "item2", "item3".
Slice
Returns specific items from a list.
Syntax
<ListName>.Slice(StartIndex, EndIndex)
​
StartIndex - index of the first item
EndIndex - index of the last item
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2", "item3", "item4", "item5")
​
myList.Items = myList.Slice(1, 3) 'Sets the items of the list to the same list but sliced. Basically it slices the list. The list is now "item2", "item3", "item4".
Concat
Returns the union of the current list with an array.
Syntax
<ListName>.Concat(OtherArray)
​
OtherArray - array which will be concated.
Example [VBA]
Dim myList As New List
myList.Items = Array("item1", "item2")
​
myList.Items = myList.Concat(Array("item3", "item4") 'Sets the items of the list to the returned array, which is the union of the list and the array. The list is now "item1", "item2", "item3", "item4".
Errors (v1.16 only)
1: "List is empty."
The method or property does not allow the List to be empty.
​
2: "'%Item_Index%' is not a valid ItemIndex. Minimum allowed is -1."
You tried to run a method, get or set a property and provided an index which is less than -1.
​
4: "'%Item_Index%' is not a valid ItemIndex. Minimum allowed is 0."
You tried to run a method, get or set a property and provided an index which is than 0. Lists can only indices greater than or equal to 0.
​
8: "'%Item_Index%' is not a valid ItemIndex. Maximum allowed is %MAX_INDEX%."
You tried to run a method, get or set a property and provided an index which is greater than MAX_INDEX. Lists can only have indices less than or equal to MAX_INDEX. MAX_INDEX = ListLength - 1
​
16: "'%Index%' is not a valid Index. Minimum allowed is -1."
You tried to run a method, get or set a property and provided an index which is less than -1.
​
32: "'%Index%' is not a valid Index. Maximum allowed is %MAX_INDEX%."
You tried to run a method, get or set a property and provided an index which is greater than MAX_INDEX. Lists can only have indices less than or equal to MAX_INDEX. MAX_INDEX = ListLength - 1
​
64: "'%StartIndex%' is not a valid StartIndex. Minimum allowed is -1."
You tried to run a method, get or set a property and provided an index which is greater than MAX_INDEX. Lists can only have indices less than or equal to MAX_INDEX. MAX_INDEX = ListLength - 1

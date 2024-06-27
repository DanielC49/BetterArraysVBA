# About
Adds a List class that works just like an array but is much easier to use and features useful functions.

To learn more visit: https://pptgamespt.wixsite.com/pptg-coding/better-arrays


# Documentation

#### Better Arrays - version 1.12 - 1.20
> NOTE: This documentation only applies to version 1.12 - 1.17 and 1.20. To see the documentation of version 1.10, download the .zip file of the 1.10 version and open the Documentation.txt file.

# Get Started
Create a new Better Arrays List. Better Arrays' arrays are called Lists.

Example \[VBA]:
```vb
Dim myList As New List
```

# Properties and Methods
Items
Sets or gets one or more items of a List. The items can be of any type. When trying to get an item of an invalid index, an error is returned.

NOTE: In version 1.20, the Items property has changed.
Syntax
## version 1.12 - 1.17

`<ListName>.Items[(<ItemIndex>)]`

\- or -

`<ListName>[(<ItemIndex>)]`

`ItemIndex` - index of the item to select. Default: -1. When is set to -1: selects all items from the List.

​
## version 1.20
​
`<ListName>.Items`
Example \[VBA]
```vb
' version 1.12 - 1.17
​
Dim myList As New List
myList.Items = Array("item1", "item2") 'Sets the items to be equal to an array
​
Debug.Print myList.Items(0) 'Prints the first item of the List, in this case "item1"
Debug.Print myList(1) 'Prints the second item of the list, in this case "item2"
​
' version 1.20
​
Dim myList As New List
myList.Items = Array("item1", "item2") 'Sets the items to be equal to an array
​
' myList.Items --> returns an array, so you can use ..Items(<index>) to return an item from the list/array.
Debug.Print myList.Items(0) 'Prints the first item of the returned array, in this case "item1"
```

#### See full documentation [here](https://pptgamespt.wixsite.com/pptg-coding/better-arrays).

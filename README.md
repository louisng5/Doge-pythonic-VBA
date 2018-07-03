# Doge-pythonic-VBA
A VBA implementation of list object, DogeList (inspired by python), Higher-order and lambda function, DogeFunc

Usage:
1) Import both Doge.bas, DogeList and DogeFung into your VBA project.
2) Add reference to Microsoft Visual Basic for Applications Extensibility 5.3 (VBA Editor -> Tool -> References)
3) Make sure this is checked: Macro Settings -> Developer Macro Settings -> "Trust access to the VBA project object model"
4) Using lambda function will complie code at run time which will disable breakpoints. To enable breakpoint while using lambda, go to DogeFunc.cls, find the following code
```VBA
#Const DEBUGMODE = False
```
change DEBUGMODE to True
```VBA
#Const DEBUGMODE = True
```
This will make all the lambda functino be compiled in a new workbook (slower performance) thus allowing breakpoints.

DogeList Example
```VBA
arr = Array(1, 2, 3, 4, 5)
Dim d As New DogeList
d.assigndata arr
Debug.Print d.join(",")
'1,2,3,4,5

d.pprint
'1,2,3,4,5   (=Debug.Print d.join(delimiter:= "," text_qualifier := """", newline_char := vbNewLine))

'Helpful factory function
Set d = newlist(arr)
d.pprint
'1,2,3,4,5

'Pythonic slicing
newlist(arr)("0:3").pprint
'1,2,3
newlist(arr)("3:0:-1").pprint
'4,3,2
newlist(arr)("3::-1").pprint
'4,3,2,1

'Pythonic slicing multidimension
Dim arr2(0 To 2, 0 To 2)
arr2(0, 0) = 1
arr2(1, 0) = 2
arr2(2, 0) = 3
arr2(0, 1) = 4
arr2(1, 1) = 5
arr2(2, 1) = 6
arr2(0, 2) = 7
arr2(1, 2) = 8
arr2(2, 2) = 9

'2D array is transformed to DogeList of DogeList
Set d = newlist(arr2)

'Multidimensional support

Debug.Print d(0)(1)
'2

d.pprint
'1,2,3
'4,5,6
'7,8,9

d("0:2", "0:2").pprint
'1,2
'4,5

'printing object with pointer
newlist(Array(Sheets(1), Sheets(1))).pprint
'Worksheet@140684903964672,Worksheet@140684903964672

'lambda! what!? yeah
newlist(Array(1, 2, 3, 4)).lambdaMap("fx(x): fx = x + 1").pprint
'2,3,4,5
'The lambda string is compiled into -->  function fx(x): fx = x + 1: end function


'lambda with external variable
tmp = 99
newlist(Array(1, 2, 3, 4)).lambdaMap("myfunc(x,y): myfunc = x + y", tmp).pprint
'100,101,102,103
```

DogeFunc Example
```VBA
Function join(x As String, y As String, z As String) As String
    join = x & y & z
End Function

Sub higher_order_function()
'basic usage
Dim func As New DogeFunc
func.init ("join")
Debug.Print func.exc("A", "B", "C")

'Helper function
Set f = newfunc("join")
Debug.Print f.exc("A", "B", "C")
'ABC

'Python function warper
Set f = newfunc("join", "A")
Debug.Print f.exc("B", "C")
'ABC

'Second parameter
Debug.Print newfunc("join", , "A").exc("B", "C")
'BAC

'lambda
Debug.Print lambda("fx(x,y,z): fx = x & y & z").exc("A", "B", "C")
'ABC

'Use DogeFunc object with DogeList
Dim fn As DogeFunc
Set lst = newlist(Array(1, 2, 3))
Set fn = lambda("fx(x): fx = x + 1")
lst.map(fn).pprint
'2,3,4
End Sub
```

More DogeList function
```VBA
Sub other_DogeList_function()
'where
arr = Array(-2, -1, 0, 1, 2, 3)
'newlist(arr).where(lambda("fx(x): fx = (x > 0)")).pprint
'1,2,3

'append
Set lst = newlist(arr)
lst.append (4)
lst.pprint
'-2,-1,0,1,2,3,4

lst.remove (0)
lst.pprint
'-1,0,1,2,3,4

'default property of DogeList is its underlying array
Dummyarr = lst
Debug.Print TypeName(Dummyarr)
'V()

'Referencing the DogeList object itself
Debug.Print TypeName(lst.self)
'DogeList

'iteration of the items in DogeList
For Each item In lst.items
    
Next

'recursively assign data to DogeList
Dim col As New Collection
Dim col2 As New Collection
For i = 0 To 3
    col.add i
Next

For i = 0 To 3
    col2.add col
Next

newlist(col2).pprint
'Collection@106102915383600,Collection@106102915383600,Collection@106102915383600,Collection@106102915383600

Set lst = New DogeList
lst.assigndata col2, recursive:=True
lst.pprint
'0,1,2,3
'0,1,2,3
'0,1,2,3
'0,1,2,3
End Sub
```

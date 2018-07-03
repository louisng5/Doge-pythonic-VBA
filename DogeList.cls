VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DogeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


''
' DogeList V1.0
' https://github.com/louisng5/Doge-pythonic-VBA
'
' Pythonic List Object for VBA
'
' @author Louis NG, louisng5@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2017, Louis Ng
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
Private data_()

Property Get items(ParamArray a())
Attribute items.VB_UserMemId = 0
    If IsMissing(a) Then
        items = data_()
        Exit Property
    End If
    If VarType(a(0)) = vbString Then
        Dim param
        param = a
        Set items = slicing(param, 0)
        Exit Property
    End If
    If IsNumeric(a(0)) Then
        assignvalue items, data_(a(0))
        Exit Property
    End If
End Property
Property Let items(ParamArray a(), data)
Attribute items.VB_UserMemId = 0
    If Not IsMissing(a) Then
        data_(a(0)) = data
        Exit Property
    End If
    assigndata data
End Property
Property Set items(ParamArray a(), data)
Attribute items.VB_UserMemId = 0
    If Not IsMissing(a) Then
        Set data_(a(0)) = data
    Else
        Set data_ = data
    End If
    assigndata data
End Property

Property Get itemArray()
    itemArray = data_
End Property

Property Get self()
    Set self = Me
End Property

Property Get count()
    count = UBound(data_) + 1
End Property
Property Get upbound()
    upbound = UBound(data_)
End Property

Public Function slicing(param, axis As Integer) As DogeList
Dim slicestring As String: slicestring = param(axis)
If InStr(slicestring, ":") = 0 Then Err.Raise 9992, , "Invalid slicing"
        Dim tmp: tmp = Split(slicestring, ":")
        
        Dim istep: istep = 1
        If UBound(tmp) = 2 Then
            If Len(tmp(2)) > 0 Then
                istep = CInt(tmp(2))
                If istep = 0 Then Err.Raise 9992, , "slice step cannot be zero"
            End If
        End If
        
        Dim istart
        If istep < 0 Then
            istart = Me.upbound
        Else
            istart = 0
        End If
        If Len(tmp(0)) > 0 Then
            istart = CInt(tmp(0))
            If istart < 0 Then istart = Me.count + istart
            If istart < -1 Then istart = IIf(istep < 0, -1, 0)
            If istart > Me.upbound Then istart = Me.upbound
        End If
        
        Dim iend
        If istep < 0 Then
            iend = -1
        Else
            iend = Me.count
        End If
        If Len(tmp(1)) > 0 Then
            iend = CInt(tmp(1))
            If iend < 0 Then iend = Me.count + iend
            If iend < 0 Then iend = -1
            If iend > Me.count Then iend = Me.count
        End If
        
        Dim tmpdogelist As New DogeList
        Dim i
        For i = istart To iend Step istep
            If Not (i < 0 Or i = iend Or i > Me.upbound) Then
                If axis <> UBound(param) Then
                    tmpdogelist.append (Me(i).slicing(param, axis + 1))
                Else
                    tmpdogelist.append (Me(i))
                End If
            End If
        Next
        Set slicing = tmpdogelist

End Function

Public Sub append(obj)
    Dim i
    i = UBound(data_) + 1
    ReDim Preserve data_(i)
    assignvalue data_(i), obj
End Sub

Public Sub insert(obj, index As Integer)
    If index < 0 Then Err.Raise 9995, , "Index cannot be smaller than zero"
    Dim i
    i = UBound(data_) + 1
    ReDim Preserve data_(i)
    If index > UBound(data_) Then
        index = UBound(data_)
    End If
    For i = UBound(data_) To 0 Step -1
        If index = i Then
            assignvalue data_(i), obj
            Exit For
        End If
        assignvalue data_(i), data_(i - 1)
    Next
End Sub

Public Sub pprint()
    Debug.Print Me.join(",", """")
End Sub


Public Sub remove(index As Integer)
    If index < 0 Then Err.Raise 9995, , "Index cannot be smaller than zero"
    If index > UBound(data_) Then Err.Raise 9995, , "Index out of range"
    Dim i
    For i = index + 1 To UBound(data_)
        assignvalue data_(i - 1), data_(i)
    Next
    ReDim Preserve data_(UBound(data_) - 1)
End Sub

Public Sub assigndata(pdata, Optional recursive As Boolean = False, Optional axis As Integer = 0)
    Dim data
    
    If TypeOf pdata Is DogeList Then
        data = pdata.itemArray
    Else
        assignobject data, pdata
    End If
    
    If Not isIterable(data) Then Err.Raise 9991, , "Data not is not Iterable"
    
    If ismultidimensionarray(data) Then
        data_ = multiDimensionalArrayToDogelist(data, axis).itemArray
        Exit Sub
    End If
    
    Dim tmpobj, i, tmpdogelist
    Dim tmparray: tmparray = Array()
    For Each tmpobj In data: Exit For
    Next
    'Performance optimization for collection of object and collection of varient
    If recursive Then
        On Error GoTo e1
        For Each tmpobj In data
            i = UBound(tmparray) + 1
            ReDim Preserve tmparray(i)
            If isIterable(tmpobj) Then
                Set tmpdogelist = New DogeList
                tmpdogelist.assigndata tmpobj, True
                Set tmparray(i) = tmpdogelist
            Else
                Set tmparray(i) = tmpobj
            End If
        Next
    Else
        If IsObject(tmpobj) Then
            On Error GoTo e1
            For Each tmpobj In data
                i = UBound(tmparray) + 1
                ReDim Preserve tmparray(i)
                Set tmparray(i) = tmpobj
            Next
    
        Else
            On Error GoTo e2
            For Each tmpobj In data
                i = UBound(tmparray) + 1
                ReDim Preserve tmparray(i)
                tmparray(i) = tmpobj
            Next
        End If
    End If
    data_ = tmparray
Exit Sub
e1:
        tmparray(i) = tmpobj
        Resume Next
e2:
        Set tmparray(i) = tmpobj
        Resume Next
End Sub

Private Function multiDimensionalArrayToDogelist(arr, Optional axis As Integer = 0) As DogeList
Dim d As Variant: d = getDimensionality(arr)
Dim base As New DogeList
Dim arr_ubound As Integer
Dim tmp_array
Dim endnodes As New Collection
buildtree base.self, d, UBound(d), axis, endnodes
Dim node As DogeList
Dim innercount: innercount = 0
Dim outtercount: outtercount = 0
Dim numsmallbatch: numsmallbatch = 1
Dim numlargebatch: numlargebatch = 1
Dim i
If axis > 0 Then
    For i = 0 To axis - 1
        numsmallbatch = numsmallbatch * (d(i) + 1)
    Next
End If
If axis < UBound(d) Then
    For i = axis + 1 To UBound(d)
        numlargebatch = numlargebatch * (d(i) + 1)
    Next
End If
Dim value, batchnum, loopcount, arrlen
batchnum = 1
loopcount = 0
arrlen = d(axis) + 1
For Each value In arr
    innercount = innercount + 1
    endnodes(innercount + outtercount * numsmallbatch).append (value)
    loopcount = loopcount + 1
    
    If innercount = numsmallbatch Then innercount = 0
    If loopcount = numsmallbatch * arrlen Then outtercount = outtercount + 1: loopcount = 0
Next
Set multiDimensionalArrayToDogelist = base
End Function

Private Function buildtree(base As DogeList, dimensionality As Variant, dimension As Integer, axis As Integer, endnodes As Collection)
    Dim i
    Dim tmpdogelist As DogeList
    If axis = dimension Then
        dimension = dimension - 1
    End If
    If dimension < 0 Then
        endnodes.add base.self
        Exit Function
    End If
    For i = 0 To dimensionality(dimension)
        Set tmpdogelist = New DogeList
        base.append tmpdogelist.self
        buildtree base:=tmpdogelist.self, dimensionality:=dimensionality, dimension:=dimension - 1, axis:=axis, endnodes:=endnodes
    Next
    
End Function

Private Sub assignobject(ByRef dest, ByRef obj)
On Error GoTo e
    Set dest = obj
Exit Sub
e:
    dest = obj
End Sub

Private Sub assignvalue(ByRef dest, ByRef obj)
If IsObject(obj) Then assignobject dest, obj: Exit Sub
On Error GoTo e
    dest = obj
Exit Sub
e:
    Set dest = obj
End Sub


Public Function where(func As DogeFunc) As DogeList
    Dim obj
    Dim tmpdogelist As New DogeList
    For Each obj In Me.itemArray
        If func.exc(obj) Then tmpdogelist.append (obj)
    Next
    Set where = tmpdogelist
End Function

Public Function map(func As DogeFunc) As DogeList
    Dim obj
    Dim tmpdogelist As New DogeList
    For Each obj In Me.itemArray
        tmpdogelist.append (func.exc(obj))
    Next
    Set map = tmpdogelist
End Function

Public Function unique() As DogeList
Dim tmp As New Collection
Dim tmpdogelist As New DogeList
Dim obj
For Each obj In Me.itemArray
    If IsArray(obj) Then
        tmpdogelist.append (obj)
    ElseIf IsObject(obj) Then
        If tryAdd("<Doge>" & TypeName(obj) & CStr(ObjPtr(obj)), 1, tmp) Then
            tmpdogelist.append (obj)
        End If
    Else
        If tryAdd("<Doge>" & TypeName(obj) & CStr(obj), 1, tmp) Then
            tmpdogelist.append (obj)
        End If
    End If
Next
Set unique = tmpdogelist
End Function

Public Function join(delimiter, Optional text_qualifier As String = "", Optional newline_char As String = vbNewLine) As String
    join = ""
    Dim tmpstr As String
    Dim obj
    Dim tmp_delimiter
    For Each obj In itemArray
        tmp_delimiter = delimiter
        If IsObject(obj) Then
            If TypeOf obj Is DogeList Then
                tmpstr = obj.join(delimiter, text_qualifier, newline_char)
                tmp_delimiter = newline_char
            Else
                tmpstr = TypeName(obj) & "@" & CStr(ObjPtr(obj))
            End If
        ElseIf IsArray(obj) Then
            tmpstr = TypeName(obj) & "@" & CStr(VarPtr(obj))
        ElseIf TypeName(obj) = "String" Then
            tmpstr = text_qualifier & obj & text_qualifier
        Else
            tmpstr = CStr(obj)
        End If
        tmpstr = tmpstr & tmp_delimiter
        join = join & tmpstr
    Next
    If join = "" Then Exit Function
    join = Left(join, Len(join) - Len(tmp_delimiter))
End Function

Public Function tryAdd(key As String, item, ByRef col As Collection) As Boolean
On Error GoTo e1
    col.add item, key
    tryAdd = True
    Exit Function
e1:
End Function

Public Function lambdaMap(funcstring As String, ParamArray param()) As DogeList
    Dim func As New DogeFunc
    Dim p()
    p = param
    Set func = func.lambda(funcstring)
    func.default_ = p
    Dim obj
    Dim tmpdogelist As New DogeList
    For Each obj In Me.itemArray
        tmpdogelist.append (func.exc(obj))
    Next
    Set lambdaMap = tmpdogelist
End Function

Private Function isIterable(obj) As Boolean:
If TypeOf obj Is DogeList Then
    isIterable = True
End If
On Error GoTo isIterable_Error: Dim i: For Each i In obj: isIterable = True: Exit Function: Next:
isIterable_Error:
End Function
Private Function ismultidimensionarray(var) As Boolean
    If IsArray(var) Then
        If UBound(getDimensionality(var)) > 0 Then
            ismultidimensionarray = True
        Else
            ismultidimensionarray = False
        End If
    Else
        ismultidimensionarray = False
    End If
End Function

Private Function getDimensionality(var As Variant) As Variant
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    Dim bound
    Dim arr
    arr = Array()
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(var, i)
        bound = UBound(arr) + 1
        ReDim Preserve arr(bound)
        arr(bound) = tmp
    Loop
Err:
    getDimensionality = arr
End Function

Private Sub Class_Initialize()
    data_ = Array()
End Sub

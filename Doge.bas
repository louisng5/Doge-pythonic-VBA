Attribute VB_Name = "Doge"
Option Explicit
Public Function lambda(function_string As String, ParamArray param()) As DogeFunc
    Dim func As New DogeFunc
    Dim p()
    p = param
    Set func = func.lambda(function_string)
    func.default_ = p
    Set lambda = func
End Function

Public Function newfunc(funcname As String, ParamArray param()) As DogeFunc
    Dim func As New DogeFunc
    Dim p()
    p = param
    func.init funcname
    func.default_ = p
    Set newfunc = func
End Function

Public Function newlist(data) As DogeList
    Dim d As New DogeList
    d.assigndata data
    Set newlist = d
End Function

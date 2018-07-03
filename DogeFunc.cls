VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DogeFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''
' DogeFunc V1.0
' https://github.com/louisng5/Doge-pythonic-VBA
'
' Higher Order Function for VBA
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
Private funcname_
Private vbComp_ As Object
Private mutable_ As Boolean
Private defaultparam_()
#Const DEBUGMODE = False
#If DEBUGMODE Then
Private lambdaWB_  As Workbook
#End If

Public Sub init(funcname As String, ParamArray defaultparam())
    If Not mutable_ Then Err.Raise 9999, , "Function object already initialized"
    funcname_ = funcname
    defaultparam_ = defaultparam
    mutable_ = False
End Sub

Public Property Get funcname()
    funcname = funcname_
End Property

Private Sub assignvalue(ByRef dest, ByRef obj)
On Error GoTo e
    dest = obj
Exit Sub
e:
    Set dest = obj
End Sub

Public Property Let default_(param)
    defaultparam_ = param
End Property

Public Function exc(ParamArray param() As Variant)
    Dim p()
    If UBound(defaultparam_) = -1 Then
        p = param
    Else
        Dim i, assigncount
        i = 0
        assigncount = 0
        p = defaultparam_
        Do While assigncount < UBound(param) + 1
            If Not i > UBound(p) Then
                If IsMissing(p(i)) Then
                    assignvalue p(i), param(assigncount)
                    assigncount = assigncount + 1
                End If
            Else
                ReDim Preserve p(i)
                assignvalue p(i), param(assigncount)
                assigncount = assigncount + 1
            End If
            i = i + 1
        Loop
    End If
      Select Case UBound(p()) + 1
        Case 0
            getresult exc, Application.Run(funcname_)
        Case 1
            getresult exc, Application.Run(funcname_, p(0))
        Case 2
            getresult exc, Application.Run(funcname_, p(0), p(1))
        Case 3
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2))
        Case 4
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3))
        Case 5
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4))
        Case 6
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5))
        Case 7
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6))
        Case 8
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7))
        Case 9
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8))
        Case 10
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9))
        Case 11
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10))
        Case 12
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11))
        Case 13
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12))
        Case 14
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13))
        Case 15
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14))
        Case 16
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15))
        Case 17
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16))
        Case 18
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17))
        Case 19
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18))
        Case 20
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19))
        Case 21
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20))
        Case 22
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21))
        Case 23
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22))
        Case 24
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23))
        Case 25
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24))
        Case 26
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25))
        Case 27
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26))
        Case 28
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27))
        Case 29
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28))
        Case 30
            getresult exc, Application.Run(funcname_, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28), p(29))
    End Select
End Function

Private Sub getresult(ByRef returnvalue, ByVal result)
On Error GoTo e
    returnvalue = result
    Exit Sub
e:
    Set returnvalue = result
End Sub

Public Function lambda(ByVal funcstring As String, ParamArray defaultparam()) As DogeFunc
If Not mutable_ Then Err.Raise 9999, , "Function object already initialized"
funcstring = lambda_phaser(funcstring)
defaultparam_ = defaultparam
Dim ProcKind As VBIDE.vbext_ProcKind
#If DEBUGMODE Then
    Set lambdaWB_ = Workbooks.add
    Set vbComp_ = lambdaWB_.VBProject.VBComponents.add(1)
    vbComp_.CodeModule.AddFromString funcstring
    funcname_ = lambdaWB_.Name & "!" & vbComp_.Name & "." & vbComp_.CodeModule.ProcOfLine(1, ProcKind)
#Else
    Set vbComp_ = ThisWorkbook.VBProject.VBComponents.add(1)
    vbComp_.CodeModule.AddFromString funcstring
    funcname_ = vbComp_.Name & "." & vbComp_.CodeModule.ProcOfLine(1, ProcKind)
#End If
    mutable_ = False
    Set lambda = Me
End Function

Private Function lambda_phaser(funcstring As String) As String
    lambda_phaser = "Function " & funcstring & ":End Function"
End Function

Private Sub Class_Initialize()
    mutable_ = True
End Sub

Private Sub Class_Terminate()
    #If DEBUGMODE Then
    If Not lambdaWB_ Is Nothing Then
        lambdaWB_.Saved = True
        lambdaWB_.Close
    End If
    #Else
    If Not vbComp_ Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.remove vbComp_
    End If
    #End If

End Sub

Public Function map(iterable) As DogeList
    If mutable_ Then Err.Raise 9999, , "Function object have not been initialized"
    Dim d As New DogeList
    d.items = iterable
    Set map = d.map(Me)
End Function

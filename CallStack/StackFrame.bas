VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IStackFrame

Private Type TStackFrame
    ModuleName As String
    MemberName As String
    values As Collection
End Type

Private this As TStackFrame

Public Function Create(ByVal module As String, ByVal member As String, ByRef parameterValues() As Variant) As IStackFrame
Attribute Create.VB_Description = "Creates a new instance of an object representing a stack frame, i.e. a procedure call and its arguments."
    With New StackFrame
        .ModuleName = module
        .MemberName = member

        Dim index As Integer
        For index = LBound(parameterValues) To UBound(parameterValues)
            .AddParameterValue parameterValues(index)
        Next

        Set Create = .Self
    End With
End Function

Public Property Get Self() As IStackFrame
Attribute Self.VB_Description = "Gets a reference to this instance."
    Set Self = Me
End Property

Public Property Get ModuleName() As String
Attribute ModuleName.VB_Description = "Gets/sets the name of the module for this instance."
    ModuleName = this.ModuleName
End Property

Public Property Let ModuleName(ByVal value As String)
    this.ModuleName = value
End Property

Public Property Get MemberName() As String
Attribute ModuleName.VB_Description = "Gets/sets the name of the member for this instance."
    MemberName = this.MemberName
End Property

Public Property Let MemberName(ByVal value As String)
    this.MemberName = value
End Property

Public Property Get ParameterValue(ByVal index As Integer) As Variant
Attribute ModuleName.VB_Description = "Gets the value of the parameter at the specified index."
    ParameterValue = this.values(index)
End Property

Public Sub AddParameterValue(ByRef value As Variant)
Attribute AddParameterValue.VB_Description = "Adds the specified parameter value to this instance."
    this.values.Add value
End Sub

Private Sub Class_Initialize()
    Set this.values = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.values = Nothing
End Sub

Private Property Get IStackFrame_MemberName() As String
    IStackFrame_MemberName = this.MemberName
End Property

Private Property Get IStackFrame_ModuleName() As String
    IStackFrame_ModuleName = this.ModuleName
End Property

Private Property Get IStackFrame_ParameterValue(ByVal index As Integer) As Variant
    IStackFrame_ParameterValue = this.values(index)
End Property

Private Function IStackFrame_ToString() As String

    Dim result As String
    result = this.ModuleName & "." & this.MemberName & "("

    Dim index As Integer
    Dim value As Variant
    For Each value In this.values

        index = index + 1

        result = result & "{" & TypeName(value) & ":"
        If IsObject(value) Then
            result = result & ObjPtr(value)
        ElseIf IsArray(value) Then
            result = result & "[" & LBound(value) & "-" & UBound(value) & "]"
        ElseIf VarType(value) = vbString Then
            result = result & Chr$(34) & value & Chr$(34)
        Else
            result = result & CStr(value)
        End If
        result = result & "}" & IIf(index = this.values.Count, vbNullString, ",")

    Next

    result = result & ")"
    IStackFrame_ToString = result

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSpecExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mvarGivenValue As Variant

Private Function get_matcher_fail_msg(actual_value As String, expected_value As String, expectation_suffix As String, _
                                        Optional custom_err_msg, Optional found_msg) As String
    'fail msg contains 3 lines: headline, found value and expected value
    Dim fail_msg As String
    'value indention is the number of characters from the first character of value_type to the ">" sign,
    '   eg. "found: >400<" has an indention of 7
    Dim value_indentation As Long
    
    If Not IsMissing(custom_err_msg) Then
        custom_err_msg = Trim(CStr(custom_err_msg))
    Else
        custom_err_msg = vbNullString
    End If
    value_indentation = Len(Trim("expected: " & expectation_suffix & " ") & " ")
    'start building the fail msg
    fail_msg = Space(4) & "failed expectation" & vbLf
    If IsMissing(found_msg) Then
        fail_msg = fail_msg & Space(4) & "found:" & Space(value_indentation - Len("found:")) & ">" & actual_value & "<" & vbLf
    Else
        fail_msg = fail_msg & Space(4) & found_msg & vbLf
    End If
    If expectation_suffix = vbNullString Then
        fail_msg = fail_msg & Space(4) & "expected: " & ">" & expected_value & "<"
    Else
        fail_msg = fail_msg & Space(4) & "expected: " & expectation_suffix & " >" & expected_value & "<"
    End If
    If custom_err_msg <> vbNullString Then
        fail_msg = fail_msg & vbLf & Space(4) & custom_err_msg
    End If
    get_matcher_fail_msg = fail_msg
End Function

Private Function get_validator_fail_msg(validation_result As String, Optional custom_err_msg) As String
    'fail msg contains 2 lines: headline, validation result
    Dim fail_msg As String
    
    If Not IsMissing(custom_err_msg) Then
        custom_err_msg = Trim(CStr(custom_err_msg))
    Else
        custom_err_msg = vbNullString
    End If
    fail_msg = Space(4) & "failed expectation" & vbLf
    fail_msg = fail_msg & Space(4) & validation_result
    If custom_err_msg <> vbNullString Then
        fail_msg = fail_msg & vbLf & Space(4) & custom_err_msg
    End If
    get_validator_fail_msg = fail_msg
End Function

Public Sub to_be(expected_value, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If Not values_are_equal(Me.given_value, expected_value) Then
        matcher_fail_msg = get_matcher_fail_msg(Me.given_value, CStr(expected_value), vbNullString, custom_err_msg)
        raise_expectation_err matcher_fail_msg
    End If
End Sub

Public Sub not_to_be(expected_value, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If values_are_equal(Me.given_value, expected_value) Then
        matcher_fail_msg = get_matcher_fail_msg(Me.given_value, CStr(expected_value), "not", custom_err_msg)
        raise_expectation_err matcher_fail_msg
    End If
End Sub

Public Sub to_be_nothing(Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If Not IsObject(Me.given_value) Then
        matcher_fail_msg = get_matcher_fail_msg(Me.given_value, "Nothing", vbNullString, custom_err_msg)
        raise_expectation_err matcher_fail_msg
    Else
        If Not Me.given_value Is Nothing Then
            matcher_fail_msg = get_matcher_fail_msg(vbNullString, "Nothing", vbNullString, custom_err_msg, "found an object")
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Sub not_to_be_nothing(Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If Not IsObject(Me.given_value) Then
        matcher_fail_msg = get_matcher_fail_msg(Me.given_value, "an object", vbNullString, custom_err_msg)
        raise_expectation_err matcher_fail_msg
    Else
        If Me.given_value Is Nothing Then
            matcher_fail_msg = get_matcher_fail_msg("Nothing", "an object", vbNullString, custom_err_msg)
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Private Sub test_comparability(pGivenValue As Variant, pExpectedValue As Variant)
    Dim matcher_fail_msg As String
    
    If (IsObject(Me.given_value) And Not IsObject(pExpectedValue)) _
     Or (IsObject(pExpectedValue) And Not IsObject(Me.given_value)) Then
        matcher_fail_msg = vbTab & "can't compare given with expected value - check if only one of them is an object"
        raise_expectation_err matcher_fail_msg
    End If
End Sub

Private Function values_are_equal(given_value As Variant, expected_value As Variant) As Boolean
    Dim index As Integer
    
    'basic data types (e.g. string, integer) can't be compared with objects
    test_comparability given_value, expected_value
    
    values_are_equal = True
    If IsObject(given_value) Then
        If Not (given_value Is expected_value) Then
            values_are_equal = False
        End If
    ElseIf IsArray(given_value) Then
        If UBound(given_value) <> UBound(expected_value) Then
            'if two arrays have a different size, they can't be equal
            values_are_equal = False
        Else
            For index = 0 To UBound(given_value)
                If given_value(index) <> expected_value(index) Then
                   values_are_equal = False
                   Exit For
                End If
            Next
        End If
    Else
        If given_value <> expected_value Then
            values_are_equal = False
        End If
    End If
End Function

Public Property Get given_value() As Variant
    On Error Resume Next
    Set given_value = mvarGivenValue
    given_value = mvarGivenValue
    On Error GoTo 0
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    On Error Resume Next
    Set mvarGivenValue = pvarGivenValue
    mvarGivenValue = pvarGivenValue
    On Error GoTo 0
End Property

Public Sub contains_member(expected_member As Variant, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    Dim expected_member_text As String

    If TypeName(Me.given_value) = "Collection" Then
        If Not ExtraVBA.collection_has_value(expected_member, Me.given_value) Then
            If IsArray(expected_member) Then
                expected_member_text = Join(expected_member, "*")
            ElseIf IsObject(expected_member) Then
                expected_member_text = "given " & TypeName(expected_member)
            Else
                expected_member_text = CStr(expected_member)
            End If
            matcher_fail_msg = get_validator_fail_msg("collection does not contain item " & ">" & expected_member_text & "<", custom_err_msg)
            raise_expectation_err matcher_fail_msg
        End If
    ElseIf IsArray(Me.given_value) Then
        If Not ExtraVBA.array_has_value(expected_member, Me.given_value) Then
            expected_member_text = CStr(expected_member)
            matcher_fail_msg = get_validator_fail_msg("array does not contain item " & ">" & expected_member_text & "<", custom_err_msg)
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Sub includes_text(include_item As Variant, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If TypeName(include_item) = "String" Then
        If InStr(Me.given_value, include_item) = 0 Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should include:" & vbTab & ">" & include_item & "<"
            If Not IsMissing(custom_err_msg) Then
                matcher_fail_msg = matcher_fail_msg & vbLf & vbTab & custom_err_msg
            End If
            raise_expectation_err matcher_fail_msg
        End If
    Else
        Debug.Print "TSpecExpectation.includes_text PENDING: include test for arrays"
    End If
End Sub

Public Sub does_not_include(include_item As Variant)
    Dim matcher_fail_msg As String
    
    If TypeName(include_item) = "String" Then
        If InStr(Me.given_value, include_item) > 0 Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should not include:" & vbTab & ">" & include_item & "<"
            raise_expectation_err matcher_fail_msg
        End If
    Else
        Debug.Print "STATUS_PENDING: include test for arrays"
    End If
End Sub

Public Sub starts_with_text(heading_text As Variant, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If TypeName(heading_text) = "String" Then
        If Left(Me.given_value, Len(heading_text)) <> heading_text Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should start with:" & vbTab & ">" & heading_text & "<"
            If Not IsMissing(custom_err_msg) Then
                matcher_fail_msg = matcher_fail_msg & vbLf & vbTab & custom_err_msg
            End If
            raise_expectation_err matcher_fail_msg
        End If
    Else
        raise_expectation_err "invalid input for expectation method starts_with_text:" & vbLf & vbTab & _
                                "starts_with_text does apply only for Strings and not for " & TypeName(heading_text)
    End If
End Sub

Public Sub ends_with_text(trailing_text As Variant, Optional custom_err_msg)
    Dim matcher_fail_msg As String
    
    If TypeName(trailing_text) = "String" Then
        If Right(Me.given_value, Len(trailing_text)) <> trailing_text Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should end with:" & vbTab & ">" & trailing_text & "<"
            If Not IsMissing(custom_err_msg) Then
                matcher_fail_msg = matcher_fail_msg & vbLf & vbTab & custom_err_msg
            End If
            raise_expectation_err matcher_fail_msg
        End If
    Else
        raise_expectation_err "invalid input for expectation method ends_with_text:" & vbLf & vbTab & _
                                "ends_with_text does apply only for Strings and not for " & TypeName(trailing_text)
    End If
End Sub

Private Sub raise_expectation_err(fail_msg)
    TSpec.LastFailMsg = fail_msg
    Err.raise Number:=ERR_ID_EXPECTATION_STATUS_FAILED, description:=fail_msg
End Sub


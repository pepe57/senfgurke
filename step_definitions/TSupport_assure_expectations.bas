Attribute VB_Name = "TSupport_assure_expectations"
Option Explicit

Private Function get_failure_msg()
    Dim err_msg As String
    
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    err_msg = TSpec.LastFailMsg
    get_failure_msg = normalize_err_msg(err_msg)
End Function

Public Function normalize_err_msg(err_msg As String) As String
    Dim msg_text As String
    
    'resolve any tabs into fixed width white space
    msg_text = Replace(err_msg, vbTab, Space(2))
    'remove any blockindention used in feature files
    msg_text = Senfgurke.ExtraVBA.align_textblock(msg_text)
    'escape linebreaks to fit the actual err msg into one line
    msg_text = Replace(msg_text, vbLf, "\n")
    normalize_err_msg = msg_text
End Function

Public Function run_search_expectation(search_list As Variant, search_item As Variant, custom_err_msg As String, _
                                        search_type As String, example_context As TContext) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo search_failed
    Select Case search_type
        Case "contains_member"
            TSpec.expect(search_list).contains_member search_item, custom_err_msg
    End Select
    example_context.set_value "confirmed", "expectation_result"
resume_after_failure:
    On Error GoTo 0
    run_search_expectation = actual_err_msg
    Exit Function
    
search_failed:
    actual_err_msg = get_failure_msg
    example_context.set_value "failed", "expectation_result"
    Err.Clear
    Resume resume_after_failure
End Function

Public Function run_comparison_expectation(actual_value As Variant, expected_value As Variant, custom_err_msg As String, _
                                            comparison_type As String, example_context As TContext) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo comparison_failed
    Select Case comparison_type
        Case "starts_with"
            TSpec.expect(actual_value).starts_with_text expected_value, custom_err_msg
        Case "ends_with"
            TSpec.expect(actual_value).ends_with_text expected_value, custom_err_msg
        Case "includes_text"
            TSpec.expect(actual_value).includes_text expected_value, custom_err_msg
        Case "to_be"
            TSpec.expect(actual_value).to_be expected_value, custom_err_msg
        Case "not_to_be"
            TSpec.expect(actual_value).not_to_be expected_value, custom_err_msg
    End Select
    example_context.set_value "confirmed", "expectation_result"
resume_after_failure:
    On Error GoTo 0
    run_comparison_expectation = actual_err_msg
    Exit Function
    
comparison_failed:
    actual_err_msg = get_failure_msg
    example_context.set_value "failed", "expectation_result"
    Err.Clear
    Resume resume_after_failure
End Function

Public Function run_validation_expectation(actual_value As Variant, comparison_type As String, example_context As TContext, _
                                            Optional custom_err_msg) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo comparison_failed
    Select Case comparison_type
        Case "to_be_nothing"
            TSpec.expect(actual_value).to_be_nothing
        Case "not_to_be_nothing"
            TSpec.expect(actual_value).not_to_be_nothing
    End Select
    example_context.set_value "confirmed", "expectation_result"
resume_after_failure:
    On Error GoTo 0
    run_validation_expectation = actual_err_msg
    Exit Function
    
comparison_failed:
    actual_err_msg = get_failure_msg
    example_context.set_value "failed", "expectation_result"
    Err.Clear
    Resume resume_after_failure
End Function

Public Function run_collection_expectation(given_collection As Collection, expected_value As Variant, validation_type, _
                                            example_context As TContext, Optional custom_err_msg) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo comparison_failed
    Select Case validation_type
        Case "contains_member"
            TSpec.expect(given_collection).contains_member expected_value, custom_err_msg
    End Select
    example_context.set_value "confirmed", "expectation_result"
resume_after_failure:
    On Error GoTo 0
    run_collection_expectation = actual_err_msg
    Exit Function
    
comparison_failed:
    actual_err_msg = get_failure_msg
    example_context.set_value "failed", "expectation_result"
    Err.Clear
    Resume resume_after_failure
End Function

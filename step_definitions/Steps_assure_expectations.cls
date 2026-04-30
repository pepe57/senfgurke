VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_assure_expectations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_expected_value_was_defined_as_INT_AA730CC5F5BC(example_context As TContext, step_expressions As Collection)
    'Given an expected value was defined as {integer}
    example_context.set_value step_expressions(1), "expected_value"
End Sub

Public Sub Given_the_actual_value_was_evaluated_as_INT_C1726D898F77(example_context As TContext, step_expressions As Collection)
    'And the actual value was evaluated as {integer}
    example_context.set_value step_expressions(1), "actual_value"
End Sub

Public Sub Given_the_actual_value_refers_to_an_object_C17237078D18(example_context As TContext)
    'Given the actual value refers to an object
    example_context.set_value New Collection, "actual_value"
End Sub

Public Sub Given_the_actual_value_refers_to_nothing_F37725446C7A(example_context As TContext)
    'Given the actual value refers to Nothing
    example_context.set_value Nothing, "actual_value"
End Sub

Public Sub When_expected_and_actual_values_are_being_com_04DE16907D42(example_context As TContext)
    'When expected and actual values are being compared using to_be
    Dim actual_err_msg As String
    
    actual_err_msg = run_comparison_expectation("to_be", example_context)
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_expected_and_actual_values_are_being_com_84BE37220FDB(example_context As TContext)
    'When expected and actual values are being compared using not_to_be
    Dim actual_err_msg As String
    
    actual_err_msg = run_comparison_expectation("not_to_be", example_context)
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_nonobject_variable_is_expected_to_be_10954F3DA4DC(example_context As TContext)
    'When the non-object variable is expected to be Nothing
    On Error GoTo expectation_failed
    TSpec.expect(example_context.get_value("sample_variable")).to_be_nothing
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

Public Sub When_the_actual_value_is_validated_using_to_b_CC83E48E0542(example_context As TContext)
    'When the actual value is validated using to_be_nothing
    Dim actual_value As Variant
    Dim actual_err_msg As String
    
    If IsObject(example_context.get_value("actual_value")) Then
        Set actual_value = example_context.get_value("actual_value")
    Else
        actual_value = example_context.get_value("actual_value")
    End If
    actual_err_msg = run_validation_expectation(actual_value, "to_be_nothing", example_context)
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_actual_value_is_validated_using_not__8C8DAF193434(example_context As TContext)
    'When the actual value is validated using not_to_be_nothing
    Dim actual_value As Variant
    Dim actual_err_msg As String
    
    If IsObject(example_context.get_value("actual_value")) Then
        Set actual_value = example_context.get_value("actual_value")
    Else
        actual_value = example_context.get_value("actual_value")
    End If
    actual_err_msg = run_validation_expectation(actual_value, "not_to_be_nothing", example_context)
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub Then_the_expectation_is_confirmed_A341C2A69D10(example_context As TContext)
    'Then the expectation is confirmed
    Dim expectation_result As String
    
    expectation_result = example_context.get_value("expectation_result")
    TSpec.expect(expectation_result).to_be "confirmed"
End Sub

Public Sub Then_the_expectation_fails_DBF0F873A79F(example_context As TContext)
    'Then the expectation fails
    Dim expectation_result As String
    
    expectation_result = example_context.get_value("expectation_result")
    TSpec.expect(expectation_result).to_be "failed"
End Sub

Public Sub Then_the_expectation_result_is_shown_as_STR_4642F94D38D6(example_context As TContext, step_expressions As Collection)
    'And the expection result is shown as {string}
    Dim actual_err_msg As String
    Dim expected_err_msg As String
    
    'ignore tabs in msg comparison
    actual_err_msg = example_context.get_value("actual_err_msg")
    expected_err_msg = TSupport_assure_expectations.normalize_err_msg(CStr(step_expressions(1)))
    TSpec.expect(actual_err_msg).to_be expected_err_msg
End Sub

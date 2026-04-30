VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_assure_numeric_expecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_expected_value_is_of_type_int_FBDDC5B46F7B(example_context As TContext)
    'Given an expected value is of type int
    example_context.set_value CInt(11), "expected_value"
End Sub

Public Sub Given_an_expected_value_is_of_type_double_72FB6915AB87(example_context As TContext)
    'Given an expected value is of type double
    example_context.set_value CDbl(0.5), "expected_value"
End Sub

Public Sub Given_an_expected_value_is_of_type_string_72FB69159910(example_context As TContext)
    'Given an expected value is of type string
    example_context.set_value CStr("a text"), "expected_value"
End Sub

Public Sub Given_an_expected_value_is_of_type_collection_C2EB8752BD54(example_context As TContext)
    'Given an expected value is of type collection
    example_context.set_value New Collection, "expected_value"
End Sub

Public Sub Given_an_expected_value_is_of_type_array_72FB69155444(example_context As TContext)
    'Given an expected value is of type array
    example_context.set_value Array(1, 2, 3), "expected_value"
End Sub

Public Sub Given_the_actual_value_is_of_type_string_F37750FA1FE4(example_context As TContext)
    'And the actual value is of type string
    example_context.set_value CStr("a text"), "actual_value"
End Sub

Public Sub Given_the_actual_value_is_of_type_boolean_F37750FAE102(example_context As TContext)
    'And the actual value is of type boolean
    example_context.set_value True, "actual_value"
End Sub

Public Sub Given_the_actual_value_is_of_type_collection_C172DFB13AD3(example_context As TContext)
    'And the actual value is of type collection
    example_context.set_value New Collection, "actual_value"
End Sub

Public Sub Given_the_actual_value_is_of_type_array_F37750FAE1B3(example_context As TContext)
    'And the actual value is of type array
    example_context.set_value Array(1, 2, 3), "actual_value"
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_84BE59E60CD0(example_context As TContext)
    'When expected and actual value are being compared using to_be with
    run_comparison_expectation "to_be", example_context
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_98C557A814BE(example_context As TContext)
    'When expected and actual value are being compared using not_to_be with
    run_comparison_expectation "not_to_be", example_context
End Sub

Public Sub When_actual_and_expected_values_are_being_com_667181D5805B(example_context As TContext, step_expressions As Collection)
    'When actual and expected values are being compared using to_be with {string}
    run_comparison_expectation "to_be", example_context, step_expressions(1)
End Sub

Public Sub When_actual_and_expected_values_are_being_com_82CE18F68AB7(example_context As TContext, step_expressions As Collection)
    'When actual and expected values are being compared using not_to_be with {string}
    run_comparison_expectation "not_to_be", example_context, step_expressions(1)
End Sub

Public Sub Then_the_comparison_results_in_an_error_messa_B924C73DD4AE(example_context As TContext)
    'Then the comparison results in an error message
    Dim actual_err_msg
    
    actual_err_msg = example_context.get_value("actual_err_msg")
    expect(Len(CStr(actual_err_msg))).not_to_be 0
    expect(CStr(actual_err_msg)).not_to_be "Expected an error being raised by a missed expectation but none found so far..."
End Sub

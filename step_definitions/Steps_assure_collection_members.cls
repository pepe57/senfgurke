VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_assure_collection_members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_collection_has_INT_members_array_INT__22433998BB74(example_context As TContext, step_expressions As Collection)
    'Given a collection has {integer} members array( {integer}, {string} ) and array( {integer}, {string} )
    Dim index As Integer
    Dim given_collection As Collection

    Set given_collection = New Collection
    example_context.set_value given_collection, "given_collection"
    For index = 1 To step_expressions(1) * 2 Step 2
        given_collection.Add Array(step_expressions(index + 1), step_expressions(index + 2))
    Next
End Sub

Public Sub Given_a_collection_has_INT_members_STR_and_ST_20E2C2F8B116(example_context As TContext, step_expressions As Collection)
    'Given a collection has {integer} members {string} and {string}
    Dim index As Integer
    Dim given_collection As Collection

    Set given_collection = New Collection
    example_context.set_value given_collection, "given_collection"
    For index = 2 To step_expressions(1) + 1
        given_collection.Add step_expressions(index)
    Next
End Sub

Public Sub Given_an_array_INT__STR__as_possible_collecti_000442533110(example_context As TContext, step_expressions As Collection)
    'And an array( {integer}  {string} ) as possible collection member
    Dim expected_member As Variant
    
    expected_member = Array(step_expressions(1), step_expressions(2))
    example_context.set_value expected_member, "expected_member"
End Sub

Public Sub When_collection_is_validated_for_the_specific_2B4B7B00B2EA(example_context As TContext)
    'When collection is validated for the specific member using contains_member
    Dim given_collection As Collection
    Dim expected_member As Variant
    Dim actual_err_msg As String
    
    Set given_collection = example_context.get_value("given_collection")
    If IsObject(example_context.get_value("expected_member")) Then
        Set expected_member = example_context.get_value("expected_member")
    Else
        expected_member = example_context.get_value("expected_member")
    End If
    actual_err_msg = run_collection_expectation(given_collection, expected_member, "contains_member", example_context)
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Ability: assure expectations
  Expectations are bridiging the gap between expected results described by
  all the examples in feature files and the actual results. So when Senfgurke
  provides an empty code template matching each step in an exmple - a developer
  will use expectations to match expected with actual values. This could look
  like this:
    TSpec.expect( variable_with_expected_value ).to_be variable_with_actual_value


  Rule: show actual and expected values if an expectation fails
    Example: compare integer type values with to_be comparison
      Given an expected value was defined as 200
        And the actual value was evaluated as 400
       When expected and actual value are being compared using to_be
       Then the expectation fails
        And the expectation result is shown as
					"""
            failed expectation
            found:    >400<
            expected: >200<
          """

    Example: compare integer type values with not_to_be comparison
      Given an expected value was defined as 200
        And the actual value was evaluated as 200
       When expected and actual value are being compared using not_to_be
       Then the expectation fails
        And the expectation result is shown as
					"""
            failed expectation
            found:        >200<
            expected: not >200<
          """

@vba-specific
Ability: assure objects exists
    VBA distinguishes between variables assigned to objects and variables
    assigned to basic data types (e.g. string, integer, boolean). This
    expectation assures a variable is assigned to an object or not. The
    expectation is using the Nothing keyword to determine if a variable is
    assigned to an objoct or not.


  Rule: is Nothing expectation should fail for variables assigned to an object
    Example: expected Nothing when object exists
      Given the actual value refers to an object
       When the actual value is validated using to_be_nothing
       Then the expectation fails
        And the expectation result is shown as
          """
            failed expectation
            found an object
            expected: >Nothing<
          """

    Example: expected Nothing when object doesn't exists
      Given the actual value refers to Nothing
       When the actual value is validated using to_be_nothing
       Then the expectation is confirmed

    Example: expected Nothing when the actual value isn't an object
      Given the actual value was evaluated as 42
       When the actual value is validated using to_be_nothing
       Then the expectation fails
        And the expectation result is shown as
           """
             failed expectation
             found:    >42<
             expected: >Nothing<
           """

  Rule: not is Nothing expectation should confirm when a variable is assigned to an object
    Example: expected not Nothing when object exists
      Given the actual value refers to an object
       When the actual value is validated using not_to_be_nothing
       Then the expectation is confirmed

    Example: expected not Nothing when object doesn't exists
      Given the actual value refers to Nothing
       When the actual value is validated using not_to_be_nothing
       Then the expectation fails
        And the expectation result is shown as
          """
            failed expectation
            found:    >Nothing<
            expected: >an object<
          """

    Example: expected Nothing when the actual value isn't an object
      Given the actual value was evaluated as 42
       When the actual value is validated using not_to_be_nothing
       Then the expectation fails
        And the expectation result is shown as
          """
            failed expectation
            found:    >42<
            expected: >an object<
          """

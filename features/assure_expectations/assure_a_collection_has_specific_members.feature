@vba-specific
Ability: assure a collection has specific members
    A collection in VBA is like a named array (see
    https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
    for more information). This expectation validates if a collection contains
    certain members or not.

  Rule: contains expectation should confirm that an array is an item of a collection
    Note: VBA knows fixed sized arrays as well. Those arrays can be added as
    items to collections.

    Example: collection contains an array of primitive values
      Given a collection has 2 members array(1,"a") and array(2,"b")
        And an array(2,"b") as possible collection member
       When collection is validated for the specific member using contains_member
       Then the expectation is confirmed

    Example: collections contains arrays but not the expected one
      Given a collection has 2 members array(1,"a") and array(2,"b")
        And an array(3,"c") as possible collection member
       When collection is validated for the specific member using contains_member
       Then the expectation fails
        And the expectation result is shown as
             """
               failed expectation
               collection does not contain item >3*c<
             """

    Example: collections doesn't contain any array
      Given a collection has 2 members "a" and "b"
        And an array(1,"a") as possible collection member
       When collection is validated for the specific member using contains_member
       Then the expectation fails
        And the expectation result is shown as
              """
                failed expectation
                collection does not contain item >1*a<
              """

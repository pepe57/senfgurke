@vba
Ability: compare numeric values
  VBA distinguishes between variables assigned to objects and variables
  assigned to primitive data types (e.g. string, integer, boolean). This
  expectation assures that two basic values have the correct relation to each
  other (eg. one is greater, equal or less than the other).


  Rule: the result of a value comparison expectation is determined by the comparison type

    Example: compare integer type values
      Given an expected value was defined as <expected_value>
        And the actual value was evaluated as <actual_value>
       When actual and expected values are being compared using <expectation_function> with <comparison_type>
       Then the expectation <result>

#      Examples: equal comparison as default
#        | expectation_function | comparison_type |expected_value | actual_value | result       |
#        | to_be                |                 |           200 |          400 | fails        |
#        | not_to_be            |                 |           200 |          400 | is confirmed |

      Examples: explicit comparison
        | expectation_function | actual_value | comparison_type | expected_value |    result    |
        | to_be                |          200 | "="             |            400 | fails        |
        | not_to_be            |          200 | "="             |            400 | is confirmed |
        | to_be                |          400 | ">"             |            200 | is confirmed |
        | not_to_be            |          400 | ">"             |            200 | fails        |
        | to_be                |          100 | "<"             |            200 | is confirmed |
        | not_to_be            |          100 | "<"             |            200 | fails        |


  Rule: return error message when comparing values with non-comparable data types
  # VBA considers boolean as an numeric data type where 0 represents false and
  # any other value represents true while -1 is the default value for true

    Example: compare primitive values with incompatible data types
      Given an expected value is of type <type_expected>
        And the actual value is of type <type_actual>
       When expected and actual values are being compared using <expectation_function>
       Then the comparison results in an error message

       Examples: integer and long values
        | expectation_function | type_expected | type_actual |
        | to_be                |           int |      string |
        | to_be                |        double |      string |
        | not_to_be            |           int |      string |
        | not_to_be            |        string |     boolean |

       Examples: objects and non-objects
        | expectation_function | type_expected | type_actual |
        | to_be                |    collection |      string |
        | not_to_be            |           int |  collection |
        | not_to_be            |         array |  collection |

       Examples: arrays and primitives
        | expectation_function | type_expected | type_actual |
        | to_be                |         array |      string |
        | not_to_be            |           int |       array |

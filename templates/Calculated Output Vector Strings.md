# Calculated Output Vector Strings

The following input strings are for calculating new vectors in either the Femap _Model_ > _Output_ > _Calculate_ window or by using the `feOutputCalculate` method in the Femap API. All strings assume the default `i` and `case` variables.

> Do _not_ input any spaces in the strings when submitting them to Femap. Femap does not recognize spaces in these function strings.

Bush Resultant Beam Shear
    `SQRT(SQR(VEC(!case;3775;!i))+SQR(VEC(!case;3376;!i)))`
Bush Resultant Beam Bending Moment
    `SQRT(SQR(VEC(!case;3778;!i))+SQR(VEC(!case;3379;!i)))`

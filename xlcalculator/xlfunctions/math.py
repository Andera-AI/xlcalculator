import decimal
import math
from typing import Tuple, Union

import numpy as np
import pandas as pd

from . import xl, xlerrors, xlcriteria, func_xltypes


@xl.register()
@xl.validate_args
def ABS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Find the absolute value of provided value.

    https://support.office.com/en-us/article/
        abs-function-3420200f-5628-4e8c-99da-c99d7c87713c
    """
    return abs(number)


@xl.register()
@xl.validate_args
def ACOS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arccosine, or inverse cosine, of a number.

    https://support.office.com/en-us/article/
        acos-function-cb73173f-d089-4582-afa1-76e5524b5d5b
    """
    return np.arccos(float(number))


@xl.register()
@xl.validate_args
def ACOSH(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the inverse hyperbolic cosine of a number.

    https://support.office.com/en-us/article/
        acosh-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe
    """
    if number < 1:
        raise xlerrors.NameExcelError(f'number {number} must be greater'
                                      f'than or equal to 1')

    return np.arccosh(float(number))


@xl.register()
@xl.validate_args
def ASIN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arcsine, or inverse sine, of a number.

    https://support.office.com/en-us/article/
        asin-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347
    """
    if number < -1 or number > 1:
        raise xlerrors.NumExcelError(f'number {number} must be less than '
                                     f'or equal to -1 or greater ot equal '
                                     f'to 1')

    return np.arcsin(float(number))


@xl.register()
@xl.validate_args
def ASINH(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the inverse hyperbolic sine of a number.

    https://support.office.com/en-us/article/
        asinh-function-4e00475a-067a-43cf-926a-765b0249717c
    """
    return np.arcsinh(float(number))


@xl.register()
@xl.validate_args
def ATAN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arctangent, or inverse tangent, of a number.

    https://support.office.com/en-us/article/
        atan-function-50746fa8-630a-406b-81d0-4a2aed395543
    """
    return np.arctan(float(number))


@xl.register()
@xl.validate_args
def ATAN2(
        x_num: func_xltypes.XlNumber,
        y_num: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arctangent, or inverse tangent, of the specified
        x- and y-coordinates.

    https://support.office.com/en-us/article/
        atan2-function-c04592ab-b9e3-4908-b428-c96b3a565033
    """
    return np.arctan2(float(x_num), float(y_num))


@xl.register()
@xl.validate_args
def CEILING(
        number: func_xltypes.XlNumber,
        significance: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns number rounded up, away from zero, to the nearest multiple of
        significance.

    https://support.office.com/en-us/article/
        ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f
    """

    if significance == 0:
        return 0

    if significance < 0 < number:
        raise xlerrors.NumExcelError('significance below zero and number \
                                      above zero is not allowed')

    number = float(number)
    significance = float(significance)

    ceiling = significance * math.ceil(number / significance)

    # If number is an exact multiple of significance, no rounding occurs
    if (number % significance) == 0:
        return ceiling

    quantize_multiplier = str(significance % 1)

    # If number is negative, and significance is negative, the value is
    # rounded down, away from zero.
    if number < 0 and significance < 0:
        result = decimal.Decimal(ceiling)
        result = result.quantize(decimal.Decimal(quantize_multiplier),
                                 rounding=decimal.ROUND_DOWN)
        return float(result)

    # If number is negative, and significance is positive, the value is
    # rounded up towards zero.
    if number < 0 < significance:
        result = decimal.Decimal(ceiling)
        result = result.quantize(decimal.Decimal(quantize_multiplier),
                                 rounding=decimal.ROUND_UP)
        return float(result)

    # Regardless of the sign of number, a value is rounded up when adjusted
    # away from zero.
    result = decimal.Decimal(ceiling)
    result = result.quantize(decimal.Decimal(quantize_multiplier),
                             rounding=decimal.ROUND_UP)
    return float(result)


@xl.register()
@xl.validate_args
def COS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the cosine of the given angle.

    https://support.office.com/en-us/article/
        cos-function-0fb808a5-95d6-4553-8148-22aebdce5f05
    """
    return np.cos(float(number))


@xl.register()
@xl.validate_args
def DEGREES(
        angle: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Converts radians into degrees.

    https://support.office.com/en-us/article/
        degrees-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1
    """
    return np.degrees(float(angle))


@xl.register()
@xl.validate_args
def FLOOR(
        number: func_xltypes.XlNumber,
        significance: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Rounds number down, toward zero, to the nearest multiple of
        significance.

    https://support.office.com/en-us/article/
        FLOOR-function-14BB497C-24F2-4E04-B327-B0B4DE5A8886
    """

    if significance < 0 < number:
        raise xlerrors.NumExcelError('number and significance needto have \
                                      the same sign')

    if number == 0:
        return 0

    if significance == 0:
        raise xlerrors.DivZeroExcelError()

    return significance * math.floor(number / significance)


@xl.register()
@xl.validate_args
def LN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the natural logarithm of a number.

    https://support.office.com/en-us/article/
        ln-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f
    """
    return math.log(number)


@xl.register()
@xl.validate_args
def MOD(
        number: func_xltypes.XlNumber,
        divisor: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the remainder after number is divided by divisor.

    https://support.office.com/en-us/article/
        mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    """
    return number % divisor


@xl.register()
def PI() -> func_xltypes.XlNumber:
    """Returns the number 3.14159265358979, the mathematical constant pi.

    Accurate to 15 digits.

    https://support.office.com/en-us/article/
        pi-function-264199d0-a3ba-46b8-975a-c4a04608989b
    """
    return math.pi


@xl.register()
@xl.validate_args
def POWER(
        number: func_xltypes.XlNumber,
        power: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the result of a number raised to a power.

    https://support.office.com/en-us/article/
        power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a
    """
    return np.power(number, power)


@xl.register()
@xl.validate_args
def RADIANS(
        angle: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Converts degrees to radians.

    https://support.office.com/en-us/article/
        radians-function-ac409508-3d48-45f5-ac02-1497c92de5bf
    """
    return np.radians(float(angle))


@xl.register()
@xl.validate_args
def ROUND(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0,
        _rounding=decimal.ROUND_HALF_UP
):
    """Rounding half up

    https://support.office.com/en-us/article/
        ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c
    """
    number = decimal.Decimal(str(number))
    dc = decimal.getcontext()
    dc.rounding = _rounding
    ans = round(number, int(num_digits))
    return float(ans)


@xl.register()
@xl.validate_args
def ROUNDUP(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Round up

    https://support.office.com/en-us/article/
         ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7
    """
    return ROUND(number, num_digits=num_digits, _rounding=decimal.ROUND_UP)


@xl.register()
@xl.validate_args
def ROUNDDOWN(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Round down

    https://support.office.com/en-us/article/
        rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53
    """
    return ROUND(number, num_digits=num_digits, _rounding=decimal.ROUND_DOWN)


@xl.register()
@xl.validate_args
def SQRT(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns a positive square root.

    https://support.office.com/en-us/article/
        sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf
    """
    if number < 0:
        raise xlerrors.NumExcelError(f'number {number} must be non-negative')

    return math.sqrt(number)


@xl.register()
@xl.validate_args
def SUM(
        *numbers: Tuple[func_xltypes.XlNumber]
) -> func_xltypes.XlNumber:
    """The SUM function adds values.

    https://support.office.com/en-us/article/
        sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89
    """
    # If no non numeric cells, return zero (is what excel does)
    if len(numbers) == 0:
        return 0

    return sum(numbers)


@xl.register()
@xl.validate_args
def SUMIF(
        range: func_xltypes.XlArray,
        criteria: func_xltypes.XlAnything,
        sum_range: func_xltypes.XlArray = None
) -> func_xltypes.XlNumber:
    """Adds the cells specified by a given criteria.

    https://support.office.com/en-us/article/
        sumif-function-169b8c99-c05c-4483-a712-1697a653039b
    """
    # WARNING:
    # - wildcards not supported

    check = xlcriteria.parse_criteria(criteria)

    if sum_range is None:
        sum_range = range

    range = range.flat
    sum_range = sum_range.cast_to_numbers().flat

    # zip() will automatically drop any range values that have indexes larger
    # than sum_range's length.
    return sum([
        sval
        for cval, sval in zip(range, sum_range)
        if check(cval)
    ])


@xl.register()
@xl.validate_args
def SUMIFS(
        sum_range: func_xltypes.XlArray,
        criteria_range: func_xltypes.XlArray,
        criteria: func_xltypes.XlAnything,
        *criteriaAndRanges: Tuple[Union[
            func_xltypes.XlAnything, func_xltypes.XlArray
        ]]
) -> func_xltypes.XlNumber:
    """Adds the cells specified by given criteria in multiple arrays.
    Requires equal length arrays, since the arrays get decomposed.

    https://support.microsoft.com/en-us/office
        /sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b
    """
    # WARNING:
    # - wildcards not supported
    ranges = [criteria_range.flat]
    checks = [xlcriteria.parse_criteria(criteria)]
    rangeLen = len(criteria_range.flat)
    newRange = []
    idx = 0
    for item in criteriaAndRanges:
        if idx == rangeLen:
            checks.append(xlcriteria.parse_criteria(item))
            ranges.append(newRange)
            newRange = []
            idx = 0
        else:
            newRange.append(item)
            idx += 1
    sum_range = sum_range.cast_to_numbers().flat
    # zip() will automatically drop any range values that have indexes larger
    # than sum_range's length.
    return sum([
        sval
        for cvals, sval in zip(zip(*ranges), sum_range)
        if all(checkfn(cvals[i]) for i, checkfn in enumerate(checks))
    ])


@xl.register()
@xl.validate_args
def SUMPRODUCT(
        *arrays: Tuple[func_xltypes.XlArray]
) -> func_xltypes.XlNumber:
    """Returns the sum of the products of corresponding arrays or arrays.

    https://support.office.com/en-us/article/
        sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e
    """
    if len(arrays) == 0:
        raise xlerrors.NullExcelError('Not enough arguments for function.')

    array1_shape = arrays[0].shape
    if array1_shape == (0, 0):
        return 0

    for array in arrays:
        array_shape = array.shape
        if array1_shape != array_shape:
            raise xlerrors.ValueExcelError(
                f"The shapes of the arrays do not match. Looking "
                f"for {array1_shape} but given array has {array_shape}")
        if any(filter(xlerrors.ExcelError.is_error, xl.flatten(array))):
            raise xlerrors.NaExcelError(
                "Excel Errors are present in the sumproduct items.")

    sumproduct = pd.concat(arrays, axis=1)
    return sumproduct.prod(axis=1).sum()


@xl.register()
@xl.validate_args
def TRUNC(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Truncate a number to the specified number of digits.

    https://support.office.com/en-us/article/
        trunc-function-8b86a64c-3127-43db-ba14-aa5ceb292721
    """
    # Simple case. We want to make sure to return an integer in this
    # case.
    if num_digits == 0:
        return math.trunc(number)

    num_digits = int(num_digits)

    return math.trunc(number * 10**num_digits) / 10**num_digits

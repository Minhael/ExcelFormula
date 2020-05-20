//  https://en.wikipedia.org/wiki/Luhn_algorithm

//  Assume Cell J2 stores the full number, including the check digit


//  Remove the last check digit from original number and split into array in reverse order

PAN=MID(TEXTJOIN("",1,MID(MID(J2, 1, LEN(J2) - 1),LEN(J2) - ROW(INDIRECT("1:"&LEN(J2)-1)),1)),ROW(INDIRECT("1:"&LEN(J2)-1)),1)


//  Prepare the multiply factor array with correct length

FACTOR=MOD(ROW(INDIRECT("1:"&LEN(J2)-1)), 2) + 1


//  Calculate the product at even position

DOUBLE = PAN * (FACTOR)


//  Convert to TEXT type, and split into array of digits within [0-9]

//  Numbers < 10 -> ['0', '1']; Numbers > 9 -> ['1', '0'], assumed that doubling single digit will not larger than 99

SUMDIGITS = MID(TEXT(DOUBLE, "00"), {1,2}, 1)


//  Sum all digits

//  https://www.extendoffice.com/documents/excel/2037-excel-sum-digits.html

SUM = SUMPRODUCT(1*SUMDIGITS)


//  Calculate check digit

CHECKDIGIT = MOD(SUM * 9, 10)


//  Final formula, examinate value at A1

=MOD(SUMPRODUCT(1*MID(TEXT(MID(TEXTJOIN("",1,MID(MID(A1, 1, LEN(A1) - 1),LEN(A1) - ROW(INDIRECT("1:"&LEN(A1)-1)),1)),ROW(INDIRECT("1:"&LEN(A1)-1)),1) * (MOD(ROW(INDIRECT("1:"&LEN(A1)-1)), 2) + 1), "00"), {1,2}, 1)) * 9, 10)

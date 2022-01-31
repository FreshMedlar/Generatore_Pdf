birth = "1984/12/24"
year = birth.partition('/')[0]

month = birth.partition('/')[2].partition('/')[0]

day = birth.partition('/')[2].partition('/')[2].partition('/')[0]

_OMOCODIA = {
    "0": "L",
    "1": "M",
    "2": "N",
    "3": "P",
    "4": "Q",
    "5": "R",
    "6": "S",
    "7": "T",
    "8": "U",
    "9": "V",
}

maketrans = "".maketrans

_OMOCODIA_DIGITS = "".join([digit for digit in _OMOCODIA])

_OMOCODIA_LETTERS = "".join([_OMOCODIA[digit] for digit in _OMOCODIA])

_OMOCODIA_DECODE_TRANS = maketrans(_OMOCODIA_LETTERS, _OMOCODIA_DIGITS)
print(_OMOCODIA_DECODE_TRANS)

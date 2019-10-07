from decimal import Decimal, ROUND_HALF_UP

def realRound(d, n):
    format = '0.'
    while(n):
        format = format+'0'
        n= n-1
    return Decimal(str(d)).quantize(Decimal(format), rounding=ROUND_HALF_UP)
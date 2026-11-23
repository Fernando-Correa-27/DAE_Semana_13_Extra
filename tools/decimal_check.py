"""Four real comparisons between float and Decimal, for the lab evidence."""
from decimal import ROUND_HALF_UP, Decimal

print("0.1 + 0.2 =", 0.1 + 0.2)
print("Decimal('0.1') + Decimal('0.2') =", Decimal("0.1") + Decimal("0.2"))
print("round(0.845, 2) =", round(0.845, 2))
print(
    "Decimal('0.845').quantize(Decimal('0.01'), ROUND_HALF_UP) =",
    Decimal("0.845").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
)

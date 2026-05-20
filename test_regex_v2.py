import re
import sys
import os

sys.path.append(os.path.abspath('src'))
from services.excel_converter import _RAW_REF_RE

# Use raw string for formula to handle $ correctly
formula = r'IF(LEN(fact[IndicatorName="S2RLCR_RANK_"; SRF_LEI=?; endDate=?].value)=1;"0000"&fact[IndicatorName="S2RLCR_RANK_"; SRF_LEI=?; endDate=?].value;"")'
print(f'Input: {formula}')

output = _RAW_REF_RE.sub(lambda m: f'REPLACED({m.group(0)})', formula)

print(f'Output: {output}')
print(f'Changed: {formula != output}')

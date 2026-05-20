import re
import sys
import os

# Put the src directory in the path to import from services
sys.path.append(os.path.abspath('src'))
from services.excel_converter import _RAW_REF_RE

formula = 'IF(LEN(fact[IndicatorName=\"S2RLCR_RANK_\"; SRF_LEI=?; endDate=?].value)=1;\"0000\"&fact[IndicatorName=\"S2RLCR_RANK_\"; SRF_LEI=?; endDate=?].value;\"\")'
print(f'Input: {formula}')

# Simulate how it might be used in the code
# Usually _RAW_REF_RE is used with a lambda or sub
output = _RAW_REF_RE.sub(lambda m: f'REPLACED({m.group(0)})', formula)

print(f'Output: {output}')
print(f'Changed: {formula != output}')

import re

matches = re.findall(r'(?:^|\s|-)([a-zA-Z]{1,2}|VUÔNG|TRÒN|Φ|ϕ)?\s*(\d+(?:[.,]\d+)?(?:mm|m)?)(?:\b|$)', "Tôn lạnh 0.4mm", re.I)
print("Matches:", matches)

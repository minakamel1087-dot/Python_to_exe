import re
from engine.rules import RuleSet
from engine.clearance_validation import validate_clearance
data = open('WIR_Clearance Log Validation.xlsm','rb').read()
r = RuleSet.from_workbook(data, config_dir='reference')
out = validate_clearance(data, r)

c2r = [x for x in out.rows if 'C2R' in str(x.get('Fixes_Applied','')).split('; ')]
mep = [x for x in c2r if 'MEP Room' in str(x.get('Area_Validated'))]
spurious = [x for x in mep if not re.search(r'MEP\s*room', str(x.get('Clearance Description') or ''), re.I)]
print('C2R rows:', len(c2r))
print('  resolved to MEP Room:', len(mep))
print('  ...with no literal "MEP room" in the description:', len(spurious))

c2a = [x for x in out.rows if 'C2A' in str(x.get('Fixes_Applied','')).split('; ')]
sr = [x for x in c2a if re.search(r'service\s*room', str(x.get('Clearance Description') or ''), re.I)]
print()
print('C2A rows:', len(c2a))
print('  whose description says "service room":', len(sr))
for x in sr[:3]:
    print('   ', x.get('Clearance no.'), '|', ' '.join(str(x.get('Clearance Description')).split())[:96])

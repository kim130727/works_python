

import pypyodbc
msa_drivers = [x for x in pypyodbc.drivers() if 'ACCESS' in x.upper()]
print(f'MS-Access drivers: \n{msa_drivers}')
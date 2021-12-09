
import pubchempy as pcp

시료명1 = 'chlorantraniliprole'
시료명1CID = pcp.get_compounds(시료명1, 'name')
print (시료명1CID[0].fingerprint)

시료명2 = 'cyantraniliprole'
시료명2CID = pcp.get_compounds(시료명2, 'name')
print (시료명2CID[0].fingerprint)

시료명3 = 'acetamiprid'
시료명3CID = pcp.get_compounds(시료명3, 'name')
print (시료명3CID[0].fingerprint)
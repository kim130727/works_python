from rdkit import Chem

m = Chem.MolFromSmiles('Cc1ccccc1')

print (m)

print(Chem.MolToMolBlock(m))

# 구조 그리기2
from rdkit.Chem import Draw
print (Draw.MolToFile(m,'cdk2.png'))

#topological fingerprints

from rdkit import DataStructs
ms = [Chem.MolFromSmiles('Cc1ccccc1'), Chem.MolFromSmiles('Cc1ccccc1')]
fps = [Chem.RDKFingerprint(x) for x in ms]
print ("similarity=", DataStructs.FingerprintSimilarity(fps[0],fps[1]))

#MACCS Keys

from rdkit.Chem import MACCSkeys
fps = [MACCSkeys.GenMACCSKeys(x) for x in ms]
print ("similarity=", DataStructs.FingerprintSimilarity(fps[0],fps[1]))

#Atom Pairs and Topological Torsions

from rdkit.Chem.AtomPairs import Pairs
ms = [Chem.MolFromSmiles('C1CCC1OCC'),Chem.MolFromSmiles('CC(C)OCC'),Chem.MolFromSmiles('CCOCC')]
pairFps = [Pairs.GetAtomPairFingerprint(x) for x in ms]

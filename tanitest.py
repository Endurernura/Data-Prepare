from rdkit import Chem
from rdkit.Chem import RDKFingerprint
from rdkit.DataStructs import FingerprintSimilarity

# 苯甲酸的分子对象
benzoic_acid_smiles = "C1=CC=CC=C1C(=O)O"
benzoic_acid_mol = Chem.MolFromSmiles(benzoic_acid_smiles)
# 甲苯的分子对象
toluene_smiles = "C1=CC=CC=C1C"
toluene_mol = Chem.MolFromSmiles(toluene_smiles)

# 计算苯甲酸的Daylight - like拓扑指纹
benzoic_acid_fp = RDKFingerprint(benzoic_acid_mol)
# 计算甲苯的Daylight - like拓扑指纹
toluene_fp = RDKFingerprint(toluene_mol)

# 计算分子相似性
similarity = FingerprintSimilarity(benzoic_acid_fp, toluene_fp)
print("苯甲酸和甲苯的Daylight - like拓扑指纹相似性:", similarity)

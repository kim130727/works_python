import numpy as np
import matplotlib.pyplot as plt
from sklearn.datasets import make_classification
from sklearn.neural_network import MLPClassifier
import random


clf = MLPClassifier(activation='logistic', solver='lbfgs', alpha=1e-03, hidden_layer_sizes=(400, 100), learning_rate_init=0.0001, max_iter=800)

X = [[0.59,	43.15,	0,	0,	0,	0,	0,	1,	0.5,	0.5,	0.5,	0,	7,	0.2,	0.1,	0.5,	0,	5,	12,	28.960 ]
,[0.59,	43.15,	0,	0,	0,	0,	0,	1,	0.5,	0.5,	0.6,	0,	7,	0.2,	0.1,	0.5,	0,	5,	12,	28.860 ]
,[0.59,	43.15,	0,	0,	0,	0,	0,	1,	0.5,	0.5,	0.9,	0,	7,	0.2,	0.1,	0.5,	0,	5,	12,	28.560 ]
,[0.59,	43.15,	0,	0,	0,	0,	0,	1,	0.5,	0.5,	0.3,	0,	7,	0.1,	0.1,	0.5,	0,	5,	12,	29.260 ]
,[0.59,	43.15,	0,	0,	0,	0,	0,	1,	0.5,	0.5,	0.4,	0,	7,	0.1,	0.1,	0.5,	0,	5,	12,	29.160 ]
,[0.597,	43.57,	1.5,	0,	0,	0,	0,	0,	0,	0.5,	0.5,	0,	7,	0.1,	0.1,	0.5,	0,	5,	12,	28.633 ]
,[0.54,	39,	1.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	34.460 ]
,[0.54,	43.6,	0.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	30.860 ]
,[0.54,	43.6,	0.5,	0,	0.5,	2.5,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	28.360 ]
,[0.54,	43.6,	0.5,	2,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	28.860 ]
,[0.61,	43.6,	1.5,	0,	0,	0,	0,	0,	0,	0.3,	0.5,	0,	7,	0.1,	0.1,	0.5,	5,	0,	10,	30.790 ]
,[0.61,	43.6,	1.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	29.790 ]
,[0.54,	43.6,	0.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	30.860 ]
,[0.54,	43.6,	0.5,	0,	0.5,	2.5,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	28.360 ]
,[0.54,	43.6,	0.5,	2,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	28.860 ]
,[0.607,	43.2,	0.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	31.193 ]
,[0.607,	43.2,	0.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0,	7,	0.1,	0.1,	0.5,	5,	0,	10,	31.693 ]
,[0.607,	43.2,	1,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0,	7,	0.1,	0.1,	0.5,	5,	0,	10,	31.193 ]
,[0.607,	43.2,	1,	0,	0,	0,	2,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	29.193 ]
,[0.607,	43.2,	0.5,	0,	0.5,	0,	0,	0,	0,	0.3,	0.5,	0.5,	12,	0.1,	0.1,	0.5,	5,	0,	10,	26.193 ]
,[0.59,	43.2,	0.5,	0,	0.5,	0,	2,	0,	0,	0.3,	0.5,	0.5,	7,	0.1,	0.1,	0.5,	5,	0,	10,	29.210 ]
]

y = [8,	8,	12,	12,	12,	11,	8,	3,	3,	5,	8,	8,	3,	3,	5,	1,	5,	5,	6,	6,	6]

print (clf.fit(X, y))
print (clf.score(X, y))

treeHit = 0
while treeHit < 2000:
    treeHit = treeHit + 1
    Tiafenacil = 0.607
    Glyphosate = 43.2
    Atlox4915 = random.uniform(0, 1.5)
    Atlox4913 = random.uniform(0, 2)
    Atlox4894 = random.uniform(0, 0.5)
    SU500 = random.uniform(0, 2.5)
    SC141C = random.uniform(0, 2)
    SC600 = random.uniform(0, 1)
    TSP150PG = random.uniform(0, 0.5)
    SAG1538 = random.uniform(0.3, 0.5)
    Citricacid = random.uniform(0.3, 0.9)
    EDTAacid = random.uniform(0, 0.5)
    PG = random.uniform(7, 12)
    GXL = random.uniform(0.1, 0.2)
    XG = 0.1
    Attagel50 = 0.5
    FKC1800 = random.uniform(0, 5)
    Teric = random.uniform(0, 5)
    Milcoside = random.uniform(10, 12)
    Water = 100 - Tiafenacil - Glyphosate - Atlox4915 - Atlox4913 - Atlox4894 - SU500 - SC141C - SC600 - TSP150PG - SAG1538 - Citricacid - EDTAacid - PG - GXL - XG - Attagel50 - FKC1800 - Teric - Milcoside

    if Water > 0:
        print(Tiafenacil, end=" ")
        print(Glyphosate, end=" ")
        print(Atlox4915, end=" ")
        print(Atlox4913, end=" ")
        print(Atlox4894, end=" ")
        print(SU500, end=" ")
        print(SC141C, end=" ")
        print(SC600, end=" ")
        print(TSP150PG, end=" ")
        print(SAG1538, end=" ")
        print(Citricacid, end=" ")
        print(EDTAacid, end=" ")
        print(PG, end=" ")
        print(GXL, end=" ")
        print(XG, end=" ")
        print(Attagel50, end=" ")
        print(FKC1800, end=" ")
        print(Teric, end=" ")
        print(Milcoside, end=" ")
        print(Water, end=" ")

        print(clf.predict([[Tiafenacil
,Glyphosate
,Atlox4915
,Atlox4913
,Atlox4894
,SU500
,SC141C
,SC600
,TSP150PG
,SAG1538
,Citricacid
,EDTAacid
,PG
,GXL
,XG
,Attagel50
,FKC1800
,Teric
,Milcoside
,Water]]))

    else:
        pass


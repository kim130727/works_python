import numpy as np
import matplotlib.pyplot as plt
from sklearn.datasets import make_classification
from sklearn.neural_network import MLPClassifier


clf = MLPClassifier(solver='lbfgs', alpha=1e-4, hidden_layer_sizes=(300, 300))

X = [[11.15],[15.70],[18.90],[19.40],[21.40],[21.70],[25.30],[26.40],[26.70],[29.10]]
y = [4,0,0,0,0,0,0,1,0,4]

print (clf.fit(X, y))

print (clf.predict([[29.10]]))
print (clf.predict_proba([[29.10]]))
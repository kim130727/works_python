import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
from sklearn.datasets import make_moons, make_circles, make_classification
from sklearn.neural_network import MLPClassifier

clf = MLPClassifier(solver='lbfgs', alpha=1e-5, hidden_layer_sizes=(10, 10), random_state=1)

names = []
classifiers = []

X, y = make_classification(n_features=2, n_redundant=0, n_informative=2, random_state=1, n_clusters_per_class=1)

print (X)
print (y)

print (clf.fit(X, y))

print (clf.predict([[0., 0.]]))

print (clf.predict_proba([[0., 0.]]))

print (clf.predict_proba([[1., 1.]]))

print (clf.predict_proba([[2., 2.]]))

plt.plot(X,y)
plt.show()
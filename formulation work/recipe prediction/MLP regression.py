import numpy as np
import matplotlib.pyplot as plt
from sklearn.datasets import make_classification
from sklearn.neural_network import MLPRegressor


clf = MLPRegressor(activation='logistic', solver='lbfgs', alpha=0.00001, hidden_layer_sizes=(400, 200), learning_rate_init=0.0001, max_iter=10000)

X = [

[0.001]
,[0.193026234108713]
,[0.92515750226166]
,[0.108243073775708]
,[0.412920323240435]
,[0.067523514751084]



]

y = [
[0.50151948151935]
,[0.642918301481307]
,[0.990600757070195]
,[0.669105694099201]
,[0.951157065422061]
,[0.548429197846704]



]

print (clf.fit(X, y))

print (clf.score(X, y))

print (clf.predict([[0.001]])) # y = 0.50151948151935

print (clf.predict([[0.480187004062368]])) # y= 0.972033342922662

print (clf.predict([[0.502146375925832]])) # y = 0.974522912747041

print (clf.predict([[0.928957120682498]])) # y= 0.99916694099192

print (clf.predict([[0.114403599478374]])) # y= 0.58006478661757
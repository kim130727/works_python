from sklearn import datasets
from sklearn import svm
import pylab as pl

iris = datasets.load_iris()
digits = datasets.load_digits()

clf = svm.SVC(gamma=0.01, C=10.)

print(clf.fit(digits.data[:-1], digits.target[:-1]))

print(digits.images.shape)

print(pl.imshow(digits.images[-1], cmap=pl.cm.gray_r))

clf.predict(digits.data[-1])
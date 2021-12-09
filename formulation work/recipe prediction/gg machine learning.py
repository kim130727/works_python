import numpy as np

f = open("c:\data automation\gg.txt")
f.readline()  # skip the header
data = np.loadtxt(f)

X = data[:,:-1]
y = data[:,-1]

print (X)
print (y)

from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

scaler = StandardScaler()
scaler.fit(X)
X = scaler.transform(X)

X_train, X_test, y_train, y_test = train_test_split(X, y, random_state=0)
logreg = LogisticRegression().fit(X_train, y_train)
print ("테스트 세트 점수: {:.2f}".format(logreg.score(X_test, y_test)))

from sklearn.dummy import DummyClassifier
eclf = DummyClassifier(strategy="most_frequent").fit(X_train, y_train)
print ("테스트 세트 점수: {:.2f}".format(eclf.score(X_test, y_test)))

from sklearn.svm import SVC
eclf1 = SVC(gamma=1).fit(X_train, y_train)
print ("테스트 세트 점수: {:.2f}".format(eclf1.score(X_test, y_test)))

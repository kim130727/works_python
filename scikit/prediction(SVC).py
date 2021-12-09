import requests
import json
import ast
import datetime
import pandas as pd

crypto = input("예측하고 싶은 암호화폐 이름을 입력하세요: ")
period = input("평가기간을 입력하세요: ")

url = "http://www.coincap.io/history/"+crypto

def date(days):
    source_code = requests.get(url)
    plain_text = source_code.text
    jsonString = json.dumps(plain_text)
    dict2 = ast.literal_eval(plain_text)

    date1 = datetime.datetime.fromtimestamp((float(dict2['market_cap'][days][0])) / 1000).strftime('%Y-%m-%d %H:%M:%S')
    date2 = dict2['price'][days][1]

    date3 = str(date1) + " " + str(date2)
    return date3

def predict(days):

    import numpy as np


    source_code = requests.get(url)
    plain_text = source_code.text

    jsonString = json.dumps(plain_text)

    dict = json.loads(jsonString)
    dict2 = ast.literal_eval(plain_text)

    volume = pd.DataFrame(dict2['volume'], columns=['Date', 'Volume'])
    price =  pd.DataFrame(dict2['price'], columns=['Date', 'Price'])
    volume = volume.drop('Date', 1)

    price.insert(len(price.columns), "Volume", volume)

    price = price[price['Volume']!=0]

    ma5 = price['Price'].rolling(window=5).mean()
    ma10 = price['Price'].rolling(window=10).mean()
    ma20 = price['Price'].rolling(window=20).mean()
    ma30 = price['Price'].rolling(window=30).mean()
    ma60 = price['Price'].rolling(window=60).mean()
    ma120 = price['Price'].rolling(window=120).mean()
    v5 = price['Volume'].rolling(window=5).mean()
    v10 = price['Volume'].rolling(window=10).mean()
    v20 = price['Volume'].rolling(window=20).mean()
    v30 = price['Volume'].rolling(window=30).mean()
    v60 = price['Volume'].rolling(window=60).mean()
    v120 = price['Volume'].rolling(window=120).mean()
    std5 = price['Price'].rolling(window=5).std()
    std10 = price['Price'].rolling(window=10).std()
    std20 = price['Price'].rolling(window=20).std()
    std30 = price['Price'].rolling(window=30).std()
    std60 = price['Price'].rolling(window=60).std()
    std120 = price['Price'].rolling(window=120).std()
    v5_s = price['Volume'].rolling(window=5).std()
    v10_s = price['Volume'].rolling(window=10).std()
    v20_s = price['Volume'].rolling(window=20).std()
    v30_s = price['Volume'].rolling(window=30).std()
    v60_s = price['Volume'].rolling(window=60).std()
    v120_s = price['Volume'].rolling(window=120).std()

    f5_1_v = v5 + (v5_s * 2)
    f5_2_v = v5 - (v5_s * 2)
    f10_1_v = v10 + (v10_s * 2)
    f10_2_v = v10 - (v10_s * 2)
    f20_1_v = v20 + (v20_s * 2)
    f20_2_v = v20 - (v20_s * 2)
    f30_1_v = v30 + (v30_s * 2)
    f30_2_v = v30 - (v30_s * 2)
    f60_1_v = v60 + (v60_s * 2)
    f60_2_v = v60 - (v60_s * 2)
    f120_1_v = v120 + (v120_s * 2)
    f120_2_v = v120 - (v120_s * 2)

    f5_1 = ma5 + (std5 * 2)
    f5_2 = ma5 - (std5 * 2)
    f10_1 = ma10 + (std10 * 2)
    f10_2 = ma10 - (std10 * 2)
    f20_1 = ma20 + (std20 * 2)
    f20_2 = ma20 - (std20 * 2)
    f30_1 = ma30 + (std30 * 2)
    f30_2 = ma30 - (std30 * 2)
    f60_1 = ma60 + (std60 * 2)
    f60_2 = ma60 - (std60 * 2)
    f120_1 = ma120 + (std120 * 2)
    f120_2 = ma120 - (std120 * 2)

    price.insert(len(price.columns), "MA5", ma5)
    price.insert(len(price.columns), "MA10", ma10)
    price.insert(len(price.columns), "MA20", ma20)
    price.insert(len(price.columns), "MA30", ma30)
    price.insert(len(price.columns), "MA60", ma60)
    price.insert(len(price.columns), "MA120", ma120)
    price.insert(len(price.columns), "V5", v5)
    price.insert(len(price.columns), "V10", v10)
    price.insert(len(price.columns), "V20", v20)
    price.insert(len(price.columns), "V30", v30)
    price.insert(len(price.columns), "V60", v60)
    price.insert(len(price.columns), "V120", v120)
    price.insert(len(price.columns), "f5_1", f5_1)
    price.insert(len(price.columns), "f5_2", f5_2)
    price.insert(len(price.columns), "f10_1", f10_1)
    price.insert(len(price.columns), "f10_2", f10_2)
    price.insert(len(price.columns), "f20_1", f20_1)
    price.insert(len(price.columns), "f20_2", f20_2)
    price.insert(len(price.columns), "f30_1", f30_1)
    price.insert(len(price.columns), "f30_2", f30_2)
    price.insert(len(price.columns), "f60_1", f60_1)
    price.insert(len(price.columns), "f60_2", f60_2)
    price.insert(len(price.columns), "f120_1", f120_1)
    price.insert(len(price.columns), "f120_2", f120_2)
    price.insert(len(price.columns), "f5_1_v", f5_1_v)
    price.insert(len(price.columns), "f5_2_v", f5_2_v)
    price.insert(len(price.columns), "f10_1_v", f10_1_v)
    price.insert(len(price.columns), "f10_2_v", f10_2_v)
    price.insert(len(price.columns), "f20_1_v", f20_1_v)
    price.insert(len(price.columns), "f20_2_v", f20_2_v)
    price.insert(len(price.columns), "f30_1_v", f30_1_v)
    price.insert(len(price.columns), "f30_2_v", f30_2_v)
    price.insert(len(price.columns), "f60_1_v", f60_1_v)
    price.insert(len(price.columns), "f60_2_v", f60_2_v)
    price.insert(len(price.columns), "f120_1_v", f120_1_v)
    price.insert(len(price.columns), "f120_2_v", f120_2_v)

    price = price.dropna(axis=0)
    price = price.values

    f1 = (price[0][1] - price[0][3]) / price[0][1]
    f2 = (price[0][1] - price[0][4]) / price[0][1]
    f3 = (price[0][1] - price[0][5]) / price[0][1]
    f4 = (price[0][1] - price[0][6]) / price[0][1]
    f5 = (price[0][1] - price[0][7]) / price[0][1]
    f6 = (price[0][1] - price[0][8]) / price[0][1]
    f7 = (price[0][2] - price[0][9]) / price[0][2]
    f8 = (price[0][2] - price[0][10]) / price[0][2]
    f9 = (price[0][2] - price[0][11]) / price[0][2]
    f10 = (price[0][2] - price[0][12]) / price[0][2]
    f11 = (price[0][2] - price[0][13]) / price[0][2]
    f12 = (price[0][2] - price[0][14]) / price[0][2]
    f13 = (price[0][1] - price[0][15]) / price[0][1]
    f14 = (price[0][1] - price[0][16]) / price[0][1]
    f15 = (price[0][1] - price[0][17]) / price[0][1]
    f16 = (price[0][1] - price[0][18]) / price[0][1]
    f17 = (price[0][1] - price[0][19]) / price[0][1]
    f18 = (price[0][1] - price[0][20]) / price[0][1]
    f19 = (price[0][1] - price[0][21]) / price[0][1]
    f20 = (price[0][1] - price[0][22]) / price[0][1]
    f21 = (price[0][1] - price[0][23]) / price[0][1]
    f22 = (price[0][1] - price[0][24]) / price[0][1]
    f23 = (price[0][1] - price[0][25]) / price[0][1]
    f24 = (price[0][1] - price[0][26]) / price[0][1]
    f25 = (price[0][2] - price[0][27]) / price[0][2]
    f26 = (price[0][2] - price[0][28]) / price[0][2]
    f27 = (price[0][2] - price[0][29]) / price[0][2]
    f28 = (price[0][2] - price[0][30]) / price[0][2]
    f29 = (price[0][2] - price[0][31]) / price[0][2]
    f30 = (price[0][2] - price[0][32]) / price[0][2]
    f31 = (price[0][2] - price[0][33]) / price[0][2]
    f32 = (price[0][2] - price[0][34]) / price[0][2]
    f33 = (price[0][2] - price[0][35]) / price[0][2]
    f34 = (price[0][2] - price[0][36]) / price[0][2]
    f35 = (price[0][2] - price[0][37]) / price[0][2]
    f36 = (price[0][2] - price[0][38]) / price[0][2]

    if ((price[5][3] - price[0][3]) / price[5][3]) > 0:
        result1 = 1
    else:
        result1 = 2

    if ((price[10][4] - price[0][4]) / price[10][4]) > 0:
        result2 = 1
    else:
        result2 = 2

    if ((price[20][5] - price[0][5]) / price[20][5]) > 0:
        result3 = 1
    else:
        result3 = 2

    if ((price[30][6] - price[0][6]) / price[30][6]) > 0:
        result4 = 1
    else:
        result4 = 2

    factor = np.array(
            [f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17, f18, f19, f20, f21, f22, f23,
             f24, f25, f26, f27, f28, f29, f30, f31, f32, f33, f34, f35, f36, result1, result2, result3, result4])

    n = 1

    try:
        while n <10000:
            f1 = (price[n][1] - price[n][3]) / price[n][1]
            f2 = (price[n][1] - price[n][4]) / price[n][1]
            f3 = (price[n][1] - price[n][5]) / price[n][1]
            f4 = (price[n][1] - price[n][6]) / price[n][1]
            f5 = (price[n][1] - price[n][7]) / price[n][1]
            f6 = (price[n][1] - price[n][8]) / price[n][1]
            f7 = (price[n][2] - price[n][9]) / price[n][2]
            f8 = (price[n][2] - price[n][10]) / price[n][2]
            f9 = (price[n][2] - price[n][11]) / price[n][2]
            f10 = (price[n][2] - price[n][12]) / price[n][2]
            f11 = (price[n][2] - price[n][13]) / price[n][2]
            f12 = (price[n][2] - price[n][14]) / price[n][2]
            f13 = (price[n][1] - price[n][15]) / price[n][1]
            f14 = (price[n][1] - price[n][16]) / price[n][1]
            f15 = (price[n][1] - price[n][17]) / price[n][1]
            f16 = (price[n][1] - price[n][18]) / price[n][1]
            f17 = (price[n][1] - price[n][19]) / price[n][1]
            f18 = (price[n][1] - price[n][20]) / price[n][1]
            f19 = (price[n][1] - price[n][21]) / price[n][1]
            f20 = (price[n][1] - price[n][22]) / price[n][1]
            f21 = (price[n][1] - price[n][23]) / price[n][1]
            f22 = (price[n][1] - price[n][24]) / price[n][1]
            f23 = (price[n][1] - price[n][25]) / price[n][1]
            f24 = (price[n][1] - price[n][26]) / price[n][1]
            f25 = (price[n][2] - price[n][27]) / price[n][2]
            f26 = (price[n][2] - price[n][28]) / price[n][2]
            f27 = (price[n][2] - price[n][29]) / price[n][2]
            f28 = (price[n][2] - price[n][30]) / price[n][2]
            f29 = (price[n][2] - price[n][31]) / price[n][2]
            f30 = (price[n][2] - price[n][32]) / price[n][2]
            f31 = (price[n][2] - price[n][33]) / price[n][2]
            f32 = (price[n][2] - price[n][34]) / price[n][2]
            f33 = (price[n][2] - price[n][35]) / price[n][2]
            f34 = (price[n][2] - price[n][36]) / price[n][2]
            f35 = (price[n][2] - price[n][37]) / price[n][2]
            f36 = (price[n][2] - price[n][38]) / price[n][2]

            if ((price[n + 5][3] - price[n][3]) / price[n + 5][3]) > 0:
                result1 = 1
            else:
                result1 = 2

            if ((price[n + 10][4] - price[n][4]) / price[n + 10][4]) > 0:
                result2 = 1
            else:
                result2 = 2

            if ((price[n + 20][5] - price[n][5]) / price[n + 20][5]) > 0:
                result3 = 1
            else:
                result3 = 2

            if ((price[n + 30][6] - price[n][6]) / price[n + 30][6]) > 0:
                result4 = 1
            else:
                result4 = 2

            factor = np.vstack((factor,
                                    [f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17, f18,
                                     f19, f20, f21, f22, f23, f24, f25, f26, f27, f28, f29, f30, f31, f32, f33, f34,
                                     f35, f36, result1, result2, result3, result4]))

            n = n+1

    except IndexError:
        pass

    import numpy as np
    from sklearn.neural_network import MLPClassifier
    from sklearn.neighbors import KNeighborsClassifier
    from sklearn.svm import SVC
    from sklearn import linear_model
    from sklearn.preprocessing import StandardScaler
    from sklearn.model_selection import cross_val_score
    from sklearn.dummy import DummyClassifier
    from sklearn.tree import DecisionTreeClassifier
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.naive_bayes import GaussianNB
    from sklearn.discriminant_analysis import QuadraticDiscriminantAnalysis
    from sklearn.model_selection import StratifiedKFold
    from sklearn.ensemble import VotingClassifier

    names = ["Nearest Neighbors", "Linear SVM", "RBF SVM", "Logistic classifier", "Neural Net", "DecisionTree", "RandomForest","Naive Bayes", "QDA","Dummy Classifier"]

    classifiers = [
        KNeighborsClassifier(),
        SVC(kernel="linear"),
        SVC(),
        linear_model.LogisticRegression(),
        MLPClassifier(activation='logistic', solver='lbfgs', alpha=0.5, hidden_layer_sizes=(1,), learning_rate_init=0.5),
        DecisionTreeClassifier(),
        RandomForestClassifier(),
        GaussianNB(),
        QuadraticDiscriminantAnalysis(),
        DummyClassifier(strategy="most_frequent")]

    a = np.array([0])
    b = np.array([0])
    c = np.array([0])
    d = np.array([0])
    e = np.array([0])
    f = np.array([0])


    xy = factor.transpose()
    X = np.transpose(xy[0:36])
    y1 = np.transpose(xy[36])
    y2 = np.transpose(xy[37])
    y3 = np.transpose(xy[38])
    y4 = np.transpose(xy[39])

    scaler = StandardScaler()
    scaler.fit(X)
    X = scaler.transform(X)

    f1 = (price[days][1]-price[days][3])/price[days][1]
    f2 = (price[days][1]-price[days][4])/price[days][1]
    f3 = (price[days][1]-price[days][5])/price[days][1]
    f4 = (price[days][1]-price[days][6])/price[days][1]
    f5 = (price[days][1]-price[days][7])/price[days][1]
    f6 = (price[days][1]-price[days][8])/price[days][1]
    f7 = (price[days][2]-price[days][9])/price[days][2]
    f8 = (price[days][2]-price[days][10])/price[days][2]
    f9 = (price[days][2]-price[days][11])/price[days][2]
    f10 = (price[days][2]-price[days][12])/price[days][2]
    f11 = (price[days][2]-price[days][13])/price[days][2]
    f12 = (price[days][2]-price[days][14])/price[days][2]
    f13 = (price[days][1] - price[days][15]) / price[days][1]
    f14 = (price[days][1] - price[days][16]) / price[days][1]
    f15 = (price[days][1] - price[days][17]) / price[days][1]
    f16 = (price[days][1] - price[days][18]) / price[days][1]
    f17 = (price[days][1] - price[days][19]) / price[days][1]
    f18 = (price[days][1] - price[days][20]) / price[days][1]
    f19 = (price[days][1] - price[days][21]) / price[days][1]
    f20 = (price[days][1] - price[days][22]) / price[days][1]
    f21 = (price[days][1] - price[days][23]) / price[days][1]
    f22 = (price[days][1] - price[days][24]) / price[days][1]
    f23 = (price[days][1] - price[days][25]) / price[days][1]
    f24 = (price[days][1] - price[days][26]) / price[days][1]
    f25 = (price[days][2] - price[days][27]) / price[days][2]
    f26 = (price[days][2] - price[days][28]) / price[days][2]
    f27 = (price[days][2] - price[days][29]) / price[days][2]
    f28 = (price[days][2] - price[days][30]) / price[days][2]
    f29 = (price[days][2] - price[days][31]) / price[days][2]
    f30 = (price[days][2] - price[days][32]) / price[days][2]
    f31 = (price[days][2] - price[days][33]) / price[days][2]
    f32 = (price[days][2] - price[days][34]) / price[days][2]
    f33 = (price[days][2] - price[days][35]) / price[days][2]
    f34 = (price[days][2] - price[days][36]) / price[days][2]
    f35 = (price[days][2] - price[days][37]) / price[days][2]
    f36 = (price[days][2] - price[days][38]) / price[days][2]

    X1 = np.array([[f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17, f18, f19, f20, f21, f22, f23, f24, f25, f26, f27, f28, f29, f30, f31, f32, f33, f34,
                                     f35, f36]])

    X_pred = X1
    X_pred = scaler.transform(X_pred)

    eclf1 = SVC(gamma=0.001).fit(X, y1)
    final1 = np.vstack((a,eclf1.predict(X_pred)))
    final1 = np.delete(final1, [0], axis=0)

    eclf2 = SVC(gamma=0.001).fit(X, y2)
    final2 = np.vstack((b,eclf2.predict(X_pred)))
    final2 = np.delete(final2, [0], axis=0)

    eclf3 = SVC(gamma=0.001).fit(X, y3)
    final3 = np.vstack((c,eclf3.predict(X_pred)))
    final3 = np.delete(final3, [0], axis=0)

    eclf4 = SVC(gamma=0.001).fit(X, y4)
    final4 = np.vstack((d,eclf4.predict(X_pred)))
    final4 = np.delete(final4, [0], axis=0)


    final = final1, final2, final3, final4

    return final

n = -1*int(period)

while n < 0:
    print (date(n)," ",predict(n))
    n = n+1
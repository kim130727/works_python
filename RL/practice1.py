#imports, nothing to see here
import numpy as np
from scipy import stats
import random
import matplotlib.pyplot as plt
import IPython

n = 10
arms = np.random.rand(n)
eps = 0.1

av = np.ones(n) #initialize action-value array
counts = np.zeros(n) #stores counts of how many times we've taken a particular action

print('arms:',arms)
print('av:',av)
print('counts:',counts)

def reward(prob):
    total = 0;
    for i in range(10):
        if random.random() < prob:
            total += 1
    return total

print('reward값 0.5가 넘지 않는 것:',reward(0.5))

def bestArm(a):
    return np.argmax(a) #최대값이 존재하는 인덱스 위치를 찾아주는 함수

print ('bestarm 값이 0,2,4,5,6일 경우 6이 가장 크니 인덱스는 4가 됨', bestArm([0,2,4,5,6]))

plt.xlabel("Plays")
plt.ylabel("Mean Reward")

for i in range(500):
    if random.random() > eps:
        choice = bestArm(av)
        counts[choice] += 1
        k = counts[choice]
        rwd =  reward(arms[choice])
        old_avg = av[choice]
        new_avg = old_avg + (1/k)*(rwd - old_avg) #update running avg
        av[choice] = new_avg
        print (i, '단계', end="")
        print ('choice', choice, end="")
        print ('counts[choice]', counts[choice], end="")
        print ('k',k, end="")
        print ('rwd',rwd, end="")
        print ('old_avg', old_avg, end="")
        print ('new_avg', new_avg, end="")
        print ('av[choice]', av[choice])

    else:
        choice = np.where(arms == np.random.choice(arms))[0][0] #randomly choose an arm (returns index)
        counts[choice] += 1
        k = counts[choice]
        rwd =  reward(arms[choice])
        old_avg = av[choice]
        new_avg = old_avg + (1/k)*(rwd - old_avg) #update running avg
        av[choice] = new_avg
        print(i, '단계', end="")
        print('choice', choice, end="")
        print('counts[choice]', counts[choice], end="")
        print('k', k, end="")
        print('rwd', rwd, end="")
        print('old_avg', old_avg, end="")
        print('new_avg', new_avg, end="")
        print('av[choice]', av[choice])
    #have to use np.average and supply the weights to get a weighted average
    runningMean = np.average(av, weights=np.array([counts[i]/np.sum(counts) for i in range(len(counts))]))
    print ('running Mean', runningMean)
    plt.scatter(i, runningMean)


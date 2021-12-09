# needed imports
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
from scipy.cluster.hierarchy import cophenet
from scipy.spatial.distance import pdist
import numpy as np
from numpy import genfromtxt

my_data = genfromtxt('test.csv', delimiter=',')

# generate the linkage matrix
Z = linkage(my_data, 'ward')
c, coph_dists = cophenet(Z, pdist(my_data))

print (Z)

np.savetxt("test1.csv", Z, delimiter=",")






# needed imports
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
from scipy.cluster.hierarchy import cophenet
from scipy.spatial.distance import pdist
from scipy.cluster import vq
import numpy as np
from numpy import genfromtxt

my_data = genfromtxt('test.csv', delimiter=',')

# Calculating the cluster centroids and variance
# from kmeans
centroids, variance = vq.kmeans(my_data, 12)
# The identified variable contains the information
# we need to separate the points in clusters
# based on the vq function.
identified, distance = vq.vq(my_data, centroids)
# Retrieving coordinates for points in each vq
# identified core
vqc1 = my_data[identified == 0]
vqc2 = my_data[identified == 1]
vqc3 = my_data[identified == 2]
vqc4 = my_data[identified == 3]
vqc5 = my_data[identified == 4]
vqc6 = my_data[identified == 5]
vqc7 = my_data[identified == 6]
vqc8 = my_data[identified == 7]
vqc9 = my_data[identified == 8]
vqc10 = my_data[identified == 9]
vqc11 = my_data[identified == 10]
vqc12 = my_data[identified == 11]

np.savetxt("vqc1.csv", vqc1, delimiter=",")
np.savetxt("vqc2.csv", vqc2, delimiter=",")
np.savetxt("vqc3.csv", vqc3, delimiter=",")
np.savetxt("vqc4.csv", vqc4, delimiter=",")
np.savetxt("vqc5.csv", vqc5, delimiter=",")
np.savetxt("vqc6.csv", vqc6, delimiter=",")
np.savetxt("vqc7.csv", vqc7, delimiter=",")
np.savetxt("vqc8.csv", vqc8, delimiter=",")
np.savetxt("vqc9.csv", vqc9, delimiter=",")
np.savetxt("vqc10.csv", vqc10, delimiter=",")
np.savetxt("vqc11.csv", vqc11, delimiter=",")
np.savetxt("vqc12.csv", vqc12, delimiter=",")



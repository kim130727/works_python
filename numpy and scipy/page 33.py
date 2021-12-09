import numpy as np
from scipy.cluster import vq
import matplotlib.pyplot as plt

# Creating data
c1 = np.random.randn(100, 2) + 5
c2 = np.random.randn(30, 2) - 5
c3 = np.random.randn(50, 2)
# Pooling all the data into one 180 x 2 array
data = np.vstack([c1, c2, c3])
# Calculating the cluster centroids and variance
# from kmeans
centroids, variance = vq.kmeans(data, 3)
# The identified variable contains the information
# we need to separate the points in clusters
# based on the vq function.
identified, distance = vq.vq(data, centroids)
# Retrieving coordinates for points in each vq
# identified core
vqc1 = data[identified == 0]
vqc2 = data[identified == 1]
vqc3 = data[identified == 2]

plt.scatter(vqc1[:,0], vqc1[:,1], color='red')
plt.scatter(vqc2[:,0], vqc2[:,1], color='yellow')
plt.scatter(vqc3[:,0], vqc3[:,1])
plt.show()
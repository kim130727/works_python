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

print (Z[:20])

# calculate full dendrogram
plt.figure(figsize=(25, 10))
plt.title('Hierarchical Clustering Dendrogram')
plt.xlabel('sample index')
plt.ylabel('distance')
dendrogram(
    Z,
    truncate_mode = 'lastp',
    p = 12,
    leaf_rotation=90.,
    leaf_font_size=12.,
    show_contracted = True,
)
plt.show()




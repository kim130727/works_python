import numpy as np
from scipy.interpolate import griddata
import matplotlib.pyplot as plt

# Defining a function
ripple = lambda x, y: np.sqrt(x**2 + y**2)+np.sin(x**2 + y**2)

# Generating gridded data. The complex number defines
# how many steps the grid data should have. Without the
# complex number mgrid would only create a grid data structure
# with 5 steps.

grid_x, grid_y = np.mgrid[0:5:1000j, 0:5:1000j]
# Generating sample that interpolation function will see

xy = np.random.rand(1000, 2)

sample = ripple(xy[:,0] * 5 , xy[:,1] * 5)
# Interpolating data with a cubic

grid_z0 = griddata(xy * 5, sample, (grid_x, grid_y), method='cubic')

print(grid_x, grid_y)

plt.imshow(grid_z0.T, extent=(0,1,0,1), origin='lower')
plt.plot(xy[:,0], xy[:,1], 'k.', ms=1)
plt.show()
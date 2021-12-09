import numpy as np
import matplotlib.pyplot as plt
import scipy.interpolate

# Generate data:
x = [0.1,0.1,0.5,0.9,0.9]
y = [0.1,0.9,0.5,0.1,0.9]
z = [403, 260, 325, 271, 129]

print (x)
print (y)
print (z)

# Set up a regular grid of interpolation points
xi, yi = np.linspace(0, 1, 100), np.linspace(0, 1, 100)
xi, yi = np.meshgrid(xi, yi)

print (xi)
print (yi)

# Interpolate
rbf = scipy.interpolate.Rbf(x, y, z, function='linear')
zi = rbf(xi, yi)

print (zi)

plt.imshow(zi, vmin=0, vmax=700, origin='lower',
           extent=[0, 1, 0, 1])
plt.scatter(x, y, c=z, cmap=plt.cm.Reds)
plt.colorbar()
plt.show()
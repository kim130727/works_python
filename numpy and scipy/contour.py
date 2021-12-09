import numpy as np
import matplotlib.pyplot as plt
import scipy.interpolate

# Generate data:
x = [0, 0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1,1,0]
y = [0, 0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1,0,1]
z = [10,60, 70, 80, 90, 99, 90, 80,70,60,10,10,10]

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

plt.contourf(xi, yi, zi, np.arange(0, 100), cmap=plt.cm.get_cmap('rainbow'))
plt.colorbar()

plt.axhline(0, color='white')
plt.axvline(0, color='white')

plt.show()
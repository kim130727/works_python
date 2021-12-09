import numpy as np
from scipy.integrate import quad
import matplotlib.pyplot as plt

# Defining function to integrate
func = lambda x: np.cos(np.exp(x)) ** 2
# Integrating function with upper and lower
# limits of 0 and 3, respectively
solution = quad(func, 0, 3)
print (solution)
# The first element is the desired value
# and the second is the error.
# (1.296467785724373, 1.397797186265988e-09)

x = np.linspace(0,3,100)
print(x)

plt.plot(x, func(x))
plt.show()
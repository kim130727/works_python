
import numpy as np

xy = np.loadtxt('1.csv', delimiter=',', dtype=np.float32)
x_data = xy[:, 0:-1]
y_data = xy[:, [-1]]

print (x_data)
print (y_data)

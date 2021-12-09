from scipy import misc
import matplotlib.pyplot as plt
import numpy as np

f = misc.imread('40sec.jpg')
misc.imsave('40sec.png', f) # uses the Image module (PIL)
arr1d = np.arange_ndarray(1000)

print (arr1d)
print (arr1d.shape)
print (arr1d.dtype)

file1 = open('c:\data automation\chromatogram_list.txt', 'a')
file1.write(str(f))
file1.close()

plt.imshow(arr1d)
plt.show()
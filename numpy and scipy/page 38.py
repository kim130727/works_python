import numpy as np
from scipy.misc import imread, imsave
from glob import glob
# Getting the list of files in the directory
files = glob('c:\data automation 40sec.jpg')
# Opening up the first image for loop
im1 = imread(files[0]).astype(np.float32)
# Starting loop and continue co-adding new images

print (files)


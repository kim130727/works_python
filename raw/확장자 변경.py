
#replace file ext.
import glob
import os.path
files = glob.glob('c:\\data automation\\COA\\*.xml')
for x in files:
    if not os.path.isdir(x):
        print (x)
        x2 = x.replace('.xml', '.xls')
        print ('==> ' + x2)
        os.rename(x, x2)
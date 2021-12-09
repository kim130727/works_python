# a stacked bar plot with errorbars
import numpy as np
import matplotlib.pyplot as plt

SC = (66, 41, 44, 44, 33, 44, 33)
WP = (32, 46, 25, 9,	15,	17,	10)
EC = (26, 52, 28, 12, 19, 26, 21)
GR = (27, 24, 40, 38, 17, 19, 9)
SL = (18, 19, 14, 9,	18,	23,	16)
WG = (17, 18, 18, 12, 15, 22, 13)
ME = (6, 6,	2,	6,	7,	4,	2)
EW = (3,	4,	5,	3,	1,	4,	2)
SE = (4,	1,	5,	6,	2,	3,	2)
DC = (2,	2,	1,	3,	6,	3,	2)
DP = (3,	2,	0,	0,	0,	4,	0)
FG = (2,	2,	4,	1,	2,	1,	0)
OD = (0,	1,	0,	1,	0,	0,	0)
WS = (2,	0,	0,	1,	0,	0,	0)
FS = (1,	1,	2,	0,	0,	1,	0)
GG = (1,	0,	2,	1,	0,	2,	0)
UG = (2,	0,	1,	2,	0,	0,	0)
DT = (2,	2,	1,	8,	3,	4,	1)
etc = (16, 3,	8,	3,	5,	5, 3)

N = 7
ind = np.arange(N)    # the x locations for the groups
width = 0.35       # the width of the bars: can also be len(x) sequence

p1 = plt.bar(ind, SC, width)
p2 = plt.bar(ind, WP, width)
p3 = plt.bar(ind, EC, width)
p4 = plt.bar(ind, GR, width)
p5 = plt.bar(ind, SL, width)
p6 = plt.bar(ind, WG, width)
p7 = plt.bar(ind, ME, width)
p8 = plt.bar(ind, EW, width)
p9 = plt.bar(ind, SE, width)
p10 = plt.bar(ind, DC, width)
p11 = plt.bar(ind, DP, width)
p12 = plt.bar(ind, FG, width)
p13 = plt.bar(ind, OD, width)
p14 = plt.bar(ind, WS, width)
p15 = plt.bar(ind, FS, width)
p16 = plt.bar(ind, GG, width)
p17 = plt.bar(ind, UG, width)
p18 = plt.bar(ind, DT, width)
p19 = plt.bar(ind, etc, width)


plt.ylabel('Registration Number')
plt.xticks(ind, ('2012', '2013', '2014', '2015', '2016', '2017', '2018'))
plt.legend((p1[0], p2[0], p3[0], p4[0],p5[0], p6[0],p7[0], p8[0],p9[0], p10[0],p11[0], p12[0],p13[0], p14[0],p15[0], p16[0],p17[0], p18[0],p19[0]),
           ('SC','WP','EC','GR','SL','WG','ME','EW','SE','DC','DP','FG','OD','WS','FS','GG','UG','DT','etc','GR'),loc='upper center', bbox_to_anchor=(0.5, 1.05),
          ncol=10, fancybox=True, shadow=True)
plt.legend(loc='center left')
plt.show()
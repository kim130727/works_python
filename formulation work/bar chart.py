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
fig, ax = plt.subplots()
ind = np.arange(N)
width = 0.35

p1 = ax.bar(ind, SC, width, label='SC')
p2 = ax.bar(ind+width, WP, width)
p3 = ax.bar(ind+width*2, EC, width)
p4 = ax.bar(ind, GR, width)
p5 = ax.bar(ind, SL, width)
p6 = ax.bar(ind, WG, width)
p7 = ax.bar(ind, ME, width)
p8 = ax.bar(ind, EW, width)
p9 = ax.bar(ind, SE, width)
p10 = ax.bar(ind, DC, width)
p11 = ax.bar(ind, DP, width)
p12 = ax.bar(ind, FG, width)
p13 = ax.bar(ind, OD, width)
p14 = ax.bar(ind, WS, width)
p15 = ax.bar(ind, FS, width)
p16 = ax.bar(ind, GG, width)
p17 = ax.bar(ind, UG, width)
p18 = ax.bar(ind, DT, width)
p19 = ax.bar(ind, etc, width)

ax.set_xlabel('Year')
ax.set_ylabel('Registration Number')
ax.set_xticklabels(('2012', '2013', '2014', '2015', '2016', '2017', '2018'))

ax.legend()

fig.tight_layout()
plt.show()
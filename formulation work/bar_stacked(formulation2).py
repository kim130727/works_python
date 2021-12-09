# a stacked bar plot with errorbars
import numpy as np
import matplotlib.pyplot as plt

GM = (300.6684,	448.7655,	111.8325,	145.48125,	53.42925,	0)
FW= (2052.617911,	3044.65138173033,	918.930626654953,	2858.23153651774,	2540.863647,	2530.3041062)
GG = (3037.276571,	4078.539939,	4261.11815161222,	4675.19380425,	4878.510614,	4997.562737)
PA = (3575.3740667289,	6084.22181704656,	5341.70596186763,	4037.29040654732,	4297.5917,	3566.43307636364)
ME = (13557.990786,	8512.787825,	9372.74154739072,	10731.6472443431,	9223.231857,	13027.9011226)
DC = (4897.8836016,	5921.43225,	5493.89100573845,	6343.2390645,	7650.066,	11825.880493)
DP = (11060.5965464471,	13175.8000926678,	10305.8043682077,	7809.39980292507,	7587.81519575896,	7580.6044062)
FG = (4593.2323392,	3357.91425,	5361.782227,	6192.368294,	5948.434007,	6251.648057)
UG = (1062.877439,	2235.038998,	1472.718544,	2666.432077,	2493.539824,	1954.15585)
SP = (4918.74531469011,	5564.17726301206,	4944.03164017457,	4207.44589309017,	3997.955885,	4193.66420484)
WP = (214484.302645849,	207587.469672675,	191924.981513064,	178064.642268101,	177685.967051262,	177511.309663197)
SC = (198768.846076367,	214749.889509686,	223709.470875778,	229454.549132874,	240668.36237625,	243720.120900313)
SL = (101378.768468259,	104224.055011726,	127533.437202594,	149721.318424545,	137751.367455661,	141985.917372293)
OD = (0,	0,	1371.068076,	1531.038061,	1118.05018,	763.783423)
EC = (123238.34837817,	139329.128230224,	139751.500796212,	133111.79327495,	137834.401652818,	135804.400369543)
EW = (8105.29886264706,	9446.85066631043,	9125.56167420178,	10018.9723547403,	10092.057657,	13190.8235659365)
SE = (8776.30883844706,	8913.87911,	10414.3346883727,	8394.76417625,	11354.309095,	10792.455362)
SG = (4944.027768,	4948.5225,	4490.4117525,	6083.811477,	5990.24325,	5990.60858329412)
WG = (81633.2794474274,	88795.4020161191,	94196.9856747263,	91850.2434208161,	90857.0343665,	102478.264631183)
GR = (181364.108761119,	178093.911252479,	175535.520276102,	175543.756866934,	184171.222526599,	197143.37027916)
WS = (1397.111717,	1445.669228,	1735.05432425,	1853.651284,	1720.727314,	1203.181008)
FS = (3229.234442,	3357.57461835886,	4942.65257885655,	7766.7258867379,	6338.925363,	6441.529991)
AL = (0,	0,	0,	0,	0,	882.114048)
DT = (8675.54388,	13324.87275,	15411.7503195,	17935.9948491578,	20300.085812,	21875.708202)
CS = (6171.629815,	3082.023624,	3649.3299785,	3808.97453,	3799.61414,	3318.281931)
GA = (2788.83811,	7599.732431,	9727.2926155,	13167.831333,	16382.731287,	14779.338308)


N = 6
ind = np.arange(N)    # the x locations for the groups
width = 0.35       # the width of the bars: can also be len(x) sequence

p1 = plt.bar(ind, GM, width, color='#E0E0E0')
p2 = plt.bar(ind, FW, width, color='#C0C0C0', bottom=GM)
p3 = plt.bar(ind, GG, width, color='#9ecae1', bottom=[GM[j] +FW[j] for j in range(len(GM))])
p4 = plt.bar(ind, PA, width, color='#c6dbef', bottom=[GM[j] +FW[j]+GG[j] for j in range(len(GM))])
p5 = plt.bar(ind, ME, width, color='#e6550d', bottom=[GM[j] +FW[j]+GG[j]+PA[j] for j in range(len(GM))])
p6 = plt.bar(ind, DC, width, color='#fd8d3c', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j] for j in range(len(GM))])
p7 = plt.bar(ind, DP, width, color='#fdae6b', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j] for j in range(len(GM))])
p8 = plt.bar(ind, FG, width, color='#fdd0a2', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j] for j in range(len(GM))])
p9 = plt.bar(ind, UG, width, color='#8c564b', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j] for j in range(len(GM))])
p10 = plt.bar(ind, SP, width, color='#74c476', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j] for j in range(len(GM))])
p11 = plt.bar(ind, WP, width, color='#a1d99b', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j] for j in range(len(GM))])
p12 = plt.bar(ind, SC, width, color='#FF0000', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j] for j in range(len(GM))])
p13 = plt.bar(ind, SL, width, color='#330000', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j] for j in range(len(GM))])
p14 = plt.bar(ind, OD, width, color='#9e9ac8', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j] for j in range(len(GM))])
p15 = plt.bar(ind, EC, width, color='#2ca02c', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j] for j in range(len(GM))])
p16 = plt.bar(ind, EW, width, color='#dadaeb', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j] for j in range(len(GM))])
p17 = plt.bar(ind, SE, width, color='#636363', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j] for j in range(len(GM))])
p18 = plt.bar(ind, SG, width, color='#969696', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j] for j in range(len(GM))])
p19 = plt.bar(ind, WG, width, color='#bdbdbd', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j] for j in range(len(GM))])
p20 = plt.bar(ind, GR, width, color='#1f77b4', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j] for j in range(len(GM))])
p21 = plt.bar(ind, WS, width, color='#393b79', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j] for j in range(len(GM))])
p22 = plt.bar(ind, FS, width, color='#5254a3', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j]+WS[j] for j in range(len(GM))])
p23 = plt.bar(ind, AL, width, color='#6b6ecf', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j]+WS[j]+FS[j] for j in range(len(GM))])
p24 = plt.bar(ind, DT, width, color='#9c9ede', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j]+WS[j]+FS[j]+AL[j] for j in range(len(GM))])
p25 = plt.bar(ind, CS, width, color='#637939', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j]+WS[j]+FS[j]+AL[j]+DT[j] for j in range(len(GM))])
p26 = plt.bar(ind, GA, width, color='#8ca252', bottom=[GM[j] +FW[j]+GG[j]+PA[j]+ME[j]+DC[j]+DP[j]+FG[j]+UG[j]+SP[j]+WP[j]+SC[j]+SL[j]+OD[j]+EC[j]+EW[j]+SE[j]+SG[j]+WG[j]+GR[j]+WS[j]+FS[j]+AL[j]+DT[j]+GA[j] for j in range(len(GM))])


plt.ylabel('Thousand Won')
plt.xticks(ind, ('2012', '2013', '2014', '2015', '2016', '2017'))
plt.legend((p1[0], p2[0], p3[0], p4[0],p5[0], p6[0],p7[0], p8[0],p9[0], p10[0],p11[0], p12[0],p13[0], p14[0],p15[0], p16[0],p17[0], p18[0],p19[0], p20[0],p21[0], p22[0],p23[0], p24[0],p25[0], p26[0]),
           ('GM','FW','GG','PA','ME','DC','DP','FG','UG','SP','WP','SC','SL','OD','EC','EW','SE','SG','WG','GR','WS','FS','AL','DT','CS','GA'),loc='upper center', bbox_to_anchor=(0.5, 1.05),
          ncol=10, fancybox=True, shadow=True)
plt.legend(loc='center left')
plt.show()
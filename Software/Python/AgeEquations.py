from IsotopeConstants import Cnsts
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
import sys

Lambda235 = 0.00098485
Lambda238 = 0.000155125
TW_RatioUranium_UPb = 1 / 137.88

# Ratio76 = float(input('Please, write a 76Ratio > 0: '))

def CalculateAge76AndPlot(Ratio76):

    y1 = []
    y2 = []
    y3 = []
    y4 = []
    y5 = []
    Zeros = []

    Age76 = NewtonRaphson_76Age(10000,Ratio76)

    if isinstance(Age76,str)  == True:
        print('Age not calculated due to ' + Age76)
        sys.exit()

    if Age76[0] - 1000 < 0.0001:
        LL = 0.1 #Lower limit of the x axis of the plot
    else:
        LL = Age76[0] - 1000

    UL = Age76[0] + 1000

    x = np.arange(LL, UL, 0.1)

    for Age in x:
        # logging.info("np.exp(Lambda235 * i) = " + str(np.exp(Lambda235 * i)))
        Eq76 = (TW_RatioUranium_UPb * (np.exp(Lambda235 * Age) - 1) / (np.exp(Lambda238 * Age) - 1) - Ratio76)
        y1.append(Eq76)
        Zeros.append(0)

    Ys = [y1, y2, y3, y4, y5, Zeros]

    # plt.figure(1)
    plt.figure(1,figsize=(12, 6))

    for i in Ys:
        if not len(i) == 0:
            plt.plot(x, i)

    plt.title('207Pb/206Pb Age \n' + '76ratio = ' + str(Ratio76) + '\n' + str(Age76[2]) + ' interations to calculate ' + str(round(Age76[0],2)) + ' Ma using the Newton-Raphson method.')
    plt.xlabel('Age Ma')
    plt.ylabel('TW_RatioUranium_UPb * (EulerNum ** (Lambda235 * Age) - 1) / \n (EulerNum ** (Lambda238 * Age) - 1) - Ratio76')

    plt.show()

def NewtonRaphson_76Age(InitialGuess,Ratio76):

    delta=0.0000000000001
    Eq76_Guess = []
    Eq76_Diff_Guess = []
    Guess = InitialGuess

    count=1

    print('Ratio76 = ', Ratio76)
    TangentParameters = []

    if Ratio76 <= 0:
        return 'Error'

    for n in range(0,100001,1):
        try:
            Eq76_Guess=(TW_RatioUranium_UPb * (np.exp(Lambda235 * Guess) - 1) / (np.exp(Lambda238 * Guess) - 1) - Ratio76)
            Eq76_Diff_Guess= np.exp(Guess * Lambda235) * Lambda235 * TW_RatioUranium_UPb * 1 / (np.exp(Guess * Lambda238) - 1) \
                             - np.exp(Guess * Lambda238) * Lambda238 * TW_RatioUranium_UPb * (np.exp(Guess * Lambda235) - 1) * 1 / \
                            (np.exp(Guess * Lambda238) - 1) ** 2

            TangentParameters.append([Eq76_Guess,Eq76_Diff_Guess,Guess])

            Result = Guess - Eq76_Guess/Eq76_Diff_Guess

            # print(Eq76_Guess)
            # print(Eq76_Diff_Guess)
            # print(Result)

            if abs(Eq76_Guess) < delta:
                return [Guess,TangentParameters,count]

            Guess = Result
            count += 1

            # break

        except:
            print('Error calculating the age for ' + str(Ratio76) + '76 ratio.')
            return 'Error 1'

    print('Interations = ', count)
    return 'Error 2'

def DiffEquation76():
    import sympy

    TW_RatioUranium_UPb = sympy.Symbol('TW_RatioUranium_UPb')
    Lambda235 = sympy.Symbol('Lambda235')
    Lambda238 = sympy.Symbol('Lambda238')
    Ratio76 = sympy.Symbol('Ratio76')
    EulerNum = sympy.Symbol('EulerNum')
    Age = sympy.Symbol('Age')

    eq = (TW_RatioUranium_UPb * (EulerNum ** (Lambda235 * Age) - 1) / (EulerNum ** (Lambda238 * Age) - 1) - Ratio76)
    Eq_Diff = eq.diff(Age)
    print(Eq_Diff)

class SampleApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.entry = tk.Entry(self)
        self.button = tk.Button(self, text="Start", command=self.on_button)
        self.button.pack()
        self.entry.pack()

    def on_button(self):

        self.EntryValue = self.entry.get()
        self.quit()
        self.destroy()

def input76():
    app = SampleApp()
    app.mainloop()
    Ratio76 = float(app.EntryValue)
    CalculateAge76AndPlot(Ratio76)

def MSWD_Zi(sigma_Xi, sigma_Yi,m, ri):

    '''
    Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.17
    Wx " Wy• = weighting factors for X-and Y -coordinates of any data Point i
    ri correlation between analytical errors of X and Y for any sample i
    '''

    try:
        MSWD_Zi = 1 / (m ^ 2 * sigma_Xi ^ 2 + sigma_Yi ^ 2 - 2 * m * ri * sigma_Xi * sigma_Yi)

    except:
        print("MSWD_Zi error")
        return 'Error'

    return MSWD_Zi

def MSWD_W(sigma):
    '''
    Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.16
    sigma is the x or y uncertainty
    '''

    try:
        MSWD_W = 1 / sigma**2

    except:
        print("MSWD_W error")
        return 'Error'

    return MSWD_W

def MSWD_WeightedCenterOfGravity(Xi_or_Yi, Zi):
    '''
    Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.18 and 4.19
    '''
    summation1 = 0
    summation2 = 0

    if len(Xi_or_Yi) != len(Zi):
        print('WeightedCenterOfGravity')
        print('Lists sizes are different')
        return 'Error'

    try:
        for n in range(0,len(Xi_or_Yi)):
            summation1 += Xi_or_Yi[n] * Zi[n]
            summation2 += Zi[n]

        WeightedCenterOfGravity =  summation1/summation2

    except:
        print("WeightedCenterOfGravity error")
        return 'Error'

    return WeightedCenterOfGravity

def MSWD_Xm(Xi,Zi):
    '''
    Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.18 and 4.19
    '''

    return WeightedCenterOfGravity(Xi,Zi)

def MSWD_Ym(Yi,Zi):
    '''
    Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.18 and 4.19
    '''

    return WeightedCenterOfGravity(Yi,Zi)

def MSWD_S(Xi,Yi,Zi,m,b):

    '''Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.14

    'Yi , Xi = measured values of the X and Y parameters of each data point
    'm = slope of the best-fit straight line
    'b = intercept on the Y -axis of the best-fit straight line
    'Zi = weighting term for each sample in the regression
    '''

    S = 0

    length = len(Xi)
    if any(len(lst) != length for lst in [Yi,Zi]): #AN INTERESTING FOR STRUCTURE IN PYTHON
        print('MSWD_S')
        print('Lists sizes are different')
        return 'Error'

    try:
        for count in range(0, len(Zi)):
            S += (Yi[count] - m * Xi[count] - b)**2 * Zi[count]
    except:
        print("MSWD_S error")
        return 'Error'

    return S

def MSWD_m(Xi,Yi,sigma_Xi,sigma_Yi):

    '''Based on Harmer and Eglington (1990) apud Faure and Messing (2005) - Equation 4.20

    'Yi , Xi = measured values of the X and Y parameters of each data point
    'sigma_Xi , sigma_Yi = Yi and Xi uncertainties
    '''

    W_Xi = [1 / n ** 2 for n in sigma_Xi]
    W_Yi = [1 / n ** 2 for n in sigma_Yi]


    Zi = MSWD_Zi(sigma_Xi,sigma_Yi,m,ri)
    return [W_Xi,W_Yi]


Xi = [1,2,3,4,5,6,7,8,9]
Yi = [9,8,7,6,5,4,3,2,1]
sigma_Xi = [2,2,2,2,2,2,2,2,2]
sigma_Yi = [4,4,4,4,4,4,4,4,4]
Zi = [10,11,12,13,14,15,16,17,18]
m=2
b=10

# print(MSWD_S(Xi,Yi,Zi,m,b))
a = MSWD_m(Xi,Yi,sigma_Xi,sigma_Yi)
print(a[0])
print(a[1])
# input76()
# DiffEquation76()



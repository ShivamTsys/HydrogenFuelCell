import math
import numpy as np
import matplotlib.pyplot as plt 
from array import*
import xlsxwriter

global i, T, PO2
#T = float(input())


workbook = xlsxwriter.Workbook("HFCParamteres.xlsx")
bold_format = workbook.add_format({'bold':True})
cell_format = workbook.add_format()
cell_format.set_text_wrap()
cell_format.set_align('top')
cell_format.set_align('left=')

worksheet = workbook.add_worksheet('HFCParameters')


worksheet.write('A1', 'Current', bold_format)
worksheet.write('B1', 'Enernst', bold_format)
worksheet.write('C1', 'Vact', bold_format)
worksheet.write('D1', 'Rm', bold_format)
worksheet.write('E1', 'Rion', bold_format)
worksheet.write('F1', 'Vohm', bold_format)
worksheet.write('G1', 'Vconc', bold_format)
worksheet.write('H1', 'Vcell', bold_format)
worksheet.write('I1', 'Vstack', bold_format)
worksheet.write('J1', 'Power', bold_format)
worksheet.write('K1', 'Power_Stack', bold_format)
worksheet.write('L1', 'Efficiency', bold_format)

rowindex = 2

T = 273.15 + 70
c = array('f',[])
v = array('f',[])
k = array('f',[])
k_stack = array('f',[])
e = array('f',[])
x = 75.00
#I = float(input("Enter the  current in Ampere:"))  # current drawn from fuel cell in ampere(2)
#float(input("Enter the Temperature in Kelvin:"))#298.15
PO2 = 1
PH2 = 1
Enernest = float(1.229 - (8.5 * 10 ** -4) * (T - 298.15) + ((4.308 * 10 ** -5) * T * (math.log(PH2) + 0.5 * (math.log(PO2)))))
R = 8.3144  # KJ/Kmole.K
n = 1
Afc = 50.6
F = 96484  # columbs/mole
B = 0.015  # (R*T)/(n*F)
Jmax = 1.5
# data for calculation of Vact
zeta_1 = -0.948
zeta_2 = 0.0030373688787134006
zeta_3 = 7.6*(10**-5)
zeta_4 = -1.93*(10**-4)
# plot VI char
for i in np.arange(1,x,1,
                   dtype=float):
    print("Current = ", i)
    J = i / Afc
    # calculation of nernest eq
    print("The value of Enernest is: ",Enernest," volts")

    CH2 = (PH2) / (1.09 * (10 ** 6) * (math.exp(77 / T)))
    CO2 = (PO2) / (5.08 * (10 ** 6) * (math.exp(-498 / T))) 
    #print("CH2 is: ",CH2)
    #print("CO2 is: ",CO2)
    print()
    Vact = -((zeta_1)+(zeta_2*T)+(zeta_3*T)*(np.log(CO2))+(zeta_4)*(np.log(i)))
    print("The value of Activation loss is: ",Vact," volts")  # -0.7172

    # data for calculation of Vohm
    L = 0.0178  # float(input("Enter the value of thickness of membrane in cm:"))
    Afc = 50.6  # float(input("Enter the value of Active area of fuel cell in sq.cm :"))
    lamda = 23  # float(input("Enter the   value of semi-imperical variable:"))
    R = 8.3144
    F = 96485

    X = 181.6 * (1 + 0.03 * (i / Afc) + 0.062 * ((T / 303) ** 2) * (i / Afc) ** 2.5)  # 181.63
    Y = (lamda - 0.634 - 3 * (i / Afc)) * (math.exp(4.18 * (T - 303) / T))

    Rm = float(X / Y)  # Resistivity in ohm.cm
    Rions = float((Rm * L) / Afc)  # Ionic Resistivity
    Vohm = i * Rions

    print("The value of ohmic loss is: ",Vohm," volts")

    # calculation for Vcon : There is another formula for Vcon

    Vcon = -B * math.log(1 - (J / Jmax))

    print("The value of Concentration loss is: ",Vcon," volts")

    Loss = Vact + Vohm + Vcon
    print("The value of Loss is: ", Loss)

    N = 40
    Vcell = Enernest - Loss
    Vstack = N * Vcell
    #Eth = 1.23 
    P = Vcell*i
    P_stack = N*P 
    mu_f = 0.95
    HHV  = 1.482
    eff = (mu_f*Vcell)/HHV
    print("The value of Vcell is: ", Vcell)
    print("The value of Vstack is: ", Vstack)
    print("The value of efficiency is: ", eff)
    print("The value of power is: ", P)
    print("\n")


    worksheet.write('A'+str(rowindex), i)
    worksheet.write('B'+str(rowindex), Enernest)
    worksheet.write('C'+str(rowindex), Vact)
    worksheet.write('D'+str(rowindex), Rm)
    worksheet.write('E'+str(rowindex), Rions)
    worksheet.write('F'+str(rowindex), Vohm)
    worksheet.write('G'+str(rowindex), Vcon)
    worksheet.write('H'+str(rowindex), Vcell)
    worksheet.write('I'+str(rowindex), Vstack)
    worksheet.write('J'+str(rowindex), P)
    worksheet.write('K'+str(rowindex), P_stack)
    worksheet.write('L'+str(rowindex), eff)
    
    c.append(float(i))
    v.append(float(Vcell))
    k.append(float(P))
    k_stack.append(float(P_stack))
    e.append(float(eff))

    rowindex+=1

    

workbook.close()

'''plt.plot(c,v)
#plt.plot(c,P_stack)
plt.xlabel('current')
plt.ylabel('Vcell')
plt.title('VI  Characteristics at 70 degrees')
plt.show()

plt.plot(c,k)
plt.xlabel('current')
plt.ylabel('power')
plt.title('PI  Characteristics at 70 degrees')
plt.show()

plt.plot(c,e)
plt.xlabel('current')
plt.ylabel('eff')
plt.title('eff  Characteristics at 70 degrees')
plt.show()'''

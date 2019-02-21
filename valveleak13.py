# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 10:53:08 2018

@author: ketul
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jul 06 18:21:45 2018

@author: ketul
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Jun 24 19:26:51 2018

@author: ketul
"""

###############################################
###Import libraries
##############################################

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os        
import seaborn as sns
import math
import xlrd
from textwrap import wrap
#from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import matplotlib.ticker as ticker
                                  ##$ sudo pip3 install openpyxl


############################################################
###Input sheets
############################################################

Input_sheet = xlrd.open_workbook('Input Sheet.xlsx')
Input_variables = Input_sheet.sheet_by_name('Input Variables')

current_directory = os.getcwd()
final_directory = os.path.join(current_directory, r'%s'%str(Input_variables.cell(6,1).value))
if not os.path.exists(final_directory):
    os.makedirs(final_directory)

####################################################################
###Merging files
####################################################################

leak_test = pd.read_excel(str(Input_variables.cell(2,1).value))

#####
###Details of columns headings

keys = list(leak_test.columns)
values = list(leak_test.iloc[0])
LT_col = dict(zip(keys, values))
LT_col['Q0b']
####
##Removing unnecessary columns

leak_test1 = leak_test.drop([keys[0],keys[1],keys[2],keys[3],keys[4],keys[5],keys[6]], axis =1)
leak_test1 = leak_test1[1:]
columns = list(leak_test1.iloc[:,2:26].columns)
leak_test1 = leak_test1.dropna(axis=0,how= 'all',subset=columns)
leak_test1 = leak_test1.drop_duplicates()
#leak_test1 = leak_test1.drop_duplicates('Q0b')

#############################################################################################################
#####SAP


if(Input_variables.cell(1,1).value <3):
    SAP_Data= pd.read_excel(str(Input_variables.cell(3,1).value))
    col_name = list(SAP_Data.columns)
    SAP_Data.shape
else:
    SAP_Data= pd.read_excel(str(Input_variables.cell(4,1).value))
    SAP_VEX = pd.read_excel(str(Input_variables.cell(5,1).value))
    SAP_Data=SAP_Data.append(SAP_VEX,ignore_index = True)

###Complete Else statement


##Actual Merging
df = pd.merge(SAP_Data,leak_test1,how='right',left_on = 'Order', right_on = 'Q0b')
df = df.dropna(subset = ['Material description'])


df100 = pd.read_excel('Input Sheet.xlsx', sheetname = 'Max Leak')
df1 = pd.merge(df,df100, how='left',left_on = 'Material description', right_on = 'Model')
df1.columns

df1['Material description'] = df1['Material description'].str.replace('CPs','CPS')

##################################################################################################################
###############################################################################################################
###BAR CHARTS
###################################################################################################################      
########################################################################################################################################################################################



model = list(df100['Model'])

dz = pd.DataFrame(df1[df1['Model'].isna()!=True])

model1 =[]        
for i in range(0,len(model)):        
    a = str.split(str(model[i]))
    if(len(a) > 3):
        model1.append(a[len(a)-2]+' ' +a[len(a)-1])
    else:
        model1.append(a[len(a)-1])        
model1 = pd.DataFrame(model1)
model1 = pd.DataFrame(model1.iloc[:,0].drop_duplicates()) 
model1 = list(model1.iloc[:,0])    

for i in range(0,len(model1)):
        x = dz[dz['Material description'].str.contains('%s'%(model1[i])) == True]
        if (x.shape[0] != 0):
            
            y = x.iloc[:,12:36]                            
            y = y.dropna(axis = 0, how = 'all') 
            y = y.dropna(axis = 1, how = 'all')
            z = list(y.values.flatten())                    
            Leakvalue = [nn for nn in z if nn == nn]                      ##this will remove NA values
            df25 = pd.DataFrame(Leakvalue)
            df25['Max Leak'] = x.iloc[0]['Max Leak']                      ##add Max Leak column
            df25.columns.values[0] = 'leakvalue'                          ##first column name as leakvalue
            df25 = pd.DataFrame(df25,dtype=float)
            df25['percentage'] = (df25['leakvalue']/df25['Max Leak'])*100
            df25 = df25.round(2)
            ##df25 = df25.round({'percentage':2})
            #df25['percent']  = df25['percentage'].astype(str) + '%'
            df25.index += 1                                           ### start index from 1
                                                                      ### add dataset into dictionary   
            df25['ValveNumber'] =  df25.index                         ### add one column as valvenumber
          
            #######
            ##Start Plotting, Valve numbers vs Leak Value
            

               
            ############################
            ###Frequency distribution
            l=[]
            k = [0,25,50,75,100,110,120,130,3000]
            for m in range(0,len(k)):
                if (m < 5):
                    l.append(int(math.floor(x.iloc[0]['Max Leak']*k[m]/100)))
                else:
                    l.append(int(math.ceil(x.iloc[0]['Max Leak']*k[m]/100)))
            
            xlab =['%s-%s'%(l[0],l[1]),'%s-%s'%(l[1]+1,l[2]),'%s-%s'%(l[2]+1,l[3]),'%s-%s'%(l[3]+1,l[4]),'%s-%s'%(l[4]+1,l[5]),'%s-%s'%(l[5]+1,l[6]),'%s-%s'%(l[6]+1,l[7]),'\n'.join(wrap('%s and above'%(l[7]), 10))]
              
            h=np.histogram(df25['leakvalue'],bins=l)
            
            bar = plt.bar(range(8),h[0],width=1,align = 'center',zorder=3)
            bar[0].set_color('darkblue')
            bar[1].set_color('darkblue')
            bar[2].set_color('darkblue')
            bar[3].set_color('darkblue')
            bar[4].set_color('darkred')
            bar[5].set_color('darkred')
            bar[6].set_color('darkred')
            bar[7].set_color('darkred')
       
            ax = plt.gca()
            plt.xticks(np.arange(0,8,1),xlab)
            minor_ticks = list(np.arange(-0.5,8.5,1))
            ax.spines['left'].set_position(('data',-0.5))
            ax.set_xticks(minor_ticks, minor=True)
            ax.grid(which='minor', axis='both', linestyle='-',alpha = 1,zorder=0)
            ax.grid(which='major', axis='y', linestyle='-',alpha = 1,zorder=0)
            plt.xlim(xmin = -0.5,xmax=7.5)
        
            
            
            
           
            plt.title('%s VALVES'%(model1[i]), fontsize = 12, weight = 'bold',y = 1.08)
            plt.xlabel('LEAK VALUE RANGE', fontsize = 10, weight = 'bold',labelpad = 3)
            plt.ylabel('NUMBER OF VALVES', fontsize = 10, weight = 'bold',labelpad =3)
            plt.tick_params(axis='x', which='major', labelsize=8,length = 0,pad = 7)
            plt.tick_params(axis='y', which='major', labelsize=8,length = 3,pad = 5)
            plt.tick_params(axis='x', which='minor', length = 3,pad = 5)
            plt.xticks(weight = 'bold')
            plt.yticks(weight = 'bold')
            
        
            from matplotlib.ticker import MaxNLocator
            
            if ((np.max(h[0]) > 20)&(np.max(h[0]) < 100)): 
                plt.ylim(ymax =np.max(h[0])+2)
            elif (np.max(h[0]) > 100):
                plt.ylim(ymax =np.max(h[0])+5)
            elif (np.max(h[0]) > 300):
                plt.ylim(ymax =np.max(h[0])+10)
            
            ax.yaxis.set_major_locator(MaxNLocator(integer=True,steps=[1,2,5,10]))
            
            
            
            
            new = os.path.join(final_directory, r'Frequency distribution')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'Frequency Leak Value_combined')
            if not os.path.exists(new1):
                os.makedirs(new1)
            plt.tight_layout()
                        
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new1)
            fig2.savefig(P_BAR + "\\%s.png"%(model1[i]),dpi=100,bbox_inches='tight',pad_inches=0.1)
            plt.close()
            
        
        
        
            
            ###########################
            ##Percentage Frequency distribution
            #h = np.histogram(df25['leakvalue'],bins = ()) 
            
            h=np.histogram(df25['percentage'],bins=(0,25,50,75,100,110,120,130,3000))
            bar = plt.bar(range(8),h[0],width=1,align = 'center',zorder=3)
            bar[0].set_color('darkblue')
            bar[1].set_color('darkblue')
            bar[2].set_color('darkblue')
            bar[3].set_color('darkblue')
            bar[4].set_color('darkred')
            bar[5].set_color('darkred')
            bar[6].set_color('darkred')
            bar[7].set_color('darkred')
            
            xlab=['0-25%','25-50%','50-75%','75-100%','100-110%','110-120%','120-130%','\n'.join(wrap('130% and above', 10))]
            plt.xticks(np.arange(0,8,1),xlab)
            #majorLocator = MultipleLocator(1)
            #majorFormatter = FormatStrFormatter('%d')
            minor_ticks = list(np.arange(-0.5,8.5,1))
            
            #ax.set_xticks(minor_ticks, minor=True)
            
            ax = plt.gca()
            #ax.spines['right'].set_visible(False)
            #ax.spines['top'].set_visible(False)
            #ax.xaxis.set_minor_locator(minorLocator)
            ax.spines['left'].set_position(('data',-0.5))
            ax.set_xticks(minor_ticks, minor=True)
            ax.grid(which='minor', axis='both', linestyle='-',alpha = 1,zorder=0)
            ax.grid(which='major', axis='y', linestyle='-',alpha = 1,zorder=0)
            plt.xlim(xmin = -0.5,xmax=7.5)
            
            #ax.get_axes_locator()
            #ax.get_xgridlines() = [np.arange(0.5,8.5,1)]
            #ax.xaxis.set_major_locator(majorLocator)
            
            #ax.locator_params()
            plt.title('%s VALVES'%(model1[i]), fontsize = 12, weight = 'bold',y = 1.08)
                        #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
                        #k = list(np.linspace(0,max(df25['percentage']),8)
                           
                        #plt.yticks(k,z)
                        #plt.yticks(bins)
            plt.xlabel('PERCENTAGE OF LEAKAGE LIMIT', fontsize = 10, weight = 'bold',labelpad = 3)
            plt.ylabel('NUMBER OF VALVES', fontsize = 10, weight = 'bold',labelpad =3)
            plt.tick_params(axis='x', which='major', labelsize=8,length = 0,pad = 7,rotation = 45)
            plt.tick_params(axis='y', which='major', labelsize=8,length = 3,pad = 5)
            plt.tick_params(axis='x', which='minor', length = 3,pad = 5)
            plt.xticks(weight = 'bold')
            plt.yticks(weight = 'bold')
            
            #yint = range(min(h[0]), int(math.ceil(max(h[0]))+1))
            from matplotlib.ticker import MaxNLocator
            #import matplotlib
            if ((np.max(h[0]) > 20)&(np.max(h[0]) < 100)): 
                plt.ylim(ymax =np.max(h[0])+2)
            elif (np.max(h[0]) > 100):
                plt.ylim(ymax =np.max(h[0])+5)
            elif (np.max(h[0]) > 300):
                plt.ylim(ymax =np.max(h[0])+10)
            ax.yaxis.set_major_locator(MaxNLocator(integer=True,steps=[1,2,5,10]))
            #ax.yaxis.set_major_locator(matplotlib.ticker.AutoLocator())
            #plt.yticks(int)
            
            
                
            new = os.path.join(final_directory, r'Frequency distribution')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'Frequency distribution Percentage_Combined')
            if not os.path.exists(new1):
                os.makedirs(new1)
            plt.tight_layout()
                        
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new1)
            fig2.savefig(P_BAR + "\\%s.png"%(model1[i]),dpi=100,bbox_inches='tight',pad_inches=0.1)
            plt.close()
            
            
            ##Bar charts        
            plt.bar(df25['ValveNumber'],df25['leakvalue'], width=0.5, color ='darkblue',label = model1[i])
              
            ###Removing  right and top spline
            ax = plt.gca()
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.autoscale(axis= 'both')
            ## adding horizontal line at Max Leak value         
            ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 2, color='maroon',label = 'Max Leak Value')             #Horizontal Line  
            
            
            ##Title, Xticks, Yticks, Xlabel, YLabel and Tick Parameters
            #plt.title(model[i], fontsize = 11, weight = 'bold')
            
            
            if (len(df25['ValveNumber']) < 21 ):
                plt.xticks(np.array(np.arange(1,len(df25['ValveNumber'])+1), 'i4'))
            else:    
                plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
            
            #plt.xticks()
            
            #np.array(np.linspace(1,len(df25['ValveNumber']),15)
            k = list(np.linspace(0,max((max(df25['leakvalue'])),df25['Max Leak'].iloc[0])+10,8))
            plt.yticks(np.array(k,'i4'), weight = 'bold')
            plt.xlabel('VALVE NUMBERS', fontsize = 11, weight = 'bold')
            plt.ylabel('LEAK VALUE', fontsize = 11, weight = 'bold')
            plt.xticks(weight = 'bold')
            plt.tick_params(axis='both', which='major', labelsize=10)
            lgd = plt.legend(bbox_to_anchor=(0.1, 1.02, 0.8, .102),fancybox=True,mode = "expand",ncol = 2,borderaxespad=0., fontsize = 12,prop=dict(weight='bold'))
    #ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.05),
    #          ncol=3, fancybox=True, shadow=True)
    
            
            
    ###################################
    #Create folder to save
            

            new = os.path.join(final_directory, r'BAR')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'BAR_COMBINED')
            if not os.path.exists(new1):
                os.makedirs(new1)
               
            ##Save figure
            fig1 = plt.gcf()
            fig1.set_size_inches(6.35,4.8)
            BAR = os.path.abspath(new1)
            fig1.savefig(BAR + "\\%s.png"%(model1[i]),bbox_extra_artists=(lgd,),bbox_inches = 'tight', pad_inches = 0.2)  #,bbox_extra_artists=(lgd,),bbox_inches = 'tight', pad_inches = 0.8)
            #fig1.savefig('%s.png' %(model[i]))
            plt.close()
                  
##################################################################################        
    ######
    ###Percentage of Leakage Limit Plotting
                     
            
            plt.bar(df25['ValveNumber'],df25['percentage'], width=0.5, color ='darkblue')
           
            #z = np.array(np.linspace(0,max(df25['percentage']),8), 'i4')
            if (np.max(df25['percentage']) < 100):
                k =100
            else:
                k = np.max(df25['percentage']) + 10
            plt.ylim(ymax=k)
            #plt.xlim(xmin=0,xmax=len(df25['ValveNumber']+1))
            ###Removing  right and top spline
            ax1 = plt.gca()
            ax1.spines['right'].set_visible(False)
            ax1.spines['top'].set_visible(False)
            ax1.yaxis.set_major_locator(ticker.MaxNLocator(integer=True,steps=[1,2,5,10],min_n_ticks=10))
            ax1.yaxis.set_major_formatter(ticker.FormatStrFormatter("%d%%"))
            #ax1.xaxis.set_major_formatter()
            #ax1.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            #ax1.xaxis.set_major_locator(ticker.LinearLocator())
            
            plt.title(model1[i], fontsize = 11, weight = 'bold')
            if (len(df25['ValveNumber']) <22):
                plt.xticks(np.array(np.arange(1,len(df25['ValveNumber'])+1), 'i4'))
            else:
                plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
            #k = list(np.linspace(0,max(df25['percentage']),8))
            #print(type(loc()))
                
            
            plt.xlabel('VALVE NUMBERS', fontsize = 11, weight = 'bold')
            plt.ylabel('PERCENTAGE OF LEAKAGE LIMIT', fontsize = 11, weight = 'bold')
            plt.tick_params(axis='both', which='major', labelsize=10)
            plt.xticks(weight = 'bold')
            plt.yticks(weight='bold')
            
            new = os.path.join(final_directory, r'Percent BAR')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'Percent BAR_COMBINED')
            if not os.path.exists(new1):
                os.makedirs(new1)
            
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new1)
            fig2.savefig(P_BAR + "\\%s.png"%(model1[i]),bbox_inches='tight',pad=0.2)
            plt.close()
            
            
            
            
            
            
            
            
            
            
            
            
################################################
######
#####Seperate
################################################            
            
        
        
if (str(Input_variables.cell(1,4).value) == 'YES'):
    for i in range(0,len(model)):
        x = df1[df1['Material description'] == model[i]]
        if (x.shape[0] != 0):
            y = x.iloc[:,12:36]                            
            y = y.dropna(axis = 0, how = 'all') 
            y = y.dropna(axis = 1, how = 'all')
            z = list(y.values.flatten())                    
            Leakvalue = [nn for nn in z if nn == nn]                      ##this will remove NA values
            df25 = pd.DataFrame(Leakvalue)
            df25['Max Leak'] = x.iloc[0]['Max Leak']                      ##add Max Leak column
            df25.columns.values[0] = 'leakvalue'                          ##first column name as leakvalue
            df25 = pd.DataFrame(df25,dtype=float)
            df25['percentage'] = (df25['leakvalue']/df25['Max Leak'])*100
            df25 = df25.round(2)
            ##df25 = df25.round({'percentage':2})
            #df25['percent']  = df25['percentage'].astype(str) + '%'
            df25.index += 1                                           ### start index from 1
                                                                      ### add dataset into dictionary   
            df25['ValveNumber'] =  df25.index                         ### add one column as valvenumber
          
            #######
            ##Start Plotting, Valve numbers vs Leak Value
            
            
            a = str.split(str(model[i]))
        
            sum1 = a[0] + ' '
            for j in range(1,len(a)):
                if (j != (len(a)-1)):
                    sum1 = sum1 + a[j] + ' ' 
                else:
                    sum1 = sum1 +a[j]
          
            sum2 = a[0][0]
            for j in range(1,len(a)):
                if (j != (len(a)-1)):
                    sum2 = sum2 + a[j][0]
                else:
                    sum2 = sum2+ ' '  +a[j]
                    
              
            ############################
            ###Frequency distribution
            l=[]
            k = [0,25,50,75,100,110,120,130,3000]
            for i in range(0,len(k)):
                if (i < 5):
                    l.append(int(math.floor(x.iloc[0]['Max Leak']*k[i]/100)))
                else:
                    l.append(int(math.ceil(x.iloc[0]['Max Leak']*k[i]/100)))
            xlab =['%s-%s'%(l[0],l[1]),'%s-%s'%(l[1]+1,l[2]),'%s-%s'%(l[2]+1,l[3]),'%s-%s'%(l[3]+1,l[4]),'%s-%s'%(l[4]+1,l[5]),'%s-%s'%(l[5]+1,l[6]),'%s-%s'%(l[6]+1,l[7]),'\n'.join(wrap('%s and above'%(l[7]), 10))]
            
            h=np.histogram(df25['leakvalue'],bins=l)
            bar = plt.bar(range(8),h[0],width=1,align = 'center',zorder=3)
            bar[0].set_color('darkblue')
            bar[1].set_color('darkblue')
            bar[2].set_color('darkblue')
            bar[3].set_color('darkblue')
            bar[4].set_color('darkred')
            bar[5].set_color('darkred')
            bar[6].set_color('darkred')
            bar[7].set_color('darkred')
            
            ax = plt.gca()
            plt.xticks(np.arange(0,8,1),xlab)
            minor_ticks = list(np.arange(-0.5,8.5,1))
            ax.spines['left'].set_position(('data',-0.5))
            ax.set_xticks(minor_ticks, minor=True)
            ax.grid(which='minor', axis='both', linestyle='-',alpha = 1,zorder=0)
            ax.grid(which='major', axis='y', linestyle='-',alpha = 1,zorder=0)
            plt.xlim(xmin = -0.5,xmax=7.5)
            
            plt.title(sum2, fontsize = 12, weight = 'bold',y = 1.08)
            plt.xlabel('LEAK VALUE RANGE', fontsize = 10, weight = 'bold',labelpad = 3)
            plt.ylabel('NUMBER OF VALVES', fontsize = 10, weight = 'bold',labelpad =3)
            plt.tick_params(axis='x', which='major', labelsize=8,length = 0,pad = 7)
            plt.tick_params(axis='y', which='major', labelsize=8,length = 3,pad = 5)
            plt.tick_params(axis='x', which='minor', length = 3,pad = 5)
            plt.xticks(weight = 'bold')
            plt.yticks(weight = 'bold')
            
        
            from matplotlib.ticker import MaxNLocator
            
            if ((np.max(h[0]) > 20)&(np.max(h[0]) < 100)): 
                plt.ylim(ymax =np.max(h[0])+2)
            elif (np.max(h[0]) > 100):
                plt.ylim(ymax =np.max(h[0])+5)
            elif (np.max(h[0]) > 300):
                plt.ylim(ymax =np.max(h[0])+10)
            
            ax.yaxis.set_major_locator(MaxNLocator(integer=True,steps=[1,2,5,10]))
            
            
            
                
            new = os.path.join(final_directory, r'Frequency distribution')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'Frequency Leak Value')
            if not os.path.exists(new1):
                os.makedirs(new1)
            plt.tight_layout()
                        
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new1)
            fig2.savefig(P_BAR + "\\%s.png"%(sum2),dpi=100,bbox_inches='tight',pad_inches=0.1)
            plt.close()
            
        
        
        
            
            ###########################
            ##Percentage Frequency distribution
            #h = np.histogram(df25['leakvalue'],bins = ()) 
            
            h=np.histogram(df25['percentage'],bins=(0,25,50,75,100,110,120,130,3000))
            bar = plt.bar(range(8),h[0],width=1,align = 'center',zorder=3)
            bar[0].set_color('darkblue')
            bar[1].set_color('darkblue')
            bar[2].set_color('darkblue')
            bar[3].set_color('darkblue')
            bar[4].set_color('darkred')
            bar[5].set_color('darkred')
            bar[6].set_color('darkred')
            bar[7].set_color('darkred')
            
            xlab=['0-25%','25-50%','50-75%','75-100%','100-110%','110-120%','120-130%','\n'.join(wrap('130% and above', 10))]
            plt.xticks(np.arange(0,8,1),xlab)
            #majorLocator = MultipleLocator(1)
            #majorFormatter = FormatStrFormatter('%d')
            minor_ticks = list(np.arange(-0.5,8.5,1))
            
            #ax.set_xticks(minor_ticks, minor=True)
            
            ax = plt.gca()
            #ax.spines['right'].set_visible(False)
            #ax.spines['top'].set_visible(False)
            #ax.xaxis.set_minor_locator(minorLocator)
            ax.spines['left'].set_position(('data',-0.5))
            ax.set_xticks(minor_ticks, minor=True)
            ax.grid(which='minor', axis='both', linestyle='-',alpha = 1,zorder=0)
            ax.grid(which='major', axis='y', linestyle='-',alpha = 1,zorder=0)
            plt.xlim(xmin = -0.5,xmax=7.5)
            
            #ax.get_axes_locator()
            #ax.get_xgridlines() = [np.arange(0.5,8.5,1)]
            #ax.xaxis.set_major_locator(majorLocator)
            
            #ax.locator_params()
            plt.title(sum2, fontsize = 12, weight = 'bold',y = 1.08)
                        #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
                        #k = list(np.linspace(0,max(df25['percentage']),8)
                           
                        #plt.yticks(k,z)
                        #plt.yticks(bins)
            plt.xlabel('PERCENTAGE OF LEAKAGE LIMIT', fontsize = 10, weight = 'bold',labelpad = 3)
            plt.ylabel('NUMBER OF VALVES', fontsize = 10, weight = 'bold',labelpad =3)
            plt.tick_params(axis='x', which='major', labelsize=8,length = 0,pad = 7,rotation = 45)
            plt.tick_params(axis='y', which='major', labelsize=8,length = 3,pad = 5)
            plt.tick_params(axis='x', which='minor', length = 3,pad = 5)
            plt.xticks(weight = 'bold')
            plt.yticks(weight = 'bold')
            
            #yint = range(min(h[0]), int(math.ceil(max(h[0]))+1))
            from matplotlib.ticker import MaxNLocator
            #import matplotlib
            if ((np.max(h[0]) > 20)&(np.max(h[0]) < 100)): 
                plt.ylim(ymax =np.max(h[0])+2)
            elif (np.max(h[0]) > 100):
                plt.ylim(ymax =np.max(h[0])+5)
            elif (np.max(h[0]) > 300):
                plt.ylim(ymax =np.max(h[0])+10)
            ax.yaxis.set_major_locator(MaxNLocator(integer=True,steps=[1,2,5,10]))
            #ax.yaxis.set_major_locator(matplotlib.ticker.AutoLocator())
            #plt.yticks(int)
            
            
                
            new = os.path.join(final_directory, r'Frequency distribution')
            if not os.path.exists(new):
                os.makedirs(new)
            new1 = os.path.join(new, r'Frequency distribution Percentage')
            if not os.path.exists(new1):
                os.makedirs(new1)
            plt.tight_layout()
                        
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new1)
            fig2.savefig(P_BAR + "\\%s.png"%(sum2),dpi=100,bbox_inches='tight',pad_inches=0.1)
            plt.close()
            
            
######################################################################################################################################3
##################
#############
           ##Bar charts        
            plt.bar(df25['ValveNumber'],df25['leakvalue'], width=0.5, color ='darkblue',label = sum2)
              
            ###Removing  right and top spline
            ax = plt.gca()
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.autoscale(axis= 'both')
            ## adding horizontal line at Max Leak value         
            ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 2, color='maroon',label = 'Max Leak Value')             #Horizontal Line  
            
            
            ##Title, Xticks, Yticks, Xlabel, YLabel and Tick Parameters
            #plt.title(model[i], fontsize = 11, weight = 'bold')
            
            
            if (len(df25['ValveNumber']) < 21 ):
                plt.xticks(np.array(np.arange(1,len(df25['ValveNumber'])+1), 'i4'))
            else:    
                plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
            
            #plt.xticks()
            
            #np.array(np.linspace(1,len(df25['ValveNumber']),15)
            k = list(np.linspace(0,max((max(df25['leakvalue'])),df25['Max Leak'].iloc[0])+10,8))
            plt.yticks(np.array(k,'i4'), weight = 'bold')
            plt.xlabel('VALVE NUMBERS', fontsize = 11, weight = 'bold')
            plt.ylabel('LEAK VALUE', fontsize = 11, weight = 'bold')
            plt.xticks(weight = 'bold')
            plt.tick_params(axis='both', which='major', labelsize=10)
            lgd = plt.legend(bbox_to_anchor=(0.1, 1.02, 0.8, .102),fancybox=True,mode = "expand",ncol = 2,borderaxespad=0., fontsize = 12,prop=dict(weight='bold'))
    #ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.05),
    #          ncol=3, fancybox=True, shadow=True)
    
            
            
    ###################################
    #Create folder to save
            

            new = os.path.join(final_directory, r'BAR')
            if not os.path.exists(new):
                os.makedirs(new)
               
            ##Save figure
            fig1 = plt.gcf()
            fig1.set_size_inches(6.35,4.8)
            BAR = os.path.abspath(new)
            fig1.savefig(BAR + "\%s.png"%(sum2),bbox_extra_artists=(lgd,),bbox_inches = 'tight', pad_inches = 0.2)  #,bbox_extra_artists=(lgd,),bbox_inches = 'tight', pad_inches = 0.8)
            #fig1.savefig('%s.png' %(model[i]))
            plt.close()
                  
##################################################################################        
    ######
    ###Percentage of Leakage Limit Plotting
                     
            
            plt.bar(df25['ValveNumber'],df25['percentage'], width=0.5, color ='darkblue')
           
            #z = np.array(np.linspace(0,max(df25['percentage']),8), 'i4')
            if (np.max(df25['percentage']) < 100):
                k =100
            else:
                k = np.max(df25['percentage']) +10
            plt.ylim(ymax=k)
            #plt.xlim(xmin=0,xmax=len(df25['ValveNumber']+1))
            ###Removing  right and top spline
            ax1 = plt.gca()
            ax1.spines['right'].set_visible(False)
            ax1.spines['top'].set_visible(False)
            ax1.yaxis.set_major_locator(ticker.MaxNLocator(integer=True,steps=[1,2,5,10],min_n_ticks=10))
            ax1.yaxis.set_major_formatter(ticker.FormatStrFormatter("%d%%"))
            #ax1.xaxis.set_major_formatter()
            #ax1.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            #ax1.xaxis.set_major_locator(ticker.LinearLocator())
            
            plt.title(sum2, fontsize = 11, weight = 'bold')
            if (len(df25['ValveNumber']) <22):
                plt.xticks(np.array(np.arange(1,len(df25['ValveNumber'])+1), 'i4'))
            else:
                plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
            #k = list(np.linspace(0,max(df25['percentage']),8))
            #print(type(loc()))
                
            
            plt.xlabel('VALVE NUMBERS', fontsize = 11, weight = 'bold')
            plt.ylabel('PERCENTAGE OF LEAKAGE LIMIT', fontsize = 11, weight = 'bold')
            plt.tick_params(axis='both', which='major', labelsize=10)
            plt.xticks(weight = 'bold')
            plt.yticks(weight='bold')
            
            new = os.path.join(final_directory, r'Percent BAR')
            if not os.path.exists(new):
                os.makedirs(new)
            
            ##Save figure
            fig2 = plt.gcf()
            fig2.set_size_inches(6.35,4.8)
            P_BAR = os.path.abspath(new)
            fig2.savefig(P_BAR + "\\%s.png"%(sum2),bbox_inches='tight',pad=0.2)
            plt.close()
            
################################################################################
################################################################################
##suction
################################################################################
################################################################################            
        
suction = df100[df100['Model'].str.contains('SUCTION')]
discharge = df100[df100['Model'].str.contains('DISCHARGE')]  


################################################################################
################################################################################
###Seperate
################################################################################
################################################################################



model = list(suction['Model'])
suc = pd.DataFrame({'dummy':np.arange(1,10000)}) 
maxleak = []
for i in range(0,len(model)):
    x = df1[df1['Material description'] == model[i]] 
    if (x.shape[0] != 0):
        y = x.iloc[:,12:36]                            
        y = y.dropna(axis = 0, how = 'all') 
        y = y.dropna(axis = 1, how = 'all')
        z = list(y.values.flatten())                    
        Leakvalue = [nn for nn in z if nn == nn]                      ##this will remove NA values
        df25 = pd.DataFrame(Leakvalue)
        df25['Max Leak'] = x.iloc[0]['Max Leak']                      ##add Max Leak column
        df25.columns.values[0] = 'leakvalue'                          ##first column name as leakvalue
        df25 = pd.DataFrame(df25,dtype=float)
        df25.index += 1                                               ###start index from 1
                                                                      ### add dataset into dictionary   
        df25['ValveNumber'] =  df25.index    
        df25 = df25.round(2)
        a = str.split(str(model[i]))
        if(len(a) > 3):
            suc['ASSY %s'%(a[len(a)-1])] = df25['leakvalue']
        else:
            suc['%s'%(a[len(a)-1])] = df25['leakvalue']
        maxleak.append(x.iloc[0]['Max Leak']) 

suc = suc.drop('dummy', axis = 1)
suc = suc.dropna(axis = 0, how = 'all') 
suc = suc.dropna(axis = 1, how = 'all')
suc_d = suc.describe()      
suc_j = suc_d.T                                                       #Transponse
suc_j['MaxLeak'] = maxleak
suc_j
##################################################
##Average Suction 
##################################################
if(str(Input_variables.cell(2,4).value) == 'YES'):     
    # set width of bar
    barWidth = 0.4
     
    # Set position of bar on X axis
    r1 = np.arange(len(suc_j['mean']))
    r2 = [x11 + barWidth for x11 in r1] 
    
    # Make the plot
    l = "Average Measured Leakage Rate"
    l1 = 'Allowable Leakage Rate'
    label = '\n'.join(wrap(l, 20))
    label1 = '\n'.join(wrap(l1, 20))
    plt.bar(r1, suc_j['mean'], color='darkblue', width=barWidth, edgecolor='white', label=label)
    plt.bar(r2,suc_j['MaxLeak'], color='maroon', width=barWidth, edgecolor='white', label=label1)
    
     
    # Add xticks on the middle of the group bars
    
    plt.xticks([r + barWidth for r in range(len(suc_j['mean']))], suc_j.index.values, rotation = 90,weight = 'bold')
    #k = (math.ceil(np.max(np.max(suc_j[['mean','MaxLeak']]))/25.0)*25)+25
    k = np.max(np.max(suc_j[['mean','MaxLeak']]))+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')
    #plt.yticks(np.array(np.arange(0,450,25), 'i4')) 
    # Create legend & Show graphic
    
    
    
    ax4 = plt.gca()
    ax4.spines['right'].set_visible(False)
    ax4.spines['top'].set_visible(False)
    #ax4.grid(which='major', axis='y', linestyle='-',alpha = 0.4)  
    
            ## adding horizontal line at Max Leak value         
            #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
            
            
            
    plt.title('SUCTION VALVES', fontsize = 16, weight = 'bold')
    #plt.ylabel('AVERAGE LEAKAGE VALUE', fontsize = 14, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=11)
    lgd = plt.legend(bbox_to_anchor=(1.15, 0.5), loc=2, borderaxespad=0.,  fontsize = 10)
    
    ax4.set_ylabel('AVERAGE LEAKAGE VALUE', fontsize = 12, weight = 'bold')
    ax4.set_yticks(np.array(np.arange(0,k,50), 'i4'))
    
    
    ax21 = ax4.twinx()
    
    ax21.set_ylabel('MAX LEAK VALUE',fontsize = 12, weight = 'bold')
    ax21.set_yticks(np.array(np.arange(0,k,50), 'i4'))
    plt.yticks(weight='bold')
    
    ##Save figure
    fig4 = plt.gcf()
    fig4.set_size_inches(6.5, 4.5)
    #plt.figure(figsize = (20,20))
    Avg = os.path.abspath(final_directory)
    fig4.savefig(Avg + "\\average_suction.png" ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4)
    plt.close()
    
    
###############################################################################
##Box Plot Suction
#################################################################################    
    
    
if(str(Input_variables.cell(4,4).value) == 'YES'): 
        
            #df = pd.DataFrame(data = np.random.random(size=(4,4)), columns = ['A','B','C','D'])
    
    sns.boxplot(x="variable", y="value", data=pd.melt(suc))
    plt.scatter(range(0,len(suc_j.index)) ,suc_j['MaxLeak'],marker ='_' ,s = 100, color = 'maroon')
    
    #plt.boxplot(suc)
    ax5 = plt.gca()
    ax5.spines['right'].set_visible(False)
    ax5.spines['top'].set_visible(False)
    ax5.grid(which='major', axis='y', linestyle='-', alpha = 0.4)  
            
    plt.title('MEASURED LEAKAGE RATE - SV', fontsize = 26, weight = 'bold')
                    #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
                    
    k = np.max(np.max(suc))+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                
    plt.xlabel('VALVE SIZE', fontsize = 20, weight = 'bold')
    plt.ylabel('LEAK VALUE', fontsize = 20, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=13)
    plt.xticks(rotation = 90,weight = 'bold')
    
    
    #plt.tick_params(axis='both', which='major', labelsize=16,)
    #plt.xticks(rotation = 90)
    #plt.set_marker('hline')
    
    
    lgd = plt.legend(bbox_to_anchor=(1.03, 0.5), loc=2, borderaxespad=0.,  fontsize = 16)
                    
    ##Save figure
    fig5 = plt.gcf()
    fig5.set_size_inches(12,12)
    Avg1 = os.path.abspath(final_directory)
    fig5.savefig(Avg1 + '\\boxplot_SV.png',bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4)
    #plt.figure(figsize = (20,20))
    plt.close()

     
##################################################################################################]
##########
##########Max_min_std+_std-_average
##########
############################################################################################### 
if(str(Input_variables.cell(6,4).value) == 'YES'): 

    suc_j['std+'] = suc_j['mean']+suc_j['std']
    suc_j['std-'] = suc_j['mean']-suc_j['std']
    #suc_j.loc['98CP']
    w = suc_j.index
    plt.xticks(range(len(suc_j['MaxLeak'])),suc_j.index)
    plt.scatter(np.arange(0,len(suc_j.index)) ,suc_j['mean'],marker ='o',color ='black' ,s = 35)   
    plt.scatter(np.arange(0,len(suc_j.index)) ,suc_j['min'],marker ='o' ,s = 12,color = 'black') 
    plt.scatter(np.arange(0,len(suc_j.index)) ,suc_j['max'],marker ='o' ,s = 12, color = 'black') 
    plt.scatter(np.arange(0,len(suc_j.index)) ,suc_j['std+'],marker ='_' ,s = 30,color = 'black' ) 
    plt.scatter(np.arange(0,len(suc_j.index)) ,suc_j['std-'],marker ='_' ,s = 30, color = 'black') 
    plt.scatter(np.arange(0,len(suc_j.index)),suc_j['MaxLeak'],marker ='_' ,s = 50,  color = 'maroon')
    
    suc_k = suc_j.drop(['count','std'],axis =1)
    for i in range(0,len(suc_j.index)):
        x111 = [suc_k.index[i],suc_k.index[i]]  
        y111 = [max(suc_k.iloc[i]),min(suc_k.iloc[i])]
    
        plt.plot(x111,y111,linewidth = 0.6,color = 'black')
    
    ax6 = plt.gca()
    ax6.spines['right'].set_visible(False)
    ax6.spines['top'].set_visible(False)
    ax6.grid(which='major', axis='y', linestyle='-',alpha = 0.6)  
    
            ## adding horizontal line at Max Leak value         
            #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
            
    plt.title('MEASURED LEAKAGE RATE - SV', fontsize = 16, weight = 'bold')
                    #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
     
    k = np.max(suc_k.values)+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                
    
    #plt.xticks(suc_j.index)
    plt.xlabel('VALVE SIZE', fontsize = 14, weight = 'bold')
    plt.ylabel('LEAK VALUE', fontsize = 14, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=12)
    plt.xticks(rotation = 90,weight = 'bold')
    
    lgd = plt.legend(bbox_to_anchor=(1.03, 0.5), loc=2, borderaxespad=0.,  fontsize = 10)
    
    #plt.scatter(range(0,len(suc_j.index)) ,suc_j['MaxLeak'],marker ='_' ,s = 300)
    #plt.tick_params(axis='both', which='major', labelsize=16,)
    #plt.xticks(rotation = 90)
    #plt.set_marker('hline')
    
                    
    ##Save figure
    fig6 = plt.gcf()
    fig6.set_size_inches(9,5)
                    #P_BAR = os.path.abspath('Percent BAR')
    Avg = os.path.abspath(final_directory)                
    fig6.savefig(Avg + '\\Measured Leakage Rate - SV.png' ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)
    plt.close()

###########################################################################################################################
###########################################################################################################################
###Discharge valve
###########################################################################################################################
###########################################################################################################################

discharge = df100[df100['Model'].str.contains('DISCHARGE')]  


model = list(discharge['Model'])
dis = pd.DataFrame({'dummy':np.arange(1,10000)}) 
maxleak = []
for i in range(0,len(model)):
    x = df1[df1['Material description'] == model[i]]
    if (x.shape[0] != 0):
        y = x.iloc[:,12:36]                            
        y = y.dropna(axis = 0, how = 'all') 
        y = y.dropna(axis = 1, how = 'all')
        z = list(y.values.flatten())                    
        Leakvalue = [nn for nn in z if nn == nn]                       ##this will remove NA values
        df25 = pd.DataFrame(Leakvalue)
        df25['Max Leak'] = x.iloc[0]['Max Leak']                       ##add Max Leak column
        df25.columns.values[0] = 'leakvalue'                           ##first column name as leakvalue
        df25 = pd.DataFrame(df25,dtype=float)
        df25.index += 1                                                ### start index from 1
                                                                       ### add dataset into dictionary   
        df25['ValveNumber'] =  df25.index    
        df25 = df25.round(2)
        a = str.split(str(model[i]))
        dis['%s'%(a[len(a)-1])] = df25['leakvalue']
        maxleak.append(x.iloc[0]['Max Leak']) 
        


dis = dis.drop('dummy', axis = 1)
dis = dis.dropna(axis = 0, how = 'all') 
dis = dis.dropna(axis = 1, how = 'all')
dis_d = dis.describe()      
dis_j = dis_d.T                                           #Transponse
dis_j['MaxLeak'] = maxleak


##################################################
##Average discharge
##################################################
if(str(Input_variables.cell(3,4).value) == 'YES'):      
    # set width of bar
    barWidth = 0.4
      
    # Set position of bar on X axis
    r1 = np.arange(len(dis_j['mean']))
    r2 = [x11 + barWidth for x11 in r1] 
    
    #Label on right side in center with wrap text
    l = "Average Measured Leakage Rate"
    l1 = 'Allowable Leakage Rate'
    label = '\n'.join(wrap(l, 20))
    label1 = '\n'.join(wrap(l1, 20))
    
    
    
    # Make the plot
    plt.bar(r1, dis_j['mean'], color='darkblue', width=barWidth, edgecolor='white', label=label)
    plt.bar(r2,dis_j['MaxLeak'], color='maroon', width=barWidth, edgecolor='white', label=label1)
    
    # Add xticks on the middle of the group bars
    plt.xticks([r + barWidth for r in range(len(dis_j['mean']))], dis_j.index.values, rotation = 90,weight = 'bold')
    k = np.max(np.max(dis_j[['mean','MaxLeak']]))+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                
    
    #plt.yticks(np.array(np.arange(0,450,25), 'i4')) 
    # Create legend & Show graphic
    ax7 = plt.gca()
    ax7.spines['right'].set_visible(False)
    ax7.spines['top'].set_visible(False)
    #ax7.grid(which='major', axis='y', linestyle='-',alpha = 0.5)  
    
            ## adding horizontal line at Max Leak value         
            #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
    plt.title('DISCHARGE VALVES', fontsize = 14, weight = 'bold')
    
           
    ax7.set_ylabel('AVERAGE LEAKAGE VALUE', fontsize = 12, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=11)
    lgd = plt.legend(bbox_to_anchor=(1.15, 0.5), loc=2, borderaxespad=0.,  fontsize = 10)
    
    ax22 = ax7.twinx()
    
    ax22.set_ylabel('MAX LEAK VALUE',fontsize = 12, weight = 'bold')
    ax22.set_yticks(np.array(np.arange(0,k,50), 'i4'))
    plt.yticks(weight='bold')
            
    ##Save figure
    fig7 = plt.gcf()
    fig7.set_size_inches(6.5,4.5)
    Avg_dis = os.path.abspath(final_directory)
    fig7.savefig(Avg_dis+'\\average_discharge.png',bbox_extra_artists=(lgd,),bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)
    plt.close()
###############################################################################
##Box Plot discharge
####################################################################    
    
if(str(Input_variables.cell(5,4).value) == 'YES'):     

        
            #df = pd.DataFrame(data = np.random.random(size=(4,4)), columns = ['A','B','C','D'])
    
    sns.boxplot(x="variable", y="value", data=pd.melt(dis))
    plt.scatter(range(0,len(dis_j.index)) ,dis_j['MaxLeak'],marker ='_' , s = 100,color = 'maroon')
    #plt.boxplot(suc)
    ax8 = plt.gca()
    ax8.spines['right'].set_visible(False)
    ax8.spines['top'].set_visible(False)
    ax8.grid(which='major', axis='y', linestyle='-', alpha = 0.4) 
            ## adding horizontal line at Max Leak value         
            #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
            
    plt.title('MEASURED LEAKAGE RATE - DV', fontsize = 26, weight = 'bold')
                    #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
    k = np.max(np.max(dis))+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                
    
    plt.xlabel('VALVE SIZE', fontsize = 20, weight = 'bold')
    plt.ylabel('LEAK VALUE', fontsize = 20, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=13)
    plt.xticks(rotation = 90,weight ='bold')
    

    #plt.tick_params(axis='both', which='major', labelsize=16,)
    #plt.xticks(rotation = 90)
    #plt.set_marker('hline')
   
                    
                    ##Save figure
    fig8 = plt.gcf()
                                                         #
                                                         #P_BAR = os.path.abspath('Percent BAR')
    fig8.set_size_inches(12,12)
    Avg1 = os.path.abspath(final_directory)
    fig8.savefig(Avg1 + '\\boxplot_DV.png',bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)
    plt.close()


#suc.shape

#plt.show()
#fig = plt.gcf()
                #
                #P_BAR = os.path.abspath('Percent BAR')

#plt.figure(figsize = (30,30))  
       
#fig.savefig('xxx1.png')        
        
#x = df1[df1['Material description'].str.contains('54CPS')] == model[0]]   
                
######################################################################################################
##########################Statistical Average_mean_std plot
###################################################################################################                
if(str(Input_variables.cell(7,4).value) == 'YES'):   
    
    dis_j['std+'] = dis_j['mean']+dis_j['std']
    dis_j['std-'] = dis_j['mean']-dis_j['std']
    plt.xticks(range(len(dis_j['MaxLeak'])),dis_j.index)
    plt.scatter(np.arange(0,len(dis_j.index)) ,dis_j['mean'],marker ='o',color ='black' ,s = 35)   
    plt.scatter(np.arange(0,len(dis_j.index)) ,dis_j['min'],marker ='o' ,s = 12,color = 'black') 
    plt.scatter(np.arange(0,len(dis_j.index)) ,dis_j['max'],marker ='o' ,s = 12, color = 'black') 
    plt.scatter(np.arange(0,len(dis_j.index)) ,dis_j['std+'],marker ='_' ,s = 30,color = 'black' ) 
    plt.scatter(np.arange(0,len(dis_j.index)) ,dis_j['std-'],marker ='_' ,s = 30, color = 'black') 
    plt.scatter(np.arange(0,len(dis_j.index)),dis_j['MaxLeak'],marker ='_' ,s = 50,  color = 'maroon')
    
    dis_k = dis_j.drop(['count','std'],axis =1)
    for i in range(0,len(dis_j.index)):
        x111 = [dis_k.index[i],dis_k.index[i]]  
        y111 = [max(dis_k.iloc[i]),min(dis_k.iloc[i])]
        plt.plot(x111,y111,linewidth = 0.6,color = 'black')
    
    ax9 = plt.gca()
    ax9.spines['right'].set_visible(False)
    ax9.spines['top'].set_visible(False)
    ax9.grid(which='major', axis='y', linestyle='-',alpha = 0.6)
    
            ## adding horizontal line at Max Leak value         
            #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
            
    plt.title('MEASURED LEAKAGE RATE - DV', fontsize = 16, weight = 'bold')
                    #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
    
    
    k = np.max(dis_k.values)+50
    plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                  
    
    #plt.xticks(suc_j.index)
    plt.xlabel('VALVE SIZE', fontsize = 14, weight = 'bold')
    plt.ylabel('LEAK VALUE', fontsize = 14, weight = 'bold')
    plt.tick_params(axis='both', which='major', labelsize=12)
    plt.xticks(rotation = 90,weight = 'bold')
    
    #plt.scatter(range(0,len(suc_j.index)) ,suc_j['MaxLeak'],marker ='_' ,s = 300)
    #plt.tick_params(axis='both', which='major', labelsize=16,)
    #plt.xticks(rotation = 90)
    #plt.set_marker('hline')
    lgd = plt.legend(bbox_to_anchor=(1.03, 0.5), loc=2, borderaxespad=0.,  fontsize = 10) 
           
                    ##Save figure
  
  
    fig9 = plt.gcf()
    fig9.set_size_inches(9, 5)
                    #P_BAR = os.path.abspath('Percent BAR')
    Avg = os.path.abspath(final_directory)                
    fig9.savefig(Avg + '\\Measured Leakage Rate - DV.png' ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)
    plt.close()
    

##########################################################################################
##############
##############    
if (str(Input_variables.cell(9,4).value) == 'YES'):    
    suc1 = pd.DataFrame(suc_j)
    dis1 = pd.DataFrame(dis_j)
    com = pd.merge(suc1,dis1,how = 'outer', right_index = True ,left_index = True)
    com = com.fillna(0)
    mean = []
    MaxLeak = []
    std =[]
    Max = []
    Min = []
    for i in range(0,com.shape[0]):
        if((com['mean_x'][i] == 0) |(com['mean_y'][i] == 0)):
            mean.append (com['mean_x'][i]+com['mean_y'][i])
            MaxLeak.append(com['MaxLeak_x'][i]+com['MaxLeak_y'][i])
            std.append(com['std_x'][i]+com['std_y'][i])
            Max.append(com['max_x'][i]+com['max_y'][i])
            Min.append(com['min_x'][i]+com['min_y'][i])
        else:
            mean.append((com['mean_x'][i]+com['mean_y'][i])/2)
            MaxLeak.append((com['MaxLeak_x'][i]+com['MaxLeak_y'][i])/2)
            std.append((com['std_x'][i]+com['std_y'][i])/2)
            Max.append((com['max_x'][i]+com['max_y'][i])/2)
            Min.append((com['min_x'][i]+com['min_y'][i])/2)
    com['mean'] = mean
    com['MaxLeak'] = MaxLeak
    com['Max'] = Max
    com['Min'] = Min
    com['std'] = std
    com['mean'] = com['mean'].round(0)
    com = com.sort_values(by = 'mean')

############
##Average combined   

    if(str(Input_variables.cell(10,4).value) == 'YES'):     
        # set width of bar
        barWidth = 0.4
         
        # Set position of bar on X axis
        r1 = np.arange(len(com['mean']))
        r2 = [x11 + barWidth for x11 in r1] 
        
        # Make the plot
        l = "Average Measured Leakage Rate"
        l1 = 'Allowable Leakage Rate'
        label = '\n'.join(wrap(l, 20))
        label1 = '\n'.join(wrap(l1, 20))
        plt.bar(r1, com['mean'], color='darkblue', width=barWidth, edgecolor='white', label=label)
        plt.bar(r2,com['MaxLeak'], color='maroon', width=barWidth, edgecolor='white', label=label1)
        
         
        # Add xticks on the middle of the group bars
        
        plt.xticks([r + barWidth for r in range(len(com['mean']))], com.index.values, rotation = 90,weight = 'bold')
        #k = (math.ceil(np.max(np.max(suc_j[['mean','MaxLeak']]))/25.0)*25)+25
        k = np.max(np.max(com[['mean','MaxLeak']]))+50
        
        #plt.yticks(np.array(np.arange(0,450,25), 'i4')) 
        # Create legend & Show graphic
        
        
        
        ax4 = plt.gca()
        ax4.spines['right'].set_visible(False)
        ax4.spines['top'].set_visible(False)
        #ax4.grid(which='major', axis='y', linestyle='-',alpha = 0.4)
        ax4.set_ylabel('AVERAGE LEAKAGE VALUE', fontsize = 12, weight = 'bold')
        ax4.set_yticks(np.array(np.arange(0,k,50), 'i4'))       
        plt.yticks(weight='bold')
        plt.title('VALVE LEAKAGE BY SIZE', fontsize = 14, weight = 'bold')
        
        plt.tick_params(axis='both', which='major', labelsize=11)
        lgd = plt.legend(bbox_to_anchor=(1.15, 0.5), loc=2, borderaxespad=0.,  fontsize = 10)
        
    
        ax21 = ax4.twinx()
    
        ax21.set_ylabel('MAX LEAK VALUE',fontsize = 12, weight = 'bold')
        ax21.set_yticks(np.array(np.arange(0,k,50), 'i4'))
        plt.yticks(weight='bold')
        
        ##Save figure
        fig4 = plt.gcf()
        fig4.set_size_inches(6.5,4.5)
        #plt.figure(figsize = (20,20))
        Avg = os.path.abspath(final_directory)
        fig4.savefig(Avg + "\\average_combined.png" ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4)
        plt.close()
############################
###Scatter
        
        if(str(Input_variables.cell(11,4).value) == 'YES'): 
    
            com['std+'] = com['mean']+com['std']
            com['std-'] = com['mean']-com['std']
            #suc_j.loc['98CP']
            
            plt.xticks(range(len(com['MaxLeak'])),com.index)
            plt.scatter(np.arange(0,len(com.index)) ,com['mean'],marker ='o',color ='black' ,s = 35)   
            plt.scatter(np.arange(0,len(com.index)) ,com['Min'],marker ='o' ,s = 12,color = 'black') 
            plt.scatter(np.arange(0,len(com.index)) ,com['Max'],marker ='o' ,s = 12, color = 'black') 
            plt.scatter(np.arange(0,len(com.index)) ,com['std+'],marker ='_' ,s = 30,color = 'black' ) 
            plt.scatter(np.arange(0,len(com.index)) ,com['std-'],marker ='_' ,s = 30, color = 'black') 
            plt.scatter(np.arange(0,len(com.index)),com['MaxLeak'],marker ='_' ,s = 50,  color = 'maroon')
            
            com_k = com[['mean','Max','Min','std+','std-','MaxLeak']]
            for i in range(0,len(com.index)):
                x111 = [com_k.index[i],com_k.index[i]]  
                y111 = [np.max(com_k.iloc[i]),np.min(com_k.iloc[i])]
            
                plt.plot(x111,y111,linewidth = 0.6,color = 'black')
            
            ax6 = plt.gca()
            ax6.spines['right'].set_visible(False)
            ax6.spines['top'].set_visible(False)
            ax6.grid(which='major', axis='y', linestyle='-',alpha = 0.6)  
            
                    ## adding horizontal line at Max Leak value         
                    #ax.axhline(y = x.iloc[0]['Max Leak'], linewidth = 4, color='r')             #Horizontal Line  
                    
            plt.title('VALVE LEAKAGE BY SIZE', fontsize = 16, weight = 'bold')
                            #plt.xticks(np.array(np.linspace(1,len(df25['ValveNumber']),15), 'i4'))
             
            k = np.max(com_k.values)+50
            plt.yticks(np.array(np.arange(0,k,50), 'i4'),weight = 'bold')                
            
            #plt.xticks(suc_j.index)
            plt.xlabel('VALVE SIZE', fontsize = 13, weight = 'bold')
            plt.ylabel('LEAK VALUE', fontsize = 13, weight = 'bold')
            plt.tick_params(axis='both', which='major', labelsize=11)
            plt.xticks(rotation = 90,weight = 'bold')
            
            lgd = plt.legend(bbox_to_anchor=(1.03, 0.5), loc=2, borderaxespad=0.,  fontsize = 10)
            
            #plt.scatter(range(0,len(suc_j.index)) ,suc_j['MaxLeak'],marker ='_' ,s = 300)
            #plt.tick_params(axis='both', which='major', labelsize=16,)
            #plt.xticks(rotation = 90)
            #plt.set_marker('hline')
            
                            
            ##Save figure
            fig6 = plt.gcf()
            fig6.set_size_inches(9, 5)   #(10,8)
                            #P_BAR = os.path.abspath('Percent BAR')
            Avg = os.path.abspath(final_directory)                
            fig6.savefig(Avg + '\\Measured Leakage Rate.png' ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)
            plt.close()
            
    
#####################################################################################################################
#####################################################################################################################
#####################################################################################################################
#####################################################################################################################
####Howmany Orders
#####################################################################################################################
#####################################################################################################################
#####################################################################################################################

df22 = pd.merge(SAP_Data,leak_test1,how='left',left_on = 'Order', right_on = 'Q0b')
df22 = df22[(df22['Material description'].str.contains('CP') == True)  |(df22['Material description'].str.contains('XP') == True)] 
df22 = df22.dropna(subset = ['Material description'])


#df100 = pd.read_excel('Input Sheet.xlsx', sheetname = 'Max Leak')
df2 = pd.merge(df22,df100, how='left',left_on = 'Material description', right_on = 'Model')
df2['Material description'] = df2['Material description'].str.replace('CPs','CPS')
df2 = pd.DataFrame(df2[df2['Model'].isna()!=True])
df2.reset_index(inplace = True)
df555 =pd.DataFrame(df2.drop_duplicates('Order'))


df2['Q1'].fillna(0,inplace = True)  
df555['Q1'].fillna(0,inplace=True)  
SAP = pd.DataFrame(df555.groupby('Actual finish date')['Confirmed quantity (GMEIN)'].size().rename('SAP'))  
VLT = pd.DataFrame(df2[df2['Q1'] > 0].groupby('Actual finish date').size().rename('VLT'))    
VLT1= pd.DataFrame(df555[df555['Q1'] > 0].groupby('Actual finish date').size().rename('VLT'))  
HMO = pd.merge(SAP,VLT, how = 'left', left_index = True,right_index = True)
HMO['date'] = pd.to_datetime(HMO.index, infer_datetime_format=True)
HMO['date1']= HMO['date'].dt.strftime('%m/%d/%Y')


#####################################################################################################################
######3Combined charts
##############################################################################################################
# set width of bar

if(str(Input_variables.cell(8,4).value) == 'YES'): 
    barWidth = 0.3
      
    # Set position of bar on X axis
    r1 = np.arange(len(SAP.index))
    r2 = [x11 + barWidth for x11 in r1] 
    
    # Make the plot
    plt.bar(r1, HMO['SAP'], color='darkblue', width=barWidth, edgecolor='white', label='SAP')
    plt.bar(r2,HMO['VLT'], color='maroon', width=barWidth, edgecolor='white', label='VLT')
    
    
    # Add xticks on the middle of the group bars
    
    plt.xticks([r + barWidth for r in range(len(SAP.index))], HMO['date1'].values, rotation = 90, fontsize = 10,weight='bold')
    k = np.max(np.max(HMO[['SAP','VLT']])) +10  
    plt.yticks(np.array(np.arange(0,k,10), 'i4'), fontsize = 10,weight = 'bold')                
    
    #plt.yticks(np.array(np.arange(0,450,25), 'i4')) 
    # Create legend & Show graphic

    
    ax0 = plt.gca()
    ax0.spines['right'].set_visible(False)
    ax0.spines['top'].set_visible(False)
    #ax0.grid(which='major', axis='y', linestyle='-',alpha = 0.4)
    
    
    plt.title('ORDERS COMPLETED', fontsize = 16, weight = 'bold')
    plt.ylabel('NUMBER OF ORDERS', fontsize = 14, weight = 'bold')
    plt.xlabel('DATE', fontsize = 14, weight = 'bold')
    
    #plt.tick_params(axis='both', which='major', labelsize=16)
    
    lgd = plt.legend(bbox_to_anchor=(1.03, 1), loc=2, borderaxespad=0.,  fontsize = 10)
    #ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.05),
    #          ncol=3, fancybox=True, shadow=True)
    
           
    ##Save figure
    fig0 = plt.gcf()
    fig0.set_size_inches(8,4.5)
    #plt.figure(figsize = (8,5))
 
    Avg = os.path.abspath(final_directory)                
    fig0.savefig(Avg + '\\HMO.png' ,bbox_extra_artists=(lgd,) ,bbox_inches = 'tight', pad_inches = 0.4,dpi = 100)

    plt.close()
    
    
    
###############################################################
###Write HMO to excel file
###############################################################    
HMO1 = HMO.drop(['date'],axis=1) 
HMO1.reset_index(inplace = True)   
HMO1.set_index(['date1'],inplace = True)
HMO1 = HMO1.drop(['Actual finish date'],axis=1) 
add = pd.DataFrame([[np.sum(HMO['SAP']),np.sum(HMO['VLT'])]],columns=['SAP','VLT'],index={'Total'})
HMO1 = HMO1.append(add)


######################################################################################################################
#######################Pivot Tables
######################################################################################################################

df33 = pd.DataFrame(df2)

failed = []
for i in range(0,df33.shape[0]):
    failed.append(np.sum(df33.iloc[i,12:36] > df33.loc[i,'Max Leak']))
df33['failed'] = failed    

df33['Q1'] = df33['Q1'].astype(int)
df3 = df33[['Material description',"Confirmed quantity (GMEIN)","Q1","failed"]]

#df3[["Confirmed quantity (GMEIN)","Q1","failed"]]=df3[["Confirmed quantity (GMEIN)","Q1","failed"]].astype(int)
table1 = pd.pivot_table(df555,index = ['Material description'],values=["Confirmed quantity (GMEIN)"],aggfunc=[np.sum],margins = True)
table2 = pd.pivot_table(df3,index = ['Material description'],values=["Q1","failed"],aggfunc=[np.sum],margins = True)

table = pd.merge(table1,table2, how = 'left', left_index = True,right_index = True)


table.columns =['Sum of Delivered quantity', 'Sum of valves tested','Sum of Failed']
table['% of valve tested'] = (table['Sum of valves tested']/ table['Sum of Delivered quantity'])*100      
table['% of valve failed_(Delivered quantity)'] =  (table['Sum of Failed']/ table['Sum of Delivered quantity'])*100 
table['% of valve failed_(Valve tested)'] =  (table['Sum of Failed']/ table['Sum of valves tested'])*100 
table = table.round(1)



#############################################################################################
#############################################################################################
###Create output Excel file
#############################################################################################
###########################################################################################


Avg = os.path.abspath(final_directory)

writer = pd.ExcelWriter(Avg+'\\output.xlsx', engine='xlsxwriter')

# Write your DataFrame to a file     
HMO1.to_excel(writer, 'HMO')
#VLT1.to_excel(writer,'VLT1')

table.to_excel(writer,'pivottable')     
writer.save() 







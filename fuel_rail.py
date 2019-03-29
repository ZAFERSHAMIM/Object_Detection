import tkinter as tkt
import tkinter.ttk as ttk
import warnings; warnings.simplefilter("ignore")
import math as mt
import pandas as pd
from tkinter import filedialog
import os
from tkinter import messagebox
import time

def start():
    global win
    win=tkt.Tk()
    win.configure(background="white")
    win.title("Fuel Rail AI Project")
    lblL=tkt.Label(win, text="Search for File location", font=("Comic Sanc MS", 10), anchor="w")
    lblL.grid(column=3,row=0, sticky="w")
#    w=tkt.Canvas(win,background="white")
#    w.grid(sticky="s")
    def search():
        global FLoc
        win2=tkt.Tk()
        win2.title("search")
        lblL2=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
        lblL2.grid(column=3,row=1, sticky="w")
        win2.filename= filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel Files","*.xlsx"),("all files","*.*")))
        FLoc=win2.filename
        lblL2.update()
        lblL2.configure(text=FLoc, fg="blue")
        win2.destroy()
        def openfile():
            os.startfile(FLoc)
        btno=tkt.Button(win, text="Open File", command=openfile)
        btno.grid(column=4, row=5, sticky="w")
    
    btn=tkt.Button(win, text="search", command=search)
    btn.grid(column=4, row=0)
    
    a={'200': [24.2,11.6,6.3], '250': [18,11.6,3.2], '350': [16.4,9.4,3.5]}
    labels=['OD','ID','T']
    cup=pd.DataFrame(index=labels,data=a)
    b={'200': [15.85,10,2.925], '250': [19,10,4.5], '350': [16.1,10,3.05]}
    labels=['OD','ID','T']
    psc=pd.DataFrame(index=labels,data=b)
    c={'200': [19,14,2.5], '250': [19,14,2.5], '350': [24.6,9.4,7.6]}
    labels=['OD','ID','T']
    tube=pd.DataFrame(index=labels,data=c)
    d={'200': [14,6.2,3.9], '250': [19,11.6,3.7], '350': [14,3.5,5.25]}
    labels=['OD','ID','T']
    inlet=pd.DataFrame(index=labels,data=d)
    
    lbl=tkt.Label(win, text="Choose Working \nPressure(Bar)", font=("Comic Sanc MS", 10), anchor="w")
    lbl.grid(column=0,row=0, sticky="w")
    lbl1=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl1.grid(column=0,row=1, sticky="w")
    win.geometry('1300x600')
    
    combo=ttk.Combobox(win)
    combo['values']=(200, 250, 350)
    combo.grid(column=1, row=0, sticky="w")
    
    lbl2=tkt.Label(win, text="Pmin(MPa)", font=("Comic Sanc MS", 10), anchor="w")
    lbl2.grid(column=0,row=3, sticky="w")
    lbl4=tkt.Label(win, text="Pmax(MPa)", font=("Comic Sanc MS", 10), anchor="w")
    lbl4.grid(column=1,row=3, sticky="w")
    
    txt = tkt.Entry(win,width=10)
    txt.grid(column=0, row=4, sticky="w")
    txt2 = tkt.Entry(win,width=10)
    txt2.grid(column=1, row=4, sticky="w")
    
    
    lbl3=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl3.grid(column=1,row=5, sticky="w")
    
    lbl5=tkt.Label(win, text="Surface Roughness Rz", font=("Comic Sanc MS", 10), anchor="w")
    lbl5.grid(column=0,row=6, sticky="w")
    
    lbl6=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl6.grid(column=1,row=9, sticky="w")
    lbl7=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl7.grid(column=1,row=10, sticky="w")
    lbl8=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl8.grid(column=1,row=11, sticky="w")
    lbl9=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl9.grid(column=1,row=12, sticky="w")
    lbl10=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl10.grid(column=1,row=13, sticky="w")
    lbl11=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl11.grid(column=1,row=14, sticky="w")
    lbl12=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl12.grid(column=1,row=15, sticky="w")
    lbl13=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl13.grid(column=1,row=16, sticky="w")
    lbl14=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbl14.grid(column=1,row=17, sticky="w")
    
    
    lblc6=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc6.grid(column=0,row=9, sticky="w")
    lblc7=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc7.grid(column=0,row=10, sticky="w")
    lblc8=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc8.grid(column=0,row=11, sticky="w")
    lblc9=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc9.grid(column=0,row=12, sticky="w")
    lblc10=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc10.grid(column=0,row=13, sticky="w")
    lblc11=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc11.grid(column=0,row=14, sticky="w")
    lblc12=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc12.grid(column=0,row=15, sticky="w")
    lblc13=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc13.grid(column=0,row=16, sticky="w")
    lblc14=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc14.grid(column=0,row=17, sticky="w")
    lblc15=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc15.grid(column=0,row=18, sticky="w")
    lblc16=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc16.grid(column=0,row=19, sticky="w")
    lblc17=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblc17.grid(column=0,row=20, sticky="w")
    
    
    lblp6=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp6.grid(column=2,row=9, sticky="w")
    lblp7=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp7.grid(column=2,row=10, sticky="w")
    lblp8=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp8.grid(column=2,row=11, sticky="w")
    lblp9=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp9.grid(column=2,row=12, sticky="w")
    lblp10=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp10.grid(column=2,row=13, sticky="w")
    lblp11=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp11.grid(column=2,row=14, sticky="w")
    lblp12=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp12.grid(column=2,row=15, sticky="w")
    lblp13=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp13.grid(column=2,row=16, sticky="w")
    lblp14=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblp14.grid(column=2,row=17, sticky="w")
    
    
    lbli6=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli6.grid(column=3,row=9, sticky="w")
    lbli7=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli7.grid(column=3,row=10, sticky="w")
    lbli8=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli8.grid(column=3,row=11, sticky="w")
    lbli9=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli9.grid(column=3,row=12, sticky="w")
    lbli10=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli10.grid(column=3,row=13, sticky="w")
    lbli11=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli11.grid(column=3,row=14, sticky="w")
    lbli12=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli12.grid(column=3,row=15, sticky="w")
    lbli13=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli13.grid(column=3,row=16, sticky="w")
    lbli14=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lbli14.grid(column=3,row=17, sticky="w")
    
    txt3 = tkt.Entry(win,width=10)
    txt3.grid(column=0, row=7, sticky="w")
    
    lbld=tkt.Label(win, text="Enter Tolerance value for dP", font=("Comic Sanc MS", 10), anchor="w")
    lbld.grid(column=2,row=3, sticky="w")
    
    txt4 = tkt.Entry(win,width=10)
    txt4.grid(column=2, row=4, sticky="w")
    
    lblt=tkt.Label(win, text="", font=("Comic Sanc MS", 10), anchor="w")
    lblt.grid(column=3,row=3, sticky="w")
    
    def clicked():
        wb=pd.ExcelFile(FLoc)
        sheet_1=pd.read_excel(wb, sheet_name=0)
        res1="Selected WP = " + combo.get() + " Bar"
        lbl1.configure(text=res1, fg="green")
        tmp=float(combo.get())
        fcheck1=sheet_1.iloc[0,8]
        fcheck2=sheet_1.iloc[0,10]
        if (fcheck1==2 and fcheck2==11.6):
            if (tmp==200):
                x=0
            elif (tmp==250):
                x=1
            elif (tmp==350):
                x=2
            if (int(x)==0):
                A=sheet_1.query('["200"] in WP')
                WPD=200
            elif (int(x)==1):
                A=sheet_1.query('["250"] in WP')
                WPD=250
            elif (int(x)==2):
                A=sheet_1.query('["350"] in WP')
                WPD=350
            dp= int(txt2.get()) - int(txt.get())
            res="Entered value of dP is " + str(dp) + " Bar"
            lbl3.configure(text=res, fg="green")
            C=A.where(A['dP']==dp)
            B=C.dropna(how='all')
            DT=B
            temp=1
            Pmin=float(txt.get())
            Pmax=float(txt2.get())
            Sa=int(dp/2)
            Sm=((Pmax+Pmin)/2)
            if (Pmin==0):
                Rr=0
            else:
                Rr=(Pmin/Pmax)
                
            if (len(B)==0):
                t=int(txt4.get())
                temp=2
                L=[]
                for j in range(dp-t,dp+t+1,1):
                    B=A['dP'].where(A['dP']==j)
                    C=B.dropna(how='all')
                    if (len(C)!=0):
                        D=C.to_string(index=False)
                        E=D[0:4]            
                        L.append(E[:])
                R=[]
                if (len(L)!=0):
                    V=len(E)-1
                    ress1="dP value not found\nAvailable dP values are\n"+ str(L[:])
                    lblt.configure(text=ress1)
                    for i in range(0,V):
                        map(float, L[i:i+1])
                        X=''.join(str(n) for n in L[i:i+1])
                        X1=float(X)
                        B=A.where(A['dP']==float(X1))
                        C=B.dropna(how='all')
                        globals()['U%s' % i]= C
                        R.append(globals()['U%s' % i])
                    Q=pd.concat(R)
                    DT=Q
                elif (len(L)==0):
                    temp=3
            if (temp==1):
                lblt.update()
                lblt.configure(text="dP value found", fg="green")
            if (temp!=3):
                B1=DT.query('["CTBJ"] in LOF')
                B2=DT.query('["CupWall"] in LOF')
                B3=DT.query('["PSC"] in LOF')
                B4=DT.query('["Tube"] in LOF')
                B5=DT.query('["IB"] in LOF')
                B1_k=B1[["K"]]
                B2_k=B2[["K"]]
                B3_k=B3[["K"]]
                B4_k=B4[["K"]]
                B5_k=B5[["K"]]
                c1=B1_k.mean()
                c2=B2_k.mean()
                c3=B3_k.mean()
                c4=B4_k.mean()
                c5=B5_k.mean()
                E1=B1[["FC"]]
                E2=B2[["FC"]]
                E3=B3[["FC"]]
                E4=B4[["FC"]]
                E5=B5[["FC"]]
                stressr=(-1*dp)
                Sa=(dp/2)
                r=float(txt3.get())
                Ka=(1/(0.912*(r**0.0829)))
                lblc6a=tkt.Label(win, text="Cup Failure", font=("Comic Sanc MS", 12))
                lblc6a.grid(column=0,row=8, sticky="w")
                if (len(B1)!=0 or len(B2)!=0):
                    od=float(cup.iloc[0,int(x)])
                    id=float(cup.iloc[1,int(x)])
                    t=float(cup.iloc[2,int(x)])
                    stressxc=(dp*(1/(((od/id)**2)-1)))
                    stresshc=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                    ab=(stressxc-stressr)**2
                    bc=(stressr-stresshc)**2
                    cd=(stresshc-stressxc)**2
                    von1=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                    Kb1=((od/7.62)**(-0.107))
                    End1=(485*0.5*Ka*Kb1*0.923)
                    A1=(((0.9*485)**2)/End1)
                    Bx=((0.9*485)/End1)
                    Ba=((-1)/3)*(mt.log(Bx,10))
                    lblc6.update()
                    res2="Axial stress= " + str(int(stressxc)) + " N/mm2"
                    lblc6.configure(text=res2, fg="black", font=("Comic Sanc MS", 10))
                    res3="Hoop stress= " + str(int(stresshc)) + " N/mm2"
                    lblc7.configure(text=res3)
                    res2="Von Misses stress= " + str(int(von1)) + " N/mm2"
                    lblc8.configure(text=res2)
                    res2="Ka= " + str(Ka)
                    lblc9.configure(text=res2)
                    res2="Kb= " + str(Kb1)
                    lblc10.configure(text=res2)
                    res2="Endurance limit= " + str(int(End1)) +" N/mm2"
                    lblc11.configure(text=res2)
                    if (len(B1)!=0):
                        N10=float((((von1*c1)/A1)**(1/Ba)))
                        lblc12.update()
                        res2="No. of Cycles(braze)= " + str(int(N10))
                        lblc12.configure(text=res2)
                        res2="Max cycles(datasheet)(braze)= " + str(int(E1.max()))
                        lblc13.configure(text=res2)
                        res2="Min cycles(datasheet)(braze)= " + str(int(E1.min()))
                        lblc14.configure(text=res2)
                    else:
                        lblc12.update()
                        res2="No Data found on Cup Braze joint"
                        lblc12.configure(text=res2, fg="red")
                        lblc13.configure(text="")
                        lblc14.configure(text="")
                    if (len(B2)!=0):
                        N11=(((von1*c2)/A1)**(1/Ba))
                        res2="No. of Cycles(wall)= " + str(int(N11))
                        lblc15.configure(text=res2, fg="black", font=("Comic Sanc MS", 10))
                        res2="Max cycles(datasheet)(braze)= " + str(int(E2.max()))
                        lblc16.configure(text=res2)
                        res2="Min cycles(datasheet)(braze)= " + str(int(E2.min()))
                        lblc17.configure(text=res2)
                    else:
                        lblc15.update()
                        res2="No Data found on Cup Wall"
                        lblc15.configure(text=res2, fg="red")
                        lblc16.configure(text="")
                        lblc17.configure(text="")
                else:
                    lblc6.update()
                    res2="No Value for dP found"
                    lblc6.configure(text=res2, fg="red")
                    lblc7.configure(text="")
                    lblc8.configure(text="")
                    lblc9.configure(text="")
                    lblc10.configure(text="")
                    lblc11.configure(text="")
                    lblc12.configure(text="")
                    lblc13.configure(text="")
                    lblc14.configure(text="")
                    lblc15.configure(text="")
                    lblc16.configure(text="")
                    lblc17.configure(text="")
                        
                        
                lbl6a=tkt.Label(win, text="Tube Failure", font=("Comic Sanc MS", 12))
                lbl6a.grid(column=1,row=8, sticky="w")    
                if (len(B4)!=0):
                    od=float(tube.iloc[0,int(x)])
                    id=float(tube.iloc[1,int(x)])
                    t=float(tube.iloc[2,int(x)])
                    stressxt=((dp*(1/(((od/id)**2)-1))))
                    stressht=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                    ab=(stressxt-stressr)**2
                    bc=(stressr-stressht)**2
                    cd=(stressht-stressxt)**2
                    von2=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                    Kb2=((od/7.62)**(-0.107))
                    End2=(485*0.5*Ka*Kb2*0.923)
                    A2=(((0.9*485)**2)/End2)
                    Bx=((0.9*485)/End2)
                    b2=((-1)/3)*(mt.log(Bx,10))
                    N2=(((von2*c4)/A2)**(1/b2))
                    lbl6.update()
                    res2="Axial stress= " + str(int(stressxt)) + " N/mm2"
                    lbl6.configure(text=res2, fg="black", font=("Comic Sanc MS", 10))
                    res3="Hoop stress= " + str(int(stressht)) + " N/mm2"
                    lbl7.configure(text=res3)
                    res2="Von Misses stress= " + str(int(von2)) + " N/mm2"
                    lbl8.configure(text=res2)
                    res2="Ka= " + str(Ka)
                    lbl9.configure(text=res2)
                    res2="Kb= " + str(Kb2)
                    lbl10.configure(text=res2)
                    res2="Endurance limit= " + str(int(End2)) +" N/mm2"
                    lbl11.configure(text=res2)
                    res2="No of Cycles= " + str(int(N2))
                    lbl12.configure(text=res2)
                    res2="Max cycles from datahseet= " + str(int(E4.max()))
                    lbl13.configure(text=res2)
                    res2="Min cycles from datahseet= " + str(int(E4.min()))
                    lbl14.configure(text=res2)
                else: 
                    lbl6.update()
                    res2="No Value for dP found"
                    lbl6.configure(text=res2, fg="red")
                    lbl7.configure(text="")
                    lbl8.configure(text="")
                    lbl9.configure(text="")
                    lbl10.configure(text="")
                    lbl11.configure(text="")
                    lbl12.configure(text="")
                    lbl13.configure(text="")
                    lbl14.configure(text="")
                
                lblp6a=tkt.Label(win, text="PSC", font=("Comic Sanc MS", 12))
                lblp6a.grid(column=2,row=8, sticky="w")
                if (len(B3)!=0):
                    od=float(psc.iloc[0,int(x)])
                    id=float(psc.iloc[1,int(x)])
                    t=float(psc.iloc[2,int(x)])
                    stressxp=((dp*(1/(((od/id)**2)-1))))
                    stresshp=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                    ab=(stressxp-stressr)**2
                    bc=(stressr-stresshp)**2
                    cd=(stresshp-stressxp)**2
                    von3=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                    Kb3=((od/7.62)**(-0.107))
                    End3=(485*0.5*Ka*Kb3*0.923)
                    A3=(((0.9*485)**2)/End3)
                    Bx=((0.9*485)/End3)
                    b3=((-1)/3)*(mt.log(Bx,10))
                    N3=(((von3*c3)/A3)**(1/b3))
                    lblp6.update()
                    res2="Axial stress= " + str(int(stressxp)) + " N/mm2"
                    lblp6.configure(text=res2, fg="black", font=("Comic Sanc MS", 10))
                    res3="Hoop stress= " + str(int(stresshp)) + " N/mm2"
                    lblp7.configure(text=res3)
                    res2="Von Misses stress= " + str(int(von3)) + " N/mm2"
                    lblp8.configure(text=res2)
                    res2="Ka= " + str(Ka)
                    lblp9.configure(text=res2)
                    res2="Kb= " + str(Kb3)
                    lblp10.configure(text=res2)
                    res2="Endurance limit= " + str(int(End3)) +" N/mm2"
                    lblp11.configure(text=res2)
                    res2="No of Cycles= " + str(int(N3))
                    lblp12.configure(text=res2)
                    res2="Max cycles from datahseet= " + str(int(E3.max()))
                    lblp13.configure(text=res2)
                    res2="Min cycles from datahseet= " + str(int(E3.min()))
                    lblp14.configure(text=res2)
                else: 
                    lblp6.update()
                    res2="No Value for dP found"
                    lblp6.configure(text=res2, fg="red")
                    lblp7.configure(text="")
                    lblp8.configure(text="")
                    lblp9.configure(text="")
                    lblp10.configure(text="")
                    lblp11.configure(text="")
                    lblp12.configure(text="")
                    lblp13.configure(text="")
                    lblp14.configure(text="")
                
                lbli6a=tkt.Label(win, text="Inlet Fitting", font=("Comic Sanc MS", 12))
                lbli6a.grid(column=3,row=8, sticky="w")
                if (len(B5)!=0):
                    od=float(inlet.iloc[0,int(x)])
                    id=float(inlet.iloc[1,int(x)])
                    t=float(inlet.iloc[2,int(x)])
                    stressxi=((dp*(1/(((od/id)**2)-1))))
                    stresshi=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                    ab=(stressxi-stressr)**2
                    bc=(stressr-stresshi)**2
                    cd=(stresshi-stressxi)**2
                    von4=((1/mt.sqrt(2))*(mt.sqrt(ab+bc+cd)))
                    Kb4=((od/7.62)**(-0.107))
                    End4=(485*0.5*Ka*Kb4*0.923)
                    A4=(((0.9*485)**2)/End4)
                    Bx=((0.9*485)/End4)
                    b4=((-1)/3)*(mt.log(Bx,10))
                    N4=(((von4*c5)/A4)**(1/b4))
                    lbli6.update()
                    res2="Axial stress= " + str(int(stressxi)) + " N/mm2"
                    lbli6.configure(text=res2, fg="black", font=("Comic Sanc MS", 10))
                    res3="Hoop stress= " + str(int(stresshi)) + " N/mm2"
                    lbli7.configure(text=res3)
                    res2="Von Misses stress= " + str(int(von4)) + " N/mm2"
                    lbli8.configure(text=res2)
                    res2="Ka= " + str(Ka)
                    lbli9.configure(text=res2)
                    res2="Kb= " + str(Kb4)
                    lbli10.configure(text=res2)
                    res2="Endurance limit= " + str(int(End4)) +" N/mm2"
                    lbli11.configure(text=res2)
                    res2="No of Cycles= " + str(int(N4))
                    lbli12.configure(text=res2)
                    res2="Max cycles from datahseet= " + str(int(E5.max()))
                    lbli13.configure(text=res2)
                    res2="Min cycles from datahseet= " + str(int(E5.min()))
                    lbli14.configure(text=res2)
                else: 
                    lbli6.update()
                    res2="No Value for dP found"
                    lbli6.configure(text=res2, fg="red")
                    lbli7.configure(text="")
                    lbli8.configure(text="")
                    lbli9.configure(text="")
                    lbli10.configure(text="")
                    lbli11.configure(text="")
                    lbli12.configure(text="")
                    lbli13.configure(text="")
                    lbli14.configure(text="")
                
                def save():
                    global progress
                    progress=ttk.Progressbar(win, length=100, mode='determinate')
                    progress.grid(column=5,row=2)
                    for fzs in range(0,20):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs) + "%"
                            lbls=tkt.Label(win, text=ressp,fg="red", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                    if (len(B1)!=0):
                        od=float(cup.iloc[0,int(x)])
                        id=float(cup.iloc[1,int(x)])
                        Dd=id/od
                        t=float(cup.iloc[2,int(x)])
                        C2=(0.427-(6.77*Dd)+(22.698*Dd*Dd)-(16.67*Dd*Dd*Dd))
                        C3=(11.357+(15.665*Dd)-(60.929*Dd*Dd)+(41.501*Dd*Dd*Dd))
                        Ktt=3+(C2*Dd)+(C3*Dd*Dd)
                        stressh=((WPD/10)*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        stressx=(((WPD/10)*(1/(((od/id)**2)-1))))
                        stressxw=((dp*(1/(((od/id)**2)-1))))
                        stresshw=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        ab=(stressxw-stressr)**2
                        bc=(stressr-stresshw)**2
                        cd=(stresshw-stressxw)**2
                        von=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                        Kb=((od/7.62)**(-0.107))
                        End=(485*0.5*Ka*Kb2*0.923)
                        A=(((0.9*485)**2)/End)
                        Bx=((0.9*485)/End)
                        Bb=((-1)/3)*(mt.log(Bx,10))
                        N=(((von*c1)/A2)**(1/Bb))
                        Nv=(((von)/A2)**(1/Bb))
                        SS=(Sm+((Sa*1.868*170)/End))
                        HSS=(stressxw*Ktt)
                        BSTC=((A)*((N)**(Bb)))
                        loc=FLoc
                        writer=pd.ExcelWriter(loc)
                        sheet_1=pd.read_excel(writer, sheet_name=0)
                        sheet_1.dropna(how='all')
                        L={'Test': ['TR'], 'WP': [float(WPD)], 'Pmin': [float(Pmin*10)], 'Pmax':[float(Pmax*10)], 'LOF':['CTBJ'], 'FC':[float(N)], 'dP':[float(dp)], 'Smean':[float(Sm)], 'Sa':[float(Sa)], 'R':[float(Rr)], 'dD':[float(Dd)], 'C2':[float(C2)], 'C3':[float(C3)], 'Kt':[float(Ktt)], 'OD':[od], 'ID':[id], 'T':[t], 'THS':[stressh], 'WHS':[stresshw], 'TAS':[stressx], 'WAS':[stressxw], 'Von':[von], 'SS':[SS], 'K':[float(c1)], 'HSS':[HSS], 'BSTC':[float(BSTC)], 'FOS':[2],'Rz':[r],'Ka':[Ka], 'Kb':[Kb], 'CEL':[End], 'A':[A], 'B':[Bb], 'BCV':[float(Nv)]}
                        Y=[]
                        df=pd.DataFrame(L)
                        new_df=[sheet_1, df]
                        Y=pd.concat(new_df)
                        Y.to_excel(writer, 'sheet_1', startrow=0, startcol=0, index=False, engine='openpyxl')
                        writer.save()
                        for fzs in range(20,40):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs) + "%"
                            lbls=tkt.Label(win, text=ressp,fg="red", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                    if (len(B2)!=0):
                        od=float(cup.iloc[0,int(x)])
                        id=float(cup.iloc[1,int(x)])
                        Dd=id/od
                        t=float(cup.iloc[2,int(x)])
                        C2=(0.427-(6.77*Dd)+(22.698*Dd*Dd)-(16.67*Dd*Dd*Dd))
                        C3=(11.357+(15.665*Dd)-(60.929*Dd*Dd)+(41.501*Dd*Dd*Dd))
                        Ktt=3+(C2*Dd)+(C3*Dd*Dd)
                        stressh=((WPD/10)*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        stressx=(((WPD/10)*(1/(((od/id)**2)-1))))
                        stressxw=((dp*(1/(((od/id)**2)-1))))
                        stresshw=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        ab=(stressxw-stressr)**2
                        bc=(stressr-stresshw)**2
                        cd=(stresshw-stressxw)**2
                        von=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                        Kb=((od/7.62)**(-0.107))
                        End=(485*0.5*Ka*Kb2*0.923)
                        A=(((0.9*485)**2)/End)
                        Bx=((0.9*485)/End)
                        Bb=((-1)/3)*(mt.log(Bx,10))
                        N=(((von*c2)/A2)**(1/Bb))
                        Nv=(((von)/A2)**(1/Bb))
                        SS=(Sm+((Sa*1.868*170)/End))
                        HSS=(stressxw*Ktt)
                        BSTC=((A)*((N)**(Bb)))
                        loc=FLoc
                        writer=pd.ExcelWriter(loc)
                        sheet_1=pd.read_excel(writer, sheet_name=0)
                        sheet_1.dropna(how='all')
                        L={'Test': ['TR'], 'WP': [float(WPD)], 'Pmin': [float(Pmin*10)], 'Pmax':[float(Pmax*10)], 'LOF':['CupWall'], 'FC':[float(N)], 'dP':[float(dp)], 'Smean':[float(Sm)], 'Sa':[float(Sa)], 'R':[float(Rr)], 'dD':[float(Dd)], 'C2':[float(C2)], 'C3':[float(C3)], 'Kt':[float(Ktt)], 'OD':[od], 'ID':[id], 'T':[t], 'THS':[stressh], 'WHS':[stresshw], 'TAS':[stressx], 'WAS':[stressxw], 'Von':[von], 'SS':[SS], 'K':[float(c2)], 'HSS':[HSS], 'BSTC':[float(BSTC)], 'FOS':[2],'Rz':[r],'Ka':[Ka], 'Kb':[Kb], 'CEL':[End], 'A':[A], 'B':[Bb], 'BCV':[float(Nv)]}
                        Y=[]
                        df=pd.DataFrame(L)
                        new_df=[sheet_1, df]
                        Y=pd.concat(new_df)
                        Y.to_excel(writer, 'sheet_1', startrow=0, startcol=0, index=False, engine='openpyxl')
                        writer.save()
                        for fzs in range(40,60):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs)+"%"
                            lbls=tkt.Label(win, text=ressp,fg="red", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                    if (len(B3)!=0):
                        od=float(psc.iloc[0,int(x)])
                        id=float(psc.iloc[1,int(x)])
                        Dd=id/od
                        t=float(psc.iloc[2,int(x)])
                        C2=(0.427-(6.77*Dd)+(22.698*Dd*Dd)-(16.67*Dd*Dd*Dd))
                        C3=(11.357+(15.665*Dd)-(60.929*Dd*Dd)+(41.501*Dd*Dd*Dd))
                        Ktt=3+(C2*Dd)+(C3*Dd*Dd)
                        stressh=((WPD/10)*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        stressx=(((WPD/10)*(1/(((od/id)**2)-1))))
                        stressxw=((dp*(1/(((od/id)**2)-1))))
                        stresshw=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        ab=(stressxw-stressr)**2
                        bc=(stressr-stresshw)**2
                        cd=(stresshw-stressxw)**2
                        von=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                        Kb=((od/7.62)**(-0.107))
                        End=(485*0.5*Ka*Kb2*0.923)
                        A=(((0.9*485)**2)/End)
                        Bx=((0.9*485)/End)
                        Bb=((-1)/3)*(mt.log(Bx,10))
                        N=(((von*c3)/A2)**(1/Bb))
                        Nv=(((von)/A2)**(1/Bb))
                        SS=(Sm+((Sa*1.868*170)/End))
                        HSS=(stressxw*Ktt)
                        BSTC=((A)*((N)**(Bb)))
                        loc=FLoc
                        writer=pd.ExcelWriter(loc)
                        sheet_1=pd.read_excel(writer, sheet_name=0)
                        sheet_1.dropna(how='all')
                        L={'Test': ['TR'], 'WP': [float(WPD)], 'Pmin': [float(Pmin*10)], 'Pmax':[float(Pmax*10)], 'LOF':['PSC'], 'FC':[float(N)], 'dP':[float(dp)], 'Smean':[float(Sm)], 'Sa':[float(Sa)], 'R':[float(Rr)], 'dD':[float(Dd)], 'C2':[float(C2)], 'C3':[float(C3)], 'Kt':[float(Ktt)], 'OD':[od], 'ID':[id], 'T':[t], 'THS':[stressh], 'WHS':[stresshw], 'TAS':[stressx], 'WAS':[stressxw], 'Von':[von], 'SS':[SS], 'K':[float(c3)], 'HSS':[HSS], 'BSTC':[float(BSTC)], 'FOS':[2],'Rz':[r],'Ka':[Ka], 'Kb':[Kb], 'CEL':[End], 'A':[A], 'B':[Bb], 'BCV':[float(Nv)]}
                        Y=[]
                        df=pd.DataFrame(L)
                        new_df=[sheet_1, df]
                        Y=pd.concat(new_df)
                        Y.to_excel(writer, 'sheet_1', startrow=0, startcol=0, index=False, engine='openpyxl')
                        writer.save()
                        for fzs in range(60,80):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs)+"%"
                            lbls=tkt.Label(win, text=ressp,fg="green", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                    if (len(B4)!=0):
                        od=float(tube.iloc[0,int(x)])
                        id=float(tube.iloc[1,int(x)])
                        Dd=id/od
                        t=float(tube.iloc[2,int(x)])
                        C2=(0.427-(6.77*Dd)+(22.698*Dd*Dd)-(16.67*Dd*Dd*Dd))
                        C3=(11.357+(15.665*Dd)-(60.929*Dd*Dd)+(41.501*Dd*Dd*Dd))
                        Ktt=3+(C2*Dd)+(C3*Dd*Dd)
                        stressh=((WPD/10)*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        stressx=(((WPD/10)*(1/(((od/id)**2)-1))))
                        stressxw=((dp*(1/(((od/id)**2)-1))))
                        stresshw=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        ab=(stressxw-stressr)**2
                        bc=(stressr-stresshw)**2
                        cd=(stresshw-stressxw)**2
                        von=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                        Kb=((od/7.62)**(-0.107))
                        End=(485*0.5*Ka*Kb2*0.923)
                        A=(((0.9*485)**2)/End)
                        Bx=((0.9*485)/End)
                        Bb=((-1)/3)*(mt.log(Bx,10))
                        N=(((von*c4)/A2)**(1/Bb))
                        Nv=(((von)/A2)**(1/Bb))
                        SS=(Sm+((Sa*1.868*170)/End))
                        HSS=(stressxw*Ktt)
                        BSTC=((A)*((N)**(Bb)))
                        loc=FLoc
                        writer=pd.ExcelWriter(loc)
                        sheet_1=pd.read_excel(writer, sheet_name=0)
                        sheet_1.dropna(how='all')
                        L={'Test': ['TR'], 'WP': [float(WPD)], 'Pmin': [float(Pmin*10)], 'Pmax':[float(Pmax*10)], 'LOF':['Tube'], 'FC':[float(N)], 'dP':[float(dp)], 'Smean':[float(Sm)], 'Sa':[float(Sa)], 'R':[float(Rr)], 'dD':[float(Dd)], 'C2':[float(C2)], 'C3':[float(C3)], 'Kt':[float(Ktt)], 'OD':[od], 'ID':[id], 'T':[t], 'THS':[stressh], 'WHS':[stresshw], 'TAS':[stressx], 'WAS':[stressxw], 'Von':[von], 'SS':[SS], 'K':[float(c4)], 'HSS':[HSS], 'BSTC':[float(BSTC)], 'FOS':[2],'Rz':[r],'Ka':[Ka], 'Kb':[Kb], 'CEL':[End], 'A':[A], 'B':[Bb], 'BCV':[float(Nv)]}
                        Y=[]
                        df=pd.DataFrame(L)
                        new_df=[sheet_1, df]
                        Y=pd.concat(new_df)
                        Y.to_excel(writer, 'sheet_1', startrow=0, startcol=0, index=False, engine='openpyxl')
                        writer.save()
                        for fzs in range(80,90):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs)+"%"
                            lbls=tkt.Label(win, text=ressp,fg="green", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                    if (len(B5)!=0):
                        od=float(inlet.iloc[0,int(x)])
                        id=float(inlet.iloc[1,int(x)])
                        Dd=id/od
                        t=float(inlet.iloc[2,int(x)])
                        C2=(0.427-(6.77*Dd)+(22.698*Dd*Dd)-(16.67*Dd*Dd*Dd))
                        C3=(11.357+(15.665*Dd)-(60.929*Dd*Dd)+(41.501*Dd*Dd*Dd))
                        Ktt=3+(C2*Dd)+(C3*Dd*Dd)
                        stressh=((WPD/10)*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        stressx=(((WPD/10)*(1/(((od/id)**2)-1))))
                        stressxw=((dp*(1/(((od/id)**2)-1))))
                        stresshw=(dp*((((od/id)**2)+1)/(((od/id)**2)-1)))
                        ab=(stressxw-stressr)**2
                        bc=(stressr-stresshw)**2
                        cd=(stresshw-stressxw)**2
                        von=(1/mt.sqrt(2))*mt.sqrt(ab+bc+cd)
                        Kb=((od/7.62)**(-0.107))
                        End=(485*0.5*Ka*Kb2*0.923)
                        A=(((0.9*485)**2)/End)
                        Bx=((0.9*485)/End)
                        Bb=((-1)/3)*(mt.log(Bx,10))
                        N=(((von*c5)/A2)**(1/Bb))
                        Nv=(((von)/A2)**(1/Bb))
                        SS=(Sm+((Sa*1.868*170)/End))
                        HSS=(stressxw*Ktt)
                        BSTC=((A)*((N)**(Bb)))
                        loc=FLoc
                        writer=pd.ExcelWriter(loc)
                        sheet_1=pd.read_excel(writer, sheet_name=0)
                        sheet_1.dropna(how='all')
                        L={'Test': ['TR'], 'WP': [float(WPD)], 'Pmin': [float(Pmin*10)], 'Pmax':[float(Pmax*10)], 'LOF':['IB'], 'FC':[float(N)], 'dP':[float(dp)], 'Smean':[float(Sm)], 'Sa':[float(Sa)], 'R':[float(Rr)], 'dD':[float(Dd)], 'C2':[float(C2)], 'C3':[float(C3)], 'Kt':[float(Ktt)], 'OD':[od], 'ID':[id], 'T':[t], 'THS':[stressh], 'WHS':[stresshw], 'TAS':[stressx], 'WAS':[stressxw], 'Von':[von], 'SS':[SS], 'K':[float(c5)], 'HSS':[HSS], 'BSTC':[float(BSTC)], 'FOS':[2],'Rz':[r],'Ka':[Ka], 'Kb':[Kb], 'CEL':[End], 'A':[A], 'B':[Bb], 'BCV':[float(Nv)]}
                        Y=[]
                        df=pd.DataFrame(L)
                        new_df=[sheet_1, df]
                        Y=pd.concat(new_df)
                        Y.to_excel(writer, 'sheet_1', startrow=0, startcol=0, index=False, engine='openpyxl')
                        writer.save()
                        for fzs in range(90,100):
                            progress['value']=fzs
                            win.update_idletasks()
                            time.sleep(0.005)
                            ressp=str(fzs)+"%"
                            lbls=tkt.Label(win, text=ressp,fg="green", font=("Comic Sanc MS", 10))            
                            lbls.grid(column=4,row=3)
                        progress.grid_forget()  
                    lbls=tkt.Label(win, text="Results Saved",fg="green", font=("Comic Sanc MS", 10))            
                    lbls.grid(column=4,row=3)
                    
                    
                btns=tkt.Button(win, text="Save Results",fg="blue", command=save)
                btns.grid(column=4, row=2)
            if (temp==3):
                lblt.update()
                ress1="No values Found even in tolerance zone"
                lblt.configure(text=ress1, fg="red", font=("Times"))
                res2="No Value for dP found"
                lblc6.configure(text=res2, fg="red", font=("Times"))
                lblc7.configure(text="")
                lblc8.configure(text="")
                lblc9.configure(text="")
                lblc10.configure(text="")
                lblc11.configure(text="")
                lblc12.configure(text="")
                lblc13.configure(text="")
                lblc14.configure(text="")
                lblc15.configure(text="")
                lblc16.configure(text="")
                lblc17.configure(text="")
                res2="No Value for dP found"
                lbl6.configure(text=res2, fg="red", font=("Times"))
                lbl7.configure(text="")
                lbl8.configure(text="")
                lbl9.configure(text="")
                lbl10.configure(text="")
                lbl11.configure(text="")
                lbl12.configure(text="")
                lbl13.configure(text="")
                lbl14.configure(text="")
                res2="No Value for dP found"
                lblp6.configure(text=res2, fg="red", font=("Times"))
                lblp7.configure(text="")
                lblp8.configure(text="")
                lblp9.configure(text="")
                lblp10.configure(text="")
                lblp11.configure(text="")
                lblp12.configure(text="")
                lblp13.configure(text="")
                lblp14.configure(text="")
                res2="No Value for dP found"
                lbli6.configure(text=res2, fg="red", font=("Times"))
                lbli7.configure(text="")
                lbli8.configure(text="")
                lbli9.configure(text="")
                lbli10.configure(text="")
                lbli11.configure(text="")
                lbli12.configure(text="")
                lbli13.configure(text="")
                lbli14.configure(text="")
        else:
            messagebox.showinfo("Error", "Invalid File")            
    
    
    btn=tkt.Button(win, text="Check", command=clicked)
    btn.grid(column=1, row=7, sticky="w")
    def restart():
        win.destroy()
        start()
    
    btncl=tkt.Button(win, text="Restart window", command=restart)
    btncl.grid(column=4, row=15)


start()   
win.mainloop()

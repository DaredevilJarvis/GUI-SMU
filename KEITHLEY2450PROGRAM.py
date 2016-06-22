from tkinter import *
import visa
from matplotlib import pyplot as plt
from matplotlib import style
import xlsxwriter
from tkinter import filedialog

root=Tk()
adr=StringVar() # addres from Intsrument
var1=IntVar()# Selecting the Bias
var2=IntVar()# Selecting the Buffer
var3=StringVar() #
srcv=DoubleVar() # Source Voltage
coun=IntVar() # for counting the Values
bufsize=IntVar() # Buffer Size
var4=IntVar() #value of wire sense
wiresense=StringVar() # Wire Sense
wiresense="ON"
var5=IntVar()  #  Route
rt=StringVar() # route front or rear
var6=IntVar() #  Measure for Voltage Bias
meas=StringVar() #  measurng Current Voltage and Resistance
curlimit=DoubleVar() #  Current Limit
currange=DoubleVar()
voltrange=DoubleVar()
svdelay=DoubleVar()
var7=IntVar() #  measure for Current
srcc=DoubleVar() # Current Level
voltlimit=DoubleVar() # Voltage Limit
voltrange=DoubleVar()
voltcrange=DoubleVar()
curbrange=DoubleVar()
scdelay=DoubleVar()  # source current 
measc=StringVar()  #  Measure the Function for Current Bias
var8=IntVar()  #  X axis Value
var9=IntVar()  #Y axis value
x=IntVar()
y=IntVar()
buff=StringVar()
# Function To Exit =======
label4=Label()
label7=Label()
entry2=Entry()
entry5=Entry()
label5=Label()
label9=Label()
label10=Label()
label11=Label()
label12=Label()
label13=Label()
label14=Label()
label15=Label()
label16=Label()
label17=Label()
label18=Label()
label19=Label()
R9=Button()
R10=Button()
R11=Button()
R12=Button()
R13=Button()
R14=Button()
R12=Button()
R19=Button()
R20=Button()
entry3=Entry()
entry6=Entry()
entry7=Entry()
entry8=Entry()
entry9=Entry()
entry10=Entry()
entry11=Entry()
entry12=Entry()
entry13=Entry()
entry14=Entry()
# ====================    
# Buffer creation
def usedef():
    global label4
    global entry2
    global label7
    global entry5
    global buff
    buff="defbuffer1"
    label4.destroy()
    label7.destroy()
    entry2.destroy()
    entry5.destroy()
def userdef():
    global label4
    global entry2
    global label7
    global entry5
    label4=Label(root,text="Buffer Name",font=16,bd=5,padx=2,pady=2,relief=RIDGE)
    label4.grid(row=1,column=3)
    entry2=Entry(root,textvariable=var3,font=16,bd=5,relief=SUNKEN,width=15) 
    entry2.grid(row=1,column=4)
    label7=Label(root,text="BufferSize ",font=10,bd=5,padx=2,pady=2,relief=GROOVE)
    label7.grid(row=1,column=5,sticky=W)
    entry5=Entry(root,textvariable=bufsize,font=10,bd=5,relief=SUNKEN,width=12) 
    entry5.grid(row=1,column=6,sticky=W)
    
    
    
    
#=========  ============ ========== ====== 

####---- Wire SEnse
def wire():
    global wiresense
    if var4.get() is 1:
        wiresense="OFF"
    else:
        wiresense="ON"
#============================
       
#========== Route
def route():
    global rt
    if var5.get() is 1:
         rt="FRON"
    else:
         rt="REAR"

#========== Measure Function for Voltage Current REsistance
def measure():
    global meas
    w1=var6.get()
    if w1 is 1:
        meas="VOLT"
    elif w1 is 2:
        meas="CURR"
    else:
        meas="RES"
#===  measure Function for Current Bias 
def measurecb():
    global measc
    if var7.get()==1:
        measc="VOLT"
    elif var7.get()==2:
        measc="CURR"
    else:
        measc="RES"

#==== X axis
def xaxis():
    global x
    if var8.get()==1:
        x=1
    elif var8.get()==2:
        x=2
    else:
        x=3

#===== Y axis 
def yaxis():
    global y
    if var9.get()==1:
        y=1
    else:
        y=2
        
 #==================




#==========  
# ----Voltage BIas    
def voltbias():

    global label5
    global label9
    global label10
    global label11
    global label12
    global label13
    global label14
    global label15
    global label16
    global label17
    global label18
    global label19
    global R9
    global R10
    global R11
    global R12
    global R13
    global R14
    global R12
    global R19
    global R20
    global entry3
    global entry6
    global entry7
    global entry8
    global entry9
    global entry10
    global entry11
    global entry12
    global entry13
    global entry14

    
    label9=Label(root,text="Measure  ",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
    label9.grid(row=4,column=0,sticky=W)
    R9=Radiobutton(root,text="Voltage ",variable=var6,value=1,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measure)
    R9.grid(row=4,column=1,sticky=W)
    R10=Radiobutton(root,text="Current ",variable=var6,value=2,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measure)
    R10.grid(row=4,column=2,sticky=W)
    R11=Radiobutton(root,text="Resistance",variable=var6,value=3,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measure)
    R11.grid(row=4,column=3,sticky=W)
    label5=Label(root,text="Voltage Level (V) ",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
    label5.grid(row=5,column=0,sticky=W)
    entry3=Entry(root,textvariable=srcv,font=16,bd=5,relief=SUNKEN,width=15) 
    entry3.grid(row=5,column=1,sticky=W)
    label10=Label(root,text="Current Limit (A) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label10.grid(row=5,column=2,sticky=W)
    entry6=Entry(root,textvariable=curlimit,font=16,bd=5,relief=SUNKEN,width=15) 
    entry6.grid(row=5,column=3,sticky=W)
    label11=Label(root,text="Current Range (A) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label11.grid(row=5,column=4,sticky=W)
    entry7=Entry(root,textvariable=currange,font=16,bd=5,relief=SUNKEN,width=15) 
    entry7.grid(row=5,column=5,sticky=W)
    label12=Label(root,text="Source Range (V) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label12.grid(row=6,column=0,sticky=W)
    entry8=Entry(root,textvariable=voltrange,font=16,bd=5,relief=SUNKEN,width=15) 
    entry8.grid(row=6,column=1,sticky=W)
    label19=Label(root,text="Source Delay ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label19.grid(row=6,column=2,sticky=W)
    entry14=Entry(root,textvariable=svdelay,font=16,bd=5,relief=SUNKEN,width=15) 
    entry14.grid(row=6,column=3,sticky=W)
    R19=Button(root,text="   Execute It!!!   ",font=20,bd=10,relief=RAISED,fg='red',highlightcolor='Blue',command=voltrun,bg='grey')
    R19.grid(row=9,column=2)
    label13.destroy()
    label14.destroy()
    label15.destroy()
    label16.destroy()
    label17.destroy()
    label18.destroy()
    R12.destroy()
    R13.destroy()
    R14.destroy()
    R20.destroy()
    entry9.destroy()
    entry10.destroy()
    entry11.destroy()
    entry12.destroy()
    entry13.destroy()
    
    
    
    
    
#==============================



def currbias():
    global label5
    global label9
    global label10
    global label11
    global label12
    global label13
    global label14
    global label15
    global label16
    global label17
    global label18
    global label19
    global R9
    global R10
    global R11
    global R12
    global R13
    global R14
    global R19
    global R20
    global entry3
    global entry6
    global entry7
    global entry8
    global entry9
    global entry10
    global entry11
    global entry12
    global entry13

    label13=Label(root,text="Measure  ",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
    label13.grid(row=4,column=0,sticky=W)
    R12=Radiobutton(root,text="Voltage ",variable=var7,value=1,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measurecb)
    R12.grid(row=4,column=1,sticky=W)
    R13=Radiobutton(root,text="Current ",variable=var7,value=2,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measurecb)
    R13.grid(row=4,column=2,sticky=W)
    R14=Radiobutton(root,text="Resistance",variable=var7,value=3,borderwidth=4,font=16,fg='black',relief=RIDGE,command=measurecb)
    R14.grid(row=4,column=3,sticky=W)
    label14=Label(root,text="Current Level (A) ",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
    label14.grid(row=5,column=0,sticky=W)
    entry9=Entry(root,textvariable=srcc,font=16,bd=5,relief=SUNKEN,width=15) 
    entry9.grid(row=5,column=1,sticky=W)
    label15=Label(root,text="Voltage Limit (V) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label15.grid(row=5,column=2,sticky=W)
    entry10=Entry(root,textvariable=voltlimit,font=16,bd=5,relief=SUNKEN,width=15) 
    entry10.grid(row=5,column=3,sticky=W)
    label16=Label(root,text="Voltage Range (V) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label16.grid(row=5,column=4,sticky=W)
    entry11=Entry(root,textvariable=voltcrange,font=16,bd=5,relief=SUNKEN,width=15) 
    entry11.grid(row=5,column=5,sticky=W)
    label17=Label(root,text="Source Range (A) ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label17.grid(row=6,column=0,sticky=W)
    entry12=Entry(root,textvariable=curbrange,font=16,bd=5,relief=SUNKEN,width=15) 
    entry12.grid(row=6,column=1,sticky=W)
    label18=Label(root,text="Source Delay ",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
    label18.grid(row=6,column=2,sticky=W)
    entry13=Entry(root,textvariable=scdelay,font=16,bd=5,relief=SUNKEN,width=15) 
    entry13.grid(row=6,column=3,sticky=W)
    R20=Button(root,text="   Execute It!!!   ",font=20,bd=10,relief=RAISED,fg='red',highlightcolor='blue',command=currun,bg='grey')
    R20.grid(row=9,column=2)
    label5.destroy()
    label9.destroy()
    label11.destroy()
    label12.destroy()
    label10.destroy()
    label19.destroy()
    R9.destroy()
    R10.destroy()
    R11.destroy()
    R19.destroy()
    entry3.destroy()
    entry6.destroy()
    entry7.destroy()
    entry8.destroy()
    entry14.destroy()

    
    
         
    
 #========================================================== VOLtage Run Program===============
def voltrun():  
    global buff
    global var2
    global bufsize
    global meas
    global srcv
    global curlimit
    global currange
    global voltrange
    global svdelay
    global rt
    global wiresense
    global coun
    global x
    global y
    
    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("All Files",".*")],initialfile="Untitled")
    workbook = xlsxwriter.Workbook(file_path)
    sheet1 = workbook.add_worksheet()
    sheet1.write(0, 0, "Relative Time(s)")
    sheet1.write(0, 1, meas)
    sheet1.write(0, 2, "Voltage  ")
    sheet1.write(0, 3, "Time Stamp ")
    rm = visa.ResourceManager('C:\\Program Files (x86)\\IVI Foundation\\VISA\\WinNT\\agvisa\\agbin\\visa32.dll')
    sar1=adr.get()
    myinst =rm.open_resource("%s"%sar1)
    myinst.timeout= 50000
    myinst.write("*RST")
    if var2.get()==2:
        buff=var3.get()
        sar2=bufsize.get()
        myinst.write("TRAC:MAKE \"%s\",%d"%(buff,sar2))
    myinst.write("SOUR:FUNC VOLT")
    myinst.write("SENS:FUNC \"%s\""%meas)
    myinst.write("SOUR:VOLT:READ:BACK ON")
    sar3=srcv.get()
    myinst.write("SOUR:VOLT %f"%sar3)
    myinst.write("TRAC:CLE \"%s\""%buff)
    sar4=curlimit.get()
    myinst.write("SOUR:VOLT:ILIM %f"%sar4)
    sar5=currange.get()
    myinst.write("SENS:CURR:RANG %f"%sar5)
    sar6=voltrange.get()
    myinst.write("SOUR:VOLT:RANG %f"%sar6)
    sar7=svdelay.get()
    myinst.write("SOUR:VOLT:DEL %f"%sar7)
    myinst.write("ROUT:TERM %s"%rt)
    myinst.write("SENS:%s:RSEN %s"%(meas,wiresense))
    myinst.write("OUTP ON")
    style.use('ggplot')
    fig = plt.figure()
    ax1 = fig.add_subplot(1,1,1)
    sar8=coun.get()
    a=[]
    b=[]
    c=[]
    d=[]
    m=1
    plt.show()
    for i in range(0,sar8):
        myinst.write("COUNT 1")
        myinst.write("READ? \"%s\",REL,READ,SOUR,TST"%buff)
        t=myinst.read()
        p,q,r,s=t.split(',')
        a.append(float(p))
        b.append(float(q))
        c.append(float(r))
        d.append(str(s))
        sheet1.write(m,0,float(p))
        sheet1.write(m,1,float(q))
        sheet1.write(m,2,float(r))
        sheet1.write(m,3,str(s))
        ax1.clear()
        if x== 1:
            p=a
            plt.xlabel("Time->")
        elif x==2:
            p=c
            plt.xlabel("Voltage")
        else:
            p=b
            plt.xlabel("Current->")
        if y==1:
            q=c
            plt.ylabel("Voltage->")
        else:
            q=b
            plt.ylabel("Current->")
        plt.plot(p,q,'g',linewidth="2")
        plt.grid("True",color='r')
        plt.pause(0.005)
        m=m+1

   
    
    
    
    
    
    
    
    myinst.write("OUTP OFF")   
    myinst.close()
    workbook.close()
    
    
#============  ================  =========
    
#=========== Ecetution of Current bias  ===========================
    
def currun():
    global buff
    global var2
    global bufsize
    global measc
    global srcc
    global voltlimit
    global curbrange
    global voltcrange
    global scdelay
    global rt
    global wiresense
    global coun
    global x
    global y
    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("All Files",".*")],initialfile="Untitled")
    workbook = xlsxwriter.Workbook(file_path)
    sheet1 = workbook.add_worksheet()
    sheet1.write(0, 0, "Relative Time(s)")
    sheet1.write(0, 1, measc)
    sheet1.write(0, 2, "CURRENT  ")
    sheet1.write(0, 3, "Time Stamp ")
    rm = visa.ResourceManager('C:\\Program Files (x86)\\IVI Foundation\\VISA\\WinNT\\agvisa\\agbin\\visa32.dll')
    gar1=adr.get()
    myinst =rm.open_resource("%s"%gar1)
    myinst.timeout= 50000
    myinst.write("*RST")
    if var2.get()==2:
        buff=var3.get()
        gar2=bufsize.get()
        myinst.write("TRAC:MAKE \"%s\",%d"%(buff,gar2))
    myinst.write("SOUR:FUNC CURR")
    myinst.write("SENS:FUNC \"%s\""%measc)
    myinst.write("SOUR:CURR:READ:BACK ON")
    gar3=srcc.get()
    myinst.write("SOUR:CURR %f"%gar3)
    myinst.write("TRAC:CLE \"%s\""%buff)
    gar4=voltlimit.get()
    myinst.write("SOUR:CURR:VLIM %f"%gar4)
    gar5=voltcrange.get()
    myinst.write("SENS:VOLT:RANG %f"%gar5)
    gar6=curbrange.get()
    myinst.write("SOUR:CURR:RANG %f"%gar6)
    gar7=scdelay.get()
    myinst.write("SOUR:CURR:DEL %f"%gar7)
    myinst.write("ROUT:TERM %s"%rt)
    myinst.write("SENS:%s:RSEN %s"%(measc,wiresense))
    myinst.write("OUTP ON")
    style.use('ggplot')
    fig = plt.figure()
    ax1 = fig.add_subplot(1,1,1)
    gar8=coun.get()
    a=[]
    b=[]
    c=[]
    d=[]
    m=1
    plt.show()
    for i in range(0,gar8):
        myinst.write("COUNT 1")
        myinst.write("READ? \"%s\",REL,READ,SOUR,TST"%buff)
        t=myinst.read()
        p,q,r,s=t.split(',')
        a.append(float(p))
        b.append(float(q))
        c.append(float(r))
        d.append(str(s))
        sheet1.write(m,0,float(p))
        sheet1.write(m,1,float(q))
        sheet1.write(m,2,float(r))
        sheet1.write(m,3,str(s))
        ax1.clear()
        if x== 1:
            p=a
            plt.xlabel("Time->")
        elif x==2:
            p=c
            plt.xlabel("Voltage")
        else:
            p=b
            plt.xlabel("Current->")
        if y==1:
            q=c
            plt.ylabel("Voltage->")
        else:
            q=b
            plt.ylabel("Current->")
        plt.plot(p,q,'g',linewidth="2")
        plt.grid("True",color='r')
        plt.pause(0.005)
        m=m+1

   
    
    
    
    
    
    
    
    myinst.write("OUTP OFF")   
    myinst.close()
    workbook.close()



























#   ***** MenuBar  ******
root.title("SourceMeter GUI")
menu=Menu(root)
root.config(menu=menu)

subMenu=Menu(menu)
menu.add_cascade(label="File",menu=subMenu)
subMenu.add_command(label="Exit",command=root.quit)   

#=============================


# Conecting the Intsrument======
label1=Label(root,text="Device Connection Address",font=16,bd=5,padx=2,pady=2,relief=SUNKEN)
label1.grid(row=0)
entry1=Entry(root,textvariable=adr,font=16,bd=5,relief=SUNKEN,width=50) 
entry1.grid(row=0,column=1)
#====================
#=== Selection of Buffer



 
label3=Label(root,text="Buffer",font=16,bd=5,padx=2,pady=2,relief=RIDGE)
label3.grid(row=1,column=0,sticky=W)
R3=Radiobutton(root,text=" Use Default",variable=var2,value=1,borderwidth=4,font=16,fg='black',relief=GROOVE,command=usedef)
R3.grid(row=1,column=1,sticky=W)
R4=Radiobutton(root,text="User Defined",variable=var2,value=2,borderwidth=4,font=16,fg='black',relief=GROOVE,command=userdef)
R4.grid(row=1,column=2,sticky=W)


    
#========== ==================
#========= Set the Count Value  ===========
label6=Label(root,text="Measurements  ",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
label6.grid(row=2,column=0,sticky=W)
entry4=Entry(root,textvariable=coun,font=16,bd=5,relief=SUNKEN,width=12) 
entry4.grid(row=2,column=1,sticky=W)

#===== Two Wire and Four Wire Sense===
R5=Radiobutton(root,text="2 Wire Sense",variable=var4,value=1,borderwidth=4,font=16,fg='black',relief=GROOVE,command=wire)
R5.grid(row=2,column=2,sticky=W)
R6=Radiobutton(root,text="4 Wire Sense",variable=var4,value=2,borderwidth=4,font=16,fg='black',relief=GROOVE,command=wire)
R6.grid(row=2,column=3,sticky=W)
label8=Label(root,text="ROUTE",font=16,bd=5,padx=2,pady=2,relief=GROOVE)
label8.grid(row=2,column=4,sticky=W)
R7=Radiobutton(root,text="FRONT",variable=var5,value=1,borderwidth=4,font=16,fg='black',relief=GROOVE,command=route)
R7.grid(row=2,column=5,sticky=W)
R8=Radiobutton(root,text="REAR",variable=var5,value=2,borderwidth=4,font=16,fg='black',relief=GROOVE,command=route)
R8.grid(row=2,column=6,sticky=W)





#    
    
# Selection of Voltage and current Bias
label2=Label(root,text="Select the Bias",font=16,bd=5,padx=2,pady=2,relief=RIDGE)
label2.grid(row=3,column=0,sticky=W)
R1=Radiobutton(root,text="Voltage Bias",variable=var1,value=1,borderwidth=4,font=16,fg='red',relief=GROOVE,command=voltbias)
R1.grid(row=3,column=1,sticky=W)
R2=Radiobutton(root,text="Current Bias",variable=var1,value=2,borderwidth=4,font=16,fg='blue',relief=GROOVE,command=currbias)
R2.grid(row=3,column=2,sticky=W)




#====== Graph Plotting Axis
label20=Label(root,text="On X Axis ",font=16,bd=5,padx=2,pady=2,relief=RAISED)
label20.grid(row=7,column=0,sticky=W)
R22=Radiobutton(root,text="TIME ",variable=var8,value=1,borderwidth=4,font=16,fg='black',relief=SUNKEN,command=xaxis)
R22.grid(row=7,column=1,sticky=W)
R23=Radiobutton(root,text="VOLTAGE",variable=var8,value=2,borderwidth=4,font=16,fg='black',relief=SUNKEN,command=xaxis)
R23.grid(row=7,column=2,sticky=W)
R24=Radiobutton(root,text="CURRENT",variable=var8,value=3,borderwidth=4,font=16,fg='black',relief=SUNKEN,command=xaxis)
R24.grid(row=7,column=3,sticky=W)
label21=Label(root,text="On Y Axis ",font=16,bd=5,padx=2,pady=2,relief=RAISED)
label21.grid(row=8,column=0,sticky=W)
R25=Radiobutton(root,text="VOLTAGE ",variable=var9,value=1,borderwidth=4,font=16,fg='black',relief=SUNKEN,command=yaxis)
R25.grid(row=8,column=1,sticky=W)
R26=Radiobutton(root,text="CURRENT",variable=var9,value=2,borderwidth=4,font=16,fg='black',relief=SUNKEN,command=yaxis)
R26.grid(row=8,column=2,sticky=W)




root.mainloop()


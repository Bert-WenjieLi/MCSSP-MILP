from gurobipy import *
import math
import xlrd
#print(gurobipy.__Version__)
Spoint=7+1
Dpoint=14+1
TCpoint=4+1
MaNum=4+1
Requests=16+1
Stages=2+1
Cpoint=8+1
Sequence=4+1
OverZone=4+1

data=xlrd.open_workbook(r'4 tower crane locations_new.xls')
supply=data.sheet_by_name('supply')
demand=data.sheet_by_name('demand')
crane=data.sheet_by_name('crane')
#supply positions data
sx={}
sy={}
sz={}   
for i in range(1,Spoint):
    sx[i]=supply.row(i)[1].value
    sy[i]=supply.row(i)[2].value
    sz[i]=supply.row(i)[3].value

#demand positions data
dx={}
dy={}
dz={}
for j in range(1,Dpoint):
    dx[j]=demand.row(j)[1].value
    dy[j]=demand.row(j)[2].value
    dz[j]=demand.row(j)[3].value

#crane positions data
cx={}
cy={}
cz={}
for k in range(1,TCpoint):
    cx[k]=crane.row(k)[1].value
    cy[k]=crane.row(k)[2].value
    cz[k]=crane.row(k)[3].value

#point A and B
ax={}
ay={}
for  p in range(1,Cpoint):
    ax[p]=demand.row(p)[6].value
    ay[p]=demand.row(p)[7].value
    #print dx[p],dy[p]
##calculate the operate time 
Dis_STC={}
Dis_DTC={}
Dis_SD={}
Time_sd_r={}
Time_sd_w={}
Time_sd_co={}
Time_sd_v={}
Time_sd_run={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint): 
            Dis_STC[i,k]=math.sqrt((sx[i]-cx[k])**2+(sy[i]-cy[k])**2)
            Dis_DTC[j,k]=math.sqrt((dx[j]-cx[k])**2+(dy[j]-cy[k])**2)
            Dis_SD[i,j]=math.sqrt((dx[j]-sx[i])**2+(dy[j]-sy[i])**2)
            Time_sd_r[i,j,k]=abs(Dis_STC[i,k]-Dis_DTC[j,k])/60
            Time_sd_w[i,j,k]=math.acos(-(Dis_SD[i,j]**2-Dis_DTC[j,k]**2-Dis_STC[i,k]**2)/(2*Dis_STC[i,k]*Dis_DTC[j,k]))/0.5
            Time_sd_co[i,j,k]=max(Time_sd_r[i,j,k],Time_sd_w[i,j,k])+0*min(Time_sd_r[i,j,k],Time_sd_w[i,j,k])
            Time_sd_v[i,j,k]=abs(sz[i]-dz[j]+3)/136
            Time_sd_run[i,j,k]=max(Time_sd_co[i,j,k],Time_sd_v[i,j,k])+min(Time_sd_co[i,j,k],Time_sd_v[i,j,k])
##if the operate path will entire the overlap space
Dis_CTC={}
Dis_CS={}
Dis_CD={}
Time_SC={}
Time_DC={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
            for  p in range(1,Cpoint):
                Dis_CTC[p,k]=math.sqrt((ax[p]-cx[k])**2+(ay[p]-cy[k])**2)#distance of tower crane and point A or B
                Dis_CS[p,i]=math.sqrt((ax[p]-sx[i])**2+(ay[p]-sy[i])**2)#distance of point A or B and supply point
                Dis_CD[p,j]=math.sqrt((dx[j]-ax[p])**2+(dy[j]-ay[p])**2)#distance of point A or B and demand point
                #operate time from supply point to point A or B
                Time_SC[p,i,k]=math.acos(-(Dis_CS[p,i]**2-Dis_CTC[p,k]**2-Dis_STC[i,k]**2)/(2*Dis_STC[i,k]*Dis_CTC[p,k]))/0.5
                Time_DC[p,j,k]=math.acos(-(Dis_CD[p,j]**2-Dis_CTC[p,k]**2-Dis_DTC[j,k]**2)/(2*Dis_DTC[j,k]*Dis_CTC[p,k]))/0.5 
                #if i==3 and k==2 and (p==5 or p==6):
                #    print ('p,i,k',p,i,k,Time_SC[p,i,k])
Time_AB={}
for k in range(1,TCpoint):
    for o in range(1,OverZone):
        Time_AB[k,o]=0
Time_AB[1,1]=2.72
Time_AB[2,1]=2.72
Time_AB[1,2]=2.85
Time_AB[3,2]=2.85
Time_AB[2,3]=2.93
Time_AB[3,3]=2.93
Time_AB[2,4]=3.60
Time_AB[4,4]=3.60
#D在干涉区域
Path_C={}
for j in range(1,Dpoint):
    for o in range(1,OverZone):
        Path_C[j,o]=0
Path_C[2,1]=1
Path_C[2,1]=1
Path_C[3,1]=1
Path_C[3,2]=1
Path_C[3,1]=1
Path_C[3,3]=1
Path_C[3,2]=1
Path_C[3,3]=1
Path_C[4,2]=1
Path_C[4,2]=1
Path_C[6,3]=1
Path_C[6,3]=1
#S在干涉区域
Path_D={}
for i in range(1,Spoint):
    for o in range(1,OverZone):
        Path_D[i,o]=0
Path_D[3,4]=1
Path_D[4,4]=1
        
Time_P1={}
Time_P2={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
            for o in range(1,OverZone):
                if Time_AB[k,o]!=0 :#and Dis_STC[i,k]<=70 and Dis_DTC[j,k]<=70:
                    p=2*o-1
                    if Path_C[j,o]==1 or Path_D[i,o]==1:
                        Time_P1[i,j,k,o] = Time_SC[p,i,k]+Time_DC[p,j,k]
                        Time_P2[i,j,k,o] = Time_SC[p+1,i,k]+Time_DC[p+1,j,k]
                        #if i==3 and j==14 and o==4:
                        #    print(i,j,k,o,Time_SC[p,i,k],Time_SC[p+1,j,k])
                    else:
                        Time_P1[i,j,k,o] = Time_SC[p,i,k]+Time_DC[p+1,j,k]+Time_AB[k,o]
                        Time_P2[i,j,k,o] = Time_SC[p+1,i,k]+Time_DC[p,j,k]+Time_AB[k,o]
Path_A={}
Path={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
            for o in range(1,OverZone):
                if Time_AB[k,o]!=0 :#and Dis_STC[i,k]<=70 and Dis_DTC[j,k]<=70:
                    p=2*o-1
                    if Time_P1[i,j,k,o]<=Time_sd_run[i,j,k]:
                        Path_A[i,j,k,o,p] = 1
                    else:
                        Path_A[i,j,k,o,p] = 0
                    if Time_P2[i,j,k,o]<=Time_sd_run[i,j,k]:
                        Path_A[i,j,k,o,p+1] = 1
                    else:
                        Path_A[i,j,k,o,p+1] = 0

for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
            for o in range(1,OverZone):
                p=2*o-1
                if Time_AB[k,o]!=0 :#and Dis_STC[i,k]<=70 and Dis_DTC[j,k]<=70:
                    if Path_A[i,j,k,o,p]==1 or Path_A[i,j,k,o,p+1]==1:
                        Path[i,j,k,o]=1
                    else:
                        Path[i,j,k,o]=0
                    #if k==2 and Path[i,j,k,o]==1 and Dis_STC[i,k]<=70 and Dis_DTC[j,k]<=70:
                    #    print(i,j,k,o,Path[i,j,k,o])

MaStore={}
for i in range(1,Spoint):
    for m in range(1,MaNum):
        MaStore[i,m]=1
MaStore[1,3]=1
MaStore[1,4]=1
MaStore[2,1]=1
MaStore[2,2]=1
MaStore[3,1]=1
MaStore[3,3]=1
MaStore[3,4]=1
MaStore[4,2]=1
MaStore[5,1]=1
MaStore[5,4]=1
MaStore[6,2]=1
MaStore[6,3]=1
MaStore[7,2]=1
MaStore[7,1]=1
MaDe={}
for j in range(1,Dpoint):
    for m in range(1,MaNum):
        for r in range(1,Requests):
            MaDe[r,j,m]=0
MaDe[12,1,2]=1        
MaDe[11,2,4]=1
MaDe[7,3,3]=1
MaDe[5,3,1]=1
MaDe[6,3,2]=1
MaDe[10,4,4]=1
MaDe[4,5,3]=1
MaDe[15,6,1]=1
MaDe[8,7,4]=1
MaDe[1,8,2]=1
MaDe[9,9,3]=1
MaDe[16,10,4]=1
MaDe[14,11,3]=1
MaDe[3,12,4]=1
MaDe[13,13,2]=1
MaDe[2,14,1]=1

MaDe[17,2,3]=1
MaDe[18,6,2]=1
MaDe[19,4,3]=1
MaDe[20,13,1]=1

MaDe[21,1,1]=1
MaDe[22,8,2]=1
MaDe[23,9,2]=1
MaDe[24,11,4]=1

MaDe[25,2,4]=1
MaDe[26,7,3]=1
MaDe[27,10,1]=1
MaDe[28,12,2]=1

MaDe[29,3,4]=1
MaDe[30,2,2]=1
MaDe[31,3,2]=1
MaDe[32,14,1]=1

MaDe[33,8,4]=1
MaDe[34,4,2]=1
MaDe[35,9,2]=1
MaDe[36,12,3]=1
MaDeW={}
for j in range(1,Dpoint):
    for m in range(1,MaNum):
        MaDeW[j,m]=0
MaDeW[1,2]=1.2        
MaDeW[2,4]=2
MaDeW[3,3]=0.7
MaDeW[3,1]=1.8
MaDeW[3,2]=2
MaDeW[4,4]=0.5
MaDeW[5,3]=0.9
MaDeW[6,1]=1.4
MaDeW[7,4]=1.5
MaDeW[8,2]=2.2
MaDeW[9,3]=0.3
MaDeW[10,4]=1.2
MaDeW[11,3]=2.5
MaDeW[12,4]=1.1
MaDeW[13,2]=0.6
MaDeW[14,1]=1

MaDeW[2,3]=1
MaDeW[6,2]=1
MaDeW[4,1]=1
MaDeW[13,4]=1

MaDeW[1,1]=1
MaDeW[8,3]=1
MaDeW[9,2]=1
MaDeW[11,4]=1

MaDeW[2,4]=1
MaDeW[7,3]=1
MaDeW[10,2]=1
MaDeW[12,2]=1
#set the variables
co=Model('collision')

MaSu={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for m in range(1,MaNum):
            for r in range(1,Requests):
                MaSu[r,i,j,m]=co.addVar(vtype=GRB.BINARY)

ReSe={}
for r in range(1,Requests):
    for s in range(1,Sequence):
        for k in range(1,TCpoint):
            ReSe[r,s,k]=co.addVar(vtype=GRB.BINARY)

SuDe={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for m in range(1,MaNum):
            for k in range(1,TCpoint):
                for s in range(1,Sequence):
                    SuDe[s,i,j,m,k]=co.addVar(vtype=GRB.BINARY)
SuDeNoM={}
for i in range(1,Spoint):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
                for s in range(1,Sequence):
                    SuDeNoM[s,i,j,k]=co.addVar(vtype=GRB.BINARY)
DeSu={}
for s in range(1,Sequence):
    for i in range(1,Spoint):
        for j in range(1,Dpoint):
            for k in range(1,TCpoint):
                DeSu[s,j,i,k]=co.addVar(vtype=GRB.BINARY)
DeSe={}
for s in range(1,Sequence):
    for j in range(1,Dpoint):
        for k in range(1,TCpoint):
            DeSe[s,j,k]=co.addVar(vtype=GRB.BINARY)

#timeline={}
#for k in range(1,TCpoint):
#    for s in range(1,Sequence):
#        for q in range(1,Stages):
#            timeline[k,s,q]=co.addVar(vtype=GRB.CONTINUOUS)
                       
judge={}
for k in range(1,TCpoint):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
                for h in range(1,Stages):
                    for o in range(1,OverZone):
                        judge[k,s,q,u,h,o]=co.addVar(vtype=GRB.BINARY)

JuCr={}                  
for s in range(1,Sequence):
    for q in range(1,Stages):
        for k in range(1,TCpoint):
            for o in range(1,OverZone):
                JuCr[k,s,q,o]=co.addVar(vtype=GRB.BINARY)
TcCo={}            
for k in range(1,TCpoint):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
                for h in range(1,Stages):
                    for o in range(1,OverZone):
                        TcCo[k,s,q,u,h,o]=co.addVar(vtype=GRB.BINARY)
judge_a={}                    
for k in range(1,TCpoint):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
                for h in range(1,Stages):
                    for o in range(1,OverZone):
                        judge_a[k,s,q,u,h,o]=co.addVar(vtype=GRB.BINARY)
co.update

#definite the constrains
#choose a supply point to finish requests
for r in range(1,Requests):
    for i in range(1,Spoint):
        for j in range(1,Dpoint):
            for m in range(1,MaNum):
                    co.addConstr(MaStore[i,m]>=MaSu[r,i,j,m])
#                    co.addConstr(MaDe[r,j,m]>=MaSu[r,i,j,m])
for r in range(1,Requests):
    for j in range(1,Dpoint):
        for m in range(1,MaNum):
            co.addConstr(quicksum(MaSu[r,i,j,m] for i in range(1,Spoint))==MaDe[r,j,m])
for r in range(1,Requests):
    co.addConstr(quicksum(ReSe[r,s,k] for s in range(1,Sequence)for k in range(1,TCpoint))==1)
for s in range(1,Sequence):
    for k in range(1,TCpoint):
        co.addConstr(quicksum(ReSe[r,s,k] for r in range(1,Requests))<=1)
        co.addConstr(quicksum(SuDe[s,i,j,m,k] for  i in range(1,Spoint) for j in range(1,Dpoint) for m in range(1,MaNum)) ==1)
        co.addConstr(quicksum(DeSu[s,j,i,k] for  i in range(1,Spoint) for j in range(1,Dpoint)) ==1)
#for r in range(1,Requests):
#    for s in range(1,Sequence):
#        for k in range(1,TCpoint):
#            for u in range(1,Requests):
#                for h in range(1,Sequence):
#                    if s < h and r > u:
#                        #pass
#                        co.addConstr(1 - ReSe[r,s,k] >= ReSe[u,h,k])
#choose a tower crane to finish requests and sequence them
for k in range(1,TCpoint):
    for r in range(1,Requests):
        for i in range(1,Spoint):
            for j in range(1,Dpoint):
                for m in range(1,MaNum):
                    for s in range(1,Sequence):
                        co.addConstr(2-MaSu[r,i,j,m]-ReSe[r,s,k]>=1-SuDe[s,i,j,m,k])
                        co.addConstr(SuDe[s,i,j,m,k]*(70-Dis_STC[i,k])>=0)
                        co.addConstr(SuDe[s,i,j,m,k]*(70-Dis_DTC[j,k])>=0)
                        co.addConstr(2-MaDe[r,j,m]-ReSe[r,s,k]>=1-DeSe[s,j,k])
#every request have just one material to supply
#for k in range(1,TCpoint):
#    for i in range(1,Spoint):
#            for j in range(1,Dpoint):
#                for s in range(1,Sequence):
#                        co.addConstr(quicksum(SuDe[s,i,j,m,k] for m in range(1,MaNum))==SuDeNoM[s,i,j,k])                       
#choose a supply point for the next request
for k in range(1,TCpoint):
    for s in range(1,Sequence-1):    
        for i in range(1,Spoint):
            for j in range(1,Dpoint):
                for o in range(1,Dpoint):
                    for m in  range(1,MaNum):
                        co.addConstr(2-DeSe[s,j,k]-SuDe[s+1,i,o,m,k]>=1-DeSu[s+1,j,i,k])

#choose a supply point from the initial position
DeSe_1={}
for j in range(1,Dpoint):
    for k in range(1,TCpoint):
        DeSe_1[j,k]=0     
DeSe_1[5,1]=1
DeSe_1[2,2]=1
DeSe_1[6,3]=1
DeSe_1[11,4]=1
for s in range(1,Sequence):
    for i in range(1,Spoint):
        for j in range(1,Dpoint):
            for o in range(1,Dpoint):
                for k in range(1,TCpoint):
                    for m in  range(1,MaNum):
                        if s==1:
                            co.addConstr(2-SuDe[s,i,o,m,k]-DeSe_1[j,k]>=1-DeSu[s,j,i,k])
                        #co.addConstr(DeSe_1[j,k]>=DeSu[s,j,i,k])
#choose no more than one supply point for the next request
#for s in range(1,Sequence):
#    for j in range(1,Dpoint):
#        for k in range(1,TCpoint):
#            co.addConstr(quicksum(DeSu[s,j,i,k] for i in range(1,Spoint))<=1)
            
###calculate the start time to every supply point and demand point 
timeline={}
for k in range(1,TCpoint):
    for s in range(1,Sequence):
        for q in range(1,Stages): #A whole Sequence have two stages
        #start time
            if s==1:
                if q==1:
                    timeline[k,s,q]=0
                if q==2:
                    timeline[k,s,q]=timeline[k,s,q-1]+quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k] for i in range(1,Spoint) for j in range(1,Dpoint))
            if s!=1:
                if q==1:
                    timeline[k,s,q]=timeline[k,s-1,q+1]+quicksum((Time_sd_run[i,j,k]+1)*SuDe[s-1,i,j,m,k] for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) 
                if q==2:
                    timeline[k,s,q]=timeline[k,s,q-1]+quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k] for i in range(1,Spoint) for j in range(1,Dpoint))
Time_In={}
Time_Out={}
#calculate the enter and leave time when demand location exist in overlapping zone 
def calculate_process(k,o,p):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            if q==1:
                if s==1:
                    Time_In[k,s,q,o] = timeline[k,s,q] + quicksum(DeSu[s,j,i,k]*(1-Path_C[j,o])*(-100*(1-Path[i,j,k,o])+(Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])) for i in range(1,Spoint) for j in range(1,Dpoint))
                    Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(DeSu[s,j,i,k]*(1-Path_C[j,o])*Path[i,j,k,o]*Time_AB[k,o]for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum(DeSu[s,j,i,k]*Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k]))for i in range(1,Spoint) for j in range(1,Dpoint))
                else:
                    Time_In[k,s,q,o] = timeline[k,s,q] + quicksum(DeSu[s,j,i,k]*(1-Path_C[j,o])*(-100*(1-Path[i,j,k,o])+(Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k]))for i in range(1,Spoint) for j in range(1,Dpoint)) - quicksum(SuDe[s-1,i,j,m,k]*Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])+1) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum))
                    Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(DeSu[s,j,i,k]*(1-Path_C[j,o])*(Path[i,j,k,o]*Time_AB[k,o])for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum(DeSu[s,j,i,k]*Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k]))for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum(SuDe[s-1,i,j,m,k]*Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])+1) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum))
            else:
                Time_In[k,s,q,o] = timeline[k,s,q] + quicksum(SuDe[s,i,j,m,k]*(-100*(1-Path[i,j,k,o])+(Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum))
                if s==Sequence-1:
                    Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(SuDe[s,i,j,m,k]*((1-Path_C[j,o])*Path[i,j,k,o]*Time_AB[k,o] + 50*Path_C[j,o]) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum))
                else:
                    Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(SuDe[s,i,j,m,k]*((1-Path_C[j,o])*Path[i,j,k,o]*Time_AB[k,o]) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) + quicksum(SuDe[s,i,j,m,k]*(Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])+1)) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) + quicksum(DeSu[s+1,j,i,k]*(Path_C[j,o]*((Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])))  for i in range(1,Spoint) for j in range(1,Dpoint))
#calculate the enter and leave time when supply location exist in overlapping zone 
def calculate_process2(k,o,p): 
    for s in range(1,Sequence):
        for q in range(1,Stages):
            if q==1:
                Time_In[k,s,q,o] = timeline[k,s,q] + quicksum(DeSu[s,j,i,k]*(-100*(1-Path[i,j,k,o])+(Path_A[i,j,k,o,p]*Time_DC[p,j,k])+(Path_A[i,j,k,o,p+1]*Time_DC[p+1,j,k])) for i in range(1,Spoint) for j in range(1,Dpoint))
                Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(DeSu[s,j,i,k]*(Path_D[i,o]*((Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])+1))for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum(DeSu[s,j,i,k]*(1-Path_D[i,o])*Path[i,j,k,o]*Time_AB[k,o]for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum(SuDe[s,i,j,m,k]*Path_D[i,o]*((Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum))
            else:
                Time_In[k,s,q,o] = timeline[k,s,q] + quicksum(SuDe[s,i,j,m,k]*(1-Path_D[i,o])*(-100*(1-Path[i,j,k,o])+(Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) - quicksum(DeSu[s,j,i,k]*Path_D[i,o]*((Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])+1)for i in range(1,Spoint) for j in range(1,Dpoint))
                Time_Out[k,s,q,o] = Time_In[k,s,q,o] + quicksum(SuDe[s,i,j,m,k]*((1-Path_D[i,o])*Path[i,j,k,o]*Time_AB[k,o]) for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) + quicksum(SuDe[s,i,j,m,k]*Path_D[i,o]*((Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k]))for i in range(1,Spoint) for j in range(1,Dpoint)for m in  range(1,MaNum)) + quicksum(DeSu[s,j,i,k]*Path_D[i,o]*((Path_A[i,j,k,o,p]*Time_SC[p,i,k])+(Path_A[i,j,k,o,p+1]*Time_SC[p+1,i,k])+1)for i in range(1,Spoint) for j in range(1,Dpoint))

def judge_equation(k,o,p):
    if sum(Path_D[i,o] for i in range(1,Spoint))>=1 :
        calculate_process2(k,o,p)
    else:
        calculate_process(k,o,p)
        
for k in range(1,TCpoint):
    for p in range(1,Cpoint):
        for o in range(1,OverZone):
            if k==1 or k==2:
                if o==1:
                    if p==1:
                       judge_equation(k,o,p) 
            if k==1 or k==3:                  
                if o==2:
                    if p==3:
                        judge_equation(k,o,p)         
            if k==2 or k==3:
                if o==3:
                    if p==5:
                        judge_equation(k,o,p)
            if k==2 or k==4:
                if o==4:
                    if p==7:
                        judge_equation(k,o,p)
                        
#for k in range(1,TCpoint):
#    for s in range(1,Sequence):
#        for q in range(1,Stages):
#            for o in range(1,OverZone):
#                if k==1 or k==2:
#                    if o==1:
#                        co.addConstr(1000*JuCr[k,s,q,o]>=Time_In[k,s,q,o]-100)
#                        co.addConstr(1000*(JuCr[k,s,q,o]-1)<=Time_In[k,s,q,o]-100)
#                if k==1 or k==3:
#                    if o==2:
#                        co.addConstr(1000*JuCr[k,s,q,o]>=Time_In[k,s,q,o]-100)
#                        co.addConstr(1000*(JuCr[k,s,q,o]-1)<=Time_In[k,s,q,o]-100)
#                if k==2 or k==3:
#                    if o==3:
#                        co.addConstr(1000*JuCr[k,s,q,o]>=Time_In[k,s,q,o]-100)
#                        co.addConstr(1000*(JuCr[k,s,q,o]-1)<=Time_In[k,s,q,o]-100)        
#                if k==2 or k==4:
#                    if o==4:
#                        co.addConstr(1000*JuCr[k,s,q,o]>=Time_In[k,s,q,o]-100)
#                        co.addConstr(1000*(JuCr[k,s,q,o]-1)<=Time_In[k,s,q,o]-100) 
#                       
#for k in range(1,TCpoint-1):
#    for s in range(1,Sequence):
#        for q in range(1,Stages):
#            for u in range(1,Sequence):
#                for h in range(1,Stages):
#                    for o in range(1,OverZone):                      
#                        if k==1:
#                            if o==1:
#                                co.addConstr(JuCr[k,s,q,o]+JuCr[k+1,u,h,o]>=1-TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k,s,q,o]>=TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k+1,u,h,o]>=TcCo[k,s,q,u,h,o])
#                        if k==1:
#                            if o==2:
#                                co.addConstr(JuCr[k,s,q,o]+JuCr[k+2,u,h,o]>=1-TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k,s,q,o]>=TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k+2,u,h,o]>=TcCo[k,s,q,u,h,o])
#                        if k==2:
#                            if o==3:
#                                co.addConstr(JuCr[k,s,q,o]+JuCr[k+1,u,h,o]>=1-TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k,s,q,o]>=TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k+1,u,h,o]>=TcCo[k,s,q,u,h,o])
#                        if k==2:
#                            if o==4:
#                                co.addConstr(JuCr[k,s,q,o]+JuCr[k+2,u,h,o]>=1-TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k,s,q,o]>=TcCo[k,s,q,u,h,o])
#                                co.addConstr(1-JuCr[k+2,u,h,o]>=TcCo[k,s,q,u,h,o])
Time_wait={}
for k in range(1,TCpoint-1):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
               for h in range(1,Stages):
                    for o in range(1,OverZone):
                        if k==1:
                            if o==1:
                                Time_wait[k,s,q,u,h,o]=Time_Out[k+1,u,h,o]-Time_In[k,s,q,o] #+0.5
                                Time_wait[k+1,u,h,s,q,o]=Time_Out[k,s,q,o]-Time_In[k+1,u,h,o] #+0.5
                        if k==1:
                            if o==2:
                                Time_wait[k,s,q,u,h,o]=Time_Out[k+2,u,h,o]-Time_In[k,s,q,o] #+0.5
                                Time_wait[k+2,u,h,s,q,o]=Time_Out[k,s,q,o]-Time_In[k+2,u,h,o] #+0.5
                        if k==2:
                            if o==3:
                                Time_wait[k,s,q,u,h,o]=Time_Out[k+1,u,h,o]-Time_In[k,s,q,o] #+0.5
                                Time_wait[k+1,u,h,s,q,o]=Time_Out[k,s,q,o]-Time_In[k+1,u,h,o] #+0.5
                        if k==2:
                            if o==4:
                                Time_wait[k,s,q,u,h,o]=Time_Out[k+2,u,h,o]-Time_In[k,s,q,o] #+0.5
                                Time_wait[k+2,u,h,s,q,o]=Time_Out[k,s,q,o]-Time_In[k+2,u,h,o] #+0.5
for k in range(1,TCpoint):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
                for h in range(1,Stages):
                    for o in range(1,OverZone):                        
                        #if wait_k1>0,judge=1;else,judge=0
                        if k==1 or k==2:
                            if o==1:
                                co.addConstr(Time_wait[k,s,q,u,h,o]>=(judge[k,s,q,u,h,o]-1)*300)
                                co.addConstr(Time_wait[k,s,q,u,h,o]<=200*judge[k,s,q,u,h,o])
                        if k==1 or k==3:
                            if o==2:
                                co.addConstr(Time_wait[k,s,q,u,h,o]>=(judge[k,s,q,u,h,o]-1)*300)
                                co.addConstr(Time_wait[k,s,q,u,h,o]<=200*judge[k,s,q,u,h,o])        
                        if k==2 or k==3:
                            if o==3:
                                co.addConstr(Time_wait[k,s,q,u,h,o]>=(judge[k,s,q,u,h,o]-1)*300)
                                co.addConstr(Time_wait[k,s,q,u,h,o]<=200*judge[k,s,q,u,h,o])
                        if k==2 or k==4:
                            if o==4:
                                co.addConstr(Time_wait[k,s,q,u,h,o]>=(judge[k,s,q,u,h,o]-1)*300)
                                co.addConstr(Time_wait[k,s,q,u,h,o]<=200*judge[k,s,q,u,h,o])      
for k in range(1,TCpoint-1):
    for s in range(1,Sequence):
        for q in range(1,Stages):
            for u in range(1,Sequence):
                for h in range(1,Stages):
                    for o in range(1,OverZone):
                        if k==1:
                            if o==1:
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k,s,q,u,h,o]>=1-judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(judge[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k+1,u,h,s,q,o]>=1-judge_a[k+1,u,h,s,q,o])
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k+1,u,h,s,q,o])
#                                co.addConstr(judge[k+1,u,h,s,q,o]>=judge_a[k+1,u,h,s,q,o])
#                                 pass
#                                co.addConstr(1-judge_a[k,s,q,u,h,o]-judge_a[k+1,u,h,s,q,o]>=0)
                                co.addConstr(1-judge[k,s,q,u,h,o]-judge[k+1,u,h,s,q,o]>=0)
                        if k==1:
                            if o==2:
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k,s,q,u,h,o]>=1-judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(judge[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k+2,u,h,s,q,o]>=1-judge_a[k+2,u,h,s,q,o])
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k+2,u,h,s,q,o])
#                                co.addConstr(judge[k+2,u,h,s,q,o]>=judge_a[k+2,u,h,s,q,o])
#                                 pass
#                                co.addConstr(1-judge_a[k,s,q,u,h,o]-judge_a[k+2,u,h,s,q,o]>=0)
                                co.addConstr(1-judge[k,s,q,u,h,o]-judge[k+2,u,h,s,q,o]>=0)
                        if k==2:
                            if o==3:
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k,s,q,u,h,o]>=1-judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(judge[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k+1,u,h,s,q,o]>=1-judge_a[k+1,u,h,s,q,o])
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k+1,u,h,s,q,o])
#                                co.addConstr(judge[k+1,u,h,s,q,o]>=judge_a[k+1,u,h,s,q,o])
#                                 pass
#                                co.addConstr(1-judge_a[k,s,q,u,h,o]-judge_a[k+1,u,h,s,q,o]>=0)
                                co.addConstr(1-judge[k,s,q,u,h,o]-judge[k+1,u,h,s,q,o]>=0)
                        if k==2:
                            if o==4:
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k,s,q,u,h,o]>=1-judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(judge[k,s,q,u,h,o]>=judge_a[k,s,q,u,h,o])                        
#                                co.addConstr(2-TcCo[k,s,q,u,h,o]-judge[k+2,u,h,s,q,o]>=1-judge_a[k+2,u,h,s,q,o])
#                                co.addConstr(TcCo[k,s,q,u,h,o]>=judge_a[k+2,u,h,s,q,o])
#                                co.addConstr(judge[k+2,u,h,s,q,o]>=judge_a[k+2,u,h,s,q,o])
#                                 pass
#                                co.addConstr(1-judge_a[k,s,q,u,h,o]-judge_a[k+2,u,h,s,q,o]>=0)
                                co.addConstr(1-judge[k,s,q,u,h,o]-judge[k+2,u,h,s,q,o]>=0)

#co.addConstr(ReSe[4,1,1]==1)
#co.addConstr(ReSe[7,2,1]==1)
#co.addConstr(ReSe[11,3,1]==1)
#co.addConstr(ReSe[12,4,1]==1)
#
#
#co.addConstr(ReSe[1,1,2]==1)
#co.addConstr(ReSe[5,2,2]==1)
#co.addConstr(ReSe[6,3,2]==1)
#co.addConstr(ReSe[8,4,2]==1)
#
#
#co.addConstr(ReSe[9,1,3]==1)
#co.addConstr(ReSe[10,2,3]==1)
#co.addConstr(ReSe[15,3,3]==1)
#co.addConstr(ReSe[16,4,3]==1)
#
#
#co.addConstr(ReSe[2,1,4]==1)
#co.addConstr(ReSe[3,2,4]==1)
#co.addConstr(ReSe[13,3,4]==1)
#co.addConstr(ReSe[14,4,4]==1)

SDT={}
DST={}
time={}
for k in range(1,TCpoint):
    DST[k]=3*quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k] for s in range(1,Sequence) for i in range(1,Spoint) for j in range(1,Dpoint))
    SDT[k]=6*quicksum((Time_sd_run[i,j,k]+1)*SuDe[s,i,j,m,k] for i in range(1,Spoint) for j in range(1,Dpoint) for s in range(1,Sequence) for m in range(1,MaNum))
    time[k]=SDT[k]+DST[k]
Time = quicksum(time[k] for k in range(1,TCpoint))

co.Params.IntFeasTol = 1e-2
#co.Params.TimeLimit = 2000
#co.Params.Presolve = 2 #presolve setting
#co.Params.MIPFocus = 3 #solve setting
#co.Params.MIPGap = 0.05 #solve setting
#co.Params.Heuristics = 0.1 #Determines the amount of time spent in MIP heuristics
co.setObjective(Time*100,GRB.MINIMIZE)

co.optimize()

if co.status==GRB.status.OPTIMAL or co.status==GRB.status.TIME_LIMIT:
    print ('\nThe best solution cost is',' %.2f' % (co.objval/100),'CNY')
    #print ('---------------------') 
    #for r in range(1,Requests):
    #    for i in range(1,Spoint):
    #        for j in range(1,Dpoint):
    #            for m in range(1,MaNum):
    #                if MaSu[r,i,j,m].x==1:
    #                    print (r,i,j,m )
    print ('---------------------' )
    print('request, Sequence, Tower crane')  
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for r in range(1,Requests):
                if ReSe[r,s,k].x > 0.9:
                    print (r,s,k,ReSe[r,s,k].x)
    print ('---------------------')  
    print('Sequence, Supply location, Demand location, Tower crane')  
    for k in range(1,TCpoint): 
        for s in range(1,Sequence):
            for j in range(1,Dpoint):
                for i in range(1,Spoint):
                    for m in range(1,MaNum):
                        if SuDe[s,i,j,m,k].x > 0.9:
                            #print  'In sequence',s, 'TC',k, 'travel from supply location',i, 'to demand location',j,'with material',m,'and the time is',Time_sd_run[i,j,k]   
                            print ('[',s,i,j,m,k,']','%.2f' %((Time_sd_run[i,j,k]+1)*SuDe[s,i,j,m,k].x*6))
#                            print ('[',s,i,j,m,k,']','%.2f' %Time_sd_run[i,j,k])
                            #print Time_sd_run[i,j,k]    
    print ('-----------------------------------')
    print('Sequence, Demand location, Supply location, Tower crane') 
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for j in range(1,Dpoint):
                for i in range(1,Spoint):
                    if DeSu[s,j,i,k].x > 0.9:
                        #pass
                        #print 'In sequence',s, 'TC',k,' travel from demand location',j, 'to supply location',i,'and the time is',Time_sd_run[i,j,k]
                        print  ('[',s,j,i,k,']','%.2f' %((Time_sd_run[i,j,k]+1)*3*DeSu[s,j,i,k].x))
#                        print  ('[',s,j,i,k,']','%.2f' %Time_sd_run[i,j,k])
                        #print  Time_sd_run[i,j,k]
    print ('---------------------')
    print(quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k].x for s in range(1,Sequence) for i in range(1,Spoint) for j in range(1,Dpoint) for k in range(1,TCpoint)))
    print(quicksum((Time_sd_run[i,j,k]+1)*SuDe[s,i,j,m,k].x for i in range(1,Spoint) for j in range(1,Dpoint) for s in range(1,Sequence)for m in range(1,MaNum) for k in range(1,TCpoint)))
    print(quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k].x for s in range(1,Sequence) for i in range(1,Spoint) for j in range(1,Dpoint) for k in range(1,TCpoint)) + quicksum((Time_sd_run[i,j,k]+1)*SuDe[s,i,j,m,k].x for i in range(1,Spoint) for j in range(1,Dpoint) for s in range(1,Sequence)for m in range(1,MaNum) for k in range(1,TCpoint)))
    print ('---------------------')
    for k in range(1,TCpoint):
        print(k, quicksum((Time_sd_run[i,j,k]+1)*DeSu[s,j,i,k].x*3 for s in range(1,Sequence) for i in range(1,Spoint) for j in range(1,Dpoint)) + quicksum((Time_sd_run[i,j,k]+1)*SuDe[s,i,j,m,k].x*6 for i in range(1,Spoint) for j in range(1,Dpoint) for s in range(1,Sequence)for m in range(1,MaNum) ))
    print('------------------') 
    print('1# and 2#')
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for q in range(1,Stages):
                for o in range(1,OverZone):
                    if k==1 or k==2:
                        if o==1:
                            print(k,s,q,o,Time_In[k,s,q,o].getValue(),Time_Out[k,s,q,o].getValue())
    print('------------------') 
    print('1# and 3#')
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for q in range(1,Stages):
                for o in range(1,OverZone):
                    if k==1 or k==3:
                        if o==2:
                            print(k,s,q,o,Time_In[k,s,q,o].getValue(),Time_Out[k,s,q,o].getValue())
    print('------------------') 
    print('2# and 3#')
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for q in range(1,Stages):
                for o in range(1,OverZone):
                    if k==2 or k==3:
                        if o==3:
                            print(k,s,q,o,Time_In[k,s,q,o].getValue(),Time_Out[k,s,q,o].getValue())
    print('------------------') 
    print('2# and 4#')
    for k in range(1,TCpoint):
        for s in range(1,Sequence):
            for q in range(1,Stages):
                for o in range(1,OverZone):
                    if k==2 or k==4:
                        if o==4:
                            print(k,s,q,o,Time_In[k,s,q,o].getValue(),Time_Out[k,s,q,o].getValue())                           
    #print ('---------------------')
    #for k in range(1,TCpoint-1):
    #    for s in range(1,Sequence):
    #        for q in range(1,Stages):
    #            for u in range(1,Sequence):
    #                for h in range(1,Stages):
    #                    for o in range(1,OverZone):
    #                        if k==1:
    #                            if o==1:
    #                                Time_wait2[k,s,q,u,h,o]=Time_Out2[k+1,u,h,o]-Time_In2[k,s,q,o]
    #                                Time_wait2[k+1,u,h,s,q,o]=Time_Out2[k,s,q,o]-Time_In2[k+1,u,h,o]
    #                        if k==1:
    #                            if o==2:
    #                                Time_wait2[k,s,q,u,h,o]=Time_Out2[k+2,u,h,o]-Time_In2[k,s,q,o]
    #                                Time_wait2[k+2,u,h,s,q,o]=Time_Out2[k,s,q,o]-Time_In2[k+2,u,h,o]
    #                                print('-------------')
    #                                print(k,s,q,u,h,o,Time_wait[k,s,q,u,h,o].getValue(),judge[k,s,q,u,h,o].x)
    #                                print(k+2,u,h,s,q,o,Time_wait[k+2,u,h,s,q,o].getValue(),judge[k+2,u,h,s,q,o].x)
    #                        if k==2:
    #                            if o==3:
    #                                Time_wait2[k,s,q,u,h,o]=Time_Out2[k+1,u,h,o]-Time_In2[k,s,q,o]
    #                                Time_wait2[k+1,u,h,s,q,o]=Time_Out2[k,s,q,o]-Time_In2[k+1,u,h,o]
    #                        if k==2:
    #                            if o==4:
    #                                Time_wait2[k,s,q,u,h,o]=Time_Out2[k+2,u,h,o]-Time_In2[k,s,q,o]
    #                                Time_wait2[k+2,u,h,s,q,o]=Time_Out2[k,s,q,o]-Time_In2[k+2,u,h,o]
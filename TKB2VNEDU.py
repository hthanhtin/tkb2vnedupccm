import  pandas as pd
df=pd.read_excel("TKB_PCCM.xlsx", sheet_name="LOPVNEDU")
df=df.to_dict(orient='list')
loptkb2vnedu=dict(zip(df["LopTKB"],df["LopVNEDU"]))

df=pd.read_excel("TKB_PCCM.xlsx", sheet_name="MONVNEDU")
df=df.to_dict(orient='list')
montkb2vnedu=dict(zip(df["MonTKB"],df["MonVNEDU"]))

df=pd.read_excel("TKB_PCCM.xlsx", sheet_name="PCGD_Mau1")
pccm=df.to_dict(orient='records')
def tenlop(s):
    s=s.strip()
    return loptkb2vnedu[s[:len(s)-3]]
pc=[]
for gv in pccm:
    x=gv['DSLOPDAY'].split(",")
    y=[tenlop(i) for i in x ]
    pc.append((gv["GIAOVIEN"],f"{gv["MONVNEDU"]} : {", ".join(y)}"))

lopsgv=dict()
for x in pc:
    lopsgv[x[0]]=""

for x in pc:
    if len(lopsgv[x[0]])==0:
        lopsgv[x[0]]=x[1]
    else:
        lopsgv[x[0]]=lopsgv[x[0]]+"; "+x[1]
a=lopsgv.keys()
b=lopsgv.values()
data={
    "GV": a,
    "Lop": b
}
dt=pd.DataFrame(data)
dt.to_excel("VNEDU.xlsx")

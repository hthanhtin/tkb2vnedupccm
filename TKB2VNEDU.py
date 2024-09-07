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
def vnedu_pccm():
    pc=[]
    a=[]
    b=[]
    c=[]
    for gv in pccm:
        x=gv['DSLOPDAY'].split(",")
        y=[tenlop(i) for i in x ]
        a.append(gv["GIAOVIEN"])
        b.append(gv["MONVNEDU"])
        c.append(", ".join(y))

    data={
        "GV": a,
        "Mon":b,
        "Lopday": c

    }
    try:
        dt=pd.DataFrame(data)
        dt.to_excel("VNEDU.xlsx")
    except:
        with open("vnedu_pccm.txt", "w", encoding="utf8") as f:
            s1 = f"{len(a)} {" ".join(a)}"+"\n"
            s2 = f"{len(b)} {" ".join(b)}" + "\n"
            s3 = f"{len(c)} {" ".join(c)}" + "\n"
            f.write(s1+s2+s3)

def lop_ppcm():
    pccm_lop=dict()
    for k in loptkb2vnedu.keys():
        pccm_lop[k]=[]

    for x in pccm:
        lops=x["DSLOPDAY"].split(",")
        for l in lops:
            l=l.strip()
            l=l[0:len(l)-3]
            pccm_lop[l].append(x["MONVNEDU"]+"-"+x["GIAOVIEN"])
    for k in pccm_lop.keys():
        l=pccm_lop[k]
        l.sort()
        pccm_lop[k]=l
    try:
        sv=pd.DataFrame(pccm_lop)
        sv.to_excel("pccm_lop.xlsx")
    except:
        with open("log_ppcm_lop.txt","w", encoding="utf8") as f:
            for k in pccm_lop.keys():
                l = pccm_lop[k]
                s=f"{len(l)};{k};{";".join(l)}"+"\n"
                f.write(s)




lop_ppcm()
vnedu_pccm()
import io
import re
import unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st

# ======================================================
# CONFIG
# ======================================================
FEC_COLUMNS = [
    "JournalCode","JournalLib","EcritureNum","EcritureDate",
    "CompteNum","CompteLib","CompAuxNum","CompAuxLib",
    "PieceRef","PieceDate","EcritureLib",
    "Debit","Credit",
    "EcritureLet","DateLet",
    "ValidDate","Montantdevise","Idevise"
]

# ======================================================
# OUTILS ROBUSTES
# ======================================================
def sstr(x):
    try:
        if pd.isna(x):
            return ""
    except:
        pass
    return str(x).strip()

def strip_accents(s):
    return "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )

def norm_key(s):
    s = strip_accents(sstr(s)).lower()
    s = s.replace("�","")
    s = re.sub(r"[^a-z0-9 ]"," ",s)
    return re.sub(r"\s+"," ",s).strip()

def parse_fr_number(x):
    s = sstr(x).replace(" ","").replace(",",".")
    try:
        return float(s)
    except:
        return 0.0

def parse_date_any(x):
    s = sstr(x)
    if " - " in s:
        s = s.split(" - ")[0]
    for f in ("%d/%m/%Y","%Y-%m-%d"):
        try:
            return datetime.strptime(s,f).strftime("%Y%m%d")
        except:
            pass
    try:
        return pd.to_datetime(s,dayfirst=True).strftime("%Y%m%d")
    except:
        return ""

def clean_invoice_no(x):
    return re.sub(r"[^\w\-\/]","",sstr(x))

def join_nonempty(lst):
    return " | ".join(sstr(x) for x in lst if sstr(x))

def fec_row(*args):
    row = dict(zip(FEC_COLUMNS,args))
    row["Debit"]  = f"{float(row['Debit']):.2f}"
    row["Credit"] = f"{float(row['Credit']):.2f}"
    return row

def read_any(file):
    data = file.getvalue()
    name = file.name.lower()
    if name.endswith(("xls","xlsx")):
        return pd.read_excel(io.BytesIO(data),dtype=str)
    txt = data.decode("utf-8",errors="replace")
    for sep in [";","|","\t",","]:
        try:
            df = pd.read_csv(io.StringIO(txt),sep=sep,dtype=str)
            if df.shape[1]>3:
                return df
        except: pass
    return pd.read_csv(io.StringIO(txt),dtype=str)

def to_fec(df,sep):
    out = io.StringIO()
    df[FEC_COLUMNS].fillna("").astype(str).to_csv(
        out,sep=sep,index=False,lineterminator="\n"
    )
    return out.getvalue().encode("utf-8")

# ======================================================
# UI
# ======================================================
st.set_page_config("DELPHY - Export FEC Opticien",layout="wide")
st.title("DELPHY - Export FEC — Factures / Encaissements")

with st.sidebar:
    st.subheader("Comptes")
    cpt_vente = st.text_input("707 Ventes", "7071")
    cpt_tva   = st.text_input("44571 TVA", "44571")
    cpt_cli   = st.text_input("411 Clients", "FCLIENTS")

    st.subheader("Caisse (classe 53)")
    cpt_esp = st.text_input("Espèces", "531")
    cpt_chq = st.text_input("Chèque", "532")
    cpt_cb  = st.text_input("Carte bleue", "533")
    cpt_aut = st.text_input("Autre / Virement", "534")

    st.subheader("Journaux")
    j_ve = st.text_input("Ventes", "VE")
    j_cs = st.text_input("Encaissements", "CS")

    st.subheader("Lettrage")
    use_aux = st.checkbox("CompAux = ID client", True)
    use_auxlib = st.checkbox("Nom client en CompAuxLib", True)

    sep = st.selectbox("Séparateur FEC",["|",";","\t"])

# ======================================================
# FICHIERS
# ======================================================
col1,col2 = st.columns(2)
with col1:
    fact_file = st.file_uploader("Factures (CSV/XLSX)",type=["csv","xls","xlsx"])
with col2:
    pay_file = st.file_uploader("Encaissements (CSV/XLSX)",type=["csv","xls","xlsx"])

if not fact_file and not pay_file:
    st.stop()

fec = []
ecriture = 1

def aux(cid,cname):
    if not use_aux:
        return "",""
    return sstr(cid)[:40], (sstr(cname)[:200] if use_auxlib else "")

# ======================================================
# FACTURES
# ======================================================
if fact_file:
    df = read_any(fact_file)
    df.columns = [norm_key(c) for c in df.columns]

    for _,r in df.iterrows():
        ht  = parse_fr_number(r.get("montant ht"))
        tva = parse_fr_number(r.get("tva"))
        ttc = abs(ht+tva)

        fact = clean_invoice_no(r.get("n fact avoir"))
        date = parse_date_any(r.get("date"))

        cid  = r.get("id client")
        cname= r.get("client")
        ax,al= aux(cid,cname)

        lib = join_nonempty(["Facture",fact,cname])

        fec += [
            fec_row(j_ve,"VENTES",ecriture,date,
                    cpt_cli,"Clients",ax,al,
                    f"FAC-{fact}",date,lib,
                    ttc,0,fact,"","",""),
            fec_row(j_ve,"VENTES",ecriture,date,
                    cpt_vente,"Ventes","","",
                    f"FAC-{fact}",date,lib,
                    0,abs(ht),fact,"","","")
        ]
        if tva:
            fec.append(
                fec_row(j_ve,"VENTES",ecriture,date,
                        cpt_tva,"TVA","","",
                        f"FAC-{fact}",date,lib,
                        0,abs(tva),fact,"","","")
            )
        ecriture+=1

# ======================================================
# ENCAISSEMENTS
# ======================================================
if pay_file:
    df = read_any(pay_file)
    df.columns = [norm_key(c) for c in df.columns]

    for _,r in df.iterrows():
        amt = parse_fr_number(r.get("montant"))
        if amt==0: continue

        mode = sstr(r.get("rgt"))
        if "cb" in mode.lower(): cpt = cpt_cb
        elif "chq" in mode.lower(): cpt = cpt_chq
        elif "esp" in mode.lower(): cpt = cpt_esp
        else: cpt = cpt_aut

        date = parse_date_any(r.get("date saisie"))
        fact = clean_invoice_no(r.get("fact"))
        cid  = r.get("id client")
        cname= r.get("client")
        ax,al= aux(cid,cname)

        lib = join_nonempty(["Encaissement",fact,cname,mode])

        fec += [
            fec_row(j_cs,"CAISSE",ecriture,date,
                    cpt,"Caisse","","",
                    f"CS-{ecriture}",date,lib,
                    abs(amt),0,fact,"","",""),
            fec_row(j_cs,"CAISSE",ecriture,date,
                    cpt_cli,"Clients",ax,al,
                    f"CS-{ecriture}",date,lib,
                    0,abs(amt),fact,"","","")
        ]
        ecriture+=1

# ======================================================
# EXPORT
# ======================================================
df_fec = pd.DataFrame(fec,columns=FEC_COLUMNS)
st.dataframe(df_fec.head(200),use_container_width=True)

st.download_button(
    "📥 Télécharger le FEC",
    data=to_fec(df_fec,sep),
    file_name="FEC_DELPHY.txt",
    mime="text/plain"
)

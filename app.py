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
    s = s.replace("�", "")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def parse_fr_number(x):
    s = sstr(x).replace("\u00A0", " ").replace(" ", "")
    # 1 234,56 -> 1234.56
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    # 1.234,56 -> 1234.56
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def parse_date_any(x):
    s = sstr(x)
    if s == "":
        return ""
    if " - " in s:
        s = s.split(" - ")[0].strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, f).strftime("%Y%m%d")
        except:
            pass
    try:
        d = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return ""
        return d.strftime("%Y%m%d")
    except:
        return ""

def clean_invoice_no(x):
    return re.sub(r"[^\w\-\/]", "", sstr(x))

def join_nonempty(lst):
    vals = []
    for x in lst:
        t = sstr(x)
        if t and t.lower() != "nan":
            vals.append(t)
    return " | ".join(vals)

def fec_row(*args):
    row = dict(zip(FEC_COLUMNS, args))
    # sécurise Debit/Credit
    row["Debit"]  = f"{float(row['Debit']):.2f}"
    row["Credit"] = f"{float(row['Credit']):.2f}"
    # sécurité sur lettrage: toujours vide
    row["EcritureLet"] = ""
    row["DateLet"] = ""
    return row

def read_any(file):
    data = file.getvalue()
    name = file.name.lower()
    if name.endswith(("xls", "xlsx")):
        return pd.read_excel(io.BytesIO(data), dtype=str)

    txt = data.decode("utf-8", errors="replace")
    for sep in [";", "|", "\t", ","]:
        try:
            df = pd.read_csv(io.StringIO(txt), sep=sep, dtype=str)
            if df.shape[1] > 3:
                return df
        except:
            pass
    return pd.read_csv(io.StringIO(txt), dtype=str)

def to_fec(df, sep):
    out = io.StringIO()
    df2 = df.copy()
    # enforce columns + fill
    df2 = df2.reindex(columns=FEC_COLUMNS).fillna("")
    # forcer lettrage vide
    df2["EcritureLet"] = ""
    df2["DateLet"] = ""
    df2.astype(str).to_csv(out, sep=sep, index=False, lineterminator="\n")
    return out.getvalue().encode("utf-8")

def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FEC")
    return buf.getvalue()

# ======================================================
# UI
# ======================================================
st.set_page_config("DELPHY - Export FEC Opticien", layout="wide")
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

    st.subheader("Auxiliaires")
    use_aux = st.checkbox("CompAux = ID client", True)
    use_auxlib = st.checkbox("Nom client en CompAuxLib", True)

    sep = st.selectbox("Séparateur FEC", ["|", ";", "\t"], index=0)

# ======================================================
# FICHIERS
# ======================================================
col1, col2 = st.columns(2)
with col1:
    fact_file = st.file_uploader("Factures (CSV/XLSX) — optionnel", type=["csv", "xls", "xlsx"])
with col2:
    pay_file = st.file_uploader("Encaissements (CSV/XLSX) — optionnel", type=["csv", "xls", "xlsx"])

if not fact_file and not pay_file:
    st.info("Charge au moins un fichier.")
    st.stop()

fec = []
ecriture = 1

def aux(cid, cname):
    if not use_aux:
        return "", ""
    ax = sstr(cid)[:40]
    al = sstr(cname)[:200] if use_auxlib else ""
    return ax, al

# ======================================================
# FACTURES
# ======================================================
if fact_file:
    df = read_any(fact_file)
    raw_cols = list(df.columns)
    df.columns = [norm_key(c) for c in df.columns]

    # mapping minimal pour ton export
    # (on tolère les variations)
    col_date = "date" if "date" in df.columns else None
    col_ht = "montant ht" if "montant ht" in df.columns else None
    col_tva = "tva" if "tva" in df.columns else None
    col_fact = None
    for c in df.columns:
        if "fact" in c and ("avoir" in c or "fact" in c):
            col_fact = c
            break
    col_cid = "id client" if "id client" in df.columns else None
    col_cname = "client" if "client" in df.columns else None

    for _, r in df.iterrows():
        ht  = parse_fr_number(r.get(col_ht)) if col_ht else 0.0
        tva = parse_fr_number(r.get(col_tva)) if col_tva else 0.0

        # si avoir (montants négatifs) : on garde le sens via abs mais tu peux ajuster plus tard
        ttc = abs(ht + tva)

        fact = clean_invoice_no(r.get(col_fact)) if col_fact else ""
        date = parse_date_any(r.get(col_date)) if col_date else ""

        cid  = r.get(col_cid) if col_cid else ""
        cname = r.get(col_cname) if col_cname else ""
        ax, al = aux(cid, cname)

        lib = join_nonempty(["Facture", fact, cname])

        piece_ref = f"FAC-{fact}" if fact else f"FAC-{ecriture}"

        fec += [
            fec_row(j_ve, "VENTES", ecriture, date,
                    cpt_cli, "Clients", ax, al,
                    piece_ref, date, lib,
                    ttc, 0.0,
                    "", "", "", "", ""),
            fec_row(j_ve, "VENTES", ecriture, date,
                    cpt_vente, "Ventes", "", "",
                    piece_ref, date, lib,
                    0.0, abs(ht),
                    "", "", "", "", "")
        ]

        if abs(tva) > 0.000001:
            fec.append(
                fec_row(j_ve, "VENTES", ecriture, date,
                        cpt_tva, "TVA collectée", "", "",
                        piece_ref, date, lib,
                        0.0, abs(tva),
                        "", "", "", "", "")
            )

        ecriture += 1

# ======================================================
# ENCAISSEMENTS
# ======================================================
if pay_file:
    df = read_any(pay_file)
    df.columns = [norm_key(c) for c in df.columns]

    # mapping encaissements
    col_date = "date saisie" if "date saisie" in df.columns else ("date" if "date" in df.columns else None)
    col_amt = "montant" if "montant" in df.columns else None
    col_mode = "rgt" if "rgt" in df.columns else ("rgt." if "rgt." in df.columns else None)
    col_fact = "fact" if "fact" in df.columns else ("fact." if "fact." in df.columns else None)
    col_cid = "id client" if "id client" in df.columns else None
    col_cname = "client" if "client" in df.columns else None

    for _, r in df.iterrows():
        amt = parse_fr_number(r.get(col_amt)) if col_amt else 0.0
        if abs(amt) < 0.000001:
            continue

        mode = sstr(r.get(col_mode)) if col_mode else ""
        mlow = mode.lower()

        if "cb" in mlow or "carte" in mlow:
            cpt = cpt_cb
            cpt_lib = "Caisse CB"
        elif "chq" in mlow or "cheq" in mlow or "chèque" in mlow:
            cpt = cpt_chq
            cpt_lib = "Caisse chèques"
        elif "esp" in mlow or "espe" in mlow or "cash" in mlow:
            cpt = cpt_esp
            cpt_lib = "Caisse espèces"
        else:
            cpt = cpt_aut
            cpt_lib = "Caisse autres"

        date = parse_date_any(r.get(col_date)) if col_date else ""
        fact = clean_invoice_no(r.get(col_fact)) if col_fact else ""

        cid = r.get(col_cid) if col_cid else ""
        cname = r.get(col_cname) if col_cname else ""
        ax, al = aux(cid, cname)

        lib = join_nonempty(["Encaissement", fact, cname, mode])

        piece_ref = f"CS-{ecriture}"

        # Sens (si amt négatif => annulation)
        if amt >= 0:
            # 53 D / 411 C
            fec += [
                fec_row(j_cs, "CAISSE", ecriture, date,
                        cpt, cpt_lib, "", "",
                        piece_ref, date, lib,
                        abs(amt), 0.0,
                        "", "", "", "", ""),
                fec_row(j_cs, "CAISSE", ecriture, date,
                        cpt_cli, "Clients", ax, al,
                        piece_ref, date, lib,
                        0.0, abs(amt),
                        "", "", "", "", "")
            ]
        else:
            # 53 C / 411 D
            fec += [
                fec_row(j_cs, "CAISSE", ecriture, date,
                        cpt, cpt_lib, "", "",
                        piece_ref, date, lib,
                        0.0, abs(amt),
                        "", "", "", "", ""),
                fec_row(j_cs, "CAISSE", ecriture, date,
                        cpt_cli, "Clients", ax, al,
                        piece_ref, date, lib,
                        abs(amt), 0.0,
                        "", "", "", "", "")
            ]

        ecriture += 1

# ======================================================
# EXPORT + AFFICHAGE COMPLET (avec contrôle)
# ======================================================
df_fec = pd.DataFrame(fec, columns=FEC_COLUMNS)

# force lettrage vide (au cas où)
if not df_fec.empty:
    df_fec["EcritureLet"] = ""
    df_fec["DateLet"] = ""

st.subheader("FEC généré")
st.write(f"**Lignes totales : {len(df_fec)}** | **Écritures : {df_fec['EcritureNum'].nunique() if not df_fec.empty else 0}**")

if df_fec.empty:
    st.warning("Aucune ligne FEC générée (vérifie les colonnes montant/date).")
    st.stop()

# Affichage paramétrable (évite la fausse impression de “pas complet”)
max_rows = len(df_fec)
to_show = st.slider("Nombre de lignes à afficher", min_value=50, max_value=max_rows, value=min(500, max_rows), step=50)
st.dataframe(df_fec.head(to_show), use_container_width=True)

# Downloads
st.download_button(
    "📥 Télécharger le FEC (txt)",
    data=to_fec(df_fec, sep),
    file_name="FEC_DELPHY.txt",
    mime="text/plain"
)

st.download_button(
    "📥 Télécharger le FEC (Excel de contrôle)",
    data=to_excel_bytes(df_fec),
    file_name="FEC_DELPHY.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st

# -----------------------------
# Helpers
# -----------------------------
FEC_COLUMNS = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib",
    "Debit", "Credit",
    "EcritureLet", "DateLet",
    "ValidDate", "Montantdevise", "Idevise"
]

def normalize_colname(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip())

def parse_fr_number(x):
    """Robuste: gère 1 234,56 / 1234.56 / '', NaN."""
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00A0", " ").replace(" ", "")
    # si virgule décimale
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    # si format 1.234,56
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def parse_date_any(x):
    """Retourne YYYYMMDD. Accepte dd/mm/yyyy, datetime, '30/11/2025 - 16:25:31'."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%Y%m%d")
    s = str(x).strip()
    if s == "":
        return ""
    # split si "dd/mm/yyyy - hh:mm:ss"
    if " - " in s:
        s = s.split(" - ")[0].strip()
    # tentative dd/mm/yyyy
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%Y%m%d")
        except:
            pass
    # tentative pandas
    try:
        d = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return ""
        return d.strftime("%Y%m%d")
    except:
        return ""

def clean_invoice_no(x):
    """Nettoie un numéro de facture pour lettrage: garde digits/lettres."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    return re.sub(r"[^\w\-\/]", "", s)

def make_piece_ref(*parts, fallback=""):
    p = [str(x).strip() for x in parts if str(x).strip() not in ("", "nan", "None")]
    if p:
        return "-".join(p)[:50]
    return fallback[:50]

def fec_row(
    JournalCode, JournalLib, EcritureNum, EcritureDate,
    CompteNum, CompteLib,
    CompAuxNum, CompAuxLib,
    PieceRef, PieceDate, EcritureLib,
    Debit, Credit,
    EcritureLet="", DateLet="",
    ValidDate="", Montantdevise="", Idevise=""
):
    return {
        "JournalCode": JournalCode,
        "JournalLib": JournalLib,
        "EcritureNum": str(EcritureNum),
        "EcritureDate": EcritureDate,
        "CompteNum": str(CompteNum),
        "CompteLib": str(CompteLib)[:200],
        "CompAuxNum": str(CompAuxNum)[:40] if CompAuxNum else "",
        "CompAuxLib": str(CompAuxLib)[:200] if CompAuxLib else "",
        "PieceRef": str(PieceRef)[:50],
        "PieceDate": PieceDate,
        "EcritureLib": str(EcritureLib)[:200],
        "Debit": f"{Debit:.2f}" if Debit else "0.00",
        "Credit": f"{Credit:.2f}" if Credit else "0.00",
        "EcritureLet": str(EcritureLet)[:20] if EcritureLet else "",
        "DateLet": DateLet if DateLet else "",
        "ValidDate": ValidDate if ValidDate else "",
        "Montantdevise": Montantdevise if Montantdevise else "",
        "Idevise": Idevise if Idevise else "",
    }

def to_fec_text(df: pd.DataFrame, sep="|", encoding="utf-8"):
    df2 = df.copy()
    # assure ordre colonnes
    df2 = df2[FEC_COLUMNS]
    out = io.StringIO()
    df2.to_csv(out, sep=sep, index=False, line_terminator="\n", encoding=encoding)
    return out.getvalue().encode(encoding)

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Export FEC Opticien (Factures + Encaissements)", layout="wide")
st.title("Export écritures FEC — Opticien (Factures + Encaissements)")

with st.sidebar:
    st.header("Paramètres Comptables")

    st.subheader("Comptes")
    compte_vente = st.text_input("Compte ventes (Crédit HT)", value="7071")
    compte_tva = st.text_input("Compte TVA collectée 20% (Crédit TVA)", value="44571")
    compte_client_global = st.text_input("Compte clients 411 (Débit ventes / Crédit règlements)", value="FCLIENTS")
    compte_banque = st.text_input("Compte banque (Débit règlements)", value="512000")

    st.subheader("Journaux")
    jnl_ve = st.text_input("Journal Ventes (code)", value="VE")
    jnl_ve_lib = st.text_input("Journal Ventes (libellé)", value="VENTES")
    jnl_bq = st.text_input("Journal Banque (code)", value="BQ")
    jnl_bq_lib = st.text_input("Journal Banque (libellé)", value="BANQUE")

    st.subheader("Options lettrage")
    use_aux_per_client = st.checkbox("Utiliser CompAuxNum = ID client (sinon vide)", value=True)
    use_client_name_in_auxlib = st.checkbox("Mettre CompAuxLib = Nom client", value=True)

    st.subheader("Export")
    sep = st.selectbox("Séparateur FEC", options=["|", ";", "\t"], index=0)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1) Fichier FACTURES (CSV)")
    fact_file = st.file_uploader("Upload CSV factures", type=["csv"], key="fact")

with col2:
    st.subheader("2) Fichier ENCAISSEMENTS (Excel ou CSV)")
    pay_file = st.file_uploader("Upload encaissements", type=["xlsx", "xls", "csv"], key="pay")

if not fact_file or not pay_file:
    st.info("Charge les deux fichiers pour générer le FEC.")
    st.stop()

# -----------------------------
# Read invoices
# -----------------------------
try:
    # Tentative csv avec séparateurs usuels
    fact_bytes = fact_file.getvalue()
    # Essai auto-detect sep
    fact_text = fact_bytes.decode("utf-8", errors="replace")
    # heuristique
    sep_candidates = [";", ",", "\t", "|"]
    best_sep = ";"
    best_cols = 0
    for s in sep_candidates:
        try:
            tmp = pd.read_csv(io.StringIO(fact_text), sep=s)
            if tmp.shape[1] > best_cols:
                best_cols = tmp.shape[1]
                best_sep = s
        except:
            pass
    df_fact = pd.read_csv(io.StringIO(fact_text), sep=best_sep, dtype=str)
except Exception as e:
    st.error(f"Erreur lecture CSV factures: {e}")
    st.stop()

df_fact.columns = [normalize_colname(c) for c in df_fact.columns]

# Expected columns in your sample:
# "Date", "ID client", "Client", "N° Fact/Avoir", "montant HT", "TVA", maybe "Vente", "Article", "Vend.", "T"
for c in ["Date", "N° Fact/Avoir", "montant HT", "TVA"]:
    if c not in df_fact.columns:
        st.error(f"Colonne manquante dans FACTURES: '{c}'")
        st.write("Colonnes trouvées:", list(df_fact.columns))
        st.stop()

# Build numeric fields
df_fact["montant_HT_num"] = df_fact["montant HT"].apply(parse_fr_number)
df_fact["TVA_num"] = df_fact["TVA"].apply(parse_fr_number)
df_fact["TTC_num"] = df_fact["montant_HT_num"] + df_fact["TVA_num"]

df_fact["date_fec"] = df_fact["Date"].apply(parse_date_any)
df_fact["fact_no"] = df_fact["N° Fact/Avoir"].apply(clean_invoice_no)

# Sign handling (avoirs)
# If column "Vente" contains "Avoir" or HT/TVA negative -> treat as credit note.
if "Vente" in df_fact.columns:
    is_avoir = df_fact["Vente"].astype(str).str.contains("avoir", case=False, na=False)
else:
    is_avoir = pd.Series([False]*len(df_fact))

is_negative = (df_fact["TTC_num"] < 0) | (df_fact["montant_HT_num"] < 0) | (df_fact["TVA_num"] < 0)
df_fact["is_avoir"] = is_avoir | is_negative

# Absolute values for posting, sign decides debit/credit direction
df_fact["HT_abs"] = df_fact["montant_HT_num"].abs()
df_fact["TVA_abs"] = df_fact["TVA_num"].abs()
df_fact["TTC_abs"] = df_fact["TTC_num"].abs()

# -----------------------------
# Read payments
# -----------------------------
try:
    if pay_file.name.lower().endswith(".csv"):
        pay_text = pay_file.getvalue().decode("utf-8", errors="replace")
        # auto sep
        best_sep = ";"
        best_cols = 0
        for s in [";", ",", "\t", "|"]:
            try:
                tmp = pd.read_csv(io.StringIO(pay_text), sep=s)
                if tmp.shape[1] > best_cols:
                    best_cols = tmp.shape[1]
                    best_sep = s
            except:
                pass
        df_pay = pd.read_csv(io.StringIO(pay_text), sep=best_sep, dtype=str)
    else:
        df_pay = pd.read_excel(pay_file, dtype=str)
except Exception as e:
    st.error(f"Erreur lecture encaissements: {e}")
    st.stop()

df_pay.columns = [normalize_colname(c) for c in df_pay.columns]

for c in ["Date Saisie", "Montant"]:
    if c not in df_pay.columns:
        st.error(f"Colonne manquante dans ENCAISSEMENTS: '{c}'")
        st.write("Colonnes trouvées:", list(df_pay.columns))
        st.stop()

df_pay["date_fec"] = df_pay["Date Saisie"].apply(parse_date_any)
df_pay["montant_num"] = df_pay["Montant"].apply(parse_fr_number)
df_pay["montant_abs"] = df_pay["montant_num"].abs()

# facture number can be in "Fact." column
if "Fact." in df_pay.columns:
    df_pay["fact_no"] = df_pay["Fact."].apply(clean_invoice_no)
else:
    df_pay["fact_no"] = ""

# filter? -> on conserve tout (y compris annulations) car ça aide à tracer
# mais on ignore les lignes à 0
df_pay = df_pay[df_pay["montant_abs"] > 0.000001].copy()

# -----------------------------
# Preview
# -----------------------------
tab1, tab2, tab3 = st.tabs(["Aperçu factures", "Aperçu encaissements", "FEC généré"])

with tab1:
    st.write(f"Factures lues: {len(df_fact)} (séparateur détecté: '{best_sep}')")
    st.dataframe(df_fact.head(50), use_container_width=True)

with tab2:
    st.write(f"Encaissements lus: {len(df_pay)}")
    st.dataframe(df_pay.head(50), use_container_width=True)

# -----------------------------
# Build FEC entries
# -----------------------------
fec_rows = []
ecriture_num = 1

def get_aux(id_client, client_name):
    if not use_aux_per_client:
        return ("", "")
    auxnum = str(id_client).strip() if id_client not in (None, "", "nan") else ""
    auxlib = str(client_name).strip() if use_client_name_in_auxlib and client_name not in (None, "", "nan") else ""
    return auxnum[:40], auxlib[:200]

# FACTURES -> 3 lignes par facture (411 / 707 / 44571)
for _, r in df_fact.iterrows():
    date_ecr = r.get("date_fec", "") or ""
    piece_date = date_ecr
    fact_no = r.get("fact_no", "") or ""
    vente_type = str(r.get("T", "")).strip()
    article = str(r.get("Article", "")).strip()
    vendeur = str(r.get("Vend.", "")).strip()
    client_id = r.get("ID client", "")
    client_name = r.get("Client", "")

    auxnum, auxlib = get_aux(client_id, client_name)

    piece_ref = make_piece_ref("FAC", fact_no, fallback=f"FAC-{ecriture_num}")
    let = fact_no  # LETTRAGE = numéro facture

    # Libellé riche
    lib = " | ".join([x for x in [
        f"Facture {fact_no}" if fact_no else "Facture",
        client_name if client_name else "",
        f"ID:{client_id}" if str(client_id).strip() not in ("", "nan") else "",
        f"Vendeur:{vendeur}" if vendeur else "",
        f"T:{vente_type}" if vente_type else "",
        article if article else ""
    ] if x])

    ht = float(r.get("HT_abs", 0.0))
    tva = float(r.get("TVA_abs", 0.0))
    ttc = float(r.get("TTC_abs", 0.0))
    is_avoir_row = bool(r.get("is_avoir", False))

    # Sens normal facture: 411 D TTC / 707 C HT / 44571 C TVA
    # Sens avoir: inverse
    if not is_avoir_row:
        # 411 débit
        fec_rows.append(fec_row(
            jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
            compte_client_global, "Clients",
            auxnum, auxlib,
            piece_ref, piece_date, lib,
            Debit=ttc, Credit=0.0,
            EcritureLet=let
        ))
        # 707 crédit HT
        fec_rows.append(fec_row(
            jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
            compte_vente, "Ventes",
            "", "",
            piece_ref, piece_date, lib,
            Debit=0.0, Credit=ht,
            EcritureLet=let
        ))
        # TVA crédit
        if tva > 0:
            fec_rows.append(fec_row(
                jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                compte_tva, "TVA collectée",
                "", "",
                piece_ref, piece_date, lib,
                Debit=0.0, Credit=tva,
                EcritureLet=let
            ))
    else:
        # Avoir: 411 crédit TTC / 707 débit HT / TVA débit
        fec_rows.append(fec_row(
            jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
            compte_client_global, "Clients",
            auxnum, auxlib,
            piece_ref, piece_date, f"AVOIR - {lib}",
            Debit=0.0, Credit=ttc,
            EcritureLet=let
        ))
        fec_rows.append(fec_row(
            jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
            compte_vente, "Ventes",
            "", "",
            piece_ref, piece_date, f"AVOIR - {lib}",
            Debit=ht, Credit=0.0,
            EcritureLet=let
        ))
        if tva > 0:
            fec_rows.append(fec_row(
                jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                compte_tva, "TVA collectée",
                "", "",
                piece_ref, piece_date, f"AVOIR - {lib}",
                Debit=tva, Credit=0.0,
                EcritureLet=let
            ))

    ecriture_num += 1

# ENCAISSEMENTS -> 2 lignes par mouvement (512 / 411)
for _, r in df_pay.iterrows():
    date_ecr = r.get("date_fec", "") or ""
    piece_date = date_ecr

    fact_no = r.get("fact_no", "") or ""
    client_id = r.get("ID client", r.get("ID client ", "")) if ("ID client" in df_pay.columns or "ID client " in df_pay.columns) else r.get("ID client", "")
    client_name = r.get("Client", "")
    mode = str(r.get("Rgt.", "")).strip()
    sous_mode = str(r.get("Sous Rgt.", "")).strip()
    typ = str(r.get("Type", "")).strip()
    bord = str(r.get("N° Bord.", "")).strip()
    echeance = str(r.get("Echéance", "")).strip()

    auxnum, auxlib = get_aux(client_id, client_name)

    amt = float(r.get("montant_abs", 0.0))
    sign = 1 if float(r.get("montant_num", 0.0)) >= 0 else -1

    # piece_ref: bordereau si dispo sinon facture sinon fallback
    piece_ref = make_piece_ref("BORD", bord, "FAC", fact_no, fallback=f"RG-{date_ecr}-{ecriture_num}")
    let = fact_no  # Lettrage sur facture si dispo

    lib = " | ".join([x for x in [
        "Encaissement" if sign > 0 else "Annulation règlement",
        f"Fact {fact_no}" if fact_no else "",
        client_name if client_name else "",
        f"ID:{client_id}" if str(client_id).strip() not in ("", "nan") else "",
        f"Mode:{mode}" if mode else "",
        f"Sous:{sous_mode}" if sous_mode else "",
        f"Type:{typ}" if typ else "",
        f"Bord:{bord}" if bord else "",
        f"Echéance:{echeance}" if echeance else "",
    ] if x])

    # Sens: encaissement + : 512 D / 411 C
    # Sens: encaissement - (annulation) : 512 C / 411 D
    if sign > 0:
        fec_rows.append(fec_row(
            jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
            compte_banque, "Banque",
            "", "",
            piece_ref, piece_date, lib,
            Debit=amt, Credit=0.0,
            EcritureLet=let
        ))
        fec_rows.append(fec_row(
            jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
            compte_client_global, "Clients",
            auxnum, auxlib,
            piece_ref, piece_date, lib,
            Debit=0.0, Credit=amt,
            EcritureLet=let
        ))
    else:
        fec_rows.append(fec_row(
            jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
            compte_banque, "Banque",
            "", "",
            piece_ref, piece_date, lib,
            Debit=0.0, Credit=amt,
            EcritureLet=let
        ))
        fec_rows.append(fec_row(
            jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
            compte_client_global, "Clients",
            auxnum, auxlib,
            piece_ref, piece_date, lib,
            Debit=amt, Credit=0.0,
            EcritureLet=let
        ))

    ecriture_num += 1

df_fec = pd.DataFrame(fec_rows, columns=FEC_COLUMNS)

# -----------------------------
# Post checks
# -----------------------------
with tab3:
    st.subheader("FEC généré")
    st.write(f"Lignes FEC: {len(df_fec)} | Écritures: {df_fec['EcritureNum'].nunique()}")

    # Contrôle équilibre par écriture
    ctrl = df_fec.copy()
    ctrl["Debit_f"] = ctrl["Debit"].apply(parse_fr_number)
    ctrl["Credit_f"] = ctrl["Credit"].apply(parse_fr_number)
    bal = ctrl.groupby(["JournalCode", "EcritureNum"], as_index=False).agg(
        Debit=("Debit_f", "sum"), Credit=("Credit_f", "sum")
    )
    bal["Diff"] = (bal["Debit"] - bal["Credit"]).round(2)

    bad = bal[bal["Diff"].abs() > 0.01]
    if len(bad) > 0:
        st.warning(f"⚠️ {len(bad)} écritures non équilibrées (écart > 0,01).")
        st.dataframe(bad.head(50), use_container_width=True)
    else:
        st.success("✅ Toutes les écritures sont équilibrées.")

    st.dataframe(df_fec.head(200), use_container_width=True)

# -----------------------------
# Download
# -----------------------------
fec_bytes = to_fec_text(df_fec, sep=sep, encoding="utf-8")
st.download_button(
    "📥 Télécharger le FEC (txt/csv)",
    data=fec_bytes,
    file_name=f"FEC_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
    mime="text/plain"
)

st.download_button(
    "📥 Télécharger le FEC en Excel (pour contrôle)",
    data=io.BytesIO(df_fec.to_excel(index=False, engine="openpyxl")),
    file_name=f"FEC_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

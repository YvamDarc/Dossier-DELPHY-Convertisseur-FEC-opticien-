import io
import re
import unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st

# -----------------------------
# Constantes FEC
# -----------------------------
FEC_COLUMNS = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib",
    "Debit", "Credit",
    "EcritureLet", "DateLet",
    "ValidDate", "Montantdevise", "Idevise"
]

# -----------------------------
# Utils
# -----------------------------
def strip_accents(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    return "".join([c for c in s if not unicodedata.combining(c)])

def norm_key(s: str) -> str:
    """Normalise une colonne: enlève accents, espaces, ponctuation, casse."""
    s = strip_accents(s)
    s = s.replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.lower()
    # remplace caractères bizarres éventuels
    s = s.replace("�", "")
    # garde lettres/chiffres/espaces/./°
    s = re.sub(r"[^a-z0-9 \.\-\/°]", "", s)
    s = s.strip()
    return s

def parse_fr_number(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00A0", " ").replace(" ", "")
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def parse_date_any(x):
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%Y%m%d")
    s = str(x).strip()
    if s == "":
        return ""
    if " - " in s:
        s = s.split(" - ")[0].strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%Y%m%d")
        except:
            pass
    d = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(d):
        return ""
    return d.strftime("%Y%m%d")

def clean_invoice_no(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    s = s.replace("�", "")
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
    df2 = df2[FEC_COLUMNS]
    out = io.StringIO()
    df2.to_csv(out, sep=sep, index=False, line_terminator="\n")
    return out.getvalue().encode(encoding)

def read_any_table(uploaded_file):
    """Lit csv/xlsx, auto-sep, retourne df str."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(data), dtype=str)

    # csv
    text = data.decode("utf-8", errors="replace")
    best_sep, best_cols = ";", 0
    for s in [";", ",", "\t", "|"]:
        try:
            tmp = pd.read_csv(io.StringIO(text), sep=s, dtype=str)
            if tmp.shape[1] > best_cols:
                best_cols, best_sep = tmp.shape[1], s
        except:
            pass
    df = pd.read_csv(io.StringIO(text), sep=best_sep, dtype=str)
    return df

def auto_map_columns(df: pd.DataFrame, kind: str):
    """
    kind = 'fact' ou 'pay'
    Retourne df + dict mapping log.
    """
    original_cols = list(df.columns)
    norm_cols = {c: norm_key(c) for c in original_cols}

    # synonymes
    if kind == "fact":
        synonyms = {
            "date": ["date", "date ", "date facture", "datefac", "date fact"],
            "invoice": ["n fact/avoir", "n fact", "n facture", "n facture/avoir", "n° fact/avoir", "n factavoir",
                        "n fact avoir", "n fact/avo", "n fact/av", "n factavoir ", "n fact/avoir ", "n factavoir",
                        "n fact/avoir", "n fact/avoir", "n fact/avoir", "n fact/avoir", "n fact/avoir",
                        "n fact/avoir", "n fact/avoir", "n fact/avoir", "n fact/avoir",
                        "n fact/avoir", "n fact/avoir", "n fact/avoir", "n fact/avoir",
                        "n fact/avoir", "n fact/avoir",
                        "n fact/avoir", "n fact/avoir", "n fact/avoir",
                        "n fact/avoir",
                        "n fact/avoir",
                        "n fact/avoir",
                        "n fact/avoir",
                        "n fact/avoir",
                        "n fact/avoir",
                        # version cassée encodage
                        "n fact/avoir", "n factavoir", "n factavoir", "n factavoir",
                        "n factavoir", "n factavoir",
                        "n fact/avoir", "n factavoir",
                        "n fact/avoir", "n fact/avoir",
                        "n factavoir",
                        "n factavoir",
                        "n fact/avoir",
                        "n factavoir",
                        "n fact/avoir",
                        "n factavoir",
                        "n factavoir",
                        "n fact/avoir",
                        "n factavoir",
                        "n factavoir",
                        "n factavoir",
                        "n factavoir",
                        # ton cas: "N� Fact/Avoir" devient souvent "n fact/avoir"
                        "n fact/avoir"
                       ],
            "client_id": ["id client", "client id", "idclient", "code client"],
            "client_name": ["client", "nom client", "client ", "client name"],
            "ht": ["montant ht", "montantht", "ht", "total ht"],
            "tva": ["tva", "montant tva", "tva "],
            "article": ["article", "libelle", "designation"],
            "vendeur": ["vend.", "vendeur", "vendez.", "vendez", "vend "],
            "vente_lib": ["vente", "type vente", "vente "],
            "t_col": ["t", "t ", "type", "type "],
        }
    else:
        synonyms = {
            "date": ["date saisie", "date", "date ", "date paiement", "date reglement"],
            "amount": ["montant", "amount", "montant "],
            "invoice": ["fact.", "fact", "facture", "n fact", "n facture"],
            "client_id": ["id client", "client id", "idclient"],
            "client_name": ["client", "nom client"],
            "mode": ["rgt.", "rgt", "mode", "mode reglement"],
            "sous_mode": ["sous rgt.", "sous rgt", "sous-mode", "sous mode"],
            "type": ["type"],
            "bord": ["n bord.", "n bord", "bordereau", "n bordereau"],
            "echeance": ["echeance", "echeance "],
        }

    # inverse lookup: norm->original
    norm_to_original = {}
    for orig, nk in norm_cols.items():
        norm_to_original.setdefault(nk, []).append(orig)

    chosen = {}
    for target, syns in synonyms.items():
        found = None
        for syn in syns:
            syn_n = norm_key(syn)
            # match exact
            if syn_n in norm_to_original:
                found = norm_to_original[syn_n][0]
                break
        # fallback: contient
        if not found:
            for nk, origs in norm_to_original.items():
                if any(norm_key(s) in nk for s in syns):
                    found = origs[0]
                    break
        chosen[target] = found

    # applique renommage minimal
    df2 = df.copy()
    rename_map = {}

    if kind == "fact":
        if chosen["date"]: rename_map[chosen["date"]] = "Date"
        if chosen["invoice"]: rename_map[chosen["invoice"]] = "InvoiceNo"
        if chosen["client_id"]: rename_map[chosen["client_id"]] = "ClientID"
        if chosen["client_name"]: rename_map[chosen["client_name"]] = "ClientName"
        if chosen["ht"]: rename_map[chosen["ht"]] = "HT"
        if chosen["tva"]: rename_map[chosen["tva"]] = "TVA"
        if chosen["article"]: rename_map[chosen["article"]] = "Article"
        if chosen["vendeur"]: rename_map[chosen["vendeur"]] = "Vendeur"
        if chosen["vente_lib"]: rename_map[chosen["vente_lib"]] = "VenteLib"
        if chosen["t_col"]: rename_map[chosen["t_col"]] = "Tcol"
    else:
        if chosen["date"]: rename_map[chosen["date"]] = "DateSaisie"
        if chosen["amount"]: rename_map[chosen["amount"]] = "Montant"
        if chosen["invoice"]: rename_map[chosen["invoice"]] = "InvoiceNo"
        if chosen["client_id"]: rename_map[chosen["client_id"]] = "ClientID"
        if chosen["client_name"]: rename_map[chosen["client_name"]] = "ClientName"
        if chosen["mode"]: rename_map[chosen["mode"]] = "Mode"
        if chosen["sous_mode"]: rename_map[chosen["sous_mode"]] = "SousMode"
        if chosen["type"]: rename_map[chosen["type"]] = "TypeLib"
        if chosen["bord"]: rename_map[chosen["bord"]] = "Bord"
        if chosen["echeance"]: rename_map[chosen["echeance"]] = "Echeance"

    df2 = df2.rename(columns=rename_map)

    log = {
        "original_cols": original_cols,
        "normalized_cols": {k: norm_cols[k] for k in original_cols},
        "chosen": chosen,
        "rename_map": rename_map
    }
    return df2, log

# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="DELPHY - Export FEC Opticien", layout="wide")
st.title("Export FEC — Factures / Encaissements (fichiers indépendants)")

with st.sidebar:
    st.header("Paramètres")

    st.subheader("Comptes")
    compte_vente = st.text_input("707 ventes (Crédit HT)", value="7071")
    compte_tva = st.text_input("44571 TVA collectée 20% (Crédit TVA)", value="44571")
    compte_client_global = st.text_input("411 clients global", value="FCLIENTS")
    compte_banque = st.text_input("512 banque", value="512000")

    st.subheader("Journaux")
    jnl_ve = st.text_input("Journal ventes", value="VE")
    jnl_ve_lib = st.text_input("Libellé ventes", value="VENTES")
    jnl_bq = st.text_input("Journal banque", value="BQ")
    jnl_bq_lib = st.text_input("Libellé banque", value="BANQUE")

    st.subheader("Lettrage")
    use_aux_per_client = st.checkbox("CompAuxNum = ID client", value=True)
    use_client_name_in_auxlib = st.checkbox("CompAuxLib = Nom client", value=True)

    st.subheader("Export")
    sep = st.selectbox("Séparateur", options=["|", ";", "\t"], index=0)

st.markdown("### 1) Charge tes fichiers (tu peux en mettre un seul)")
col1, col2 = st.columns(2)
with col1:
    fact_file = st.file_uploader("Factures (CSV/XLSX) — optionnel", type=["csv","xlsx","xls"], key="fact")
with col2:
    pay_file = st.file_uploader("Encaissements (CSV/XLSX) — optionnel", type=["csv","xlsx","xls"], key="pay")

if not fact_file and not pay_file:
    st.info("Charge au moins un fichier (factures ou encaissements).")
    st.stop()

def get_aux(client_id, client_name):
    if not use_aux_per_client:
        return ("", "")
    auxnum = str(client_id).strip() if client_id not in (None, "", "nan") else ""
    auxlib = str(client_name).strip() if use_client_name_in_auxlib and client_name not in (None, "", "nan") else ""
    return auxnum[:40], auxlib[:200]

fec_rows = []
ecriture_num = 1

tabs = st.tabs(["Logs mapping", "Aperçu", "FEC"])

mapping_logs = {}

# -----------------------------
# FACTURES
# -----------------------------
df_fact = None
if fact_file:
    try:
        df_raw = read_any_table(fact_file)
        df_raw.columns = [str(c) for c in df_raw.columns]
        df_fact, log = auto_map_columns(df_raw, "fact")
        mapping_logs["factures"] = log

        # champs calculés (si absents -> 0)
        df_fact["HT_num"] = df_fact["HT"].apply(parse_fr_number) if "HT" in df_fact.columns else 0.0
        df_fact["TVA_num"] = df_fact["TVA"].apply(parse_fr_number) if "TVA" in df_fact.columns else 0.0
        df_fact["TTC_num"] = df_fact["HT_num"] + df_fact["TVA_num"]

        df_fact["date_fec"] = df_fact["Date"].apply(parse_date_any) if "Date" in df_fact.columns else ""
        df_fact["fact_no"] = df_fact["InvoiceNo"].apply(clean_invoice_no) if "InvoiceNo" in df_fact.columns else ""

        # détection avoir (si TTC/HT/TVA négatifs ou si VenteLib contient "avoir")
        is_negative = (df_fact["TTC_num"] < 0) | (df_fact["HT_num"] < 0) | (df_fact["TVA_num"] < 0)
        if "VenteLib" in df_fact.columns:
            is_avoir_txt = df_fact["VenteLib"].astype(str).str.contains("avoir", case=False, na=False)
        else:
            is_avoir_txt = False
        df_fact["is_avoir"] = is_negative | is_avoir_txt

        df_fact["HT_abs"] = df_fact["HT_num"].abs()
        df_fact["TVA_abs"] = df_fact["TVA_num"].abs()
        df_fact["TTC_abs"] = df_fact["TTC_num"].abs()

        # Génération écritures ventes
        for _, r in df_fact.iterrows():
            date_ecr = r.get("date_fec", "") or ""
            piece_date = date_ecr
            fact_no = r.get("fact_no", "") or ""

            client_id = r.get("ClientID", "")
            client_name = r.get("ClientName", "")

            auxnum, auxlib = get_aux(client_id, client_name)

            vendeur = str(r.get("Vendeur", "")).strip()
            tcol = str(r.get("Tcol", "")).strip()
            article = str(r.get("Article", "")).strip()

            piece_ref = make_piece_ref("FAC", fact_no, fallback=f"FAC-{ecriture_num}")
            let = fact_no  # lettrage si dispo

            lib = " | ".join([x for x in [
                f"Facture {fact_no}" if fact_no else "Facture",
                client_name if client_name else "",
                f"ID:{client_id}" if str(client_id).strip() not in ("", "nan") else "",
                f"Vendeur:{vendeur}" if vendeur else "",
                f"T:{tcol}" if tcol else "",
                article if article else "",
            ] if x])

            ht = float(r.get("HT_abs", 0.0))
            tva = float(r.get("TVA_abs", 0.0))
            ttc = float(r.get("TTC_abs", 0.0))
            is_avoir_row = bool(r.get("is_avoir", False))

            if not is_avoir_row:
                fec_rows.append(fec_row(
                    jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                    compte_client_global, "Clients",
                    auxnum, auxlib,
                    piece_ref, piece_date, lib,
                    Debit=ttc, Credit=0.0,
                    EcritureLet=let
                ))
                fec_rows.append(fec_row(
                    jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                    compte_vente, "Ventes",
                    "", "",
                    piece_ref, piece_date, lib,
                    Debit=0.0, Credit=ht,
                    EcritureLet=let
                ))
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

    except Exception as e:
        st.error(f"Erreur traitement FACTURES: {e}")

# -----------------------------
# ENCAISSEMENTS
# -----------------------------
df_pay = None
if pay_file:
    try:
        df_raw = read_any_table(pay_file)
        df_raw.columns = [str(c) for c in df_raw.columns]
        df_pay, log = auto_map_columns(df_raw, "pay")
        mapping_logs["encaissements"] = log

        df_pay["date_fec"] = df_pay["DateSaisie"].apply(parse_date_any) if "DateSaisie" in df_pay.columns else ""
        df_pay["montant_num"] = df_pay["Montant"].apply(parse_fr_number) if "Montant" in df_pay.columns else 0.0
        df_pay["montant_abs"] = df_pay["montant_num"].abs()

        if "InvoiceNo" in df_pay.columns:
            df_pay["fact_no"] = df_pay["InvoiceNo"].apply(clean_invoice_no)
        else:
            df_pay["fact_no"] = ""

        # ignore lignes 0
        df_pay = df_pay[df_pay["montant_abs"] > 0.000001].copy()

        for _, r in df_pay.iterrows():
            date_ecr = r.get("date_fec", "") or ""
            piece_date = date_ecr

            fact_no = r.get("fact_no", "") or ""
            client_id = r.get("ClientID", "")
            client_name = r.get("ClientName", "")

            auxnum, auxlib = get_aux(client_id, client_name)

            mode = str(r.get("Mode", "")).strip()
            sous_mode = str(r.get("SousMode", "")).strip()
            typ = str(r.get("TypeLib", "")).strip()
            bord = str(r.get("Bord", "")).strip()
            echeance = str(r.get("Echeance", "")).strip()

            amt = float(r.get("montant_abs", 0.0))
            sign = 1 if float(r.get("montant_num", 0.0)) >= 0 else -1

            piece_ref = make_piece_ref("BORD", bord, "FAC", fact_no, fallback=f"RG-{date_ecr}-{ecriture_num}")
            let = fact_no  # lettrage facture si présent

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

    except Exception as e:
        st.error(f"Erreur traitement ENCAISSEMENTS: {e}")

# -----------------------------
# Affichages
# -----------------------------
with tabs[0]:
    st.subheader("Logs mapping (pour diagnostiquer les colonnes)")
    st.write(mapping_logs)

with tabs[1]:
    if df_fact is not None:
        st.subheader("Aperçu factures (après mapping)")
        st.dataframe(df_fact.head(50), use_container_width=True)
    if df_pay is not None:
        st.subheader("Aperçu encaissements (après mapping)")
        st.dataframe(df_pay.head(50), use_container_width=True)

with tabs[2]:
    df_fec = pd.DataFrame(fec_rows, columns=FEC_COLUMNS)
    st.subheader("FEC généré")
    st.write(f"Lignes: {len(df_fec)} | Écritures: {df_fec['EcritureNum'].nunique() if len(df_fec)>0 else 0}")
    st.dataframe(df_fec.head(200), use_container_width=True)

    if len(df_fec) > 0:
        # contrôle équilibre
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

        fec_bytes = to_fec_text(df_fec, sep=sep, encoding="utf-8")
        st.download_button(
            "📥 Télécharger le FEC (txt/csv)",
            data=fec_bytes,
            file_name=f"FEC_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )

# -----------------------------
# Fin
# -----------------------------

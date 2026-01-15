import io
import re
import unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st

FEC_COLUMNS = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib",
    "Debit", "Credit",
    "EcritureLet", "DateLet",
    "ValidDate", "Montantdevise", "Idevise"
]

# -----------------------------
# Utils robustes
# -----------------------------
def strip_accents(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    return "".join([c for c in s if not unicodedata.combining(c)])

def norm_key(s: str) -> str:
    s = strip_accents(s)
    s = s.replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.lower()
    s = s.replace("�", "")
    s = re.sub(r"[^a-z0-9 \.\-\/°]", "", s)
    return s.strip()

def sstr(x) -> str:
    """String safe: NaN -> '', float -> string, etc."""
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except:
        pass
    return str(x).strip()

def parse_fr_number(x):
    s = sstr(x)
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
    s = sstr(x)
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
    s = sstr(x).replace("�", "")
    if s == "":
        return ""
    return re.sub(r"[^\w\-\/]", "", s)

def make_piece_ref(*parts, fallback=""):
    p = [sstr(x) for x in parts if sstr(x) not in ("", "nan", "None")]
    if p:
        return "-".join(p)[:50]
    return sstr(fallback)[:50]

def join_nonempty(items, sep=" | "):
    """Join safe (stringify) + drop empties."""
    out = []
    for it in items:
        t = sstr(it)
        if t != "" and t.lower() != "nan":
            out.append(t)
    return sep.join(out)

def fec_row(
    JournalCode, JournalLib, EcritureNum, EcritureDate,
    CompteNum, CompteLib,
    CompAuxNum, CompAuxLib,
    PieceRef, PieceDate, EcritureLib,
    Debit, Credit,
    EcritureLet="", DateLet="",
    ValidDate="", Montantdevise="", Idevise=""
):
    # sécurise toutes les strings
    return {
        "JournalCode": sstr(JournalCode),
        "JournalLib": sstr(JournalLib)[:200],
        "EcritureNum": sstr(EcritureNum),
        "EcritureDate": sstr(EcritureDate),
        "CompteNum": sstr(CompteNum),
        "CompteLib": sstr(CompteLib)[:200],
        "CompAuxNum": sstr(CompAuxNum)[:40],
        "CompAuxLib": sstr(CompAuxLib)[:200],
        "PieceRef": sstr(PieceRef)[:50],
        "PieceDate": sstr(PieceDate),
        "EcritureLib": sstr(EcritureLib)[:200],
        "Debit": f"{float(Debit):.2f}" if float(Debit) != 0 else "0.00",
        "Credit": f"{float(Credit):.2f}" if float(Credit) != 0 else "0.00",
        "EcritureLet": sstr(EcritureLet)[:20],
        "DateLet": sstr(DateLet),
        "ValidDate": sstr(ValidDate),
        "Montantdevise": sstr(Montantdevise),
        "Idevise": sstr(Idevise),
    }

def to_fec_text(df: pd.DataFrame, sep="|", encoding="utf-8"):
    df2 = df.copy()
    df2 = df2[FEC_COLUMNS]

    # Important: forcer en str sur colonnes texte pour éviter surprises
    for c in FEC_COLUMNS:
        if c in df2.columns:
            if c in ("Debit", "Credit"):
                continue
            df2[c] = df2[c].astype(str).replace("nan", "").replace("None", "")

    out = io.StringIO()
    # pandas: utiliser lineterminator (plus stable)
    df2.to_csv(out, sep=sep, index=False, lineterminator="\n")
    return out.getvalue().encode(encoding)

def read_any_table(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(data), dtype=str)

    text = data.decode("utf-8", errors="replace")
    best_sep, best_cols = ";", 0
    for s in [";", ",", "\t", "|"]:
        try:
            tmp = pd.read_csv(io.StringIO(text), sep=s, dtype=str)
            if tmp.shape[1] > best_cols:
                best_cols, best_sep = tmp.shape[1], s
        except:
            pass
    return pd.read_csv(io.StringIO(text), sep=best_sep, dtype=str)

def auto_map_columns(df: pd.DataFrame, kind: str):
    original_cols = list(df.columns)
    norm_cols = {c: norm_key(c) for c in original_cols}

    if kind == "fact":
        synonyms = {
            "date": ["date"],
            "invoice": ["n fact/avoir", "n factavoir", "n facture", "facture", "n fact", "n fact/avoir"],
            "client_id": ["id client", "client id"],
            "client_name": ["client", "nom client"],
            "ht": ["montant ht", "ht"],
            "tva": ["tva"],
            "article": ["article", "designation", "libelle"],
            "vendeur": ["vend.", "vendeur", "vendez.", "vendez"],
            "vente_lib": ["vente"],
            "t_col": ["t", "type"],
        }
    else:
        synonyms = {
            "date": ["date saisie", "date"],
            "amount": ["montant"],
            "invoice": ["fact.", "fact", "facture", "n facture", "n fact"],
            "client_id": ["id client", "client id"],
            "client_name": ["client", "nom client"],
            "mode": ["rgt.", "rgt", "mode"],
            "sous_mode": ["sous rgt.", "sous rgt", "sous mode"],
            "type": ["type"],
            "bord": ["n bord.", "n bord", "bordereau"],
            "echeance": ["echeance"],
        }

    norm_to_original = {}
    for orig, nk in norm_cols.items():
        norm_to_original.setdefault(nk, []).append(orig)

    chosen = {}
    for target, syns in synonyms.items():
        found = None
        for syn in syns:
            syn_n = norm_key(syn)
            if syn_n in norm_to_original:
                found = norm_to_original[syn_n][0]
                break
        if not found:
            # contient
            for nk, origs in norm_to_original.items():
                for syn in syns:
                    if norm_key(syn) in nk:
                        found = origs[0]
                        break
                if found:
                    break
        chosen[target] = found

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
        "chosen": chosen,
        "rename_map": rename_map
    }
    return df2, log

# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="DELPHY - Export FEC Opticien", layout="wide")
st.title("DELPHY - Export FEC — Factures / Encaissements (fichiers indépendants)")

with st.sidebar:
    st.header("Paramètres")

    st.subheader("Comptes")
    compte_vente = st.text_input("707 ventes (Crédit HT)", value="7071")
    compte_tva = st.text_input("44571 TVA collectée 20%", value="44571")
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
    auxnum = sstr(client_id)
    auxlib = sstr(client_name) if use_client_name_in_auxlib else ""
    return auxnum[:40], auxlib[:200]

fec_rows = []
ecriture_num = 1
mapping_logs = {}

tabs = st.tabs(["Logs mapping", "Aperçu", "FEC"])

# -----------------------------
# FACTURES
# -----------------------------
df_fact = None
if fact_file:
    try:
        df_raw = read_any_table(fact_file)
        df_raw.columns = [sstr(c) for c in df_raw.columns]
        df_fact, log = auto_map_columns(df_raw, "fact")
        mapping_logs["factures"] = log

        df_fact["HT_num"] = df_fact["HT"].apply(parse_fr_number) if "HT" in df_fact.columns else 0.0
        df_fact["TVA_num"] = df_fact["TVA"].apply(parse_fr_number) if "TVA" in df_fact.columns else 0.0
        df_fact["TTC_num"] = df_fact["HT_num"] + df_fact["TVA_num"]

        df_fact["date_fec"] = df_fact["Date"].apply(parse_date_any) if "Date" in df_fact.columns else ""
        df_fact["fact_no"] = df_fact["InvoiceNo"].apply(clean_invoice_no) if "InvoiceNo" in df_fact.columns else ""

        is_negative = (df_fact["TTC_num"] < 0) | (df_fact["HT_num"] < 0) | (df_fact["TVA_num"] < 0)
        is_avoir_txt = df_fact["VenteLib"].astype(str).str.contains("avoir", case=False, na=False) if "VenteLib" in df_fact.columns else False
        df_fact["is_avoir"] = is_negative | is_avoir_txt

        df_fact["HT_abs"] = df_fact["HT_num"].abs()
        df_fact["TVA_abs"] = df_fact["TVA_num"].abs()
        df_fact["TTC_abs"] = df_fact["TTC_num"].abs()

        for _, r in df_fact.iterrows():
            date_ecr = sstr(r.get("date_fec", ""))
            piece_date = date_ecr
            fact_no = sstr(r.get("fact_no", ""))

            client_id = r.get("ClientID", "")
            client_name = r.get("ClientName", "")
            auxnum, auxlib = get_aux(client_id, client_name)

            vendeur = r.get("Vendeur", "")
            tcol = r.get("Tcol", "")
            article = r.get("Article", "")

            piece_ref = make_piece_ref("FAC", fact_no, fallback=f"FAC-{ecriture_num}")
            let = fact_no

            lib = join_nonempty([
                f"Facture {fact_no}" if fact_no else "Facture",
                client_name,
                f"ID:{client_id}" if sstr(client_id) else "",
                f"Vendeur:{vendeur}" if sstr(vendeur) else "",
                f"T:{tcol}" if sstr(tcol) else "",
                article
            ])

            ht = float(r.get("HT_abs", 0.0))
            tva = float(r.get("TVA_abs", 0.0))
            ttc = float(r.get("TTC_abs", 0.0))
            is_avoir_row = bool(r.get("is_avoir", False))

            if not is_avoir_row:
                fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                        compte_client_global, "Clients",
                                        auxnum, auxlib,
                                        piece_ref, piece_date, lib,
                                        Debit=ttc, Credit=0.0,
                                        EcritureLet=let))
                fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                        compte_vente, "Ventes",
                                        "", "",
                                        piece_ref, piece_date, lib,
                                        Debit=0.0, Credit=ht,
                                        EcritureLet=let))
                if tva > 0:
                    fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                            compte_tva, "TVA collectée",
                                            "", "",
                                            piece_ref, piece_date, lib,
                                            Debit=0.0, Credit=tva,
                                            EcritureLet=let))
            else:
                fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                        compte_client_global, "Clients",
                                        auxnum, auxlib,
                                        piece_ref, piece_date, f"AVOIR - {lib}",
                                        Debit=0.0, Credit=ttc,
                                        EcritureLet=let))
                fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                        compte_vente, "Ventes",
                                        "", "",
                                        piece_ref, piece_date, f"AVOIR - {lib}",
                                        Debit=ht, Credit=0.0,
                                        EcritureLet=let))
                if tva > 0:
                    fec_rows.append(fec_row(jnl_ve, jnl_ve_lib, ecriture_num, date_ecr,
                                            compte_tva, "TVA collectée",
                                            "", "",
                                            piece_ref, piece_date, f"AVOIR - {lib}",
                                            Debit=tva, Credit=0.0,
                                            EcritureLet=let))

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
        df_raw.columns = [sstr(c) for c in df_raw.columns]
        df_pay, log = auto_map_columns(df_raw, "pay")
        mapping_logs["encaissements"] = log

        df_pay["date_fec"] = df_pay["DateSaisie"].apply(parse_date_any) if "DateSaisie" in df_pay.columns else ""
        df_pay["montant_num"] = df_pay["Montant"].apply(parse_fr_number) if "Montant" in df_pay.columns else 0.0
        df_pay["montant_abs"] = df_pay["montant_num"].abs()

        df_pay["fact_no"] = df_pay["InvoiceNo"].apply(clean_invoice_no) if "InvoiceNo" in df_pay.columns else ""
        df_pay = df_pay[df_pay["montant_abs"] > 0.000001].copy()

        for _, r in df_pay.iterrows():
            date_ecr = sstr(r.get("date_fec", ""))
            piece_date = date_ecr

            fact_no = sstr(r.get("fact_no", ""))
            client_id = r.get("ClientID", "")
            client_name = r.get("ClientName", "")
            auxnum, auxlib = get_aux(client_id, client_name)

            mode = r.get("Mode", "")
            sous_mode = r.get("SousMode", "")
            typ = r.get("TypeLib", "")
            bord = r.get("Bord", "")
            echeance = r.get("Echeance", "")

            amt = float(r.get("montant_abs", 0.0))
            sign = 1 if float(r.get("montant_num", 0.0)) >= 0 else -1

            piece_ref = make_piece_ref("BORD", bord, "FAC", fact_no, fallback=f"RG-{date_ecr}-{ecriture_num}")
            let = fact_no

            lib = join_nonempty([
                "Encaissement" if sign > 0 else "Annulation règlement",
                f"Fact {fact_no}" if fact_no else "",
                client_name,
                f"ID:{client_id}" if sstr(client_id) else "",
                f"Mode:{mode}" if sstr(mode) else "",
                f"Sous:{sous_mode}" if sstr(sous_mode) else "",
                f"Type:{typ}" if sstr(typ) else "",
                f"Bord:{bord}" if sstr(bord) else "",
                f"Echéance:{echeance}" if sstr(echeance) else "",
            ])

            if sign > 0:
                fec_rows.append(fec_row(jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
                                        compte_banque, "Banque",
                                        "", "",
                                        piece_ref, piece_date, lib,
                                        Debit=amt, Credit=0.0,
                                        EcritureLet=let))
                fec_rows.append(fec_row(jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
                                        compte_client_global, "Clients",
                                        auxnum, auxlib,
                                        piece_ref, piece_date, lib,
                                        Debit=0.0, Credit=amt,
                                        EcritureLet=let))
            else:
                fec_rows.append(fec_row(jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
                                        compte_banque, "Banque",
                                        "", "",
                                        piece_ref, piece_date, lib,
                                        Debit=0.0, Credit=amt,
                                        EcritureLet=let))
                fec_rows.append(fec_row(jnl_bq, jnl_bq_lib, ecriture_num, date_ecr,
                                        compte_client_global, "Clients",
                                        auxnum, auxlib,
                                        piece_ref, piece_date, lib,
                                        Debit=amt, Credit=0.0,
                                        EcritureLet=let))

            ecriture_num += 1

    except Exception as e:
        st.error(f"Erreur traitement ENCAISSEMENTS: {e}")

# -----------------------------
# Affichages + Download
# -----------------------------
with tabs[0]:
    st.subheader("Logs mapping")
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
            "📥 Télécharger le FEC",
            data=fec_bytes,
            file_name=f"FEC_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )

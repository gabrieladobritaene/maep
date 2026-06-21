"""
Analiză cheltuieli din subvenție — aplicație Streamlit.

Utilizatorul încarcă fișierul Excel (.xls / .xlsx) folosit pentru raportare.
Aplicația RECALCULEAZĂ singură toate totalurile, pornind DOAR de la rândurile
de tranzacții (cele cu dată reală), ignorând complet rândurile de subtotal/total
din fișier. Astfel rezultatele nu depind de formulele din Excel (care pot fi greșite).

Structura fișierului se recunoaște după coloanele FIXE (Luna, data, Total cheltuită,
Total primită, Venituri în clarificare, Sume transferate, Sold cont), iar categoriile
de cheltuieli sunt deduse automat (tot ce stă între coloana de date și „Total cheltuită"),
deci merge chiar dacă se schimbă categoriile de la an la an.
"""

import io
import re
import datetime
import unicodedata

import numpy as np
import pandas as pd


# ----------------------------------------------------------------------------- #
#  LOGICĂ (pură, testabilă fără Streamlit)
# ----------------------------------------------------------------------------- #

def _norm(s):
    """Normalizează un text: fără diacritice, litere mici, spații colapsate."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s.lower()).strip()


def _to_date(v):
    """Întoarce un Timestamp DOAR pentru valori care sunt cu adevărat date
    (datetime). Textele de tip 'IANUARIE' NU sunt considerate date."""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        try:
            return pd.Timestamp(v)
        except Exception:
            return pd.NaT
    return pd.NaT


def _to_num(series):
    """Coerce la numeric, tratând textul/golul ca 0."""
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def read_excel_any(file_like, filename):
    engine = "xlrd" if filename.lower().rsplit(".", 1)[-1] == "xls" else "openpyxl"
    return pd.read_excel(file_like, header=None, engine=engine, dtype=object)


def detect_header_row(raw):
    """Caută rândul de antet (cel cu 'Luna' și 'Sold cont' / 'Total cheltuită')."""
    for i in range(min(15, len(raw))):
        vals = [_norm(v) for v in raw.iloc[i].tolist()]
        has_luna = any(v == "luna" or v.startswith("luna") for v in vals)
        has_sold = any("sold cont" in v for v in vals)
        has_tot = any("total" in v and "cheltuit" in v for v in vals)
        if has_luna and (has_sold or has_tot):
            return i
    return 0


def find_anchor_columns(headers):
    """headers = listă de antete normalizate. Întoarce indecșii coloanelor fixe."""
    def find(pred):
        for idx, h in enumerate(headers):
            if pred(h):
                return idx
        return None

    return {
        "luna": find(lambda h: h == "luna" or h.startswith("luna")),
        "total_cheltuit": find(lambda h: "total" in h and "cheltuit" in h),
        "total_primit": find(lambda h: "total" in h and "primit" in h),
        # IMPORTANT: 'venituri ... clarificare' (AE), nu categoria 'cheltuieli in curs de clarificare'
        "venit_clarificare": find(lambda h: "venitur" in h and "clarificare" in h),
        "transferat": find(lambda h: "transferate" in h or ("sume" in h and "filiale" in h)),
        "sold": find(lambda h: "sold cont" in h or h == "sold"),
        "furnizor": find(lambda h: "furnizor" in h),
    }


def detect_date_col(df, luna_idx, total_cheltuit_idx):
    """Coloana de date = cea cu cele mai multe date reale, între Luna și Total cheltuită."""
    start = (luna_idx + 1) if luna_idx is not None else 0
    end = total_cheltuit_idx if total_cheltuit_idx is not None else df.shape[1]
    best, best_cnt = None, -1
    for c in range(start, max(start, end)):
        cnt = int(df.iloc[:, c].map(lambda v: not pd.isna(_to_date(v))).sum())
        if cnt > best_cnt:
            best_cnt, best = cnt, c
    return best


MONTHS_RO = ["ianuarie", "februarie", "martie", "aprilie", "mai", "iunie",
             "iulie", "august", "septembrie", "octombrie", "noiembrie", "decembrie"]


def _luna_ro(period):
    return f"{MONTHS_RO[period.month - 1]} {period.year}"


def analyze(raw):
    """Analiza completă. Întoarce un dict cu rezultate sau ridică ValueError clar."""
    hdr = detect_header_row(raw)
    headers_norm = [_norm(v) for v in raw.iloc[hdr].tolist()]
    headers_raw = ["" if pd.isna(v) else str(v).strip() for v in raw.iloc[hdr].tolist()]
    cols = find_anchor_columns(headers_norm)

    if cols["total_cheltuit"] is None:
        raise ValueError("Nu am găsit coloana „Total subvenție cheltuită”. "
                         "Verifică antetul fișierului.")

    body_full = raw.iloc[hdr + 1:].reset_index(drop=True)

    # Limita blocului de sinteză: primul rând unde coloana Luna conține „total"
    # (ex. „TOTAL CATEGORIE"). De acolo în jos = sinteză manuală, o excludem.
    boundary, grand = None, None
    if cols["luna"] is not None:
        for i in range(len(body_full)):
            if "total" in _norm(body_full.iloc[i, cols["luna"]]):
                boundary = i
                val = pd.to_numeric(pd.Series([body_full.iloc[i, cols["total_cheltuit"]]]),
                                    errors="coerce").iloc[0]
                if pd.notna(val):
                    grand = {"rand": i + hdr + 2, "total_fisier": float(val)}
                break
    body = body_full.iloc[:boundary] if boundary is not None else body_full

    date_idx = detect_date_col(body, cols["luna"], cols["total_cheltuit"])
    if date_idx is None:
        raise ValueError("Nu am găsit coloana cu date (datele tranzacțiilor).")

    cat_idxs = list(range(date_idx + 1, cols["total_cheltuit"]))
    if not cat_idxs:
        raise ValueError("Nu am găsit nicio coloană de categorie între dată și „Total cheltuită”.")
    cat_names = [headers_raw[i] or f"Coloana {i + 1}" for i in cat_idxs]

    dates_all = body.iloc[:, date_idx].map(_to_date)
    mask_tx = dates_all.notna()
    tx = body[mask_tx]              # păstrăm indexul original (pt. nr. rând corect)
    tx_dates = dates_all[mask_tx]

    # --- sume pe rând ---
    cat_matrix = tx.iloc[:, cat_idxs].apply(_to_num)
    cat_matrix.columns = cat_names
    row_cheltuit = cat_matrix.sum(axis=1)
    primit = _to_num(tx.iloc[:, cols["total_primit"]]) if cols["total_primit"] is not None else pd.Series(0.0, index=tx.index)
    venit = _to_num(tx.iloc[:, cols["venit_clarificare"]]) if cols["venit_clarificare"] is not None else pd.Series(0.0, index=tx.index)
    transferat = _to_num(tx.iloc[:, cols["transferat"]]) if cols["transferat"] is not None else pd.Series(0.0, index=tx.index)
    movement = primit + venit - row_cheltuit - transferat

    # --- sold inițial (preluat din fișier, e dată reală nu formulă) ---
    opening = 0.0
    opening_source = "presupus 0 (nu am găsit coloana Sold cont)"
    if cols["sold"] is not None:
        sold_typed = pd.to_numeric(tx.iloc[:, cols["sold"]], errors="coerce")
        fv = sold_typed.first_valid_index()
        if fv is not None:
            opening = float(sold_typed.loc[fv] - movement.loc[fv])
            opening_source = "preluat din fișier (primul Sold cont)"

    running = opening + movement.cumsum()
    final_sold = float(opening + movement.sum())

    # --- pe categorii ---
    cat_tot = cat_matrix.sum(axis=0)
    cat_tot = cat_tot[cat_tot != 0].sort_values(ascending=False)
    total_cheltuit = float(row_cheltuit.sum())
    cat_table = pd.DataFrame({
        "Categorie": cat_tot.index,
        "Total (lei)": cat_tot.values,
        "% din cheltuieli": (cat_tot.values / total_cheltuit * 100) if total_cheltuit else 0,
    })

    # --- pe luni ---
    period = tx_dates.dt.to_period("M")
    mdf = pd.DataFrame({
        "p": period.values,
        "primit": primit.values, "cheltuit": row_cheltuit.values,
        "venit": venit.values, "transferat": transferat.values,
        "movement": movement.values,
    })
    g = mdf.groupby("p", sort=True).sum(numeric_only=True)
    g["sold_sf_luna"] = opening + g["movement"].cumsum()
    month_table = pd.DataFrame({
        "Luna": [_luna_ro(p) for p in g.index],
        "Primit (lei)": g["primit"].values,
        "Cheltuit (lei)": g["cheltuit"].values,
        "Sold la sfârșit de lună (lei)": g["sold_sf_luna"].values,
    })

    # --- pe furnizor ---
    supplier_table = None
    if cols["furnizor"] is not None:
        sdf = pd.DataFrame({
            "Furnizor": tx.iloc[:, cols["furnizor"]].map(lambda v: "(necompletat)" if pd.isna(v) else str(v).strip()).values,
            "cheltuit": row_cheltuit.values,
        })
        sdf = sdf[sdf["cheltuit"] != 0]
        supplier_table = (sdf.groupby("Furnizor", as_index=False)["cheltuit"].sum()
                          .sort_values("cheltuit", ascending=False)
                          .rename(columns={"cheltuit": "Total cheltuit (lei)"}))

    # --- sold pe timp (pentru grafic) ---
    sold_curve = pd.DataFrame({"Data": tx_dates.values, "Sold (lei)": running.values})
    sold_curve = sold_curve.dropna().sort_values("Data")

    # ======================= VERIFICĂRI =======================
    checks = {}

    # 1) Total pe rând tastat ≠ suma categoriilor
    if cols["total_cheltuit"] is not None:
        typed_ac = pd.to_numeric(tx.iloc[:, cols["total_cheltuit"]], errors="coerce")
        diff = (typed_ac - row_cheltuit)
        bad = diff.abs() > 0.01
        rows = []
        for i in np.where(bad & typed_ac.notna())[0]:
            rows.append({
                "Rând în fișier": int(tx.index[i]) + hdr + 2,
                "Total tastat": round(float(typed_ac.iloc[i]), 2),
                "Suma categoriilor": round(float(row_cheltuit.iloc[i]), 2),
                "Diferență": round(float(diff.iloc[i]), 2),
            })
        checks["total_rand_gresit"] = pd.DataFrame(rows)

    # 2) Tranzacții reale FĂRĂ dată: au furnizor + bani, dar nu au dată -> NU sunt
    #    numărate în totaluri (rândurile de subtotal lunar nu au furnizor, deci sunt excluse aici).
    no_date = body[~mask_tx]
    if len(no_date) and cols["furnizor"] is not None:
        ndc = no_date.iloc[:, cat_idxs].apply(_to_num).sum(axis=1)
        furn = no_date.iloc[:, cols["furnizor"]]

        def _real_supplier(v):
            if pd.isna(v):
                return False
            s = str(v).strip()
            if s == "":
                return False
            return not re.fullmatch(r"[-\d.,\s]+", s)  # nume real conține litere

        has_furn = furn.map(_real_supplier)
        skipped = no_date[(ndc != 0) & has_furn]
        rows = []
        for idx in skipped.index:
            rows.append({
                "Rând în fișier": int(idx) + hdr + 2,
                "Furnizor": str(furn.loc[idx]).strip(),
                "Sumă (lei)": round(float(ndc.loc[idx]), 2),
            })
        checks["orfane"] = pd.DataFrame(rows)

    # 3) Sold din fișier ≠ sold recalculat — rezumat (nu listă lungă)
    if cols["sold"] is not None:
        sold_typed = pd.to_numeric(tx.iloc[:, cols["sold"]], errors="coerce")
        sdiff = (sold_typed - running)
        mism = (sdiff.abs() > 0.01) & sold_typed.notna()
        idxs = np.where(mism)[0]
        if len(idxs):
            first = int(idxs[0])
            last_valid = sold_typed.last_valid_index()
            checks["sold"] = {
                "n": int(len(idxs)),
                "first_row": int(tx.index[first]) + hdr + 2,
                "first_typed": round(float(sold_typed.iloc[first]), 2),
                "first_correct": round(float(running.iloc[first]), 2),
                "final_typed": round(float(sold_typed.loc[last_valid]), 2) if last_valid is not None else None,
                "final_correct": round(float(running.iloc[-1]), 2),
            }
        else:
            checks["sold"] = {"n": 0}

    return {
        "header_row": hdr, "cols": cols, "date_idx": date_idx,
        "cat_names": cat_names, "n_tx": int(len(tx)),
        "period_min": tx_dates.min(), "period_max": tx_dates.max(),
        "opening": opening, "opening_source": opening_source,
        "total_cheltuit": total_cheltuit,
        "total_primit": float(primit.sum()),
        "total_venit": float(venit.sum()),
        "total_transferat": float(transferat.sum()),
        "final_sold": final_sold,
        "cat_table": cat_table, "month_table": month_table,
        "supplier_table": supplier_table, "sold_curve": sold_curve,
        "checks": checks, "grand": grand,
    }


def build_export(res):
    """Construiește un Excel curat cu totalurile corecte."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        sumar = pd.DataFrame({
            "Indicator": ["Perioadă", "Număr tranzacții", "Sold inițial",
                          "Total primit", "Total cheltuit", "Sume transferate",
                          "Sold final"],
            "Valoare": [
                f"{res['period_min']:%d.%m.%Y} – {res['period_max']:%d.%m.%Y}",
                res["n_tx"], round(res["opening"], 2), round(res["total_primit"], 2),
                round(res["total_cheltuit"], 2), round(res["total_transferat"], 2),
                round(res["final_sold"], 2),
            ],
        })
        sumar.to_excel(xl, sheet_name="Sumar", index=False)
        res["cat_table"].round(2).to_excel(xl, sheet_name="Pe categorii", index=False)
        res["month_table"].round(2).to_excel(xl, sheet_name="Pe luni", index=False)
        if res["supplier_table"] is not None:
            res["supplier_table"].round(2).to_excel(xl, sheet_name="Pe furnizor", index=False)
    buf.seek(0)
    return buf.getvalue()


# ----------------------------------------------------------------------------- #
#  INTERFAȚĂ STREAMLIT
# ----------------------------------------------------------------------------- #

def _fmt(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " lei"


def main():
    import streamlit as st

    st.set_page_config(page_title="Verificare cheltuieli subvenție", page_icon="📊", layout="wide")

    st.title("📊 Verificare cheltuieli din subvenție")
    st.markdown(
        "Încarcă fișierul Excel folosit pentru raportare. Aplicația **recalculează singură** "
        "toate totalurile, pornind doar de la rândurile de tranzacții, și îți arată unde "
        "fișierul are eventuale diferențe. **Tu nu trebuie să faci nimic tehnic — doar încarci fișierul.**"
    )

    up = st.file_uploader("Încarcă fișierul (.xls sau .xlsx)", type=["xls", "xlsx"])
    if up is None:
        st.info("Aștept fișierul. După încărcare, vei vedea totalurile pe secțiuni.")
        st.stop()

    try:
        raw = read_excel_any(up, up.name)
        res = analyze(raw)
    except Exception as e:
        st.error(f"Nu am putut procesa fișierul: {e}")
        st.stop()

    # ---- 1. SUMAR ----
    st.header("1. Sumar general")
    st.caption(f"Perioadă acoperită: **{res['period_min']:%d.%m.%Y} – {res['period_max']:%d.%m.%Y}** · "
               f"**{res['n_tx']}** tranzacții · {len(res['cat_names'])} categorii de cheltuieli detectate.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total primit", _fmt(res["total_primit"]))
    c2.metric("Total cheltuit", _fmt(res["total_cheltuit"]))
    c3.metric("Sume transferate", _fmt(res["total_transferat"]))
    c4.metric("Sold final", _fmt(res["final_sold"]))
    st.caption(f"Sold inițial folosit: **{_fmt(res['opening'])}** ({res['opening_source']}).")

    st.divider()

    # ---- 2. PE CATEGORII ----
    st.header("2. Cheltuieli pe categorii")
    st.caption("Cât s-a cheltuit pe fiecare categorie, pe toată perioada. Sortat descrescător.")
    ct = res["cat_table"]
    st.dataframe(
        ct.style.format({"Total (lei)": "{:,.2f}", "% din cheltuieli": "{:.1f}%"}),
        use_container_width=True, hide_index=True,
    )
    st.bar_chart(ct.set_index("Categorie")["Total (lei)"], horizontal=True)

    st.divider()

    # ---- 3. PE LUNI ----
    st.header("3. Pe luni")
    st.caption("Pentru fiecare lună: cât s-a primit, cât s-a cheltuit și soldul la sfârșitul lunii.")
    mt = res["month_table"]
    st.dataframe(
        mt.style.format({c: "{:,.2f}" for c in mt.columns if c != "Luna"}),
        use_container_width=True, hide_index=True,
    )

    st.divider()

    # ---- 4. EVOLUȚIA SOLDULUI ----
    st.header("4. Evoluția soldului în timp")
    st.caption("Soldul contului după fiecare tranzacție (recalculat corect).")
    st.line_chart(res["sold_curve"].set_index("Data")["Sold (lei)"])

    st.divider()

    # ---- 5. PE FURNIZOR ----
    if res["supplier_table"] is not None and len(res["supplier_table"]):
        st.header("5. Pe furnizor")
        st.caption("Top furnizori după suma totală cheltuită.")
        top = st.slider("Câți furnizori să afișez?", 5, 50, 15)
        sup = res["supplier_table"].head(top)
        st.dataframe(
            sup.style.format({"Total cheltuit (lei)": "{:,.2f}"}),
            use_container_width=True, hide_index=True,
        )
        st.divider()

    # ---- 6. VERIFICĂRI ----
    st.header("6. Verificări — unde fișierul ar putea avea probleme")
    st.caption("Acestea sunt diferențe între ce e scris în fișier și calculul corect. "
               "Nu modifică fișierul — doar îți arată ce merită verificat.")
    checks = res["checks"]

    # headline dublă-numărare
    if res.get("grand"):
        g = res["grand"]
        diff = g["total_fisier"] - res["total_cheltuit"]
        if abs(diff) > 0.01:
            st.warning(
                f"**Totalul general din fișier (rândul {g['rand']}) = {_fmt(g['total_fisier'])}**, "
                f"dar totalul corect recalculat = **{_fmt(res['total_cheltuit'])}**. "
                f"Diferență: {_fmt(diff)}. (De obicei înseamnă dublă-numărare în rândul de total.)"
            )
        else:
            st.success(f"Totalul general din fișier coincide cu calculul corect ({_fmt(res['total_cheltuit'])}).")

    def show_check(key, ok_msg, bad_title):
        df = checks.get(key)
        if df is None:
            return
        if len(df) == 0:
            st.success(ok_msg)
        else:
            with st.expander(f"⚠️ {bad_title} — {len(df)} rând(uri)"):
                st.dataframe(df, use_container_width=True, hide_index=True)

    show_check("total_rand_gresit",
               "Pe fiecare rând, totalul tastat coincide cu suma categoriilor.",
               "Rânduri unde totalul tastat diferă de suma categoriilor")
    show_check("orfane",
               "Toate cheltuielile au dată — nimic exclus din calcul.",
               "Tranzacții cu cheltuieli dar FĂRĂ dată (NU au fost numărate — adaugă data în fișier)")

    sold = checks.get("sold")
    if sold is not None:
        if sold["n"] == 0:
            st.success("Soldul din fișier coincide cu soldul recalculat pe tot parcursul.")
        else:
            st.warning(
                f"**Soldul din fișier diferă de cel corect.** Prima diferență apare la "
                f"**rândul {sold['first_row']}** (fișier: {_fmt(sold['first_typed'])}, "
                f"corect: {_fmt(sold['first_correct'])}). La final, fișierul arată "
                f"**{_fmt(sold['final_typed'])}**, iar soldul corect este "
                f"**{_fmt(sold['final_correct'])}**."
            )

    st.divider()

    # ---- 7. EXPORT ----
    st.header("7. Descarcă rezumatul")
    st.caption("Un fișier Excel curat, cu totalurile corecte (sumar, pe categorii, pe luni, pe furnizor).")
    st.download_button(
        "⬇️ Descarcă rezumat Excel",
        data=build_export(res),
        file_name="rezumat_subventie.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()

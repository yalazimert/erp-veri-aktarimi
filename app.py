import io
import json
import os
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------
# App config
# ---------------------------
APP_TITLE = "ERP Veri Aktarımı"

# templates klasörünü app.py’nin yanına sabitle
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TEMPLATE_DIR, exist_ok=True)

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.write(
    "Kaynak Excel ve Hedef Excel yükleyin. Hedef Excel’in kolon sırasını koruyarak eşleştirme yapın. "
    "İsterseniz eşleştirmeyi şablon olarak kaydedip tekrar kullanın."
)

# ---------------------------
# Helpers
# ---------------------------
def read_excel(file) -> dict:
    """Return {sheet_name: DataFrame}"""
    xls = pd.ExcelFile(file)
    return {s: pd.read_excel(xls, sheet_name=s, engine="openpyxl") for s in xls.sheet_names}

def list_templates():
    return sorted([fn for fn in os.listdir(TEMPLATE_DIR) if fn.endswith(".json")])

def save_template(name: str, payload: dict):
    safe = "".join(c for c in name if c.isalnum() or c in ("-", "_", " ")).strip().replace(" ", "_")
    if not safe:
        safe = "template"
    path = os.path.join(TEMPLATE_DIR, f"{safe}.json")
    payload["_saved_at"] = datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return path

def load_template(filename: str) -> dict:
    path = os.path.join(TEMPLATE_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def normalize_rule(rule, blank_option, manual_option):
    """
    Backward compatible:
    - Old templates: mapping[tgt] = "SourceCol" or "(Boş)"
    - New templates: mapping[tgt] = {"type": "...", "value": "..."}
    """
    if rule is None:
        return {"type": "blank"}

    if isinstance(rule, dict) and "type" in rule:
        t = rule.get("type")
        if t == "source":
            return {"type": "source", "value": str(rule.get("value", ""))}
        if t == "manual":
            return {"type": "manual", "value": str(rule.get("value", ""))}
        return {"type": "blank"}

    if isinstance(rule, str):
        if rule == blank_option:
            return {"type": "blank"}
        if rule == manual_option:
            return {"type": "manual", "value": ""}
        return {"type": "source", "value": rule}

    return {"type": "blank"}

def transform(src: pd.DataFrame, tgt_cols_order: list, mapping: dict) -> pd.DataFrame:
    out = pd.DataFrame()
    for tgt_col in tgt_cols_order:
        rule = mapping.get(tgt_col, {"type": "blank"})

        if rule["type"] == "source":
            src_col = rule.get("value", "")
            out[tgt_col] = src[src_col] if src_col in src.columns else pd.NA

        elif rule["type"] == "manual":
            out[tgt_col] = rule.get("value", "")

        else:
            out[tgt_col] = pd.NA

    return out

# ---------------------------
# UI: Uploads
# ---------------------------
c1, c2 = st.columns([1.1, 1])

with c1:
    st.subheader("1) Dosyaları yükle")
    src_file = st.file_uploader("Kaynak Excel", type=["xlsx"], key="src")
    tgt_file = st.file_uploader("Hedef Excel (kolon sırası buradan alınır)", type=["xlsx"], key="tgt")

if not src_file or not tgt_file:
    st.info("Devam etmek için hem Kaynak hem Hedef Excel yükleyin.")
    st.stop()

src_sheets = read_excel(src_file)
tgt_sheets = read_excel(tgt_file)

with c1:
    st.subheader("2) Sayfa seç")
    src_sheet = st.selectbox("Kaynak sayfa", list(src_sheets.keys()), index=0)
    tgt_sheet = st.selectbox("Hedef sayfa", list(tgt_sheets.keys()), index=0)

src_df = src_sheets[src_sheet]
tgt_df = tgt_sheets[tgt_sheet]

src_cols = [str(c) for c in list(src_df.columns)]
tgt_cols = [str(c) for c in list(tgt_df.columns)]

with c2:
    st.subheader("Özet")
    st.write(f"Kaynak kolon sayısı: **{len(src_cols)}**")
    st.write(f"Hedef kolon sayısı: **{len(tgt_cols)}**")
    with st.expander("Hedef kolonları (sıra korunur)", expanded=False):
        st.write(tgt_cols)

st.divider()

# ---------------------------
# Templates UI
# ---------------------------
st.subheader("3) Şablon (yükle / kaydet)")

t1, t2, t3 = st.columns([2, 2, 3])

with t3:
    if st.button("Şablon listesini yenile"):
        st.rerun()

templates = list_templates()

with t1:
    chosen_tpl = st.selectbox("Kayıtlı şablon yükle", ["(Seçme)"] + templates, key="tpl_select")

loaded_tpl = None
if chosen_tpl != "(Seçme)":
    try:
        loaded_tpl = load_template(chosen_tpl)
        st.success(f"Şablon yüklendi: {chosen_tpl}")
    except Exception as e:
        st.error(f"Şablon okunamadı: {e}")

# Options
blank_option = "(Boş)"
manual_option = "(Manuel Değer Gir)"
options = [blank_option, manual_option] + src_cols

prefill_mapping = loaded_tpl.get("mapping", {}) if loaded_tpl else {}

# ---------------------------
# Mapping UI
# ---------------------------
st.subheader("4) Kolon eşleştirme")
st.caption("Hedefteki her kolon için kaynaktan kolon seçin veya sabit değer girmek için '(Manuel Değer Gir)' seçin.")

mapping = {}

mleft, mright = st.columns([3, 2])

with mleft:
    # UI kalabalık olmasın diye arama kutusu: hedef kolon adında filtre
    search = st.text_input("Hedef kolonlarda ara (opsiyonel)", value="")
    visible_tgts = [c for c in tgt_cols if search.strip().lower() in c.lower()] if search.strip() else tgt_cols

    for i, tgt in enumerate(visible_tgts):
        pre_rule = normalize_rule(prefill_mapping.get(tgt), blank_option, manual_option)

        if pre_rule["type"] == "blank":
            default_choice = blank_option
        elif pre_rule["type"] == "manual":
            default_choice = manual_option
        else:
            default_choice = pre_rule.get("value", blank_option)

        if default_choice not in options:
            default_choice = blank_option

        choice = st.selectbox(
            f"Hedef: {tgt}",
            options,
            index=options.index(default_choice),
            key=f"map_{tgt}"
        )

        if choice == manual_option:
            default_manual = pre_rule.get("value", "") if pre_rule["type"] == "manual" else ""
            val = st.text_input(
                f"{tgt} için manuel değer",
                value=default_manual,
                key=f"manual_{tgt}"
            )
            mapping[tgt] = {"type": "manual", "value": val}
        elif choice == blank_option:
            mapping[tgt] = {"type": "blank"}
        else:
            mapping[tgt] = {"type": "source", "value": choice}

# Hedef kolonların hepsini mapping’e koymak lazım (arama ile filtrelesek bile)
# Görünmeyenleri de prefill üzerinden veya boş olarak dolduralım
for tgt in tgt_cols:
    if tgt not in mapping:
        pre_rule = normalize_rule(prefill_mapping.get(tgt), blank_option, manual_option)
        mapping[tgt] = pre_rule if pre_rule else {"type": "blank"}

with mright:
    st.subheader("Özet")
    used_sources = [v.get("value") for v in mapping.values() if v.get("type") == "source"]
    used_manual = [k for k, v in mapping.items() if v.get("type") == "manual" and str(v.get("value", "")).strip() != ""]

    st.write(f"Dolu eşleşme sayısı: **{len(set(used_sources)) + len(used_manual)} / {len(tgt_cols)}**")
    st.write("Manuel değer girilen kolonlar:")
    st.write(used_manual if used_manual else "—")

st.divider()

with t2:
    tpl_name = st.text_input("Şablon adı", value="")

with t3:
    if st.button("Şablonu Kaydet", type="primary"):
        payload = {
            "source_sheet": src_sheet,
            "target_sheet": tgt_sheet,
            "source_columns_snapshot": src_cols,
            "target_columns_snapshot": tgt_cols,
            "mapping": mapping
        }
        path = save_template(tpl_name or "eslestirme", payload)
        st.success(f"Şablon kaydedildi: {path}")
        # Kaydettikten sonra selectbox listesini hemen güncelle
        st.rerun()

st.divider()

# ---------------------------
# Transform + download
# ---------------------------
st.subheader("5) Dönüştür ve indir")

out_df = transform(src_df, tgt_cols, mapping)

with st.expander("Önizleme (ilk 20 satır)", expanded=True):
    st.dataframe(out_df.head(20), use_container_width=True)

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    out_df.to_excel(writer, sheet_name="Output", index=False)
output.seek(0)

st.download_button(
    label="Çıktı Excel'i indir",
    data=output,
    file_name="erp_aktarim.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

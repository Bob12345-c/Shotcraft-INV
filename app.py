
import streamlit as st
import pandas as pd
import math
from io import BytesIO

st.set_page_config(page_title="Shotcraft Case-Based Inventory (Upload)", layout="wide")
st.title("Shotcraft — Simple Upload & Calculate")
st.caption("Upload your Excel, enter Cases Sold, and see Required & Remaining per material.")

def load_excel(file):
    # Try common sheet names
    xls = pd.ExcelFile(file)
    sheets = [s.lower() for s in xls.sheet_names]
    # Find formula-like sheet
    formula_sheet = None
    for name in ["formula_695_cases","formula"]:
        if name in sheets:
            formula_sheet = xls.sheet_names[sheets.index(name)]
            break
    if formula_sheet is None:
        # fallback to first sheet
        formula_sheet = xls.sheet_names[0]

    df_formula = pd.read_excel(xls, sheet_name=formula_sheet)
    # Normalize columns
    cols = {c.lower().strip(): c for c in df_formula.columns}
    # Expect at least Component + Per_Case; UOM is nice to have
    comp_col = cols.get("component") or list(df_formula.columns)[0]
    per_case_col = cols.get("per_case")
    uom_col = cols.get("uom")

    # If no Per_Case, try Batch_Qty / detected batch cases cell
    if per_case_col is None:
        batch_col = cols.get("batch_qty")
        # try to detect batch cases (G2 used in our templates)
        batch_cases = 695.0
        if batch_col:
            df_formula["Per_Case"] = df_formula[cols["batch_qty"]] / float(batch_cases)
            per_case_col = "Per_Case"
        else:
            st.error("Couldn't find a Per_Case or Batch_Qty column in your formula sheet.")
            return None, None, None

    components = df_formula[[comp_col, per_case_col]].copy()
    components.columns = ["Component","Per_Case"]
    if uom_col:
        components["UOM"] = df_formula[uom_col]
    else:
        components["UOM"] = ""

    # Drop blanks and zeros safely
    components = components[components["Component"].notna()].reset_index(drop=True)
    # Try to load on-hand from INVENTORY if present
    onhand_df = None
    inv_sheet = None
    for name in xls.sheet_names:
        if name.lower() == "inventory":
            inv_sheet = name
            break
    if inv_sheet:
        inv = pd.read_excel(xls, sheet_name=inv_sheet)
        inv_cols = {c.lower().strip(): c for c in inv.columns}
        if "component" in inv_cols and "on_hand" in inv_cols:
            onhand_df = inv[[inv_cols["component"], inv_cols["on_hand"]]].rename(columns={inv_cols["component"]:"Component", inv_cols["on_hand"]:"On_Hand"})
    return components, onhand_df, formula_sheet

def compute_results(components, onhand_df, cases_sold):
    df = components.copy()
    # Merge on-hand if available
    if onhand_df is not None:
        df = df.merge(onhand_df, on="Component", how="left")
    if "On_Hand" not in df.columns:
        df["On_Hand"] = 0.0

    # Ensure numeric
    df["Per_Case"] = pd.to_numeric(df["Per_Case"], errors="coerce").fillna(0.0)
    df["On_Hand"] = pd.to_numeric(df["On_Hand"], errors="coerce").fillna(0.0)

    df["Required"] = df["Per_Case"] * float(cases_sold)
    df["Remaining"] = df["On_Hand"] - df["Required"]

    # Bottleneck (max sellable cases from current On_Hand)
    candidates = df[(df["Per_Case"]>0)]
    if not candidates.empty:
        df["MaxCasesByItem"] = df.apply(lambda r: (r["On_Hand"]/r["Per_Case"]) if r["Per_Case"]>0 else float("inf"), axis=1)
        max_sellable = math.floor(df["MaxCasesByItem"].min())
        bottleneck_row = df.loc[df["MaxCasesByItem"].idxmin()]
        bottleneck = {
            "Item": bottleneck_row["Component"],
            "UOM": bottleneck_row.get("UOM",""),
            "On_Hand": bottleneck_row["On_Hand"],
            "Per_Case": bottleneck_row["Per_Case"],
            "MaxCasesByItem": bottleneck_row["MaxCasesByItem"],
        }
    else:
        max_sellable = 0
        bottleneck = None

    # Shortages for this order
    shortages = df[df["Remaining"] < 0][["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].copy()

    # Order results table nicely
    display = df[["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].sort_values("Component")
    return display, int(max_sellable), shortages

def download_updated_inventory(display, original_onhand, formula_sheet_name):
    # Create an Excel with two sheets: FORMULA (components with per_case) and INVENTORY (updated on_hand + computed fields)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        # Write formula-like sheet
        formula = display[["Component","UOM","Per_Case"]].copy()
        formula.to_excel(writer, sheet_name=formula_sheet_name or "FORMULA", index=False)
        # Write inventory sheet
        inv = display.copy()
        inv.to_excel(writer, sheet_name="INVENTORY", index=False)
    out.seek(0)
    return out

uploaded = st.file_uploader("Upload your Shotcraft Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    comps, onhand_df, formula_name = load_excel(uploaded)
    if comps is not None:
        st.success(f"Loaded {len(comps)} components from sheet: {formula_name}")
        st.write("Per-case usage (from your file):")
        st.dataframe(comps, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("Enter order size")
        cases = st.number_input("Cases Sold (e.g., LCBO order)", min_value=0.0, step=1.0, value=0.0)
        results, max_sellable, shortages = compute_results(comps, onhand_df, cases)

        st.markdown("### Results")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Max sellable cases from current stock", max_sellable)
        with col2:
            st.metric("Order size entered (cases)", int(cases))

        st.dataframe(results, use_container_width=True, hide_index=True)

        if not shortages.empty:
            st.warning("Shortages for this order:")
            st.dataframe(shortages, use_container_width=True, hide_index=True)
        else:
            st.info("No shortages detected for this order.")

        # Allow download of updated snapshot
        st.markdown("### Download updated snapshot")
        buf = download_updated_inventory(results, onhand_df, formula_name)
        st.download_button("Download Excel snapshot", buf, file_name="Shotcraft_Inventory_Snapshot.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Could not load required columns. Please ensure your file has a sheet with 'Component' and 'Per_Case' columns.")
else:
    st.info("Upload your Excel to begin.")

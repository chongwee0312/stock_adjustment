import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
from fuzzywuzzy import process, fuzz
import io

st.set_page_config(page_title="Stock & Order Matcher", layout="wide")
st.title("📊 Stock and Order Matcher")

# --- File upload section ---
st.sidebar.header("1. Upload Files")
st.sidebar.info('Export from PowerClinic: Inventory Management -> Stock Management -> Stock Take -> Print -> Excel')
stock_file = st.sidebar.file_uploader("Upload Stock File (.xls)", type="xls")
st.sidebar.info('Your Self-Defined Stock Order')
order_file = st.sidebar.file_uploader("Upload Desired Order File (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"])

if stock_file and order_file:
    try:
        # --- Load Stock File ---
        stock = pd.read_excel(stock_file)
        stock = stock.dropna(how='all', axis=1).dropna(how='all').reset_index(drop=True)
        stock.columns = stock.loc[6]
        stock.columns = [str(col).strip().lower().replace(' ', '_').replace('.', '').replace("'", '') for col in stock.columns]
        stock = stock.loc[7:].reset_index(drop=True)

        # Rename unnamed columns
        new_columns = []
        counter = 0
        for col in stock.columns:
            if col == 'nan':
                counter += 1
                new_columns.append(f'na_{counter}')
            else:
                new_columns.append(col)
        stock.columns = new_columns

        stock['on_hand_qty'] = stock['on_hand_qty'].fillna(stock.get('na_3'))
        stock = stock.drop(columns=[col for col in ['na_1', 'na_2', 'na_3'] if col in stock.columns])
        stock = stock.dropna(how='all')
        stock = stock[stock['name'].notna() & (stock['actual_qty'] != "Actual Q'ty")].reset_index(drop=True)
        stock['item_no'] = stock['item_no'].str.strip().str.upper()

        # --- Load Order File ---
        if order_file.name.endswith("csv"):
            order_df = pd.read_csv(order_file)
        else:
            order_df = pd.read_excel(order_file)

        order_df = order_df.dropna(how='all', axis=1).dropna(how='all').reset_index(drop=True)

        for col in order_df.columns:
            if 'Unnamed' in col and 'item_no' in order_df[col].astype(str).str.lower().values:
                name_row = order_df[order_df[col].astype(str).str.lower() == 'item_no'].index[0]
                order_df.columns = order_df.loc[name_row]
                order_df = order_df.loc[name_row + 1:].reset_index(drop=True)
                break

        order_df.columns = [str(col).lower().strip().replace('.', '').replace(' ', '_') for col in order_df.columns]

        if 'item_no' not in order_df.columns:
            st.error("❌ 'item_no' column not found in the order file. Please reupload with correct headers.")
            st.stop()

        order_df['item_no'] = order_df['item_no'].str.strip().str.upper()

        # --- Categorize ---
        consume = [item for item in stock['item_no'].unique() if item.startswith('(C)')]
        pharma = [item for item in stock['item_no'].unique() if item not in consume]
        pharma_df = stock[stock['item_no'].isin(pharma)]
        consume_df = stock[stock['item_no'].isin(consume)]

        # --- Stock Summary ---
        total_items = stock.shape[0]
        total_pharma = pharma_df.shape[0]
        total_consume = consume_df.shape[0]

        st.markdown("### 📊 Stock Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Items", total_items)
        col2.metric("Pharma Items", total_pharma)
        col3.metric("Consumables", total_consume)

        # --- Check for duplicates ---
        if order_df['item_no'].duplicated().any():
            st.warning("⚠️ Duplicated items found in the order file.")
            st.dataframe(order_df[order_df['item_no'].duplicated(keep=False)])

        # --- Merge stock with order ---
        final = order_df.merge(
            stock[['item_no', 'name', 'on_hand_qty', 'actual_qty']],
            on='item_no', how='left', indicator=True
        )

        # --- Check unmatched ---
        unmatched = final[final['_merge'] == 'left_only']
        if not unmatched.empty:
            st.error("Some items in the desired order file could not be found in the system stock file.")
            not_found = unmatched['item_no'].tolist()
            search_pool = stock['item_no'].unique().tolist()
            match_list = []

            for item in not_found:
                match = process.extractOne(item, search_pool)
                matched = match[0] if match and match[1] >= 90 else ''
                match_list.append(matched)

            possible_matches = pd.DataFrame({
                'Index': range(len(not_found)),
                'Item Not Found': not_found,
                'Suggested Match': match_list
            })

            all_indices = possible_matches['Index'].tolist()
            select_all = st.checkbox("Select all unmatched items")

            selected_indices = st.multiselect(
                "Select the index of items to replace with suggested matches",
                help="This will update the name and merge the stock quantities.",
                options=all_indices,
                default=all_indices if select_all else []
            )

            for i in selected_indices:
                original = possible_matches.at[i, 'Item Not Found']
                replacement = possible_matches.at[i, 'Suggested Match']
                final.loc[final['item_no'] == original, 'item_no'] = replacement

            # Re-merge to update quantities and names
            final = final.drop(columns=['name', 'on_hand_qty', 'actual_qty', '_merge'], errors='ignore')
            final = final.merge(
                stock[['item_no', 'name', 'on_hand_qty', 'actual_qty']],
                on='item_no', how='left'
            )

            st.dataframe(possible_matches)
            st.markdown("""
                ✅ Selected items will now show their stock quantities.<br>
                ❌ Unselected unmatched items remain unchanged.
            """, unsafe_allow_html=True)

        else:
            final = final.drop(columns=['_merge'])

        # --- Download merged file ---
        buffer = io.BytesIO()
        final.to_csv(buffer, index=False)
        st.download_button("Download Merged File", data=buffer.getvalue(), file_name="merged_stock_order.csv")

    except Exception as e:
        st.error(f"❌ An error occurred while processing the files: {e}")

else:
    st.info("Please upload both stock and order files to begin.")

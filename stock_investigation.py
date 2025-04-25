import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
from fuzzywuzzy import process, fuzz

st.set_page_config(page_title='Stock Adjustment', layout='wide')
st.title('🧾 Stock Adjustment')

# Initial files needed
master_file = None
system_file = None

# Master file upload
if not master_file:
    st.info('ℹ️ Please upload the file with Stock To Be Checked.')
    master_file = st.file_uploader('Upload: Stock To Be Checked (csv format with "item_no" and "stock_on_hand" columns)')
    system_file = None

# Load master file
if master_file:
    if not master_file.name.endswith('.csv'):
        st.error('❌ Invalid file format. Please upload a "csv" file.')
        st.stop()
    
    master = pd.read_csv(master_file)

    master.columns = [col.lower().strip().replace('.', '').replace(' ', '_') for col in master.columns]

    try:
        master = master[['item_no', 'stock_on_hand']]
        master['item_no'] = master['item_no'].str.strip().str.upper()
    except:
        st.error('❌ Required columns ("item_no" or "stock_on_hand") cannot be found.')
        st.error('Please upload the file with required columns.')
        st.stop()
        
    # Check for duplicates
    duplicated = master[master['item_no'].duplicated(keep=False)]
    if not duplicated.empty:
        st.subheader('Duplicated items:')
        st.dataframe(duplicated.sort_values('item_no'))
        st.warning('⚠️ Duplicated items found in the Stock To Be Checked.')
        st.warning('⚠️ Please reupload the correct data.')
        st.stop()

    # Check for missing values
    item_na = master['item_no'].isna().sum()
    count_na = master['stock_on_hand'].isna().sum()

    if item_na > 0:        
        st.subheader('Items without Item No:')
        st.dataframe(master[master['item_no'].isna()])
        st.error(f'❌ There are {item_na} unidentified items in Stock To Be Checked list.')
        st.error('❌ Please reupload the file with "item_no" filled in.')
        st.stop()
        
    if count_na > 0:
        st.subheader('Items without Stock Count:')
        st.dataframe(master[master['stock_on_hand'].isna()])
        st.error(f'❌ There are {count_na} items without stock count in Stock To Be Checked list.')
        st.error('❌ Please reupload the file with "stock_on_hand" filled in.')
        st.stop()
        
    # Ensure numeric type
    non_numeric = pd.to_numeric(master['stock_on_hand'], errors='coerce').isna()
    if non_numeric.any():        
        st.subheader('Non-Numeric "stock_on_hand" Values')
        st.dataframe(master[non_numeric])
        st.error('❌ Please ensure that "stock_on_hand" is numeric.')
        st.error('❌ Please reupload the file with correct stock count.')
        st.stop()
    else:
        master['stock_on_hand'] = master['stock_on_hand'].astype(float)

    # Ensure stock count > 0
    negative_stock = master[master['stock_on_hand'] < 0]
    if negative_stock.shape[0] > 0:
        st.subheader('Negative "stock_on_hand" Values')
        st.dataframe(negative_stock)
        st.error('❌ Please ensure that "stock_on_hand" is positive.')
        st.error('❌ Please reuploaad the file with correct stock count.')
        st.stop()
        
    # Check total items to be check
    total_item = master.shape[0]
    if total_item == 0:
        st.warning('⚠️ No items in the file. Please check again.')
        st.stop()
    elif total_item == 1:
        st.success(f'There is {total_item} item to be checked.')
    else:
        st.success(f'There are {total_item} items to be checked.')

# System file upload
if master_file:
    st.info('ℹ️ Please upload the file with System Stock Count.')
    system_file = st.file_uploader('Upload: System Stock (PowerClinic -> Stock Management -> Stock Status -> Export Grid)')

# Load system file
if system_file:
    if not system_file.name.endswith('.xlsx'):
        st.error('❌ Invalid file format. Please upload a "xlsx" file.')
        st.stop()
    
    system = pd.read_excel(system_file, header=1)

    # Clean system
    system.columns = [col.lower().strip().replace("'", '').replace(' ', '_').replace('.', '') for col in system.columns]
    system = system[['item_no', 'item_name', 'on_hand_qty']]
    system.columns = ['item_no', 'item_name', 'stock_in_system']
    system['item_no'] = system['item_no'].str.strip().str.upper()

    # Merge
    check = master.merge(system, on='item_no', how='left', indicator=True)
    check = check[['item_no', 'item_name', 'stock_on_hand', 'stock_in_system', '_merge']]
    
    # Check unmatched items
    left_only = check.query('_merge == "left_only"')
    unmatched_count = left_only.shape[0]
    
    if unmatched_count > 0:
        st.subheader('Items Not Found in System')
                
        not_found_list = left_only['item_no'].tolist()
        search_list = system['item_no'].unique().tolist()
        match_list = []

        for item in not_found_list:
            match = process.extract(item, search_list, limit=1)
            match = match[0][0] if match[0][1] >= 80 else ''

            if match == '':
                try:
                    matches = []
                    for word in item.split():
                        matched_items = system[system['item_no'].str.contains(word, regex=False, na=False)]['item_no'].unique().tolist()
                        matches.extend(matched_items)

                    if matches:
                        result = process.extract(item, matches, limit=1, scorer=fuzz.token_set_ratio)
                        match = result[0][0] if result else ''
                except:
                    match = ''

            match_list.append(match)

        possible_match = pd.DataFrame({
            'Not Found Item': not_found_list,
            'Possible Match in System': match_list
        })

        st.dataframe(possible_match)        
        st.warning(f"⚠️ {unmatched_count} item(s) cannot be found in the system.")
        st.warning('⚠️ Please ensure the "item_no" of Stock To Be Checked are identical to the System.')

        # Ask user to proceed with stock adjustment calculation if unmatched count > 0
        proceed = st.checkbox("Continue calculating stock adjustment with available data?")
    
    else:
        proceed = True
        
    # Drop merge indicator
    check = check.drop(columns=['_merge'])
    
    if proceed:
        check = check.dropna(subset=['stock_in_system'])  # Drop unmatched items
        check['stock_adjustment'] = check['stock_on_hand'] - check['stock_in_system']
        to_be_adjusted = check[check['stock_adjustment'] != 0].reset_index(drop=True)
    
        if not to_be_adjusted.empty:
            st.subheader("Stock to be Adjusted")
            st.dataframe(to_be_adjusted)
    
            # Download button (if not already included)
            from datetime import datetime
            today = datetime.today().strftime('%Y%m%d')
            csv = to_be_adjusted.to_csv(index=False).encode('utf-8')
            file_name = f"stock_adjustment_{today}.csv"
    
            st.download_button(
                label="📥 Download Stock Adjustment",
                data=csv,
                file_name=file_name,
                mime='text/csv'
            )
        else:
            st.success("✅ All matched stocks are accurate. No adjustment needed.")
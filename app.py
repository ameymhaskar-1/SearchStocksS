import streamlit as st
import io
from datetime import datetime
from stock_utils import fetch_stock_data, generate_excel

# Page Configuration
st.set_page_config(page_title="NSE Stock Analyzer", page_icon="📈", layout="wide")

st.title("📈 NSE Stock Analysis Tool")
st.markdown("""
Paste NSE company names below to generate a detailed momentum analysis report in Excel.
""")

# User Input
example_text = "Reliance, TCS, HDFC Bank\nInfosys\nZomato, Tata Motors"
input_data = st.text_area(
    "Enter Company Names (separated by commas or new lines):",
    placeholder=example_text,
    height=150
)

col1, col2 = st.columns([1, 5])
with col1:
    analyze_btn = st.button("🚀 Analyze Stocks")

if analyze_btn:
    if not input_data.strip():
        st.warning("Please enter at least one company name.")
    else:
        # Process input names
        names = []
        for line in input_data.split('\n'):
            for name in line.split(','):
                if name.strip():
                    names.append(name.strip())
        
        names = list(set(names)) # Remove duplicates
        
        with st.spinner(f'Fetching data for {len(names)} companies...'):
            df, failed = fetch_stock_data(names)
        
        if not df.empty:
            st.success(f"Successfully analyzed {len(df)} companies!")
            
            if failed:
                st.warning(f"Could not find data for: {', '.join(failed)}")
            
            # Display Preview
            st.subheader("Data Preview")
            
            # Updated .map() function for newer Pandas versions
            styled_df = df.style.map(
                lambda x: 'color: green' if isinstance(x, (int, float)) and x > 0 else ('color: red' if isinstance(x, (int, float)) and x < 0 else ''),
                subset=['1W Return %', '1M Return %', '3M Return %']
            )
            
            st.dataframe(styled_df, use_container_width=True)

            # Generate Excel
            output = io.BytesIO()
            workbook = generate_excel(df)
            workbook.save(output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel Report",
                data=processed_data,
                file_name=f"NSE_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Could not fetch data for any of the provided names. Please check the spellings.")

# Footer
st.markdown("---")
st.caption("Data provided by Yahoo Finance. This tool is for educational purposes.")

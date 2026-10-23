
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Excel Duplicate Remover",
    page_icon="📊"
)

st.title("📊 Excel Duplicate Remover")
st.write("Upload an Excel file and remove duplicate records.")

uploaded_file = st.file_uploader(
    "Choose your Excel file",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # Clean column names
        df.columns = df.columns.str.strip()

        if df.empty:
            st.warning("Your Excel file contains no data.")
        else:
            st.subheader("Original Data")
            st.write("Total rows:", len(df))

            # Updated: use width="stretch"
            st.dataframe(df, width="stretch")

            # Select a column
            column = st.selectbox(
                "Select the column to check for duplicates",
                options=df.columns.tolist()
            )

            if st.button("Remove Duplicates"):
                cleaned_df = df.drop_duplicates(
                    subset=[column],
                    keep="first"
                ).copy()

                st.subheader("Cleaned Data")
                st.dataframe(
                    cleaned_df,
                    width="stretch"
                )

                removed = len(df) - len(cleaned_df)

                st.success(
                    f"Removed {removed} duplicate rows."
                )

                # Create Excel file for download
                output = BytesIO()

                with pd.ExcelWriter(
                    output, engine="openpyxl"
                ) as writer:
                    cleaned_df.to_excel(
                        writer,
                        index=False
                    )

                st.download_button(
                    label="Download Cleaned Excel",
                    data=output.getvalue(),
                    file_name="student_cleaned.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

    except Exception as e:
        st.error(f"Something went wrong: {e}")

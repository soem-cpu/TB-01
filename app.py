import streamlit as st
import pandas as pd
import io

from rules.ssr11 import run_ssr11

st.image("assets/TB_image2.jpg", width=200)

st.title("📊 IHRP: TB data verification & indicator calculation")

st.markdown("""
Upload your **Excel file** for TB data verification.  
The app will apply built-in rules and show validation results.  
You can download results as Excel.
""")

# ---------------- Indicator Selection ----------------

indicator = st.selectbox(
    "📌 Select Indicator",
    [
        "SSR1.1",
        "SSR1.2",
        "SSR1.3",
        "SSR1.4",
        "SSR1.5",
        "SSR1.6",
        "SSR1.6K"
    ]
)

# ---------------- File Upload ----------------

data_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])


if data_file:

    try:

        if indicator == "SSR1.1":
            results = run_ssr11(data_file)

        else:
            st.warning(f"⚠️ {indicator} rules not implemented yet")
            st.stop()

        excel_output = io.BytesIO()

        with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:

            st.markdown("## 📑 Validation Results")

            for name, df in results.items():

                st.write(f"### {name}")

                if not df.empty:
                    st.dataframe(df, use_container_width=True)
                else:
                    st.success("No issues found ✅")

                df.to_excel(writer, sheet_name=name[:31], index=False)

        excel_output.seek(0)

        st.download_button(
            "⬇️ Download Excel Results",
            excel_output,
            file_name=f"{indicator}_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error running rules: {e}")

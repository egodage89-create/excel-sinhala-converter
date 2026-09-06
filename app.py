import streamlit as st
import pandas as pd
from deep_translator import GoogleTranslator
import io
import time

SINHALA_MONTHS = {
    1: "ජනවාරි", 2: "පෙබරවාරි", 3: "මාර්තු", 4: "අප්‍රේල්",
    5: "මැයි", 6: "ජූනි", 7: "ජූලි", 8: "අගෝස්තු",
    9: "සැප්තැම්බර්", 10: "ඔක්තෝබර්", 11: "නොවැම්බර්", 12: "දෙසැම්බර්"
}

def format_disconnection_date(date_val):
    try:
        dt = pd.to_datetime(date_val)
        return f"{dt.year} {SINHALA_MONTHS.get(dt.month, '')}"
    except Exception:
        return date_val

def translate_with_retry(text, retries=3, delay=1.2):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    
    text_str = str(text).strip()
    for attempt in range(retries):
        try:
            time.sleep(delay)
            translated = GoogleTranslator(source='auto', target='si').translate(text_str)
            if translated and "Error 500" not in str(translated):
                return translated
        except Exception:
            time.sleep(delay * 2) # Failure එකකදී තවත් තත්පර කිහිපයක් රඳී සිටී
            
    return text_str

st.title("📊 Excel Sinhala Converter Tool")
st.write("Singlish ලිපින සහ Disconnection Dates සිංහලට පරිවර්තනය කිරීම")

uploaded_file = st.file_uploader("ඔබගේ Excel File එක මෙතැනට Upload කරන්න", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("---")
    st.subheader("📋 Upload කළ Data (Preview)")
    st.dataframe(df.head())

    address_col = st.selectbox("ලිපිනය සහිත Column එක තෝරන්න:", df.columns)
    date_col = st.selectbox("Disconnection Date සහිත Column එක තෝරන්න:", df.columns)

    if st.button("Process & Convert Excel"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        translated_addresses = []
        total_rows = len(df)

        for idx, row in enumerate(df[address_col]):
            status_text.text(f"පරිවර්තනය වෙමින් පවතී: {idx + 1}/{total_rows}")
            translated_addresses.append(translate_with_retry(row))
            progress_bar.progress((idx + 1) / total_rows)

        df['Address_Sinhala'] = translated_addresses
        df['Disconnection_Date_Sinhala'] = df[date_col].apply(format_disconnection_date)

        status_text.text("පරිවර්තනය අවසන්!")
        st.success("සියලුම Data සාර්ථකව පරිවර්තනය විය!")
        st.dataframe(df.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Converted_Data')

        st.download_button(
            label="📥 Converted Excel Sheet එක Download කරගන්න",
            data=output.getvalue(),
            file_name="Converted_Data_Sinhala.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

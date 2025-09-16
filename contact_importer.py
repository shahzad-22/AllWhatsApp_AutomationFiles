import streamlit as st
import pandas as pd

st.set_page_config(page_title="Contact Import Automation", layout="centered")
st.title("📇 Contact Import Automation")

uploaded_file = st.file_uploader("📤 Upload Excel File (Name + Number columns)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    df = df[['Name', 'Number']].dropna()
    df['Phone'] = df['Number'].apply(lambda x: x.replace(" ", "").replace("-", "").strip())
    st.success("✅ Contacts loaded successfully!")
    st.dataframe(df)

    gmail = st.text_input("✉️ Enter your Gmail address (where contacts will be imported)")

    if gmail and st.button("📥 Generate Google Contacts CSV"):
        export_df = pd.DataFrame()
        export_df["Name"] = df["Name"]
        export_df["Phone 1 - Type"] = "Mobile"
        export_df["Phone 1 - Value"] = df["Phone"]

        csv_path = "contacts_google_upload.csv"
        export_df.to_csv(csv_path, index=False)

        with open(csv_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Google Contacts CSV",
                data=f,
                file_name="contacts_google_upload.csv",
                mime="text/csv"
            )

        st.info(
            f"Now go to [Google Contacts Import](https://contacts.google.com/import) and upload this file to your Gmail account: **{gmail}**"
        )

        st.markdown("---")
        st.success("🎯 Contacts prepared for import.")

        choice = st.radio("Do you want to start WhatsApp automation now?", ["✅ Yes", "❌ No"])

        if choice == "✅ Yes":
            st.markdown("🚀 Please run `whatsapp_sender.py` to continue messaging.")
        else:
            st.markdown("🛑 You chose not to proceed. You can run the sender script later.")
import subprocess

# After successful contact import
st.success("✅ Your contacts have been imported successfully.")

# Ask user if they want to send WhatsApp messages
send_now = st.button("📤 Start WhatsApp Messaging")

if send_now:
    st.info("Launching WhatsApp automation...")
    # Run the second script
    subprocess.Popen(["python", "whatsapp_sender.py"])

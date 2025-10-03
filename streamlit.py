import streamlit as st
import requests

# --- CONFIG ---
BACKEND_URL = "http://localhost:8502"   # Update if backend runs on another host/port

st.set_page_config(page_title="SES", layout="centered")

# --- SAP Connection ---
if st.button("Connect to SAP"):
    try:
        resp = requests.post(f"{BACKEND_URL}/connect-sap")
        if resp.status_code == 200:
            st.success("Connected")
        else:
            st.error(f"Failed to connect")
    except Exception as e:
        st.error(f"Could not reach backend")

# --- SES Form ---
with st.form("ses_form"):
    po_number = st.text_input("Enter Purchase Order Number:")
    distribution_name = st.text_input("Enter Distribution Name:")
    submitted = st.form_submit_button("Generate SES")

if submitted:
    if not po_number.strip():
        st.error("❌ Please enter a PO number.")
    else:
        try:
            with st.spinner("Generating SES document..."):
                resp = requests.post(
                    f"{BACKEND_URL}/generate-ses",
                    json={
                        "po_number": po_number,
                        "distribution_name": distribution_name,
                    },
                )

            if resp.status_code == 200:
                data = resp.json()
                st.success(f"SES generated for PO {data['po_number']} ✅")

                # Provide download link
                po_number = data["po_number"]
                download_url = f"{BACKEND_URL}/download/{po_number}"
                response = requests.get(download_url)
                st.download_button(
    label="📥 Download SES Document",
    data=response.content,
    file_name=f"SES_PO_{po_number}.docx",  # or .pdf etc.
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
            else:
                st.error(f"❌ Error: {resp.text}")
        except Exception as e:
            st.error(f"⚠️ Backend error: {e}")


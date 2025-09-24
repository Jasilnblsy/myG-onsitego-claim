import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ----------------------
# CONFIG
# ----------------------
EXCEL_FILE = "/workspaces/myG-onsitego-claim/MYG India Pvt Ltd report for AUGUST 2025_Revised (1).xlsx"
TARGET_EMAIL = "jmon11829@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "gamersplatform0@gmail.com"
SENDER_PASSWORD = "arbz roqo stto fxvk"  # Gmail App Password

# ----------------------
# LOAD DATA
# ----------------------
@st.cache_data
def load_data():
    return pd.read_excel(EXCEL_FILE)

df = load_data()

st.title("📌 OSG Warranty Claim Registration")

# ----------------------
# CUSTOMER SELECTION
# ----------------------
mobile_no_input = st.text_input("Enter Customer Mobile No")

if mobile_no_input:
    customer_data = df[df["Mobile No"].astype(str) == mobile_no_input.strip()]

    if not customer_data.empty:
        customer_name = customer_data["Customer"].iloc[0]

        st.subheader("Customer Details")
        st.write(f"**Customer Name:** {customer_name}")
        st.write(f"**Email:** {customer_data['Email'].iloc[0]}")
        st.write(f"**Mobile:** {mobile_no_input.strip()}")

        # Manual inputs
        customer_address = st.text_area("Enter Customer Address")
        issue_description = st.text_area("Describe the Issue")

        st.subheader("Purchased Products")
        customer_data["Product Display"] = (
            "Invoice: " + customer_data["Invoice No"].astype(str) +
            " | Model: " + customer_data["Model"].astype(str) +
            " | Serial/OSID: " + customer_data["Serial No"].astype(str)
        )

        product_choices = st.multiselect(
            "Select Product(s) for Claim",
            options=customer_data["Product Display"].tolist()
        )

        # File upload
        uploaded_file = st.file_uploader("Upload Invoice / Supporting Document", type=["pdf", "jpg", "png"])

        if st.button("Submit Claim"):
            if not product_choices:
                st.warning("Please select at least one product.")
            elif not customer_address.strip():
                st.warning("Please enter customer address.")
            elif not issue_description.strip():
                st.warning("Please describe the issue.")
            else:
                selected_products = customer_data[customer_data["Product Display"].isin(product_choices)]

                # ----------------------
                # EMAIL BODY (Professional, Product Details in Separate Rows)
                # ----------------------
                product_info = "\n\n".join([
                    f"Invoice  : {row['Invoice No']}\n"
                    f"Model    : {row['Model']}\n"
                    f"OSID     : {row['Serial No']}\n"
                    f"Plan     : {row['Plan Type']}\n"
                    f"Duration : {row['Duration (Year)']} Year(s)"
                    for _, row in selected_products.iterrows()
                ])

                body = f"""
Subject: Warranty Claim Submission - {customer_name}

Dear Team,

We have received a warranty claim with the following details:

----------------------------------------
Customer Information
----------------------------------------
Name       : {customer_name}
Mobile No  : {mobile_no_input.strip()}
Address    : {customer_address}

----------------------------------------
Product(s) Details
----------------------------------------
{product_info}

----------------------------------------
Issue Description
----------------------------------------
{issue_description}

Kindly review the claim and process it at the earliest convenience.
"""

                # ----------------------
                # SEND EMAIL
                # ----------------------
                msg = MIMEMultipart()
                msg["From"] = SENDER_EMAIL
                msg["To"] = TARGET_EMAIL
                msg["Subject"] = f"Warranty Claim - {customer_name}"
                msg.attach(MIMEText(body, "plain"))

                # Attach uploaded file if any
                if uploaded_file is not None:
                    file_attachment = MIMEApplication(uploaded_file.read(), Name=uploaded_file.name)
                    file_attachment['Content-Disposition'] = f'attachment; filename="{uploaded_file.name}"'
                    msg.attach(file_attachment)

                try:
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                        server.starttls()
                        server.login(SENDER_EMAIL, SENDER_PASSWORD)
                        server.sendmail(SENDER_EMAIL, TARGET_EMAIL, msg.as_string())

                    st.success("✅ Claim submitted successfully and sent to email.")
                except Exception as e:
                    st.error(f"❌ Error sending email: {e}")

    else:
        st.warning("No products found for this mobile number.")

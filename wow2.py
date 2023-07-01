import streamlit as st
from PIL import Image
import requests
from io import BytesIO
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from streamlit_option_menu import option_menu







response = requests.get('https://static.mycareersfuture.gov.sg/images/company/logos/eb6e0f752982dff2188b6cbc89eef734/mindsprint.png')
image = Image.open(BytesIO(response.content))


st.sidebar.image(image, caption='MINDSPRINT')

st.sidebar.write("")
st.sidebar.write("")

st.sidebar.success("# `This Web App built in Python and Streamlit by  MAYUKH BHAUMIK 👨🏼‍💻`")

col1, col2, col3= st.columns(3)

with col1:
    st.title("MINDSPRINT")
with col2:
    st.write()

with col3:
    st.image(image, caption='MINDSPRINT')

st.sidebar.write("")
st.sidebar.write("")

with st.sidebar:
        selected = option_menu("Menu", ["Home"], 
        icons=['house', ], default_index=0)

st.write("")
st.write("")
st.write("")
st.write("")
st.write("")
st.write("")



excel_file_1 = 'PAYMENT DATA.XLSX'
excel_file_2 = 'POC DATA.XLSX'
df_a = pd.read_excel(excel_file_1)
df_b = pd.read_excel(excel_file_2, sheet_name='B DATA')

date_format = '%d.%m.%Y'

dff= pd.DataFrame( columns=["Vendor","Vendor Name","Payment Number","Amount in USD","Payment Date","Account Number","Status"])

for i in range(0,len(df_a)) :
    print(i)
    row_a = df_a.iloc[i]
    df = df_b[df_b["Vendor"].isin([row_a["Vendor"]])]
    row_b = df.iloc[0]
    date_a = datetime.strptime(row_a["Posting Date"], date_format)
    date_b = datetime.strptime(row_b["Date"], date_format)

    if date_b == date_a :
        dff.loc[len(dff.index)] = [row_a["Vendor"],row_a["Vendor Name"],row_a["Payment Document Number"],row_a["Amount in loc.curr.2"],row_a["Posting Date"],row_a["Vendor Bank Ac"],"HAVE TO CHECK"]
            
    if date_b < date_a :
        dff.loc[len(dff.index)] = [row_a["Vendor"],row_a["Vendor Name"],row_a["Payment Document Number"],row_a["Amount in loc.curr.2"],row_a["Posting Date"],row_a["Vendor Bank Ac"],"ISSU"]

    if date_b > date_a :
        dff.loc[len(dff.index)] = [row_a["Vendor"],row_a["Vendor Name"],row_a["Payment Document Number"],row_a["Amount in loc.curr.2"],row_a["Posting Date"],row_a["Vendor Bank Ac"],"NOT ISSUS"]


st.sidebar.write("")
st.sidebar.write("")
st.sidebar.write("")

status_chek = st.sidebar.selectbox('SELECT Status :',["SELECT","HAVE TO CHECK","ISSU","NOT ISSUS"])

if status_chek != "SELECT" :
    dff = dff[dff["Status"].isin([status_chek])]

st.write(dff)



st.write("")
st.write("")
st.write("")
st.success("DATA is ready to Downlode or Sending MAIL.")
st.write("")


def download_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
        output.seek(0)
        return output

button = st.download_button(label='Download Excel', data=download_excel(dff), file_name='data_Paymant.xlsx')

def gmail() :

        dff.to_excel('data_Paymant.xlsx', index=False)

        try:
            s = smtplib.SMTP('smtp.gmail.com', 587)

            s.starttls()

            s.login("bhaumikmayukh@gmail.com", "qipywyfskjmofnky")

            sender_email = "bhaumikmayukh@gmail.com"
            recipient_email = "mayukh.bhaumik1999@gmail.com"
            cc_email = "bhaumikmayukh@gmail.com"
            subject = "Hello from Python!"
            body = "This is a test email sent from Python."
            attachment_path = "data_Paymant.xlsx"

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Cc'] = cc_email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(attachment_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {attachment_path}")
            msg.attach(part)

            s.sendmail(sender_email, [recipient_email, cc_email], msg.as_string())

            s.quit()
            st.success("Email sent successfully!")
        except Exception as e:
            st.warning("Error: Unable to send email.")
            st.warning(e)
            st.warning("Please Send again")

st.write("")
st.write("")
st.write("")
    

if st.button(" SEND MAIL "):
        gmail()






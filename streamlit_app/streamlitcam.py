import streamlit as st
import pyqrcode
from io import BytesIO
import cv2
from pyzbar import pyzbar
import pandas as pd
import os
import webbrowser

st.title("ElektroXen App")

def decode_qr_code(frame):
    decoded_objects = pyzbar.decode(frame)
    qr_codes = []
    for obj in decoded_objects:
        qr_code_data = obj.data.decode('utf-8')
        qr_codes.append((qr_code_data, obj.rect))
    return qr_codes

def create_qr_code():
    st.title("QR Code Generator")
    st.write("Enter text to generate a QR code.")

    input_text = st.text_input("Enter text:")
    if st.button('Generate QR Code'):
        if input_text:
            qr_code = pyqrcode.create(input_text)
            buffer = BytesIO()
            qr_code.png(buffer, scale=8)
            st.image(buffer.getvalue(), caption='Generated QR Code')
        else:
            st.error("Please enter text to generate a QR code.")

def save_to_excel(decoded_message, file_path):
    if file_path:
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            df = pd.DataFrame(columns=["QR Code Data", "S.No", "Component Location", "Confidence", "x1", "y1", "x2", "y2", "Actual Results", "Prediction Accuracy"])  # Define columns here

            # Add a note about column appearance
            st.info("Column names are shown in grey color.")

        # Create a new row with the provided decoded_message
        new_row = pd.DataFrame({
            "QR Code Data": [decoded_message],
            "S.No": [""],
            "Component Location": [""],
            "Confidence": [""],
            "x1": [""],
            "y1": [""],
            "x2": [""],
            "y2": [""],
            "Actual Results": [""],
            "Prediction Accuracy": [""]
        })

        # Concatenate the new row to the existing DataFrame
        df = pd.concat([df, new_row], ignore_index=True)

        # Save the DataFrame to Excel file
        df.to_excel(file_path, index=False)

def scan_qr_code():
    st.title("QR Code Scanner")
    st.write("Click the button below to start the camera and scan a QR code.")

    if "scanning" not in st.session_state:
        st.session_state.scanning = False

    if "decoded_messages" not in st.session_state:
        st.session_state.decoded_messages = []

    file_path = st.text_input("Enter Excel file path:")

    if st.button('Start Scanning') or st.session_state.scanning:
        st.session_state.scanning = True
        cap = cv2.VideoCapture(1, cv2.CAP_DSHOW)  # Use CAP_DSHOW backend for potentially faster initialization

        if not cap.isOpened():
            st.error("Failed to open webcam. Please ensure that the webcam is connected and not being used by other applications.")
            return

        stframe = st.empty()

        try:
            while st.session_state.scanning:
                ret, frame = cap.read()

                if not ret:
                    st.write("Failed to capture image.")
                    break

                qr_codes = decode_qr_code(frame)

                for qr_code, rect in qr_codes:
                    x, y, w, h = rect.left, rect.top, rect.width, rect.height 
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

                stframe.image(frame, channels="BGR")

                if qr_codes:
                    decoded_message = qr_codes[0][0]
                    if decoded_message not in st.session_state.decoded_messages:
                        st.session_state.decoded_messages.append(decoded_message)
                        st.success(f"Decoded QR Code: {decoded_message}")
                        save_to_excel(decoded_message, file_path)  # Save decoded message to Excel
                        st.session_state.scanning = False
                        cap.release()
                        webbrowser.open_new_tab("http://localhost:8501")  # Redirect to the URL
                        st.experimental_rerun()  # Rerun the app
                        return

        except Exception as e:
            st.error(f"Error during object detection: {e}")

        finally:
            cap.release()

def main():
    st.sidebar.title("QR Code Tool")
    menu = ["Generate QR Code", "Scan QR Code"]
    choice = st.sidebar.selectbox("Select Activity", menu)

    if choice == "Generate QR Code":
        create_qr_code()
    elif choice == "Scan QR Code":
        scan_qr_code()

if __name__ == "__main__":
    main()


#Robot Motion Control

import os
import sys
import cv2
from ultralytics import YOLO
import streamlit as st
import altair as alt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import base64

st.title("ElektroXen App")

# Get the path to the model file
def get_model_path():
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, use this path for the model file
        return os.path.join(sys._MEIPASS, 'streamlit_app/best_F3.pt')
    else:
        # If running in a normal environment, use the usual path
        return os.path.join(os.getcwd(), 'best_F3.pt')

model_path = get_model_path()

# Debugging: Print current working directory and list of files
st.write(f"Current working directory: {os.getcwd()}")
st.write(f"Files in current directory: {os.listdir(os.getcwd())}")
st.write(f"Model path: {model_path}")  # Debug statement to check the model path

# Check if the model file exists
if not os.path.exists(model_path):
    st.error(f"Model file not found at {model_path}")
else:
    # Load the YOLO model
    try:
        model = YOLO(model_path)
        st.success("Model loaded successfully.")
    except Exception as e:
        st.error(f"Error loading YOLO model: {e}")

classNames = ["Capacitor", "Diode", "Dot-Cut Mark", "Excess-Solder", "IC", "MCU", "Missing Com.", "Non-Good com.",
              "Resistor", "Short", "Soldering-Missing", "Tilt-Com"]

# Check if the webcam is opened correctly
cap = cv2.VideoCapture(0)
if not cap.isOpened():
    st.error("Error: Couldn't open webcam.")
else:
    st.success("Webcam is opened successfully.")

# Create a Streamlit placeholder to display the detected image
detected_image_placeholder = st.empty()

# Create a multiselect widget for selecting labels
selected_labels = st.sidebar.multiselect("Select Labels", classNames, default=classNames)

# Create a placeholder to display bounding box predictions in the sidebar
bounding_box_placeholder = st.sidebar.empty()

# Upload XLSX file
uploaded_file = st.file_uploader("Upload XLSX file", type=["xlsx"])

# Initialize the workbook and worksheet for storing results
filename = "results.xlsx"
try:
    if uploaded_file is not None:
        wb = load_workbook(uploaded_file)
        ws = wb.active
    else:
        wb = load_workbook(filename) if os.path.exists(filename) else Workbook()
        ws = wb.active
        # Add the custom column headers if the worksheet is empty
        if ws.max_row == 1:
            ws.append(["Label", "S.No", "Confidence", "x1", "y1", "x2", "y2", "Actual Results", "Prediction Accuracy"])
            
            # Apply bold font to headers
            for cell in ws[1]:
                cell.font = Font(bold=True)
except Exception as e:
    st.error(f"Error accessing workbook: {e}")

# Create download button for Excel file
if st.button("Download Report results.xlsx"):
    try:
        wb.save(filename)
        with open(filename, "rb") as file:
            btn = file.read()
            b64 = base64.b64encode(btn).decode()
            href = f'<a href="data:file/xlsx;base64,{b64}" download="{filename}">Download {filename}</a>'
            st.markdown(href, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error saving workbook: {e}")

# Initialize S.No
row_number = ws.max_row

# Set to store unique bounding box predictions
unique_predictions = set()

while cap.isOpened():
    # Read the frame from the webcam
    ret, frame = cap.read()

    if not ret:
        st.error("Error: Couldn't read frame.")
        break

    # Perform object detection
    try:
        result_img, bounding_box_predictions = predict_and_detect(model, frame, classes=selected_labels, conf=0.5)

        # Save bounding box predictions to Excel file
        for prediction in bounding_box_predictions:
            # Convert tensor values to Python native types (float or int)
            confidence = float(prediction["Confidence"]) * 100
            confidence_str = f"{confidence:.2f}%"
            x1 = int(prediction["x1"])
            y1 = int(prediction["y1"])
            x2 = int(prediction["x2"])
            y2 = int(prediction["y2"])
            Actual_Results = "UNKNOWN"
            
            # Determine success rate based on class
            if prediction["Label"] in ["Capacitor", "Diode", "IC", "MCU", "Dot-Cut Mark", "Resistor"]:
                Actual_Results = "OK"
                cell_color = "00FF00"  # Green color in HEX format
            elif prediction["Label"] in ["Excess-Solder", "Missing Com.", "Non-Good com.", "Short", "Soldering-Missing", "Tilt-Com"]:
                Actual_Results = "FAIL"
                cell_color = "FF0000"  # Red color in HEX format
            if 50 <= confidence < 90:
                Actual_Results = "NOT OK"
                cell_color = "FFFF00"  # Yellow color in HEX format
 
            # Determine Prediction Accuracy
            if 85 <= confidence <= 100:
                Prediction_Accuracy = "PASS"
                accuracy_color = "00FF00"  # Green color in HEX format
            elif 50 <= confidence < 85:
                Prediction_Accuracy = "FAIL"
                accuracy_color = "FF0000"  # Red color in HEX format
            else:
                Prediction_Accuracy = "UNKNOWN"
                accuracy_color = "FFFFFF"  # White color in HEX format

            # Create a tuple of the prediction to check for duplicates
            prediction_tuple = (prediction["Label"], confidence, x1, y1, x2, y2)
            
            # Only add unique predictions
            if prediction_tuple not in unique_predictions:
                unique_predictions.add(prediction_tuple)
                # Append the bounding box predictions to the Excel sheet
                ws.append([prediction["Label"], row_number, confidence_str, x1, y1, x2, y2, Actual_Results, Prediction_Accuracy])
                
                # Apply color formatting to the "success_rate" column
                row = ws.max_row
                ws.cell(row=row, column=8).fill = PatternFill(start_color=cell_color, end_color=cell_color, fill_type="solid")

                # Apply color formatting to the "Prediction Accuracy" column
                ws.cell(row=row, column=9).fill = PatternFill(start_color=accuracy_color, end_color=accuracy_color, fill_type="solid")

                # Apply red color formatting to the "Label" column if the label is in the specified list
                if prediction["Label"] in ["Excess-Solder", "Missing Com.", "Non-Good com.", "Short", "Soldering-Missing", "Tilt-Com"]:
                    ws.cell(row=row, column=1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

                row_number += 1  # Increment row number

        # Save the changes to the Excel file
        wb.save(filename)

        # Display the detected image
        detected_image_placeholder.image(result_img, channels="BGR", caption='Detected Objects', use_column_width=True)

        # Update the bounding box predictions in the sidebar                    
        bounding_box_placeholder.text("Bounding Box Predictions:")
        for prediction in bounding_box_predictions:
            bounding_box_placeholder.text(f"Class: {prediction['Label']}, Bounding Box: ({prediction['x1']}, {prediction['y1']}) - ({prediction['x2']}, {prediction['y2']})")

    except Exception as e:
        st.error(f"Error during object detection: {e}")

# Release the VideoCapture and close all OpenCV windows
cap.release()
# cv2.destroyAllWindows()

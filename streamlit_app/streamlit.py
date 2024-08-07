import cv2

cap = cv2.VideoCapture(0)

if not cap.isOpened():
    error_code = cap.get(cv2.CAP_PROP_SETTINGS)
    st.error(f"Error: Couldn't open webcam. Error code: {error_code}")
    logger.error(f"Error: Couldn't open webcam. Error code: {error_code}")
else:
    st.success("Webcam is opened successfully.")
    logger.info("Webcam opened successfully.")

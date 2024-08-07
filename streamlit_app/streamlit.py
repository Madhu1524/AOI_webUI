import streamlit as st
import cv2

def main():
    st.title("Live Webcam Feed")

    # Try to open the webcam
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        st.error(f"Error: Could not open webcam. Check the camera index and connection.")
        st.stop()  # Stop the Streamlit app if camera cannot be accessed

    frame_placeholder = st.empty()

    while True:
        ret, frame = cap.read()
        if not ret:
            st.error("Error: Failed to capture image.")
            break

        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame_placeholder.image(frame_rgb)

        if st.button("Stop"):
            break

    cap.release()

if __name__ == "__main__":
    main()

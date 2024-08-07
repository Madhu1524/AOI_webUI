import streamlit as st
import cv2

def main():
    st.title("Live Webcam Feed")

    # Open the webcam
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        st.error("Error: Could not open webcam.")
        return

    # Create a placeholder for the webcam feed
    frame_placeholder = st.empty()

    while True:
        ret, frame = cap.read()
        if not ret:
            st.error("Error: Failed to capture image.")
            break

        # Convert the frame to RGB (Streamlit requires RGB format)
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # Display the frame
        frame_placeholder.image(frame_rgb)

        if st.button("Stop"):
            break

    # Release the camera when done
    cap.release()

if __name__ == "__main__":
    main()

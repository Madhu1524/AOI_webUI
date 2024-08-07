import streamlit as st

st.title("Webcam Streamlit App")

# Use Streamlit's built-in camera input widget
image = st.camera_input("Take a picture")

if image:
    st.image(image, caption="Captured Image", use_column_width=True)

import streamlit as st
from streamlit_webrtc import webrtc_streamer

def main():
    st.title("Streamlit WebRTC Test")
    webrtc_streamer(key="test")

if __name__ == "__main__":
    main()

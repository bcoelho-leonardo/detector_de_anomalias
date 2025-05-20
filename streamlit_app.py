# -*- coding: utf-8 -*-
import streamlit as st

# set_page_config MUST be the first Streamlit command
st.set_page_config(page_title="Test App", layout="centered")

# Basic content
st.title("Test App")
st.write("Hello World! If you can see this, the app is working.")

# Simple interaction
if st.button("Click Me"):
    st.success("Button clicked!")
import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from utils import report, create_cmap, dict_langs2

st.title("Template app")

with st.sidebar:
    st.header("Upload file")
    uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file, sep=';', usecols=["Language",
                                                            "Profession",
                                                            "Student's Email",
                                                            "Reviewer's Email",
                                                            "Course",
                                                            "Lesson",
                                                           "Attempt number",
                                                            "Attempt status",
                                                           "Current Status",
                                                            "Upload Date",
                                                            "Review Start",
                                                            "Review Finish",
                                                            "Ticket ID"]).dropna()
    st.dataframe(df)

#download file
try:
    cmap, cmap2 = create_cmap("#025464", "#E57C23", "#E8AA42", "#F8F1F1")

    binary_output = report(df)



    st.download_button(label = 'Download ppw',
                    data = binary_output.getvalue(),
                    file_name = 'my_power.pptx')
    
except NameError:
    st.warning("Load data from Iterations info in doublecloud")


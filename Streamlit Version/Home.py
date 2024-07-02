import streamlit as st


def main():
    st.set_page_config(
        page_title="Transcript Converter",
        page_icon="📂",
        layout="centered",
        initial_sidebar_state="auto",
    )
    st.title("Transcript Converter")
    st.image("logo.png", width=200)  # Replace "logo.png" with your logo file

    st.markdown(
        """
    ## Welcome to the Transcript Converter
    This website helps you convert bilingual one-on-one scripted vocabulary lesson transcripts into Excel sheets.
    
    ### How it works:
    1. **Stage 1:** Upload a transcript file and convert it to a raw Excel file.
    2. **Stage 2:** Upload the raw Excel file and apply a template to it for further analysis and coding.
    
    ### Instructions:
    - Make sure your transcript file follows the specified format.
    - Ensure your template file has the required columns.
    
    Navigate through the stages using the sidebar.
    """
    )


if __name__ == "__main__":
    main()

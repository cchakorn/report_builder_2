import streamlit as st

def main_page():
    st.markdown("# Main Page 🎈")
    st.sidebar.markdown("# Main page 🎈")

def page2():
    st.markdown("# REPORT Cleaner ❄️")
    st.sidebar.markdown("# REPORT Builder ❄️")

def page2():
    st.markdown("# REPORT Builder ❄️")
    st.sidebar.markdown("# REPORT Builder ❄️")


page_names_to_funcs = {
    "Main Page": main_page,
    "REPORT Builder": page2,
}

st.title('REPORT WEB APP')
st.header('Select application in sidebar:')
st.subheader('Current Options Are:')
st.write('page 2 : Report Cleaner')
st.write('page 3 : Report Builder')




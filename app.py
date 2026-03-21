import streamlit as st
import os
import time
import shutil
from agent import SurcotecAgent
from tools import process_file_to_text

# ==========================================
# 1. CONFIGURATION & UI SETUP
# ==========================================
# ⚠️ Ensure your key is placed here
MY_API_KEY = ""
OUTPUT_DIR = "generated_docs"
EXAMPLES_DIR = "Examples"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

st.set_page_config(
    page_title="Surcotec Architect",
    page_icon="🏗️",
    layout="wide"
)

# Custom CSS for Professional Industrial Branding
st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; color: #111111; }
    .stApp p, .stApp label, .stApp span, .stApp div,
    .stApp .stMarkdown, .stApp .stText { color: #111111; }
    h1, h2, h3, h4 { color: #001f3f !important; font-family: 'Segoe UI', sans-serif; font-weight: 700; }
    [data-testid="stSidebar"] { background-color: #001f3f; }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    [data-testid="stSidebar"] .stFileUploader label { color: #ffffff !important; }
    .stButton>button {
        background-color: #001f3f;
        color: white !important;
        border-radius: 6px;
        width: 100%;
    }
    .stDownloadButton>button {
        background-color: #218838 !important;
        color: white !important;
        font-weight: bold;
        border-radius: 6px;
        width: 100%;
    }
    .stChatMessage {
        background-color: #ffffff;
        border: 1px solid #d1d9e0;
        border-radius: 12px;
    }
    .stChatMessage p, .stChatMessage div, .stChatMessage span { color: #111111 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- Initialize Session States ---
if "agent" not in st.session_state:
    st.session_state.agent = SurcotecAgent(MY_API_KEY)

if "messages" not in st.session_state:
    st.session_state.messages = []

# ==========================================
# 2. SIDEBAR - CONTROL PANEL
# ==========================================
with st.sidebar:
    st.title("🏗️ Surcotec Ops")
    st.caption("Intelligent Document Control")
    st.divider()

    if os.path.exists(EXAMPLES_DIR) and len(os.listdir(EXAMPLES_DIR)) > 0:
        st.write("✅ **Learning Examples Active**")
    else:
        st.warning("⚠️ No Examples folder found")

    st.subheader("📁 1. Upload Quote")
    uploaded_file = st.file_uploader("Drop PDF/Excel quote here", type=['pdf', 'xlsx', 'xls', 'csv'])

    if uploaded_file:
        st.success(f"File Ready: {uploaded_file.name}")

        if st.button("🔍 2. Analyze & Compare"):
            with st.spinner("Analyzing..."):
                file_bytes = uploaded_file.getvalue()
                extracted_text = process_file_to_text(file_bytes, uploaded_file.name)

                prompt = (
                    f"I have uploaded a new quotation: {uploaded_file.name}. "
                    f"Content: \n{extracted_text}\n\n"
                    "Step 1: Compare this to the Master Template and Examples. "
                    "Step 2: List the changes clearly for me to review."
                )
                response = st.session_state.agent.ask(prompt)
                st.session_state.messages.append({"role": "assistant", "content": response})
                st.rerun()

    st.divider()
    st.subheader("📥 3. Download Results")

    if os.path.exists(OUTPUT_DIR):
        files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.xlsx')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_DIR, x)), reverse=True)

        if files:
            # Latest file — always front and centre
            latest = files[0]
            latest_path = os.path.join(OUTPUT_DIR, latest)
            with open(latest_path, "rb") as f_ptr:
                st.download_button(
                    label="📊 Download Latest",
                    data=f_ptr.read(),
                    file_name=latest,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_latest_{os.path.getmtime(latest_path)}",
                    use_container_width=True,
                )

            # Older files — hidden in an expander with a selector
            if len(files) > 1:
                with st.expander("🗂️ Download older files"):
                    selected = st.selectbox(
                        "Select a file",
                        options=files[1:],
                        label_visibility="collapsed",
                    )
                    selected_path = os.path.join(OUTPUT_DIR, selected)
                    with open(selected_path, "rb") as f_ptr:
                        st.download_button(
                            label=f"⬇️ Download: {selected[:22]}",
                            data=f_ptr.read(),
                            file_name=selected,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_old_{selected}_{os.path.getmtime(selected_path)}",
                            use_container_width=True,
                        )

            if st.button("🚨 Clear File History", use_container_width=True):
                for f in files:
                    os.remove(os.path.join(OUTPUT_DIR, f))
                st.toast("Cleared!", icon="🔥")
                time.sleep(1)
                st.rerun()
        else:
            st.info("No files ready.")

    st.divider()
    if st.button("🗑️ Clear Chat", use_container_width=True):
        st.session_state.messages = []
        st.rerun()

# ==========================================
# 3. MAIN CHAT INTERFACE
# ==========================================
st.title("Surcotec Document Architect")
st.markdown("---")

col_chat, col_status = st.columns([3, 1])

with col_chat:
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_query := st.chat_input("Ask for corrections or type 'Produce' to save..."):
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)

        with st.chat_message("assistant"):
            with st.spinner("Processing..."):
                # INSTRUCTION TO AGENT: We remind it to incorporate CHAT HISTORY into the final save
                response = st.session_state.agent.ask(user_query)
                st.markdown(response)
                st.session_state.messages.append({"role": "assistant", "content": response})

                if "SUCCESS" in response:
                    st.toast("Excel Generated!", icon="📊")
                    time.sleep(1.5)
                    st.rerun()

with col_status:
    st.subheader("System Intel")
    st.info("🤖 **Engine:** Gemini 2.5 Flash")
    st.success("📝 **Template:** Active")

    with st.expander("Applying Corrections"):
        st.write("""
        If you see a mistake in the preview:
        1. Tell the bot: *"Change X to Y"*
        2. Wait for confirmation.
        3. Then say **'Produce'**.
        The bot will now include your chat edits in the final Excel file.
        """)
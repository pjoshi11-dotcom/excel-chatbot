import streamlit as st
import pandas as pd

# Load all sheets from the uploaded Excel file
@st.cache_data
def load_qa_pairs(uploaded_file):
    qa_list = []
    xls = pd.ExcelFile(uploaded_file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        # Try to find columns that look like questions/answers
        q_cols = [col for col in df.columns if 'question' in col.lower()]
        a_cols = [col for col in df.columns if 'answer' in col.lower() or 'reply' in col.lower() or 'expected response' in col.lower()]
        if q_cols and a_cols:
            for _, row in df.iterrows():
                question = str(row[q_cols[0]])
                answer = str(row[a_cols[0]])
                if question and answer and question != 'nan' and answer != 'nan':
                    qa_list.append({'sheet': sheet_name, 'question': question, 'answer': answer})
    return qa_list

st.title("📊 Excel Knowledge Base Chatbot")

uploaded_file = st.file_uploader("Upload your Excel knowledge base", type=["xlsx"])
if uploaded_file:
    qa_pairs = load_qa_pairs(uploaded_file)
    st.success(f"Loaded {len(qa_pairs)} Q&A pairs from {len(set([q['sheet'] for q in qa_pairs]))} sheets.")

    # Chat interface
    if "history" not in st.session_state:
        st.session_state.history = []

    user_input = st.chat_input("Ask your question...")
    if user_input:
        # Simple keyword matching for best answer
        best_match = None
        best_score = 0
        for qa in qa_pairs:
            score = sum(1 for word in user_input.lower().split() if word in qa['question'].lower())
            if score > best_score:
                best_score = score
                best_match = qa
        if best_match and best_score > 0:
            answer = f"**[{best_match['sheet']}]** {best_match['answer']}"
        else:
            answer = "Sorry, I couldn't find an answer to your question in the knowledge base."
        st.session_state.history.append(("user", user_input))
        st.session_state.history.append(("bot", answer))

    # Display chat history
    for role, msg in st.session_state.history:
        st.chat_message(role).markdown(msg)
else:
    st.info("Please upload your Excel knowledge base to start chatting.")
"""
Copilot Studio Chat Interface
A Streamlit chat app using the M365 Agents SDK.
"""

import asyncio
import os
import streamlit as st
from streamlit_msal import Msal
from dotenv import load_dotenv

from copilot_client import CopilotStudioClient, clean_citations, format_references_html

load_dotenv()

# Page config
st.set_page_config(
    page_title="Copilot Studio",
    page_icon="💬",
    layout="centered",
)

# Minimal styling
st.markdown("""
<style>
    header {visibility: hidden;}
    .block-container {padding-top: 2rem;}
</style>
""", unsafe_allow_html=True)


def init_session():
    """Initialize session state."""
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "client" not in st.session_state:
        st.session_state.client = None


def main():
    init_session()

    st.title("💬 Copilot Studio")

    # Authentication via streamlit-msal
    auth_data = Msal.initialize_ui(
        client_id=os.getenv("AZURE_APP_CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}",
        scopes=["https://api.powerplatform.com/.default"],
        sign_in_label="Sign in with Microsoft",
        sign_out_label="Sign out",
    )

    if not auth_data:
        st.info("Please sign in to chat with Copilot Studio.")
        st.stop()

    # New Chat button in header area
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("🗑️ New", help="Start a new conversation"):
            st.session_state.messages = []
            st.session_state.client = None
            st.rerun()

    # Get access token
    access_token = auth_data.get("accessToken")
    if not access_token:
        st.error("Failed to get access token.")
        st.stop()

    # Initialize client if needed
    if st.session_state.client is None:
        with st.spinner("Connecting to Copilot Studio..."):
            client = CopilotStudioClient(access_token)
            welcome = asyncio.run(client.start_conversation())
            st.session_state.client = client

            if welcome:
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": welcome
                })

    # Display messages
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            # Use unsafe_allow_html for assistant messages (may contain citation links)
            if msg["role"] == "assistant":
                st.markdown(msg["content"], unsafe_allow_html=True)
            else:
                st.markdown(msg["content"])

    # Chat input
    if prompt := st.chat_input("Message Copilot..."):
        # Add user message
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Stream response
        with st.chat_message("assistant"):
            # Container for reasoning/thoughts (collapsible)
            thinking_container = st.empty()
            status_placeholder = st.empty()
            content_placeholder = st.empty()

            async def process_response():
                content_parts = []
                suggestions = None
                citation_metadata = {}
                search_results = []  # Collect search results by index
                thoughts = []  # Collect chain-of-thought
                got_streaming = False

                async for msg_type, msg_content in st.session_state.client.send_message(prompt):
                    if msg_type == 'status':
                        status_placeholder.caption(f"_{msg_content}_")
                    elif msg_type == 'thought':
                        # Collect reasoning/chain-of-thought
                        thoughts.append(msg_content)
                        # Update thinking display
                        with thinking_container.status("Thinking...", expanded=False) as status:
                            for t in thoughts:
                                task = t.get('task', 'Processing')
                                text = t.get('text', '')
                                st.write(f"**{task}**: {text}")
                    elif msg_type == 'search_result':
                        # Collect search results (contain URLs)
                        search_results.append(msg_content)
                    elif msg_type == 'content':
                        got_streaming = True
                        content_parts.append(msg_content)
                        # Show accumulated content with citations cleaned (plain text during streaming)
                        accumulated = "".join(content_parts)
                        cleaned, _ = clean_citations(accumulated)
                        content_placeholder.markdown(cleaned)
                    elif msg_type == 'final_content':
                        # Non-streaming response - use this only if we didn't get streaming chunks
                        if not got_streaming:
                            content_parts = [msg_content]
                            cleaned, _ = clean_citations(msg_content)
                            content_placeholder.markdown(cleaned)
                    elif msg_type == 'citations':
                        # Merge citation metadata from entities
                        # Try to enrich with URLs from search results
                        for cite_id, cite_info in msg_content.items():
                            # Citation IDs are like 'turn52search0' - extract index
                            import re
                            match = re.search(r'search(\d+)$', cite_id)
                            if match and not cite_info.get('url'):
                                idx = int(match.group(1))
                                # Find matching search result by index
                                for sr in search_results:
                                    if sr.get('index') == idx:
                                        cite_info['url'] = sr.get('url', '')
                                        if not cite_info.get('title'):
                                            cite_info['title'] = sr.get('title', '')
                                        break
                        citation_metadata.update(msg_content)
                    elif msg_type == 'suggestion':
                        suggestions = msg_content

                # Finalize thinking display
                if thoughts:
                    with thinking_container.status("Reasoning", expanded=False, state="complete") as status:
                        for t in thoughts:
                            task = t.get('task', 'Processing')
                            text = t.get('text', '')
                            st.write(f"**{task}**: {text}")

                # Clear status when done
                status_placeholder.empty()

                # Return cleaned final content with clickable HTML citations
                raw_content = "".join(content_parts)
                cleaned_text, citations = clean_citations(raw_content, use_html=True, citation_metadata=citation_metadata)

                # Add references section with clickable links
                if citations:
                    cleaned_text += format_references_html(citations)

                return cleaned_text, citations, suggestions

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                response, citations, suggestions = loop.run_until_complete(process_response())
            finally:
                loop.close()

            # Render final response (with clickable HTML citations if any)
            content_placeholder.markdown(response, unsafe_allow_html=True)

            # Show suggestions if any
            if suggestions:
                st.caption(f"**Suggestions:** {suggestions}")

        # Store response with HTML citations for history
        st.session_state.messages.append({"role": "assistant", "content": response})


if __name__ == "__main__":
    main()

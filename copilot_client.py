"""
Copilot Studio Client
Simple wrapper for the M365 Agents SDK.
"""

import os
import re
from typing import AsyncIterator, Optional

from microsoft_agents.activity import ActivityTypes
from microsoft_agents.copilotstudio.client import ConnectionSettings, CopilotClient


def clean_citations(text: str, use_html: bool = False, citation_metadata: dict = None) -> tuple[str, dict]:
    """
    Clean up citation markers from Copilot Studio responses.
    Citations use Unicode markers: \ue200cite\ue202{id}\ue201

    Args:
        text: The text containing citation markers
        use_html: If True, return clickable HTML links to external sources
        citation_metadata: Dict mapping citation IDs to {url, title}

    Returns:
        tuple: (cleaned_text, citations_dict) where citations_dict maps numbers to {id, url, title}
    """
    if not text:
        return text, {}

    citation_metadata = citation_metadata or {}

    # Track unique citations to number them
    citations = {}
    citation_counter = [0]

    def replace_citation(match):
        citation_id = match.group(1)
        if citation_id not in citations:
            citation_counter[0] += 1
            num = citation_counter[0]
            meta = citation_metadata.get(citation_id, {})
            citations[citation_id] = {
                'num': num,
                'url': meta.get('url', ''),
                'title': meta.get('title', f'Source {num}')
            }

        cite_info = citations[citation_id]
        num = cite_info['num']
        url = cite_info['url']

        if use_html and url:
            # Clickable superscript that opens external URL
            return f'<a href="{url}" target="_blank" style="text-decoration:none;color:#0066cc;"><sup>[{num}]</sup></a>'
        else:
            # Plain text for streaming display
            return f"[{num}]"

    # Pattern: \ue200cite\ue202{citation_id}\ue201 (using non-greedy match)
    cleaned = re.sub(r'\ue200cite\ue202(.+?)\ue201', replace_citation, text)

    # Return dict mapping numbers to citation info
    citations_by_num = {v['num']: {'id': k, 'url': v['url'], 'title': v['title']}
                        for k, v in citations.items()}

    return cleaned, citations_by_num


def format_references_html(citations: dict) -> str:
    """Format citations as an HTML references section with clickable links."""
    if not citations:
        return ""

    html = '<div style="margin-top:1rem;padding-top:0.5rem;border-top:1px solid #ddd;font-size:0.9em;">'
    html += '<strong>References:</strong><br>'
    for num in sorted(citations.keys()):
        cite = citations[num]
        title = cite.get('title', f'Source {num}')
        url = cite.get('url', '')
        if url:
            html += f'<a href="{url}" target="_blank" style="color:#0066cc;">[{num}] {title}</a><br>'
        else:
            html += f'<span>[{num}] {title}</span><br>'
    html += '</div>'
    return html


class CopilotStudioClient:
    """Wrapper for Copilot Studio interactions."""

    def __init__(self, access_token: str):
        """Initialize with an access token from MSAL."""
        self.settings = ConnectionSettings(
            environment_id=os.getenv("COPILOT_ENVIRONMENT_ID", ""),
            agent_identifier=os.getenv("COPILOT_AGENT_IDENTIFIER", ""),
            cloud=None,
            copilot_agent_type=None,
            custom_power_platform_cloud=None,
        )
        self._client = CopilotClient(self.settings, access_token)
        self._conversation_id: Optional[str] = None

    async def start_conversation(self) -> Optional[str]:
        """Start a new conversation and return welcome message."""
        welcome_text = ""
        async for activity in self._client.start_conversation():
            if activity.text:
                welcome_text += activity.text + "\n"
            if activity.conversation:
                self._conversation_id = activity.conversation.id

        return welcome_text.strip() if welcome_text else None

    async def send_message(self, message: str) -> AsyncIterator[tuple[str, any]]:
        """
        Send a message and yield response tuples.

        Yields:
            tuple[str, any]: (type, content) where type is:
                - 'status': Informative messages like "Generating plan..."
                - 'content': Actual response content chunks
                - 'suggestion': Suggested actions
                - 'citations': Dict mapping citation IDs to metadata (url, title)
        """
        if not self._conversation_id:
            yield ('content', "Error: No active conversation.")
            return

        import json
        debug_activities = []

        async for reply in self._client.ask_question(message, self._conversation_id):
            channel_data = getattr(reply, 'channel_data', None) or getattr(reply, 'channelData', {}) or {}

            # Capture full activity for debugging - properly serialize entities
            entities_data = []
            raw_entities = getattr(reply, 'entities', None) or []
            for ent in raw_entities:
                if hasattr(ent, '__dict__'):
                    entities_data.append(vars(ent))
                elif isinstance(ent, dict):
                    entities_data.append(ent)
                else:
                    entities_data.append(str(ent))

            activity_debug = {
                'type': str(reply.type),
                'text': reply.text[:200] if reply.text else None,
                'channel_data': channel_data,
                'entities': entities_data,
                'attachments': getattr(reply, 'attachments', None),
                'value': getattr(reply, 'value', None),
            }
            debug_activities.append(activity_debug)
            with open('/tmp/activities_debug.json', 'w') as f:
                json.dump(debug_activities, f, indent=2, default=str)

            # Capture chain-of-thought and search results from event activities
            if reply.type == ActivityTypes.event:
                value = getattr(reply, 'value', None) or {}
                if isinstance(value, dict):
                    # Chain of thought / reasoning
                    thought = value.get('thought')
                    if thought:
                        task_id = value.get('taskDialogId', '')
                        state = value.get('state', '')
                        # Extract just the tool name from identifiers like:
                        # MCcr981_guildhallAssistant.action.GuildhallMCPServer-InvokeServer:list_quests
                        # P:UniversalSearchTool
                        task_name = task_id
                        if ':' in task_id:
                            task_name = task_id.split(':')[-1]  # Get part after last colon
                        elif '.' in task_id:
                            task_name = task_id.split('.')[-1]  # Get part after last dot
                        task_name = task_name.replace('P:', '').replace('-InvokeServer', '')
                        yield ('thought', {
                            'text': thought,
                            'task': task_name,
                            'state': state
                        })

                    # Search results (contain URLs)
                    observation = value.get('observation', {})
                    if isinstance(observation, dict):
                        search_result = observation.get('search_result', {})
                        if isinstance(search_result, dict):
                            results = search_result.get('search_results', [])
                            for idx, res in enumerate(results):
                                if isinstance(res, dict):
                                    # Store by index for potential mapping
                                    yield ('search_result', {
                                        'index': idx,
                                        'url': res.get('Url') or res.get('url') or '',
                                        'title': res.get('Name') or res.get('name') or '',
                                        'source_id': res.get('SourceId') or ''
                                    })

            if reply.type == ActivityTypes.typing:
                if isinstance(channel_data, dict):
                    stream_type = channel_data.get('streamType')
                    chunk_type = channel_data.get('chunkType')

                    # Informative status messages
                    if stream_type == 'informative' and reply.text:
                        yield ('status', reply.text)

                    # Streaming content chunks
                    elif chunk_type == 'delta' and reply.text:
                        yield ('content', reply.text)

            elif reply.type == ActivityTypes.message:
                # Extract citation metadata from entities (schema.org Claim objects)
                entities = getattr(reply, 'entities', None) or []
                citation_map = {}
                for ent in entities:
                    # Handle both dict and object forms
                    if hasattr(ent, '__dict__'):
                        ent_dict = vars(ent)
                    elif isinstance(ent, dict):
                        ent_dict = ent
                    else:
                        continue

                    # Look for schema.org Claim type entities with citation data
                    ent_type = ent_dict.get('type', '')
                    if 'Claim' in str(ent_type) or 'citation' in str(ent_type).lower():
                        cite_id = ent_dict.get('@id') or ent_dict.get('id') or ''
                        if cite_id:
                            citation_map[cite_id] = {
                                'url': ent_dict.get('url') or ent_dict.get('Url') or ent_dict.get('uri') or ent_dict.get('sameAs') or '',
                                'title': ent_dict.get('name') or ent_dict.get('title') or ent_dict.get('Name') or ''
                            }

                if citation_map:
                    yield ('citations', citation_map)

                # For non-streaming responses, the final message contains the full text
                # (streamType 'final' without prior delta chunks)
                if reply.text:
                    stream_type = channel_data.get('streamType') if isinstance(channel_data, dict) else None
                    # Yield content if it's a final message (non-streamed response)
                    if stream_type == 'final' or stream_type is None:
                        yield ('final_content', reply.text)

                # Check for suggested actions
                if reply.suggested_actions:
                    actions = [a.title for a in reply.suggested_actions.actions]
                    if actions:
                        yield ('suggestion', ", ".join(actions))

            elif reply.type == ActivityTypes.end_of_conversation:
                yield ('status', "Conversation ended.")
                break

    @property
    def conversation_id(self) -> Optional[str]:
        return self._conversation_id

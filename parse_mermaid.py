import re

def parse_mermaid_sequence_diagram(mermaid_code, default_participant_color="#FFFFFF"):
    participants_info = []
    messages_info = []

    lines = mermaid_code.strip().split('\n')

    participant_pattern = re.compile(r'\s*participant\s+(\w+)(?:\s+as\s+)?(.+)?(?:\s*<<(\w+)>>)?')
    message_pattern = re.compile(r'\s*(\w+)-\>\>(\w+):\s*(.+)')

    participant_to_alias = {}

    for line in lines:
        # Ignore comments
        if line.strip().startswith("%%"):
            continue

        participant_match = participant_pattern.match(line)
        message_match = message_pattern.match(line)

        if participant_match:
            participant_text = participant_match.group(1)
            participant_alias = participant_match.group(2) if participant_match.group(2) else participant_text
            participant_color = participant_match.group(3) if participant_match.group(3) else default_participant_color
            participants_info.append({
                'text': participant_alias.strip(),
                'color': participant_color.strip(),
            })
            participant_to_alias[participant_text] = participant_alias.strip()
        elif message_match:
            from_participant = participant_to_alias.get(message_match.group(1), message_match.group(1))
            to_participant = participant_to_alias.get(message_match.group(2), message_match.group(2))
            try:
                from_index = participants_info.index(next(p for p in participants_info if p['text'] == from_participant))
                to_index = participants_info.index(next(p for p in participants_info if p['text'] == to_participant))
            except ValueError:
                print(f"Error: Participant not found in the sequence diagram for message: {message_match.group(3)}")
                continue

            messages_info.append({
                'from': from_index,
                'to': to_index,
                'text': message_match.group(3),
            })

    return participants_info, messages_info

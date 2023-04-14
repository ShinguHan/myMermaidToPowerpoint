import re

def parse_mermaid_sequence_diagram(mermaid_code):
    participants_info = []
    messages_info = []

    lines = mermaid_code.strip().split('\n')

    participant_pattern = re.compile(r'\s*participant\s+(\w+)(?:\s+as\s+)?(.+)?')
    message_pattern = re.compile(r'\s*(\w+)-\>\>(\w+):\s*(.+)')

    for line in lines:
        participant_match = participant_pattern.match(line)
        message_match = message_pattern.match(line)

        if participant_match:
            participant_color = participant_match.group(2) if participant_match.group(2) else ''
            participants_info.append({
                'text': participant_match.group(1),
                'color': participant_color.strip(),
            })
        elif message_match:
            try:
                from_index = participants_info.index(next(p for p in participants_info if p['text'] == message_match.group(1)))
                to_index = participants_info.index(next(p for p in participants_info if p['text'] == message_match.group(2)))
            except ValueError:
                print(f"Error: Participant not found in the sequence diagram for message: {message_match.group(3)}")
                continue

            messages_info.append({
                'from': from_index,
                'to': to_index,
                'text': message_match.group(3),
            })

    return participants_info, messages_info

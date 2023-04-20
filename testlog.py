import re
import datetime
import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx

# 시나리오 정의
scenario = {
    "S6F11": "S6F12",
}

def extract_secs_messages(log_file):
    with open(log_file, 'r') as f:
        log_data = f.read()
    secs_messages = re.findall(r'<SECS Message[\s\S]*?</SECS Message>', log_data)
    return secs_messages

def parse_secs_message(message):
    message_type = re.search(r'<Message Type="(.*?)"', message).group(1)
    timestamp = re.search(r'<Timestamp>(.*?)</Timestamp>', message).group(1)
    timestamp = datetime.datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S.%f')
    content = re.search(r'<Content>([\s\S]*?)</Content>', message).group(1)
    return message_type, timestamp, content.strip()

def create_secs_message_dataframe(messages):
    data = {'Type': [], 'Timestamp': [], 'Content': []}
    for message in messages:
        message_type, timestamp, content = parse_secs_message(message)
        data['Type'].append(message_type)
        data['Timestamp'].append(timestamp)
        data['Content'].append(content)
    return pd.DataFrame(data)

def analyze_message_frequency(df):
    df['Timestamp'] = pd.to_datetime(df['Timestamp'])
    df['Hour'] = df['Timestamp'].dt.hour
    hourly_count = df.groupby('Hour').count()['Type']
    hourly_count.plot(kind='bar', title='Hourly Message Frequency')
    plt.show()

def verify_scenario(df, scenario):
    errors = []
    for index, row in df.iterrows():
        if row['Type'] in scenario:
            expected_response = scenario[row['Type']]
            if index + 1 < len(df) and df.iloc[index + 1]['Type'] == expected_response:
                continue
            else:
                errors.append((index, row['Type'], expected_response))
    return errors

def visualize_scenario(df, scenario, errors):
    G = nx.DiGraph()
    for src, dest in scenario.items():
        G.add_edge(src, dest)

    pos = nx.spring_layout(G)
    nx.draw(G, pos, with_labels=True, node_size=3000, node_color='lightblue', edge_color='gray', font_size=10)

    for error in errors:
        index, src, dest = error
        nx.draw_networkx_edges(G, pos, edgelist=[(src, dest)], edge_color='red', width=2)

    plt.show()

def analyze_secs_messages(log_file):
    messages = extract_secs_messages(log_file)
    df = create_secs_message_dataframe(messages)
    errors = verify_scenario(df, scenario)
    if errors:
        print("Scenario Errors:")
        for error in errors:
            index, src, dest = error
            print(f"Index: {index}, Expected: {src} -> {dest}")
        print("\n---\n")
    visualize_scenario(df, scenario, errors)
    analyze_message_frequency(df)
    return df

if __name__ == '__main__':
    log_file = 'secs_log.txt'  # 로그 파일명을 변경하여 사용하세요
    df = analyze_secs_messages(log_file)
    print(df.head())

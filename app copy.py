import tkinter as tk
from tkinter import filedialog
from parse_mermaid import parse_mermaid_sequence_diagram
from create_presentation import create_powerpoint_sequence_diagram

# Custom colors
bg_color = "#f0f0f0"
button_color = "#d0d0d0"
button_text_color = "#000000"

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.md")])
    if not file_path:
        return  # User pressed cancel

    with open(file_path, "r") as f:
        mermaid_code = f.read()
        mermaid_code_text.delete(1.0, tk.END)
        mermaid_code_text.insert(tk.END, mermaid_code)

def save_presentation():
    mermaid_code = mermaid_code_text.get(1.0, tk.END)
    participants_info, messages_info = parse_mermaid_sequence_diagram(mermaid_code)

    if not participants_info or not messages_info:
        tk.messagebox.showerror("Error", "Failed to parse the Mermaid code. Please check the input.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])

    if save_path:
        create_powerpoint_sequence_diagram(participants_info, messages_info, save_path)

# Custom button class with rounded corners
class RoundedButton(tk.Canvas):
    def __init__(self, parent, width, height, corner_radius, color, text, text_color, command=None):
        tk.Canvas.__init__(self, parent, width=width, height=height, highlightthickness=0)
        self.command = command
        self.create_oval((0, 0, corner_radius*2, corner_radius*2), fill=color, outline="")
        self.create_oval((width-corner_radius*2, 0, width, corner_radius*2), fill=color, outline="")
        self.create_oval((0, height-corner_radius*2, corner_radius*2, height), fill=color, outline="")
        self.create_oval((width-corner_radius*2, height-corner_radius*2, width, height), fill=color, outline="")
        self.create_rectangle((corner_radius, 0, width-corner_radius, height), fill=color, outline="")
        self.create_rectangle((0, corner_radius, width, height-corner_radius), fill=color, outline="")
        self.create_text((width/2, height/2), text=text, fill=text_color, font=("Helvetica", 12))

        self.bind("<Button-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)

    def on_press(self, event):
        if self.command:
            self.command()

    def on_release(self, event):
        pass

# Main window
root = tk.Tk()
root.title("Mermaid to PowerPoint")
root.configure(bg=bg_color)
root.geometry("1000x800")

# Label
title = tk.Label(root, text="Mermaid to PowerPoint", bg=bg_color, font=("Helvetica", 16))
title.pack(pady=10)

# Text input
mermaid_code_text = tk.Text(root, wrap=tk.WORD, font=("Helvetica", 12))
mermaid_code_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# Buttons
buttons_frame = tk.Frame(root, bg=bg_color)
buttons_frame.pack(pady=10)

load_button = RoundedButton(buttons_frame, 150, 30, 10, button_color, "Load Mermaid File", button_text_color, command=browse_file)
load_button.pack(side=tk.LEFT, padx=10)

save_button = RoundedButton(buttons_frame, 150, 30, 10, button_color, "Save PowerPoint", button_text_color, command=save_presentation)
save_button.pack(side=tk.RIGHT, padx=10)

root.mainloop()

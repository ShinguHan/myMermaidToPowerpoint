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
    # Get settings values from the UI
    left_start = int(left_start_var.get())
    top_start = int(top_start_var.get())
    width = int(width_var.get())
    height = int(height_var.get())
    messages_per_slide = int(messages_per_slide_var.get())
    
    mermaid_code = mermaid_code_text.get(1.0, tk.END)
    participants_info, messages_info, autonumber = parse_mermaid_sequence_diagram(mermaid_code)

    if not participants_info or not messages_info:
        tk.messagebox.showerror("Error", "Failed to parse the Mermaid code. Please check the input.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])

    if save_path:
        create_powerpoint_sequence_diagram(participants_info, messages_info, save_path, autonumber, left_start, top_start, width, height, messages_per_slide)
        

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

# Settings input
settings_frame = tk.Frame(root, bg=bg_color)
settings_frame.pack(pady=10)

left_start_label = tk.Label(settings_frame, text="Left Start:", bg=bg_color)
left_start_label.pack(side=tk.LEFT, padx=2)
left_start_var = tk.StringVar()
left_start_var.set("100")
left_start_entry = tk.Entry(settings_frame, textvariable=left_start_var, width=5)
left_start_entry.pack(side=tk.LEFT, padx=2)

top_start_label = tk.Label(settings_frame, text="Top Start:", bg=bg_color)
top_start_label.pack(side=tk.LEFT, padx=2)
top_start_var = tk.StringVar()
top_start_var.set("100")
top_start_entry = tk.Entry(settings_frame, textvariable=top_start_var, width=5)
top_start_entry.pack(side=tk.LEFT, padx=2)

width_label = tk.Label(settings_frame, text="Width:", bg=bg_color)
width_label.pack(side=tk.LEFT, padx=2)
width_var = tk.StringVar()
width_var.set("800")
width_entry = tk.Entry(settings_frame, textvariable=width_var, width=5)
width_entry.pack(side=tk.LEFT, padx=2)

height_label = tk.Label(settings_frame, text="Height:", bg=bg_color)
height_label.pack(side=tk.LEFT, padx=2)
height_var = tk.StringVar()
height_var.set("600")
height_entry = tk.Entry(settings_frame, textvariable=height_var, width=5)
height_entry.pack(side=tk.LEFT, padx=2)

messages_per_slide_label = tk.Label(settings_frame, text="Messages per Slide:", bg=bg_color)
messages_per_slide_label.pack(side=tk.LEFT, padx=2)
messages_per_slide_var = tk.StringVar()
messages_per_slide_var.set("5")
messages_per_slide_entry = tk.Entry(settings_frame, textvariable=messages_per_slide_var, width=5)
messages_per_slide_entry.pack(side=tk.LEFT, padx=2)


root.mainloop()

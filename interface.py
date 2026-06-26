import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import docx
import threading

from editorial_proofreader import editorial_proofreader
import time

# Ensure the folder exists with correct permissions
import stat

def ensure_folder_exists(folder_path):
	if not os.path.exists(folder_path):
		os.makedirs(folder_path, exist_ok=True)

	# Ensure the folder is writable (Fix permission issue on Windows)
	if not os.access(folder_path, os.W_OK):
		try:
			os.chmod(folder_path, stat.S_IWRITE)
		except Exception as e:
			print(f"Warning: Could not modify folder permissions for {folder_path} -> {e}")

def upload_file():
	global uploaded_file_name

	# Ask the user to select a file
	file_path = filedialog.askopenfilename()

	if file_path:
		if not file_path.endswith(".docx"):
			messagebox.showerror("Error", f"File {file_path} must be a .docx Word document.")
			return
		
		target_folder = f"data/input/{datetime.now().strftime('%d-%m-%Y')}/"
		ensure_folder_exists(target_folder)  # ✅ Ensure folder exists before using it
	
		# Get the file name
		file_name = os.path.basename(file_path)
		uploaded_file_name = file_name

		# Define the target file path
		target_file_path = os.path.join(target_folder, file_name)

		# Copy the selected file to the target folder
		shutil.copy(file_path, target_file_path)

		# Update the label with the uploaded file name
		file_name_label.config(text=f"{uploaded_file_name}", fg="steel blue")
	else:
		messagebox.showerror("Error", f"File {file_path} does not exist.")
		return

uploaded_file_name = None

def load_ignored_styles():
	try:
		with open("ignored_styles.txt", "r", encoding="utf-8") as f:
			for line in f:
				ignored_styles.insert(tk.END, line.strip())
	except FileNotFoundError:
		messagebox.showerror("Error", "ignored_styles.txt file was not found")

def save_ignored_styles():
	with open("ignored_styles.txt", "w", encoding="utf-8") as f:
		list = ignored_styles.get(0, tk.END)
		for item in list:
			f.write(f"{item}\n")

def add_ignore_style_popup():
	# Create a pop-up window
	popup = tk.Toplevel()
	popup.title("Select a style to ignore")

	available_styles = []
	# List of available styles to choose from
	if uploaded_file_name is not None:
		available_styles = [style.name for style in docx.Document(f"data/input/{datetime.now().strftime('%d-%m-%Y')}/{uploaded_file_name}").styles]

	# Listbox to display the available styles
	style_listbox = tk.Listbox(popup, selectmode=tk.SINGLE, height=10)
	style_listbox.pack(padx=10, pady=10)
	
	# Add the available styles to the listbox
	for style in available_styles:
		style_listbox.insert(tk.END, style)

	# Function to add selected style to the main ignored_styles Listbox
	def add_selected_style():
		selected = style_listbox.curselection()
		if selected:
			selected_style = style_listbox.get(selected)
			ignored_styles.insert(tk.END, selected_style)
			save_ignored_styles()
			popup.destroy()  # Close the pop-up window after adding the item
	
	# Add button to confirm the selection
	add_button = tk.Button(popup, text="Add", command=add_selected_style)
	add_button.pack(pady=10)

def delete_ignored_style():
	try:
		selected_item_index = ignored_styles.curselection()[0]
		ignored_styles.delete(selected_item_index)
		save_ignored_styles()
	except IndexError:
		messagebox.showwarning("Selection Error", "Select an ignored style to remove.")

def generate_proofread_docx():
	file_path = filedialog.asksaveasfilename(
		initialfile=f"{uploaded_file_name[:-5]}_edited.docx",
		title="Select destination and filename",
		defaultextension=".docx",
		filetypes=[("Word Documents", "*.docx")]
	)

	if file_path:
		save_ignored_styles()
		target_folder = f"data/output/{datetime.now().strftime('%d-%m-%Y')}/"
		ensure_folder_exists(target_folder)
		generate_status_label.config(text=f"Editing {uploaded_file_name} ...")

		try:
            # Convert the radio button's value to a Boolean flag:
			# Use the radio button's value: if pipeline_var is 1, then use both models; else only GPT-4o.
			use_both = True if pipeline_var.get() == 1 else False
			editorial_proofreader(
				docx_file_name=f"{uploaded_file_name[:-5]}",
				styles_excluded=ignored_styles.get(0, tk.END),
				use_both_models=use_both
			)

			shutil.copy(f"{target_folder}{uploaded_file_name[:-5]}_edited.docx", file_path)
			generate_status_label.config(text=f"Successfully edited {uploaded_file_name}", fg="green3")
		except Exception as e:
			generate_status_label.config(text=f"Error: {e}", fg="red")

# Create the main root window
root = tk.Tk()
root.title("Editorial Proofreader")
root.geometry("1000x600")
root.configure(background='gray90')
root.resizable(False, False) # Disable resizing the window

# Upload and status elements
upload_button = tk.Button(root, text="Upload Word document", command=upload_file, bg="steel blue", fg="white", width=30, font=("Arial", 10))
upload_button.grid(row=0, column=0, padx=10, pady=20)
file_name_label = tk.Label(root, text="No file uploaded", font=("Arial", 12), bg="white", width=78, anchor="w")
file_name_label.grid(row=0, column=1, padx=10, pady=20)

generate_button = tk.Button(root, text="Edit document", command=generate_proofread_docx, bg="steel blue", fg="white", font=("Arial", 10), width=30)
generate_button.grid(row=1, column=0, padx=10, pady=0)
generate_status_label = tk.Label(root, text="", font=("Arial", 12), bg="white", width=78, anchor="w")
generate_status_label.grid(row=1, column=1, padx=10, pady=0)

separator_line = ttk.Separator(root, orient="horizontal")
separator_line.grid(row=2, column=0, columnspan=2, sticky="ew", pady=25)

# Ignored styles section
ignored_styles_label = tk.Label(root, text="Ignored styles", bg="white", font=("Arial", 12), justify="left")
ignored_styles_label.grid(row=3, column=0, padx=10, pady=0)
ignored_styles_frame = tk.Frame(root)
ignored_styles_frame.grid(row=4, column=0, padx=10, pady=10)
ignored_styles_frame_scrollbar = tk.Scrollbar(ignored_styles_frame)
ignored_styles_frame_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
ignored_styles = tk.Listbox(ignored_styles_frame, yscrollcommand=ignored_styles_frame_scrollbar.set, width=35, height=20)
ignored_styles.pack(side=tk.LEFT, fill=tk.X)
ignored_styles_frame_scrollbar.config(command=ignored_styles.yview)
load_ignored_styles()

add_ignored_style_button = tk.Button(root, text="Add ignored style", width=30, command=add_ignore_style_popup)
add_ignored_style_button.grid(row=5, column=0, padx=10, pady=10)
delete_ignored_style_button = tk.Button(root, text="Remove ignored style", width=30, command=delete_ignored_style)
delete_ignored_style_button.grid(row=6, column=0, padx=10, pady=0)

# Pipeline radio buttons section (right column)
pipeline_var = tk.IntVar(master=root, value=0)  # 0 = GPT‑4o Only, 1 = GPT‑4o + o3‑mini
pipeline_frame = tk.Frame(root, bg="lightgray", bd=2, relief="sunken")
pipeline_frame.grid(row=3, column=1, rowspan=4, padx=10, pady=10, sticky="n")
tk.Label(pipeline_frame, text="Select pipeline:", bg="lightgray", font=("Arial", 10)).pack(anchor="w", padx=5, pady=2)
tk.Radiobutton(pipeline_frame, text="Claude Haiku only", variable=pipeline_var, value=0).pack(anchor="w")
tk.Radiobutton(pipeline_frame, text="Claude Haiku + Sonnet", variable=pipeline_var, value=1).pack(anchor="w")

root.mainloop()
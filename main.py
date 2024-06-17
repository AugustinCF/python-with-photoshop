import pandas as pd
import win32com.client
import os
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Configuration file path
config_file = 'config.json'

# Function to load the configuration
def load_config():
    print("incarc config")
    if os.path.exists(config_file):
        with open(config_file, 'r') as file:
            return json.load(file)
    return {}

# Function to save the configuration
def save_config():
    config = {
        'template_path': template_path.get(),
        'csv_file_path': csv_file_path.get(),
        'output_path': output_path.get(),
        'save_psd': save_psd_var.get(),
        'save_jpg': save_jpg_var.get()
    }
    with open(config_file, 'w') as file:
        json.dump(config, file)
        print("finish save")

def read_csv_headers(file_path):
    df = pd.read_csv(file_path, nrows=0)
    headers = list(df.columns)
    return headers

file_path = r"C:\Users\Panda\Documents\buldo generate\data.csv"
headers = read_csv_headers(file_path)
print(headers)

# Function to update text layers
def update_text_layer(doc, layer_name, text):
    try:
            textLayer = doc.ArtLayers[layer_name]
            if textLayer.Kind == 2:  # 2 corresponds to the text layer
                textItem = textLayer.TextItem
                words = text.split()
                if len(words) > 1 and 10 < len(text) < 12:
                    textItem.Size = 50  # Set font size to 50 points
                elif len(words) == 2 and len(text) > 12:
                    textItem.Contents = "\r".join(words)  # Use "\r" to represent a line break
                    textItem.Position = (232, 205)
                    textItem.Size = 47
                    luna_layer = doc.ArtLayers['luna']
                    luna_layer.TextItem.Size = 39.55  # Set font size to 39.55 points
                    luna_layer.TextItem.Position = (232, 240)  # Move the 'luna' layer
                elif len(words) == 3:
                    first_line = " ".join(words[:2])
                    second_line = " ".join(words[2:])
                    textItem.Contents = f"{first_line}\r{second_line}"
                    textItem.Position = (232, 205)
                    textItem.Size = 44  # Set font size to 85 points
                    luna_layer = doc.ArtLayers['luna']
                    luna_layer.TextItem.Size = 39.55
                    luna_layer.TextItem.Position = (232, 240)
                else:
                    textItem.Contents = text  # Set the text content without modifications
                    if len(text) > 10:
                        textItem.Size = 50  # Set font size to 50 points

    except Exception as e:
        print(f"Error updating layer {layer_name}: {e}")

def generate_images(template_path, csv_file_path, output_path, save_psd, save_jpg, progress_bar, progress_label):
    data = pd.read_csv(csv_file_path, delimiter=',')
    psApp = win32com.client.Dispatch("Photoshop.Application")
    psApp.Visible = True

    os.makedirs(output_path, exist_ok=True)
    print(data)
    total_images = len(data)
    progress_bar["maximum"] = total_images

    for index, row in data.iterrows():
        city = row['oras']
        month = row['luna']

        month_directory = os.path.join(output_path, month)
        os.makedirs(month_directory, exist_ok=True)

        doc = psApp.Open(template_path)

        update_text_layer(doc, 'oras', city)
        update_text_layer(doc, 'luna', month)

        if save_psd:
            output_path_psd = os.path.join(month_directory, f"{city}_image_{index + 1}.psd")
            doc.SaveAs(output_path_psd)

        if save_jpg:
            output_path_jpg = os.path.join(month_directory, f"{city}_image_{index + 1}.jpg")
            jpg_options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
            jpg_options.Format = 6
            jpg_options.Quality = 100
            doc.Export(ExportIn=output_path_jpg, ExportAs=2, Options=jpg_options)

        doc.Close(2)

        progress_bar["value"] = index + 1
        progress_label["text"] = f"Progress: {index + 1}/{total_images} images processed"
        root.update_idletasks()

    print("Images generated successfully.")
    messagebox.showinfo("Success", "Images generated successfully.")

def browse_template():
    template_path.set(filedialog.askopenfilename(filetypes=[("PSD files", "*.psd")]))
    save_config()

def browse_csv():
    csv_file_path.set(filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")]))
    save_config()

def browse_output():
    output_path.set(filedialog.askdirectory())
    save_config()

def start_generation():
    template = template_path.get()
    csv_file = csv_file_path.get()
    output_dir = output_path.get()
    save_psd = save_psd_var.get()
    save_jpg = save_jpg_var.get()

    if not template or not csv_file or not output_dir:
        messagebox.showwarning("Input Error", "Please select the template, data file, and output directory.")
        return

    generate_images(template, csv_file, output_dir, save_psd, save_jpg, progress_bar, progress_label)

root = tk.Tk()
root.title("Image Generation")

config = load_config()

template_path = tk.StringVar(value=config.get('template_path', ''))
csv_file_path = tk.StringVar(value=config.get('csv_file_path', ''))
output_path = tk.StringVar(value=config.get('output_path', ''))

save_psd_var = tk.BooleanVar(value=config.get('save_psd', True))
save_jpg_var = tk.BooleanVar(value=config.get('save_jpg', True))

ttk.Label(root, text="Template PSD File:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
ttk.Entry(root, textvariable=template_path, width=50).grid(row=0, column=1, padx=10, pady=5)
ttk.Button(root, text="Browse", command=browse_template).grid(row=0, column=2, padx=10, pady=5)

ttk.Label(root, text="CSV Data File:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
ttk.Entry(root, textvariable=csv_file_path, width=50).grid(row=1, column=1, padx=10, pady=5)
ttk.Button(root, text="Browse", command=browse_csv).grid(row=1, column=2, padx=10, pady=5)

ttk.Label(root, text="Output Directory:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
ttk.Entry(root, textvariable=output_path, width=50).grid(row=2, column=1, padx=10, pady=5)
ttk.Button(root, text="Browse", command=browse_output).grid(row=2, column=2, padx=10, pady=5)

ttk.Checkbutton(root, text="Save as PSD", variable=save_psd_var, command=save_config).grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
ttk.Checkbutton(root, text="Save as JPEG", variable=save_jpg_var, command=save_config).grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)

ttk.Button(root, text="Generate Images", command=start_generation).grid(row=4, column=0, columnspan=3, padx=10, pady=20)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

progress_label = ttk.Label(root, text="Progress: 0/0 images processed")
progress_label.grid(row=6, column=0, columnspan=3, padx=10, pady=5)

root.protocol("WM_DELETE_WINDOW", lambda: [save_config(), root.destroy()])

root.mainloop()

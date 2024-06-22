import tkinter as tk
from tkinter import ttk
import csv
import pandas as pd
import json

class DynamicRulesApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Dynamic Rules Application")

        self.rule_count = 0
        self.rules = []
        self.column_names = []

        # Load configuration from JSON
        self.load_config()

        # Add new rules button
        self.add_rule_button = ttk.Button(self, text="Add New Rules", command=self.add_new_rule)
        self.add_rule_button.pack(pady=10)

        # Container for all rules
        self.rules_frame = ttk.Frame(self)
        self.rules_frame.pack(pady=10)

        # Save button
        self.save_button = ttk.Button(self, text="Save", command=self.save_rules)
        self.save_button.pack(pady=20)

    def load_config(self):
        with open('config.json', 'r') as file:
            config = json.load(file)
            self.csv_file_path = config['csv_file_path']
            self.load_csv(self.csv_file_path)

    def load_csv(self, file_path):
        try:
            df = pd.read_csv(file_path)
            self.column_names = df.columns.tolist()
            print("CSV loaded successfully!")
        except Exception as e:
            print(f"Error loading CSV: {e}")

    def add_new_rule(self):
        if not self.column_names:
            print("Please load a CSV file first!")
            return

        self.rule_count += 1
        rule_frame = ttk.Frame(self.rules_frame, borderwidth=2, relief="groove", padding=10)
        rule_frame.pack(pady=5, fill="x")

        # Column name dropdown
        column_label = ttk.Label(rule_frame, text="Column name")
        column_label.grid(row=0, column=0, padx=5, pady=5)
        column_combo = ttk.Combobox(rule_frame, values=self.column_names)
        column_combo.grid(row=0, column=1, padx=5, pady=5)

        # Condition dropdown
        condition_label = ttk.Label(rule_frame, text="Condition")
        condition_label.grid(row=0, column=2, padx=5, pady=5)
        condition_combo = ttk.Combobox(rule_frame, values=[">", ">=", "!==", "==="])
        condition_combo.grid(row=0, column=3, padx=5, pady=5)

        # Value entry
        value_entry = ttk.Entry(rule_frame)
        value_entry.grid(row=0, column=4, padx=5, pady=5)

        # Text size entry
        text_size_label = ttk.Label(rule_frame, text="Text size")
        text_size_label.grid(row=1, column=0, padx=5, pady=5)
        text_size_entry = ttk.Entry(rule_frame)
        text_size_entry.grid(row=1, column=1, padx=5, pady=5)

        # Position entries
        position_label = ttk.Label(rule_frame, text="Position")
        position_label.grid(row=2, column=0, padx=5, pady=5)
        x_entry = ttk.Entry(rule_frame, width=5)
        x_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        y_entry = ttk.Entry(rule_frame, width=5)
        y_entry.grid(row=2, column=2, padx=5, pady=5, sticky="w")

        self.rules.append((column_combo, condition_combo, value_entry, text_size_entry, x_entry, y_entry))

    def save_rules(self):
        with open('rules.json', 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(["Column Name", "Condition", "Value", "Text Size", "Position X", "Position Y"])

            for rule in self.rules:
                column_combo, condition_combo, value_entry, text_size_entry, x_entry, y_entry = rule
                csvwriter.writerow([
                    column_combo.get(),
                    condition_combo.get(),
                    value_entry.get(),
                    text_size_entry.get(),
                    x_entry.get(),
                    y_entry.get()
                ])

        print("Rules saved successfully!")

if __name__ == "__main__":
    app = DynamicRulesApp()
    app.mainloop()

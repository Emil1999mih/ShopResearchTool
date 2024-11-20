import tkinter as tk
from tkinter import filedialog, messagebox
from excelbuilder import save_to_excel, add_price_analysis, add_price_ranges, add_graphs
from skimmer import extract_data

class DataScraperGUI:
    def __init__(self, master):
        self.master = master
        master.title("Website Research Tool")

        # url
        self.url_label = tk.Label(master, text="Enter URL:")
        self.url_label.grid(row=0, column=0, padx=10, pady=10)
        self.url_entry = tk.Entry(master)
        self.url_entry.grid(row=0, column=1, padx=10, pady=10)

        # locatie salvare
        self.save_location_label = tk.Label(master, text="Save Location:")
        self.save_location_label.grid(row=1, column=0, padx=10, pady=10)
        self.save_location_entry = tk.Entry(master, state="readonly")
        self.save_location_entry.grid(row=1, column=1, padx=10, pady=10)
        self.browse_button = tk.Button(master, text="Browse", command=self.handle_browse)
        self.browse_button.grid(row=1, column=2, padx=10, pady=10)

        # buton salvare
        self.scrape_button = tk.Button(master, text="Save Data to Excel", command=self.handle_scrape)
        self.scrape_button.grid(row=2, column=1, padx=10, pady=10)

    def handle_browse(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            self.save_location_entry.config(state="normal")
            self.save_location_entry.delete(0, tk.END)
            self.save_location_entry.insert(0, save_path)
            self.save_location_entry.config(state="readonly")

    def handle_scrape(self):
        url = self.url_entry.get()
        save_path = self.save_location_entry.get()

        if not url:
            tk.messagebox.showerror("Error", "Please enter a valid URL.")
            return
        if not save_path:
            tk.messagebox.showerror("Error", "Please select a save location.")
            return

        try:
            data = extract_data(url)
            save_to_excel(data, save_path)
            add_price_analysis(data, save_path)
            add_price_ranges(data, save_path)
            add_graphs(data, save_path)
            tk.messagebox.showinfo("Success", "Data saved successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataScraperGUI(root)
    root.mainloop()
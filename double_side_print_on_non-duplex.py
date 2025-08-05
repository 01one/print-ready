# this program depends on pyMuPDF
# pip install PyMuPDF

# The consideration.. This program currently does not use ocr… 
# So if the page text are scanned or image.. it may add top of the page "this page intentionally left blank"
# To fix this.. you need to integrate ocr.. or custom blank page detector..


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz
import os
import threading

class PDFDuplexSplitterGUI:
	def __init__(self, root):
		self.root = root
		self.root.title("PDF Duplex Splitter")
		self.root.geometry("600x500")
		self.root.resizable(True, True)
		
		self.input_file_var = tk.StringVar()
		self.output_dir_var = tk.StringVar()
		self.progress_var = tk.IntVar()
		self.status_var = tk.StringVar(value="Ready")
		
		self.setup_ui()
		
	def setup_ui(self):
		main_frame = ttk.Frame(self.root, padding="10")
		main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

		self.root.columnconfigure(0, weight=1)
		self.root.rowconfigure(0, weight=1)
		main_frame.columnconfigure(1, weight=1)
		
		title_label = ttk.Label(main_frame, text="PDF Duplex Splitter", 
							   font=("Arial", 16, "bold"))
		title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
		
		desc_text = ("Split PDF into odd and even pages for manual duplex printing.\n"
					"Perfect for printers without automatic duplex support!")
		desc_label = ttk.Label(main_frame, text=desc_text, justify="center")
		desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
		
		ttk.Label(main_frame, text="Input PDF File:").grid(row=2, column=0, sticky="w", pady=5)
		
		input_frame = ttk.Frame(main_frame)
		input_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
		input_frame.columnconfigure(0, weight=1)
		
		self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file_var, width=50)
		self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
		
		ttk.Button(input_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=1)
		
		ttk.Label(main_frame, text="Output Directory:").grid(row=3, column=0, sticky="w", pady=5)
		
		output_frame = ttk.Frame(main_frame)
		output_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
		output_frame.columnconfigure(0, weight=1)
		
		self.output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, width=50)
		self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
		
		ttk.Button(output_frame, text="Browse", command=self.browse_output_dir).grid(row=0, column=1)
		
		note_label = ttk.Label(main_frame, text="(Leave empty to save in same folder as input file)", 
							  font=("Arial", 8), foreground="gray")
		note_label.grid(row=4, column=1, columnspan=2, sticky="w", pady=(0, 10))
		
		self.split_button = ttk.Button(main_frame, text="Split PDF for Duplex Printing", 
									 command=self.split_pdf_threaded)
		self.split_button.grid(row=5, column=0, columnspan=3, pady=20)
		
		self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
										  maximum=100, mode='determinate')
		self.progress_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
		
		self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
		self.status_label.grid(row=7, column=0, columnspan=3, pady=5)
		
		self.results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
		self.results_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), 
							  pady=(20, 0))
		self.results_frame.columnconfigure(0, weight=1)
		
		self.results_text = tk.Text(self.results_frame, height=8, wrap=tk.WORD)
		scrollbar = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.results_text.yview)
		self.results_text.configure(yscrollcommand=scrollbar.set)
		
		self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
		scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
		
		main_frame.rowconfigure(8, weight=1)
		
		instructions_frame = ttk.LabelFrame(main_frame, text="Printing Instructions", padding="10")
		instructions_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
		
		instructions_text = (
			"1. First, print the ODD pages PDF\n"
			"2. After printing is complete, take the printed pages\n"
			"3. Flip the entire stack and put it back in the printer tray\n"
			"4. Then print the EVEN pages PDF\n"
			"5. Your pages should now be printed on both sides!"
		)
		ttk.Label(instructions_frame, text=instructions_text, justify="left").grid(row=0, column=0, sticky="w")

	def browse_input_file(self):
		filename = filedialog.askopenfilename(
			title="Select PDF file",
			filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
		)
		if filename:
			self.input_file_var.set(filename)
			if not self.output_dir_var.get():
				self.output_dir_var.set(os.path.dirname(filename))
	
	def browse_output_dir(self):
		directory = filedialog.askdirectory(title="Select output directory")
		if directory:
			self.output_dir_var.set(directory)
	
	def is_page_blank(self, page):
		words = page.get_text("words")
		return len(words) == 0

	def add_watermark(self, page, text="This page intentionally left blank"):
		"""Add watermark to prevent printer skipping"""
		page.insert_text((10, 10), text,
						fontsize=8,
						color=(0.7, 0.7, 0.7),
						overlay=True)
	
	def split_pdf_threaded(self):
		if not self.input_file_var.get():
			messagebox.showerror("Error", "Please select an input PDF file!")
			return
		
		self.split_button.config(state="disabled")
		self.progress_var.set(0)
		self.status_var.set("Processing...")
		
		self.results_text.delete(1.0, tk.END)
		
		thread = threading.Thread(target=self.split_pdf)
		thread.daemon = True
		thread.start()

	def split_pdf(self):
		try:
			input_pdf_path = self.input_file_var.get()
			output_dir = self.output_dir_var.get() or os.path.dirname(input_pdf_path)
			
			if not os.path.exists(input_pdf_path):
				raise FileNotFoundError(f"File '{input_pdf_path}' not found!")
			
			if not input_pdf_path.lower().endswith('.pdf'):
				raise ValueError("Input file must be a PDF!")
			
			if not os.path.exists(output_dir):
				os.makedirs(output_dir)
			
			self.update_status("Opening PDF...")
			self.update_progress(10)
			
			pdf_document = fitz.open(input_pdf_path)
			total_pages = len(pdf_document)
			
			for page_num in range(total_pages):
				page = pdf_document.load_page(page_num)
				if self.is_page_blank(page):
					self.add_watermark(page)
			
			self.update_results(f"Processing PDF: {os.path.basename(input_pdf_path)}")
			self.update_results(f"Total pages: {total_pages}")
			
			base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
			
			self.update_status("Creating odd pages PDF...")
			self.update_progress(30)
			
			odd_pdf = fitz.open()
			odd_pages = []
			for page_num in range(0, total_pages, 2):
				odd_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
				odd_pages.append(page_num + 1)
			
			odd_output_path = os.path.join(output_dir, f"Print_First_{base_name}_odd_pages.pdf")
			odd_pdf.save(odd_output_path)
			odd_pdf.close()
			
			self.update_results(f"\nOdd pages saved: {odd_output_path}")
			self.update_results(f"Odd pages: {odd_pages}")
			
			self.update_status("Creating even pages PDF...")
			self.update_progress(60)
			
			even_pdf = fitz.open()
			even_pages = []

			blank_pdf = fitz.open()
			blank_page = blank_pdf.new_page(width=pdf_document[0].rect.width,
										  height=pdf_document[0].rect.height)
			self.add_watermark(blank_page)

			if total_pages % 2 != 0:
				even_pdf.insert_pdf(blank_pdf)
				even_pages.append("Blank")
			
			start_page = total_pages - 1 if total_pages % 2 == 0 else total_pages - 2
			for page_num in range(start_page, 0, -2):
				even_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
				even_pages.append(page_num + 1)

			blank_pdf.close()
			
			even_output_path = os.path.join(output_dir, f"Print_Second_{base_name}_even_pages.pdf")
			even_pdf.save(even_output_path)
			even_pdf.close()
			
			self.update_results(f"\nEven pages saved: {even_output_path}")
			self.update_results(f"Even pages (in reverse order): {even_pages}")
			
			pdf_document.close()
			
			self.update_progress(100)
			self.update_status("Complete! Ready for printing.")
			
			self.update_results(f"\n{'='*50}")
			self.update_results("SUCCESS! Files created:")
			self.update_results(f"• Print FIRST: {os.path.basename(odd_output_path)}")
			self.update_results(f"• Print SECOND: {os.path.basename(even_output_path)}")
			self.update_results(f"{'='*50}")
			
			messagebox.showinfo("Success", 
							  f"PDF prepared for duplex printing!\n\n"
							  f"1. FIRST print: {os.path.basename(odd_output_path)}\n"
							  f"2. Flip pages, reload same paper\n"
							  f"3. SECOND print: {os.path.basename(even_output_path)}")
			
		except Exception as e:
			self.update_status(f"Error: {str(e)}")
			self.update_results(f"\nERROR: {str(e)}")
			messagebox.showerror("Error", f"Failed to process PDF:\n{str(e)}")
		
		finally:
			self.root.after(0, lambda: self.split_button.config(state="normal"))

	def update_status(self, message):
		self.root.after(0, lambda: self.status_var.set(message))
	
	def update_progress(self, value):
		self.root.after(0, lambda: self.progress_var.set(value))
	
	def update_results(self, message):
		def _update():
			self.results_text.insert(tk.END, message + "\n")
			self.results_text.see(tk.END)
		self.root.after(0, _update)

def main():
	root = tk.Tk()
	app = PDFDuplexSplitterGUI(root)
	root.mainloop()

if __name__ == "__main__":
	main()

"""
Editor for HAZOP deviations/notes.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
from models import Deviation


class DeviationEditor(tk.Toplevel):
    """Window for editing a deviation."""
    
    def __init__(self, parent, deviation: Deviation, on_save_callback=None):
        super().__init__(parent)
        self.deviation = deviation
        self.on_save_callback = on_save_callback
        self.result = None
        
        self.title("Edit Deviation")
        self.geometry("600x700")
        
        self.create_widgets()
        self.load_data()
    
    def create_widgets(self):
        """Create the editor widgets."""
        # Main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Deviation
        ttk.Label(main_frame, text="Deviation:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.deviation_entry = ttk.Entry(main_frame, width=50)
        self.deviation_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        
        # Causes
        ttk.Label(main_frame, text="Causes:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Entry widget and Add button on same row
        self.cause_entry = ttk.Entry(main_frame, width=50)
        self.cause_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        self.cause_entry.bind("<Return>", lambda e: self.add_cause())
        
        causes_btn_frame = ttk.Frame(main_frame)
        causes_btn_frame.grid(row=1, column=2, padx=5)
        ttk.Button(causes_btn_frame, text="Add", command=self.add_cause).pack(side=tk.LEFT, padx=2)
        ttk.Button(causes_btn_frame, text="Remove", command=self.remove_cause).pack(side=tk.LEFT, padx=2)
        
        # Listbox below entry and buttons
        causes_frame = ttk.Frame(main_frame)
        causes_frame.grid(row=2, column=1, sticky=tk.EW, pady=5)
        
        self.causes_listbox = tk.Listbox(causes_frame, height=5)
        self.causes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        causes_scroll = ttk.Scrollbar(causes_frame, orient=tk.VERTICAL, command=self.causes_listbox.yview)
        causes_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.causes_listbox.config(yscrollcommand=causes_scroll.set)
        
        # Consequence
        ttk.Label(main_frame, text="Consequence:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.consequence_text = scrolledtext.ScrolledText(main_frame, width=50, height=4)
        self.consequence_text.grid(row=3, column=1, sticky=tk.EW, pady=5)
        
        # Safeguards
        ttk.Label(main_frame, text="Safeguards:").grid(row=4, column=0, sticky=tk.W, pady=5)
        
        # Entry widget and Add button on same row
        self.safeguard_entry = ttk.Entry(main_frame, width=50)
        self.safeguard_entry.grid(row=4, column=1, sticky=tk.EW, pady=5)
        self.safeguard_entry.bind("<Return>", lambda e: self.add_safeguard())
        
        safeguards_btn_frame = ttk.Frame(main_frame)
        safeguards_btn_frame.grid(row=4, column=2, padx=5)
        ttk.Button(safeguards_btn_frame, text="Add", command=self.add_safeguard).pack(side=tk.LEFT, padx=2)
        ttk.Button(safeguards_btn_frame, text="Remove", command=self.remove_safeguard).pack(side=tk.LEFT, padx=2)
        
        # Listbox below entry and buttons
        safeguards_frame = ttk.Frame(main_frame)
        safeguards_frame.grid(row=5, column=1, sticky=tk.EW, pady=5)
        
        self.safeguards_listbox = tk.Listbox(safeguards_frame, height=5)
        self.safeguards_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        safeguards_scroll = ttk.Scrollbar(safeguards_frame, orient=tk.VERTICAL, command=self.safeguards_listbox.yview)
        safeguards_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.safeguards_listbox.config(yscrollcommand=safeguards_scroll.set)
        
        # Recommendations
        ttk.Label(main_frame, text="Recommendations:").grid(row=6, column=0, sticky=tk.W, pady=5)
        
        # Entry widget and Add button on same row
        self.recommendation_entry = ttk.Entry(main_frame, width=50)
        self.recommendation_entry.grid(row=6, column=1, sticky=tk.EW, pady=5)
        self.recommendation_entry.bind("<Return>", lambda e: self.add_recommendation())
        
        recommendations_btn_frame = ttk.Frame(main_frame)
        recommendations_btn_frame.grid(row=6, column=2, padx=5)
        ttk.Button(recommendations_btn_frame, text="Add", command=self.add_recommendation).pack(side=tk.LEFT, padx=2)
        ttk.Button(recommendations_btn_frame, text="Remove", command=self.remove_recommendation).pack(side=tk.LEFT, padx=2)
        
        # Listbox below entry and buttons
        recommendations_frame = ttk.Frame(main_frame)
        recommendations_frame.grid(row=7, column=1, sticky=tk.EW, pady=5)
        
        self.recommendations_listbox = tk.Listbox(recommendations_frame, height=5)
        self.recommendations_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        recommendations_scroll = ttk.Scrollbar(recommendations_frame, orient=tk.VERTICAL, command=self.recommendations_listbox.yview)
        recommendations_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.recommendations_listbox.config(yscrollcommand=recommendations_scroll.set)
        
        # Comments
        ttk.Label(main_frame, text="Comments:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.comments_text = scrolledtext.ScrolledText(main_frame, width=50, height=4)
        self.comments_text.grid(row=8, column=1, sticky=tk.EW, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=9, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)  # Causes listbox row
        main_frame.rowconfigure(5, weight=1)  # Safeguards listbox row
        main_frame.rowconfigure(7, weight=1)  # Recommendations listbox row
    
    def load_data(self):
        """Load deviation data into widgets."""
        self.deviation_entry.delete(0, tk.END)
        self.deviation_entry.insert(0, self.deviation.deviation)
        
        self.causes_listbox.delete(0, tk.END)
        for cause in self.deviation.causes:
            self.causes_listbox.insert(tk.END, cause)
        
        self.consequence_text.delete(1.0, tk.END)
        self.consequence_text.insert(1.0, self.deviation.consequence)
        
        self.safeguards_listbox.delete(0, tk.END)
        for safeguard in self.deviation.safeguards:
            self.safeguards_listbox.insert(tk.END, safeguard)
        
        self.recommendations_listbox.delete(0, tk.END)
        for recommendation in self.deviation.recommendations:
            self.recommendations_listbox.insert(tk.END, recommendation)
        
        self.comments_text.delete(1.0, tk.END)
        self.comments_text.insert(1.0, self.deviation.comments)
    
    def add_cause(self):
        """Add a cause."""
        cause = self.cause_entry.get().strip()
        if cause:
            self.causes_listbox.insert(tk.END, cause)
            self.cause_entry.delete(0, tk.END)
    
    def remove_cause(self):
        """Remove selected cause."""
        selection = self.causes_listbox.curselection()
        if selection:
            self.causes_listbox.delete(selection[0])
    
    def add_safeguard(self):
        """Add a safeguard."""
        safeguard = self.safeguard_entry.get().strip()
        if safeguard:
            self.safeguards_listbox.insert(tk.END, safeguard)
            self.safeguard_entry.delete(0, tk.END)
    
    def remove_safeguard(self):
        """Remove selected safeguard."""
        selection = self.safeguards_listbox.curselection()
        if selection:
            self.safeguards_listbox.delete(selection[0])
    
    def add_recommendation(self):
        """Add a recommendation."""
        recommendation = self.recommendation_entry.get().strip()
        if recommendation:
            self.recommendations_listbox.insert(tk.END, recommendation)
            self.recommendation_entry.delete(0, tk.END)
    
    def remove_recommendation(self):
        """Remove selected recommendation."""
        selection = self.recommendations_listbox.curselection()
        if selection:
            self.recommendations_listbox.delete(selection[0])
    
    def save(self):
        """Save deviation data."""
        self.deviation.deviation = self.deviation_entry.get()
        self.deviation.causes = list(self.causes_listbox.get(0, tk.END))
        self.deviation.consequence = self.consequence_text.get(1.0, tk.END).strip()
        self.deviation.safeguards = list(self.safeguards_listbox.get(0, tk.END))
        self.deviation.recommendations = list(self.recommendations_listbox.get(0, tk.END))
        self.deviation.comments = self.comments_text.get(1.0, tk.END).strip()
        
        if self.on_save_callback:
            self.on_save_callback(self.deviation)
        
        self.destroy()
    
    def cancel(self):
        """Cancel editing."""
        self.destroy()


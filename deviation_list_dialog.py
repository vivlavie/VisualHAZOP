"""
Dialog for viewing and managing deviations for a node.
"""
import tkinter as tk
from tkinter import ttk, messagebox
from models import Node, Deviation
from deviation_editor import DeviationEditor


class DeviationListDialog(tk.Toplevel):
    """Dialog for viewing and managing deviations for a node."""
    
    def __init__(self, parent, node: Node, on_update_callback=None):
        super().__init__(parent)
        self.node = node
        self.on_update_callback = on_update_callback
        
        self.title(f"Deviations for {node.name}")
        self.geometry("600x400")
        self.transient(parent)
        
        self.create_widgets()
        self.refresh_list()
    
    def create_widgets(self):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.deviation_listbox = tk.Listbox(list_frame, height=10)
        self.deviation_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.deviation_listbox.bind("<Double-Button-1>", self.on_double_click)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.deviation_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.deviation_listbox.config(yscrollcommand=scrollbar.set)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Add New", command=self.add_new).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Edit", command=self.edit_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Delete", command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Close", command=self.destroy).pack(side=tk.RIGHT, padx=2)
    
    def refresh_list(self):
        """Refresh the deviation list."""
        self.deviation_listbox.delete(0, tk.END)
        for i, dev in enumerate(self.node.deviations):
            display_text = dev.deviation or f"Deviation {i+1}"
            if len(display_text) > 50:
                display_text = display_text[:47] + "..."
            self.deviation_listbox.insert(tk.END, display_text)
    
    def on_double_click(self, event):
        """Handle double-click on deviation."""
        self.edit_selected()
    
    def add_new(self):
        """Add a new deviation."""
        deviation = Deviation()
        editor = DeviationEditor(self, deviation, on_save_callback=self.save_deviation)
        self.wait_window(editor)
    
    def edit_selected(self):
        """Edit selected deviation."""
        selection = self.deviation_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        deviation = self.node.deviations[index]
        editor = DeviationEditor(self, deviation, on_save_callback=lambda dev: self.update_deviation(index, dev))
        self.wait_window(editor)
    
    def delete_selected(self):
        """Delete selected deviation."""
        selection = self.deviation_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        deviation = self.node.deviations[index]
        if messagebox.askyesno("Delete Deviation", f"Delete deviation '{deviation.deviation}'?"):
            self.node.deviations.pop(index)
            self.refresh_list()
            if self.on_update_callback:
                self.on_update_callback()
    
    def save_deviation(self, deviation):
        """Save a new deviation."""
        self.node.deviations.append(deviation)
        self.refresh_list()
        if self.on_update_callback:
            self.on_update_callback()
    
    def update_deviation(self, index, deviation):
        """Update an existing deviation."""
        self.node.deviations[index] = deviation
        self.refresh_list()
        if self.on_update_callback:
            self.on_update_callback()


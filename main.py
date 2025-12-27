"""
Main application window for VisualHAZOP.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pdf_viewer import PDFViewer
from deviation_editor import DeviationEditor
from deviation_list_dialog import DeviationListDialog
from spreadsheet_view import SpreadsheetView
from models import Node, Deviation, HAZOPData
import os


class VisualHAZOPApp(tk.Tk):
    """Main application class."""
    
    def __init__(self):
        super().__init__()
        self.title("VisualHAZOP - HAZOP Analysis Tool")
        self.geometry("1200x800")
        
        self.hazop_data = HAZOPData()
        self.selected_node = None
        self.spreadsheet_window = None
        self._save_path = None  # Track save path
        
        self.create_menu()
        self.create_toolbar()
        self.create_main_content()
        
        # Bind keyboard shortcuts
        self.bind("<Control-l>", lambda e: self.start_line_creation())
        self.bind("<Control-o>", lambda e: self.open_pdf())
        self.bind("<Control-s>", lambda e: self.save_data())
        self.bind("<Control-Shift-S>", lambda e: self.save_data_as())
        # Bind PageUp/PageDown at main window level as fallback
        self.bind_all("<KeyPress-Prior>", lambda e: self.handle_page_up())
        self.bind_all("<KeyPress-Next>", lambda e: self.handle_page_down())
    
    def create_menu(self):
        """Create the menu bar."""
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open PDF...", command=self.open_pdf, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Load Analysis...", command=self.load_analysis)
        file_menu.add_command(label="Save Analysis", command=self.save_data, accelerator="Ctrl+S")
        file_menu.add_command(label="Save Analysis As...", command=self.save_data_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Create Line", command=self.start_line_creation, accelerator="Ctrl+L")
        edit_menu.add_command(label="Edit Node Properties", command=self.edit_selected_node_properties)
        edit_menu.add_command(label="Add Deviation", command=self.add_deviation_to_selected)
        edit_menu.add_command(label="Manage Deviations", command=self.manage_deviations)
        edit_menu.add_separator()
        edit_menu.add_command(label="Delete Node", command=self.delete_selected_node)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Spreadsheet View", command=self.show_spreadsheet)
        view_menu.add_separator()
        view_menu.add_command(label="Zoom In", command=self.zoom_in, accelerator="Ctrl+=")
        view_menu.add_command(label="Zoom Out", command=self.zoom_out, accelerator="Ctrl+-")
        view_menu.add_command(label="Reset Zoom", command=self.reset_zoom, accelerator="Ctrl+0")
        view_menu.add_separator()
        view_menu.add_command(label="Next Page", command=self.next_page, accelerator="PgDn")
        view_menu.add_command(label="Previous Page", command=self.prev_page, accelerator="PgUp")
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Export to Excel", command=self.export_to_excel)
    
    def create_toolbar(self):
        """Create the toolbar."""
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Open PDF", command=self.open_pdf).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Create Line (Ctrl+L)", command=self.start_line_creation).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(toolbar, text="Edit Node", command=self.edit_selected_node_properties).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Add Deviation", command=self.add_deviation_to_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Manage Deviations", command=self.manage_deviations).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(toolbar, text="Spreadsheet", command=self.show_spreadsheet).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(toolbar, text="Zoom In", command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Zoom Out", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Reset Zoom", command=self.reset_zoom).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(toolbar, text="Save", command=self.save_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Load", command=self.load_analysis).pack(side=tk.LEFT, padx=2)
    
    def create_main_content(self):
        """Create the main content area."""
        # Main container
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # PDF viewer frame
        viewer_frame = ttk.LabelFrame(main_container, text="PDF Viewer", padding="5")
        viewer_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create PDF viewer with shared hazop_data
        self.pdf_viewer = PDFViewer(viewer_frame, hazop_data=self.hazop_data, bg="white")
        self.pdf_viewer.pack(fill=tk.BOTH, expand=True)
        self.pdf_viewer.parent = self
        # Set focus to PDF viewer for keyboard events
        self.pdf_viewer.focus_set()
        
        # Status bar
        self.status_bar = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def open_pdf(self):
        """Open a PDF file."""
        filename = filedialog.askopenfilename(
            title="Open PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.hazop_data.pdf_path = filename
            self.pdf_viewer.load_pdf(filename)
            self.update_status(f"Opened: {os.path.basename(filename)}")
    
    def load_analysis(self):
        """Load HAZOP analysis data from JSON."""
        filename = filedialog.askopenfilename(
            title="Load Analysis",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.hazop_data = HAZOPData.from_json(filename)
                self._save_path = filename  # Set save path when loading
                self.pdf_viewer.load_data(self.hazop_data)
                
                # Load PDF if path exists
                if self.hazop_data.pdf_path and os.path.exists(self.hazop_data.pdf_path):
                    self.pdf_viewer.load_pdf(self.hazop_data.pdf_path)
                elif self.hazop_data.pdf_path:
                    # PDF path in data but file doesn't exist - ask user
                    if messagebox.askyesno("PDF Not Found", 
                                         f"PDF file not found at:\n{self.hazop_data.pdf_path}\n\nWould you like to locate it?"):
                        pdf_file = filedialog.askopenfilename(
                            title="Locate PDF File",
                            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
                        )
                        if pdf_file:
                            self.hazop_data.pdf_path = pdf_file
                            self.pdf_viewer.load_pdf(pdf_file)
                
                self.update_status(f"Loaded analysis from {os.path.basename(filename)}")
                # Refresh spreadsheet if open
                if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
                    self.spreadsheet_window.hazop_data = self.hazop_data
                    self.spreadsheet_window.refresh_data()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load analysis: {str(e)}")
    
    def save_data(self):
        """Save HAZOP analysis data to JSON."""
        if not hasattr(self, '_save_path') or not self._save_path:
            self.save_data_as()
        else:
            try:
                # Ensure we have the latest data from PDF viewer
                self.pdf_viewer.hazop_data = self.hazop_data
                self.hazop_data.to_json(self._save_path)
                node_count = len(self.hazop_data.nodes)
                dev_count = sum(len(node.deviations) for node in self.hazop_data.nodes)
                self.update_status(f"Saved {node_count} nodes, {dev_count} deviations to {os.path.basename(self._save_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")
    
    def save_data_as(self):
        """Save HAZOP analysis data to JSON with file dialog."""
        filename = filedialog.asksaveasfilename(
            title="Save Analysis",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.hazop_data.to_json(filename)
                self._save_path = filename
                self.update_status(f"Saved to {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")
    
    def start_line_creation(self):
        """Start creating a new line."""
        if not self.pdf_viewer.doc:
            messagebox.showwarning("Warning", "Please open a PDF file first.")
            return
        self.pdf_viewer.start_line_creation()
        self.update_status("Line creation mode: Click to add points, ESC or Right-click to finish")
    
    def on_line_creation_started(self):
        """Callback when line creation starts."""
        self.update_status("Line creation mode: Click to add points, ESC or Right-click to finish")
    
    def on_line_creation_ended(self):
        """Callback when line creation ends."""
        self.update_status("Line creation finished")
    
    def on_node_selected(self, node):
        """Callback when a node is selected."""
        self.selected_node = node
        self.update_status(f"Node selected: {node.name}")
    
    def on_node_deselected(self):
        """Callback when a node is deselected."""
        self.selected_node = None
        self.update_status("Node deselected")
    
    def edit_selected_node_properties(self):
        """Edit properties of the selected node."""
        if not self.selected_node:
            messagebox.showwarning("Warning", "Please select a node first.")
            return
        self.edit_node_properties(self.selected_node)
    
    def edit_node_properties(self, node):
        """Edit properties of a node."""
        dialog = NodePropertiesDialog(self, node)
        self.wait_window(dialog)
        if dialog.result:
            # Update node properties
            node.name = dialog.result['name']
            node.color = dialog.result['color']
            node.thickness = dialog.result['thickness']
            node.transparency = dialog.result['transparency']
            node.has_arrow = dialog.result['has_arrow']
            node.font_size = dialog.result['font_size']
            
            self.pdf_viewer.render_page()
            if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
                self.spreadsheet_window.hazop_data = self.hazop_data
                self.spreadsheet_window.refresh_data()
    
    def add_deviation_to_selected(self):
        """Add a deviation to the selected node."""
        if not self.selected_node:
            messagebox.showwarning("Warning", "Please select a node first.")
            return
        self.add_deviation(self.selected_node)
    
    def add_deviation(self, node):
        """Add a deviation to a node."""
        deviation = Deviation()
        editor = DeviationEditor(self, deviation, on_save_callback=lambda dev: self.save_deviation(node, dev))
        self.wait_window(editor)
    
    def save_deviation(self, node, deviation):
        """Save a deviation to a node."""
        node.deviations.append(deviation)
        self.pdf_viewer.render_page()
        if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
            self.spreadsheet_window.hazop_data = self.hazop_data
            self.spreadsheet_window.refresh_data()
    
    def manage_deviations(self):
        """Open deviation management dialog for selected node."""
        if not self.selected_node:
            messagebox.showwarning("Warning", "Please select a node first.")
            return
        
        def on_update():
            self.pdf_viewer.render_page()
            if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
                self.spreadsheet_window.hazop_data = self.hazop_data
                self.spreadsheet_window.refresh_data()
        
        dialog = DeviationListDialog(self, self.selected_node, on_update_callback=on_update)
        self.wait_window(dialog)
    
    def manage_deviations_for_node(self, node):
        """Open deviation management dialog for a specific node."""
        def on_update():
            self.pdf_viewer.render_page()
            if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
                self.spreadsheet_window.hazop_data = self.hazop_data
                self.spreadsheet_window.refresh_data()
        
        dialog = DeviationListDialog(self, node, on_update_callback=on_update)
        self.wait_window(dialog)
    
    def delete_selected_node(self):
        """Delete the selected node."""
        if not self.selected_node:
            messagebox.showwarning("Warning", "Please select a node first.")
            return
        self.pdf_viewer.delete_node(self.selected_node)
        if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
            self.spreadsheet_window.hazop_data = self.hazop_data
            self.spreadsheet_window.refresh_data()
    
    def show_spreadsheet(self):
        """Show the spreadsheet view."""
        if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
            # Update data reference and refresh
            self.spreadsheet_window.hazop_data = self.hazop_data
            self.spreadsheet_window.refresh_data()
            self.spreadsheet_window.lift()
        else:
            self.spreadsheet_window = SpreadsheetView(self, self.hazop_data)
    
    def handle_page_up(self, event=None):
        """Handle PageUp key press."""
        if self.pdf_viewer.doc:
            self.prev_page()
        return "break"
    
    def handle_page_down(self, event=None):
        """Handle PageDown key press."""
        if self.pdf_viewer.doc:
            self.next_page()
        return "break"
    
    def next_page(self):
        """Go to next page."""
        if self.pdf_viewer.doc:
            self.pdf_viewer.next_page()
            self.update_status(f"Page {self.pdf_viewer.current_page + 1} of {self.pdf_viewer.total_pages}")
            # Ensure PDF viewer has focus
            self.pdf_viewer.focus_set()
    
    def prev_page(self):
        """Go to previous page."""
        if self.pdf_viewer.doc:
            self.pdf_viewer.prev_page()
            self.update_status(f"Page {self.pdf_viewer.current_page + 1} of {self.pdf_viewer.total_pages}")
            # Ensure PDF viewer has focus
            self.pdf_viewer.focus_set()
    
    def zoom_in(self):
        """Zoom in the PDF viewer."""
        self.pdf_viewer.zoom_in()
    
    def zoom_out(self):
        """Zoom out the PDF viewer."""
        self.pdf_viewer.zoom_out()
    
    def reset_zoom(self):
        """Reset zoom to fit window."""
        self.pdf_viewer.reset_zoom()
    
    def export_to_excel(self):
        """Export to Excel."""
        if not self.hazop_data.nodes:
            messagebox.showwarning("Warning", "No data to export.")
            return
        
        if self.spreadsheet_window and self.spreadsheet_window.winfo_exists():
            self.spreadsheet_window.export_to_excel()
        else:
            # Create temporary spreadsheet window for export
            temp_window = SpreadsheetView(self, self.hazop_data)
            temp_window.export_to_excel()
            temp_window.destroy()
    
    def update_status(self, message):
        """Update the status bar."""
        self.status_bar.config(text=message)


class NodePropertiesDialog(tk.Toplevel):
    """Dialog for editing node properties."""
    
    def __init__(self, parent, node: Node):
        super().__init__(parent)
        self.node = node
        self.result = None
        
        self.title("Edit Node Properties")
        self.geometry("400x350")
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.load_data()
    
    def create_widgets(self):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        ttk.Label(main_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(main_frame, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        
        # Color
        ttk.Label(main_frame, text="Color:").grid(row=1, column=0, sticky=tk.W, pady=5)
        color_frame = ttk.Frame(main_frame)
        color_frame.grid(row=1, column=1, sticky=tk.EW, pady=5)
        self.color_label = tk.Label(color_frame, width=10, relief=tk.SUNKEN)
        self.color_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(color_frame, text="Choose...", command=self.choose_color).pack(side=tk.LEFT)
        
        # Thickness
        ttk.Label(main_frame, text="Thickness:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.thickness_var = tk.IntVar(value=2)
        thickness_frame = ttk.Frame(main_frame)
        thickness_frame.grid(row=2, column=1, sticky=tk.EW, pady=5)
        ttk.Scale(thickness_frame, from_=1, to=10, orient=tk.HORIZONTAL, 
                  variable=self.thickness_var, command=self.update_thickness_label).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.thickness_label = ttk.Label(thickness_frame, text="2")
        self.thickness_label.pack(side=tk.LEFT, padx=5)
        
        # Transparency
        ttk.Label(main_frame, text="Transparency:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.transparency_var = tk.DoubleVar(value=0.7)
        transparency_frame = ttk.Frame(main_frame)
        transparency_frame.grid(row=3, column=1, sticky=tk.EW, pady=5)
        ttk.Scale(transparency_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL,
                  variable=self.transparency_var, command=self.update_transparency_label).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.transparency_label = ttk.Label(transparency_frame, text="0.7")
        self.transparency_label.pack(side=tk.LEFT, padx=5)
        
        # Font size
        ttk.Label(main_frame, text="Font Size:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.font_size_var = tk.IntVar(value=12)
        font_frame = ttk.Frame(main_frame)
        font_frame.grid(row=4, column=1, sticky=tk.EW, pady=5)
        ttk.Scale(font_frame, from_=8, to=24, orient=tk.HORIZONTAL,
                  variable=self.font_size_var, command=self.update_font_size_label).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.font_size_label = ttk.Label(font_frame, text="12")
        self.font_size_label.pack(side=tk.LEFT, padx=5)
        
        # Arrow
        self.has_arrow_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="Show Arrow", variable=self.has_arrow_var).grid(row=5, column=0, columnspan=2, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        main_frame.columnconfigure(1, weight=1)
    
    def load_data(self):
        """Load node data into widgets."""
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.node.name)
        self.color_label.config(bg=self.node.color)
        self.thickness_var.set(self.node.thickness)
        self.transparency_var.set(self.node.transparency)
        self.font_size_var.set(self.node.font_size)
        self.has_arrow_var.set(self.node.has_arrow)
        self.update_thickness_label(self.node.thickness)
        self.update_transparency_label(self.node.transparency)
        self.update_font_size_label(self.node.font_size)
    
    def choose_color(self):
        """Choose a color."""
        color = colorchooser.askcolor(initialcolor=self.node.color)[1]
        if color:
            self.color_label.config(bg=color)
    
    def update_thickness_label(self, value):
        """Update thickness label."""
        self.thickness_label.config(text=str(int(float(value))))
    
    def update_transparency_label(self, value):
        """Update transparency label."""
        self.transparency_label.config(text=f"{float(value):.2f}")
    
    def update_font_size_label(self, value):
        """Update font size label."""
        self.font_size_label.config(text=str(int(float(value))))
    
    def ok(self):
        """Save and close."""
        self.result = {
            'name': self.name_entry.get(),
            'color': self.color_label.cget('bg'),
            'thickness': int(self.thickness_var.get()),
            'transparency': float(self.transparency_var.get()),
            'has_arrow': self.has_arrow_var.get(),
            'font_size': int(self.font_size_var.get())
        }
        self.destroy()
    
    def cancel(self):
        """Cancel and close."""
        self.destroy()


if __name__ == "__main__":
    app = VisualHAZOPApp()
    app.mainloop()


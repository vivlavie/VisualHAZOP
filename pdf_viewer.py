"""
PDF viewer with overlay drawing capabilities for HAZOP analysis.
"""
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
from models import Node, Deviation, HAZOPData
import math
import io


class PDFViewer(tk.Canvas):
    """Canvas widget for displaying PDF pages with overlay drawings."""
    
    def __init__(self, parent, hazop_data=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.parent = parent
        # Make canvas focusable for keyboard events
        self.config(takefocus=True)
        self.doc = None
        self.current_page = 0
        self.total_pages = 0
        self.scale = 1.0
        self.page_image = None
        self.overlay_image = None
        self.photo = None
        self.hazop_data = hazop_data if hazop_data is not None else HAZOPData()
        
        # Zoom and pan state
        self.zoom_level = 1.0  # Current zoom level (1.0 = 100%)
        self.pan_x = 0  # Pan offset in screen coordinates
        self.pan_y = 0  # Pan offset in screen coordinates
        self.base_zoom = 1.5  # Base zoom for PDF rendering
        self.fit_to_window = True  # Whether to fit to window initially
        self.pdf_page_width = 0  # Original PDF page width (at base_zoom)
        self.pdf_page_height = 0  # Original PDF page height (at base_zoom)
        
        # Panning state
        self.panning = False
        self.pan_start_x = 0
        self.pan_start_y = 0
        
        # Line creation state
        self.creating_line = False
        self.current_line_points = []
        self.current_node = None
        
        # Selection state
        self.selected_node = None
        self.hover_node = None
        
        # Point editing state
        self.editing_node = None  # Node being edited
        self.dragging_point = None  # Index of point being dragged, or None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_original_point = None  # Original PDF coordinates of point being dragged
        
        # Bind events
        self.bind("<Button-1>", self.on_click)
        self.bind("<Double-Button-1>", self.on_double_click)
        self.bind("<Button-3>", self.on_right_click)
        self.bind("<Button-2>", self.on_middle_click)  # Middle mouse button for panning
        self.bind("<B2-Motion>", self.on_middle_drag)  # Middle mouse drag
        self.bind("<ButtonRelease-2>", self.on_middle_release)
        self.bind("<B1-Motion>", self.on_drag)  # Mouse drag for point editing
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Motion>", self.on_motion)
        self.bind("<KeyPress>", self.on_key_press)
        self.bind("<Configure>", self.on_configure)
        self.bind("<Prior>", lambda e: self.handle_page_up())  # Page Up
        self.bind("<Next>", lambda e: self.handle_page_down())    # Page Down
        # Also bind at root level for global handling (works even without focus)
        root = self.winfo_toplevel()
        root.bind_all("<KeyPress-Prior>", lambda e: self.handle_page_up())
        root.bind_all("<KeyPress-Next>", lambda e: self.handle_page_down())
        self.bind("<MouseWheel>", self.on_mouse_wheel)  # Windows/Linux
        self.bind("<Button-4>", self.on_mouse_wheel)  # Linux scroll up
        self.bind("<Button-5>", self.on_mouse_wheel)  # Linux scroll down
        self.focus_set()
        
        # Font for text rendering
        try:
            self.default_font = ImageFont.truetype("arial.ttf", 12)
        except:
            self.default_font = ImageFont.load_default()
    
    def load_pdf(self, filepath: str):
        """Load a PDF file."""
        try:
            self.doc = fitz.open(filepath)
            self.total_pages = len(self.doc)
            self.current_page = 0
            if self.hazop_data:
                self.hazop_data.pdf_path = filepath
            self.render_page()
            # Ensure focus for keyboard events
            self.focus_set()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
    
    def load_data(self, hazop_data: HAZOPData):
        """Load HAZOP data."""
        self.hazop_data = hazop_data
        self.render_page()
    
    def render_page(self):
        """Render the current PDF page with overlays."""
        if not self.doc:
            return
        
        # Get page
        page = self.doc[self.current_page]
        
        # Store original page dimensions (in PDF points) for coordinate conversion
        if self.pdf_page_width == 0:
            # Get page rect in PDF points
            page_rect = page.rect
            self.pdf_page_width = page_rect.width
            self.pdf_page_height = page_rect.height
        
        # Calculate zoom for PDF rendering
        pdf_zoom = self.base_zoom * self.zoom_level
        mat = fitz.Matrix(pdf_zoom, pdf_zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img_data = pix.tobytes("ppm")
        self.page_image = Image.open(io.BytesIO(img_data)).convert("RGBA")
        
        # Create overlay at same size as page image
        self.overlay_image = Image.new("RGBA", self.page_image.size, (0, 0, 0, 0))
        self.draw_overlays()
        
        # Combine images
        combined = Image.alpha_composite(self.page_image, self.overlay_image)
        
        # Calculate display size
        canvas_width = self.winfo_width()
        canvas_height = self.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            if self.fit_to_window and self.zoom_level == 1.0:
                # Fit to window mode
                img_ratio = combined.width / combined.height
                canvas_ratio = canvas_width / canvas_height
                
                if img_ratio > canvas_ratio:
                    display_width = canvas_width
                    display_height = int(canvas_width / img_ratio)
                else:
                    display_height = canvas_height
                    display_width = int(canvas_height * img_ratio)
                
                combined = combined.resize((display_width, display_height), Image.Resampling.LANCZOS)
                # Scale: pixels per PDF point, accounting for current zoom
                # In fit mode: zoom_level = 1.0, so scale = display_width / (pdf_width * base_zoom)
                self.scale = display_width / (self.pdf_page_width * self.base_zoom * self.zoom_level)
                self.pan_x = 0
                self.pan_y = 0
            else:
                # Zoomed mode - combined is already at the correct size from PDF rendering
                # No need to resize, just use it as is
                display_width = combined.width
                display_height = combined.height
                # Scale: pixels per PDF point, accounting for current zoom
                # display_width = pdf_width * base_zoom * zoom_level
                # So scale = display_width / (pdf_width * base_zoom * zoom_level) = 1.0
                # But we need: screen = pdf * scale * base_zoom * zoom_level
                # So scale should be: display_width / (pdf_width * base_zoom * zoom_level)
                self.scale = display_width / (self.pdf_page_width * self.base_zoom * self.zoom_level)
        else:
            self.scale = 1.0
        
        # Convert to PhotoImage
        self.photo = ImageTk.PhotoImage(combined)
        
        # Clear and draw
        self.delete("all")
        
        # Apply pan offset
        x_pos = self.pan_x
        y_pos = self.pan_y
        
        self.create_image(x_pos, y_pos, anchor=tk.NW, image=self.photo)
        
        # Store image reference to prevent garbage collection
        self.image_ref = self.photo
    
    def draw_overlays(self):
        """Draw all node overlays on the overlay image."""
        if not self.overlay_image:
            return
        
        draw = ImageDraw.Draw(self.overlay_image)
        nodes = self.hazop_data.get_nodes_for_page(self.current_page)
        
        # Calculate scale factor to convert PDF coordinates to rendered image coordinates
        # The overlay_image is the same size as page_image, which is rendered at base_zoom * zoom_level
        # So we need to scale PDF coordinates by base_zoom * zoom_level
        render_scale = self.base_zoom * self.zoom_level
        
        for node in nodes:
            if len(node.points) < 2:
                continue
            
            # Convert color to RGBA
            color_rgb = self.hex_to_rgb(node.color)
            alpha = int(255 * node.transparency)
            color_rgba = (*color_rgb, alpha)
            
            # Check if this node is selected or being edited
            is_selected = (node == self.selected_node)
            is_editing = (node == self.editing_node)
            
            # Scale line thickness with zoom
            line_thickness = int(node.thickness * render_scale)
            if is_selected:
                line_thickness = max(int((node.thickness + 3) * render_scale), int(node.thickness * 2 * render_scale))
            
            # Scale points from PDF coordinates to rendered image coordinates
            scaled_points = [(int(p[0] * render_scale), int(p[1] * render_scale)) for p in node.points]
            
            if is_editing:
                # Draw dot-dashed line for nodes being edited
                self.draw_dot_dashed_line(draw, scaled_points, color_rgba, line_thickness)
                # Draw corner points
                self.draw_editing_points(draw, scaled_points, color_rgb, render_scale)
            elif is_selected:
                # Draw dashed line for selected nodes
                self.draw_dashed_line(draw, scaled_points, color_rgba, line_thickness)
            else:
                # Draw solid line for unselected nodes
                for i in range(len(scaled_points) - 1):
                    p1 = scaled_points[i]
                    p2 = scaled_points[i + 1]
                    draw.line([p1, p2], fill=color_rgba, width=line_thickness)
            
            # Draw arrow at end
            if node.has_arrow and len(scaled_points) >= 2:
                self.draw_arrow(draw, scaled_points[-2], scaled_points[-1], color_rgba, line_thickness)
            
            # Draw name along longest segment (using scaled points)
            if node.name:
                self.draw_name(draw, node, color_rgb, scaled_points, render_scale)
            
            # Draw deviation indicators (using scaled points)
            if node.deviations:
                self.draw_deviation_indicators(draw, node, scaled_points, render_scale)
    
    def draw_arrow(self, draw, start, end, color, thickness):
        """Draw an arrow at the end of a line."""
        dx = end[0] - start[0]
        dy = end[1] - start[1]
        length = math.sqrt(dx*dx + dy*dy)
        if length == 0:
            return
        
        # Normalize
        dx /= length
        dy /= length
        
        # Arrow size
        arrow_size = max(10, thickness * 3)
        
        # Arrow points
        angle = math.atan2(dy, dx)
        arrow1 = (
            end[0] - arrow_size * math.cos(angle - math.pi/6),
            end[1] - arrow_size * math.sin(angle - math.pi/6)
        )
        arrow2 = (
            end[0] - arrow_size * math.cos(angle + math.pi/6),
            end[1] - arrow_size * math.sin(angle + math.pi/6)
        )
        
        draw.line([end, arrow1], fill=color, width=thickness)
        draw.line([end, arrow2], fill=color, width=thickness)
    
    def draw_name(self, draw, node, color_rgb, scaled_points=None, render_scale=None):
        """Draw node name along the longest segment."""
        if scaled_points is None:
            scaled_points = node.points
        if len(scaled_points) < 2:
            return
        
        # Find longest segment
        max_length = 0
        longest_start = scaled_points[0]
        longest_end = scaled_points[1]
        
        for i in range(len(scaled_points) - 1):
            p1 = scaled_points[i]
            p2 = scaled_points[i + 1]
            length = math.sqrt((p2[0]-p1[0])**2 + (p2[1]-p1[1])**2)
            if length > max_length:
                max_length = length
                longest_start = p1
                longest_end = p2
        
        # Calculate position and angle
        mid_x = (longest_start[0] + longest_end[0]) / 2
        mid_y = (longest_start[1] + longest_end[1]) / 2
        
        dx = longest_end[0] - longest_start[0]
        dy = longest_end[1] - longest_start[1]
        angle = math.degrees(math.atan2(dy, dx))
        
        # Scale font size with zoom
        scaled_font_size = int(node.font_size * render_scale) if render_scale else node.font_size
        
        # Try to get font
        try:
            font = ImageFont.truetype("arial.ttf", scaled_font_size)
        except:
            font = ImageFont.load_default()
        
        # Draw text with background for visibility
        bbox = draw.textbbox((0, 0), node.name, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # Create text image
        text_img = Image.new("RGBA", (text_width + 4, text_height + 4), (255, 255, 255, 200))
        text_draw = ImageDraw.Draw(text_img)
        text_draw.text((2, 2), node.name, fill=(*color_rgb, 255), font=font)
        
        # Rotate if needed
        if abs(angle) > 45 and abs(angle) < 135:
            text_img = text_img.rotate(90, expand=True)
        
        # Paste onto overlay
        paste_x = int(mid_x - text_img.width / 2)
        paste_y = int(mid_y - text_img.height / 2)
        self.overlay_image.paste(text_img, (paste_x, paste_y), text_img)
    
    def draw_deviation_indicators(self, draw, node, scaled_points=None, render_scale=None):
        """Draw indicators for deviations on a node."""
        if not node.deviations:
            return
        
        if scaled_points is None:
            scaled_points = node.points
        
        count = len(node.deviations)
        
        if len(scaled_points) < 2:
            return
        
        # Find the center segment of the line
        total_length = 0
        segment_lengths = []
        for i in range(len(scaled_points) - 1):
            p1 = scaled_points[i]
            p2 = scaled_points[i + 1]
            seg_len = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
            segment_lengths.append(seg_len)
            total_length += seg_len
        
        if total_length == 0:
            return
        
        # Find the center point along the line
        target_length = total_length / 2
        current_length = 0
        center_x, center_y = scaled_points[0]
        
        for i in range(len(scaled_points) - 1):
            seg_len = segment_lengths[i]
            if current_length + seg_len >= target_length:
                # Center is on this segment
                p1 = scaled_points[i]
                p2 = scaled_points[i + 1]
                t = (target_length - current_length) / seg_len
                center_x = p1[0] + t * (p2[0] - p1[0])
                center_y = p1[1] + t * (p2[1] - p1[1])
                break
            current_length += seg_len
        
        # Get line direction at center for arranging circles
        # Find the segment containing the center
        current_length = 0
        line_dx, line_dy = 0, 0
        for i in range(len(scaled_points) - 1):
            seg_len = segment_lengths[i]
            if current_length + seg_len >= target_length:
                p1 = scaled_points[i]
                p2 = scaled_points[i + 1]
                line_dx = p2[0] - p1[0]
                line_dy = p2[1] - p1[1]
                line_len = math.sqrt(line_dx**2 + line_dy**2)
                if line_len > 0:
                    line_dx /= line_len
                    line_dy /= line_len
                break
            current_length += seg_len
        
        # If we couldn't determine direction, use first segment
        if line_dx == 0 and line_dy == 0 and len(scaled_points) >= 2:
            p1 = scaled_points[0]
            p2 = scaled_points[1]
            line_dx = p2[0] - p1[0]
            line_dy = p2[1] - p1[1]
            line_len = math.sqrt(line_dx**2 + line_dy**2)
            if line_len > 0:
                line_dx /= line_len
                line_dy /= line_len
        
        # Calculate perpendicular direction for offset
        perp_dx = -line_dy
        perp_dy = line_dx
        
        # Get node color
        color_rgb = self.hex_to_rgb(node.color)
        circle_color = (*color_rgb, 255)
        outline_color = (255, 255, 255, 255)  # White outline
        
        # Scale radius and spacing with zoom
        base_radius = 8
        radius = int(base_radius * render_scale) if render_scale else base_radius
        spacing = radius * 2.5  # Spacing between circles
        
        # Calculate starting position (centered)
        total_spread = (count - 1) * spacing if count > 1 else 0
        start_offset = -total_spread / 2
        
        for i in range(count):
            # Position along the line
            offset = start_offset + i * spacing
            
            # Calculate circle center
            circle_x = center_x + offset * line_dx + perp_dx * (radius + 2)  # Offset perpendicular to line
            circle_y = center_y + offset * line_dy + perp_dy * (radius + 2)
            
            # Draw circle with node color
            draw.ellipse([circle_x - radius, circle_y - radius, 
                         circle_x + radius, circle_y + radius], 
                        fill=circle_color, outline=outline_color, width=2)
    
    def hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple."""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def screen_to_pdf_coords(self, x, y):
        """Convert screen coordinates to PDF coordinates."""
        if not self.photo:
            return x, y
        # Account for pan offset and scaling
        # scale = display_width / (pdf_width * base_zoom * zoom_level)
        # So: screen_coord = pdf_coord * scale * base_zoom * zoom_level + pan
        # Therefore: pdf_coord = (screen_coord - pan) / (scale * base_zoom * zoom_level)
        # But scale * base_zoom * zoom_level = display_width / pdf_width
        # So: pdf_coord = (screen_coord - pan) * pdf_width / display_width
        # Or: pdf_coord = (screen_coord - pan) / scale / base_zoom / zoom_level
        effective_scale = self.scale * self.base_zoom * self.zoom_level
        pdf_x = (x - self.pan_x) / effective_scale
        pdf_y = (y - self.pan_y) / effective_scale
        return int(pdf_x), int(pdf_y)
    
    def pdf_to_screen_coords(self, x, y):
        """Convert PDF coordinates to screen coordinates."""
        # screen_coord = pdf_coord * scale * base_zoom * zoom_level + pan
        effective_scale = self.scale * self.base_zoom * self.zoom_level
        screen_x = x * effective_scale + self.pan_x
        screen_y = y * effective_scale + self.pan_y
        return int(screen_x), int(screen_y)
    
    def on_click(self, event):
        """Handle mouse click."""
        # Ensure canvas has focus for keyboard events
        self.focus_set()
        
        if self.creating_line:
            # Add point to current line
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            self.current_line_points.append((pdf_x, pdf_y))
            
            if not self.current_node:
                # Create new node
                self.current_node = Node(
                    name=f"Line {len(self.hazop_data.nodes) + 1}",
                    page_number=self.current_page,
                    points=[(pdf_x, pdf_y)]
                )
                self.hazop_data.add_node(self.current_node)
            else:
                self.current_node.points.append((pdf_x, pdf_y))
            
            self.render_page()
        elif self.editing_node:
            # Check if clicking on a point to drag
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            point_index = self.find_point_near(pdf_x, pdf_y, self.editing_node, tolerance=15)
            
            if point_index is not None:
                # Start dragging this point
                self.dragging_point = point_index
                self.drag_start_x = event.x
                self.drag_start_y = event.y
                self.drag_original_point = self.editing_node.points[point_index]
            else:
                # Clicked elsewhere, exit editing mode
                self.end_editing()
        else:
            # Check if clicking on a node (with improved selection)
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            clicked_node = self.find_node_at_point(pdf_x, pdf_y, tolerance=25)  # Increased tolerance
            
            if clicked_node:
                self.selected_node = clicked_node
                self.render_page()  # Re-render to show selection
                if self.parent:
                    self.parent.on_node_selected(clicked_node)
            else:
                self.selected_node = None
                self.render_page()  # Re-render to clear selection
                if self.parent:
                    self.parent.on_node_deselected()
    
    def on_double_click(self, event):
        """Handle double-click to enter editing mode."""
        if self.creating_line or self.editing_node:
            return
        
        # Check if double-clicking on a node
        pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
        clicked_node = self.find_node_at_point(pdf_x, pdf_y, tolerance=25)
        
        if clicked_node and len(clicked_node.points) >= 2:
            self.start_editing(clicked_node)
    
    def on_drag(self, event):
        """Handle mouse drag for point editing."""
        if self.dragging_point is not None and self.editing_node:
            # Calculate movement
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            
            # Convert screen movement to PDF coordinates
            effective_scale = self.scale * self.base_zoom * self.zoom_level
            pdf_dx = dx / effective_scale
            pdf_dy = dy / effective_scale
            
            # Get original point
            orig_x, orig_y = self.drag_original_point
            
            # Allow free movement in any direction
            new_x = orig_x + pdf_dx
            new_y = orig_y + pdf_dy
            
            # Update point
            self.editing_node.points[self.dragging_point] = (new_x, new_y)
            
            # Update drag start for next movement
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.drag_original_point = (new_x, new_y)
            
            # Throttle rendering to avoid lag - only render every N drag events
            if not hasattr(self, '_drag_render_count'):
                self._drag_render_count = 0
            self._drag_render_count += 1
            
            # Render every 3rd drag event for smoother performance
            if self._drag_render_count % 3 == 0:
                self.render_page()
            else:
                # Quick update: just redraw the canvas without full PDF re-render
                self.quick_redraw()
    
    def on_release(self, event):
        """Handle mouse release."""
        if self.dragging_point is not None:
            self.dragging_point = None
            # Final render to ensure everything is up to date
            if hasattr(self, '_drag_render_count'):
                self._drag_render_count = 0
            self.render_page()
    
    def quick_redraw(self):
        """Quick redraw of overlays without full PDF re-render."""
        if not self.doc or not self.overlay_image or not self.page_image:
            return
        
        # Recreate overlay with updated points
        self.overlay_image = Image.new("RGBA", self.page_image.size, (0, 0, 0, 0))
        self.draw_overlays()
        
        # Combine images
        combined = Image.alpha_composite(self.page_image, self.overlay_image)
        
        # Calculate display size (same as in render_page)
        canvas_width = self.winfo_width()
        canvas_height = self.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            if self.fit_to_window and self.zoom_level == 1.0:
                img_ratio = combined.width / combined.height
                canvas_ratio = canvas_width / canvas_height
                
                if img_ratio > canvas_ratio:
                    display_width = canvas_width
                    display_height = int(canvas_width / img_ratio)
                else:
                    display_height = canvas_height
                    display_width = int(canvas_height * img_ratio)
                
                combined = combined.resize((display_width, display_height), Image.Resampling.LANCZOS)
            else:
                display_width = combined.width
                display_height = combined.height
        else:
            return
        
        # Convert to PhotoImage
        self.photo = ImageTk.PhotoImage(combined)
        
        # Update canvas
        self.delete("all")
        x_pos = self.pan_x
        y_pos = self.pan_y
        self.create_image(x_pos, y_pos, anchor=tk.NW, image=self.photo)
        
        # Store image reference
        self.image_ref = self.photo
    
    def on_right_click(self, event):
        """Handle right mouse click."""
        if self.creating_line:
            # End line creation
            self.end_line_creation()
        elif self.editing_node:
            # In editing mode: check if clicking on a point or on the line
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            point_index = self.find_point_near(pdf_x, pdf_y, self.editing_node, tolerance=15)
            
            if point_index is not None:
                # Right-click on a point: show menu to edit/remove
                self.show_point_context_menu(event.x, event.y, point_index)
            else:
                # Right-click on the line: add a new point
                self.add_point_to_line(pdf_x, pdf_y)
        else:
            # Show context menu
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            clicked_node = self.find_node_at_point(pdf_x, pdf_y)
            if clicked_node:
                self.show_node_context_menu(event.x, event.y, clicked_node)
    
    def on_motion(self, event):
        """Handle mouse motion."""
        if not self.creating_line:
            pdf_x, pdf_y = self.screen_to_pdf_coords(event.x, event.y)
            hover_node = self.find_node_at_point(pdf_x, pdf_y)
            if hover_node != self.hover_node:
                self.hover_node = hover_node
                # Could update cursor here
    
    def on_key_press(self, event):
        """Handle key press."""
        if event.keysym == 'Escape':
            if self.creating_line:
                self.end_line_creation()
            elif self.editing_node:
                self.end_editing()
        elif event.keysym == 'Prior':  # Page Up
            self.prev_page()
        elif event.keysym == 'Next':  # Page Down
            self.next_page()
        elif event.state & 0x4 and event.keysym.lower() == 'l':  # Ctrl-L
            self.start_line_creation()
        elif event.state & 0x4 and event.keysym == '0':  # Ctrl-0
            self.reset_zoom()
        elif event.state & 0x4 and (event.keysym == 'plus' or event.keysym == 'equal'):  # Ctrl-+ or Ctrl-=
            self.zoom_in()
        elif event.state & 0x4 and event.keysym == 'minus':  # Ctrl--
            self.zoom_out()
    
    def on_configure(self, event):
        """Handle canvas resize."""
        if self.doc and event.width > 1 and event.height > 1:
            # Re-render on resize (with debouncing to avoid too many renders)
            if self.fit_to_window and self.zoom_level == 1.0:
                self.after(100, self.render_page)
    
    def on_mouse_wheel(self, event):
        """Handle mouse wheel for zooming."""
        if not self.doc:
            return
        
        # Check if Ctrl is held for zoom, otherwise allow normal scrolling
        if event.state & 0x4:  # Ctrl key
            # Zoom at mouse position
            if event.delta:  # Windows/Mac
                delta = event.delta
            elif event.num == 4:  # Linux scroll up
                delta = 120
            elif event.num == 5:  # Linux scroll down
                delta = -120
            else:
                return
            
            # Get mouse position
            mouse_x = event.x
            mouse_y = event.y
            
            # Convert to PDF coordinates before zoom
            pdf_x, pdf_y = self.screen_to_pdf_coords(mouse_x, mouse_y)
            
            # Zoom
            old_zoom = self.zoom_level
            zoom_factor = 1.1 if delta > 0 else 1/1.1
            new_zoom = old_zoom * zoom_factor
            new_zoom = max(0.1, min(5.0, new_zoom))  # Limit zoom between 10% and 500%
            
            if new_zoom != old_zoom:
                # Calculate effective scale before zoom change
                old_effective_scale = self.scale * self.base_zoom * old_zoom
                
                self.zoom_level = new_zoom
                self.fit_to_window = False
                
                # Calculate new effective scale (scale will be updated in render, but we approximate)
                # The scale changes proportionally with zoom in zoomed mode
                zoom_ratio = new_zoom / old_zoom if old_zoom > 0 else 1.0
                new_effective_scale = old_effective_scale * zoom_ratio
                
                # Adjust pan to keep the point under mouse in the same place
                new_screen_x = pdf_x * new_effective_scale + self.pan_x
                new_screen_y = pdf_y * new_effective_scale + self.pan_y
                
                self.pan_x += mouse_x - new_screen_x
                self.pan_y += mouse_y - new_screen_y
                
                self.render_page()
    
    def on_middle_click(self, event):
        """Start panning with middle mouse button."""
        if not self.creating_line:
            self.panning = True
            self.pan_start_x = event.x
            self.pan_start_y = event.y
    
    def on_middle_drag(self, event):
        """Pan the view with middle mouse drag."""
        if self.panning:
            dx = event.x - self.pan_start_x
            dy = event.y - self.pan_start_y
            self.pan_x += dx
            self.pan_y += dy
            self.pan_start_x = event.x
            self.pan_start_y = event.y
            self.render_page()
    
    def on_middle_release(self, event):
        """Stop panning."""
        self.panning = False
    
    def zoom_in(self, factor=1.2, mouse_x=None, mouse_y=None):
        """Zoom in around mouse cursor position."""
        if not self.doc:
            return
        
        # Get mouse position - use current cursor position if not provided
        if mouse_x is None or mouse_y is None:
            try:
                # Get current mouse position relative to canvas
                mouse_x = self.winfo_pointerx() - self.winfo_rootx()
                mouse_y = self.winfo_pointery() - self.winfo_rooty()
            except:
                # Fallback to center of canvas if we can't get mouse position
                mouse_x = self.winfo_width() / 2
                mouse_y = self.winfo_height() / 2
        
        # Convert to PDF coordinates before zoom
        pdf_x, pdf_y = self.screen_to_pdf_coords(mouse_x, mouse_y)
        
        # Zoom
        old_zoom = self.zoom_level
        new_zoom = old_zoom * factor
        new_zoom = max(0.1, min(5.0, new_zoom))
        
        if new_zoom != old_zoom:
            # Calculate effective scale before zoom change
            old_effective_scale = self.scale * self.base_zoom * old_zoom
            
            self.zoom_level = new_zoom
            self.fit_to_window = False
            
            # Calculate new effective scale (scale will be updated in render, but we approximate)
            # The scale changes proportionally with zoom in zoomed mode
            zoom_ratio = new_zoom / old_zoom if old_zoom > 0 else 1.0
            new_effective_scale = old_effective_scale * zoom_ratio
            
            # Adjust pan to keep the point under mouse in the same place
            new_screen_x = pdf_x * new_effective_scale + self.pan_x
            new_screen_y = pdf_y * new_effective_scale + self.pan_y
            
            self.pan_x += mouse_x - new_screen_x
            self.pan_y += mouse_y - new_screen_y
            
            self.render_page()
    
    def zoom_out(self, factor=1.2, mouse_x=None, mouse_y=None):
        """Zoom out around mouse cursor position."""
        if not self.doc:
            return
        
        # Get mouse position - use current cursor position if not provided
        if mouse_x is None or mouse_y is None:
            try:
                # Get current mouse position relative to canvas
                mouse_x = self.winfo_pointerx() - self.winfo_rootx()
                mouse_y = self.winfo_pointery() - self.winfo_rooty()
            except:
                # Fallback to center of canvas if we can't get mouse position
                mouse_x = self.winfo_width() / 2
                mouse_y = self.winfo_height() / 2
        
        # Convert to PDF coordinates before zoom
        pdf_x, pdf_y = self.screen_to_pdf_coords(mouse_x, mouse_y)
        
        # Zoom
        new_zoom = self.zoom_level / factor
        new_zoom = max(0.1, min(5.0, new_zoom))
        
        if new_zoom != self.zoom_level:
            self.zoom_level = new_zoom
            self.fit_to_window = False
            
            # Adjust pan to keep the point under mouse in the same place
            effective_scale = self.scale * self.base_zoom * self.zoom_level
            new_screen_x = pdf_x * effective_scale + self.pan_x
            new_screen_y = pdf_y * effective_scale + self.pan_y
            
            self.pan_x += mouse_x - new_screen_x
            self.pan_y += mouse_y - new_screen_y
            
            self.render_page()
    
    def reset_zoom(self):
        """Reset zoom to fit window."""
        if not self.doc:
            return
        self.zoom_level = 1.0
        self.fit_to_window = True
        self.pan_x = 0
        self.pan_y = 0
        self.render_page()
    
    def start_line_creation(self):
        """Start creating a new line."""
        self.creating_line = True
        self.current_line_points = []
        self.current_node = None
        if self.parent:
            self.parent.on_line_creation_started()
    
    def end_line_creation(self):
        """End line creation."""
        if self.current_node and len(self.current_node.points) < 2:
            # Remove node if it has less than 2 points
            self.hazop_data.remove_node(self.current_node)
        
        self.creating_line = False
        self.current_line_points = []
        self.current_node = None
        if self.parent:
            self.parent.on_line_creation_ended()
        self.render_page()
    
    def find_node_at_point(self, x, y, tolerance=20):
        """Find node near a point by checking distance to line segments."""
        nodes = self.hazop_data.get_nodes_for_page(self.current_page)
        tolerance_pdf = tolerance / self.scale if self.scale > 0 else tolerance
        
        closest_node = None
        closest_distance = float('inf')
        
        for node in nodes:
            if len(node.points) < 2:
                continue
            
            # Check distance to each line segment
            for i in range(len(node.points) - 1):
                p1 = node.points[i]
                p2 = node.points[i + 1]
                
                # Calculate distance from point to line segment
                dist = self.point_to_line_distance((x, y), p1, p2)
                
                if dist < closest_distance:
                    closest_distance = dist
                    closest_node = node
            
            # Also check distance to endpoints (for easier selection)
            for point in node.points:
                dist = math.sqrt((point[0] - x)**2 + (point[1] - y)**2)
                if dist < closest_distance:
                    closest_distance = dist
                    closest_node = node
        
        # Return node if within tolerance
        if closest_node and closest_distance <= tolerance_pdf:
            return closest_node
        
        return None
    
    def point_to_line_distance(self, point, line_start, line_end):
        """Calculate distance from a point to a line segment."""
        px, py = point
        x1, y1 = line_start
        x2, y2 = line_end
        
        # Vector from line_start to line_end
        dx = x2 - x1
        dy = y2 - y1
        
        # If line segment is actually a point
        if dx == 0 and dy == 0:
            return math.sqrt((px - x1)**2 + (py - y1)**2)
        
        # Calculate t (parameter along line segment)
        t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)))
        
        # Closest point on line segment
        closest_x = x1 + t * dx
        closest_y = y1 + t * dy
        
        # Distance from point to closest point on line
        return math.sqrt((px - closest_x)**2 + (py - closest_y)**2)
    
    def draw_dashed_line(self, draw, points, color, width, dash_length=10, gap_length=5):
        """Draw a dashed line through multiple points."""
        for i in range(len(points) - 1):
            p1 = points[i]
            p2 = points[i + 1]
            
            # Calculate line length and direction
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            length = math.sqrt(dx*dx + dy*dy)
            
            if length == 0:
                continue
            
            # Normalize direction
            dx /= length
            dy /= length
            
            # Draw dashed segments
            current_length = 0
            draw_dash = True
            
            while current_length < length:
                if draw_dash:
                    # Start of dash
                    start_x = p1[0] + current_length * dx
                    start_y = p1[1] + current_length * dy
                    
                    # End of dash (or end of line)
                    dash_end = min(current_length + dash_length, length)
                    end_x = p1[0] + dash_end * dx
                    end_y = p1[1] + dash_end * dy
                    
                    draw.line([(start_x, start_y), (end_x, end_y)], fill=color, width=width)
                
                # Move to next segment
                segment_length = dash_length if draw_dash else gap_length
                current_length += segment_length
                draw_dash = not draw_dash
    
    def draw_dot_dashed_line(self, draw, points, color, width, dot_length=3, dash_length=8, gap_length=4):
        """Draw a dot-dashed line through multiple points."""
        for i in range(len(points) - 1):
            p1 = points[i]
            p2 = points[i + 1]
            
            # Calculate line length and direction
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            length = math.sqrt(dx*dx + dy*dy)
            
            if length == 0:
                continue
            
            # Normalize direction
            dx /= length
            dy /= length
            
            # Draw dot-dashed segments: dot, gap, dash, gap, dot, gap, dash, gap...
            current_length = 0
            segment_type = 0  # 0=dot, 1=gap, 2=dash, 3=gap
            
            while current_length < length:
                if segment_type == 0:  # Dot
                    # Draw a small dot
                    dot_x = p1[0] + current_length * dx
                    dot_y = p1[1] + current_length * dy
                    dot_radius = max(2, width // 2)
                    draw.ellipse([dot_x - dot_radius, dot_y - dot_radius,
                                dot_x + dot_radius, dot_y + dot_radius],
                               fill=color)
                    current_length += dot_length
                    segment_type = 1
                elif segment_type == 1:  # Gap after dot
                    current_length += gap_length
                    segment_type = 2
                elif segment_type == 2:  # Dash
                    start_x = p1[0] + current_length * dx
                    start_y = p1[1] + current_length * dy
                    dash_end = min(current_length + dash_length, length)
                    end_x = p1[0] + dash_end * dx
                    end_y = p1[1] + dash_end * dy
                    draw.line([(start_x, start_y), (end_x, end_y)], fill=color, width=width)
                    current_length = dash_end
                    segment_type = 3
                elif segment_type == 3:  # Gap after dash
                    current_length += gap_length
                    segment_type = 0
    
    def draw_editing_points(self, draw, scaled_points, color_rgb, render_scale):
        """Draw corner points for editing as rectangles."""
        point_size = max(8, int(8 * render_scale))
        point_color = (*color_rgb, 255)
        outline_color = (255, 255, 255, 255)
        
        for point in scaled_points:
            x, y = point
            # Draw filled rectangle
            half_size = point_size // 2
            draw.rectangle([x - half_size, y - half_size,
                           x + half_size, y + half_size],
                          fill=point_color, outline=outline_color, width=2)
    
    def find_point_near(self, x, y, node, tolerance=15):
        """Find point index near a PDF coordinate."""
        render_scale = self.base_zoom * self.zoom_level
        tolerance_pdf = tolerance / render_scale if render_scale > 0 else tolerance
        
        closest_index = None
        closest_distance = float('inf')
        
        for i, point in enumerate(node.points):
            dist = math.sqrt((point[0] - x)**2 + (point[1] - y)**2)
            if dist < closest_distance:
                closest_distance = dist
                closest_index = i
        
        if closest_index is not None and closest_distance <= tolerance_pdf:
            return closest_index
        return None
    
    def find_segment_for_point(self, x, y, node, tolerance=25):
        """Find the line segment closest to a point, returning the index after which to insert."""
        render_scale = self.base_zoom * self.zoom_level
        tolerance_pdf = tolerance / render_scale if render_scale > 0 else tolerance
        
        if len(node.points) < 2:
            return None
        
        closest_segment = None
        closest_distance = float('inf')
        insert_index = 0
        
        for i in range(len(node.points) - 1):
            p1 = node.points[i]
            p2 = node.points[i + 1]
            
            # Calculate distance from point to line segment
            dist = self.point_to_line_distance((x, y), p1, p2)
            
            if dist < closest_distance:
                closest_distance = dist
                closest_segment = i
                # Determine which side of the segment the point is closer to
                dist_to_p1 = math.sqrt((p1[0] - x)**2 + (p1[1] - y)**2)
                dist_to_p2 = math.sqrt((p2[0] - x)**2 + (p2[1] - y)**2)
                insert_index = i + 1 if dist_to_p2 < dist_to_p1 else i
        
        if closest_segment is not None and closest_distance <= tolerance_pdf:
            return insert_index
        return None
    
    def show_point_context_menu(self, x, y, point_index):
        """Show context menu for a point in editing mode."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Remove Point", command=lambda: self.remove_point(point_index))
        menu.post(x, y)
    
    def remove_point(self, point_index):
        """Remove a point from the editing node."""
        if self.editing_node and len(self.editing_node.points) > 2:
            # Need at least 2 points for a line
            self.editing_node.points.pop(point_index)
            self.render_page()
    
    def add_point_to_line(self, pdf_x, pdf_y):
        """Add a new point to the line at the specified location."""
        if not self.editing_node:
            return
        
        # Find the segment closest to the click point
        insert_index = self.find_segment_for_point(pdf_x, pdf_y, self.editing_node, tolerance=25)
        
        if insert_index is not None:
            # Insert the new point
            self.editing_node.points.insert(insert_index, (pdf_x, pdf_y))
            self.render_page()
    
    def start_editing(self, node):
        """Start editing a node's points."""
        self.editing_node = node
        self.selected_node = node
        self.dragging_point = None
        self.render_page()
        if self.parent:
            self.parent.on_node_selected(node)
    
    def end_editing(self):
        """End editing mode."""
        self.editing_node = None
        self.dragging_point = None
        self.render_page()
    
    def show_node_context_menu(self, x, y, node):
        """Show context menu for a node."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Edit Properties", command=lambda: self.parent.edit_node_properties(node))
        menu.add_command(label="Add Deviation", command=lambda: self.parent.add_deviation(node))
        menu.add_command(label="Manage Deviations", command=lambda: self.parent.manage_deviations_for_node(node))
        menu.add_separator()
        menu.add_command(label="Delete Node", command=lambda: self.delete_node(node))
        menu.post(x, y)
    
    def delete_node(self, node):
        """Delete a node."""
        if messagebox.askyesno("Delete Node", f"Delete node '{node.name}'?"):
            self.hazop_data.remove_node(node)
            self.selected_node = None
            self.render_page()
            if self.parent:
                self.parent.on_node_deselected()
    
    def go_to_page(self, page_number):
        """Go to a specific page."""
        if 0 <= page_number < self.total_pages:
            self.current_page = page_number
            # Reset pan when changing pages, but keep zoom level
            self.pan_x = 0
            self.pan_y = 0
            # Reset PDF page dimensions for new page
            self.pdf_page_width = 0
            self.pdf_page_height = 0
            if self.zoom_level == 1.0:
                self.fit_to_window = True
            self.render_page()
            # Ensure focus for keyboard events
            self.focus_set()
    
    def handle_page_up(self, event=None):
        """Handle PageUp key press."""
        if self.doc:
            self.prev_page()
        return "break"  # Prevent event propagation
    
    def handle_page_down(self, event=None):
        """Handle PageDown key press."""
        if self.doc:
            self.next_page()
        return "break"  # Prevent event propagation
    
    def next_page(self):
        """Go to next page."""
        if self.current_page < self.total_pages - 1:
            self.go_to_page(self.current_page + 1)
        # Ensure focus for keyboard events
        self.focus_set()
    
    def prev_page(self):
        """Go to previous page."""
        if self.current_page > 0:
            self.go_to_page(self.current_page - 1)
        # Ensure focus for keyboard events
        self.focus_set()


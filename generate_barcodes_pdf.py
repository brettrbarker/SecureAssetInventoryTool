"""
Generate a PDF with barcodes formatted for Avery 5160 labels.
Takes 3 input strings: barcode_value (used for barcode + printed above), label2, label3 (printed below as "label2 | label3")
"""
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER

from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image


class BarcodeGenerator:
    """Generate barcodes formatted for Avery 5160 labels."""
    
    # Avery 5160 specifications (30 labels per sheet, 3 columns x 10 rows)
    LABELS_PER_ROW = 3
    LABELS_PER_COL = 10
    LABELS_PER_PAGE = 30
    
    # Label dimensions in inches (Avery 5160: 2.625" x 1")
    LABEL_WIDTH = 2.625 * inch
    LABEL_HEIGHT = 1.0 * inch
    
    # Page margins and spacing
    LEFT_MARGIN = 0.1875 * inch  # Avery 5160 left margin
    TOP_MARGIN = 0.5 * inch      # Avery 5160 top margin
    LABEL_SPACING_X = 0.125 * inch  # Horizontal spacing between labels
    LABEL_SPACING_Y = 0.0 * inch    # Vertical spacing between labels
    
    # Barcode sizing within label
    BARCODE_WIDTH = 2.2 * inch
    BARCODE_HEIGHT = 0.4 * inch
    
    def __init__(self):
        self.styles = getSampleStyleSheet()
    
    def generate_barcode_image(self, barcode_value):
        """Generate barcode image in memory."""
        try:
            fp = io.BytesIO()
            # Configure barcode with appropriate sizing
            Code128(barcode_value, writer=ImageWriter()).write(
                fp, 
                options={
                    "module_height": 8.0,
                    "font_size": 1,  # Minimum font size (we'll overlay our own text)
                    "text_distance": 0,
                    "write_text": False,  # Don't write text in barcode image
                    "quiet_zone": 2.0
                }
            )
            fp.seek(0)
            return Image.open(fp)
        except Exception as e:
            print(f"Error generating barcode for '{barcode_value}': {e}")
            return None
    
    def draw_label(self, c, x, y, barcode_value, label2, label3):
        """Draw a single label with barcode and text."""
        # Calculate center positions
        label_center_x = x + (self.LABEL_WIDTH / 2)
        
        # Tighter spacing - position for barcode value text (above barcode)
        text_above_y = y + self.LABEL_HEIGHT - 0.25 * inch
        
        # Position for barcode image (moved down to avoid covering text)
        barcode_y = y + self.LABEL_HEIGHT - 0.65 * inch
        barcode_x = x + (self.LABEL_WIDTH - self.BARCODE_WIDTH) / 2
        
        # Position for text below barcode (closer to barcode)
        text_below_y = y + 0.25 * inch
        
        # Draw barcode value text above barcode (centered)
        c.setFont("Helvetica-Bold", 9)
        text_width = c.stringWidth(barcode_value, "Helvetica-Bold", 9)
        c.drawString(label_center_x - text_width/2, text_above_y, barcode_value)
        
        # Generate and draw barcode
        barcode_img = self.generate_barcode_image(barcode_value)
        if barcode_img:
            c.drawInlineImage(barcode_img, barcode_x, barcode_y, 
                            width=self.BARCODE_WIDTH, height=self.BARCODE_HEIGHT)
        
        # Draw label2 | label3 below barcode (centered)
        c.setFont("Helvetica", 8)
        bottom_text = f"{label2} | {label3}"
        text_width = c.stringWidth(bottom_text, "Helvetica", 8)
        c.drawString(label_center_x - text_width/2, text_below_y, bottom_text)
        
        # Optional: Draw label border for debugging/alignment (comment out for production)
        # c.setStrokeColor("lightgray")
        # c.rect(x, y, self.LABEL_WIDTH, self.LABEL_HEIGHT, fill=0)
    
    def calculate_label_position(self, label_index):
        """Calculate x, y position for a label based on its index."""
        page_label_index = label_index % self.LABELS_PER_PAGE
        row = page_label_index // self.LABELS_PER_ROW
        col = page_label_index % self.LABELS_PER_ROW
        
        x = self.LEFT_MARGIN + col * (self.LABEL_WIDTH + self.LABEL_SPACING_X)
        y = letter[1] - self.TOP_MARGIN - (row + 1) * self.LABEL_HEIGHT - row * self.LABEL_SPACING_Y
        
        return x, y
    
    def generate_pdf(self, labels_data, filename="asset_labels.pdf"):
        """
        Generate PDF with barcode labels.
        
        Args:
            labels_data: List of tuples (barcode_value, label2, label3)
            filename: Output PDF filename
        """
        c = canvas.Canvas(filename, pagesize=letter)
        
        for index, (barcode_value, label2, label3) in enumerate(labels_data):
            # Start new page if needed
            if index > 0 and index % self.LABELS_PER_PAGE == 0:
                c.showPage()
            
            # Calculate position for this label
            x, y = self.calculate_label_position(index)
            
            # Draw the label
            self.draw_label(c, x, y, barcode_value, label2, label3)
        
        c.save()
        print(f"PDF '{filename}' created with {len(labels_data)} labels")


def main():
    """Example usage with sample data."""
    # Sample data - replace with your actual data
    sample_labels = [
        ("ASSET001", "Monitor", "Dell P2414H"),
        ("ASSET002", "Computer", "HP EliteDesk"),
        ("ASSET003", "Keyboard", "Logitech K380"),
        ("ASSET004", "Mouse", "Logitech M705"),
        ("ASSET005", "Printer", "HP LaserJet"),
        ("ASSET006", "Switch", "Cisco SG300"),
        ("ASSET007", "Router", "Cisco ISR4331"),
        ("ASSET008", "Laptop", "Dell Latitude"),
        ("ASSET009", "Tablet", "iPad Pro"),
        ("ASSET010", "Phone", "iPhone 13"),
    ]
    
    generator = BarcodeGenerator()
    generator.generate_pdf(sample_labels, "asset_labels.pdf")
    
    print("\nGenerated labels:")
    for barcode_value, label2, label3 in sample_labels:
        print(f"Barcode: {barcode_value}, Text: {label2} | {label3}")


if __name__ == "__main__":
    main()

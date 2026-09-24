from PIL import Image
from openpyxl.drawing.image import Image as ExcelImage

def load_image(filepath):
    """Loads the image for the Gemini AI model."""
    return Image.open(filepath)

def prepare_excel_image(filepath, max_entries):
    """Embeds and scales the image for the Excel output."""
    excel_img = ExcelImage(filepath)
    excel_img.width = 140
    excel_img.height = max(140, max_entries * 25)
    return excel_img
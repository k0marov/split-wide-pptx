import fitz
def create_triptych(input_pdf_path, output_pdf_path):
    """
    Creates a triptych presentation from a PDF with a number of slides divisible by 3.

    Args:
        input_pdf_path (str): The path to the input PDF file.
        output_pdf_path (str): The path to the output PDF file.
    """
    doc = fitz.open(input_pdf_path)
    if len(doc) % 3 != 0:
        print("Error: The number of slides in the input PDF must be a multiple of 3.")
        return

    w = doc.pages().__next__().bound().width
    h = doc.pages().__next__().bound().height

    new_doc = fitz.open()
    for i in range(0, len(doc), 3):
        # Assuming a 16:9 aspect ratio, a common landscape format is 1920x1080
        # We'll create a new page that can hold three of these side-by-side
        new_page = new_doc.new_page(width=w * 3, height=h)

        # Place the three pages onto the new page
        new_page.show_pdf_page(fitz.Rect(0, 0, w, h), doc, i)
        new_page.show_pdf_page(fitz.Rect(w, 0, w * 2, h), doc, i + 1)
        new_page.show_pdf_page(fitz.Rect(w * 2, 0, w * 3, h), doc, i + 2)

    new_doc.save(output_pdf_path)
    print(f"Triptych presentation saved to {output_pdf_path}")


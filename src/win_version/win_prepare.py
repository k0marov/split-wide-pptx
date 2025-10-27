import os
import sys
import json
import win32com.client
from win32com.client import constants as ppt_constants

MIN_TITLE_TEXT_LENGTH = 3
MAX_TITLE_TEXT_LENGTH = 100

def analyze_slide_layout(slide, title_font_size_threshold):
    """
    Determine if a slide is a title slide based on font size threshold.
    A slide is classified as 'title' if it contains any text element with
    font size larger than FONT_SIZE_THRESHOLD.
    """
    if slide.Shapes.Count == 0:
        return "content"

    large_text_found = False

    for shape in slide.Shapes:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            text_frame = shape.TextFrame
            text_range = text_frame.TextRange

            # Check the entire text range first
            if text_range.Font.Size >= title_font_size_threshold:
                text_content = text_range.Text.strip()
                if (len(text_content) >= MIN_TITLE_TEXT_LENGTH and
                        len(text_content) <= MAX_TITLE_TEXT_LENGTH):
                    large_text_found = True
                    break

            # If text range has multiple paragraphs with different formatting,
            # check each paragraph individually
            if text_range.Paragraphs().Count > 1:
                for i in range(1, text_range.Paragraphs().Count + 1):
                    paragraph = text_range.Paragraphs(i)
                    if paragraph.Font.Size >= title_font_size_threshold:
                        text_content = paragraph.Text.strip()
                        if (len(text_content) >= MIN_TITLE_TEXT_LENGTH and
                                len(text_content) <= MAX_TITLE_TEXT_LENGTH):
                            large_text_found = True
                            break
                if large_text_found:
                    break

    return "title" if large_text_found else "content"


def get_shape_type(shape):
    """
    Get the type of shape as string
    """
    try:
        if shape.Type == ppt_constants.msoPlaceholder:
            return "Placeholder"
        elif shape.Type == ppt_constants.msoTextBox:
            return "TextBox"
        elif shape.Type == ppt_constants.msoAutoShape:
            return "AutoShape"
        else:
            return f"Type_{shape.Type}"
    except:
        return "Unknown"


def triplicate_non_title_slides(input_pptx, output_pptx, title_font_size_threshold: float):
    """
    Creates a new presentation where only NON-TITLE slides are triplicated,
    and outputs JSON mapping showing slide types and clone indices.

    Args:
        input_pptx (str): Path to input PowerPoint file
        output_pptx (str): Path where output will be saved

    Returns:
        dict: JSON-compatible mapping of slide information
    """
    # Initialize PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = False  # Run in background

    slide_mapping = {}

    try:
        # Open the source presentation
        source_pres = powerpoint.Presentations.Open(os.path.abspath(input_pptx))

        # Create a new presentation for output
        output_pres = powerpoint.Presentations.Add()

        output_slide_index = 0

        print(f"Analyzing slides with font size threshold: {title_font_size_threshold}pt")
        print("-" * 50)

        # Process each slide in the source
        for i in range(1, source_pres.Slides.Count + 1):
            source_slide = source_pres.Slides.Item(i)

            # Get largest text element info for analysis
            # Analyze slide type based on font size
            slide_type = analyze_slide_layout(source_slide, title_font_size_threshold)

            if slide_type == "title":
                # Copy title slide as-is (no triplication)
                source_slide.Copy()
                output_pres.Slides.Paste()
                output_slide_index += 1

                slide_mapping[output_slide_index] = {
                    "original_slide_number": i,
                    "type": "title",
                    "clone_index": None,
                }

            else:
                # Triplicate non-title slides
                for clone_idx in range(3):
                    source_slide.Copy()
                    output_pres.Slides.Paste()
                    output_slide_index += 1

                    slide_mapping[output_slide_index] = {
                        "original_slide_number": i,
                        "type": "content",
                        "clone_index": clone_idx + 1,  # 1, 2, or 3
                    }

        # Save the result
        output_pres.SaveAs(os.path.abspath(output_pptx))
        print(f"\nSuccessfully created {output_pptx} with triplicated non-title slides")

        return slide_mapping

    except Exception as e:
        print(f"Error: {str(e)}")
        return None
    finally:
        # Clean up
        if 'source_pres' in locals():
            source_pres.Close()
        if 'output_pres' in locals():
            output_pres.Close()
        powerpoint.Quit()


def save_slide_mapping_to_json(mapping, json_path):
    """
    Save slide mapping to JSON file
    """
    try:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)
        print(f"Slide mapping saved to {json_path}")
        return True
    except Exception as e:
        print(f"Error saving JSON: {str(e)}")
        return False


def classify_and_triplicate(input_path, output_path, title_font_size_threshold) -> dict:
    # Generate output file names
    base_name = os.path.splitext(output_path)[0]
    json_output_path = f"{base_name}_mapping.json"

    # Process the presentation
    slide_mapping = triplicate_non_title_slides(input_path, output_path, title_font_size_threshold)
    if slide_mapping:
        return slide_mapping
    else:
        print("Failed to process the presentation")
        sys.exit(1)

if __name__ == '__main__':
    FONT_SIZE_THRESHOLD = 40
    if len(sys.argv) != 3:
        print("Usage: python triplicate_pptx.py <input.pptx> <output.pptx>")
        print("Example: python triplicate_pptx.py presentation.pptx triplicated.pptx")
        print("This will create:")
        print("  - triplicated.pptx: Output PowerPoint with triplicated non-title slides")
        print("  - slide_mapping.json: JSON file with slide mapping information")
        print(f"Configuration: Font size threshold = {FONT_SIZE_THRESHOLD}pt")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"Error: Input file not found - {input_path}")
        sys.exit(1)

    if not input_path.lower().endswith('.pptx'):
        print("Error: Input file must be a .pptx file")
        sys.exit(1)
    classify_and_triplicate(input_path, output_path, FONT_SIZE_THRESHOLD)

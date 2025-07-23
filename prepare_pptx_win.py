import os
import sys
import win32com.client
from win32com.client import constants as ppt_constants

def triplicate_slides(input_pptx, output_pptx):
    """
    Creates a new presentation where each slide from the input
    is replaced with 3 copies of itself in the output.
    
    Args:
        input_pptx (str): Path to input PowerPoint file
        output_pptx (str): Path where output will be saved
    """
    # Initialize PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = False  # Run in background
    
    try:
        # Open the source presentation
        source_pres = powerpoint.Presentations.Open(os.path.abspath(input_pptx))
        
        # Create a new presentation for output
        output_pres = powerpoint.Presentations.Add()
        
        # Process each slide in the source
        for i in range(1, source_pres.Slides.Count + 1):
            source_slide = source_pres.Slides.Item(i)
            
            # Make 3 copies of the slide
            for _ in range(3):
                source_slide.Copy()
                output_pres.Slides.Paste()
        
        # Save the result
        output_pres.SaveAs(os.path.abspath(output_pptx))
        print(f"Successfully created {output_pptx} with triplicated slides")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return False
    finally:
        # Clean up
        if 'source_pres' in locals():
            source_pres.Close()
        if 'output_pres' in locals():
            output_pres.Close()
        powerpoint.Quit()
    return True

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python triplicate_pptx.py <input.pptx> <output.pptx>")
        print("Example: python triplicate_pptx.py presentation.pptx triplicated.pptx")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(input_path):
        print(f"Error: Input file not found - {input_path}")
        sys.exit(1)
    
    if not input_path.lower().endswith('.pptx'):
        print("Error: Input file must be a .pptx file")
        sys.exit(1)
    
    success = triplicate_slides(input_path, output_path)
    sys.exit(0 if success else 1)

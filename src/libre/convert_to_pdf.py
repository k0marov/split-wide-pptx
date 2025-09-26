#!/usr/bin/env python3
import asyncio
import sys
import os


async def convert_pptx_to_pdf(input_file, output_file):
    """
    Convert PPTX to PDF using LibreOffice
    """
    try:
        # Check if input file exists
        if not os.path.exists(input_file):
            print(f"Error: Input file '{input_file}' not found.")
            return False

        # Ensure output directory exists
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # LibreOffice command for conversion
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to',
            'pdf',
            '--outdir',
            os.path.dirname(output_file) or '.',
            input_file
        ]

        print(f"Converting {input_file} to PDF...")
        # result = subprocess.run(cmd, capture_output=True, text=True)
        process = await asyncio.create_subprocess_shell(' '.join(cmd), stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
        returncode = await process.wait()
        stdin, stderr = await process.communicate()

        if returncode == 0:
            # LibreOffice creates output in the specified directory with the same name but .pdf extension
            expected_output = os.path.join(
                os.path.dirname(output_file) or '.',
                os.path.splitext(os.path.basename(input_file))[0] + '.pdf'
            )

            # Rename to the desired output filename
            if expected_output != output_file and os.path.exists(expected_output):
                os.rename(expected_output, output_file)
                print(f"Successfully converted to {output_file}")
            else:
                print(f"Successfully converted to {expected_output}")

            return True
        else:
            print(f"Error during conversion: {stderr}")
            return False

    except Exception as e:
        print(f"Error: {str(e)}")
        return False


async def main():
    if len(sys.argv) != 3:
        print("Usage: python pptx_to_pdf.py <input.pptx> <output.pdf>")
        print("Example: python pptx_to_pdf.py presentation.pptx output.pdf")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    # Ensure output has .pdf extension
    if not output_file.lower().endswith('.pdf'):
        output_file += '.pdf'

    success = await convert_pptx_to_pdf(input_file, output_file)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    asyncio.run(main())
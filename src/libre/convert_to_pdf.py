#!/usr/bin/env python3
import asyncio
import sys
import os
import src.config as config
import subprocess


async def convert_pptx_to_pdf_windows(input_file, output_file):
    """
    Convert PPTX to PDF using PowerPoint (Windows)
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

        # PowerShell script to convert using PowerPoint
        ps_script = f"""
        $powerpoint = New-Object -ComObject PowerPoint.Application
        $powerpoint.Visible = $false

        try {{
            $presentation = $powerpoint.Presentations.Open("{os.path.abspath(input_file)}")
            $presentation.SaveAs("{os.path.abspath(output_file)}", 32)  # 32 = ppSaveAsPDF
            $presentation.Close()
            Write-Output "Successfully converted to {output_file}"
            exit 0
        }}
        catch {{
            Write-Error "Error during conversion: $_"
            exit 1
        }}
        finally {{
            $powerpoint.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
        }}
        """

        print(f"Converting {input_file} to PDF using PowerPoint...")

        # Run PowerShell script
        cmd = [
            'powershell',
            '-ExecutionPolicy', 'Bypass',
            '-Command', ps_script
        ]

        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )

        stdout, stderr = await process.communicate()

        if process.returncode == 0:
            print(f"Successfully converted to {output_file}")
            return True
        else:
            print(f"Error during conversion: {stderr.decode() if stderr else 'Unknown error'}")
            return False

    except Exception as e:
        print(f"Error: {str(e)}")
        return False


async def convert_pptx_to_pdf_linux(input_file, output_file):
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

        print(f"Converting {input_file} to PDF using LibreOffice...")
        process = await asyncio.create_subprocess_exec(*cmd, stdout=asyncio.subprocess.PIPE,
                                                       stderr=asyncio.subprocess.PIPE)
        stdout, stderr = await process.communicate()

        if process.returncode == 0:
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
            print(f"Error during conversion: {stderr.decode() if stderr else 'Unknown error'}")
            return False

    except Exception as e:
        print(f"Error: {str(e)}")
        return False


async def convert_pptx_to_pdf(input_file, output_file):
    """
    Convert PPTX to PDF using the appropriate method based on ALGORITHM_TYPE
    """
    algorithm_type = config.ALGORITHM_TYPE

    if algorithm_type == 'windows':
        return await convert_pptx_to_pdf_windows(input_file, output_file)
    else:
        return await convert_pptx_to_pdf_linux(input_file, output_file)


async def main():
    if len(sys.argv) != 3:
        print("Usage: python pptx_to_pdf.py <input.pptx> <output.pdf>")
        print("Example: python pptx_to_pdf.py presentation.pptx output.pdf")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    # Ensure input has .pptx extension
    if not input_file.lower().endswith('.pptx'):
        print("Error: Input file must be a .pptx file")
        sys.exit(1)

    # Ensure output has .pdf extension
    if not output_file.lower().endswith('.pdf'):
        output_file += '.pdf'

    success = await convert_pptx_to_pdf(input_file, output_file)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    asyncio.run(main())
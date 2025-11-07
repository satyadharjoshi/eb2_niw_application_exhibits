#!/bin/bash
input="ds_v7_temp.pdf"
output="ds_v7.docx"
tempdir="temp_pages"
repaired_pdf="ds_v7_repaired.pdf"

echo "🔄 Attempting to repair PDF..."
rm -rf "$tempdir"
mkdir -p "$tempdir"

# Method 1: Use Ghostscript to repair PDF
if command -v gs &> /dev/null; then
    echo "Repairing PDF with Ghostscript..."
    gs -o "$repaired_pdf" -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress "$input" 2>/dev/null
    if [ $? -eq 0 ] && [ -f "$repaired_pdf" ]; then
        input="$repaired_pdf"
        echo "✅ PDF repaired successfully"
    else
        echo "❌ Ghostscript repair failed"
    fi
else
    echo "Ghostscript not installed. Install with:"
    echo "  sudo apt install ghostscript  # Ubuntu/Debian"
    echo "  brew install ghostscript      # macOS"
fi

# Get number of pages from repaired PDF
pages=$(pdfinfo "$input" 2>/dev/null | grep "Pages:" | awk '{print $2}')
echo "Total pages in repaired PDF: $pages"

if [ -z "$pages" ] || [ "$pages" -eq 0 ]; then
    echo "❌ Could not read PDF page count. Trying alternative method..."
    # Try to get page count using Python as fallback
    pages=$(python3 -c "
import sys
try:
    from pypdf import PdfReader
    reader = PdfReader('$input')
    print(len(reader.pages))
except:
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader('$input')
        print(len(reader.pages))
    except:
        print('0')
")
fi

echo "Processing $pages pages..."

# Convert using Python with error handling
python3 - <<EOF
import sys
import os
from pdf2docx import Converter
import glob

def convert_pdf_safely(input_pdf, temp_dir, total_pages):
    """Convert PDF to DOCX with robust error handling"""
    
    # Method 1: Try converting the entire PDF at once
    try:
        print("Attempting full PDF conversion...")
        cv = Converter(input_pdf)
        cv.convert('${output}', start=0, end=None)
        cv.close()
        print("✅ Full conversion successful!")
        return True
    except Exception as e:
        print(f"Full conversion failed: {e}")
    
    # Method 2: Convert page by page with individual error handling
    successful_pages = 0
    for page_num in range(1, total_pages + 1):
        try:
            output_file = f"{temp_dir}/page_{page_num}.docx"
            print(f"Converting page {page_num}...")
            
            cv = Converter(input_pdf)
            cv.convert(output_file, pages=[page_num])
            cv.close()
            
            if os.path.exists(output_file) and os.path.getsize(output_file) > 1000:
                successful_pages += 1
                print(f"✅ Page {page_num} converted successfully")
            else:
                print(f"⚠️  Page {page_num} produced small/empty file")
                
        except Exception as e:
            print(f"❌ Page {page_num} failed: {str(e)[:100]}...")
            continue
    
    return successful_pages

# Execute conversion
success = convert_pdf_safely('$input', '$tempdir', $pages)

if success is True:
    print("🎉 Full PDF conversion completed!")
    sys.exit(0)
elif success > 0:
    print(f"📄 {success} out of $pages pages converted successfully")
    
    # Merge the successful pages
    try:
        from docx import Document
        files = sorted(glob.glob("$tempdir/page_*.docx"))
        if files:
            print(f"Merging {len(files)} converted pages...")
            merged = Document(files[0])
            for f in files[1:]:
                subdoc = Document(f)
                merged.add_page_break()
                for element in subdoc.element.body:
                    merged.element.body.append(element)
            merged.save('$output')
            print(f"✅ Merge completed: $output")
        else:
            print("❌ No files to merge")
    except Exception as e:
        print(f"❌ Merge failed: {e}")
else:
    print("❌ No pages could be converted")
    sys.exit(1)
EOF

# Clean up
if [ -f "$repaired_pdf" ]; then
    rm "$repaired_pdf"
fi

echo "Script execution completed"

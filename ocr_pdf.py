#!/usr/bin/env python3
"""
Use macOS Vision framework via pyobjc to OCR a PDF page by page.
"""
import sys
import os
import subprocess
import tempfile

def ocr_page_with_vision(image_path):
    """Use macOS Vision framework to OCR an image."""
    script = f'''
import Vision
import Quartz
import objc
from Foundation import NSURL, NSString
import threading

result_text = []
done_event = threading.Event()

def handler(request, error):
    if error:
        print(f"Error: {{error}}")
        done_event.set()
        return
    observations = request.results()
    for obs in observations:
        result_text.append(obs.topCandidates_(1)[0].string())
    done_event.set()

url = NSURL.fileURLWithPath_("{image_path}")
image_source = Quartz.CGImageSourceCreateWithURL(url, None)
image = Quartz.CGImageSourceCreateImageAtIndex(image_source, 0, None)

request = Vision.VNRecognizeTextRequest.alloc().initWithCompletionHandler_(handler)
request.setRecognitionLanguages_(["zh-Hans", "zh-Hant", "en-US"])
request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)

handler_obj = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(image, {{}})
handler_obj.performRequests_error_([request], None)

done_event.wait(timeout=30)
print("\\n".join(result_text))
'''
    result = subprocess.run(
        ['python3', '-c', script],
        capture_output=True, text=True, timeout=60
    )
    if result.returncode == 0:
        return result.stdout
    else:
        return f"OCR error: {result.stderr[:200]}"

def pdf_page_to_image(pdf_path, page_num, output_path):
    """Convert a PDF page to PNG using sips/quartz."""
    script = f'''
import Quartz
from Foundation import NSURL
import objc

pdf_url = NSURL.fileURLWithPath_("{pdf_path}")
pdf_doc = Quartz.PDFDocument.alloc().initWithURL_(pdf_url)
page = pdf_doc.pageAtIndex_({page_num})

bounds = page.boundsForBox_(Quartz.kPDFDisplayBoxMediaBox)
width = int(bounds.size.width * 2)  # 2x for better OCR
height = int(bounds.size.height * 2)

color_space = Quartz.CGColorSpaceCreateDeviceRGB()
context = Quartz.CGBitmapContextCreate(
    None, width, height, 8, 0, color_space,
    Quartz.kCGImageAlphaPremultipliedLast
)

Quartz.CGContextSetFillColorWithColor(context, Quartz.CGColorGetConstantColor("white"))
Quartz.CGContextFillRect(context, Quartz.CGRectMake(0, 0, width, height))
Quartz.CGContextScaleCTM(context, 2.0, 2.0)

page.drawWithBox_toContext_(Quartz.kPDFDisplayBoxMediaBox, context)

image = Quartz.CGBitmapContextCreateImage(context)
dest_url = NSURL.fileURLWithPath_("{output_path}")
dest = Quartz.CGImageDestinationCreateWithURL(dest_url, "public.png", 1, None)
Quartz.CGImageDestinationAddImage(dest, image, None)
Quartz.CGImageDestinationFinalize(dest)
print("ok")
'''
    result = subprocess.run(
        ['python3', '-c', script],
        capture_output=True, text=True, timeout=30
    )
    return result.returncode == 0

def ocr_pdf(pdf_path, start_page=0, end_page=None):
    """OCR all pages of a PDF and return extracted text."""
    # First get page count
    count_script = f'''
import Quartz
from Foundation import NSURL
pdf_url = NSURL.fileURLWithPath_("{pdf_path}")
pdf_doc = Quartz.PDFDocument.alloc().initWithURL_(pdf_url)
print(pdf_doc.pageCount())
'''
    result = subprocess.run(['python3', '-c', count_script], capture_output=True, text=True)
    page_count = int(result.stdout.strip())
    
    if end_page is None:
        end_page = page_count
    
    print(f"PDF has {page_count} pages, OCR-ing pages {start_page+1} to {end_page}", file=sys.stderr)
    
    all_text = {}
    with tempfile.TemporaryDirectory() as tmpdir:
        for page_num in range(start_page, end_page):
            img_path = os.path.join(tmpdir, f"page_{page_num:03d}.png")
            print(f"Processing page {page_num+1}/{end_page}...", file=sys.stderr)
            
            if pdf_page_to_image(pdf_path, page_num, img_path):
                text = ocr_page_with_vision(img_path)
                all_text[page_num+1] = text
            else:
                all_text[page_num+1] = "(failed to render page)"
    
    return all_text

if __name__ == "__main__":
    pdf_path = "/Users/lin/Documents/Code/music_slides/八下  培训用书.pdf"
    start = int(sys.argv[1]) - 1 if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    
    results = ocr_pdf(pdf_path, start, end)
    
    output_file = f"/tmp/ocr_output_{start+1}_{end or 'end'}.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        for page_num, text in sorted(results.items()):
            f.write(f"\n{'='*60}\n页面 {page_num}\n{'='*60}\n")
            f.write(text)
            f.write("\n")
    
    print(f"Output saved to {output_file}")
    
    # Also print to stdout
    for page_num, text in sorted(results.items()):
        print(f"\n{'='*60}")
        print(f"页面 {page_num}")
        print('='*60)
        print(text)

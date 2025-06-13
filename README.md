# Adding watermark using spire.doc

https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Watermark/Python-Insert-Watermarks-in-Word.html

from spire.doc import *
from spire.doc.common import *

# Create a Document object
document = Document()

# Load a Word document
document.LoadFromFile("test.docx")

# Create a PictureWatermark object
picture = PictureWatermark()

# Set the format of the picture watermark
picture.SetPicture("logo.png")
picture.Scaling = 100
picture.IsWashout = False

# Add the image watermark to document
document.Watermark = picture

#Save the result document
document.SaveToFile("Output/ImageWatermark.docx", FileFormat.Docx)
document.Close()


# watermarking to word doc

how to add watermark to word document in python

ChatGPT said:
To add a watermark to a Word document using Python, you can use the python-docx library. However, python-docx has limited support for advanced features like watermarks, which are often implemented as WordArt or headers/footers with background images or text.

Here’s a common workaround: insert watermark text in the header using a textbox shape, which gives a watermark-like effect.

🔧 Steps:
Install python-docx:

bash
Copy
Edit
pip install python-docx
Use this sample script:

python
Copy
Edit
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

def add_watermark(doc, watermark_text="CONFIDENTIAL"):
    section = doc.sections[0]
    header = section.header
    paragraph = header.paragraphs[0]

    # Create watermark using WordArt-like VML shape
    watermark = OxmlElement('w:pict')
    shape = OxmlElement('v:shape')
    shape.set(qn('id'), 'WordArtWatermark')
    shape.set(qn('style'), 'position:absolute; width:468pt; height:117pt')
    shape.set(qn('type'), '#_x0000_t136')  # WordArt style
    shape.set(qn('rotation'), '315')  # Diagonal text
    shape.set(qn('fillcolor'), '#d3d3d3')  # Light gray
    shape.set(qn('stroked'), 'f')  # No border

    textbox = OxmlElement('v:textpath')
    textbox.set(qn('style'), "font-family:Calibri; font-size:1pt")
    textbox.set(qn('string'), watermark_text)

    shape.append(textbox)
    watermark.append(shape)

    paragraph._p.append(watermark)

# Create a new doc and apply watermark
doc = Document()
doc.add_paragraph("This is a sample document with a watermark.")
add_watermark(doc)
doc.save("watermarked.docx")
⚠️ Limitations:
Word’s watermark feature isn’t officially supported by python-docx, so this is a workaround using XML.

This may not render perfectly across all Word versions or viewers.

✅ Alternative:
If you want image-based watermarks, insert an image in the header/footer.

Would you prefer an image watermark example instead?




You said:
Is there any other python package?

ChatGPT said:
Yes, aside from python-docx, there are a few other Python libraries that can help you add watermarks to Word documents (with different capabilities):

✅ 1. docxcompose
Best for: Merging and adding headers/footers including watermark-like features
Install:
bash
Copy
Edit
pip install docxcompose
Example (with a watermark template):
Create a Word file watermark.docx with just your watermark text in the header or footer.

Use this script to apply it to another document:

python
Copy
Edit
from docx import Document
from docxcompose.composer import Composer

# Open the base and watermark documents
main_doc = Document("original.docx")
watermark_doc = Document("watermark.docx")

composer = Composer(main_doc)
composer.append(watermark_doc)  # merges content

composer.save("watermarked_output.docx")
This doesn’t “layer” a watermark over each page but can be used creatively with watermark templates in headers.

✅ 2. Aspose.Words for Python
Best for: Advanced document processing (commercial product with free trial)
Install:
bash
Copy
Edit
pip install aspose-words
Example (text watermark):
python
Copy
Edit
import aspose.words as aw

doc = aw.Document("input.docx")
builder = aw.DocumentBuilder(doc)

# Create watermark shape
watermark = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
watermark.text_path.text = "CONFIDENTIAL"
watermark.width = 500
watermark.height = 100
watermark.rotation = -40
watermark.fill.color = aw.Color.gray
watermark.stroke_color = aw.Color.gray

# Position it
watermark.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
watermark.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
watermark.wrap_type = aw.drawing.WrapType.NONE
watermark.vertical_alignment = aw.drawing.VerticalAlignment.CENTER
watermark.horizontal_alignment = aw.drawing.HorizontalAlignment.CENTER

# Insert watermark into all headers
for section in doc.sections:
    header = section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY]
    header.child_nodes.insert(0, watermark.deep_clone(True))

doc.save("aspose_watermarked.docx")
✅ This is the most reliable and professional method — but it's not free for commercial use.

✅ 3. PyWin32 (Windows only)
Best for: Using Microsoft Word directly via COM automation (if Word is installed)
bash
Copy
Edit
pip install pywin32
python
Copy
Edit
import win32com.client as win32

word = win32.gencache.EnsureDispatch('Word.Application')
doc = word.Documents.Open(r"C:\path\to\your\document.docx")

# Add watermark
word.ActiveDocument.Sections(1).Headers(1).Range.Select()
word.Selection.HeaderFooter.Shapes.AddTextEffect(
    PresetTextEffect=0, Text="CONFIDENTIAL", FontName="Arial", FontSize=36,
    FontBold=True, FontItalic=False, Left=100, Top=100
).Select()
word.Selection.ShapeRange.Fill.ForeColor.RGB = 192 + (192 << 8) + (192 << 16)  # Light gray
word.Selection.ShapeRange.Rotation = 315
word.Selection.ShapeRange.LockAspectRatio = True

doc.SaveAs(r"C:\path\to\output.docx")
doc.Close()
word.Quit()
⚠️ Works only on Windows with MS Word installed.

📝 Summary:
Library	Platform	Watermark Type	Free?
python-docx	Cross-platform	Workaround via headers	✅ Yes
docxcompose	Cross-platform	Template-based merging	✅ Yes
Aspose.Words	Cross-platform	True watermark support	⚠️ Paid (trial)
PyWin32	Windows only	MS Word automation	✅ Yes

Let me know which method fits your needs best and I can help you with a complete working example

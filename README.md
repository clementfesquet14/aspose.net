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

# Download all packages for spire-doc

https://huggingface.co/tensorwa/spire-doc-pack/tree/main

From above link, you can download all package files for spire doc

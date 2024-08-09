from flask import Flask,request
import aspose.pdf as file
from PIL import Image

app = Flask(__name__)

def pdf_to_docx(dir_input,dir_out,name):
    income_pdf = dir_input
    outcome_pdf= dir_out+name+".docx"
    doc = file.Document(income_pdf)
    
    options = file.DocSaveOptions()
    options.format = file.DocSaveOptions.DocFormat.DOC_X
    options.mode = file.DocSaveOptions.RecognitionMode.FLOW
    options.relative_horizontal_proximity = 2.5
    options.recognize_bullets = True
    return doc.save(outcome_pdf,options)    

def docx_to_pdf(dir_input, dir_out, name):
    income_docx = dir_input
    outcome_pdf = dir_out + name + ".pdf"
    
    # Load the DOCX document
    doc = file.Document(income_docx)
    
    # Define the options for PDF saving
    options = file.PdfSaveOptions()
    
    # Save the document as PDF
    doc.save(outcome_pdf, options)
    return outcome_pdf

def image_conversion(dir_input,name,target_ext):
    dir_out = "convertor\File\Saver\IMAGE"
    img = Image.open(f"{dir_input}")
    img.save(f"{dir_out}\{name}.{target_ext}")
    
    
@app.route('/convert',methods = ['GET'])
def convert():
    category =  request.args.get('category')
    name = request.args.get('name')
    ext_to = request.args.get('to')
    dir_from = request.args.get('pathway')
    
    match category:
        case "Image":
            image_conversion(dir_input=dir_from,name=name,target_ext=ext_to)
        case "Audio":
            pass
        case "Document":
            if ext_to == ""
from flask import Flask, request, jsonify
import aspose.pdf as file
from PIL import Image
import os

app = Flask(__name__)

def pdf_to_docx(dir_input, dir_out, name):
    income_pdf = dir_input
    outcome_docx = os.path.join(dir_out, name + ".docx")
    doc = file.Document(income_pdf)
    
    options = file.DocSaveOptions()
    options.format = file.DocSaveOptions.DocFormat.DOC_X
    options.mode = file.DocSaveOptions.RecognitionMode.FLOW
    options.relative_horizontal_proximity = 2.5
    options.recognize_bullets = True
    doc.save(outcome_docx, options)
    return outcome_docx    

def docx_to_pdf(dir_input, dir_out, name):
    income_docx = dir_input
    outcome_pdf = os.path.join(dir_out, name + ".pdf")
    
    doc = file.Document(income_docx)
    
    options = file.PdfSaveOptions()
    
    doc.save(outcome_pdf, options)
    return outcome_pdf

def image_conversion(dir_input, name, target_ext):
    dir_out = "convertor/File/Saver/IMAGE"
    os.makedirs(dir_out, exist_ok=True)
    img = Image.open(dir_input)
    output_path = os.path.join(dir_out, f"{name}.{target_ext}")
    img.save(output_path)
    return output_path

@app.route('/convert', methods=['GET'])
def convert():
    category = request.args.get('category')
    name = request.args.get('name')
    ext_to = request.args.get('to')
    dir_from = request.args.get('pathway')
    dir_out = request.args.get('output_path', '')  # Optional parameter for output directory

    try:
        if category == "Image":
            result = image_conversion(dir_input=dir_from, name=name, target_ext=ext_to)

        elif category == "Document":
            if ext_to == "docx":
                result = pdf_to_docx(dir_input=dir_from, dir_out=dir_out, name=name)
            elif ext_to == "pdf":
                result = docx_to_pdf(dir_input=dir_from, dir_out=dir_out, name=name)
            else:
                return jsonify({"error": "Unsupported document format conversion requested"}), 400
                
        else:
            return jsonify({"error": "Unsupported category"}), 400
        
        return jsonify({"result": result}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

from flask import Flask, request, jsonify
import aspose.words as aw
from pydub import AudioSegment
from PIL import Image
import os

app = Flask(__name__)

def pdf_to_docx(dir_input, dir_out, name):
    income_pdf = dir_input
    outcome_docx = os.path.join(dir_out, name + ".docx")
    doc = aw.Document(income_pdf)
    
    options = aw.saving.DocSaveOptions(aw.SaveFormat.DOCX)
    options.mode = aw.saving.DocSaveOptions.RecognitionMode.TEXT_FLOW
    options.relative_horizontal_proximity = 2.5
    options.recognize_bullets = True
    doc.save(outcome_docx, options)
    return 200    

def docx_to_pdf(dir_input, dir_out, name):
    income_docx = dir_input
    outcome_pdf = os.path.join(dir_out, name + ".pdf")
    
    doc = aw.Document(income_docx)
    
    options = aw.saving.PdfSaveOptions()
    
    doc.save(outcome_pdf, options)
    return 200

def audio_conversion(dir_input, name, target_ext):
    audio = AudioSegment.from_file(dir_input)
    dir_out = "convertor/File/Saver/Audio"
    os.makedirs(dir_out, exist_ok=True)
    
    output_path = os.path.join(dir_out, f"{name}.{target_ext}")
    audio.export(output_path, format=target_ext)
    return 200

def image_conversion(dir_input, name, target_ext):
    dir_out = "convertor/File/Saver/Image"
    os.makedirs(dir_out, exist_ok=True)
    
    img = Image.open(dir_input)
    output_path = os.path.join(dir_out, f"{name}.{target_ext}")
    img.save(output_path)
    return 200

@app.route('/convert', methods=['GET'])
def convert():
    category = request.args.get('category')
    name = request.args.get('name')
    ext_to = request.args.get('to')
    dir_from = request.args.get('pathway')
    dir_out = request.args.get('output_path', '') 

    try:
        if not category or not name or not ext_to or not dir_from:
            return jsonify({"error": "Missing required parameters"}), 400

        match category:
            case "Image":
                result = image_conversion(dir_input=dir_from, name=name, target_ext=ext_to)

            case "Document":
                if ext_to == "docx":
                    result = pdf_to_docx(dir_input=dir_from, dir_out=dir_out, name=name)
                elif ext_to == "pdf":
                    result = docx_to_pdf(dir_input=dir_from, dir_out=dir_out, name=name)
                else:
                    return jsonify({"error": "Unsupported document format conversion requested"}), 400
                
            case "Audio":
                result = audio_conversion(dir_input=dir_from, name=name, target_ext=ext_to)
                
            case _:
                return jsonify({"error": "Unsupported category"}), 400

        return jsonify({"result": result}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

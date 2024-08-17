from flask import Flask, request, jsonify
import aspose.words as aw
import ffmpeg
from pydub import AudioSegment
from pathlib import Path
from PIL import Image
import os

app = Flask(__name__)

def get_downloads_folder():
    if os.name == 'nt':  # Windows
        return str(Path(os.environ['USERPROFILE']) / 'Downloads')
    elif os.name == 'posix':
        return str(Path.home() / 'Downloads')
    else:
        raise EnvironmentError("Unsupported operating system")
    
def generate_unique_filename(dir_path, base_name, extension):
    i = 1
    filename = f"{base_name}.{extension}"
    while os.path.exists(os.path.join(dir_path, filename)):
        filename = f"{base_name}_{i}.{extension}"
        i += 1
    return filename

def pdf_to_docx(dir_input, name):
    try:
        if not os.path.exists(dir_input):
            raise FileNotFoundError(f"File not found: {dir_input}")
        
        dir_out = get_downloads_folder()
        os.makedirs(dir_out, exist_ok=True)
        
        outcome_docx = os.path.join(dir_out, name + ".docx")
        print(f"Saving DOCX to: {outcome_docx}")
        doc = aw.Document(dir_input)
        
        options = aw.saving.DocSaveOptions(aw.SaveFormat.DOCX)
        options.mode = aw.saving.DocSaveOptions.RecognitionMode.TEXT_FLOW
        options.relative_horizontal_proximity = 2.5
        options.recognize_bullets = True
        doc.save(outcome_docx, options)
        print(f"File saved successfully: {outcome_docx}")

        return 200
    except FileNotFoundError as fnf_error:
        print(f"FileNotFoundError: {fnf_error}")
        return 404
    except Exception as e:
        print(f"Error in pdf_to_docx: {str(e)}")
        return 500

   

def docx_to_pdf(dir_input, name):
    income_docx = dir_input
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    outcome_pdf = os.path.join(dir_out, name + ".pdf")
    doc = aw.Document(income_docx)
    options = aw.saving.PdfSaveOptions()
    doc.save(outcome_pdf, options)
    return 200

def txt_to_docx(dir_input,name):
    income_pdf = dir_input
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    outcome_docx = os.path.join(dir_out, name + ".docx")
    doc = aw.Document(income_pdf)
    
    options = aw.saving.DocSaveOptions(aw.SaveFormat.DOCX)
    options.mode = aw.saving.DocSaveOptions.RecognitionMode.TEXT_FLOW
    options.relative_horizontal_proximity = 2.5
    options.recognize_bullets = True
    doc.save(outcome_docx, options)
    return 200

def txt_to_pdf(dir_input,name):
    income_docx = dir_input
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    outcome_pdf = os.path.join(dir_out, name + ".pdf")
    doc = aw.Document(income_docx)
    options = aw.saving.PdfSaveOptions()
    doc.save(outcome_pdf, options)
    return 200

def docx_to_txt(dir_input,name):
    income_txt = dir_input
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    outcome_txt = os.path.join(dir_out, name + ".txt")
    txt = aw.Document(income_txt)
    options = aw.saving.TxtSaveOptions()
    txt.save(outcome_txt,options)
    return 200 

def pdf_to_txt(dir_input,name):
    income_txt = dir_input
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    outcome_txt = os.path.join(dir_out, name + ".txt")
    txt = aw.Document(income_txt)
    options = aw.saving.TxtSaveOptions()
    txt.save(outcome_txt,options)
    return 200

def audio_conversion(dir_input, name, target_ext):
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    audio = AudioSegment.from_file(dir_input)
    output_path = os.path.join(dir_out, f"{name}.{target_ext}")
    audio.export(output_path, format=target_ext)
    return 200

def image_conversion(dir_input, name, target_ext):
    try:
        dir_out = get_downloads_folder()
        os.makedirs(dir_out, exist_ok=True)
        
        img = Image.open(dir_input)
        
        unique_filename = generate_unique_filename(dir_out, name, target_ext)
        output_path = os.path.join(dir_out, unique_filename)
        
        print(f"Saving image to: {output_path}")
        
        if target_ext.lower() in ['jpg', 'jpeg']:
            img = img.convert('RGB')
        
        img.save(output_path)
        return 200
    except Exception as e:
        print(f"Error saving image: {str(e)}")
        return 500

def video_conversion(dir_input, name, target_ext):
    dir_out = get_downloads_folder()
    os.makedirs(dir_out, exist_ok=True)
    output_path = os.path.join(dir_out, f"{name}.{target_ext}")
    ffmpeg.input(dir_input).output(output_path).run()
    return 200


@app.route('/convert', methods=['GET'])
def convert():
    category = request.args.get('category')
    name = request.args.get('name')
    ext_from = request.args.get('from')
    ext_to = request.args.get('to')
    dir_from = request.args.get('pathway')
    
    print(f"Category: {category}, Name: {name}, From: {ext_from}, To: {ext_to}, Pathway: {dir_from}")

    try:
        if not category or not name or not ext_to or not dir_from:
            return jsonify({"error": "Missing required parameters"}), 400

        match category:
            case "Image":
                if ext_to.lower() in ['jpg', 'jpeg', 'png', 'bmp', 'gif']:
                    result = image_conversion(dir_input=dir_from, name=name, target_ext=ext_to)
                else:
                    return jsonify({"error": "Unsupported image format"}), 400
                
            case "Document":
                if ext_to == "DOCX":
                    if ext_from == "PDF":
                        result = pdf_to_docx(dir_input=dir_from, name=name)
                    elif ext_from == "TXT":
                        result = txt_to_docx(dir_input=dir_from, name=name)
                    else:
                        return jsonify({"error": "Unsupported source document format"}), 400
                
                elif ext_to == "PDF":
                    if ext_from == "DOCX":
                        result = docx_to_pdf(dir_input=dir_from, name=name)
                    elif ext_from == "TXT":
                        result = txt_to_pdf(dir_input=dir_from, name=name)
                    else:
                        return jsonify({"error": "Unsupported source document format"}), 400
                
                elif ext_to == "TXT":
                    if ext_from == "PDF":
                        result = pdf_to_txt(dir_input=dir_from, name=name)   
                    elif ext_from == "DOCX":
                        result = docx_to_txt(dir_input=dir_from, name=name)
                    else:
                        return jsonify({"error": "Unsupported source document format"}), 400
                
                else:
                    return jsonify({"error": "Unsupported document format conversion requested"}), 400
            
            case "Video":
                if ext_to.lower() in ['mp4', 'avi', 'mkv', 'mov']:
                    result = video_conversion(dir_input=dir_from, name=name, target_ext=ext_to)
                else:
                    return jsonify({"error": "Unsupported video format"}), 400
                
            case "Audio":
                if ext_to.lower() in ['mp3', 'wav', 'flac', 'aac']:
                    result = audio_conversion(dir_input=dir_from, name=name, target_ext=ext_to)
                else:
                    return jsonify({"error": "Unsupported audio format"}), 400

            case _:
                return jsonify({"error": "Unsupported category"}), 400

        return jsonify({"result": result}), 200
    except Exception as e:
        return jsonify({"error": "SERVER Error"}), 500

if __name__ == '__main__':
    app.run(debug=True)

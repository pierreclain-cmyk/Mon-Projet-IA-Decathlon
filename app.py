from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import os
import shutil

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/appliquer-charte")
async def appliquer_charte(fichier: UploadFile = File(...)):
    # 1. Sauvegarder le fichier envoyé par l'utilisateur
    input_path = f"temp_{fichier.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(fichier.file, buffer)

    # 2. Ouvrir la présentation avec python-pptx
    prs = Presentation(input_path)

    # 3. Appliquer la charte graphique Decathlon à tout le texte existant
    # Bleu Decathlon : RGB(0, 130, 195)
    decathlon_blue = RGBColor(0, 130, 195)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
                
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # Forcer la police (remplacez par la police officielle si besoin)
                    run.font.name = 'Arial' 
                    
                    # Si c'est un titre (généralement gros texte), on le met en bleu
                    if run.font.size and run.font.size > Pt(20):
                        run.font.color.rgb = decathlon_blue

    # 4. Sauvegarder le fichier modifié
    output_path = "Decathlon_Modifie.pptx"
    prs.save(output_path)

    # Nettoyer le fichier temporaire
    os.remove(input_path)

    return FileResponse(
        output_path, 
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
        filename="Decathlon_Modifie.pptx"
    )

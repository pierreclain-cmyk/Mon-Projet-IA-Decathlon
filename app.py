from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from fastapi.responses import HTMLResponse
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
    # 1. Sauvegarder temporairement le fichier
    input_path = f"temp_{fichier.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(fichier.file, buffer)

    # 2. Ouvrir la présentation
    prs = Presentation(input_path)

    # Couleurs de la charte d'après la capture d'écran
    bleu_fond = RGBColor(59, 73, 184) # Le fond bleu indigo
    texte_blanc = RGBColor(255, 255, 255) # Texte blanc
    
    # 3. Appliquer la charte à chaque diapositive
    for slide in prs.slides:
        
        # A. Modifier le fond de la diapositive
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bleu_fond

        # B. Parcourir tous les éléments pour modifier les textes
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
                
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # Appliquer la police "Inter"
                    run.font.name = 'Inter'
                    
                    # Appliquer la couleur blanche au texte
                    run.font.color.rgb = texte_blanc
                    
                    # Si c'est un titre (taille existante > 24), on le met en gras (Semibold)
                    if run.font.size and run.font.size > Pt(24):
                        run.font.bold = True
                    # Si c'est du texte standard (storytelling), on peut forcer autour de 18pt
                    elif run.font.size and run.font.size < Pt(20):
                        run.font.size = Pt(18)

    # 4. Sauvegarder la présentation mise à jour
    output_path = "Decathlon_Inter_Charte.pptx"
    prs.save(output_path)
    os.remove(input_path)

    return FileResponse(
        output_path, 
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
        filename="Decathlon_Inter_Charte.pptx"
    )

@app.get("/")
async def afficher_page_accueil():
    # Lit et affiche votre fichier index.html
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

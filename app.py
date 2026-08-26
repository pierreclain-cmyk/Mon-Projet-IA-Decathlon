from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_CONNECTOR
import os
import re

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Couleurs de la charte Decathlon
BLEU_INDIGO = RGBColor(59, 73, 184)   # #3B49B8
VERT_FLUO = RGBColor(110, 240, 160)   # #6EF0A0 (Métrique / Accent)
TEXTE_BLANC = RGBColor(255, 255, 255) # #FFFFFF

def est_une_metrique(texte: str) -> bool:
    """Détecte si un texte contient un pourcentage, un +, ou un chiffre clé"""
    return bool(re.search(r'(\d+[\.,]?\d*\s*%|\+\d+|\bCible\b)', texte))

@app.post("/api/appliquer-charte")
async def appliquer_charte(fichier: UploadFile = File(...)):
    input_path = f"temp_{fichier.filename}"
    
    try:
        with open(input_path, "wb") as buffer:
            buffer.write(await fichier.read())

        prs = Presentation(input_path)

        for slide in prs.slides:
            # 1. Fond bleu indigo uni sur la slide
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = BLEU_INDIGO

            for shape in slide.shapes:
                
                # 2. Nettoyage et contraste des formes / cartes / rectangles
                if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.RECTANGLE]:
                    # Rendre les conteneurs transparents avec une bordure fine contrastée
                    shape.fill.background() 
                    shape.line.color.rgb = TEXTE_BLANC
                    shape.line.width = Pt(1)

                # 3. Traitement des flèches et connecteurs
                elif shape.shape_type == MSO_SHAPE_TYPE.CONNECTOR or "Arrow" in shape.name:
                    shape.line.color.rgb = VERT_FLUO
                    shape.line.width = Pt(2.5)

                # 4. Traitement intelligent de la typographie
                if shape.has_text_frame:
                    is_title = (shape == slide.shapes.title)
                    
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Inter'
                            texte_net = run.text.strip()

                            # CAS 1 : C'est une métrique (% ou + points) -> VERT FLUO
                            if est_une_metrique(texte_net):
                                run.font.color.rgb = VERT_FLUO
                                run.font.bold = True
                                if run.font.size and run.font.size < Pt(24):
                                    run.font.size = Pt(28) # Donner de l'impact aux chiffres

                            # CAS 2 : Titres principaux -> Blanc imposant
                            elif is_title:
                                run.font.color.rgb = TEXTE_BLANC
                                run.font.size = Pt(32)
                                run.font.bold = True

                            # CAS 3 : Textes d'explication / Storytelling -> Blanc propre
                            else:
                                run.font.color.rgb = TEXTE_BLANC
                                run.font.bold = False

        output_path = "Decathlon_Clean.pptx"
        prs.save(output_path)
        return FileResponse(output_path, filename="Decathlon_Presentation.pptx")
        
    except Exception as e:
        print(f"Erreur : {e}")
        raise HTTPException(status_code=500, detail="Erreur lors du traitement.")
        
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

@app.get("/")
async def afficher_page_accueil():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

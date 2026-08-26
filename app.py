from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image, ImageOps
import io
import os
import shutil

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def appliquer_filtre_image(image_blob):
    """Applique un filtre bleu Decathlon (Duotone) à l'image"""
    # Ouvrir l'image et la convertir en nuances de gris
    img = Image.open(io.BytesIO(image_blob)).convert("L")
    
    # Appliquer le duotone : ombres foncées, tons moyens en Bleu Decathlon, lumières en blanc
    # Bleu Decathlon : (59, 73, 184)
    img_colorized = ImageOps.colorize(
        img, 
        black=(20, 25, 60),      # Bleu très sombre pour les ombres
        white=(255, 255, 255),   # Blanc pour les lumières
        mid=(59, 73, 184)        # Bleu Decathlon pour les tons moyens
    )
    
    out_io = io.BytesIO()
    img_colorized.save(out_io, format="PNG")
    return out_io

@app.post("/api/appliquer-charte")
async def appliquer_charte(fichier: UploadFile = File(...)):
    input_path = f"temp_{fichier.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(fichier.file, buffer)

    prs = Presentation(input_path)
    bleu_fond = RGBColor(59, 73, 184)
    texte_blanc = RGBColor(255, 255, 255)
    
    # Dimensions max de la slide (pour limiter la taille des images)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    for slide in prs.slides:
        # 1. Appliquer la couleur de fond
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bleu_fond

        shapes_to_delete = []

        # 2. Parcours des éléments (Textes et Images)
        for shape in slide.shapes:
            
            # --- GESTION DES IMAGES ---
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # Extraire l'image originale
                image_blob = shape.image.blob
                
                # Appliquer le filtre couleur
                nouvelle_image_io = appliquer_filtre_image(image_blob)
                
                # Récupérer la position et la taille
                left, top = shape.left, shape.top
                width, height = shape.width, shape.height
                
                # Optimisation de la disposition : empêcher le débordement
                if width > slide_width:
                    ratio = slide_width / width
                    width = slide_width
                    height = int(height * ratio)
                    left = 0
                if height > slide_height:
                    ratio = slide_height / height
                    height = slide_height
                    width = int(width * ratio)
                    top = 0
                
                # Ajouter la nouvelle image modifiée au même endroit
                slide.shapes.add_picture(nouvelle_image_io, left, top, width, height)
                
                # Marquer l'ancienne image pour suppression
                shapes_to_delete.append(shape)
            
            # --- GESTION DES TEXTES ---
            elif shape.has_text_frame:
                is_title = shape == slide.shapes.title # Détecte si c'est le titre principal
                
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Inter'
                        run.font.color.rgb = texte_blanc
                        
                        # Hiérarchie des tailles
                        if is_title:
                            run.font.size = Pt(32)
                            run.font.bold = True
                        else:
                            # Si c'était déjà assez gros, on le passe en sous-titre
                            if run.font.size and run.font.size > Pt(20):
                                run.font.size = Pt(24)
                                run.font.bold = True
                            # Sinon, c'est du corps de texte (Storytelling à 18pt)
                            else:
                                run.font.size = Pt(18)
                                run.font.bold = False

        # Supprimer les anciennes images (non filtrées) de la diapositive
        for shape in shapes_to_delete:
            shape._element.getparent().remove(shape._element)

    output_path = "Decathlon_V2_Ameliore.pptx"
    prs.save(output_path)
    os.remove(input_path)

    return FileResponse(
        output_path, 
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
        filename="Decathlon_V2_Ameliore.pptx"
    )

@app.get("/")
async def afficher_page_accueil():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

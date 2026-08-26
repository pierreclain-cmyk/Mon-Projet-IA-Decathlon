from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import os
import shutil
import io
from google import genai # Import de l'API Google Gemini

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/appliquer-charte")
async def appliquer_charte(
    fichier: UploadFile = File(...),
    api_key: str = Form(""),
    image_prompt: str = Form("")
):
    input_path = f"temp_{fichier.filename}"
    
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(fichier.file, buffer)

        prs = Presentation(input_path)
        bleu_fond = RGBColor(59, 73, 184)
        texte_blanc = RGBColor(255, 255, 255)

        # 1. APPLICATION DE LA CHARTE (Textes et fond uniquement)
        for slide in prs.slides:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = bleu_fond

            for shape in slide.shapes:
                if shape.has_text_frame:
                    is_title = shape == slide.shapes.title
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Inter'
                            run.font.color.rgb = texte_blanc
                            
                            if is_title:
                                run.font.size = Pt(32)
                                run.font.bold = True
                            else:
                                if run.font.size and run.font.size > Pt(20):
                                    run.font.size = Pt(24)
                                    run.font.bold = True
                                else:
                                    run.font.size = Pt(18)
                                    run.font.bold = False

        # 2. GÉNÉRATION D'IMAGE IA AVEC GEMINI (Imagen 3)
        if api_key and image_prompt:
            try:
                # Connexion à l'API Google
                client = genai.Client(api_key=api_key)
                
                # On force le style pour correspondre à votre charte
                prompt_optimise = f"Style minimaliste, vecteur, couleurs bleu profond et vert fluo, sport. Sujet : {image_prompt}"
                
                # Appel du modèle Imagen 3
                result = client.models.generate_images(
                    model='imagen-3.0-generate-001',
                    prompt=prompt_optimise,
                    config=dict(
                        number_of_images=1,
                        aspect_ratio="1:1",
                        output_mime_type="image/jpeg"
                    )
                )
                
                # Récupération des données binaires de l'image
                image_bytes = result.generated_images[0].image.image_bytes
                img_stream = io.BytesIO(image_bytes)
                
                # Création d'une nouvelle diapositive vide à la fin
                blank_slide_layout = prs.slide_layouts[6] # Index 6 = diapo vide
                new_slide = prs.slides.add_slide(blank_slide_layout)
                
                # Fond bleu pour la nouvelle diapositive
                new_slide_bg = new_slide.background
                new_slide_bg.fill.solid()
                new_slide_bg.fill.fore_color.rgb = bleu_fond
                
                # Ajout et centrage de l'image IA générée
                new_slide.shapes.add_picture(img_stream, Inches(2), Inches(1), height=Inches(5.5))
                
            except Exception as e:
                print(f"L'image IA (Gemini) n'a pas pu être générée : {e}")
                # Le script ne plante pas si l'API échoue, il renvoie quand même la présentation formatée

        output_path = "Decathlon_IA_Gemini.pptx"
        prs.save(output_path)
        
        return FileResponse(
            output_path, 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
            filename="Decathlon_IA_Gemini.pptx"
        )
        
    except Exception as e:
        print(f"Erreur globale : {e}")
        raise HTTPException(status_code=500, detail="Erreur du serveur lors du formatage.")
        
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

@app.get("/")
async def afficher_page_accueil():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

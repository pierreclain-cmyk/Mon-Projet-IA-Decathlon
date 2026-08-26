from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import os
import shutil
import io
from google import genai

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
    image_prompt: str = Form(""),
    donnees: UploadFile = File(None)  # Le nouveau fichier de données optionnel
):
    input_path = f"temp_{fichier.filename}"
    
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(fichier.file, buffer)

        prs = Presentation(input_path)
        bleu_fond = RGBColor(59, 73, 184)
        texte_blanc = RGBColor(255, 255, 255)

        # 1. APPLICATION EXPERTE DE LA CHARTE TYPOGRAPHIQUE
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
                            # Forcer la police globale
                            run.font.name = 'Inter'
                            run.font.color.rgb = texte_blanc
                            
                            texte_brut = run.text.strip()
                            taille_orig = run.font.size
                            
                            # --- NOUVELLE LOGIQUE TYPOGRAPHIQUE ---
                            if is_title:
                                run.font.size = Pt(32)
                                run.font.bold = True
                                
                            # Si c'est un mot en majuscule (ex: "WATERPROOFING")
                            elif texte_brut.isupper() and len(texte_brut) > 3:
                                run.font.size = Pt(14)
                                run.font.bold = False
                                
                            # Si c'est un gros chiffre/métrique d'origine (ex: "30,000mm")
                            elif taille_orig and taille_orig > Pt(24):
                                run.font.size = Pt(36)
                                run.font.bold = True
                                
                            # Si c'est une petite note (ex: "Minimum", "Externe")
                            elif taille_orig and taille_orig < Pt(14):
                                run.font.size = Pt(12)
                                run.font.bold = False
                                
                            # Sinon, c'est le Storytelling par défaut (ex: votre capture)
                            else:
                                run.font.size = Pt(18)
                                run.font.bold = True # Semibold dans PPT est géré par le bold standard

        # 2. GÉNÉRATION IA AVEC DONNÉES EXTERNES
        # Récupération sécurisée de la clé API depuis le serveur
        api_key = os.getenv("GEMINI_API_KEY")
        
        if api_key and image_prompt:
            try:
                # Lecture des données externes si fournies
                contexte_data = ""
                if donnees:
                    contenu = await donnees.read()
                    # On décode le fichier (CSV, TXT, JSON) en texte lisible par l'IA
                    contexte_data = f"\n\nPrends impérativement en compte ces données pour générer l'image : {contenu.decode('utf-8')[:1000]}"

                client = genai.Client(api_key=api_key)
                
                # Le prompt fusionne vos consignes visuelles, la demande utilisateur, et la donnée
                prompt_optimise = f"Style vectoriel plat, minimaliste. Couleurs strictes : fond bleu indigo, accents vert menthe vif. Sujet : {image_prompt} {contexte_data}"
                
                result = client.models.generate_images(
                    model='imagen-3.0-generate-001',
                    prompt=prompt_optimise,
                    config=dict(number_of_images=1, aspect_ratio="16:9", output_mime_type="image/jpeg")
                )
                
                image_bytes = result.generated_images[0].image.image_bytes
                img_stream = io.BytesIO(image_bytes)
                
                # Ajout de l'image sur une nouvelle diapositive
                blank_slide_layout = prs.slide_layouts[6]
                new_slide = prs.slides.add_slide(blank_slide_layout)
                
                new_slide_bg = new_slide.background
                new_slide_bg.fill.solid()
                new_slide_bg.fill.fore_color.rgb = bleu_fond
                
                # Image en plein écran (ou presque)
                new_slide.shapes.add_picture(img_stream, Inches(1), Inches(1), width=Inches(8))
                
            except Exception as e:
                print(f"Erreur IA : {e}")

        output_path = "Decathlon_Expert.pptx"
        prs.save(output_path)
        
        return FileResponse(
            output_path, 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
            filename="Decathlon_Expert.pptx"
        )
        
    except Exception as e:
        print(f"Erreur : {e}")
        raise HTTPException(status_code=500, detail="Erreur serveur.")
        
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

@app.get("/")
async def afficher_page_accueil():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

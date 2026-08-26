from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import os
import shutil
import io
import yaml
from google import genai

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def charger_design_tokens():
    """Extrait les tokens YAML et la prose Markdown du fichier DESIGN.md"""
    if not os.path.exists("DESIGN.md"):
        # Valeurs par défaut de secours si le fichier n'existe pas
        return {
            "primary": RGBColor(59, 73, 184),
            "text": RGBColor(255, 255, 255),
            "prose": "Style minimaliste bleu indigo et vert menthe"
        }
    
    with open("DESIGN.md", "r", encoding="utf-8") as f:
        contenu = f.read()
    
    # Séparation du YAML (front matter) et du Markdown
    parties = contenu.split("---")
    yaml_content = parties[1] if len(parties) > 1 else ""
    markdown_prose = parties[2] if len(parties) > 2 else ""
    
    tokens = yaml.safe_load(yaml_content) or {}
    
    # Convertir la couleur Hex "#3B49B8" en RGBColor pour python-pptx
    hex_color = tokens.get("colors", {}).get("primary", "#3B49B8").lstrip("#")
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    return {
        "primary": RGBColor(r, g, b),
        "text": RGBColor(255, 255, 255),
        "prose": markdown_prose.strip()
    }

@app.post("/api/appliquer-charte")
async def appliquer_charte(
    fichier: UploadFile = File(...),
    image_prompt: str = Form(""),
    donnees: UploadFile = File(None)
):
    input_path = f"temp_{fichier.filename}"
    
    try:
        # Chargement des tokens depuis DESIGN.md
        design = charger_design_tokens()

        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(fichier.file, buffer)

        prs = Presentation(input_path)

        # 1. APPLICATION DE LA CHARTE (Basée sur DESIGN.md)
        for slide in prs.slides:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = design["primary"]

            for shape in slide.shapes:
                if shape.has_text_frame:
                    is_title = shape == slide.shapes.title
                    
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Inter'
                            run.font.color.rgb = design["text"]
                            
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

        # 2. IA GEMINI (Injecte les règles du DESIGN.md dans le prompt)
        api_key = os.getenv("GEMINI_API_KEY")
        
        if api_key and image_prompt:
            try:
                contexte_data = ""
                if donnees:
                    contenu = await donnees.read()
                    contexte_data = f"\nContexte des données : {contenu.decode('utf-8')[:1000]}"

                client = genai.Client(api_key=api_key)
                
                # Injection de la prose du DESIGN.md pour guider Imagen 3
                prompt_global = f"""Respecte impérativement cette charte visuelle :
                {design['prose']}
                
                Sujet de l'image : {image_prompt} {contexte_data}"""
                
                result = client.models.generate_images(
                    model='imagen-3.0-generate-001',
                    prompt=prompt_global,
                    config=dict(number_of_images=1, aspect_ratio="16:9", output_mime_type="image/jpeg")
                )
                
                image_bytes = result.generated_images[0].image.image_bytes
                img_stream = io.BytesIO(image_bytes)
                
                new_slide = prs.slides.add_slide(prs.slide_layouts[6])
                new_slide.background.fill.solid()
                new_slide.background.fill.fore_color.rgb = design["primary"]
                new_slide.shapes.add_picture(img_stream, Inches(1), Inches(1), width=Inches(8))
                
            except Exception as e:
                print(f"Erreur IA : {e}")

        output_path = "Decathlon_DesignMD.pptx"
        prs.save(output_path)
        
        return FileResponse(
            output_path, 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
            filename="Decathlon_DesignMD.pptx"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail="Erreur serveur.")
        
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

@app.get("/")
async def afficher_page_accueil():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

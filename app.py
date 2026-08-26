from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
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
    input_path = f"temp_{fichier.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(fichier.file, buffer)

    prs = Presentation(input_path)
    # Couleurs de la charte d'après la capture d'écran
    bleu_fond = RGBColor(59, 73, 184)
    texte_blanc = RGBColor(255, 255, 255)
    
    for slide in prs.slides:
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bleu_fond

        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Inter'
                    run.font.color.rgb = texte_blanc
                    if run.font.size and run.font.size > Pt(24):
                        run.font.bold = True
                    elif run.font.size and run.font.size < Pt(20):
                        run.font.size = Pt(18)

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
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

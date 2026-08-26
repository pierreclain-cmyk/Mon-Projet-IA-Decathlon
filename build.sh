#!/usr/bin/env bash

# 1. On installe les dépendances Python
pip install -r requirements.txt

# 2. On télécharge le modèle IA (Mistral 7B quantifié) UNIQUEMENT s'il n'est pas déjà là
if [ ! -f "mistral-7b-instruct.Q4_K_M.gguf" ]; then
    echo "Téléchargement du modèle IA en cours..."
    wget https://huggingface.co/TheBloke/Mistral-7B-Instruct-v0.1-GGUF/resolve/main/mistral-7b-instruct-v0.1.Q4_K_M.gguf -O mistral-7b-instruct.Q4_K_M.gguf
fi

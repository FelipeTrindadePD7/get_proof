#!/bin/bash
# Script para criar executável do Extrator de Comprovantes para Linux

echo "🔧 Criando executável para Linux..."
echo ""

# Instalar PyInstaller se necessário
if ! command -v pyinstaller &> /dev/null; then
    echo "📦 Instalando PyInstaller..."
    pip install --user pyinstaller
fi

# Criar executável
echo "🚀 Compilando executável..."
pyinstaller --onefile \
    --windowed \
    --name="Extrator_Comprovantes_Linux" \
    --add-data="/usr/lib/python3/dist-packages/tkinter:tkinter" \
    get_proof.py

echo ""
echo "✅ Executável criado em: dist/Extrator_Comprovantes_Linux"
echo ""
echo "Para distribuir:"
echo "  - Copie o arquivo dist/Extrator_Comprovantes_Linux para outra máquina Linux"
echo "  - Dê permissão de execução: chmod +x Extrator_Comprovantes_Linux"
echo "  - Execute: ./Extrator_Comprovantes_Linux"

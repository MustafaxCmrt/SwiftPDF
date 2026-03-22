set -e
cd "$(dirname "$0")"
source .venv/bin/activate

pyinstaller \
  --onefile \
  --windowed \
  --name "SwiftPDF" \
  --icon assets/icon1.icns \
  --add-data "assets:assets" \
  main.py

echo ""
echo "Çıktı: dist/SwiftPDF.app"

# PPT Voice Generator (Streamlit)

This app lets users upload a PowerPoint file (.pptx), generate per-slide narration, embed autoplay audio, and download the output .pptx.

## Local Run

1. Install dependencies:
   pip install -r requirements.txt

2. (Optional) Create local secrets file:
   - Copy .streamlit/secrets.toml.example to .streamlit/secrets.toml
   - Fill your Gemini API key and optional model settings

3. Start app:
   streamlit run streamlit_app.py

## Deploy to Streamlit Community Cloud

1. Push this project to a GitHub repository.
2. Open https://share.streamlit.io and sign in.
3. Click New app and select:
   - Repository: your repo
   - Branch: main (or your target branch)
   - Main file path: streamlit_app.py
4. In Advanced settings -> Secrets, add keys from .streamlit/secrets.toml.example.
5. Deploy.

After deployment, Streamlit gives you a public app URL to share.

## Notes

- Audio speed defaults to 1.5x and can be changed in the app UI.
- App uses Gemini API directly (set GEMINI_API_KEY in secrets).
- On Linux/Cloud, embedding uses python-pptx first for better compatibility, then OOXML fallback.
- packages.txt includes ffmpeg for audio tempo processing.

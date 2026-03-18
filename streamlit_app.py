from __future__ import annotations

import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

import generate_voice_ppt as g

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def _default_model() -> str:
    return os.environ.get("DRAUP_LLM_MODEL", "gemini/gemini-2.5-flash-lite")


def _load_secret_env_vars() -> None:
    """Populate process env from Streamlit secrets for hosted deployments."""
    direct_keys = [
        "DRAUP_LLM_ENV",
        "DRAUP_LLM_USER",
        "DRAUP_LLM_PROVIDER",
        "DRAUP_LLM_PROCESS",
        "DRAUP_LLM_MODEL",
        "DRAUP_INPUT_PRICE_PER_MILLION",
        "DRAUP_OUTPUT_PRICE_PER_MILLION",
        "OPENAI_API_KEY",
        "AZURE_OPENAI_API_KEY",
        "GOOGLE_API_KEY",
    ]

    for key in direct_keys:
        if key in st.secrets and key not in os.environ:
            os.environ[key] = str(st.secrets[key])

    if "env" in st.secrets:
        env_block = st.secrets["env"]
        for key, value in env_block.items():
            if key not in os.environ:
                os.environ[key] = str(value)


def _run_pipeline(
    uploaded_name: str,
    uploaded_bytes: bytes,
    model: str,
    audio_speed: float,
    embed_method: str,
) -> tuple[bytes, str, dict[str, float], int]:
    previous_speed = os.environ.get("PPT_AUDIO_SPEED")
    previous_embed = os.environ.get("PPT_AUDIO_EMBED_METHOD")

    try:
        os.environ["PPT_AUDIO_SPEED"] = f"{audio_speed:.2f}"
        os.environ["PPT_AUDIO_EMBED_METHOD"] = embed_method

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_name = Path(uploaded_name).name or "input.pptx"
            input_path = tmp_path / input_name
            input_path.write_bytes(uploaded_bytes)

            output_path = tmp_path / f"{input_path.stem}_with_audio.pptx"
            audio_dir = tmp_path / "audio"
            audio_dir.mkdir(parents=True, exist_ok=True)

            llm_handler = g._create_draup_llm_handler()
            slide_texts = g.extract_slide_texts(str(input_path))
            voiceovers, token_summary = g.generate_voiceovers(slide_texts, llm_handler, model)
            audio_paths = g.generate_audio_files(voiceovers, str(audio_dir))
            g.embed_audio(str(input_path), audio_paths, str(output_path))

            return output_path.read_bytes(), output_path.name, token_summary, len(slide_texts)
    finally:
        if previous_speed is None:
            os.environ.pop("PPT_AUDIO_SPEED", None)
        else:
            os.environ["PPT_AUDIO_SPEED"] = previous_speed

        if previous_embed is None:
            os.environ.pop("PPT_AUDIO_EMBED_METHOD", None)
        else:
            os.environ["PPT_AUDIO_EMBED_METHOD"] = previous_embed


def main() -> None:
    load_dotenv()
    _load_secret_env_vars()

    st.set_page_config(page_title="PPT Voice Generator", page_icon="🔊", layout="centered")
    st.title("PPT Voice Generator")
    st.write(
        "Upload a .pptx file, generate slide voiceovers, and download a new PPT with embedded autoplay audio."
    )

    with st.expander("Settings", expanded=True):
        model = st.text_input("LLM model", value=_default_model())
        audio_speed = st.slider("Audio speed", min_value=0.75, max_value=2.00, value=1.50, step=0.05)
        embed_method = st.selectbox("Embed method", options=["auto", "com", "ooxml"], index=0)

    uploaded_file = st.file_uploader("Upload PowerPoint", type=["pptx"])
    run_clicked = st.button("Generate Output PPT", type="primary", use_container_width=True)

    if run_clicked:
        if uploaded_file is None:
            st.error("Upload a .pptx file to continue.")
            return

        with st.spinner("Generating voiceover PPT. This can take a few minutes..."):
            try:
                output_bytes, output_name, token_summary, slide_count = _run_pipeline(
                    uploaded_name=uploaded_file.name,
                    uploaded_bytes=uploaded_file.getvalue(),
                    model=model,
                    audio_speed=audio_speed,
                    embed_method=embed_method,
                )
            except Exception as exc:
                st.error(f"Generation failed: {exc}")
                st.stop()

        st.success("Done. Download your output file below.")
        st.download_button(
            label="Download Output PPT",
            data=output_bytes,
            file_name=output_name,
            mime=PPTX_MIME,
            use_container_width=True,
        )

        st.subheader("Run Summary")
        st.write(f"Slides processed: {slide_count}")
        st.write(f"Input tokens: {int(token_summary.get('input_tokens', 0)):,}")
        st.write(f"Output tokens: {int(token_summary.get('output_tokens', 0)):,}")
        st.write(f"Total tokens: {int(token_summary.get('total_tokens', 0)):,}")
        st.write(f"Estimated total LLM cost: ${token_summary.get('total_cost', 0.0):.6f}")


if __name__ == "__main__":
    main()

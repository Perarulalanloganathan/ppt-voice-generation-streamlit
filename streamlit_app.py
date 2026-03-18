from __future__ import annotations

import os
import tempfile
from importlib import import_module
from pathlib import Path
from typing import Any

import streamlit as st
from dotenv import load_dotenv

import generate_voice_ppt as g

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def _default_model() -> str:
    return os.environ.get("GEMINI_MODEL", "gemini-2.5-flash-lite")


def _load_secret_env_vars() -> None:
    """Populate process env from Streamlit secrets for hosted deployments."""
    direct_keys = [
        "GEMINI_API_KEY",
        "GOOGLE_API_KEY",
        "GEMINI_MODEL",
        "GEMINI_INPUT_PRICE_PER_MILLION",
        "GEMINI_OUTPUT_PRICE_PER_MILLION",
    ]

    for key in direct_keys:
        if key in st.secrets and key not in os.environ:
            os.environ[key] = str(st.secrets[key])

    if "env" in st.secrets:
        env_block = st.secrets["env"]
        for key, value in env_block.items():
            if key not in os.environ:
                os.environ[key] = str(value)


def _extract_usage_counts(response: Any) -> tuple[int, int, int]:
    usage = getattr(response, "usage_metadata", None)
    if usage is None:
        return 0, 0, 0

    in_tokens = int(getattr(usage, "prompt_token_count", 0) or 0)
    out_tokens = int(getattr(usage, "candidates_token_count", 0) or 0)
    total_tokens = int(getattr(usage, "total_token_count", 0) or 0)
    if total_tokens == 0:
        total_tokens = in_tokens + out_tokens
    return in_tokens, out_tokens, total_tokens


def _extract_gemini_text(response: Any) -> str:
    text = (getattr(response, "text", "") or "").strip()
    if text:
        return text

    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        parts = getattr(content, "parts", None) or []
        chunks = []
        for part in parts:
            part_text = getattr(part, "text", None)
            if part_text:
                chunks.append(str(part_text).strip())
        merged = "\n".join([c for c in chunks if c]).strip()
        if merged:
            return merged

    return ""


def _generate_voiceovers_with_gemini(slide_texts: list[str], model: str) -> tuple[list[str], dict[str, float]]:
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Missing Gemini API key. Set GEMINI_API_KEY (or GOOGLE_API_KEY) in Streamlit secrets."
        )

    try:
        genai = import_module("google.genai")
    except ImportError as exc:
        raise RuntimeError(
            "google-genai is not installed. Add it to requirements and redeploy."
        ) from exc

    client = genai.Client(api_key=api_key)

    voiceovers: list[str] = []
    total_input_tokens = 0
    total_output_tokens = 0
    total_tokens = 0

    input_price_per_million = float(os.environ.get("GEMINI_INPUT_PRICE_PER_MILLION", "0.10"))
    output_price_per_million = float(os.environ.get("GEMINI_OUTPUT_PRICE_PER_MILLION", "0.40"))

    for i, text in enumerate(slide_texts):
        if not text.strip():
            voiceovers.append(f"Slide {i + 1}.")
            continue

        prompt = (
            "You are a presentation narrator. Write a concise, natural spoken voiceover "
            "(2-4 sentences) for the following slide content. Write it conversationally, "
            "as if explaining to an audience. Do NOT start with 'This slide' or 'In this slide'.\n\n"
            f"Slide content:\n{text}\n\nVoiceover:"
        )

        response = client.models.generate_content(model=model, contents=prompt)
        vo = _extract_gemini_text(response)
        if not vo:
            vo = f"Slide {i + 1}."

        in_tok, out_tok, tot_tok = _extract_usage_counts(response)
        total_input_tokens += in_tok
        total_output_tokens += out_tok
        total_tokens += tot_tok

        voiceovers.append(vo)

    input_cost = input_price_per_million * (total_input_tokens / 1_000_000)
    output_cost = output_price_per_million * (total_output_tokens / 1_000_000)
    total_cost = input_cost + output_cost

    token_summary = {
        "input_tokens": float(total_input_tokens),
        "output_tokens": float(total_output_tokens),
        "total_tokens": float(total_tokens),
        "input_cost": input_cost,
        "output_cost": output_cost,
        "total_cost": total_cost,
        "input_price_per_million": input_price_per_million,
        "output_price_per_million": output_price_per_million,
    }
    return voiceovers, token_summary


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

            slide_texts = g.extract_slide_texts(str(input_path))
            voiceovers, token_summary = _generate_voiceovers_with_gemini(slide_texts, model)
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
        embed_method = st.selectbox("Embed method", options=["auto", "com", "python-pptx", "ooxml"], index=0)

    if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        st.warning("Set GEMINI_API_KEY in Streamlit secrets before running generation.")

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

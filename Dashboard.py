import streamlit as st
from utils import (
    AppConfig,
    combine_source_files,
    extract_text_from_uploaded_prompt,
    generate_outline_with_gpt,
    generate_ppt_content_with_gpt,
    create_docx_bytes,
    create_gamma_from_template,
    wait_for_gamma_completion,
    find_gamma_link,
    mask_key,
)

st.set_page_config(page_title="Morning Setup PPT Generator", layout="wide")
st.title("Morning Setup Report → Gamma PPT Generator")

st.sidebar.header("API Configuration")

openai_api_key = st.sidebar.text_input(
    "OpenAI API Key",
    type="password",
    help="Paste your OpenAI API key here. It is used only for this session."
)

gamma_api_key = st.sidebar.text_input(
    "Gamma API Key",
    type="password",
    help="Paste your Gamma API key here. It is used only for this session."
)

st.sidebar.caption(
    f"OpenAI: {mask_key(openai_api_key)} | Gamma: {mask_key(gamma_api_key)}"
)

st.sidebar.header("Gamma Settings")

gamma_id = st.sidebar.text_input(
    "Gamma Template ID",
    value="g_pepbfk69p9lagj1"
)

gamma_theme_id = st.sidebar.text_input(
    "Gamma Theme ID",
    value="ge1kywkagyzapfv"
)

st.sidebar.header("Model Settings")

openai_model = st.sidebar.text_input(
    "OpenAI Model",
    value="gpt-5"
)

st.sidebar.header("Files")

source_files = st.file_uploader(
    "Upload Source PDF/DOCX Files",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

gpt_prompt_file = st.file_uploader(
    "Upload Extractor Prompt DOCX",
    type=["docx"],
    key="gpt_prompt"
)

ppt_prompt_file = st.file_uploader(
    "Upload PPT Prompt DOCX",
    type=["docx"],
    key="ppt_prompt"
)

run_btn = st.button("Generate Morning Setup PPT", use_container_width=True)

if run_btn:
    if not openai_api_key:
        st.error("Please enter OpenAI API key in the sidebar.")
        st.stop()

    if not gamma_api_key:
        st.error("Please enter Gamma API key in the sidebar.")
        st.stop()

    if not source_files:
        st.error("Please upload at least one source PDF/DOCX file.")
        st.stop()

    if not gpt_prompt_file:
        st.error("Please upload the Extractor Prompt DOCX.")
        st.stop()

    if not ppt_prompt_file:
        st.error("Please upload the PPT Prompt DOCX.")
        st.stop()

    config = AppConfig(
        openai_api_key=openai_api_key,
        gamma_api_key=gamma_api_key,
        openai_model=openai_model,
        gamma_id=gamma_id,
        gamma_theme_id=gamma_theme_id,
    )

    progress = st.progress(0, text="Starting workflow...")
    status_box = st.empty()

    try:
        progress.progress(10, text="Extracting text from source files...")
        status_box.info("Reading uploaded source files...")

        raw_text = combine_source_files(source_files)

        st.subheader("Extracted Source Preview")
        st.text_area("Source text", raw_text[:20000], height=250)

        progress.progress(25, text="Reading prompt files...")
        status_box.info("Reading prompt DOCX files...")

        gpt_prompt_text = extract_text_from_uploaded_prompt(gpt_prompt_file)
        ppt_prompt_text = extract_text_from_uploaded_prompt(ppt_prompt_file)

        progress.progress(40, text="Generating combined outline with OpenAI...")
        status_box.info("Creating Combined_Premarket outline...")

        outline_text = generate_outline_with_gpt(
            raw_text=raw_text,
            gpt_prompt_text=gpt_prompt_text,
            config=config,
        )

        st.subheader("Outline Preview")
        st.text_area("Outline", outline_text[:20000], height=250)

        outline_docx_bytes = create_docx_bytes("Combined Premarket Outline", outline_text)

        progress.progress(60, text="Generating PPT content with OpenAI...")
        status_box.info("Creating Gamma-ready PPT content...")

        ppt_content = generate_ppt_content_with_gpt(
            outline_text=outline_text,
            ppt_prompt_text=ppt_prompt_text,
            config=config,
        )

        st.subheader("PPT Content Preview")
        st.text_area("PPT content", ppt_content[:20000], height=250)

        progress.progress(75, text="Sending content to Gamma...")
        status_box.info("Creating presentation in Gamma...")

        generation_id = create_gamma_from_template(
            ppt_content=ppt_content,
            config=config,
        )

        progress.progress(85, text="Waiting for Gamma to finish...")
        status_box.info(f"Gamma generation started. ID: {generation_id}")

        gamma_status_resp = wait_for_gamma_completion(
            generation_id=generation_id,
            config=config,
            poll_interval=5,
            timeout_seconds=900,
        )

        gamma_link = find_gamma_link(gamma_status_resp)

        progress.progress(100, text="Done")
        status_box.success("Presentation workflow completed successfully.")

        st.success("Gamma deck created successfully.")

        if gamma_link:
            st.markdown(f"### Open in Gamma\n[Click here]({gamma_link})")
        else:
            st.info("Gamma completed, but no Gamma URL was found in the response.")

        st.download_button(
            "Download Combined_Premarket.docx",
            data=outline_docx_bytes,
            file_name="Combined_Premarket.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        with st.expander("Gamma Response JSON"):
            st.json(gamma_status_resp)

    except Exception as e:
        progress.progress(100, text="Workflow stopped")
        status_box.error("Workflow failed.")
        st.error(str(e))
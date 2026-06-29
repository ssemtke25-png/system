import google.generativeai as genai
from .prompt_manager import build_main_prompt

def generate_document(api_key, doc_text):
    genai.configure(api_key=api_key)

    model = genai.GenerativeModel("gemini-2.5-flash")

    prompt = build_main_prompt(doc_text)

    response = model.generate_content(prompt)

    return response.text

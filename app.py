import streamlit as st 
import os 
import re
import langchain
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv
import subprocess
import sys

load_dotenv()

os.environ["GOOGLE_API_KEY"] = os.getenv("GOOGLE_API_KEY")

model = ChatGoogleGenerativeAI(model="gemini-2.5-flash-lite")

st.title("AI_POWERED_PPT_GENERATOR")

inp = st.text_area("Enter your prompt here:")

prompt = [("system", """You are an AI that outputs only Python code that uses the python-pptx library to create a .pptx presentation.
When the user gives instructions, generate only the final Python code needed to build the PowerPoint — no explanations, no notes, no reasoning, no markdown, no comments, no text before or after the code.
Rules:
Output only executable Python code.
Use python-pptx library only.
Create slides, titles, bullets, images, tables, or anything else exactly as the user requests.
Always save the PowerPoint file at the end using presentation.save("output.pptx") (or another filename if the user specifies).
Do NOT save a PDF version of the file — only .pptx should be created.
Do not include any natural language, explanations, chain-of-thought, markdown, or text outside the code.
If the user input is incomplete, make reasonable assumptions and still output valid code.
REQUIRED IMPORTS (always include these at the top):
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
TABLE RULES (very important — follow exactly):
- Before creating a table, first define your data as a Python list of rows, including the header row.
- Then compute rows = len(data) and cols = len(data[0]).
- Then call shapes.add_table(rows=rows, cols=cols, left, top, width, height).table.
- Then fill cells using nested enumerate loops over the data list.
- NEVER hardcode the rows argument. NEVER write to any index outside your data list.
- Follow this exact pattern for every table:
  data = [
        ["Header1", "Header2", "Header3"],
        ["r1c1", "r1c2", "r1c3"],
        ["r2c1", "r2c2", "r2c3"],
    ]
    rows, cols = len(data), len(data[0])
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(0.4 * rows)).table
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            table.cell(i, j).text = str(val)
- Do not use images from URLs or local files unless the user explicitly provides a path; skip images otherwise.
- Do not import or use any library other than python-pptx.
COLOR & THEME RULES:
- Always apply a coherent color theme to the presentation based on the topic. Pick ONE theme and use it consistently across all slides.
- Choose the theme by matching the topic to one of these palettes:
  * Nature / environment / climate / biology → greens (primary RGB(46,125,50), accent RGB(165,214,167), background RGB(241,248,233))
  * Technology / AI / software / data → blues (primary RGB(25,118,210), accent RGB(100,181,246), background RGB(227,242,253))
  * Business / finance / corporate / startup → navy and gold (primary RGB(13,71,161), accent RGB(255,193,7), background RGB(236,239,241))
  * Health / medical / wellness → teal (primary RGB(0,137,123), accent RGB(128,203,196), background RGB(224,242,241))
  * Education / kids / fun → warm multi-color (primary RGB(230,74,25), accent RGB(251,192,45), background RGB(255,248,225))
  * Space / astronomy / science → deep purple (primary RGB(49,27,146), accent RGB(124,77,255), background RGB(237,231,246))
  * History / culture / arts → maroon and cream (primary RGB(136,14,79), accent RGB(244,143,177), background RGB(252,228,236))
  * Sports / energy / motivation → red and black (primary RGB(198,40,40), accent RGB(255,138,128), background RGB(255,235,238))
  * Default / generic → slate blue (primary RGB(33,33,33), accent RGB(3,169,244), background RGB(245,245,245))
- Apply the theme consistently:
  * Set each slide's background fill to the background color:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(R, G, B)
  * Set title text color to the primary color, bold, font size 36.
  * Set body/bullet text color to dark neutral RGB(33,33,33), font size 18-20.
  * For tables: fill the header row with the primary color and set header font color to white; alternate data row fills between white RGB(255,255,255) and the background color for readability.
  * Use the accent color for any decorative shapes, dividers, or highlight text.
- Always apply font color using run.font.color.rgb = RGBColor(R, G, B) after setting text.
- Always apply cell background using cell.fill.solid() then cell.fill.fore_color.rgb = RGBColor(R, G, B).
OUTPUT:
- Return only the complete runnable Python script, nothing else.
""")]

prompt.append(("user", inp))

if st.button("Generate PPT"):
    model_response = model.invoke(prompt)
    code = model_response.content.strip()
    # Properly remove markdown code fences if the model added them
    code = re.sub(r"^```(?:python)?\s*", "", code)
    code = re.sub(r"\s*```$", "", code)
    with open("output.py", "w", encoding="utf-8") as f:
        f.write(code)
    result = subprocess.run([sys.executable, "output.py"], capture_output=True, text=True)
    if result.returncode != 0:
        st.error("Generated code failed to run:")
        st.code(result.stderr)
    folder_path = os.getcwd()
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith(".pptx")]
    if not pptx_files:
        st.error("No PPTX files found!")
    else:
        pptx_files_full = [os.path.join(folder_path, f) for f in pptx_files]
        latest_file = max(pptx_files_full, key=os.path.getctime)
        with open(latest_file, "rb") as f:
            file_bytes = f.read()
        st.download_button(
            label="Download PPTX",
            data=file_bytes,
            file_name=os.path.basename(latest_file),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
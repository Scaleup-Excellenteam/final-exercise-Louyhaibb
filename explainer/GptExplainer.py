import asyncio
import json
import os
import logging
from pptx import Presentation
import openai
from pathlib import Path

logging.basicConfig(level=logging.INFO)

openai.api_key = os.environ.get('OPENAI_API_KEY')
if not openai.api_key:
    raise ValueError("Missing OpenAI API key in environment variables.")

def collect_slide_content(slide):
    return "\n".join(shape.text for shape in slide.shapes if hasattr(shape, "text"))

async def interpret_content(content):
    try:
        result = openai.Completion.create(
            model="gpt-3.5-turbo",
            prompt=f"Provide an explanation for:\n\n{content}\n\nInterpretation:",
            max_tokens=150
        )
        return result.choices[0].text.strip()
    except Exception as e:
        logging.error(f"Interpretation failed: {str(e)}")
        return None

async def analyze_presentation(file_path):
    logging.info(f"Analyzing presentation: {file_path}")
    try:
        deck = Presentation(file_path)
    except Exception as e:
        logging.error(f"Presentation access failed: {str(e)}")
        return None

    slide_contents = [collect_slide_content(slide) for slide in deck.slides]
    interpretations = []
    for i, content in enumerate(slide_contents, 1):
        if content:
            logging.info(f"Interpreting slide {i}/{len(slide_contents)}...")
            interpretation = await interpret_content(content)
            if interpretation is None:
                return None
            interpretations.append(interpretation)
            await asyncio.sleep(20 if i % 3 != 0 else 60)
    return interpretations

def store_interpretations(file_path, interpretations):
    output_file = Path(file_path).with_suffix('.json')
    logging.info(f"Storing interpretations in {output_file}")
    with open(output_file, 'w') as f:
        json.dump(interpretations, f, indent=4)

def process_presentation(path):
    file_path = Path(path)
    interpretations = asyncio.run(analyze_presentation(file_path))
    if interpretations is None:
        logging.error("Interpretation generation failed.")
        return 1
    store_interpretations(file_path, interpretations)
    logging.info(f"Interpretations stored in {file_path.with_suffix('.json')}")
    return 0

if __name__ == "__main__":
    process_presentation(" ")
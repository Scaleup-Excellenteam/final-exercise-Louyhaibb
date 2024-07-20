import os
import json
import logging
import asyncio
from pptx import Presentation
import openai
from openai.error import APIError, APIConnectionError, RateLimitError, InvalidRequestError

logging.basicConfig(level=logging.INFO)

API_KEY = os.environ.get("OPENAI_API_KEY")
if not API_KEY:
    logging.critical("OpenAI API key not found in environment variables.")
    raise SystemExit(1)

openai.api_key = API_KEY

def extract_slide_content(slide):
    return " ".join(shape.text for shape in slide.shapes if hasattr(shape, "text")).strip()

async def fetch_interpretation(content, max_attempts=5):
    wait_time = 1
    for attempt in range(max_attempts):
        try:
            response = await openai.ChatCompletion.acreate(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": f"Provide a concise explanation for this slide content:\n\n{content}"}
                ],
                max_tokens=1500,
                temperature=0.5,
            )
            interpretation = response.choices[0].message['content'].strip()
            logging.info(f"Interpretation received: {interpretation[:60]}...")
            return interpretation
        except RateLimitError:
            logging.warning("API rate limit reached. Pausing for 60 seconds...")
            await asyncio.sleep(60)
        except InvalidRequestError as e:
            logging.error(f"Invalid request: {str(e)}")
            return "Unable to process due to insufficient quota. Please review your OpenAI account settings."
        except (APIConnectionError, APIError) as e:
            logging.error(f"API error encountered: {str(e)}. Retrying in {wait_time} seconds...")
            await asyncio.sleep(wait_time)
            wait_time = min(wait_time * 2, 60)
        except Exception as e:
            logging.error(f"Unexpected error occurred: {str(e)}")
            return f"Failed to interpret slide: {str(e)}"
    return "Interpretation failed after multiple attempts."

async def analyze_presentation(file_path):
    logging.info(f"Analyzing presentation: {file_path}")
    try:
        deck = Presentation(file_path)
    except Exception as e:
        logging.error(f"Failed to access presentation: {str(e)}")
        return []

    slide_contents = [extract_slide_content(slide) for slide in deck.slides]
    interpretations = []
    for i, content in enumerate(slide_contents, 1):
        if content:
            logging.info(f"Processing slide {i}/{len(slide_contents)}...")
            interpretation = await fetch_interpretation(content)
            interpretations.append(interpretation)
            await asyncio.sleep(20 if i % 3 != 0 else 60)
    return interpretations

def store_interpretations(file_path, interpretations):
    output_file = f"{os.path.splitext(file_path)[0]}.json"
    logging.info(f"Storing interpretations in {output_file}")
    with open(output_file, 'w') as f:
        json.dump(interpretations, f, indent=4)

def process_presentation():
    file_path = "logging, debugging, getting into a large codebase.pptx"
    interpretations = asyncio.run(analyze_presentation(file_path))
    if interpretations:
        store_interpretations(file_path, interpretations)
        logging.info(f"Interpretations saved to {os.path.splitext(file_path)[0]}.json")
    else:
        logging.error("No interpretations generated due to errors.")

if __name__ == "__main__":
    process_presentation()
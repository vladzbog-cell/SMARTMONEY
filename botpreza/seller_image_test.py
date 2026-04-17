from io import BytesIO

from PIL import Image
from google import genai
from google.genai import types


client = genai.Client(
    api_key='sk-gnOixK_NWy9P76mqF6Q4lg',
    http_options=types.HttpOptions(base_url='https://api.artemox.com'),
)

prompt = 'draw red cat'

response = client.models.generate_content(
    model='gemini-3-pro-image-preview',
    contents=[prompt],
)

print(response)

parts = []
for cand in (getattr(response, 'candidates', None) or []):
    content = getattr(cand, 'content', None)
    if content and getattr(content, 'parts', None):
        parts.extend(content.parts)

image_parts = []

for part in parts:
    txt = getattr(part, 'text', None)
    if txt:
        print(txt)

    if getattr(part, 'inline_data', None) is not None:
        image_parts.append(part)

if not image_parts:
    raise RuntimeError('Model error.')

last_image_part = image_parts[-1]

try:
    img = last_image_part.as_image()
except Exception:
    data = last_image_part.inline_data.data
    img = Image.open(BytesIO(data))

img.save('generated_image.png')
print('saved: generated_image.png')
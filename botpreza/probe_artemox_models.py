from google.genai import types

from artemox_client import get_artemox_client


PROMPT = "A simple red apple on a white background, 16:9"
TEST_MODELS = [
    "gemini-2.5-flash-image",
    "gemini-3-pro-image-preview",
    "imagen-3.0-generate-002",
    "imagen-4.0-generate-preview-06-06",
    "imagen-4.0-generate-001",
]


def print_error(prefix: str, error: Exception) -> None:
    print(f"{prefix}_TYPE: {type(error)}")
    print(f"{prefix}_ERROR: {repr(error)}")
    response = getattr(error, "response", None)
    if response is not None:
        print(f"{prefix}_STATUS: {getattr(response, 'status_code', None)}")
        print(f"{prefix}_TEXT: {getattr(response, 'text', None)}")


def list_models() -> None:
    client = get_artemox_client()
    print("== MODEL LIST ==")
    try:
        pager = client.models.list(config={"page_size": 50})
        shown = 0
        for model in pager:
            name = getattr(model, "name", None)
            if name and any(token in name.lower() for token in ("image", "imagen", "gemini")):
                print(name)
                shown += 1
                if shown >= 30:
                    break
        print(f"MODEL_COUNT_SHOWN: {shown}")
    except Exception as error:
        print_error("LIST", error)


def probe_generate_images() -> None:
    client = get_artemox_client()
    print("== GENERATE IMAGES ==")
    for model in TEST_MODELS:
        print(f"MODEL: {model}")
        try:
            response = client.models.generate_images(
                model=model,
                prompt=PROMPT,
                config=types.GenerateImagesConfig(
                    number_of_images=1,
                    aspect_ratio="16:9",
                    output_mime_type="image/png",
                    include_rai_reason=True,
                ),
            )
            images = getattr(response, "generated_images", None) or getattr(response, "images", None)
            image_count = 0 if images is None else len(images)
            print(f"OK_RESPONSE_TYPE: {type(response)}")
            print(f"IMAGE_COUNT: {image_count}")
            if images:
                first = images[0]
                image = getattr(first, "image", None)
                image_bytes = getattr(image, "image_bytes", None) if image is not None else None
                print(f"FIRST_IMAGE_BYTES: {0 if image_bytes is None else len(image_bytes)}")
                return
        except Exception as error:
            print_error("GEN", error)
        print("---")


if __name__ == "__main__":
    list_models()
    probe_generate_images()
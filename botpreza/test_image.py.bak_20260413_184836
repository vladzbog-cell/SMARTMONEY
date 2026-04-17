import os
from pprint import pprint

from dotenv import load_dotenv
from google import genai
from google.genai import types


def main() -> None:
    load_dotenv("keys.env")

    api_key = os.getenv("ARTEMOX_API_KEY", "").strip()
    base_url = os.getenv("ARTEMOX_BASE_URL", "https://api.artemox.com").strip()
    model = os.getenv("ARTEMOX_IMAGE_GENERATION_MODEL", "gemini-2.5-flash-image").strip()

    if not api_key:
        raise RuntimeError("ARTEMOX_API_KEY is empty in keys.env")

    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(
            api_version="v1alpha",
            base_url=base_url,
        ),
    )

    prompt = "A simple red apple on a white background, 16:9"

    print("=== REQUEST INFO ===")
    print(f"BASE_URL: {base_url}")
    print(f"MODEL: {model}")
    print(f"PROMPT: {prompt}")

    try:
        response = client.models.generate_content(
            model=model,
            contents=[prompt],
        )

        print("=== SUCCESS ===")
        print(f"RESPONSE_TYPE: {type(response)}")
        print(f"HAS_CANDIDATES: {bool(getattr(response, 'candidates', None))}")

        parts = []
        if getattr(response, "candidates", None):
            parts = getattr(response.candidates[0].content, "parts", [])

        print(f"PARTS_COUNT: {len(parts)}")
        for index, part in enumerate(parts, start=1):
            print(f"PART_{index}_HAS_TEXT: {bool(getattr(part, 'text', None))}")
            inline_data = getattr(part, "inline_data", None)
            print(f"PART_{index}_HAS_INLINE_DATA: {bool(inline_data and getattr(inline_data, 'data', None))}")
            if inline_data and getattr(inline_data, "data", None):
                mime_type = getattr(inline_data, "mime_type", None)
                data = inline_data.data
                print(f"PART_{index}_MIME: {mime_type}")
                print(f"PART_{index}_DATA_TYPE: {type(data)}")
                print(f"PART_{index}_DATA_LEN: {len(data)}")

    except Exception as e:
        print("=== ERROR ===")
        print(f"TYPE: {type(e)}")
        print(f"ERROR: {repr(e)}")

        status_code = getattr(e, "status_code", None)
        if status_code is not None:
            print(f"STATUS_CODE: {status_code}")

        response = getattr(e, "response", None)
        if response is not None:
            response_status = getattr(response, "status_code", None)
            if response_status is not None:
                print(f"RESPONSE_STATUS_CODE: {response_status}")

            response_text = getattr(response, "text", None)
            if response_text is not None:
                try:
                    text_value = response_text() if callable(response_text) else response_text
                except Exception as response_text_error:
                    text_value = f"<failed to read response.text: {response_text_error!r}>"
                print(f"RESPONSE_TEXT: {text_value}")

            response_body = getattr(response, "body", None)
            if response_body is not None:
                print(f"RESPONSE_BODY: {response_body}")

            response_headers = getattr(response, "headers", None)
            if response_headers is not None:
                print("RESPONSE_HEADERS:")
                pprint(dict(response_headers))

        sdk_http_response = getattr(e, "sdk_http_response", None)
        if sdk_http_response is not None:
            print(f"SDK_HTTP_RESPONSE: {sdk_http_response}")


if __name__ == "__main__":
    main()
"""
agent/config.py
Configuracoes centralizadas do agente com failover multiprovedor.
Se um provedor falhar, tenta o proximo automaticamente.
"""
import base64
import os
import time

SUPPORTED_PROVIDERS = ["gemini", "groq", "openai", "grok"]

GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash-lite")
GROQ_MODEL_TEXT = os.environ.get("GROQ_MODEL_TEXT", "llama-3.3-70b-versatile")
GROQ_MODEL_VISION = os.environ.get("GROQ_MODEL_VISION", "llama-3.2-11b-vision-preview")
OPENAI_MODEL_TEXT = os.environ.get("OPENAI_MODEL_TEXT", "gpt-4o-mini")
OPENAI_MODEL_VISION = os.environ.get("OPENAI_MODEL_VISION", "gpt-4o-mini")
GROK_MODEL_TEXT = os.environ.get("GROK_MODEL_TEXT", "grok-4-latest")
GROK_MODEL_VISION = os.environ.get("GROK_MODEL_VISION", "grok-4-latest")


def _provider_order() -> list[str]:
    explicit = os.environ.get("LLM_PROVIDER_ORDER", "").strip()
    if explicit:
        order = [p.strip().lower() for p in explicit.split(",") if p.strip()]
    else:
        order = ["gemini", "groq", "openai", "grok"]

    primary = os.environ.get("LLM_PROVIDER", "").strip().lower()
    if primary in SUPPORTED_PROVIDERS:
        order = [primary] + [p for p in order if p != primary]

    dedup = []
    for p in order:
        if p in SUPPORTED_PROVIDERS and p not in dedup:
            dedup.append(p)
    return dedup or ["gemini", "groq"]


def _is_retryable_error(err_text: str) -> bool:
    err = (err_text or "").lower()
    markers = [
        "429", "rate_limit", "resource_exhausted", "quota",
        "timeout", "temporarily unavailable", "connection", "503", "502"
    ]
    return any(m in err for m in markers)


def _build_openai_messages(prompt: str, image_data: dict | None):
    if not image_data:
        return [{"role": "user", "content": prompt}]

    base64_image = base64.b64encode(image_data["data"]).decode("utf-8")
    image_url = f"data:{image_data['mime_type']};base64,{base64_image}"

    return [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": image_url}},
            ],
        }
    ]


def get_llm_response(prompt: str, image_data: dict = None, max_retries: int = 3):
    """Bridge universal com failover multiprovedor."""
    providers = _provider_order()
    last_error = None

    for provider in providers:
        try:
            return _query_by_provider(provider, prompt, image_data, max_retries)
        except Exception as e:
            last_error = e
            retryable = _is_retryable_error(str(e))
            kind = "retryable" if retryable else "non-retryable"
            print(f">>> [Failover] {provider} falhou ({kind}): {e}", flush=True)
            continue

    raise Exception(f"Todos os provedores falharam. Ultimo erro: {last_error}")


def _query_by_provider(provider, prompt, image_data, max_retries):
    if provider == "groq":
        return _query_groq(prompt, image_data, max_retries)
    if provider == "openai":
        return _query_openai(prompt, image_data, max_retries)
    if provider == "grok":
        return _query_grok(prompt, image_data, max_retries)
    return _query_gemini(prompt, image_data, max_retries)


def _query_gemini(prompt, image_data, max_retries):
    from google import genai
    from google.genai import types

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise Exception("GEMINI_API_KEY nao configurada.")

    client = genai.Client(api_key=api_key)

    if image_data:
        image_part = types.Part.from_bytes(data=image_data["data"], mime_type=image_data["mime_type"])
        contents = [prompt, image_part]
    else:
        contents = prompt

    for i in range(max_retries):
        try:
            response = client.models.generate_content(model=GEMINI_MODEL, contents=contents)
            return (response.text or "").strip()
        except Exception as e:
            if _is_retryable_error(str(e)) and i < max_retries - 1:
                wait = 10
                print(f">>> [Gemini] Retentando em {wait}s...", flush=True)
                time.sleep(wait)
            else:
                raise e


def _query_groq(prompt, image_data, max_retries):
    from groq import Groq

    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise Exception("GROQ_API_KEY nao configurada.")

    if image_data and not GROQ_MODEL_VISION:
        raise Exception("GROQ_MODEL_VISION nao configurado; pulando Groq para imagem.")

    client = Groq(api_key=api_key)
    model = GROQ_MODEL_VISION if image_data else GROQ_MODEL_TEXT
    messages = _build_openai_messages(prompt, image_data)

    for i in range(max_retries):
        try:
            completion = client.chat.completions.create(messages=messages, model=model)
            return (completion.choices[0].message.content or "").strip()
        except Exception as e:
            if _is_retryable_error(str(e)) and i < max_retries - 1:
                wait = 5
                print(f">>> [Groq] Retentando em {wait}s...", flush=True)
                time.sleep(wait)
            else:
                raise e


def _query_openai(prompt, image_data, max_retries):
    from openai import OpenAI

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise Exception("OPENAI_API_KEY nao configurada.")

    client = OpenAI(api_key=api_key)
    model = OPENAI_MODEL_VISION if image_data else OPENAI_MODEL_TEXT
    messages = _build_openai_messages(prompt, image_data)

    for i in range(max_retries):
        try:
            completion = client.chat.completions.create(model=model, messages=messages, temperature=0)
            return (completion.choices[0].message.content or "").strip()
        except Exception as e:
            if _is_retryable_error(str(e)) and i < max_retries - 1:
                wait = 5
                print(f">>> [OpenAI] Retentando em {wait}s...", flush=True)
                time.sleep(wait)
            else:
                raise e


def _query_grok(prompt, image_data, max_retries):
    """Grok via endpoint OpenAI-compat da xAI."""
    from openai import OpenAI

    api_key = os.environ.get("XAI_API_KEY") or os.environ.get("GROK_API_KEY")
    if not api_key:
        raise Exception("XAI_API_KEY (ou GROK_API_KEY) nao configurada.")

    client = OpenAI(api_key=api_key, base_url="https://api.x.ai/v1")
    model = GROK_MODEL_VISION if image_data else GROK_MODEL_TEXT
    messages = _build_openai_messages(prompt, image_data)

    for i in range(max_retries):
        try:
            completion = client.chat.completions.create(model=model, messages=messages, stream=False, temperature=0)
            return (completion.choices[0].message.content or "").strip()
        except Exception as e:
            if _is_retryable_error(str(e)) and i < max_retries - 1:
                wait = 5
                print(f">>> [Grok/xAI] Retentando em {wait}s...", flush=True)
                time.sleep(wait)
            else:
                raise e

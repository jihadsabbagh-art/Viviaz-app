# Photorealistic render endpoint — turns a sketch + material references into a
# realistic image using Google's Gemini image model ("Nano Banana").
#
# Configuration (set in Vercel → Project → Settings → Environment Variables):
#   GEMINI_API_KEY        required — your Google AI Studio API key
#   GEMINI_IMAGE_MODEL    optional — defaults to "gemini-2.5-flash-image"
from http.server import BaseHTTPRequestHandler
import json, os, urllib.request, urllib.error

MODEL = os.environ.get("GEMINI_IMAGE_MODEL", "gemini-2.5-flash-image")
API_KEY = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"


def generate_image(prompt, images):
    """Call Gemini and return (base64_image, mime_type).

    images: list of {"data": <base64 str>, "mimeType": <str>}.
    The first image is treated as the base/structure to preserve; any further
    images are material/finish references.
    """
    parts = [{"text": prompt}]
    for img in images:
        data = img.get("data")
        if not data:
            continue
        parts.append({
            "inline_data": {
                "mime_type": img.get("mimeType", "image/jpeg"),
                "data": data,
            }
        })

    body = {
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {"responseModalities": ["TEXT", "IMAGE"]},
    }

    url = ENDPOINT.format(model=MODEL, key=API_KEY)
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    text_out = ""
    for cand in result.get("candidates", []):
        for part in cand.get("content", {}).get("parts", []):
            inline = part.get("inlineData") or part.get("inline_data")
            if inline and inline.get("data"):
                mime = inline.get("mimeType") or inline.get("mime_type") or "image/png"
                return inline["data"], mime
            if part.get("text"):
                text_out += part["text"]

    # No image came back — surface any explanation the model gave, plus a hint
    # about safety blocks.
    reason = text_out.strip()
    block = result.get("promptFeedback", {}).get("blockReason")
    if block:
        reason = (reason + " " if reason else "") + "(blocked: %s)" % block
    raise ValueError(reason or "The model did not return an image. Try rephrasing the prompt.")


class handler(BaseHTTPRequestHandler):
    def _send(self, code, obj):
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(payload)

    def do_POST(self):
        try:
            if not API_KEY:
                return self._send(500, {"error": "Server not configured: set the GEMINI_API_KEY environment variable in Vercel."})
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length)) if length else {}
            prompt = (data.get("prompt") or "").strip()
            images = data.get("images") or []
            if not prompt:
                return self._send(400, {"error": "A prompt is required."})
            if not images:
                return self._send(400, {"error": "At least one input image (the sketch) is required."})
            image_b64, mime = generate_image(prompt, images)
            return self._send(200, {"image": image_b64, "mimeType": mime})
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "ignore")
            return self._send(502, {"error": "Image API error (%s): %s" % (e.code, detail[:600])})
        except urllib.error.URLError as e:
            return self._send(502, {"error": "Could not reach the image API: %s" % e.reason})
        except Exception as e:
            return self._send(500, {"error": str(e)})

    def do_OPTIONS(self):
        self._send(200, {})

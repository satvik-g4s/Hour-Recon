def generate_fix_suggestion(
    user_prompt,
    file_description,
    expected_format,
    error_message
):
    import requests
    import json

    # 🔐 Try getting API key from Streamlit secrets
    try:
        import streamlit as st
        api_key = st.secrets["OPENROUTER_API_KEY"]
    except Exception:
        return "❌ API key not found. Please set OPENROUTER_API_KEY in Streamlit secrets."

    API_URL = "https://openrouter.ai/api/v1/chat/completions"
    MODEL = "openai/gpt-oss-120b:free"

    structured_prompt = f"""
You are an Excel support assistant for a finance team.

Your job:
- Help non-technical users fix their Excel file
- DO NOT use programming terms (no pandas, no code, no traceback explanation)
- Give simple step-by-step instructions

----------------------------------------
USER ACTION:
{user_prompt}

UPLOADED FILE DETAILS:
{file_description}

EXPECTED FORMAT:
{expected_format}

ERROR MESSAGE:
{error_message}
----------------------------------------

INSTRUCTIONS:
1. Identify what is wrong with the file
2. Explain it in SIMPLE terms
3. Give clear steps to fix it in Excel
4. Be specific (column names, formats, etc.)
5. Keep it short and practical

OUTPUT FORMAT:

❌ Problem:
(Explain what is wrong in simple terms)

✅ Fix:
1. Step 1
2. Step 2
3. Step 3

📌 Example (if helpful):
(Optional)

Final Answer:
"""

    try:
        response = requests.post(
            url=API_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            data=json.dumps({
                "model": MODEL,
                "messages": [
                    {"role": "user", "content": structured_prompt}
                ],
                "temperature": 0.2,
                "max_tokens": 500
            }),
            timeout=30
        )

        data = response.json()

        if "choices" in data and len(data["choices"]) > 0:
            return data["choices"][0]["message"]["content"]

        return f"❌ Unable to generate fix suggestion.\n\nRaw response:\n{data}"

    except requests.exceptions.Timeout:
        return "❌ Request timed out. Please try again."

    except requests.exceptions.RequestException as e:
        return f"❌ Network error: {str(e)}"

    except Exception as e:
        return f"❌ Unexpected error: {str(e)}"

#a direct way to consume Mistral Le Plataforme API as an internal endpoint
#part of the Hermes AI Text Sanitizer Outlook Add-in by Eduardo Arana <info@arananet.net>.

from flask import Flask, request, jsonify
from mistralai.client import MistralClient
from mistralai.models.chat_completion import ChatMessage
import os

# Replace with your actual API key
api_key = os.environ["API_KEY"]

def run_mistral(sys_message, user_message, model="mistral-small-latest"):
    client = MistralClient(api_key=api_key)
    messages = [
        ChatMessage(role="system", content=sys_message),
        ChatMessage(role="user", content=user_message)
    ]
    chat_response = client.chat(
        model=model,
        messages=messages
    )
    return chat_response.choices[0].message.content

app = Flask(__name__)

@app.route('/rephrase', methods=['POST'])
def rephrase_text():
    try:
        input_text = request.json['text']

        sys_message = """
Improve the following text by rephrasing it, correcting any spelling or grammar errors, and ensuring it is concise and adheres to ADA style. Maintain a formal, professional tone and preserve the original format, including carriage returns and HTML elements. In addition, incorporate a layer of responsible AI by removing any sexually explicit, non-content-related, or damaging text. Finally, estimate and include a metric on how much time (in minutes) and CO2 emissions have been saved by using the Hermes AI Text Sanitizer. Add this information at the end of the curated message in the following format: '<br><br> Time saved by using the Hermes AI Text Sanitizer: X minutes and an estimated CO2 savings of X.' Please process the given text and do not include the original text in the response:
        """

        response = run_mistral(sys_message, input_text, model="mistral-small-latest")

        return response
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == "__main__":
	port = int(os.environ.get('PORT', 5001))
	app.run(debug=True, host='0.0.0.0', port=port)
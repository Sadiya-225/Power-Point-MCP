from flask import Flask, render_template, request, jsonify
import os
import sys
import traceback
from powerpoint.launcher import PowerPointLauncher
from powerpoint.macro import MacroGenerator
from gemini.client import GeminiClient
from flask_cors import CORS

# Initialize Flask application
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Load configuration
from config import DEBUG, HOST, PORT, GEMINI_API_KEY

# Initialize components
try:
    gemini_client = GeminiClient(GEMINI_API_KEY)
    print("Initialized Gemini client with API key")
except Exception as e:
    print(f"Error initializing Gemini client: {e}")
    gemini_client = GeminiClient()  # Fallback to mock mode
    print("Using mock Gemini client")

macro_generator = MacroGenerator()
powerpoint_launcher = PowerPointLauncher()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_presentation():
    prompt = request.form.get('prompt', '')
    
    if not prompt:
        return jsonify({'error': 'No prompt provided'}), 400
    
    try:
        # Generate macro using Gemini
        print(f"Generating macro for prompt: '{prompt}'")
        macro_response = gemini_client.generate_macro(prompt)
        
        # Convert Gemini response to PowerPoint macro
        macro_code = macro_generator.convert_to_macro(macro_response)
        
        # Launch PowerPoint and execute the macro
        result = powerpoint_launcher.execute_macro(macro_code)
        
        if "error" in result:
            return jsonify({
                'status': 'partial_success',
                'message': 'A basic presentation was created, but the custom macro could not be fully executed.',
                'details': result
            })
        
        return jsonify({
            'status': 'success',
            'message': 'Presentation generated!',
            'details': result
        })
    
    except Exception as e:
        print(f"Error generating presentation: {e}")
        traceback.print_exc()
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

# Run the app if this file is executed directly
if __name__ == '__main__':
    app.run(debug=DEBUG, host=HOST, port=PORT)
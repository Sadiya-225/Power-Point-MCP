# PowerPoint Presentation Generator

This project is a Flask-based web application that generates PowerPoint presentations automatically using AI-generated VBA macros. It leverages the Google Gemini API to create VBA macros based on user prompts, which are then executed in PowerPoint via COM automation on Windows.

## Features

- Web interface to input prompts for presentation generation
- Uses Google Gemini API to generate complete VBA macros for PowerPoint
- Automates PowerPoint to create presentations with custom formatting and multiple slides
- Handles PowerPoint automation including launching the app, injecting macros, and running them
- Provides feedback on success, partial success, or errors during macro execution
- CORS enabled for local frontend development

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/Sadiya-125/Power-Point-MCP.git
   cd Power-Point-MCP
   ```

2. Create and activate a Python virtual environment (optional but recommended):

   ```bash
   python -m venv venv
   venv\Scripts\activate   # On Windows
   source venv/bin/activate  # On Linux/macOS
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Set up environment variables by creating a `.env` file in the project root with the following variables:
   ```
   POWERPOINT_PATH="C:\\Path\\To\\POWERPNT.EXE"
   GEMINI_API_KEY="your_gemini_api_key_here"
   DEBUG=True
   HOST=0.0.0.0
   PORT=5000
   ```

## Configuration

- `POWERPOINT_PATH`: Full path to the PowerPoint executable on your Windows machine.
- `GEMINI_API_KEY`: API key for accessing the Google Gemini generative AI service.
- `DEBUG`: Enable or disable Flask debug mode.
- `HOST`: Host IP address to run the Flask app.
- `PORT`: Port number for the Flask app.

## Usage

1. Run the Flask application:

   ```bash
   python app.py
   ```

2. Open your browser and navigate to `http://localhost:5000` to access the web interface.

3. Enter a prompt describing the presentation content you want to generate and submit.

4. The backend will generate a VBA macro using the Gemini API, inject it into PowerPoint, and execute it to create a presentation with 6 slides styled with specific colors and fonts.

## How It Works

- **Flask App (`app.py`)**: Serves the web interface and handles requests to generate presentations.
- **GeminiClient (`gemini/client.py`)**: Communicates with the Google Gemini API to generate VBA macro code based on the user prompt.
- **MacroGenerator (`powerpoint/macro.py`)**: Converts the Gemini API response into a VBA macro script.
- **PowerPointLauncher (`powerpoint/launcher.py`)**: Automates PowerPoint on Windows using COM to launch the app, inject the macro, and run it to create the presentation.

## Dependencies

- Flask
- Flask-CORS
- python-dotenv
- requests
- google-generativeai
- pywin32 (Windows only)

## Notes

- This project requires Windows OS with Microsoft PowerPoint installed.
- The PowerPoint VBA project must allow programmatic access; the app attempts to enable this via registry modification.
- The generated macro creates a 6-slide presentation with hot pink titles and dark blue backgrounds, using Arial font.
- The app includes error handling for partial macro execution or failures.

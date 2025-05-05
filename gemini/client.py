import google.generativeai as genai
import config
import re

class GeminiClient:
    def __init__(self, api_key=None):
        # Use your API key
        self.api_key = config.GEMINI_API_KEY
        
        # Configure the API key
        genai.configure(api_key=self.api_key)
        
        # Create model instance with the correct model name
        self.model = genai.GenerativeModel('gemini-2.0-pro-exp')
    
    def generate_macro(self, prompt):
        """Generate PowerPoint VBA macro using Gemini API"""
        enhanced_prompt = f"""
            Create a complete, self-contained PowerPoint VBA macro for a presentation on the following resume content: "{prompt}"

            Requirements:
            1. The macro MUST be a complete, executable VBA script with Sub Main() as the entry point
            2. The macro should create a PowerPoint presentation with 6 slides, each with a title and content
            3. Use hot pink (#FF00B8) for titles and dark blue (#002060) for backgrounds
            4. Use Arial font for all text and set the font size to 44 for titles and 24 for content
            5. IMPORTANT: Define ALL PowerPoint constants explicitly at the top of your code (e.g., Const ppLayoutTitle = 1, Const ppLayoutTitleAndContent = 2, etc.)
            6. DO NOT include any explanatory text, only provide the VBA code
            7. The code must be fully functional without any external requirements
            8. IMPORTANT: Use ONLY ASCII characters in the code (no special characters, emojis, or non-ASCII symbols)
            9. Replace any non-ASCII characters with their ASCII equivalents (e.g., use ' instead of ')

            Format: Return ONLY the complete VBA code without any explanation or additional text.
            """
        
        try:
            # Generate content using the model
            response = self.model.generate_content(enhanced_prompt)
            
            print("API Response received, length:", len(response.text) if response.text else 0)
            
            # Clean up the response to get just the VBA code
            vba_code = response.text
            
            # Remove any markdown code block markers if present
            if vba_code.startswith("```vba"):
                vba_code = vba_code[6:]
            if vba_code.startswith("```"):
                vba_code = vba_code[3:]
            if vba_code.endswith("```"):
                vba_code = vba_code[:-3]
                
            vba_code = vba_code.strip()
            
            # Ensure the macro has a Main subroutine
            if "Sub Main()" not in vba_code:
                vba_code = f"""
                Sub Main()
                    {vba_code}
                End Sub
                """
            
            # Clean up any remaining non-ASCII characters
            vba_code = re.sub(r'[^\x00-\x7F]+', ' ', vba_code)
            
            return vba_code
                
        except Exception as e:
            print(f"Error calling Gemini API: {e}")
            # Return a default VBA script in case of failure
            return """
            Sub Main()
                ' Default PowerPoint macro
                
                ' Create presentation
                Dim pres As Presentation
                Set pres = ActivePresentation
                
                ' Define colors
                Const HOT_PINK = 12058879 ' RGB(255, 0, 184)
                Const ELECTRIC_PURPLE = 16711841 ' RGB(161, 0, 255)
                
                ' Clear existing slides
                While pres.Slides.Count > 0
                    pres.Slides(1).Delete
                Wend
                
                ' Add title slide
                Dim titleSlide As Slide
                Set titleSlide = pres.Slides.Add(1, ppLayoutTitle)
                titleSlide.Shapes.Title.TextFrame.TextRange.Text = "Generated Presentation"
                
                ' Format title
                titleSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = HOT_PINK
                titleSlide.Shapes.Title.TextFrame.TextRange.Font.Size = 44
                
                ' Set background
                titleSlide.Background.Fill.ForeColor.RGB = ELECTRIC_PURPLE
                
                ' Add content slides
                AddContentSlide "Overview", "• Point 1" & vbNewLine & "• Point 2" & vbNewLine & "• Point 3"
                AddContentSlide "Details", "• More information" & vbNewLine & "• Additional details" & vbNewLine & "• Technical aspects"
                AddContentSlide "Summary", "• Recap" & vbNewLine & "• Conclusion" & vbNewLine & "• Next steps"
            End Sub

            Sub AddContentSlide(title As String, content As String)
                ' Add a content slide
                Dim slide As Slide
                Set slide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutText)
                
                ' Set title and content
                slide.Shapes.Title.TextFrame.TextRange.Text = title
                slide.Shapes(2).TextFrame.TextRange.Text = content
                
                ' Format
                slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = 12058879 ' Hot Pink
                slide.Background.Fill.ForeColor.RGB = 16711841 ' Electric Purple
            End Sub
            """
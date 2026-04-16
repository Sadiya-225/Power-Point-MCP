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
        self.model = genai.GenerativeModel('gemini-2.5-flash-lite')
    
    def extract_resume_data(self, resume_content):
        """Extract key information from resume content"""
        lines = resume_content.split('\n')

        # Try to extract name (first line is usually the name)
        name = "Resume"
        email = "user@example.com"

        for line in lines:
            if '@' in line:  # Found email
                email = line.split()[0] if ' ' in line else line
            # First non-empty line might be name
            if line.strip() and not name.startswith("##"):
                name = line.replace("##", "").replace("<div", "").replace("</div>", "").strip()
                if name and len(name) > 0:
                    break

        # Clean up name
        name = re.sub(r'[<>"/|]', '', name).strip()[:50] or "Resume"

        return name, email

    def generate_macro(self, prompt):
        """Generate PowerPoint VBA macro using Gemini API"""
        enhanced_prompt = f"""
            Create a complete, self-contained PowerPoint VBA macro for a presentation based on this resume content: "{prompt}"

            Use this exact template structure and fill in the placeholders with relevant information from the resume:

            ```
            Sub Main()
                ' PowerPoint VBA Macro Template
                ' Constants for PowerPoint layouts
                Const ppLayoutTitle = 1
                Const ppLayoutTitleAndContent = 2
                Const ppLayoutText = 3
                Const ppLayoutTwoContent = 4
                
                ' Color constants
                Const HOT_PINK = 12058879 ' RGB(255, 0, 184)
                Const DARK_BLUE = 15773696 ' RGB(0, 32, 96)
                Const WHITE = 16777215 ' RGB(255, 255, 255)
                
                ' Get active presentation
                Dim pres As Presentation
                Set pres = ActivePresentation
                
                ' Clear existing slides
                While pres.Slides.Count > 0
                    pres.Slides(1).Delete
                Wend
                
                ' Slide 1: Title Slide
                Dim slide1 As Slide
                Set slide1 = pres.Slides.Add(1, ppLayoutTitle)
                slide1.Shapes.Title.TextFrame.TextRange.Text = "[FULL_NAME]"
                slide1.Shapes(2).TextFrame.TextRange.Text = "[PROFESSIONAL_TITLE]" & vbNewLine & "[CONTACT_INFO]"
                
                ' Format title slide
                FormatSlide slide1, HOT_PINK, DARK_BLUE
                
                ' Slide 2: Professional Summary
                Dim slide2 As Slide
                Set slide2 = pres.Slides.Add(2, ppLayoutTitleAndContent)
                slide2.Shapes.Title.TextFrame.TextRange.Text = "Professional Summary"
                slide2.Shapes(2).TextFrame.TextRange.Text = "[PROFESSIONAL_SUMMARY]"
                
                ' Slide 3: Skills
                Dim slide3 As Slide
                Set slide3 = pres.Slides.Add(3, ppLayoutTitleAndContent)
                slide3.Shapes.Title.TextFrame.TextRange.Text = "Technical Skills"
                slide3.Shapes(2).TextFrame.TextRange.Text = "[SKILLS_LIST]"
                
                ' Slide 4: Work Experience
                Dim slide4 As Slide
                Set slide4 = pres.Slides.Add(4, ppLayoutTitleAndContent)
                slide4.Shapes.Title.TextFrame.TextRange.Text = "Work Experience"
                slide4.Shapes(2).TextFrame.TextRange.Text = "[WORK_EXPERIENCE]"
                
                ' Slide 5: Education
                Dim slide5 As Slide
                Set slide5 = pres.Slides.Add(5, ppLayoutTitleAndContent)
                slide5.Shapes.Title.TextFrame.TextRange.Text = "Education"
                slide5.Shapes(2).TextFrame.TextRange.Text = "[EDUCATION_DETAILS]"
                
                ' Slide 6: Projects
                Dim slide6 As Slide
                Set slide6 = pres.Slides.Add(6, ppLayoutTitleAndContent)
                slide6.Shapes.Title.TextFrame.TextRange.Text = "Key Projects"
                slide6.Shapes(2).TextFrame.TextRange.Text = "[PROJECTS_DETAILS]"
                
                ' Format all slides
                FormatSlide slide2, HOT_PINK, DARK_BLUE
                FormatSlide slide3, HOT_PINK, DARK_BLUE
                FormatSlide slide4, HOT_PINK, DARK_BLUE
                FormatSlide slide5, HOT_PINK, DARK_BLUE
                FormatSlide slide6, HOT_PINK, DARK_BLUE
            End Sub

            Sub FormatSlide(slide As Slide, titleColor As Long, bgColor As Long)
                ' Format slide title
                slide.Shapes.Title.TextFrame.TextRange.Font.Name = "Arial"
                slide.Shapes.Title.TextFrame.TextRange.Font.Size = 44
                slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = titleColor
                slide.Shapes.Title.TextFrame.TextRange.Font.Bold = True
                
                ' Format slide content
                If slide.Shapes.Count > 1 Then
                    slide.Shapes(2).TextFrame.TextRange.Font.Name = "Arial"
                    slide.Shapes(2).TextFrame.TextRange.Font.Size = 24
                    slide.Shapes(2).TextFrame.TextRange.Font.Color.RGB = WHITE
                End If
                
                ' Set background
                slide.Background.Fill.ForeColor.RGB = bgColor
            End Sub
            ```

            INSTRUCTIONS:
            1. Replace [FULL_NAME] with the person's full name from the resume
            2. Replace [PROFESSIONAL_TITLE] with their current role or a professional title
            3. Replace [CONTACT_INFO] with email, phone, and LinkedIn (formatted nicely)
            4. Replace [PROFESSIONAL_SUMMARY] with a concise 2-3 sentence summary
            5. Replace [SKILLS_LIST] with key technical skills formatted as bullet points
            6. Replace [WORK_EXPERIENCE] with job title, company, duration, and key achievements
            7. Replace [EDUCATION_DETAILS] with degree, institution, and GPA/achievements
            8. Replace [PROJECTS_DETAILS] with project names, descriptions, and outcomes
            
            Requirements:
            - Use ONLY ASCII characters (no special symbols, emojis, or non-ASCII characters)
            - Keep content concise and professional
            - Format bullet points using "• "
            - Ensure all placeholders are replaced with actual resume content
            - Return ONLY the complete VBA code without any explanation
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
            # Extract actual resume data for the fallback
            name, email = self.extract_resume_data(prompt)

            # Return a VBA script using extracted resume data
            return f"""
            Sub Main()
                ' PowerPoint VBA Macro - Generated from Resume
                ' Constants for PowerPoint layouts
                Const ppLayoutTitle = 1
                Const ppLayoutTitleAndContent = 2
                Const ppLayoutText = 3
                Const ppLayoutTwoContent = 4

                ' Color constants
                Const HOT_PINK = 12058879 ' RGB(255, 0, 184)
                Const DARK_BLUE = 15773696 ' RGB(0, 32, 96)
                Const WHITE = 16777215 ' RGB(255, 255, 255)

                ' Get active presentation
                Dim pres As Presentation
                Set pres = ActivePresentation

                ' Clear existing slides
                While pres.Slides.Count > 0
                    pres.Slides(1).Delete
                Wend

                ' Slide 1: Title Slide
                Dim slide1 As Slide
                Set slide1 = pres.Slides.Add(1, ppLayoutTitle)
                slide1.Shapes.Title.TextFrame.TextRange.Text = "{name}"
                slide1.Shapes(2).TextFrame.TextRange.Text = "Professional Resume" & vbNewLine & "{email}"

                ' Format title slide
                FormatSlide slide1, HOT_PINK, DARK_BLUE

                ' Slide 2: Professional Summary
                Dim slide2 As Slide
                Set slide2 = pres.Slides.Add(2, ppLayoutTitleAndContent)
                slide2.Shapes.Title.TextFrame.TextRange.Text = "Professional Summary"
                slide2.Shapes(2).TextFrame.TextRange.Text = "See complete resume for detailed information about professional background, skills, experience, and achievements."

                ' Slide 3: Skills
                Dim slide3 As Slide
                Set slide3 = pres.Slides.Add(3, ppLayoutTitleAndContent)
                slide3.Shapes.Title.TextFrame.TextRange.Text = "Technical Skills"
                slide3.Shapes(2).TextFrame.TextRange.Text = "See resume for complete list of technical skills, programming languages, and tools."

                ' Slide 4: Work Experience
                Dim slide4 As Slide
                Set slide4 = pres.Slides.Add(4, ppLayoutTitleAndContent)
                slide4.Shapes.Title.TextFrame.TextRange.Text = "Work Experience"
                slide4.Shapes(2).TextFrame.TextRange.Text = "See resume for detailed work history and achievements."

                ' Slide 5: Education
                Dim slide5 As Slide
                Set slide5 = pres.Slides.Add(5, ppLayoutTitleAndContent)
                slide5.Shapes.Title.TextFrame.TextRange.Text = "Education"
                slide5.Shapes(2).TextFrame.TextRange.Text = "See resume for educational background and qualifications."

                ' Slide 6: Projects
                Dim slide6 As Slide
                Set slide6 = pres.Slides.Add(6, ppLayoutTitleAndContent)
                slide6.Shapes.Title.TextFrame.TextRange.Text = "Key Projects"
                slide6.Shapes(2).TextFrame.TextRange.Text = "See resume for portfolio and key project details."

                ' Format all slides
                FormatSlide slide2, HOT_PINK, DARK_BLUE
                FormatSlide slide3, HOT_PINK, DARK_BLUE
                FormatSlide slide4, HOT_PINK, DARK_BLUE
                FormatSlide slide5, HOT_PINK, DARK_BLUE
                FormatSlide slide6, HOT_PINK, DARK_BLUE
            End Sub

            Sub FormatSlide(slide As Slide, titleColor As Long, bgColor As Long)
                ' Format slide title
                slide.Shapes.Title.TextFrame.TextRange.Font.Name = "Arial"
                slide.Shapes.Title.TextFrame.TextRange.Font.Size = 44
                slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = titleColor
                slide.Shapes.Title.TextFrame.TextRange.Font.Bold = True

                ' Format slide content
                If slide.Shapes.Count > 1 Then
                    slide.Shapes(2).TextFrame.TextRange.Font.Name = "Arial"
                    slide.Shapes(2).TextFrame.TextRange.Font.Size = 24
                    slide.Shapes(2).TextFrame.TextRange.Font.Color.RGB = WHITE
                End If

                ' Set background
                slide.Background.Fill.ForeColor.RGB = bgColor
            End Sub
            """
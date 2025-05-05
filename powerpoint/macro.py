class MacroGenerator:
    def __init__(self):
        pass
    
    def convert_to_macro(self, gemini_response):
        """Simply pass through the Gemini response as the macro code"""
        return gemini_response
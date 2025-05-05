import os
import subprocess
import tempfile
import time
import win32com.client
import pythoncom
import traceback
import win32api
import win32con

class PowerPointLauncher:
    def __init__(self, powerpoint_path=None):
        from config import POWERPOINT_PATH
        self.powerpoint_path = powerpoint_path or POWERPOINT_PATH
    
    def enable_vba_access(self, enable=True):
        """Enable or disable programmatic access to VBA project"""
        try:
            # Open the registry key for PowerPoint security
            key = win32api.RegOpenKeyEx(
                win32con.HKEY_CURRENT_USER,
                "Software\\Microsoft\\Office\\16.0\\PowerPoint\\Security", 
                0, 
                win32con.KEY_ALL_ACCESS
            )
            
            # Set the AccessVBOM value (1 to enable, 0 to disable)
            value = 1 if enable else 0
            win32api.RegSetValueEx(key, "AccessVBOM", 0, win32con.REG_DWORD, value)
            
            print(f"VBA project access {'enabled' if enable else 'disabled'}")
            return True
        except Exception as e:
            print(f"Error modifying registry for VBA access: {e}")
            traceback.print_exc()
            return False
    
    def launch_powerpoint(self):
        """Launch PowerPoint application with COM initialization"""
        pythoncom.CoInitialize()
        
        try:
            # Try to connect to running instance first
            print("Trying to connect to running PowerPoint instance...")
            powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            powerpoint.Visible = True
            print("Connected to existing PowerPoint instance")
            return powerpoint
        except:
            # Launch new instance if none is running
            print(f"Launching PowerPoint from: {self.powerpoint_path}")
            try:
                subprocess.Popen([self.powerpoint_path])
                time.sleep(5)  # Give PowerPoint more time to start
                
                # Connect to the newly launched instance
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                powerpoint.Visible = True
                print("PowerPoint launched successfully")
                return powerpoint
            except Exception as e:
                print(f"Error launching PowerPoint: {e}")
                traceback.print_exc()
                pythoncom.CoUninitialize()
                raise
    
    def execute_macro(self, macro_code):
        """Execute VBA macro in PowerPoint without modifications"""
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Create temporary macro file
        temp_dir = tempfile.gettempdir()
        macro_file = os.path.join(temp_dir, "temp_macro.bas")
        
        try:
            # Enable VBA access through registry
            self.enable_vba_access(True)
            
            print(f"Writing macro to {macro_file}")
            with open(macro_file, "w", encoding='utf-8') as f:
                f.write(macro_code)
            
            print(f"Macro file length: {len(macro_code)} characters")
            print("First 200 chars of macro:")
            print(macro_code[:200] + "..." if len(macro_code) > 200 else macro_code)
            
            # Launch PowerPoint
            powerpoint = self.launch_powerpoint()
            
            # Create new presentation if none exists
            if powerpoint.Presentations.Count == 0:
                print("Creating new presentation")
                presentation = powerpoint.Presentations.Add()
            else:
                print("Using active presentation")
                presentation = powerpoint.ActivePresentation
            
            print("Attempting to run macro...")
            
            # Add a module and run Main
            vba_project = presentation.VBProject
            
            # Remove any existing modules with the same name
            for component in vba_project.VBComponents:
                if component.Name == "Module1":
                    vba_project.VBComponents.Remove(component)
            
            # Add new module
            component = vba_project.VBComponents.Add(1)  # 1 = Module
            component.Name = "Module1"
            component.CodeModule.AddFromString(macro_code)
            
            # Save the presentation to ensure the macro is properly added
            temp_presentation = os.path.join(temp_dir, "temp_presentation.pptm")
            presentation.SaveAs(temp_presentation)
            
            # Compile the VBA project to ensure the macro is properly compiled
            try:
                vba_project.VBComponents.Compile()
                print("VBA project compiled successfully")
            except Exception as e:
                print(f"Warning: Could not compile VBA project: {e}")
            
            # Run the macro
            try:
                powerpoint.Run("Module1.Main")
                print("Macro executed successfully")
            except Exception as e:
                print(f"Error running macro: {e}")
                # Try alternative method to run the macro
                try:
                    vba_project.VBComponents("Module1").CodeModule.Run("Main")
                    print("Macro executed successfully using alternative method")
                except Exception as e2:
                    print(f"Error running macro with alternative method: {e2}")
                    raise e2
            
            try:
                # Try to get the active presentation after macro execution
                if powerpoint.Presentations.Count > 0:
                    presentation = powerpoint.ActivePresentation
                    result = {
                        "slides_count": presentation.Slides.Count,
                        "presentation_name": presentation.Name
                    }
                else:
                    result = {
                        "error": "No active presentation found after macro execution",
                        "slides_count": 0,
                        "presentation_name": "Unknown"
                    }
            except Exception as e:
                print(f"Error getting presentation details: {e}")
                result = {
                    "error": str(e),
                    "slides_count": 0,
                    "presentation_name": "Unknown"
                }
            
            return result
            
        except Exception as e:
            print(f"Error executing macro: {e}")
            traceback.print_exc()
            raise
        finally:
            # Disable VBA access
            self.enable_vba_access(False)
            pythoncom.CoUninitialize()
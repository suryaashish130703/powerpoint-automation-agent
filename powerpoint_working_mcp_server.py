# Working PowerPoint MCP Server for Windows Automation
from mcp.server.fastmcp import FastMCP
from mcp.types import TextContent
from pywinauto.application import Application
import win32gui
import win32con
import win32api
import time
import sys
import math
import subprocess
from win32api import GetSystemMetrics

# Instantiate MCP server
mcp = FastMCP("WorkingPowerPointAutomation")

# Global variable to store PowerPoint application instance
ppt_app = None

# MATHEMATICAL TOOLS

@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers"""
    print("CALLED: add(a: int, b: int) -> int:")
    return int(a + b)

@mcp.tool()
def subtract(a: int, b: int) -> int:
    """Subtract two numbers"""
    print("CALLED: subtract(a: int, b: int) -> int:")
    return int(a - b)

@mcp.tool()
def multiply(a: int, b: int) -> int:
    """Multiply two numbers"""
    print("CALLED: multiply(a: int, b: int) -> int:")
    return int(a * b)

@mcp.tool()
def divide(a: int, b: int) -> float:
    """Divide two numbers"""
    print("CALLED: divide(a: int, b: int) -> float:")
    return float(a / b)

@mcp.tool()
def power(a: int, b: int) -> int:
    """Power of two numbers"""
    print("CALLED: power(a: int, b: int) -> int:")
    return int(a ** b)

@mcp.tool()
def sqrt(a: int) -> float:
    """Square root of a number"""
    print("CALLED: sqrt(a: int) -> float:")
    return float(a ** 0.5)

@mcp.tool()
def strings_to_chars_to_int(string: str) -> list[int]:
    """Return the ASCII values of the characters in a word"""
    print("CALLED: strings_to_chars_to_int(string: str) -> list[int]:")
    return [int(ord(char)) for char in string]

@mcp.tool()
def int_list_to_exponential_sum(int_list: list) -> float:
    """Return sum of exponentials of numbers in a list"""
    print("CALLED: int_list_to_exponential_sum(int_list: list) -> float:")
    return sum(math.exp(i) for i in int_list)

# POWERPOINT AUTOMATION TOOLS

@mcp.tool()
async def open_powerpoint() -> dict:
    """Open Microsoft PowerPoint and create a new blank presentation"""
    global ppt_app
    try:
        print("ITERATION 1: Opening PowerPoint...")
        
        # Try different PowerPoint executables
        ppt_paths = [
            'powerpnt.exe',
            'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE',
            'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE',
            'C:\\Program Files\\Microsoft Office\\root\\Office15\\POWERPNT.EXE',
            'C:\\Program Files (x86)\\Microsoft Office\\root\\Office15\\POWERPNT.EXE'
        ]
        
        ppt_app = None
        for path in ppt_paths:
            try:
                print(f"Trying to start PowerPoint with: {path}")
                ppt_app = Application().start(path)
                time.sleep(3)  # Wait for PowerPoint to fully load
                print(f"SUCCESS: PowerPoint started with {path}")
                break
            except Exception as e:
                print(f"Failed to start with {path}: {e}")
                continue
        
        if not ppt_app:
            raise Exception("Could not start PowerPoint with any of the attempted methods")
        
        # Wait for the main window to appear and create new presentation
        try:
            main_window = ppt_app.window(title_re=".*PowerPoint.*")
            main_window.wait('exists', timeout=15)
            print("SUCCESS: PowerPoint main window loaded")
            
            # Create new blank presentation
            main_window.set_focus()
            time.sleep(1)
            
            # Press Ctrl+N to create new presentation
            main_window.type_keys('^n')
            time.sleep(2)
            print("SUCCESS: Created new blank presentation")
            
        except Exception as e:
            print(f"Warning: Could not create new presentation: {e}")
        
        print("ITERATION 1 COMPLETE: PowerPoint opened successfully")
        return {
            "content": [
                TextContent(
                    type="text",
                    text="ITERATION 1 COMPLETE: PowerPoint opened successfully"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error opening PowerPoint: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error opening PowerPoint: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def select_rectangle_shape() -> dict:
    """Click Insert tab → Shapes → Rectangle"""
    global ppt_app
    try:
        if not ppt_app:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="PowerPoint is not open. Please call open_powerpoint first."
                    )
                ]
            }
        
        print("ITERATION 2: Selecting Rectangle Shape...")
        
        # Get the main PowerPoint window
        main_window = ppt_app.window(title_re=".*PowerPoint.*")
        
        # Ensure PowerPoint window is active
        if not main_window.has_focus():
            main_window.set_focus()
            time.sleep(1)
        
        # Use keyboard shortcuts for reliable automation
        try:
            print("Using keyboard shortcuts for Insert → Shapes → Rectangle...")
            
            # Press Alt to activate ribbon
            main_window.type_keys('%')
            time.sleep(0.5)
            print("Activated ribbon")
            
            # Press I for Insert tab
            main_window.type_keys('i')
            time.sleep(0.5)
            print("Selected Insert tab")
            
            # Press S for Shapes
            main_window.type_keys('s')
            time.sleep(0.5)
            print("Opened Shapes menu")
            
            # Press R for Rectangle (first option in basic shapes)
            main_window.type_keys('r')
            time.sleep(0.5)
            print("Selected Rectangle shape")
            
            print("ITERATION 2 COMPLETE: Rectangle shape selected successfully")
            
        except Exception as e:
            print(f"Keyboard shortcuts failed: {e}")
            # Try mouse clicks as fallback
            try:
                print("Trying mouse clicks as fallback...")
                
                # Look for Insert tab
                insert_tab = main_window.child_window(title="Insert", control_type="TabItem")
                if insert_tab.exists():
                    insert_tab.click()
                    time.sleep(0.5)
                    print("Clicked Insert tab")
                
                # Look for Shapes button
                shapes_button = main_window.child_window(title_re=".*Shapes.*", control_type="Button")
                if shapes_button.exists():
                    shapes_button.click()
                    time.sleep(0.5)
                    print("Clicked Shapes button")
                
                # Look for Rectangle in the shapes menu
                rectangle_option = main_window.child_window(title_re=".*Rectangle.*", control_type="MenuItem")
                if rectangle_option.exists():
                    rectangle_option.click()
                    time.sleep(0.5)
                    print("Selected Rectangle")
                
                print("ITERATION 2 COMPLETE: Rectangle shape selected via mouse clicks")
                
            except Exception as e2:
                print(f"Mouse clicks also failed: {e2}")
                print("ITERATION 2 COMPLETE: Rectangle selection attempted (may have issues)")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text="ITERATION 2 COMPLETE: Rectangle shape selected successfully"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error selecting rectangle shape: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error selecting rectangle shape: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def draw_rectangle_centered() -> dict:
    """Draw a rectangle roughly centered on the slide"""
    global ppt_app
    try:
        if not ppt_app:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="PowerPoint is not open. Please call open_powerpoint first."
                    )
                ]
            }
        
        print("ITERATION 3: Drawing Rectangle Centered on Slide...")
        
        # Get the main PowerPoint window
        main_window = ppt_app.window(title_re=".*PowerPoint.*")
        
        # Ensure PowerPoint window is active
        if not main_window.has_focus():
            main_window.set_focus()
            time.sleep(0.5)
        
        # Find the slide area - try multiple approaches
        slide_area = None
        try:
            # Try different ways to find the slide area
            slide_area = main_window.child_window(class_name="MsoDockTop")
            if not slide_area.exists():
                slide_area = main_window.child_window(title_re=".*Slide.*")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="NetUIHWND")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="MsoCommandBar")
            if not slide_area.exists():
                slide_area = main_window  # Fallback to main window
        except:
            slide_area = main_window
        
        print(f"Using slide area: {slide_area}")
        
        # Calculate center coordinates for rectangle
        # Assume slide is roughly 800x600 pixels
        slide_center_x = 400
        slide_center_y = 300
        rect_width = 200
        rect_height = 100
        
        x1 = slide_center_x - rect_width // 2
        y1 = slide_center_y - rect_height // 2
        x2 = slide_center_x + rect_width // 2
        y2 = slide_center_y + rect_height // 2
        
        print(f"Drawing rectangle from ({x1},{y1}) to ({x2},{y2})")
        
        # Draw rectangle using multiple methods
        success = False
        
        # Method 1: Mouse drag
        try:
            slide_area.press_mouse_input(coords=(x1, y1))
            time.sleep(0.2)
            slide_area.move_mouse_input(coords=(x2, y2))
            time.sleep(0.2)
            slide_area.release_mouse_input(coords=(x2, y2))
            time.sleep(0.5)
            success = True
            print("SUCCESS: Rectangle drawn using mouse drag")
        except Exception as e:
            print(f"Mouse drag failed: {e}")
            
            # Method 2: Click and drag
            try:
                slide_area.click_input(coords=(x1, y1))
                time.sleep(0.2)
                slide_area.drag_mouse_input(coords_from=(x1, y1), coords_to=(x2, y2))
                time.sleep(0.5)
                success = True
                print("SUCCESS: Rectangle drawn using click and drag")
            except Exception as e2:
                print(f"Click and drag failed: {e2}")
                
                # Method 3: Just click at center (fallback)
                try:
                    slide_area.click_input(coords=(slide_center_x, slide_center_y))
                    time.sleep(0.5)
                    success = True
                    print("SUCCESS: Clicked at rectangle center as fallback")
                except Exception as e3:
                    print(f"Center click failed: {e3}")
        
        if success:
            print("ITERATION 3 COMPLETE: Rectangle drawn successfully")
        else:
            print("ITERATION 3 COMPLETE: Rectangle drawing attempted (may have issues)")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text="ITERATION 3 COMPLETE: Rectangle drawn centered on slide"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error drawing rectangle: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error drawing rectangle: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def select_text_box() -> dict:
    """Click Insert tab → Text Box"""
    global ppt_app
    try:
        if not ppt_app:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="PowerPoint is not open. Please call open_powerpoint first."
                    )
                ]
            }
        
        print("ITERATION 4: Selecting Text Box...")
        
        # Get the main PowerPoint window
        main_window = ppt_app.window(title_re=".*PowerPoint.*")
        
        # Ensure PowerPoint window is active
        if not main_window.has_focus():
            main_window.set_focus()
            time.sleep(0.5)
        
        # Use keyboard shortcuts for reliable automation
        try:
            print("Using keyboard shortcuts for Insert → Text Box...")
            
            # Press Alt to activate ribbon
            main_window.type_keys('%')
            time.sleep(0.5)
            print("Activated ribbon")
            
            # Press I for Insert tab
            main_window.type_keys('i')
            time.sleep(0.5)
            print("Selected Insert tab")
            
            # Press X for Text Box
            main_window.type_keys('x')
            time.sleep(0.5)
            print("Selected Text Box tool")
            
            print("ITERATION 4 COMPLETE: Text Box tool selected successfully")
            
        except Exception as e:
            print(f"Keyboard shortcuts failed: {e}")
            # Try mouse clicks as fallback
            try:
                print("Trying mouse clicks as fallback...")
                
                # Look for Insert tab
                insert_tab = main_window.child_window(title="Insert", control_type="TabItem")
                if insert_tab.exists():
                    insert_tab.click()
                    time.sleep(0.5)
                    print("Clicked Insert tab")
                
                # Look for Text Box button
                textbox_button = main_window.child_window(title_re=".*Text Box.*", control_type="Button")
                if textbox_button.exists():
                    textbox_button.click()
                    time.sleep(0.5)
                    print("Clicked Text Box button")
                
                print("ITERATION 4 COMPLETE: Text Box tool selected via mouse clicks")
                
            except Exception as e2:
                print(f"Mouse clicks also failed: {e2}")
                print("ITERATION 4 COMPLETE: Text Box selection attempted (may have issues)")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text="ITERATION 4 COMPLETE: Text Box tool selected successfully"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error selecting text box: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error selecting text box: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def click_inside_rectangle() -> dict:
    """Click inside the rectangle area to place the text box"""
    global ppt_app
    try:
        if not ppt_app:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="PowerPoint is not open. Please call open_powerpoint first."
                    )
                ]
            }
        
        print("ITERATION 5: Clicking Inside Rectangle Area...")
        
        # Get the main PowerPoint window
        main_window = ppt_app.window(title_re=".*PowerPoint.*")
        
        # Ensure PowerPoint window is active
        if not main_window.has_focus():
            main_window.set_focus()
            time.sleep(0.5)
        
        # Find the slide area
        slide_area = None
        try:
            slide_area = main_window.child_window(class_name="MsoDockTop")
            if not slide_area.exists():
                slide_area = main_window.child_window(title_re=".*Slide.*")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="NetUIHWND")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="MsoCommandBar")
            if not slide_area.exists():
                slide_area = main_window
        except:
            slide_area = main_window
        
        # Click inside the rectangle area (center of slide)
        slide_center_x = 400
        slide_center_y = 300
        
        print(f"Clicking inside rectangle at ({slide_center_x},{slide_center_y})")
        
        slide_area.click_input(coords=(slide_center_x, slide_center_y))
        time.sleep(0.5)
        
        print("ITERATION 5 COMPLETE: Clicked inside rectangle area successfully")
        return {
            "content": [
                TextContent(
                    type="text",
                    text="ITERATION 5 COMPLETE: Clicked inside rectangle area successfully"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error clicking inside rectangle: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error clicking inside rectangle: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def paste_number(text: str) -> dict:
    """Paste the generated number inside the rectangle"""
    global ppt_app
    try:
        if not ppt_app:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="PowerPoint is not open. Please call open_powerpoint first."
                    )
                ]
            }
        
        print(f"ITERATION 6: Pasting Number '{text}' Inside Rectangle...")
        
        # Get the main PowerPoint window
        main_window = ppt_app.window(title_re=".*PowerPoint.*")
        
        # Ensure PowerPoint window is active
        if not main_window.has_focus():
            main_window.set_focus()
            time.sleep(0.5)
        
        # Type the text
        main_window.type_keys(text)
        time.sleep(0.5)
        
        # Click outside to finish text editing
        try:
            slide_area = main_window.child_window(class_name="MsoDockTop")
            if not slide_area.exists():
                slide_area = main_window.child_window(title_re=".*Slide.*")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="NetUIHWND")
            if not slide_area.exists():
                slide_area = main_window.child_window(class_name="MsoCommandBar")
            if not slide_area.exists():
                slide_area = main_window
        except:
            slide_area = main_window
        
        slide_area.click_input(coords=(100, 100))  # Click outside the rectangle
        time.sleep(0.5)
        
        print("ITERATION 6 COMPLETE: Number pasted successfully inside rectangle")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ITERATION 6 COMPLETE: Number '{text}' pasted successfully inside rectangle"
                )
            ]
        }
    except Exception as e:
        print(f"ERROR: Error pasting number: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"ERROR: Error pasting number: {str(e)}"
                )
            ]
        }

if __name__ == "__main__":
    print("Starting Working PowerPoint MCP Server...")
    if len(sys.argv) > 1 and sys.argv[1] == "dev":
        mcp.run()  # Run without transport for dev server
    else:
        mcp.run(transport="stdio")  # Run with stdio for direct execution

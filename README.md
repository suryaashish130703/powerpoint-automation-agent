# Paint Agent with MCP Server

This project demonstrates an AI agent that solves mathematical problems and visualizes the results using Microsoft Paint through an MCP (Model Context Protocol) server.

## Features

- 🤖 AI-powered mathematical problem solving using Google Gemini
- 🎨 Automated Paint application control
- 📊 Visual representation of mathematical results
- 🔄 Iterative problem-solving with detailed logging
- 🖥️ Single monitor support (optimized for primary display)

## Prerequisites

- Windows 10/11
- Python 3.8+
- Microsoft Paint (built into Windows)
- Google Gemini API key

## Installation

1. **Clone or download the project files**

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your Google Gemini API key:**
   - Create a `.env` file in the project directory
   - Add your API key:
     ```
     GEMINI_API_KEY=your_api_key_here
     ```

## Files Overview

- `powerpoint_working_agent.py` - Main agent script that solves math problems and controls PowerPoint
- `powerpoint_working_mcp_server.py` - MCP server that provides PowerPoint automation tools
- `requirements.txt` - Python dependencies
- `README.md` - This documentation file

## Usage

### Quick Start

1. **Run the PowerPoint agent:**
   ```bash
   python powerpoint_working_agent.py
   ```

### What the Agent Does

The agent follows this workflow:

1. **Mathematical Problem Solving:**
   - Solves: "Find ASCII values of characters in 'INDIA' and return sum of exponentials"
   - Uses iterative approach with detailed logging
   - Shows each step in the terminal

2. **PowerPoint Visualization (6 steps):**
   - **Step 1:** Opens Microsoft PowerPoint and creates new presentation
   - **Step 2:** Selects rectangle shape (Insert → Shapes → Rectangle)
   - **Step 3:** Draws a rectangle centered on the slide
   - **Step 4:** Selects text box tool (Insert → Text Box)
   - **Step 5:** Clicks inside rectangle area to place text box
   - **Step 6:** Pastes the final result inside the rectangle

### Terminal Output

The agent provides detailed logging with clear iteration headers:

```
============================================================
ITERATION 1: Processing
============================================================
Starting LLM generation...
LLM generation completed
LLM Response: FUNCTION_CALL: strings_to_chars_to_int|INDIA
Calling function: strings_to_chars_to_int
Parameters: ['INDIA']
Function result: [73, 78, 68, 73, 65]

=== AUTOMATIC POWERPOINT WORKFLOW STARTING ===
============================================================
ITERATION 1: Opening PowerPoint and Creating New Presentation
============================================================
ITERATION 1 COMPLETE: PowerPoint opened successfully
```

## MCP Server Tools

The PowerPoint MCP server provides these tools:

- `open_powerpoint()` - Opens PowerPoint and creates new presentation
- `select_rectangle_shape()` - Selects rectangle shape (Insert → Shapes → Rectangle)
- `draw_rectangle_centered()` - Draws rectangle centered on slide
- `select_text_box()` - Selects text box tool (Insert → Text Box)
- `click_inside_rectangle()` - Clicks inside rectangle area to place text box
- `paste_number(text)` - Pastes text inside the rectangle

## Customization

### Changing Rectangle Position
Edit the coordinates in `powerpoint_working_mcp_server.py`:
```python
# In draw_rectangle_centered function
slide_center_x = 400  # Center X
slide_center_y = 300  # Center Y
rect_width = 200      # Rectangle width
rect_height = 100     # Rectangle height
```

### Modifying the Math Problem
Change the query in `powerpoint_working_agent.py`:
```python
query = """Your custom math problem here. After getting the final answer, open PowerPoint, draw a rectangle, and write the result inside it."""
```

## Troubleshooting

### Common Issues

1. **PowerPoint doesn't open:**
   - Ensure Microsoft PowerPoint is installed
   - Check if PowerPoint is accessible via `powerpnt.exe`

2. **Tool selection fails:**
   - PowerPoint interface may vary by version
   - The server uses keyboard shortcuts for reliability

3. **Rectangle drawing issues:**
   - Adjust coordinates based on your screen resolution
   - The server tries multiple drawing methods automatically

4. **API key errors:**
   - Verify `.env` file exists and contains valid `GEMINI_API_KEY`
   - Check API key permissions

### Debug Mode

Run with verbose output:
```bash
python -u powerpoint_working_agent.py
```

## Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────┐
│powerpoint_agent │───▶│powerpoint_mcp_   │───▶│ Microsoft   │
│   (Main Logic)  │    │   server         │    │ PowerPoint  │
└─────────────────┘    └──────────────────┘    └─────────────┘
         │                        │
         ▼                        ▼
┌─────────────────┐    ┌──────────────────┐
│  Google Gemini  │    │  Windows API     │
│   (AI Model)    │    │ (pywinauto, etc) │
└─────────────────┘    └──────────────────┘
```

## License

This project is for educational and demonstration purposes.

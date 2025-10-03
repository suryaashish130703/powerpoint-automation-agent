# PowerPoint Automation Agent with MCP Server

This project demonstrates an AI agent that solves mathematical problems and automatically visualizes the results using Microsoft PowerPoint through an MCP (Model Context Protocol) server, with email logging capabilities.

## Features

- 🤖 AI-powered mathematical problem solving using Google Gemini
- 📊 Automated PowerPoint presentation creation and manipulation
- 🎨 Automatic rectangle drawing and text insertion
- 📧 Email logging and notifications (success/error reports)
- 🔄 Iterative problem-solving with detailed logging
- 🖥️ Single monitor support (optimized for primary display)

## Prerequisites

- Windows 10/11
- Python 3.8+
- Microsoft PowerPoint (Office 365 or standalone)
- Google Gemini API key
- Email account for logging (Gmail recommended)

## Installation

1. **Clone or download the project files**

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your environment configuration:**
   - Create a `.env` file in the project directory
   - Copy settings from `email_config_template.txt`
   - Add your API key and email settings:
     ```
     GEMINI_API_KEY=your_api_key_here
     SENDER_EMAIL=your-email@gmail.com
     SENDER_PASSWORD=your_app_password_here
     RECIPIENT_EMAIL=recipient@example.com
     SMTP_SERVER=smtp.gmail.com
     SMTP_PORT=587
     ```

4. **Configure Email Logging (Optional but Recommended):**
   - For Gmail users:
     1. Enable 2-Factor Authentication on your Google account
     2. Go to Google Account Settings > Security > 2-Step Verification > App passwords
     3. Generate an app password for "Mail"
     4. Use this app password (not your regular Gmail password) as `SENDER_PASSWORD`
   - Test your email configuration:
     ```bash
     python test_email_logger.py
     ```

## Files Overview

- `powerpoint_working_agent.py` - Main agent script that solves math problems and controls PowerPoint
- `powerpoint_working_mcp_server.py` - MCP server that provides PowerPoint automation tools
- `email_logger.py` - Email logging module for sending execution logs and notifications
- `test_email_logger.py` - Test script to verify email configuration
- `email_config_template.txt` - Template for email configuration settings
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

## Email Logging Features

The agent includes comprehensive email logging capabilities:

### 📧 **Automatic Email Notifications**
- **Success Emails:** Sent when the agent completes successfully
- **Error Emails:** Sent when the agent encounters errors
- **Detailed Logs:** All execution steps are included in the email

### 📋 **Email Content Includes:**
- Final mathematical result
- Execution time and timestamp
- Complete step-by-step logs
- Status (SUCCESS/ERROR)
- Formatted HTML emails with color-coded logs

### 🔧 **Email Configuration:**
- Supports Gmail, Outlook, Yahoo, and custom SMTP servers
- Uses App Passwords for Gmail (more secure than regular passwords)
- Configurable sender and recipient emails
- Test script included to verify configuration

### 📤 **Email Examples:**
- **Success:** "✅ PowerPoint Automation - SUCCESS" with final result and logs
- **Error:** "❌ PowerPoint Automation - ERROR" with error details and logs

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

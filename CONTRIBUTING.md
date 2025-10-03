# Contributing to PowerPoint Automation Agent

Thank you for your interest in contributing to this project! This document provides guidelines for contributing to the PowerPoint Automation Agent.

## Getting Started

1. Fork the repository
2. Clone your fork locally
3. Create a new branch for your feature or bugfix
4. Make your changes
5. Test your changes thoroughly
6. Submit a pull request

## Development Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up environment variables:**
   - Create a `.env` file
   - Add your Google Gemini API key:
     ```
     GEMINI_API_KEY=your_api_key_here
     ```

3. **Test the system:**
   ```bash
   python powerpoint_working_agent.py
   ```

## Code Style

- Follow PEP 8 Python style guidelines
- Use meaningful variable and function names
- Add comments for complex logic
- Keep functions focused and small

## Testing

Before submitting a pull request, please ensure:

- [ ] The code runs without errors
- [ ] PowerPoint automation works correctly
- [ ] All terminal logs are clear and informative
- [ ] No sensitive information is exposed

## Reporting Issues

When reporting issues, please include:

- Operating system and version
- Python version
- PowerPoint version
- Steps to reproduce the issue
- Expected vs actual behavior
- Any error messages or logs

## Feature Requests

For feature requests, please:

- Describe the feature clearly
- Explain why it would be useful
- Provide examples of how it would work
- Consider the impact on existing functionality

## Pull Request Process

1. Ensure your code follows the project's style guidelines
2. Add tests if applicable
3. Update documentation if needed
4. Provide a clear description of your changes
5. Reference any related issues

## Questions?

If you have any questions about contributing, please open an issue or contact the maintainers.

Thank you for contributing! 🚀

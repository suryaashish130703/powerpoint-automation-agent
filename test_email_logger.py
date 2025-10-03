"""
Test script for email logging functionality
Run this to test your email configuration before running the main agent
"""

import os
from dotenv import load_dotenv
from email_logger import EmailLogger

def test_email_configuration():
    """Test email configuration and send a test email"""
    print("Testing Email Logger Configuration...")
    print("=" * 50)
    
    # Load environment variables
    load_dotenv()
    
    # Initialize email logger
    logger = EmailLogger()
    
    if not logger.enabled:
        print("❌ Email configuration incomplete!")
        print("\nPlease configure the following in your .env file:")
        print("SENDER_EMAIL=your-email@gmail.com")
        print("SENDER_PASSWORD=your-app-password")
        print("RECIPIENT_EMAIL=recipient@example.com")
        print("\nFor Gmail users:")
        print("1. Enable 2-Factor Authentication")
        print("2. Generate an App Password")
        print("3. Use the App Password (not your regular password)")
        return False
    
        print("SUCCESS: Email configuration found!")
        print(f"Sender: {logger.sender_email}")
        print(f"Recipient: {logger.recipient_email}")
        print(f"SMTP Server: {logger.smtp_server}:{logger.smtp_port}")
    
    # Test connection
    print("\nTesting SMTP connection...")
    if logger.test_email_connection():
        print("SUCCESS: SMTP connection successful!")
        
        # Send test email
        print("\nSending test email...")
        test_logs = [
            "Test log message 1 - Agent started",
            "Test log message 2 - Connection established", 
            "SUCCESS: Test email functionality working",
            "Test log message 3 - Agent completed successfully"
        ]
        
        success = logger.send_log_email(
            test_logs, 
            "PowerPoint Agent - Email Test"
        )
        
        if success:
            print("SUCCESS: Test email sent successfully!")
            print(f"Check your inbox at: {logger.recipient_email}")
            return True
        else:
            print("ERROR: Failed to send test email")
            return False
    else:
        print("ERROR: SMTP connection failed!")
        print("Please check your email credentials and network connection")
        return False

def test_success_email():
    """Test success email format"""
    print("\nTesting Success Email Format...")
    print("=" * 50)
    
    logger = EmailLogger()
    if logger.enabled:
        test_logs = [
            "Agent started successfully",
            "Mathematical calculation completed",
            "PowerPoint opened successfully", 
            "Rectangle drawn successfully",
            "Text pasted successfully",
            "SUCCESS: All operations completed"
        ]
        
        logger.send_success_email("42.123456", 15.67, test_logs)
        print("SUCCESS: Success email test completed!")

def test_error_email():
    """Test error email format"""
    print("\nTesting Error Email Format...")
    print("=" * 50)
    
    logger = EmailLogger()
    if logger.enabled:
        test_logs = [
            "Agent started successfully",
            "Mathematical calculation completed",
            "ERROR: PowerPoint failed to open",
            "ERROR: Connection timeout occurred"
        ]
        
        logger.send_error_email("PowerPoint application failed to start", test_logs)
        print("SUCCESS: Error email test completed!")

if __name__ == "__main__":
    print("PowerPoint Automation Agent - Email Logger Test")
    print("=" * 60)
    
    # Test basic configuration
    if test_email_configuration():
        print("\n" + "=" * 60)
        print("SUCCESS: All email tests completed successfully!")
        print("Your email configuration is working correctly")
        print("You can now run the main agent with email logging")
        
        # Optional: Test different email formats
        test_success_email()
        test_error_email()
    else:
        print("\n" + "=" * 60)
        print("ERROR: Email configuration needs to be fixed")
        print("Please check the email_config_template.txt file for instructions")

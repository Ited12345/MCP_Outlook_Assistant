from fastmcp import FastMCP
import win32com.client
from datetime import datetime

mcp = FastMCP("mcp_assistant")

def connect():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")    
    return namespace
    
@mcp.tool
def date_tool() -> str:
    """
    Returns the current date in the format YYYYMMDD.

    Returns:
        str: The current date in the format YYYYMMDD.
    """
    return datetime.now().strftime("%Y%m%d")

@mcp.tool
def email_brief(num: int = None, date: int = None) -> str:
    """
    Reads the specified number of emails from the default inbox.

    Args:
        num (int): The number of emails to read. Default is 100. Set to higher number in order to read more emails.
        date (int): The date to filter emails in the format YYYYMMDD. Default is None.

    Returns:
        str: A string containing the details (email index, subject, sender, received_time) of the emails.
    """
    
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox folder
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending

    if date:
        # Convert date parameter to datetime object for comparison
        target_date = datetime.strptime(str(date), "%Y%m%d")
        # Filter messages to only include those from the specified date
        messages = [msg for msg in messages if msg.ReceivedTime.date() == target_date.date()]

    email_info = []  # Store info of each email
    count = 0

    for message in messages:
        if num and count >= num:  # Stop after num emails
            break
        try:
            subject = message.Subject if message.Subject else "No Subject"
            sender = message.SenderName if message.SenderName else "Unknown Sender"
            received_time = message.ReceivedTime if message.ReceivedTime else "Unknown Time"

            # Store details for return value
            email_info.append(
                f"Email_index: {count}\n Subject: {subject}\n From: {sender}\n Received Time: {received_time}\n###"
            )
            count += 1
        except Exception as e:
            print(f"Error processing email: {e}")
            continue

    return "\n".join(email_info)  # Return concatenated email info 

@mcp.tool
def email_body(email_index: int, date: int = None) -> str:
    """
    Reads the body of the email at the specified index.

    Args:
        email_index (int): The index of the email to read the body from.
        date (int): The date to filter emails in the format YYYYMMDD. Default is None.

    Returns:
        str: The body of the email at the specified index.
    """

    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox folder
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending
    
    if date:
        # Convert date parameter to datetime object for comparison
        target_date = datetime.strptime(str(date), "%Y%m%d")
        # Filter messages to only include those from the specified date
        messages = [msg for msg in messages if msg.ReceivedTime.date() == target_date.date()]
    try:
        for i, message in enumerate(messages):
            if i == email_index:
                body = message.Body if message.Body else message.HTMLBody
                body = body[:3000].strip() + "..."  # Limit to 3000 characters
                return f"Email_index: {email_index}\n Body: {body}\n###"
    except Exception as e:
        raise Exception(f"Error reading email body: {e}")
            
    return "Email not found at the specified index." 

@mcp.tool
def create_calendar_event(event_start: str, event_end: str, event_title: str) -> str:
    """
    Creates a calendar event with the given due date and title.

    Args:
        event_start (str): The start date and time in 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' format.
        event_end (str): The end date and time in 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' format.
        event_title (str): The title of the event to be marked in the calendar.
    Returns:
        str: A string indicating the success of the operation.  
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # 9 refers to the calendar folder
        event = calendar.Items.Add()
        event.Subject = event_title
        event.Start = event_start
        event.End = event_end
        event.Save()
        return f"Event '{event_title}' with due date '{due_date}' marked in calendar."

    except Exception as e:
        raise Exception(f"Error marking event in calendar: {str(e)}")

if __name__ == "__main__":
    try:
        namespace = connect()
        mcp.run()   
    except Exception as e:
        print(f"Error running MCP assistant: {e}")

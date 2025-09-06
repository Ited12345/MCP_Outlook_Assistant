# MCP Assistant

This MCP Assistant provides tools to interact with Outlook, allowing you to read emails and create calendar events.

## Features

### Email Management
- `email_brief(num: int = 100, date: int = None)`: Reads a specified number of emails from your default inbox. You can filter by date.
- `email_body(email_index: int)`: Reads the body of a specific email by its index.

### Calendar Management
- `create_calendar_event(event_start: str, event_end: str, event_title: str)`: Creates a new calendar event with a specified title, start date, and end date.

## Setup and Usage

### Prerequisites

- Python 3.x
- Outlook installed and configured
- Windows operating system

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Ited12345/MCP_Outlook_Assistant.git
   cd MCP_Outlook_Assistant
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the main application:
   ```bash
   python outlook_assistant.py
   ```
4. Add the json to agent mcp
```
{
  "mcpServers": {
    "mcp_assistant": {
      "command": "uv",
      "args": [
        "--directory",
        "Replace your path/MCP_Outlook_Assistant",
        "run",
        "outlook_assistant.py"
      ]
    }
  }
}
```

Once running, the MCP Assistant will be available to process requests related to email and calendar operations.

## Example 

### Listing emails by numbers

```
User: Can you list the last 5 emails I received?
```

### Listing emails by date

```
User: Can you list the emails I received on 2025-06-09?
```

### Read email body

```
User: Can you read the body of the email with subject: Dinner on Sunday?
```

### Create calendar event

```
User: Read the emails of yesterday and mark it in calender if it is a event.
```
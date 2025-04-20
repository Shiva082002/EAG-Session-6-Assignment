# basic import 
from mcp.server.fastmcp import FastMCP, Image
from mcp.server.fastmcp.prompts import base
from mcp.types import TextContent
from PIL import Image as PILImage
import math
import sys
from pywinauto.application import Application
import win32gui
import time
from win32api import GetSystemMetrics
import pyautogui
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from dotenv import load_dotenv
import win32com.client as win32
from models import (
    AddInput, AddOutput, AddListInput, AddListOutput,
    SubtractInput, SubtractOutput, MultiplyInput, MultiplyOutput,
    DivideInput, DivideOutput, PowerInput, PowerOutput,
    SqrtInput, SqrtOutput, CbrtInput, CbrtOutput,
    FactorialInput, FactorialOutput, LogInput, LogOutput,
    RemainderInput, RemainderOutput, SinInput, SinOutput,
    CosInput, CosOutput, TanInput, TanOutput,
    MineInput, MineOutput, StringsToIntsInput, StringsToIntsOutput,
    ExpSumInput, ExpSumOutput, FibonacciInput, FibonacciOutput,
    CreateThumbnailInput, CreateThumbnailOutput,mergerCellInput,mergerCellOutput,
    takeScreenshotInput, takeScreenshotOutput,enterTextCenteredInput,enterTextCenteredOutput
)

# instantiate an MCP server client
mcp = FastMCP("Calculator")

# Load environment variables
load_dotenv()

# DEFINE TOOLS

#addition tool
@mcp.tool()
def add(input: AddInput) -> AddOutput:
    """Add two numbers together.
    
    Args:
        input (AddInput): Input model containing two integers to add
        
    Returns:
        AddOutput: Output model containing the sum of the two numbers
    """
    print("CALLED: add(AddInput) -> AddOutput")
    return AddOutput(result=int(input.a + input.b))

@mcp.tool()
def add_list(input: AddListInput) -> AddListOutput:
    """Calculate the sum of all numbers in a list.
    
    Args:
        input (AddListInput): Input model containing a list of integers to sum
        
    Returns:
        AddListOutput: Output model containing the sum of all numbers in the list
    """
    print("CALLED: add_list(AddListInput) -> AddListOutput")
    return AddListOutput(result=sum(input.l))

# subtraction tool
@mcp.tool()
def subtract(input: SubtractInput) -> SubtractOutput:
    """Subtract one number from another.
    
    Args:
        input (SubtractInput): Input model containing two integers to subtract
        
    Returns:
        SubtractOutput: Output model containing the difference between the two numbers
    """
    print("CALLED: subtract(a: int, b: int) -> int:")
    return SubtractOutput(result=int(input.a - input.b))

# multiplication tool
@mcp.tool()
def multiply(input: MultiplyInput) -> MultiplyOutput:
    """Multiply two numbers together.
    
    Args:
        input (MultiplyInput): Input model containing two integers to multiply
        
    Returns:
        MultiplyOutput: Output model containing the product of the two numbers
    """
    print("CALLED: multiply(a: int, b: int) -> int:")
    return MultiplyOutput(result=int(input.a * input.b))

#  division tool
@mcp.tool() 
def divide(input: DivideInput) -> DivideOutput:
    """Divide one number by another.
    
    Args:
        input (DivideInput): Input model containing two integers to divide
        
    Returns:
        DivideOutput: Output model containing the quotient of the division
    """
    print("CALLED: divide(a: int, b: int) -> float:")
    return DivideOutput(result=float(input.a / input.b))

# power tool
@mcp.tool()
def power(input: PowerInput) -> PowerOutput:
    """Calculate the power of a number.
    
    Args:
        input (PowerInput): Input model containing base and exponent integers
        
    Returns:
        PowerOutput: Output model containing the result of base raised to the exponent
    """
    print("CALLED: power(a: int, b: int) -> int:")
    return PowerOutput(result=int(input.a ** input.b))

# square root tool
@mcp.tool()
def sqrt(input: SqrtInput) -> SqrtOutput:
    """Calculate the square root of a number.
    
    Args:
        input (SqrtInput): Input model containing the number to calculate square root of
        
    Returns:
        SqrtOutput: Output model containing the square root of the input number
    """
    print("CALLED: sqrt(a: int) -> float:")
    return SqrtOutput(result=float(input.a ** 0.5))

# cube root tool
@mcp.tool()
def cbrt(input: CbrtInput) -> CbrtOutput:
    """Calculate the cube root of a number.
    
    Args:
        input (CbrtInput): Input model containing the number to calculate cube root of
        
    Returns:
        CbrtOutput: Output model containing the cube root of the input number
    """
    print("CALLED: cbrt(a: int) -> float:")
    return CbrtOutput(result=float(input.a ** (1/3)))

# factorial tool
@mcp.tool()
def factorial(input: FactorialInput) -> FactorialOutput:
    """Calculate the factorial of a number.
    
    Args:
        input (FactorialInput): Input model containing the number to calculate factorial of
        
    Returns:
        FactorialOutput: Output model containing the factorial of the input number
    """
    print("CALLED: factorial(a: int) -> int:")
    return FactorialOutput(result=int(math.factorial(input.a)))

# log tool
@mcp.tool()
def log(input: LogInput) -> LogOutput:
    """Calculate the natural logarithm of a number.
    
    Args:
        input (LogInput): Input model containing the number to calculate logarithm of
        
    Returns:
        LogOutput: Output model containing the natural logarithm of the input number
    """
    print("CALLED: log(a: int) -> float:")
    return LogOutput(result=float(math.log(input.a)))

# remainder tool
@mcp.tool()
def remainder(input: RemainderInput) -> RemainderOutput:
    """Calculate the remainder of division between two numbers.
    
    Args:
        input (RemainderInput): Input model containing dividend and divisor integers
        
    Returns:
        RemainderOutput: Output model containing the remainder of the division
    """
    print("CALLED: remainder(a: int, b: int) -> int:")
    return RemainderOutput(result=int(input.a % input.b))

# sin tool
@mcp.tool()
def sin(input: SinInput) -> SinOutput:
    """Calculate the sine of a number.
    
    Args:
        input (SinInput): Input model containing the number to calculate sine of
        
    Returns:
        SinOutput: Output model containing the sine of the input number
    """
    print("CALLED: sin(a: int) -> float:")
    return SinOutput(result=float(math.sin(input.a)))

# cos tool
@mcp.tool()
def cos(input: CosInput) -> CosOutput:
    """Calculate the cosine of a number.
    
    Args:
        input (CosInput): Input model containing the number to calculate cosine of
        
    Returns:
        CosOutput: Output model containing the cosine of the input number
    """
    print("CALLED: cos(a: int) -> float:")
    return CosOutput(result=float(math.cos(input.a)))

# tan tool
@mcp.tool()
def tan(input: TanInput) -> TanOutput:
    """Calculate the tangent of a number.
    
    Args:
        input (TanInput): Input model containing the number to calculate tangent of
        
    Returns:
        TanOutput: Output model containing the tangent of the input number
    """
    print("CALLED: tan(a: int) -> float:")
    return TanOutput(result=float(math.tan(input.a)))

# mine tool
@mcp.tool()
def mine(input: MineInput) -> MineOutput:
    """Perform a special mining calculation.
    
    Args:
        input (MineInput): Input model containing two integers for the mining calculation
        
    Returns:
        MineOutput: Output model containing the result of the mining calculation
    """
    print("CALLED: mine(a: int, b: int) -> int:")
    return MineOutput(result=int(input.a - input.b - input.b))

@mcp.tool()
def create_thumbnail(input: CreateThumbnailInput) -> CreateThumbnailOutput:
    """Create a thumbnail image from a source image.
    
    Args:
        input (CreateThumbnailInput): Input model containing the path to the source image
        
    Returns:
        CreateThumbnailOutput: Output model containing the thumbnail image data and format
    """
    print("CALLED: create_thumbnail(image_path: str) -> Image:")
    img = PILImage.open(input.image_path)
    img.thumbnail((100, 100))
    return CreateThumbnailOutput(data=img.tobytes(), format="png")

@mcp.tool()
def strings_to_chars_to_int(input: StringsToIntsInput) -> StringsToIntsOutput:
    """Convert a string to a list of ASCII values.
    
    Args:
        input (StringsToIntsInput): Input model containing the string to convert
        
    Returns:
        StringsToIntsOutput: Output model containing the list of ASCII values
    """
    print("CALLED: strings_to_chars_to_int(StringsToIntsInput) -> StringsToIntsOutput:")
    return StringsToIntsOutput(ascii_values=[int(ord(char)) for char in input.string])

@mcp.tool()
def int_list_to_exponential_sum(input: ExpSumInput) -> ExpSumOutput:
    """Calculate the sum of exponentials of numbers in a list.
    
    Args:
        input (ExpSumInput): Input model containing the list of numbers
        
    Returns:
        ExpSumOutput: Output model containing the sum of exponentials
    """
    print("CALLED: int_list_to_exponential_sum(ExpSumInput) -> ExpSumOutput:")
    return ExpSumOutput(result=sum(math.exp(i) for i in input.int_list))

@mcp.tool()
def fibonacci_numbers(input: FibonacciInput) -> FibonacciOutput:
    """Generate the first n Fibonacci numbers.
    
    Args:
        input (FibonacciInput): Input model containing the number of Fibonacci numbers to generate
        
    Returns:
        FibonacciOutput: Output model containing the sequence of Fibonacci numbers
    """
    print("CALLED: fibonacci_numbers(FibonacciInput) -> FibonacciOutput:")
    if input.n <= 0:
        return FibonacciOutput(sequence=[])
    fib_sequence = [0, 1]
    for _ in range(2, input.n):
        fib_sequence.append(fib_sequence[-1] + fib_sequence[-2])
    return FibonacciOutput(sequence=fib_sequence[:input.n])
        
@mcp.tool()
async def open_excel() -> dict:
    """Open Microsoft Excel"""
    global excel
    try:
        # excel_app = Application().start(r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')
        # time.sleep(3)
        # excel_app = win32.Dispatch("Excel.Application")
        # excel_app.Visible = True
        # excel_app.Workbooks.Add()  # Create a new blank workbook
        excel = win32.gencache.EnsureDispatch("Excel.Application")

        # Make Excel visible to the user
        excel.Visible = True

        # Add a new workbook
        workbook = excel.Workbooks.Add()
        return {
            "content": [
                TextContent(
                    type="text",
                    text="Excel opened successfully."
                )
            ]
        }
    except Exception as e:
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error opening Excel: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def merge_cells(input : mergerCellInput) -> mergerCellOutput:
    """Merge cells in a predefined range in Excel (A1:O27)"""
    global excel
    try:
        # Access the first worksheet
        worksheet = excel.ActiveSheet
        
        # Predefined range to merge (you can change this range as needed)
        range_to_merge = f"{input.starting_Cell}:{input.ending_Cell}"  # e.g., "A1:O27"
        
        # Get the range of cells to merge
        cell_range = worksheet.Range(range_to_merge)
        
        # Merge the selected range
        cell_range.Merge()

        # Optionally, apply a red border and center text in the merged cells
        cell_range.Borders.Color = 255  # Red color
        cell_range.Borders.Weight = 4   # Apply border thickness
        cell_range.HorizontalAlignment = -4108  # Center align horizontally
        cell_range.VerticalAlignment = -4108    # Center align vertically
        cell_range.Value = "Merged Cells"  # You can customize the text

        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Cells {range_to_merge} merged successfully with red border and centered text."
                )
            ]
        }
    except Exception as e:
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error merging cells: {str(e)}"
                )
            ]
        }

@mcp.tool()
async def enter_text_centered(input : enterTextCenteredInput) -> enterTextCenteredOutput:
    """Enter text in Excel and center it"""
    global excel
    try:
        worksheet = excel.ActiveSheet
        cell_range = worksheet.Range(input.input_cell)
        cell_range.Value = input.text  # Set the text

        # Center the text
        cell_range.HorizontalAlignment = -4108  # Horizontal center
        cell_range.VerticalAlignment = -4108    # Vertical center
        cell_range.Font.Size = 100  # Set font size to 72
        worksheet.Rows("1:1").RowHeight = 100  # Set row height to 72

        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Text '{input.text}' entered and centered in cell A1."
                )
            ]
        }

    except Exception as e:
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error entering text: {str(e)}"
                )
            ]
        }



@mcp.tool()
async def take_screenshot_and_send_email(input : takeScreenshotInput) -> takeScreenshotOutput:
    """Take a screenshot of the Excel window and send it via email."""
    global excel
    try:
        if not excel:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="Excel is not open. Please call open_excel first."
                    )
                ]
            }
        
        excel_win = Application(backend='uia').connect(path='EXCEL.EXE').top_window()
        excel_win.set_focus()
        time.sleep(1)

        hwnd = excel_win.handle
        rect = win32gui.GetWindowRect(hwnd)
        x, y, width, height = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]

        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        temp_file = "excel_screenshot.png"
        screenshot.save(temp_file)
        
        # Enhance the message with a custom answer (you can adjust this as needed)
        enhanced_message = f"""
{input.message}

Please see the attached screenshot for the visual representation of the Excel sheet.
"""
        
        # Send email with the screenshot
        success = send_email_with_attachment(input.recipient_email, input.subject, enhanced_message, temp_file)
        
        # Clean up the temporary file
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        if success:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text=f"Screenshot taken and email sent successfully to {input.recipient_email}"
                    )
                ]
            }
        else:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text=f"Screenshot taken but failed to send email to {input.recipient_email}"
                    )
                ]
            }
    except Exception as e:
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error taking screenshot and sending email: {str(e)}"
                )
            ]
        }

def send_email_with_attachment(recipient_email: str, subject: str, message: str, attachment_path: str) -> bool:
    """Send email with attachment using SMTP."""
    try:
        # Create message
        msg = MIMEMultipart()
        msg["From"] = os.getenv("SENDER_EMAIL")
        msg["To"] = recipient_email
        msg["Subject"] = subject

        # Create email body
        body = f"""
        {message}

        Best regards,
        Your Name.
        """

        msg.attach(MIMEText(body, "plain"))

        # Attach the screenshot
        with open(attachment_path, "rb") as attachment:
            part = MIMEImage(attachment.read())
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment_path)}",
            )
            msg.attach(part)

        # Create SMTP session
        smtp_server = os.getenv("SMTP_SERVER")
        smtp_port = int(os.getenv("SMTP_PORT"))
        sender_email = os.getenv("SENDER_EMAIL")
        sender_password = os.getenv("SENDER_PASSWORD")
        
        print(f"Connecting to SMTP server: {smtp_server}:{smtp_port}")
        server = smtplib.SMTP(smtp_server, smtp_port)
        
        print("Starting TLS...")
        server.starttls()
        
        print(f"Logging in as {sender_email}...")
        server.login(sender_email, sender_password)

        # Send email
        print("Sending email...")
        server.send_message(msg)
        server.quit()
        print("Email sent successfully!")
        return True
    except smtplib.SMTPAuthenticationError as e:
        print(f"SMTP Authentication Error: {str(e)}")
        return False
    except Exception as e:
        print(f"Error sending email: {str(e)}")
        return False

# DEFINE RESOURCES

# Add a dynamic greeting resource
@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""
    print("CALLED: get_greeting(name: str) -> str:")
    return f"Hello, {name}!"


# DEFINE AVAILABLE PROMPTS
@mcp.prompt()
def review_code(code: str) -> str:
    return f"Please review this code:\n\n{code}"
    print("CALLED: review_code(code: str) -> str:")


@mcp.prompt()
def debug_error(error: str) -> list[base.Message]:
    return [
        base.UserMessage("I'm seeing this error:"),
        base.UserMessage(error),
        base.AssistantMessage("I'll help debug that. What have you tried so far?"),
    ]

if __name__ == "__main__":
    # Check if running with mcp dev command
    print("STARTING")
    if len(sys.argv) > 1 and sys.argv[1] == "dev":
        mcp.run()  # Run without transport for dev server
    else:
        mcp.run(transport="stdio")  # Run with stdio for direct execution
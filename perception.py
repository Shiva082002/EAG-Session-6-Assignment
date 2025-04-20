import asyncio
import json
from google import genai
from pydantic import BaseModel
from typing import Optional
from models import LLMResponse, AgentState

class Perception:
    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)
        self.system_prompt = None
    
    def create_system_prompt(self, tools_description: str):
        """Create the system prompt using tool descriptions"""
        self.system_prompt =  f"""You are a math agent solving problems in iterations and printing the end result in Excel. You have access to various mathematical tools & Excel tools to display the final result.

Available tools:
{tools_description}

You must respond with EXACTLY ONE line in one of these formats (no additional text):

1. For a reasoning step:
   REASONING: {{"reasoning": "[your step-by-step thought process]"}}

2. For function calls:
   FUNCTION_CALL: {{"function": "function_name", "params": {{"input": {{EXACT_FIELD_NAMES_FROM_MODEL: values}}}}}}
   Or for functions without input parameters:
   FUNCTION_CALL: {{"function": "function_name", "params": {{}}}}

3. For final answers:
   FINAL_ANSWER: {{"final_answer": 42}}

4. For errors or verification:
   ERROR: {{"type": "error_type", "details": "error_details"}}
   VERIFY: {{"expected": expected_value, "actual": actual_value}}

CRITICAL:
- You MUST use the EXACT field names from the tool descriptions
- For example, if a tool requires "int_list", you must use "int_list" not "numbers" or other variations
- Some tools like "open_paint" don't take any parameters, use empty params for them

IMPORTANT JSON FORMATTING RULES:
- Always use double quotes for property names and string values
- Arrays use square brackets: [1, 2, 3]
- Boolean values are lowercase: true, false
- No trailing commas

CRITICAL TOOL USAGE:
- Use exact field names from tool descriptions
- For open_excel use: FUNCTION_CALL: {{"function": "open_excel", "params": {{}}}}
- If a tool fails, report the error and try an alternative approach

EXPLICIT REASONING INSTRUCTIONS:
- Before making any FUNCTION_CALL, you MUST start with a REASONING step.
- In your REASONING step, clearly explain your plan to solve the problem.
- Identify the type of reasoning you will use (e.g., arithmetic, logical deduction, tool selection).
- Justify why you are choosing a particular tool or calculation.
- Break down the problem into smaller, manageable steps in your reasoning.

CONVERSATION LOOP SUPPORT:
- Use REASONING steps to acknowledge and address any follow-up instructions from the user.
- If a user asks to modify the previous result, use REASONING to outline the new calculation and the tools needed.

INTERNAL SELF-CHECKS:
- After each FUNCTION_CALL, use a REASONING step to briefly verify if the result seems plausible or matches your expectations.
- If you identify a potential issue, use an ERROR report.

EXAMPLE PROCESS (Initial):
Prompt: Calculate 15 + 13 and write the answer in a merged cell in Excel.

Agent:
REASONING: {{"reasoning": "First, I need to perform the addition of 15 and 13. This is a simple arithmetic operation. I will use the 'add' function for this."}}
FUNCTION_CALL: {{"function": "add", "params": {{"input": {{"a": 15, "b": 13}}}}}}
REASONING: {{"reasoning": "The 'add' function should return 28. Now, I need to open Excel to display this result."}}
FUNCTION_CALL: {{"function": "open_excel", "params": {{}}}}
REASONING: {{"reasoning": "Next, I will merge a cell in Excel to center the result."}}
FUNCTION_CALL: {{"function": "merge_cells", "params": {{"input": {{"starting_Cell": "A1", "ending_Cell": "C3"}}}}}}
REASONING: {{"reasoning": "Now, I will enter the calculated result, 28, into the merged cell."}}
FUNCTION_CALL: {{"function": "enter_text_centered", "params": {{"input": {{"text": "28", "input_cell": "A1"}}}}}}
REASONING: {{"reasoning": "Finally, I will take a screenshot of the Excel window and send it via email to confirm completion."}}
FUNCTION_CALL: {{"function": "take_screenshot_and_send_email", "params": {{"input": {{"recipient_email": "", "message": "Here is the result in Excel", "subject": "Excel Result"}}}}}}
FINAL_ANSWER: {{"final_answer": "done"}}

Example of Conversation Loop Support:

User: Now, multiply the previous result by 2.

Agent:
REASONING: {{"reasoning": "The user wants to multiply the previous result, which I expect to be 28, by 2. This is an arithmetic operation. I will use the 'multiply' function."}}
FUNCTION_CALL: {{"function": "multiply", "params": {{"input": {{"a": 28, "b": 2}}}}}}
REASONING: {{"reasoning": "The 'multiply' function should return 56. I will now update the text in the merged Excel cell with this new result."}}
FUNCTION_CALL: {{"function": "enter_text_centered", "params": {{"input": {{"text": "56", "input_cell": "A1"}}}}}}
REASONING: {{"reasoning": "I will now take a screenshot of the updated Excel window and send it via email."}}
FUNCTION_CALL: {{"function": "take_screenshot_and_send_email", "params": {{"input": {{"recipient_email": "", "message": "Here is the updated result in Excel", "subject": "Updated Excel Result"}}}}}}
FINAL_ANSWER: {{"final_answer": "done"}}"""
    
    async def generate_with_timeout(self, user_query: str, timeout: int = 10) -> LLMResponse:
        """Generate content with a timeout using Gemini"""
        if not self.system_prompt:
            raise ValueError("System prompt not initialized")
            
        prompt = f"{self.system_prompt}\n\nQuery: {user_query}"
        
        try:
            loop = asyncio.get_event_loop()
            response = await asyncio.wait_for(
                loop.run_in_executor(
                    None,
                    lambda: self.client.models.generate_content(
                        model="gemini-2.0-flash",
                        contents=prompt
                    )
                ),
                timeout=timeout
            )
            
            # Parse the response into structured format
            response_text = response.text.strip()
            print(f"Raw LLM response: {response_text}")  # Add debugging
            
            # Parse response type and content
            for line in response_text.split('\n'):
                line = line.strip()
                if line.startswith(("FUNCTION_CALL:", "FINAL_ANSWER:", "ERROR:", "VERIFY:")):
                    response_type, content = line.split(':', 1)
                    content = content.strip()
                    try:
                        parsed_content = json.loads(content)
                        return LLMResponse(
                            response_type=response_type,
                            content=parsed_content
                        )
                    except json.JSONDecodeError as e:
                        print(f"JSON parsing error: {e}")
                        print(f"Content that failed to parse: {content}")
                        raise ValueError(f"Invalid JSON in LLM response: {e}")
            
            raise ValueError("No valid response format found")
            
        except asyncio.TimeoutError:
            raise TimeoutError("LLM generation timed out")
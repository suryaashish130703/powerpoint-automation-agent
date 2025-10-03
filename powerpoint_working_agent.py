import os
import time
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters, types
from mcp.client.stdio import stdio_client
import asyncio
from google import genai
from concurrent.futures import TimeoutError
from functools import partial
from email_logger import EmailLogger

# Load environment variables from .env file
load_dotenv()

# Access your API key and initialize Gemini client correctly
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key)

max_iterations = 10  # For PowerPoint operations
last_response = None
iteration = 0
iteration_response = []

# Initialize email logger
email_logger = EmailLogger()
execution_logs = []
start_time = None

async def generate_with_timeout(client, prompt, timeout=10):
    """Generate content with a timeout"""
    print("Starting LLM generation...")
    try:
        # Convert the synchronous generate_content call to run in a thread
        loop = asyncio.get_event_loop()
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None, 
                lambda: client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=prompt
                )
            ),
            timeout=timeout
        )
        print("LLM generation completed")
        return response
    except TimeoutError:
        print("LLM generation timed out!")
        raise
    except Exception as e:
        print(f"Error in LLM generation: {e}")
        raise

def log_message(message, log_type="INFO"):
    """Log a message to both console and email logs"""
    global execution_logs
    timestamp = time.strftime("%H:%M:%S")
    formatted_message = f"[{timestamp}] {log_type}: {message}"
    print(formatted_message)
    execution_logs.append(formatted_message)

def log_raw_output(message):
    """Log raw terminal output to email logs (without timestamp)"""
    global execution_logs
    print(message)
    execution_logs.append(message)

def reset_state():
    """Reset all global variables to their initial state"""
    global last_response, iteration, iteration_response, execution_logs, start_time
    last_response = None
    iteration = 0
    iteration_response = []
    execution_logs = []
    start_time = None

def print_iteration_header(iteration_num, action):
    """Print a formatted iteration header"""
    print(f"\n{'='*60}")
    print(f"ITERATION {iteration_num}: {action}")
    print(f"{'='*60}")

async def main():
    global start_time
    reset_state()  # Reset at the start of main
    start_time = time.time()
    
    log_message("Starting Working PowerPoint Agent Execution...", "INFO")
    log_message("This agent will solve a math problem and automatically visualize the result in PowerPoint", "INFO")
    
    try:
        # Create a single MCP server connection to Working PowerPoint server
        log_message("Establishing connection to Working PowerPoint MCP server...", "INFO")
        server_params = StdioServerParameters(
            command="python",
            args=["powerpoint_working_mcp_server.py"]
        )

        async with stdio_client(server_params) as (read, write):
            log_message("Connection established, creating session...", "SUCCESS")
            async with ClientSession(read, write) as session:
                log_message("Session created, initializing...", "INFO")
                await session.initialize()
                
                # Get available tools
                log_message("Requesting tool list...", "INFO")
                tools_result = await session.list_tools()
                tools = tools_result.tools
                log_message(f"Successfully retrieved {len(tools)} tools", "SUCCESS")
                
                # Create system prompt with available tools
                log_message("Creating system prompt...", "INFO")
                
                try:
                    tools_description = []
                    for i, tool in enumerate(tools):
                        try:
                            # Get tool properties
                            params = tool.inputSchema
                            desc = getattr(tool, 'description', 'No description available')
                            name = getattr(tool, 'name', f'tool_{i}')
                            
                            # Format the input schema in a more readable way
                            if 'properties' in params:
                                param_details = []
                                for param_name, param_info in params['properties'].items():
                                    param_type = param_info.get('type', 'unknown')
                                    param_details.append(f"{param_name}: {param_type}")
                                params_str = ', '.join(param_details)
                            else:
                                params_str = 'no parameters'

                            tool_desc = f"{i+1}. {name}({params_str}) - {desc}"
                            tools_description.append(tool_desc)
                        except Exception as e:
                            print(f"Error processing tool {i}: {e}")
                            tools_description.append(f"{i+1}. Error processing tool")
                    
                    tools_description = "\n".join(tools_description)
                    log_message("Successfully created tools description", "SUCCESS")
                except Exception as e:
                    log_message(f"Error creating tools description: {e}", "ERROR")
                    tools_description = "Error loading tools"
                
                system_prompt = f"""You are a math agent solving problems in iterations. You have access to various mathematical tools and PowerPoint automation tools.

Available tools:
{tools_description}

You must respond with EXACTLY ONE line in one of these formats (no additional text):
1. For function calls:
   FUNCTION_CALL: function_name|param1|param2|...
   
2. For final answers:
   FINAL_ANSWER: [number]

Important:
- When a function returns multiple values, you need to process all of them
- Only give FINAL_ANSWER when you have completed all necessary calculations
- Do not repeat function calls with the same parameters
- For PowerPoint operations, follow this complete sequence: open_powerpoint -> select_rectangle_shape -> draw_rectangle_centered -> select_text_box -> click_inside_rectangle -> paste_number
- You must complete ALL steps in the sequence before giving FINAL_ANSWER

Examples:
- FUNCTION_CALL: add|5|3
- FUNCTION_CALL: strings_to_chars_to_int|INDIA
- FUNCTION_CALL: int_list_to_exponential_sum|[73,78,68,73,65]
- FUNCTION_CALL: open_powerpoint
- FUNCTION_CALL: draw_rectangle_centered
- FINAL_ANSWER: [42]

DO NOT include any explanations or additional text.
Your entire response should be a single line starting with either FUNCTION_CALL: or FINAL_ANSWER:"""

                query = """Find the ASCII values of characters in INDIA and then return sum of exponentials of those values. After getting the final answer, open PowerPoint, draw a rectangle, and write the result inside it."""
                log_message("Starting iteration loop...", "INFO")
                
                # Use global iteration variables
                global iteration, last_response
                
                while iteration < max_iterations:
                    log_raw_output(f"\n{'='*60}")
                    log_raw_output(f"ITERATION {iteration + 1}: Processing")
                    log_raw_output(f"{'='*60}")
                    
                    if last_response is None:
                        current_query = query
                    else:
                        current_query = current_query + "\n\n" + " ".join(iteration_response)
                        current_query = current_query + "  What should I do next?"

                    # Get model's response with timeout
                    log_raw_output("Preparing to generate LLM response...")
                    prompt = f"{system_prompt}\n\nQuery: {current_query}"
                    try:
                        response = await generate_with_timeout(client, prompt)
                        response_text = response.text.strip()
                        log_raw_output(f"LLM Response: {response_text}")
                        
                        # Find the FUNCTION_CALL line in the response
                        for line in response_text.split('\n'):
                            line = line.strip()
                            if line.startswith("FUNCTION_CALL:"):
                                response_text = line
                                break
                        
                    except Exception as e:
                        print(f"Failed to get LLM response: {e}")
                        break

                    if response_text.startswith("FUNCTION_CALL:"):
                        _, function_info = response_text.split(":", 1)
                        parts = [p.strip() for p in function_info.split("|")]
                        func_name, params = parts[0], parts[1:]
                        
                        log_raw_output(f"Calling function: {func_name}")
                        log_raw_output(f"Parameters: {params}")
                        
                        try:
                            # Find the matching tool to get its input schema
                            tool = next((t for t in tools if t.name == func_name), None)
                            if not tool:
                                log_raw_output(f"Available tools: {[t.name for t in tools]}")
                                raise ValueError(f"Unknown tool: {func_name}")

                            # Prepare arguments according to the tool's input schema
                            arguments = {}
                            schema_properties = tool.inputSchema.get('properties', {})

                            for param_name, param_info in schema_properties.items():
                                if not params:  # Check if we have enough parameters
                                    raise ValueError(f"Not enough parameters provided for {func_name}")
                                    
                                value = params.pop(0)  # Get and remove the first parameter
                                param_type = param_info.get('type', 'string')
                                
                                # Convert the value to the correct type based on the schema
                                if param_type == 'integer':
                                    arguments[param_name] = int(value)
                                elif param_type == 'number':
                                    arguments[param_name] = float(value)
                                elif param_type == 'array':
                                    # Handle array input
                                    if isinstance(value, str):
                                        # Remove brackets and split by comma
                                        value = value.strip('[]').split(',')
                                        # Filter out empty strings and convert to int
                                        arguments[param_name] = [int(x.strip()) for x in value if x.strip()]
                                    else:
                                        arguments[param_name] = value
                                else:
                                    arguments[param_name] = str(value)

                            log_raw_output(f"Final arguments: {arguments}")
                            
                            result = await session.call_tool(func_name, arguments=arguments)
                            
                            # Get the full result content
                            if hasattr(result, 'content'):
                                if isinstance(result.content, list):
                                    iteration_result = [
                                        item.text if hasattr(item, 'text') else str(item)
                                        for item in result.content
                                    ]
                                else:
                                    iteration_result = str(result.content)
                            else:
                                iteration_result = str(result)
                            
                            # Format the response based on result type
                            if isinstance(iteration_result, list):
                                result_str = f"[{', '.join(iteration_result)}]"
                            else:
                                result_str = str(iteration_result)
                            
                            log_raw_output(f"Function result: {result_str}")
                            
                            iteration_response.append(
                                f"In iteration {iteration + 1} you called {func_name} with {arguments} parameters, "
                                f"and the function returned {result_str}."
                            )
                            last_response = iteration_result

                        except Exception as e:
                            log_raw_output(f"Error details: {str(e)}")
                            import traceback
                            traceback.print_exc()
                            iteration_response.append(f"Error in iteration {iteration + 1}: {str(e)}")
                            break

                    elif response_text.startswith("FINAL_ANSWER:"):
                        log_raw_output("\n=== Agent Execution Complete ===")
                        log_raw_output(f"Final Answer: {response_text}")
                        
                        # Extract the final number for PowerPoint
                        final_number = response_text.replace("FINAL_ANSWER:", "").strip()
                        log_raw_output(f"Final number to display: {final_number}")
                        
                        # Now automatically perform PowerPoint operations following your exact workflow
                        log_raw_output("\n=== AUTOMATIC POWERPOINT WORKFLOW STARTING ===")
                        
                        # Step 1: Open PowerPoint
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 1: Opening PowerPoint and Creating New Presentation")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("open_powerpoint")
                        log_raw_output(result.content[0].text)

                        # Step 2: Select Rectangle Shape
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 2: Selecting Rectangle Shape (Insert -> Shapes -> Rectangle)")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("select_rectangle_shape")
                        log_raw_output(result.content[0].text)

                        # Step 3: Draw Rectangle Centered
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 3: Drawing Rectangle Centered on Slide")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("draw_rectangle_centered")
                        log_raw_output(result.content[0].text)

                        # Step 4: Select Text Box
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 4: Selecting Text Box (Insert -> Text Box)")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("select_text_box")
                        log_raw_output(result.content[0].text)

                        # Step 5: Click Inside Rectangle
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 5: Clicking Inside Rectangle Area to Place Text Box")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("click_inside_rectangle")
                        log_raw_output(result.content[0].text)

                        # Step 6: Paste the Generated Number
                        log_raw_output(f"\n{'='*60}")
                        log_raw_output("ITERATION 6: Pasting Generated Number Inside Rectangle")
                        log_raw_output(f"{'='*60}")
                        result = await session.call_tool("paste_number", arguments={
                            "text": final_number
                        })
                        log_raw_output(result.content[0].text)
                        
                        log_message("AUTOMATIC POWERPOINT WORKFLOW COMPLETE", "SUCCESS")
                        log_message("PowerPoint should now be open with a slide containing a rectangle and the number inside it!", "SUCCESS")
                        log_message("Check your PowerPoint window to see the result.", "INFO")
                        
                        # Send success email with logs
                        execution_time = time.time() - start_time
                        log_message(f"Sending success email with execution logs...", "INFO")
                        email_logger.send_success_email(final_number, execution_time, execution_logs)
                        
                        break

                    iteration += 1

    except Exception as e:
        log_message(f"Error in main execution: {e}", "ERROR")
        import traceback
        traceback.print_exc()
        
        # Send error email with logs
        log_message("Sending error email with execution logs...", "ERROR")
        email_logger.send_error_email(str(e), execution_logs)
        
    finally:
        reset_state()  # Reset at the end of main

if __name__ == "__main__":
    asyncio.run(main())
 
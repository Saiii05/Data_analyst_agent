import os
import json
import traceback
from io import BytesIO, StringIO
import sys

# Third-party libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import requests
import docx # For python-docx
import PyPDF2
from PIL import Image, ImageDraw, ImageFont # Added ImageDraw, ImageFont
# import openai # If using openai client for TogetherAI - Commented out as not used directly
import pytesseract
from reportlab.pdfgen import canvas # For sample PDF creation
from reportlab.lib.pagesizes import letter # For sample PDF creation


# Configure Matplotlib for inline plotting in Jupyter
# %matplotlib inline

print("Libraries imported successfully.")
# --- Code Cell ---
# Configuration for Together.ai API
# The API key will be sourced from the TOGETHER_API_KEY environment variable if it's set.
# Otherwise, it will use the default key provided below. For security, using an environment variable is recommended.
TOGETHER_API_KEY = os.environ.get("TOGETHER_API_KEY", "17d95aea878cfc0bf5c0fa8530f59e8b2a9a03013375c630b44d7f6d50902dd7") # Uses environment variable if set, otherwise defaults to the provided key.
MODEL_NAME = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"

if not TOGETHER_API_KEY:
    # This condition will likely not be met if a default is provided, unless the default is explicitly None or empty.
    # However, keeping the structure for clarity if the default is removed later.
    print("TOGETHER_API_KEY environment variable not set and no default key available. Please set it before proceeding.")
elif TOGETHER_API_KEY == "17d95aea878cfc0bf5c0fa8530f59e8b2a9a03013375c630b44d7f6d50902dd7" and not os.environ.get("TOGETHER_API_KEY"):
    print(f"Using default Together.ai API Key. Using model: {MODEL_NAME}")
    print("It's recommended to set your own TOGETHER_API_KEY environment variable for security and to avoid interruptions.")
else: # This means os.environ.get("TOGETHER_API_KEY") is set and is not the default.
    print(f"Together.ai API Key loaded from environment variable. Using model: {MODEL_NAME}")

# Test Together.ai API Key and Model (Optional - can be a separate step/function later)
# This is a placeholder for now. Actual API call will be in the LLM interaction module.
# print("Attempting a test API call...")
# headers = {
#     "Authorization": f"Bearer {TOGETHER_API_KEY}",
#     "Content-Type": "application/json"
# }
# data = {
#     "model": MODEL_NAME,
#     "prompt": "Hello, who are you?",
#     "max_tokens": 50
# }
# try:
#     response = requests.post("https://api.together.xyz/inference", headers=headers, json=data)
#     if response.status_code == 200:
#         print("Test API call successful.")
#         # print("Response:", response.json())
#     else:
#         print(f"Test API call failed. Status Code: {response.status_code}")
#         print("Response Text:", response.text)
# except requests.exceptions.RequestException as e:
#     print(f"Test API call failed with an exception: {e}")
# --- Code Cell ---
# Cell for File Ingestion Module

# Ensure pytesseract is installed if you plan to use it.
# You might need to install Tesseract OCR engine on the system as well.
# # !pip install pytesseract
# import pytesseract # Moved to top
# from PIL import Image # Pillow for image handling - already imported

# For DOCX
# import docx # Already imported

# For PDF
# import PyPDF2 # Already imported

# For CSV/Excel
# import pandas as pd # Already imported

# import os # Already imported

def load_txt(file_path):
    """Reads text from a .txt file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        return f"Error reading TXT file: {e}"

def load_docx(file_path):
    """Reads text from a .docx file."""
    try:
        doc_obj = docx.Document(file_path) # Renamed to avoid conflict with imported module
        full_text = []
        for para in doc_obj.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        return f"Error reading DOCX file: {e}"

def load_pdf(file_path):
    """Reads text from a .pdf file."""
    try:
        text = ""
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text += page.extract_text() or "" # Add empty string if None
        return text
    except Exception as e:
        return f"Error reading PDF file: {e}"

def load_csv_excel(file_path):
    """Reads data from .csv or .xlsx files into a Pandas DataFrame."""
    try:
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)
        elif file_path.endswith(('.xls', '.xlsx')):
            return pd.read_excel(file_path)
    except Exception as e:
        return f"Error reading CSV/Excel file: {e}"

def load_image_ocr(file_path):
    """Loads an image and performs OCR to extract text using pytesseract."""
    try:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        # Fallback if pytesseract is not configured or fails
        return f"Error performing OCR: {e}. Make sure Tesseract OCR is installed and configured."


def load_document(file_path):
    """
    Loads a document based on its file extension.
    Returns the content (text, DataFrame, or OCRed text).
    Returns a string error message if loading fails or file type is unsupported.
    """
    _, file_extension = os.path.splitext(file_path.lower())

    if not os.path.exists(file_path):
        return "Error: File not found."

    if file_extension == '.txt':
        return load_txt(file_path)
    elif file_extension == '.docx':
        return load_docx(file_path)
    elif file_extension == '.pdf':
        return load_pdf(file_path)
    elif file_extension == '.csv':
        # For CSV, load_csv_excel returns a DataFrame
        content = load_csv_excel(file_path)
        if isinstance(content, pd.DataFrame):
            return content
        else: # Error message from load_csv_excel
            return content
    elif file_extension in ['.xls', '.xlsx']:
        # For Excel, load_csv_excel returns a DataFrame
        content = load_csv_excel(file_path)
        if isinstance(content, pd.DataFrame):
            return content
        else: # Error message from load_csv_excel
            return content
    elif file_extension in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff']:
        return load_image_ocr(file_path)
    else:
        return "Error: Unsupported file type."

# Example Usage (comment out before final run or use example files)
# print("Testing file loaders...")
# Create dummy files for testing
# with open("sample.txt", "w") as f: f.write("This is a test text file.")
# doc = docx.Document()
# doc.add_paragraph("This is a test Word document.")
# doc.save("sample.docx")
# df_sample = pd.DataFrame({'col1': [1, 2], 'col2': [3, 4]})
# df_sample.to_csv("sample.csv", index=False)

# print("TXT:", load_document("sample.txt"))
# print("DOCX:", load_document("sample.docx"))
# print("CSV:", load_document("sample.csv"))
# For PDF and Image, manual creation is more complex, assume files exist or test separately.
# print("PDF:", load_document("sample.pdf")) # Needs a sample.pdf
# print("Image:", load_document("sample.png")) # Needs a sample.png and Tesseract setup

print("File ingestion module defined.")
# --- Code Cell ---
# Cell for LLM Interaction Module

# import requests # Already imported
# import json # Already imported
# import os # Already imported

# Ensure API key and model name are accessible (they are defined in a previous cell)
# If not, you might need to re-fetch them or pass them explicitly.
# TOGETHER_API_KEY = os.environ.get("TOGETHER_API_KEY_FALLBACK") # Or your actual key
# MODEL_NAME = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"


def query_llm(prompt_text, max_tokens=1024, temperature=0.7, top_p=0.7, top_k=50, repetition_penalty=1):
    """
    Queries the Together.ai LLM with the given prompt and parameters.

    Args:
        prompt_text (str): The prompt to send to the LLM.
        max_tokens (int): Maximum number of tokens to generate.
        temperature (float): Controls randomness. Lower is more deterministic.
        top_p (float): Nucleus sampling parameter.
        top_k (int): Top-k sampling parameter.
        repetition_penalty (float): Penalty for repeating tokens.

    Returns:
        str: The LLM's generated text, or an error message if the API call fails.
    """
    global TOGETHER_API_KEY, MODEL_NAME # Accessing global vars defined in setup

    if not TOGETHER_API_KEY:
        return "Error: TOGETHER_API_KEY is not set."
    if not MODEL_NAME:
        return "Error: MODEL_NAME is not set."

    endpoint = 'https://api.together.xyz/v1/chat/completions' # Or /inference for older models/formats

    headers = {
        "Authorization": f"Bearer {TOGETHER_API_KEY}",
        "Content-Type": "application/json"
    }

    # Using the chat completions endpoint structure
    data = {
        "model": MODEL_NAME,
        "messages": [{"role": "user", "content": prompt_text}],
        "max_tokens": max_tokens,
        "temperature": temperature,
        "top_p": top_p,
        "top_k": top_k,
        "repetition_penalty": repetition_penalty,
        # "stream_tokens": False # Set to True if you want to stream, but handling would change
    }

    try:
        response = requests.post(endpoint, headers=headers, json=data)
        response.raise_for_status() # Raises an HTTPError for bad responses (4XX or 5XX)

        response_json = response.json()

        # Accessing the content from the chat completions response structure
        if response_json.get("choices") and len(response_json["choices"]) > 0:
            # Check for message content
            message = response_json["choices"][0].get("message")
            if message and message.get("content"):
                 return message["content"].strip()
            # Fallback for older/different structures if needed (less likely for chat/completions)
            elif response_json["choices"][0].get("text"):
                 return response_json["choices"][0]["text"].strip()
            else:
                return "Error: Could not extract text from LLM response. Structure might have changed."
        else:
            return f"Error: LLM response does not contain 'choices' or 'choices' is empty. Full response: {response_json}"

    except requests.exceptions.HTTPError as http_err:
        # Attempt to get more details from the response body if available
        error_detail = response.text if response else "No response body"
        return f"HTTP error occurred: {http_err}. Response: {error_detail}"
    except requests.exceptions.RequestException as req_err:
        return f"Request error occurred: {req_err}"
    except json.JSONDecodeError:
        return f"Error decoding JSON response. Response text: {response.text if response else 'No response text'}"
    except Exception as e:
        return f"An unexpected error occurred: {e}"

# Example Usage (comment out or remove for final notebook)
# print("Testing LLM query function...")
# Make sure TOGETHER_API_KEY is set and valid
# if TOGETHER_API_KEY and MODEL_NAME:
#     test_prompt = "Explain what a large language model is in one sentence."
#     llm_response = query_llm(test_prompt)
#     print(f"Prompt: {test_prompt}")
#     print(f"LLM Response: {llm_response}")
# else:
#     print("Skipping LLM test because API key or model name is not set.")

print("LLM interaction module defined.")
# --- Code Cell ---
# Cell for Core Agent Logic - DataAnalysisAgent Class

# import pandas as pd # Already imported
# import matplotlib.pyplot as plt # Already imported
# import seaborn as sns # Already imported
# import traceback # Moved to top

# Assuming load_document and query_llm are defined in previous cells
# and TOGETHER_API_KEY, MODEL_NAME are globally available or passed appropriately.

class DataAnalysisAgent:
    def __init__(self):
        """
        Initializes the DataAnalysisAgent.
        Relies on global TOGETHER_API_KEY and MODEL_NAME for LLM interaction.
        """
        self.data = None  # To store loaded data (e.g., DataFrame, text)
        self.data_type = None # 'tabular', 'text', 'image_text'
        self.data_summary = None # A summary/context of the loaded data
        self.conversation_history = []

        # Ensure API key and model are loaded from global scope (defined in setup cell)
        global TOGETHER_API_KEY, MODEL_NAME
        if not TOGETHER_API_KEY or not MODEL_NAME:
            raise ValueError("TOGETHER_API_KEY or MODEL_NAME is not set. Please run the setup cell.")
        self.api_key = TOGETHER_API_KEY
        self.model_name = MODEL_NAME

    def upload_and_prepare_data(self, file_path):
        """
        Loads data from the given file_path, determines its type,
        and generates an initial summary.
        """
        self.conversation_history = [] # Reset history for new data
        loaded_content = load_document(file_path)

        if isinstance(loaded_content, str) and loaded_content.startswith("Error:"):
            self.data = None
            self.data_type = None
            self.data_summary = None
            return loaded_content # Return error message

        if isinstance(loaded_content, pd.DataFrame):
            self.data = loaded_content
            self.data_type = 'tabular'
            # Generate a summary for tabular data
            info_str = self.data.info(verbose=False, buf=None) # Get info as string
            if info_str is None: # pandas >= 1.0.0 info() prints to sys.stdout by default
                from io import StringIO
                buffer = StringIO()
                self.data.info(buf=buffer)
                info_str = buffer.getvalue()

            self.data_summary = (
                f"Loaded tabular data with {self.data.shape[0]} rows and {self.data.shape[1]} columns.\n"
                f"Columns: {', '.join(self.data.columns)}\n"
                f"Data types and non-null counts:\n{info_str}\n"
                f"First 3 rows:\n{self.data.head(3).to_string()}"
            )
            return "Tabular data loaded successfully. Summary generated."
        elif isinstance(loaded_content, str):
            self.data = loaded_content
            # Heuristic: if it's from OCR or a text file
            # A bit simplistic, might need refinement based on actual load_document outputs for images
            if file_path and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                 self.data_type = 'image_text' # Specifically from OCR
            elif file_path and file_path.endswith(('.txt', '.pdf', '.docx')):
                 self.data_type = 'text'
            else: # Default for other string data or if file_path is None (e.g. direct text input)
                 self.data_type = 'text'

            # Generate a summary for text data (could use LLM for longer texts)
            if len(self.data) > 500:
                summary_prompt = f"Provide a concise summary of the following text (max 100 words):\n\n{self.data[:1000]}"
                llm_gen_summary = query_llm(summary_prompt)
                if llm_gen_summary.startswith("Error:"):
                     self.data_summary = f"{self.data_type.capitalize()} data loaded. Length: {len(self.data)} characters. First 200 chars: {self.data[:200]}"
                else:
                    self.data_summary = f"{self.data_type.capitalize()} data loaded. LLM Summary: {llm_gen_summary}"
            else:
                self.data_summary = f"{self.data_type.capitalize()} data loaded. Content: {self.data[:200]}"
            return f"{self.data_type.capitalize()} data loaded successfully. Summary generated."
        else:
            self.data = None
            self.data_type = None
            self.data_summary = None
            return "Error: Unknown data type after loading."

    def _get_context_for_prompt(self):
        if not self.data_summary:
            return "No data loaded yet. Please upload a file first."

        history_str = "\n".join([f"{msg['role']}: {msg['content']}" for msg in self.conversation_history[-5:]]) # Last 5 exchanges
        return f"Current Data Context:\n{self.data_summary}\n\nConversation History (last 5 exchanges):\n{history_str}\n\nUser Question:"

    def ask_question(self, question_text):
        """
        Asks a question about the loaded data, using conversation history.
        """
        if not self.data_summary:
            return "Please upload and prepare data first using `upload_and_prepare_data('file_path')`."

        prompt_context = self._get_context_for_prompt()
        full_prompt = f"{prompt_context} {question_text}"

        # Add user question to history
        self.conversation_history.append({"role": "user", "content": question_text})

        llm_response = query_llm(full_prompt)

        if not llm_response.startswith("Error:"):
            # Add LLM response to history
            self.conversation_history.append({"role": "assistant", "content": llm_response})
        return llm_response

    def perform_analysis(self, analysis_request_text):
        """
        Performs data analysis based on the request.
        For tabular data, it asks the LLM to generate Python/Pandas code and executes it.
        For text data, it asks the LLM to perform the analysis directly.
        """
        if not self.data_summary:
            return "Please upload and prepare data first."

        self.conversation_history.append({"role": "user", "content": f"Perform analysis: {analysis_request_text}"})

        if self.data_type == 'tabular':
            prompt = (
                f"Current Data Context:\n{self.data_summary}\n\n"
                f"You are a data analysis assistant. Based on the data context, write Python code using Pandas "
                f"to perform the following analysis: '{analysis_request_text}'.\n"
                f"The data is available in a Pandas DataFrame called `df`.\n"
                f"The code should:\n"
                f"1. Perform the analysis.\n"
                f"2. Store the primary result in a variable called `analysis_result` (this could be a DataFrame, Series, string, or number).\n"
                f"3. If the result is a DataFrame or Series, print its string representation. If it's a scalar, print it.\n"
                f"4. Do NOT include any example DataFrame creation (e.g., `df = pd.DataFrame(...)`). Assume `df` is already loaded.\n"
                f"5. Only output the Python code block, without any explanations before or after the ```python ... ``` block.\n"
                f"Example of desired output format:\n"
                f"```python\n"
                f"# Code for analysis\n"
                f"analysis_result = df['some_column'].mean()\n"
                f"print(analysis_result)\n"
                f"```"
            )

            code_response = query_llm(prompt, max_tokens=500)

            if code_response.startswith("Error:"):
                self.conversation_history.append({"role": "assistant", "content": f"Error from LLM for code gen: {code_response}"})
                return f"Error getting analysis code from LLM: {code_response}"

            # Extract code from markdown block if present
            if "```python" in code_response:
                python_code = code_response.split("```python")[1].split("```")[0].strip()
            elif "```" in code_response: # a plain ``` block
                python_code = code_response.split("```")[1].split("```")[0].strip()
            else: # Assume raw code if no triple backticks (less ideal)
                python_code = code_response

            # Prepare a local scope for exec, including the DataFrame
            local_scope = {'df': self.data.copy(), 'pd': pd, 'np': np, 'analysis_result': None}

            final_output = ""
            try:
                # Capture print output from exec
                # from io import StringIO # Moved to top
                # import sys # Moved to top
                # from io import StringIO # Moved to top
                # import sys # Moved to top
                old_stdout = sys.stdout
                redirected_output = StringIO()
                sys.stdout = redirected_output

                exec(python_code, {'pd': pd, 'np': np}, local_scope)

                sys.stdout = old_stdout # Restore stdout

                printed_output = redirected_output.getvalue()
                result_variable = local_scope.get('analysis_result')

                final_output = "Analysis Code Executed.\n"
                if printed_output:
                    final_output += f"Printed output:\n{printed_output}\n"

                if result_variable is not None:
                    final_output += f"Result variable content:\n{str(result_variable)}"
                elif not printed_output:
                    final_output += "No specific result variable captured, and no printed output from code."
                self.conversation_history.append({"role": "assistant", "content": f"Code execution successful. Output: {final_output}"})
                return final_output

            except Exception as e:
                error_message = f"Error executing analysis code: {e}\nTraceback:\n{traceback.format_exc()}\nCode was:\n{python_code}"
                self.conversation_history.append({"role": "assistant", "content": f"Code execution failed. Error: {error_message}"})
                return error_message

        elif self.data_type == 'text' or self.data_type == 'image_text':
            prompt = (
                f"Current Data Context:\n{self.data_summary}\n\n"
                f"Based on the text data, perform the following analysis: '{analysis_request_text}'.\n"
                f"Provide a comprehensive answer."
            )
            llm_response = query_llm(prompt)
            self.conversation_history.append({"role": "assistant", "content": llm_response})
            return llm_response
        else:
            no_support_msg = "Analysis is not supported for the current data type."
            self.conversation_history.append({"role": "assistant", "content": no_support_msg})
            return no_support_msg

    def generate_visualization(self, plot_request_text):
        """
        Generates a visualization based on the request.
        Currently only supports tabular data and asks LLM to generate Matplotlib/Seaborn code.
        The code should save the plot to 'plot.png'.
        """
        if self.data_type != 'tabular':
            return "Visualization is primarily supported for tabular data. Please load a CSV or Excel file."

        self.conversation_history.append({"role": "user", "content": f"Generate visualization: {plot_request_text}"})

        prompt = (
            f"Current Data Context:\n{self.data_summary}\n\n"
            f"You are a data visualization assistant. Based on the data context, write Python code using Matplotlib or Seaborn "
            f"to generate a plot for the following request: '{plot_request_text}'.\n"
            f"The data is available in a Pandas DataFrame called `df`.\n"
            f"The code should:\n"
            f"1. Generate the plot.\n"
            f"2. Save the plot to a file named 'plot.png'.\n"
            f"3. Ensure the plot is complete and self-contained (e.g., include `plt.show()` if interactive, but saving is key).\n"
            f"4. Do NOT include any example DataFrame creation (e.g., `df = pd.DataFrame(...)`). Assume `df` is already loaded.\n"
            f"5. Only output the Python code block, without any explanations before or after the ```python ... ``` block.\n"
            f"Example of desired output format for saving a plot:\n"
            f"```python\n"
            f"import matplotlib.pyplot as plt\n"
            f"import seaborn as sns\n"
            f"# Code for plotting\n"
            f"plt.figure(figsize=(10, 6))\n"
            f"sns.histplot(data=df, x='some_column')\n"
            f"plt.title('Histogram of Some Column')\n"
            f"plt.xlabel('Value')\n"
            f"plt.ylabel('Frequency')\n"
            f"plt.savefig('plot.png')\n"
            f"# plt.show() # Optional for interactive, but savefig is crucial\n"
            f"```"
        )

        code_response = query_llm(prompt, max_tokens=600)

        if code_response.startswith("Error:"):
            self.conversation_history.append({"role": "assistant", "content": f"Error from LLM for viz code gen: {code_response}"})
            return f"Error getting visualization code from LLM: {code_response}"

        # Extract code from markdown block
        if "```python" in code_response:
            python_code = code_response.split("```python")[1].split("```")[0].strip()
        elif "```" in code_response:
            python_code = code_response.split("```")[1].split("```")[0].strip()
        else:
            python_code = code_response

        # Prepare a local scope for exec, including the DataFrame and plotting libraries
        local_scope = {'df': self.data.copy(), 'pd': pd, 'np': np, 'plt': plt, 'sns': sns, 'os': os}
        plot_file_name = 'plot.png'

        try:
            # Ensure any previous plots are cleared
            plt.close('all')
            exec(python_code, {'pd': pd, 'np': np, 'plt': plt, 'sns': sns, 'os': os}, local_scope)

            if os.path.exists(plot_file_name):
                success_msg = f"Visualization generated and saved to {plot_file_name}"
                self.conversation_history.append({"role": "assistant", "content": success_msg})
                return success_msg
            else:
                no_file_msg = f"Visualization code executed, but '{plot_file_name}' was not created. The LLM might have failed to include the save command or there was an issue in the generated code. \nCode was:\n{python_code}"
                self.conversation_history.append({"role": "assistant", "content": no_file_msg})
                return no_file_msg
        except Exception as e:
            error_message = f"Error executing visualization code: {e}\nTraceback:\n{traceback.format_exc()}\nCode was:\n{python_code}"
            self.conversation_history.append({"role": "assistant", "content": f"Viz code execution failed. Error: {error_message}"})
            return error_message

print("DataAnalysisAgent class defined.")

# --- Code Cell ---
# Cell for Creating Sample Files for Demonstration

import pandas as pd
import docx
# from PyPDF2 import PdfWriter, PageObject # Not used directly for text content
from PIL import Image, ImageDraw, ImageFont # For creating a sample image

# Create sample TXT file
with open("sample.txt", "w", encoding="utf-8") as f:
    f.write("This is a sample text file for demonstration purposes. It contains some general information about data analysis and large language models. The agent should be able to read this and answer questions about its content.")
print("sample.txt created.")

# Create sample DOCX file
doc = docx.Document()
doc.add_heading("Sample Document Title", 0)
doc.add_paragraph("This is a paragraph in a sample Word document. It discusses the importance of data visualization.")
doc.add_paragraph("Another paragraph talking about how LLMs can help in summarizing documents.")
doc.save("sample.docx")
print("sample.docx created.")

# Create sample CSV file
data_csv = {
    'ID': [1, 2, 3, 4, 5],
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
    'Age': [25, 30, 35, 40, 22],
    'City': ['New York', 'London', 'Paris', 'New York', 'London'],
    'Salary': [70000, 80000, 90000, 75000, 60000]
}
df_sample_csv = pd.DataFrame(data_csv)
df_sample_csv.to_csv("sample.csv", index=False)
print("sample.csv created.")

# Create sample PDF file (simple one-page PDF)
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    c = canvas.Canvas("sample.pdf", pagesize=letter)
    c.drawString(72, 800, "Sample PDF Document") # Coordinates from bottom-left
    c.drawString(72, 780, "This is a test PDF file created for the Data Analyst Agent.")
    c.drawString(72, 760, "It contains textual data that the agent should extract.")
    c.drawString(72, 740, "The agent will use PyPDF2 to read this content.")
    c.save()
    print("sample.pdf created using reportlab.")
except ImportError:
    print("Could not create sample.pdf: reportlab library not found. Please install it: pip install reportlab")
except Exception as e:
    print(f"Could not create sample.pdf: {e}.")


# Create a sample PNG image with text for OCR
try:
    img = Image.new('RGB', (450, 100), color = (255, 255, 255))
    d = ImageDraw.Draw(img)
    # Attempt to load a common font, fallback to default if not found
    try:
        font = ImageFont.truetype("arial.ttf", 20)
    except IOError:
        font = ImageFont.load_default()
    d.text((10,10), "Sample Text in Image for OCR\nData Agent Test 123", fill=(0,0,0), font=font)
    img.save("sample.png")
    print("sample.png created.")
except Exception as e:
    print(f"Could not create sample.png: {e}. Pillow or a system font might be missing.")

print("\nSample files created (or attempted). Ensure Tesseract OCR is installed for sample.png processing and reportlab for PDF.")
# --- Code Cell ---
# Cell for Instantiating the Agent

# Make sure the API key is set in the environment or directly in the config cell
# For example, if running locally and key is in environment:
# import os
# os.environ['TOGETHER_API_KEY'] = 'YOUR_ACTUAL_KEY_HERE' # Replace if needed

# Instantiate the agent
# This relies on TOGETHER_API_KEY and MODEL_NAME being set in the initial setup cells.
try:
    agent = DataAnalysisAgent()
    print("DataAnalysisAgent instantiated successfully.")
    # Quick check of API readiness (optional, actual calls will test it)
    # if TOGETHER_API_KEY and MODEL_NAME:
    #    test_ping = query_llm("Hello!", max_tokens=5)
    #    print(f"LLM Ping response: {test_ping}")
    # else:
    #    print("Skipping LLM ping as API key or model name not found.")
except ValueError as e:
    print(f"Error instantiating agent: {e}")
    print("Please ensure TOGETHER_API_KEY and MODEL_NAME are correctly set in the first code cells.")
except Exception as e:
    print(f"An unexpected error occurred during agent instantiation: {e}")

# --- Code Cell ---
# Demonstrate TXT file handling
if 'agent' in locals() and agent is not None:
    print("--- TXT File Handling ---")
    txt_result = agent.upload_and_prepare_data("sample.txt")
    print(txt_result)
    if agent.data_summary:
        print("\nData Summary (TXT):")
        print(agent.data_summary)
        print("\nAsking a question about TXT:")
        answer = agent.ask_question("What is this document about?")
        print(f"Answer: {answer}")
else:
    print("Agent not instantiated or failed to instantiate. Skipping TXT demo.")
# --- Code Cell ---
# Demonstrate DOCX file handling
if 'agent' in locals() and agent is not None:
    print("\n--- DOCX File Handling ---")
    docx_result = agent.upload_and_prepare_data("sample.docx")
    print(docx_result)
    if agent.data_summary:
        print("\nData Summary (DOCX):")
        print(agent.data_summary)
        print("\nAsking a question about DOCX:")
        answer = agent.ask_question("What are the key topics mentioned in this document?")
        print(f"Answer: {answer}")
else:
    print("Agent not instantiated or failed to instantiate. Skipping DOCX demo.")
# --- Code Cell ---
# Demonstrate PDF file handling
if 'agent' in locals() and agent is not None:
    print("\n--- PDF File Handling ---")
    if os.path.exists("sample.pdf"):
        pdf_result = agent.upload_and_prepare_data("sample.pdf")
        print(pdf_result)
        if agent.data_summary:
            print("\nData Summary (PDF):")
            print(agent.data_summary)
            print("\nAsking a question about PDF:")
            answer = agent.ask_question("What tool is used to read this PDF's content according to the text?")
            print(f"Answer: {answer}")
    else:
        print("sample.pdf not found. Skipping PDF demo. Ensure reportlab is installed and ran successfully in the file creation cell.")
else:
    print("Agent not instantiated or failed to instantiate. Skipping PDF demo.")
# --- Code Cell ---
# Demonstrate CSV file handling and Analysis/Visualization
if 'agent' in locals() and agent is not None:
    print("\n--- CSV File Handling ---")
    csv_result = agent.upload_and_prepare_data("sample.csv")
    print(csv_result)
    if agent.data_summary and agent.data_type == 'tabular':
        print("\nData Summary (CSV):")
        print(agent.data_summary)

        print("\nAsking a question about CSV (e.g., describe data):")
        answer = agent.ask_question("How many people are listed in New York?")
        print(f"Q: How many people are listed in New York? \nA: {answer}")

        print("\nPerforming Analysis on CSV (Average Salary):")
        analysis_output = agent.perform_analysis("What is the average salary? Calculate and show it.")
        print(f"Analysis Result: {analysis_output}")

        print("\nPerforming Analysis on CSV (People older than 30):")
        analysis_output_2 = agent.perform_analysis("Show me the names and ages of people older than 30.")
        print(f"Analysis Result: {analysis_output_2}")

        print("\nGenerating Visualization for CSV (Bar chart of salaries by city):")
        # Ensure the display import is available
        try:
            from IPython.display import Image as IPImage, display
        except ImportError:
            print("IPython.display not available. Cannot display images inline.")
            def display(x): pass # dummy display
            def IPImage(filename): return None # dummy IPImage

        viz_output = agent.generate_visualization("Create a bar chart showing the average salary for each city.")
        print(viz_output)
        if "plot.png" in viz_output and os.path.exists("plot.png"):
            display(IPImage(filename='plot.png'))
            print("Displayed plot.png above.")
        else:
            print("Plot image 'plot.png' not found or not generated.")

        print("\nGenerating Visualization for CSV (Histogram of Ages):")
        viz_output_2 = agent.generate_visualization("Create a histogram of ages and save it as plot.png") # Explicitly ask to save
        print(viz_output_2)
        if "plot.png" in viz_output_2 and os.path.exists("plot.png"):
            display(IPImage(filename='plot.png')) # Will overwrite or display the new plot
            print("Displayed plot.png above.")
        else:
            print("Plot image 'plot.png' not found or not generated for the second plot.")
else:
    print("Agent not instantiated or failed to instantiate. Skipping CSV demo.")
# --- Code Cell ---
# Demonstrate Image (OCR) file handling
if 'agent' in locals() and agent is not None:
    print("\n--- Image (OCR) File Handling ---")
    print("Note: This requires Tesseract OCR to be correctly installed and configured on your system.")
    if os.path.exists("sample.png"):
        image_result = agent.upload_and_prepare_data("sample.png")
        print(image_result)
        if agent.data_summary:
            print("\nData Summary (Image OCR):")
            print(agent.data_summary)
            if agent.data and not (isinstance(agent.data, str) and "Error performing OCR" in agent.data):
                print("\nAsking a question about Image Content (requires successful OCR):")
                answer = agent.ask_question("What is the text content of the image?")
                print(f"Answer: {answer}")
            elif agent.data and (isinstance(agent.data, str) and "Error performing OCR" in agent.data):
                 print("\nSkipping question about image content due to OCR error.")
                 print(f"OCR Error details: {agent.data}")
            else:
                print("\nOCR might have produced empty text or an unexpected result.")
    else:
        print("sample.png not found. Skipping Image OCR demo.")
else:
    print("Agent not instantiated or failed to instantiate. Skipping Image OCR demo.")


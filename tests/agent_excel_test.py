import json
import os
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

Templets_dir = Path(__file__).resolve().parent.parent / "templates" / "tests"


def inspect_template(template_name: str) -> dict:
    file_path = Templets_dir / template_name
    wb = openpyxl.load_workbook(file_path, data_only=False)
    slots = list(wb.defined_names.keys())
    return {"template": template_name, "editable_slots": slots}


def fill_templet(template_name: str, data: dict):
    templet_path = Templets_dir / template_name
    output_path = Path(__file__).resolve().parent / f"output_{template_name}"
    wb = openpyxl.load_workbook(templet_path, data_only=False)
    updated = []
    for name, value in data.items():
        if name in wb.defined_names:
            defined_name = wb.defined_names[name]
            for sheet_name, cell_ref in defined_name.destinations:
                cell = wb[sheet_name][cell_ref.replace("$", "")]
                if str(cell.value).startswith("="):
                    print(f"Skipping formula in {name}")
                    continue
                cell.value = value
                updated.append(name)
    wb.save(output_path)
    return {"status": "success", "file": str(output_path), "updated_slots": updated}


AVAILABLE_TOOLS = {"inspect_template": inspect_template, "fill_template": fill_templet}


tools = [
    {
        "type": "function",
        "function": {
            "name": "inspect_template",
            "description": "Inspects an Excel template to discover available editable named slots",
            "parameters": {
                "type": "object",
                "properties": {
                    "template_name": {
                        "type": "string",
                        "description": "e.g. startup_burn_report.xlsx",
                    }
                },
                "required": ["template_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "fill_template",
            "description": "Fills approved slots in an Excel template using a key-value dictionary",
            "parameters": {
                "type": "object",
                "properties": {
                    "template_name": {
                        "type": "string",
                        "description": "The template filename",
                    },
                    "data": {
                        "type": "object",
                        "description": "Dictionary mapping slot names to their values",
                    },
                },
                "required": ["template_name", "data"],
            },
        },
    },
]


client = OpenAI(
    api_key=os.getenv("GEMINI_API_KEY"),
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
)
messages = [
    {
        "role": "user",
        "content": "Please inspect startup_burn_report.xlsx and fill it for 'Nexus AI Labs' with $4.5 million starting cash in Q4 2026.",
    }
]

while True:
    response = client.chat.completions.create(
        model="gemini-3.5-flash", messages=messages, tools=tools
    )
    message = response.choices[0].message
    messages.append(message)

    if not message.tool_calls:
        print("\nFINAL RESPONSE:")
        print(message.content)
        break

    for tool_call in message.tool_calls:
        func_name = tool_call.function.name
        func_args = json.loads(tool_call.function.arguments)
        print(f"\nCalling tool: {func_name} with args: {func_args}")

        fun_to_run = AVAILABLE_TOOLS.get(func_name)
        if fun_to_run:
            result = fun_to_run(**func_args)

            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "name": func_name,
                    "content": str(result),
                }
            )

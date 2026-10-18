import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

import json


def add_numbers(a: int, b: int) -> int:
    return a + b

tools=[{
    "type":"function",
    "function":{
        "name":"add_numbers",
        "description":"Add two numbers",
        "parameters":{
            "type":"object",
            "properties":{
                "a":{"type":"integer","description":"First number"},
                "b":{"type":"integer","description":"Second number"}
            },
            "required":["a","b"]
        }
    }
}]



client = OpenAI(
    api_key=os.getenv("GEMINI_API_KEY"),
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/"
)

response = client.chat.completions.create(
    model="gemini-3.5-flash",
    # model="gemma-4-31b-it",
    messages=[
        {"role": "user", "content": "What is 2 + 2? "}
    ],
    tools=tools
)

tool_call = response.choices[0].message.tool_calls[0]
args = json.loads(tool_call.function.arguments)
result = add_numbers(args["a"], args["b"])
print(f"Calculated result locally: {result}")

messages = [
        {"role": "user", "content": "What is 2 + 2? "},
        response.choices[0].message,
        {
            "role": "tool",
            "tool_call_id": tool_call.id,
            "content": str(result)
        }
    ]

final_response = client.chat.completions.create(
    model="gemini-3.5-flash",
    messages=messages
)
print("Agent Final Answer:", final_response.choices[0].message.content)








import json
import os
from abc import ABC

from dotenv import load_dotenv
from openai import OpenAI

from agentoffice.tools.base_tools import BaseTool

load_dotenv()


class BaseAgent(ABC):
    def __init__(
        self,
        model: str = "gemini-3.5-flash",
        system_prompt: str = "",
        tools: list[BaseTool] = None,
    ):
        self.client = OpenAI(
            api_key=os.getenv("GEMINI_API_KEY"),
            base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        )
        self.model = model
        self.system_prompt = system_prompt
        self.tool_map: dict[str, BaseTool] = {}
        if tools:
            for tool in tools:
                self.register_tool(tool)

    def register_tool(self, tool: BaseTool):
        self.tool_map[tool.name] = tool

    def execute_tool(self, name: str, argunments_json: str) -> str:
        tool = self.tool_map.get(name)
        if not tool:
            return f"Tool {name} not found"
        try:
            kwargs = json.loads(argunments_json)
        except Exception as e:
            return f"Error: Invalid JSON arguments: {e}"
        try:
            result = tool.execute(**kwargs)
            return str(result)
        except Exception as e:
            return f"Error executing {name}: {type(e).__name__} - {e}"

    def run(self, message: str, max_iterations: int = 10) -> str:
        messages = []
        if self.system_prompt:
            messages.append({"role": "system", "content": self.system_prompt})
        messages.append({"role": "user", "content": message})
        tools_schema = [tool.to_schema() for tool in self.tool_map.values()] or None

        for turn in range(max_iterations):
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=tools_schema,
            )

            response_message = response.choices[0].message
            messages.append(response_message)

            if not response_message.tool_calls:
                return response_message.content

            for tool_call in response_message.tool_calls:
                result = self.execute_tool(
                    tool_call.function.name, tool_call.function.arguments
                )
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "name": tool_call.function.name,
                        "content": result,
                    }
                )
        return "Max iterations reached"

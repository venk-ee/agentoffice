import json
import os
from abc import ABC

from dotenv import load_dotenv
from openai import OpenAI

from agentoffice.core.memory import AgentMemory
from agentoffice.core.tool import BaseTool

load_dotenv()


class BaseAgent(ABC):
    def __init__(
        self,
        model: str = "gemini-3.5-flash-lite",
        system_prompt: str = "",
        tools: list[BaseTool] | None = None,
        history: list[dict] | None = None,
        on_tool_call=None,
        memory=None,
    ):
        self.client = OpenAI(
            api_key=os.getenv("GEMINI_API_KEY"),
            base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        )
        self.model = model
        self.system_prompt = system_prompt
        self.tool_map: dict[str, BaseTool] = {}
        self.on_tool_call = on_tool_call
        self.memory = memory or AgentMemory(system_prompt=system_prompt)

        if tools:
            for tool in tools:
                self.register_tool(tool)

    def register_tool(self, tool: BaseTool):
        self.tool_map[tool.name] = tool

    def execute_tool(self, name: str, arguments: str) -> str:
        tool = self.tool_map.get(name)
        if not tool:
            return json.dumps({"error": f"Tool '{name}' not found"})

        try:
            kwargs = json.loads(arguments) if arguments else {}
            result = tool.execute(**kwargs)
            return json.dumps(result, default=str)
        except Exception as e:
            return json.dumps({"error": f"{type(e).__name__}: {e}"})

    def run(self, message: str, max_iterations: int = 10) -> str:
        self.memory.add_user_message(message)
        tools_schema = [tool.to_schema() for tool in self.tool_map.values()] or None

        for turn in range(max_iterations):
            response = self.client.chat.completions.create(
                model=self.model,
                tools=tools_schema,
                messages=self.memory.history,
            )

            response_message = response.choices[0].message
            if response.usage:
                self.memory.record_usage(response.usage)
            self.memory.add_assistant_message(response_message)

            if not response_message.tool_calls:
                return response_message.content

            for tool_call in response_message.tool_calls:
                result = self.execute_tool(
                    tool_call.function.name, tool_call.function.arguments
                )
                if self.on_tool_call:
                    print(f"Tool called: {tool_call.function.name}")
                    self.on_tool_call(
                        tool_call.function.name, tool_call.function.arguments, result
                    )
                self.memory.add_tool_message(
                    tool_call.id, tool_call.function.name, result
                )
        return "Max tool calls reached"

    def clear_history(self):
        self.memory.clear_history()
        return self.memory.history

    @property
    def history(self) -> list[dict]:
        return self.memory.history

    @property
    def session_usage(self) -> dict:
        return self.memory.session_usage

    @property
    def last_usage(self) -> dict:
        return self.memory.last_usage

    def compact(self, force: bool = True):
        return self.memory.compact(self.client, self.model, force=force)

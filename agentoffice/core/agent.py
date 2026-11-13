import json
import os
from abc import ABC

from dotenv import load_dotenv
from openai import OpenAI

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
    ):
        self.client = OpenAI(
            api_key=os.getenv("GEMINI_API_KEY"),
            base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        )
        self.model = model
        self.system_prompt = system_prompt
        self.tool_map: dict[str, BaseTool] = {}
        self.history = history or []
        self.on_tool_call = on_tool_call
        self.last_usage: dict[str, int] = {
        "prompt_tokens": 0,
        "completion_tokens": 0,
        "cached_tokens": 0,
        "total_tokens": 0,
            }
        self.session_usage: dict[str, int] = {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "cached_tokens": 0,
            "total_tokens": 0,
            "compactions": 0,
        }

        if self.system_prompt and not self.history:
            self.history.append({"role": "system", "content": self.system_prompt})
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
        user_message = {"role": "user", "content": message}
        self.history.append(user_message)
        tools_schema = [tool.to_schema() for tool in self.tool_map.values()] or None

        for turn in range(max_iterations):
            response = self.client.chat.completions.create(
                model=self.model,
                tools=tools_schema,
                messages=self.history,
            )

            response_message = response.choices[0].message
            if response.usage:
                details = getattr(response.usage, "prompt_tokens_details", None)
                cached = (getattr(details, "cached_tokens", 0) or 0) if details else 0
                p_toks = response.usage.prompt_tokens or 0
                c_toks = response.usage.completion_tokens or 0
                t_toks = response.usage.total_tokens or 0

                self.last_usage = {
                    "prompt_tokens": p_toks,
                    "completion_tokens": c_toks,
                    "cached_tokens": cached,
                    "total_tokens": t_toks,
                }
                self.session_usage["prompt_tokens"] += p_toks
                self.session_usage["completion_tokens"] += c_toks
                self.session_usage["cached_tokens"] += cached
                self.session_usage["total_tokens"] += t_toks
            self.history.append(response_message)

            if not response_message.tool_calls:
                return response_message.content

            for tool_call in response_message.tool_calls:
                result = self.execute_tool(
                    tool_call.function.name, tool_call.function.arguments
                )
                if self.on_tool_call:
                    print(f"Tool called: {tool_call.function.name}")
                    self.on_tool_call(tool_call.function.name, tool_call.function.arguments, result)
                self.history.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "name": tool_call.function.name,
                        "content": result,
                    }
                )
        return "Max iterations reached"
    
    def clear_history(self):
        self.history = []
        if self.system_prompt:
            self.history.append({"role": "system", "content": self.system_prompt})
        return self.history

    def compact(self):
        pass

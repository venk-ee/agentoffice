class AgentMemory:
    def __init__(
        self,
        system_prompt: str = "",
        token_threshold: int = 12000,
        max_tool_chars: int = 40000,
    ):
        self.system_prompt = system_prompt
        self.token_threshold = token_threshold
        self.max_tool_chars = max_tool_chars
        self.history: list[dict] = []

        if self.system_prompt:
            self.history.append({"role": "system", "content": self.system_prompt})

        self.last_usage = {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "cached_tokens": 0,
            "total_tokens": 0,
        }
        self.session_usage = {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "cached_tokens": 0,
            "total_tokens": 0,
            "compactions": 0,
        }

    def add_user_message(self, content: str) -> None:
        self.history.append({"role": "user", "content": content})

    def add_tool_message(self, tool_call_id: str, name: str, content: str) -> None:
        if len(content) > self.max_tool_chars:
            content = content[: self.max_tool_chars] + "\n... [Truncated: output too large]"

        self.history.append(
            {
                "role": "tool",
                "tool_call_id": tool_call_id,
                "name": name,
                "content": content,
            }
        )

    def add_assistant_message(self, message) -> None:
        if isinstance(message, dict):
            self.history.append(message)
            return

        msg_dict: dict = {
            "role": "assistant",
            "content": getattr(message, "content", None) or "",
        }

        tool_calls = getattr(message, "tool_calls", None)
        if tool_calls:
            msg_dict["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": "function",
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments,
                    },
                }
                for tc in tool_calls
            ]

        self.history.append(msg_dict)

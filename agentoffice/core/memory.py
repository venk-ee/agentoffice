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
            "tokens_saved": 0,
        }

    def add_user_message(self, content: str) -> None:
        self.history.append({"role": "user", "content": content})

    def add_tool_message(self, tool_call_id: str, name: str, content: str) -> None:
        if len(content) > self.max_tool_chars:
            content = (
                content[: self.max_tool_chars] + "\n... [Truncated: output too large]"
            )

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
            msg_dict["tool_calls"] = []
            for tc in tool_calls:
                tc_data = {
                    "id": tc.id,
                    "type": "function",
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments,
                    },
                }
                extra_content = getattr(tc, "extra_content", None)
                if extra_content:
                    tc_data["extra_content"] = extra_content
                msg_dict["tool_calls"].append(tc_data)

        self.history.append(msg_dict)

    def record_usage(self, usage) -> None:
        if not usage:
            return

        cached = (
            getattr(usage.prompt_tokens_details, "cached_tokens", 0)
            if hasattr(usage, "prompt_tokens_details") and usage.prompt_tokens_details
            else 0
        ) or 0

        self.last_usage = {
            "prompt_tokens": usage.prompt_tokens,
            "completion_tokens": usage.completion_tokens,
            "cached_tokens": cached,
            "total_tokens": usage.total_tokens,
        }
        self.session_usage["count"] = self.session_usage.get("count", 0) + 1
        for key in self.last_usage:
            self.session_usage[key] = (
                self.session_usage.get(key, 0) + self.last_usage[key]
            )

    def clear_history(self):
        self.history = []
        if self.system_prompt:
            self.history.append({"role": "system", "content": self.system_prompt})
        return self.history

    def get_history(self):
        return self.history

    def get_last_usage(self):
        return self.last_usage

    def get_session_usage(self):
        return self.session_usage

    def compact(self, client, model: str, force: bool = True):
        if not force and self.last_usage.get("prompt_tokens", 0) < self.token_threshold:
            return None
        slice_index = self._find_safe_slice_index()
        if slice_index is None:
            return None
        start_index = (
            1 if (self.history and self.history[0].get("role") == "system") else 0
        )
        old_messages = self.history[start_index:slice_index]

        if not old_messages:
            return None

        text_for_compaction = []
        for msg in old_messages:
            role = msg.get("role", "")
            content = msg.get("content", "")
            if content:
                text_for_compaction.append(f"{role.upper()}: {content}")

            # Include tool calls even if content is empty/None
            if msg.get("tool_calls"):
                tool_info = []
                for tc in msg["tool_calls"]:
                    tool_name = tc.get("function", {}).get("name", "")
                    tool_args = str(tc.get("function", {}).get("arguments", ""))
                    tool_info.append(f"{tool_name}({tool_args})")
                if tool_info:
                    text_for_compaction.append(
                        f"{role.upper()}: called {'; '.join(tool_info)}"
                    )

        transcript = "\n\n".join(text_for_compaction)

        summary_messages = [
            {
                "role": "system",
                "content": (
                    "You are a helpful assistant. Summarize the following conversation history concisely. "
                    "Crucially preserve: any file names, paths, identifiers, technical decisions, key "
                    "results, and the current task status."
                ),
            },
            {"role": "user", "content": f"Conversation:\n{transcript}"},
        ]

        tokens_before = self.last_usage.get("total_tokens", 0)

        response = client.chat.completions.create(
            model=model,
            messages=summary_messages,
        )
        if response.usage:
            # Record summary usage into session totals only without corrupting last_usage
            self.session_usage["count"] = self.session_usage.get("count", 0) + 1
            self.session_usage["prompt_tokens"] += (
                getattr(response.usage, "prompt_tokens", 0) or 0
            )
            self.session_usage["completion_tokens"] += (
                getattr(response.usage, "completion_tokens", 0) or 0
            )
            cached = (
                getattr(response.usage.prompt_tokens_details, "cached_tokens", 0)
                if hasattr(response.usage, "prompt_tokens_details")
                and response.usage.prompt_tokens_details
                else 0
            ) or 0
            self.session_usage["cached_tokens"] += cached
            self.session_usage["total_tokens"] += (
                getattr(response.usage, "total_tokens", 0) or 0
            )

        summary = (response.choices[0].message.content or "").strip()
        if not summary:
            return None

        new_history = (
            self.history[:1]
            if self.history and self.history[0].get("role") == "system"
            else []
        )
        new_history.append(
            {
                "role": "assistant",
                "content": f"[Summary of previous conversation]\n{summary}",
            }
        )
        new_history.extend(self.history[slice_index:])
        self.history = new_history
        self.session_usage["compactions"] += 1
        self.session_usage["tokens_saved"] += tokens_before
        return summary

    def _find_safe_slice_index(self, keep_ratio: float = 0.25) -> int:

        if len(self.history) <= 3:
            return None

        target = int(len(self.history) * (1 - keep_ratio))

        for i in range(target, 0, -1):
            if self.history[i]["role"] == "user":
                return i
        return None

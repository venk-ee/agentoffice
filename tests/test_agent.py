import json
import unittest
from unittest.mock import MagicMock

from agentoffice.core.agent import BaseAgent
from agentoffice.core.tool import BaseTool


class DummyTool(BaseTool):
    def __init__(self):
        super().__init__(
            name="dummy_tool",
            description="A dummy test tool.",
            parameters={
                "type": "object",
                "properties": {"query": {"type": "string"}},
                "required": ["query"],
            },
        )

    def execute(self, query: str = ""):
        if query == "fail":
            raise ValueError("Intentional tool failure")
        return {"result": f"processed: {query}"}


class TestBaseAgent(unittest.TestCase):
    def setUp(self):
        self.mock_client = MagicMock()
        self.tool = DummyTool()
        self.agent = BaseAgent(
            model="test-model",
            system_prompt="Test system prompt",
            tools=[self.tool],
            client=self.mock_client,
        )

    def test_agent_initialization(self):
        self.assertEqual(self.agent.model, "test-model")
        self.assertEqual(self.agent.system_prompt, "Test system prompt")
        self.assertIn("dummy_tool", self.agent.tool_map)
        self.assertEqual(len(self.agent.history), 1)
        self.assertEqual(self.agent.history[0]["role"], "system")

    def test_agent_initialization_with_history(self):
        initial_history = [
            {"role": "user", "content": "hello"},
            {"role": "assistant", "content": "hi"},
        ]
        agent = BaseAgent(
            model="test-model",
            history=initial_history,
            client=self.mock_client,
        )
        self.assertEqual(len(agent.history), 2)
        self.assertEqual(agent.history[0]["content"], "hello")

    def test_execute_tool_success(self):
        result_str = self.agent.execute_tool("dummy_tool", json.dumps({"query": "abc"}))
        result = json.loads(result_str)
        self.assertEqual(result, {"result": "processed: abc"})

    def test_execute_tool_not_found(self):
        result_str = self.agent.execute_tool("unknown_tool", "{}")
        result = json.loads(result_str)
        self.assertIn("error", result)
        self.assertIn("not found", result["error"])

    def test_execute_tool_handles_exception(self):
        result_str = self.agent.execute_tool(
            "dummy_tool", json.dumps({"query": "fail"})
        )
        result = json.loads(result_str)
        self.assertIn("error", result)
        self.assertIn("Intentional tool failure", result["error"])

    def test_run_single_turn_reply(self):
        mock_msg = MagicMock()
        mock_msg.content = "Here is the answer."
        mock_msg.tool_calls = None
        mock_choice = MagicMock()
        mock_choice.message = mock_msg

        mock_resp = MagicMock()
        mock_resp.choices = [mock_choice]
        mock_resp.usage = None
        self.mock_client.chat.completions.create.return_value = mock_resp

        reply = self.agent.run("What is 2+2?")
        self.assertEqual(reply, "Here is the answer.")
        self.assertEqual(self.agent.history[-1]["role"], "assistant")
        self.assertEqual(self.agent.history[-1]["content"], "Here is the answer.")

    def test_run_tool_call_loop(self):
        # First turn: model requests dummy_tool
        tc = MagicMock()
        tc.id = "call_123"
        tc.function.name = "dummy_tool"
        tc.function.arguments = json.dumps({"query": "data"})
        mock_msg_1 = MagicMock()
        mock_msg_1.content = None
        mock_msg_1.tool_calls = [tc]
        mock_choice_1 = MagicMock(message=mock_msg_1)
        mock_resp_1 = MagicMock(choices=[mock_choice_1], usage=None)

        # Second turn: model replies with final answer
        mock_msg_2 = MagicMock()
        mock_msg_2.content = "Done with data."
        mock_msg_2.tool_calls = None
        mock_choice_2 = MagicMock(message=mock_msg_2)
        mock_resp_2 = MagicMock(choices=[mock_choice_2], usage=None)

        self.mock_client.chat.completions.create.side_effect = [
            mock_resp_1,
            mock_resp_2,
        ]

        callback_calls = []

        def callback(name, args, res):
            callback_calls.append((name, args, res))

        self.agent.on_tool_call = callback
        reply = self.agent.run("Process data")
        self.assertEqual(reply, "Done with data.")
        self.assertEqual(len(callback_calls), 1)
        self.assertEqual(callback_calls[0][0], "dummy_tool")

    def test_run_callback_exception_does_not_break_execution(self):
        tc = MagicMock()
        tc.id = "call_123"
        tc.function.name = "dummy_tool"
        tc.function.arguments = json.dumps({"query": "data"})
        mock_msg_1 = MagicMock(content=None, tool_calls=[tc])
        mock_resp_1 = MagicMock(choices=[MagicMock(message=mock_msg_1)], usage=None)

        mock_msg_2 = MagicMock(content="Finished smoothly.", tool_calls=None)
        mock_resp_2 = MagicMock(choices=[MagicMock(message=mock_msg_2)], usage=None)

        self.mock_client.chat.completions.create.side_effect = [
            mock_resp_1,
            mock_resp_2,
        ]

        def crashing_callback(name, args, res):
            raise BrokenPipeError("Stdout broken")

        self.agent.on_tool_call = crashing_callback
        reply = self.agent.run("Process data safely")
        self.assertEqual(reply, "Finished smoothly.")

import runpy
from types import SimpleNamespace
from unittest.mock import MagicMock

from agentoffice.core.agent import BaseAgent
from agentoffice.core.tool import BaseTool
from agentoffice.interface import chat


class _DummyTool(BaseTool):
    def __init__(self):
        super().__init__(
            name="dummy_tool",
            description="Dummy tool for testing.",
            parameters={"type": "object", "properties": {}, "required": []},
        )

    def execute(self, **kwargs):
        return {"ok": True}


def test_main_starts_chat_interface_and_exits(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "workspace").mkdir()
    monkeypatch.setattr("builtins.input", lambda _: "quit")

    fake_agent = MagicMock()
    created_agents = []

    def fake_base_agent(**kwargs):
        created_agents.append(kwargs)
        return fake_agent

    monkeypatch.setattr(chat, "BaseAgent", fake_base_agent)
    chat.main()

    assert len(created_agents) == 1
    assert created_agents[0]["on_tool_call"] is chat.print_tool_call


def test_python_m_agentoffice_calls_main(monkeypatch):
    called = []
    monkeypatch.setattr(chat, "main", lambda: called.append(True))

    runpy.run_module("agentoffice", run_name="__main__")

    assert called == [True]


def test_base_agent_on_tool_call_does_not_print_duplicate_log(capsys):
    tool_call = SimpleNamespace(
        id="tc_1",
        function=SimpleNamespace(name="dummy_tool", arguments="{}"),
    )
    first_response = SimpleNamespace(
        choices=[
            SimpleNamespace(message=SimpleNamespace(content="", tool_calls=[tool_call]))
        ],
        usage=None,
    )
    second_response = SimpleNamespace(
        choices=[
            SimpleNamespace(message=SimpleNamespace(content="Done", tool_calls=None))
        ],
        usage=None,
    )

    mock_client = MagicMock()
    mock_client.chat.completions.create.side_effect = [first_response, second_response]

    recorded = []
    agent = BaseAgent(
        tools=[_DummyTool()],
        client=mock_client,
        on_tool_call=lambda name, args, res: recorded.append((name, args, res)),
    )

    reply = agent.run("run dummy")

    captured = capsys.readouterr()
    assert reply == "Done"
    assert len(recorded) == 1
    assert "Tool called:" not in captured.out

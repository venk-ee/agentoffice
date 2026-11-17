from pathlib import Path

from agentoffice.core.agent import BaseAgent
from agentoffice.tools.excel import create_excel_tools


def print_tool_call(name: str, arguments: str, result: str) -> None:
    # print(f"\n⚙️  [Tool] {name}({arguments})")
    print(f"\n⚙️  [Tool] {name}({arguments}) -> {result[:120]}")


class ChatInterface:
    def __init__(self, agent: BaseAgent):
        self.agent = agent

    def run(self):
        while True:
            user_input = input("You: ")
            reply = self.agent.run(user_input)
            print(f"Agent: {reply}")



if __name__ == "__main__":
    workspace = Path("./workspace").resolve()
    tools = create_excel_tools(workspace)
    agent = BaseAgent(
        model="gemini-3.5-flash-lite",
        tools=tools,
        system_prompt="You are an office assistant. You can inspect, read, and update Excel workbooks in the workspace.",
        on_tool_call=print_tool_call,
    )
    interface = ChatInterface(agent=agent)
    interface.run()
from pathlib import Path

from agentoffice.core.agent import BaseAgent
from agentoffice.tools.excel import create_excel_tools


def print_tool_call(name: str, arguments: str, result: str) -> None:
    print(f"\n[Tool] {name}({arguments}) -> {result[:120]}")


class ChatInterface:
    def __init__(self, agent: BaseAgent):
        self.agent = agent

    def run(self):
        while True:
            try:
                user_input = input("You: ")
            except (KeyboardInterrupt, EOFError):
                print("\nExiting chat session.")
                break
            if user_input.strip().lower() in {"exit", "quit", "q"}:
                print("Exiting chat session.")
                break
            if not user_input.strip():
                continue
            try:
                reply = self.agent.run(user_input)
                print(f"Agent: {reply}")
            except Exception as e:
                print(
                    f"Error during request: {e}. "
                    "Tool calls completed earlier in this request may have already changed the workbook. "
                    "Please check the workbook before repeating a write."
                )


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

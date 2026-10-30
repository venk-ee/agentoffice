from dotenv import load_dotenv

load_dotenv()

from agentoffice.agents.agent import BaseAgent
from agentoffice.tools.office.excel import InspectWorkbookTool, ReadRangeTool

# 1. Initialize agent with both Excel tools
agent = BaseAgent(
    model="gemini-3.8-flash",
    system_prompt=(
        "You are an expert Office Assistant. "
        "Use your Excel tools to inspect files and answer questions precisely."
    ),
    tools=[InspectWorkbookTool(), ReadRangeTool()],
)

prompt = (
    "Please inspect the workbook 'testing_templates/tests/startup_burn_report.xlsx'. "
    "Tell me: what company is this for, what is the starting cash, "
    "and what was the total expenses in September 2026?"
)

print("--- Running Agent ---")
response = agent.run(prompt)
print("\n--- Agent Answer ---")
print(response)

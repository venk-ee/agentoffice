from agentoffice import BaseTool


class DivisionTool(BaseTool):
    def __init__(self):
        super().__init__(
            name="divide_numbers",
            description="Divides number a by number b",
            parameters={
                "type": "object",
                "properties": {
                    "a": {"type": "number", "description": "Numerator"},
                    "b": {"type": "number", "description": "Denominator"},
                },
                "required": ["a", "b"],
            },
        )

    def execute(self, a: float, b: float) -> float:
        return a / b


if __name__ == "__main__":
    from agentoffice import BaseAgent

    agent = BaseAgent(
        model="gemini-3.5-flash",
        system_prompt="You are a helpful math assistant. Always use tools for arithmetic.",
        tools=[DivisionTool()],
    )

    print(agent.run("What is 100 divided by 4?"))
    print(agent.run("What is 100 divided by 0?"))

from abc import ABC, abstractmethod


class BaseTool(ABC):
    def __init__(self, name: str, description: str, parameters: dict):
        self.name = name
        self.description = description
        self.parameters = dict(parameters)
        if "additionalProperties" not in self.parameters:
            self.parameters["additionalProperties"] = False

    @abstractmethod
    def execute(self, **kwargs):
        pass

    def to_schema(self):
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": self.parameters,
            },
        }

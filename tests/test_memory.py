from unittest.mock import MagicMock

from agentoffice.core.memory import AgentMemory


def test_record_usage_with_cache():
    memory = AgentMemory()
    usage = MagicMock()
    usage.prompt_tokens = 5000
    usage.completion_tokens = 200
    usage.total_tokens = 5200
    usage.prompt_tokens_details.cached_tokens = 4096

    memory.record_usage(usage)
    assert memory.last_usage["cached_tokens"] == 4096
    assert memory.session_usage["cached_tokens"] == 4096
    assert memory.session_usage["total_tokens"] == 5200


def test_memory():
    memory = AgentMemory(system_prompt="You are an assistant")

    assert len(memory.history) == 1
    assert memory.history[0] == {"role": "system", "content": "You are an assistant"}
    assert memory.last_usage["total_tokens"] == 0
    assert memory.session_usage["total_tokens"] == 0
    
def test_add_user_message():
    memory = AgentMemory(system_prompt="You are an assistant")
    memory.add_user_message("Hello")
    assert len(memory.history) == 2
    assert memory.history[1] == {"role": "user", "content": "Hello"}

def test_add_assistant_message():
    memory = AgentMemory(system_prompt="You are an assistant")
    memory.add_assistant_message({"role": "assistant", "content": "Hello"})
    assert len(memory.history) == 2
    assert memory.history[1] == {"role": "assistant", "content": "Hello"}

def test_add_tool_message():
    memory = AgentMemory(system_prompt="You are an assistant")
    memory.add_tool_message("tool_id", "tool_name", "tool_result")
    assert len(memory.history) == 2
    assert memory.history[1] == {"role": "tool", "tool_call_id": "tool_id", "name": "tool_name", "content": "tool_result"}

def test_clear_history():
    memory = AgentMemory(system_prompt="You are an assistant")
    memory.add_user_message("Hello")
    memory.add_assistant_message({"role": "assistant", "content": "Hello"})
    memory.add_tool_message("tool_id", "tool_name", "tool_result")
    memory.clear_history()
    assert len(memory.history) == 1
    assert memory.history[0] == {"role": "system", "content": "You are an assistant"}

def test_memory_initialization_empty():
    memory = AgentMemory()
    assert len(memory.history) == 0
    assert memory.last_usage["total_tokens"] == 0
    assert memory.session_usage["total_tokens"] == 0

def test_find_safe_slice_index_too_short():
    memory = AgentMemory()
    memory.add_user_message("Hello")
    assert memory._find_safe_slice_index() is None


def test_find_safe_slice_index_finds_user_boundary():
    memory = AgentMemory(system_prompt="System")  # index 0
    memory.add_user_message("Turn 1")  # index 1
    memory.add_assistant_message({"role": "assistant", "content": "R1"})  # index 2
    memory.add_tool_message("t1", "tool", "out1")  # index 3
    memory.add_user_message("Turn 2")  # index 4
    memory.add_assistant_message({"role": "assistant", "content": "R2"})  # index 5
    memory.add_tool_message("t2", "tool", "out2")  # index 6

    # Total length is 7. If we keep the last ~30% (keep_ratio=0.3):
    # target = int(7 * 0.7) = 4
    # Index 4 is a "user" message!
    slice_idx = memory._find_safe_slice_index(keep_ratio=0.3)
    assert slice_idx is not None
    assert memory.history[slice_idx]["role"] == "user"


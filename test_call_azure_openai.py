"""
Unit tests for call_azure_openai and the full-auto POE document generation flow.
Verifies robust handling of:
- Normal ChatCompletion responses
- Unexpected string responses (the 'str' object has no attribute 'choices' bug)
- Empty choices
- Empty content
- API version-based parameter selection (max_tokens vs max_completion_tokens)
"""

import sys
import os
import unittest
from unittest.mock import MagicMock, patch, PropertyMock

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

# Mock heavy dependencies before importing app
mock_st = MagicMock()
mock_st.secrets = {
    "AZURE_OPENAI_KEY": "test-key",
    "AZURE_OPENAI_ENDPOINT": "https://test.openai.azure.com/",
    "AZURE_OPENAI_DEPLOYMENT": "gpt-4o",
    "AZURE_OPENAI_API_VERSION": "2024-06-01",
}
mock_st.set_page_config = MagicMock()
mock_st.cache_data = lambda f=None, **kwargs: f if f else (lambda fn: fn)
mock_st.session_state = {}

sys.modules['streamlit'] = mock_st
sys.modules['msal'] = MagicMock()
sys.modules['docx'] = MagicMock()
sys.modules['docx.shared'] = MagicMock()
sys.modules['docx.enum'] = MagicMock()
sys.modules['docx.enum.text'] = MagicMock()
sys.modules['docx.oxml'] = MagicMock()
sys.modules['docx.oxml.ns'] = MagicMock()
sys.modules['requests'] = MagicMock()

mock_ui = MagicMock()
sys.modules['frontend'] = MagicMock()
sys.modules['frontend.ui'] = mock_ui

import importlib.util

spec = importlib.util.spec_from_file_location("app", os.path.join(APP_DIR, "app.py"))
app_module = importlib.util.module_from_spec(spec)

with patch.object(mock_st, 'set_page_config'):
    try:
        spec.loader.exec_module(app_module)
    except Exception as e:
        print(f"Note: Module load produced non-fatal error: {type(e).__name__}: {e}")

call_azure_openai = app_module.call_azure_openai
generate_solution_artifact = app_module.generate_solution_artifact
generate_pov_artifact = app_module.generate_pov_artifact


class FakeMessage:
    def __init__(self, content):
        self.content = content


class FakeChoice:
    def __init__(self, content="Test response", finish_reason="stop"):
        self.message = FakeMessage(content)
        self.finish_reason = finish_reason


class FakeChatCompletion:
    def __init__(self, choices=None):
        self.choices = choices if choices is not None else [FakeChoice()]


class TestCallAzureOpenAI(unittest.TestCase):
    """Test call_azure_openai with various response scenarios."""

    @patch.object(app_module, 'get_openai_client')
    def test_normal_response(self, mock_get_client):
        """Normal ChatCompletion response should return content string."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content="Hello, world!")]
        )
        mock_get_client.return_value = mock_client

        result = call_azure_openai("system", "user")
        self.assertEqual(result, "Hello, world!")

    @patch.object(app_module, 'get_openai_client')
    def test_string_response_raises_runtime_error(self, mock_get_client):
        """If SDK returns a raw string, should raise RuntimeError instead of AttributeError."""
        mock_client = MagicMock()
        # Simulate the bug: SDK returns a raw string instead of ChatCompletion
        mock_client.chat.completions.create.return_value = "raw string response from Azure"
        mock_get_client.return_value = mock_client

        with self.assertRaises(RuntimeError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("非预期的原始字符串响应", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_empty_choices_raises_runtime_error(self, mock_get_client):
        """If response has empty choices list, should raise RuntimeError."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(choices=[])
        mock_get_client.return_value = mock_client

        with self.assertRaises(RuntimeError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("无效响应结构", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_none_choices_raises_runtime_error(self, mock_get_client):
        """If response.choices is None, should raise RuntimeError."""
        mock_client = MagicMock()
        completion = FakeChatCompletion()
        completion.choices = None
        mock_client.chat.completions.create.return_value = completion
        mock_get_client.return_value = mock_client

        with self.assertRaises(RuntimeError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("无效响应结构", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_empty_content_raises_value_error(self, mock_get_client):
        """If response content is empty, should raise ValueError."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content="")]
        )
        mock_get_client.return_value = mock_client

        with self.assertRaises(ValueError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("API 返回了空内容", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_none_content_raises_value_error(self, mock_get_client):
        """If response content is None, should raise ValueError."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content=None)]
        )
        mock_get_client.return_value = mock_client

        with self.assertRaises(ValueError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("API 返回了空内容", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_whitespace_content_raises_value_error(self, mock_get_client):
        """If response content is only whitespace, should raise ValueError."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content="   \n  ")]
        )
        mock_get_client.return_value = mock_client

        with self.assertRaises(ValueError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("API 返回了空内容", str(ctx.exception))

    @patch.object(app_module, 'get_openai_client')
    def test_no_choices_attribute_raises_runtime_error(self, mock_get_client):
        """If response object has no 'choices' attribute, should raise RuntimeError."""
        mock_client = MagicMock()
        # Return an object without 'choices'
        bad_response = MagicMock(spec=[])  # spec=[] means no attributes
        mock_client.chat.completions.create.return_value = bad_response
        mock_get_client.return_value = mock_client

        with self.assertRaises(RuntimeError) as ctx:
            call_azure_openai("system", "user")
        self.assertIn("无效响应结构", str(ctx.exception))


class TestApiVersionParameterSelection(unittest.TestCase):
    """Test that max_tokens is always used with the OpenAI-compatible client."""

    @patch.object(app_module, 'get_openai_client')
    def test_always_uses_max_tokens(self, mock_get_client):
        """OpenAI-compatible client should always use max_tokens."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content="test")]
        )
        mock_get_client.return_value = mock_client

        call_azure_openai("system", "user")

        call_kwargs = mock_client.chat.completions.create.call_args
        self.assertIn("max_tokens", call_kwargs.kwargs)
        self.assertNotIn("max_completion_tokens", call_kwargs.kwargs)
        self.assertEqual(call_kwargs.kwargs["max_tokens"], 16384)

    @patch.object(app_module, 'get_openai_client')
    def test_uses_correct_model(self, mock_get_client):
        """Should pass the deployment/model name from secrets."""
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = FakeChatCompletion(
            [FakeChoice(content="test")]
        )
        mock_get_client.return_value = mock_client

        call_azure_openai("system", "user")

        call_kwargs = mock_client.chat.completions.create.call_args
        self.assertEqual(call_kwargs.kwargs["model"], mock_st.secrets["AZURE_OPENAI_DEPLOYMENT"])
        self.assertEqual(call_kwargs.kwargs["model"], mock_st.secrets["AZURE_OPENAI_DEPLOYMENT"])


class TestGenerateSolutionArtifact(unittest.TestCase):
    """Test generate_solution_artifact handles errors from call_azure_openai."""

    @patch.object(app_module, 'create_solution_docx', return_value=b"fake docx")
    @patch.object(app_module, 'call_azure_openai')
    def test_normal_generation(self, mock_call, mock_docx):
        """Should return artifact dict with content, bytes, and file_name."""
        mock_call.return_value = "# Test Solution\n\nContent here"

        result = generate_solution_artifact(
            current_doc_type="AI",
            customer_name="TestCo",
            account_name="TestCo",
            customer_bg="Test background",
            solution_ref="",
            infra_ref="",
        )
        self.assertEqual(result["content"], "# Test Solution\n\nContent here")
        self.assertEqual(result["file_name"], "TestCo-Solution Architecture.docx")
        self.assertEqual(result["bytes"], b"fake docx")

    @patch.object(app_module, 'call_azure_openai')
    def test_str_response_error_propagates(self, mock_call):
        """If call_azure_openai raises RuntimeError due to str response, it propagates."""
        mock_call.side_effect = RuntimeError("Azure OpenAI 返回了非预期的原始字符串响应")

        with self.assertRaises(RuntimeError) as ctx:
            generate_solution_artifact(
                current_doc_type="AI",
                customer_name="TestCo",
                account_name="TestCo",
                customer_bg="Test background",
                solution_ref="",
                infra_ref="",
            )
        self.assertIn("非预期的原始字符串响应", str(ctx.exception))


class TestGeneratePovArtifact(unittest.TestCase):
    """Test generate_pov_artifact handles errors from call_azure_openai."""

    @patch.object(app_module, 'create_pov_docx', return_value=b"fake pov docx")
    @patch.object(app_module, 'call_azure_openai')
    def test_normal_generation(self, mock_call, mock_docx):
        """Should return artifact dict for POV."""
        mock_call.return_value = "# POV Plan\n\nDeployment plan here"
        import datetime

        result = generate_pov_artifact(
            solution_text="# Solution\n\nSolution text",
            customer_name="TestCo",
            account_name="TestCo",
            pov_ref="",
            pov_start=datetime.date(2025, 1, 1),
            pov_end=datetime.date(2025, 3, 31),
            vendor_team="PM: 张三\nArch: 李四",
        )
        self.assertEqual(result["content"], "# POV Plan\n\nDeployment plan here")
        self.assertEqual(result["file_name"], "TestCo-PostAssessment POVdeployment.docx")


if __name__ == "__main__":
    unittest.main()

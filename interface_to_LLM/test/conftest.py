import pytest
import logging

from openai.types.chat.chat_completion import ChatCompletion, Choice
from openai.types.chat.chat_completion_message import ChatCompletionMessage
from openai.types.completion_usage import CompletionUsage

from interface_to_LLM.interface_to_llm import BaseClient, OpenAIClient


class TestClient(BaseClient):
	__test__ = False

	def __init__(self):
		super().__init__()
		self.client_name = "TEST"

	def authenticate(self, **kwargs):
		self.client = True
		logging.info("\t Successful authentication of test client")

	def send_prompt(self,
			prompt: str, system_prompt: str,
			seed: int, max_tokens: int,	temperature: float = 0,
			**kwargs) -> dict:

		return self._parse_response(response=self._get_response(prompt=prompt))

	def _get_response(self, prompt: str):
		response = {
			"text": "Lorem ipsum.",
			"input_tokens": self.count_tokens(prompt=prompt),
			"output_tokens": 2,
			"origin_tokens": "client"
		}
		return response

	@staticmethod
	def _parse_response(response, **kwargs) -> dict:
		return response

	@staticmethod
	def count_tokens(prompt: str, **kwargs) -> int:
		return len(prompt.split(" "))


class TestOpenAIClient(OpenAIClient):
	__test__ = False

	def authenticate(self, **kwargs):
		self.client = True

	def _get_response(self, params):
		return ChatCompletion(
			id="x", created=0, object="chat.completion",
			model="gpt-3.5-turbo",
			choices=[Choice(
				index=0, finish_reason="stop", message=ChatCompletionMessage(content="Lorem ipsum.", role="assistant")
			)],
			usage=CompletionUsage(completion_tokens=2, prompt_tokens=1, total_tokens=3)
		)

from abc import ABC, abstractmethod
import logging

# OpenAIClient related imports
from openai import OpenAI
from openai.types.chat import ChatCompletion
from openai._exceptions import APIStatusError as OpenAIError
import tiktoken

# AnthropicClient related imports
import anthropic
from anthropic._exceptions import APIStatusError as AnthropicError

from interface_to_LLM.exceptions import (ClientNotAuthenticatedError, ClientAuthenticationError,
										 PromptEmptyError, PromptsEmptyError)


def client_authenticated(func):
	def wrapper(self, *args, **kwargs):
		if self.client.client is None:
			raise ClientNotAuthenticatedError("Client is not authenticated, call .authenticate().")
		return func(self, *args, **kwargs)
	return wrapper

def minimum_response_elements(func):
	def wrapper(*args, **kwargs):
		response = func(**kwargs)
		required_keys = ["text", "input_tokens", "output_tokens", "origin_tokens"]
		if not all(k in response.keys() for k in required_keys):
			raise KeyError(f"Parsed response requires: {required_keys}")
		return response
	return wrapper


class BaseClient(ABC):
	def __init__(self):
		self.client_name = None
		self.client = None

	@abstractmethod
	def authenticate(self, **kwargs):
		pass

	@abstractmethod
	def send_prompt(self,
			prompt: str, system_prompt: str,
			seed: int, max_tokens: int,	temperature: float = 0.2,
			**kwargs) -> dict:
		pass

	@abstractmethod
	def _get_response(self, **kwargs):
		pass

	@minimum_response_elements  # Required in the implementation of the child class
	@staticmethod
	def _parse_response(response, **kwargs) -> dict:
		pass

	@staticmethod
	def count_tokens(prompt: str, **kwargs) -> int:
		pass


class OpenAIClient(BaseClient):
	def __init__(self, default_model: str = "gpt-4.2"):
		super().__init__()
		self.client_name = "OpenAI"
		self.default_model = default_model

	def authenticate(self, api_key: str, **kwargs):
		"""
		Initializes and authenticates OpenAI client
		:param api_key: OpenAI api key
		:param kwargs: To handle unexpected parameters
		"""

		try:
			self.client = OpenAI(api_key=api_key)
			self.client.models.list()  # Lightweight api call to validate the api key
			logging.info("\t Successful authentication of OpenAI")
		except OpenAIError as e:
			self.client = None
			logging.error(f"\tError in authentication of OpenAI: {e}")
			raise ClientAuthenticationError

	def send_prompt(self,
			prompt: str, system_prompt: str,
			seed: int, max_tokens: int,	temperature: float = 0.2,
			model: str = None, log_probs_bool: bool = False, **kwargs) -> dict:
		"""
		Sends the given prompt to the model, obtains and parses the model response
		:param prompt: Prompt string to be sent to the model and obtain a response
		:param system_prompt: Prompt string to define the model responses behaviour
		:param seed: Optional parameter to control the 'randomness' of the model response for reproducibility
		:param max_tokens: Optional parameter to limit the number of tokens used
		:param temperature: Optional parameter to control the 'creativity' of the model response
		:param model: Optional parameter to specify the model to be used (if None uses the default model)
		:param log_probs_bool: Optional parameter on whether to obtain model response log probabilities
		:param kwargs: To handle unexpected parameters
		:return: Dictionary containing the response text as well as some additional information
		"""

		params = {
			"messages": [
				{"role": "system", "content": system_prompt},
				{"role": "user", "content": prompt}
			],
			"model": model if model is not None else self.default_model,
			"seed": seed,
			"logprobs": log_probs_bool
		}

		# Only add max_tokens for models that support it
		if "o3-" not in params["model"]:  # Exclude `o3-mini`
			params["max_tokens"] = max_tokens  # Use only for GPT-4 and older models
			params["temperature"] = temperature

		response = self._get_response(params)

		# Obtain parsed response dictionary of the relevant information
		parsed_response = self._parse_response(response=response)

		logging.info(
			f"\t[Interface to LLM] <get_response>: temperature = {temperature}, log probs bool = {log_probs_bool}"
			f"\n\t -> Prompt ({parsed_response['input_tokens']} tokens):\n\t{prompt}"
			f"\n\t <- Response ({parsed_response['output_tokens']} tokens):\n\t{parsed_response['text']}"
			f"\n\t -\tfinish reason = {parsed_response['finish_reason']}"
			f"\n\t -\tlog probabilities = {parsed_response['log_probs'] if log_probs_bool else 'None'}"
		)

		return parsed_response

	def _get_response(self, params) -> ChatCompletion:
		return self.client.chat.completions.create(**params)

	@minimum_response_elements
	@staticmethod
	def _parse_response(response: ChatCompletion, **kwargs) -> dict:
		"""
		Parses the chat completion response object into a dictionary with the relevant attributes
		:param response: Raw chat completion response to parse
		:return parsed_response: Dictionary containing the response text as well as some additional information
		"""

		# Check correct response structure, if not raise correspondent error
		if not (hasattr(response, "choices") and hasattr(response, "usage")):
			raise AttributeError

		# Choices data contains all the versions of the response, by default there's only 1 response
		parsed_response = {
			"text": response.choices[0].message.content,  # Model response text
			"log_probs": response.choices[0].logprobs,  # Model response text log probabilities
			"finish_reason": response.choices[0].finish_reason,  # Either 'stop' or 'length' (max_tokens limit reached)
			"input_tokens": response.usage.prompt_tokens,  # Number of tokens of the input prompt (+ system prompt)
			"output_tokens": response.usage.completion_tokens,  # Number of tokens of the outputted response
			"origin_tokens": "client"  # 'client' if the tokens come from the client or 'calculated' otherwise
		}

		return parsed_response

	@staticmethod
	def count_tokens(prompt: str, model: str = None, **kwargs) -> int:
		"""
		Counts the number of tokens for a given prompt and model
		:param prompt: Prompt string
		:param model: Name of the model used, each model tokenizes different
		:param kwargs: To handle unexpected parameters
		:return num_tokens: Number of tokens in the given prompt
		"""

		# If no model is specified use tiktoken base model encoding, else use the encoding for the given model
		if model is None:
			# Tiktoken base model supports most of the gpt models
			encoding = tiktoken.get_encoding("cl100k_base")
		else:
			encoding = tiktoken.encoding_for_model(model)

		return len(encoding.encode(prompt))


class AnthropicClient(BaseClient):
	def __init__(self, default_model: str = "claude-sonnet-4-6"):
		super().__init__()
		self.client_name = "Anthropic"
		self.default_model = default_model

	def authenticate(self, api_key: str, **kwargs):
		"""
		Initializes and authenticates Anthropic client.
		:param api_key: Anthropic API key
		:param kwargs: To handle unexpected parameters
		"""
		try:
			self.client = anthropic.Anthropic(api_key=api_key)
			self.client.models.list()  # Lightweight API call to validate the key
			logging.info("\t Successful authentication of Anthropic")
		except AnthropicError as e:
			self.client = None
			logging.error(f"\tError in authentication of Anthropic: {e}")
			raise ClientAuthenticationError

	def send_prompt(self,
			prompt: str, system_prompt: str,
			seed: int, max_tokens: int, temperature: float = 0.2,
			model: str = None, **kwargs) -> dict:
		"""
		Sends the given prompt to the model, obtains and parses the model response.
		:param prompt: Prompt string to be sent to the model
		:param system_prompt: System prompt string passed via the messages API system parameter
		:param seed: Not used by Anthropic API (kept for interface compatibility)
		:param max_tokens: Maximum number of tokens in the response (required by Anthropic)
		:param temperature: Controls response randomness (0.0–1.0)
		:param model: Optional model override; defaults to self.default_model
		:param kwargs: To handle unexpected parameters
		:return: Dictionary containing the response text and token usage
		"""
		params = {
			"model": model if model is not None else self.default_model,
			"max_tokens": max_tokens if max_tokens is not None else 4096,
			"temperature": temperature,
			# System prompt uses cache_control so the style guide is cached after the first call,
			# saving ~90% on repeated input tokens across all batches in a document run.
			"system": [
				{
					"type": "text",
					"text": system_prompt,
					"cache_control": {"type": "ephemeral"}
				}
			],
			"messages": [{"role": "user", "content": prompt}]
		}

		response = self._get_response(params)
		parsed_response = self._parse_response(response=response)

		logging.info(
			f"\t[Interface to LLM] <get_response>: temperature = {temperature}"
			f"\n\t -> Prompt ({parsed_response['input_tokens']} tokens):\n\t{prompt}"
			f"\n\t <- Response ({parsed_response['output_tokens']} tokens):\n\t{parsed_response['text']}"
			f"\n\t -\tfinish reason = {parsed_response['finish_reason']}"
		)

		return parsed_response

	def _get_response(self, params) -> anthropic.types.Message:
		return self.client.messages.create(**params)

	@minimum_response_elements
	@staticmethod
	def _parse_response(response: anthropic.types.Message, **kwargs) -> dict:
		"""
		Parses the Anthropic Message response into a dictionary matching the BaseClient contract.
		:param response: Raw Anthropic Message response
		:return parsed_response: Dictionary with text, token counts, and finish reason
		"""
		if not (hasattr(response, "content") and hasattr(response, "usage")):
			raise AttributeError("Unexpected Anthropic response structure")

		# Warn if the model hit max_tokens — batch parse may be incomplete
		finish_reason = response.stop_reason  # "end_turn" or "max_tokens"
		if finish_reason == "max_tokens":
			logging.warning(
				"[AnthropicClient] Response stopped at max_tokens limit. "
				"Batch output may be truncated — consider increasing max_tokens or reducing BATCH_SIZE."
			)

		# cache_read_input_tokens and cache_creation_input_tokens are present when
		# prompt caching is active; we count only non-cached tokens for billing tracking.
		input_tokens = response.usage.input_tokens
		cache_creation = getattr(response.usage, "cache_creation_input_tokens", 0) or 0
		cache_read = getattr(response.usage, "cache_read_input_tokens", 0) or 0

		parsed_response = {
			"text": response.content[0].text,
			"log_probs": None,  # Anthropic does not expose log probabilities
			"finish_reason": finish_reason,
			"input_tokens": input_tokens,
			"output_tokens": response.usage.output_tokens,
			"origin_tokens": "client",
			# Extra fields for cost tracking with prompt caching
			"cache_creation_tokens": cache_creation,
			"cache_read_tokens": cache_read,
		}

		return parsed_response

	@staticmethod
	def count_tokens(prompt: str, **kwargs) -> int:
		"""
		Approximates token count for an Anthropic prompt.
		Anthropic does not ship a public tokenizer library, so we use the
		cl100k_base tiktoken encoding as a close approximation (~5% error).
		For exact counts, use the Anthropic token-counting API endpoint instead.
		:param prompt: Prompt string
		:return: Approximate token count
		"""
		encoding = tiktoken.get_encoding("cl100k_base")
		return len(encoding.encode(prompt))


class InterfaceToLLM:
	def __init__(self, client: BaseClient, system_prompt: str, seed: int = None, max_tokens: int = None):
		"""
		Initializes InterfaceToLLM class, the specified client, the system prompt and the tokens usage counter
		:param client: Client child class of the BaseClient class to manage model connection
		:param system_prompt: Prompt string to define the model responses behaviour
		:param seed: Optional parameter to control the 'randomness' of the model response for reproducibility
		:param max_tokens: Optional parameter to limit the number of tokens used
		"""

		self.client = client
		self.seed = seed
		self.max_tokens = max_tokens

		# System prompt information
		self.system_prompt = system_prompt
		self.system_prompt_tokens = self.client.count_tokens(prompt=system_prompt)

		logging.info(
			f"\t[Interface to LLM]: client = {self.client.client_name}"
			f"\n\t (seed = {self.seed}, max tokens = {self.max_tokens})"
			f"\n\t-> System prompt ({self.system_prompt_tokens} tokens):\n\t{self.system_prompt}"
		)

		# Initialize tokens counters
		self.tokens_counter = 0
		self.output_tokens_counter = 0

	def authenticate(self, **kwargs):
		"""
		Authenticates the specified client, must be run before any other client methods
		:param kwargs: Client specific authenticate() parameters:
		- OpenAI: api_key
		"""

		self.client.authenticate(**kwargs)
		logging.info(
			f"Client = {self.client.client_name} successful authentication"
		)

	@client_authenticated
	def get_responses(self, prompts: list, temperature: float = 0.2, **kwargs) -> list:
		"""
		Obtains the responses to a list of prompts, by iterating with the 'get_response' function for each prompt
		:param prompts: List of prompts to be sent to the model and get responses
		:param temperature: Optional parameter to control the 'creativity' of the model response
		:param kwargs: Client specific send_prompt() parameters:
		- OpenAI: model, log_probs_bool
		:return parsed_responses: List of parsed responses to the given list of prompts
		"""

		if not isinstance(prompts, list):
			raise ValueError("Prompts must be a list")

		# Check for empty list
		if not len(prompts):
			raise PromptsEmptyError("Prompts must not be an empty list")

		logging.info(f"\t[Interface to LLM] <get_responses>: # of prompts = {len(prompts)}")

		# Iterates the list of prompts and obtains the parsed response for each one
		parsed_responses = []
		for prompt in prompts:
			parsed_response = self.get_response(prompt=prompt, temperature=temperature, **kwargs)
			parsed_responses.append(parsed_response)

		return parsed_responses

	@client_authenticated
	def get_response(self, prompt: str, temperature: float = 0.2, **kwargs) -> dict:
		"""
		Obtains the parsed prompt response and increments the tokens usage counter
		:param prompt: Prompt to be sent to the model get response
		:param temperature: Optional parameter to control the 'creativity' of the model response
		:param kwargs: Client specific send_prompt() parameters:
		- OpenAI: model, log_probs_bool
		:return parsed_response: Dictionary containing the response text as well as some additional information
		"""

		if not isinstance(prompt, str):
			raise ValueError("Prompt must be an string")

		# Check for empty prompt
		if prompt.strip() == "":
			raise PromptEmptyError("Prompt must not be an empty string or only contain whitespace characters")

		parsed_response = self.client.send_prompt(
			prompt=prompt,
			system_prompt=self.system_prompt,
			seed=self.seed,
			max_tokens=self.max_tokens,
			temperature=temperature,
			**kwargs
		)

		# Update tokens counters
		self.tokens_counter += parsed_response["input_tokens"] + parsed_response["output_tokens"]
		self.output_tokens_counter += parsed_response["output_tokens"]

		return parsed_response
import pytest
from interface_to_LLM.test.conftest import TestClient, TestOpenAIClient

from interface_to_LLM.interface_to_llm import InterfaceToLLM, OpenAIClient, minimum_response_elements
from interface_to_LLM.exceptions import (ClientNotAuthenticatedError, ClientAuthenticationError,
										 PromptEmptyError, PromptsEmptyError)


# -- InterfaceToLLM class tests --

@pytest.fixture
def create_interface_to_llm():
	interface = InterfaceToLLM(client=TestClient(), system_prompt="Hello World")
	yield interface


def test_interface_to_llm_init_attributes(create_interface_to_llm):
	"""
	Check InterfaceToLLM attributes when initialized
	"""

	interface = create_interface_to_llm

	assert hasattr(interface, "client") and interface.client.client is None
	assert hasattr(interface, "seed") and interface.seed is None
	assert hasattr(interface, "max_tokens") and interface.max_tokens is None
	assert hasattr(interface, "system_prompt") and interface.system_prompt is not None
	assert hasattr(interface, "system_prompt_tokens") and interface.system_prompt_tokens is not None
	assert hasattr(interface, "tokens_counter") and interface.tokens_counter == 0
	assert hasattr(interface, "output_tokens_counter") and interface.output_tokens_counter == 0


def test_interface_to_llm_init_attributes_optional_params():
	"""
	Check InterfaceToLLM attributes when initialized (with optional parameters)
	"""

	interface = InterfaceToLLM(client=TestClient(), system_prompt="Hello World", seed=1234, max_tokens=420)

	# Optional parameters
	assert hasattr(interface, "seed") and interface.seed == 1234
	assert hasattr(interface, "max_tokens") and interface.max_tokens == 420


def test_interface_to_llm_get_response_invalid_prompt_type(create_interface_to_llm):
	"""
	Check when given an invalid prompt type in InterfaceToLLM.get_response() the correspondent error is raised
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	# Some other types as parameter
	with pytest.raises(ValueError):
		interface.get_response(prompt=0)  # Integer
	with pytest.raises(ValueError):
		interface.get_response(prompt=[])  # List
	with pytest.raises(ValueError):
		interface.get_response(prompt={})  # Dict


def test_interface_to_llm_get_response_empty_prompt(create_interface_to_llm):
	"""
	Check when given an empty prompt in InterfaceToLLM.get_response() the correspondent error is raised
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	with pytest.raises(PromptEmptyError):
		interface.get_response(prompt="")  # Empty prompt
	with pytest.raises(PromptEmptyError):
		interface.get_response(prompt=" \n\t")  # Prompt with only whitespace characters


def test_interface_to_llm_get_response_tokens_counters(create_interface_to_llm):
	"""
	Check when given prompt in InterfaceToLLM.get_response() the token counters are incremented
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	# Record previous tokens counters
	prev_tokens_counter = interface.tokens_counter
	prev_output_tokens_counter = interface.output_tokens_counter

	# Send a prompt
	interface.get_response(prompt="Lorem ipsum.")

	assert prev_tokens_counter <= interface.tokens_counter
	assert prev_output_tokens_counter <= interface.output_tokens_counter
	assert interface.output_tokens_counter <= interface.tokens_counter  # Should always hold


def test_interface_to_llm_get_responses_invalid_prompts_type(create_interface_to_llm):
	"""
	Check when given an invalid prompts type in InterfaceToLLM.get_responses() the correspondent error is raised
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	# Some other types as parameter
	with pytest.raises(ValueError):
		interface.get_responses(prompts=0)  # Int
	with pytest.raises(ValueError):
		interface.get_responses(prompts="Lorem ipsum.")  # String
	with pytest.raises(ValueError):
		interface.get_responses(prompts={})  # Dict


def test_interface_to_llm_get_responses_empty_prompts(create_interface_to_llm):
	"""
	Check when given an empty list prompts in InterfaceToLLM.get_responses() the correspondent error is raised
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	with pytest.raises(PromptsEmptyError):
		interface.get_responses(prompts=[])


def test_interface_to_llm_get_responses_matching_lengths(create_interface_to_llm):
	"""
	Check when given prompts in InterfaceToLLM.get_responses() the number of responses equal the number of prompts
	"""

	# Initialize class and authenticate
	interface = create_interface_to_llm
	interface.authenticate()

	prompts = ["Hello World", "Lorem ipsum.", "This is a test"]
	# Record number of prompts
	n_prompts = len(prompts)

	responses = interface.get_responses(prompts=prompts)
	n_responses = len(responses)

	assert n_prompts == n_responses


# -- Client classes tests --

# OpenAIClient tests

@pytest.fixture
def create_OpenAIClient_interface_to_llm():
	interface = InterfaceToLLM(client=OpenAIClient(), system_prompt="Hello World")
	yield interface


def test_OpenAIClient_init_client_name(create_OpenAIClient_interface_to_llm):
	"""
	Check OpenAIClient client name attribute when initialized
	"""

	# Initialize class
	interface = create_OpenAIClient_interface_to_llm

	assert hasattr(interface.client, "client_name") and interface.client.client_name == "OpenAI"


def test_OpenAIClient_authenticate_failure(create_OpenAIClient_interface_to_llm):
	"""
	Check when given a wrong api key OpenAIClient.authenticate() raises the correspondent error and client is None
	"""

	# Initialize class
	interface = create_OpenAIClient_interface_to_llm

	with pytest.raises(ClientAuthenticationError):
		# Authentication failure
		interface.authenticate(api_key="Monkey")

	assert interface.client.client is None


def test_OpenAIClient_parse_response_missing_attributes(create_OpenAIClient_interface_to_llm):
	"""
	Check when given a response object with missing attributes OpenAIClient._parse_response() raises the correspondent error
	"""

	# Initialize class
	interface = create_OpenAIClient_interface_to_llm

	with pytest.raises(AttributeError):
		# Parse invalid response object with missing attributes (None)
		interface.client._parse_response(response=None)


@pytest.fixture
def create_TestOpenAIClient_interface_to_llm():
	interface = InterfaceToLLM(client=TestOpenAIClient(), system_prompt="Hello World")
	yield interface

def test_OpenAIClient_send_prompt(create_TestOpenAIClient_interface_to_llm):
	"""
	Check OpenAIClient.send_prompt() parsed response attributes
	"""

	# Initialize class and authenticate
	interface = create_TestOpenAIClient_interface_to_llm
	interface.authenticate()

	# Obtain parsed response
	parsed_response = interface.client.send_prompt(
		prompt="Lorem ipsum.",
		system_prompt=interface.system_prompt,
		seed=interface.seed,
		max_tokens=interface.max_tokens
	)

	assert isinstance(parsed_response, dict)
	assert "text" in parsed_response.keys() and isinstance(parsed_response["text"], str)
	assert "log_probs" in parsed_response.keys()  # logprobs type complicated to assert
	assert "finish_reason" in parsed_response.keys() and isinstance(parsed_response["finish_reason"], str)
	assert "input_tokens" in parsed_response.keys() and isinstance(parsed_response["input_tokens"], int)
	assert "output_tokens" in parsed_response.keys() and isinstance(parsed_response["output_tokens"], int)
	assert "origin_tokens" in parsed_response.keys() and parsed_response["origin_tokens"] in ["client", "calculated"]

def test_OpenAIClient_count_tokens(create_TestOpenAIClient_interface_to_llm):
	"""
	Check OpenAIClient.count_tokens() expected behavior
	"""

	# Initialize class
	interface = create_TestOpenAIClient_interface_to_llm

	short_prompt = "Lorem ipsum"
	long_prompt = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "*42

	# Count tokens with default model
	short_prompt_tokens = interface.client.count_tokens(prompt=short_prompt, model=interface.client.default_model)
	long_prompt_tokens = interface.client.count_tokens(prompt=long_prompt, model=interface.client.default_model)

	assert short_prompt_tokens == 2
	assert long_prompt_tokens == 421


def test_OpenAIClient_count_tokens_model_none(create_TestOpenAIClient_interface_to_llm):
	"""
	Check OpenAIClient.count_tokens() expected behavior when no model name specified (default tiktoken model)
	"""

	# Initialize class
	interface = create_TestOpenAIClient_interface_to_llm

	short_prompt = "Lorem ipsum"
	long_prompt = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "*42

	# Count tokens with model None (default tiktoken model)
	short_prompt_tokens = interface.client.count_tokens(prompt=short_prompt)
	long_prompt_tokens = interface.client.count_tokens(prompt=long_prompt)

	assert short_prompt_tokens == 2
	assert long_prompt_tokens == 421


def test_OpenAIClient_count_tokens_invalid_model(create_TestOpenAIClient_interface_to_llm):
	"""
	Check when given an invalid model name OpenAIClient.count_tokens() correspondent error is raised
	"""

	# Initialize class
	interface = create_TestOpenAIClient_interface_to_llm

	with pytest.raises(KeyError):
		# Count tokens with invalid model
		interface.client.count_tokens(prompt="Lorem Ipsum", model="404ModelNotFound")


# -- Decorators tests --

def test_decorator_client_authenticated(create_interface_to_llm):
	"""
	Check client_authenticated decorator raise error for all InterfaceToLLM class method that require authentication
	"""

	# Initialize interface without authentication
	interface = create_interface_to_llm

	# Methods that require authentication
	with pytest.raises(ClientNotAuthenticatedError):
		interface.get_response(prompt="Lorem ipsum.")
	with pytest.raises(ClientNotAuthenticatedError):
		interface.get_responses(prompts=["Lorem ipsum.", "Hello World"])


def test_decorator_minimum_response_elements():
	"""
	Check minimum_response_elements decorator raises error when function does not return dictionary with required keys
	"""

	@minimum_response_elements
	def mock_parse_response_with_no_minimum_elements():
		return {"a": 42}

	with pytest.raises(KeyError):
		mock_parse_response_with_no_minimum_elements()

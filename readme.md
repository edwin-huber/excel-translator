# Excel Translator
Version 1.0
2024-06-19

This python app will load an excel file and translate all text cells from English to a given language using the [Azure AI Translator](https://azure.microsoft.com/en-us/products/ai-services/ai-translator) service.
  
Header rows are not translated.
  
The target language can be chosen based on the list of supported languages returned by the endpoint.

# Usage

Make sure that you have at least [Python version 3.12.2](https://www.python.org/downloads/) installed.

Install the requirements using:

```pip install -r ./requirements.txt```

You will need an Azure subscription, and an Azure AI Translation service to set the following 2 Environment Variables:

```
TRANSLATOR_TEXT_SUBSCRIPTION_KEY
TRANSLATOR_TEXT_REGION
```

For more information on the [Azure AI Translator](https://azure.microsoft.com/en-us/products/ai-services/ai-translator) Service, please review the documentation.


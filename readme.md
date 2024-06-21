# Excel Translator
Version 1.0  
2024-06-19  

Welcome!  
This python project is for translating excel files with Azure AI!  
This sample python app is a handy tool that can automatically translate any text cells in an excel file from English to your desired language, using the powerful [Azure AI Translator](https://azure.microsoft.com/en-us/products/ai-services/ai-translator) service.  

You can choose from a wide range of languages supported by the Azure AI Translator service, which it will offer to you by querying the translator service’s endpoint.

As always, human oversight is recommended to make sure that the translations make sense!  

You don't have to worry about translating the header rows, as the app will skip them and only translate the content.
If you want it to translate the header row, take a look at the function ```translate_excel_file``` on line 88.

## What next?
Do you want it to translate from a different language?
Check out the ``` translate_text``` function on line 53.

## Usage

Make sure that you have at least [Python version 3.12.2](https://www.python.org/downloads/) installed.

Install the requirements using:

```pip install -r ./requirements.txt```

You will need an Azure subscription, and an Azure AI Translation service to set the following 2 Environment Variables:

```
TRANSLATOR_TEXT_SUBSCRIPTION_KEY
TRANSLATOR_TEXT_REGION
```

For more information on the [Azure AI Translator](https://azure.microsoft.com/en-us/products/ai-services/ai-translator) Service, please review the documentation.


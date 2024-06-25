### Introduction 
Hermes (the ancient Greek god who served as a messenger and guide to the other gods) is a proof-of-concept add-in for Outlook that uses Mistral Small Model to sanitize, summarize, reduce, expand, and rephrase text. It allows users to select a block of text and apply various text-related functions to it, such as removing sensitive information, summarizing long passages, and expanding abbreviations.

### Functionality
Say goodbye to spelling and grammar errors in your written content. Not only does it automatically sanitize the selected text, but it also can translate the text into a different language and measures the time saved compared to manual grammar checks. What's more, it even calculates the estimated CO2 emissions saved, making it a smart choice for both efficiency and sustainability.

### Mistral Le plataforme API
I built Hermes using Microsoft's Outlook add-in framework and Mistral Le Plataforme API. I used Python along with Flask to create an API endpoint for accessing the Mistral Le Plataforme API. For the user interface, I use HTML, and JavaScript.

### Requirements for local testing

You need to run the app.py python script in order to enable the rephrase endpoint, which will call the Mistral Le Plataforme API. You also need ngrok (with a payed account) to have a secure url, if this is not in place, you will get an error when importing the manifest.xml from outlook because it does not allow to work with local non http urls, such as http://127.0.0.1:5001, you need ngrok to serv it in a secure way.

For ngrok go to ngrok.com, register for free, install the client using brew (mac based users), use the token (follow the basic instructions) and then running with the command from below to get the url address.

You can start the endpoint locally by running python3 app.py, this will open the API listener locally.

python3 app.py (this will run the flask listener at port 5001, http://127.0.0.1:5001)

ngrok http http://127.0.0.1:5001 (this will give you the url, something like https://41cb-45-250-XXX-165.ngrok-free.app to be set in the manifest.xml and the commands.js file)

If you want to have this on cloud to avoid using ngrok, you can easily integrate in any webapp python based, this will give you a secure url to later call the rephrase endpoint.

### Installation
Instructions:

Go to Outlook and follow the next instructions.

On the right side you will see an icon with three dots...

Select 'Get Add-Ins'

Then select 'My Add-Ins'

In 'Custom Add-ins', select 'Add a custom add-in' > 'From URL' and insert this: xxx (commented for the moment, since is a POC) or From File and select the manifest.xml file available in the root directory.

This will open a small window and then click on 'Install'.

### Developer information

If you wish to get in contact or have any questions, feel free to reach out via email at [info@arananet.net](info@arananet.net)

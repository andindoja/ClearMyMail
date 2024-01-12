ClearMyMail
ClearMyMail is a desktop application designed to enhance email management efficiency. Its primary function is to help users identify and delete unwanted or spam emails in their Gmail account, streamlining their inbox and saving valuable time.

Features
Email Address Extraction: Gathers and compiles a list of all email addresses from the user's Gmail account, storing this data in an Excel file for easy review.
Targeted Email Deletion: Allows users to specify email addresses from which they want to delete emails. Users can upload an Excel file with this list, and ClearMyMail will handle the deletion process from the Gmail account.
Local Processing: Runs entirely on the user's device, ensuring data privacy and security. No personal data is transmitted externally or stored by the application.
Getting Started
To use ClearMyMail, follow these steps:

Clone the Repository: Clone this repository to your local machine using Git.
Install Dependencies: Ensure you have Python and the necessary libraries installed (see Dependencies section below).
Run the Application: Execute the main script to start the application.
Dependencies
ClearMyMail requires the following to run:

Python 3.x
Google API Python Client
Pandas Library
OAuth2Client
Install these dependencies using pip:

bash
Copy code
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pandas
Usage
After launching ClearMyMail, you will be prompted to authenticate with your Gmail account. Follow the on-screen instructions to allow the application to access your Gmail data.

Once authenticated:

The application will compile and save a list of email addresses in an Excel file.
You can then provide an Excel file with specific email addresses from which you want to delete emails.
Privacy and Security
Your privacy and data security are paramount. ClearMyMail does not store or transmit your data outside of your local device. All operations are performed locally, and your Gmail data is accessed securely using OAuth2 authentication.


Contact
For support, feedback, or inquiries, please contact Andi Ndoja at andi.ndoja@gmail.com.
# gmail_mail_merge
This script for a google sheet allows me to send email based off of several templates to a list of people and customize the email. It can send, bump, draft, test, and more. 

Setup:
1) Create a google sheets document with two sheets. On the first put the following columns:
- Field 1 to customize (i.e. First Name)
- Field 2 to customize (i.e. Company)
... (add as many as needed; can be changed)
- Email Address	
- cc
- bcc
- Template Row
- Status
- Thread ID
- Emails Sent

On the second put the following columns:
- Subject
- Body
- Signature

2) In tools > script editor you can add this script

3) Fill in:
- one or more templates in sheet 2:
  - Fields in sheet 1 are called upon with this format: ${"Field 1"}
  - Any formating should be done directly in html without linebreaks
- the customised fields in sheet 1:
  - "template row" (the row the template you have selected is on)
  - "email address" unless "status" is set to "test"
  - the "status" column is a sort of "command line", put in the action you wish to take. Prompts it can take below.
  - don't touch "thread ID or "Emails Sent"
  
5) ALWAYS TEST before sending emails out. Use the "test" or "test bump" command for this purpose. 
For this to work you need to add your email to the code manually in the place of the placeholder "your.email@address.com"

4) There is a snooze and a send later option but it must be configured with other scripts. For send later, find the other script here:
https://code.google.com/archive/p/gmail-delay-send/wikis/GmailDelayFAQ_8.wiki


"Status" column inputs / outputs:
  => Triggers an action:
    - "ready": send an email
    - "draft": make a draft
    - "bump": send a follow up
    - "draft bump": draft a follow up
    - "force draft bump": draft follow up despite error message.
    - "test": send to a test address.
    - "test bump": bump to a test address.

  => Information purposes:
    - "sent": email was sent by the script
    - "programmed": made a draft for a delayed send
    - "drafted": a draft has been created
    - "tested": sent to test address
    - "standby": no action to take yet (an empty column will be replaced by this tag)
    - "done": no need to do anything on this thread anymore
    - "..... bump": same as above, but for a follow up
    - "..... test": same as above but to a test address

  => Error messages
    - "bump fail": could not send the bump (Check the original email)
    - "input_error": mistake in the input (Check Log console for possible bugs)
    - "incorrect ID": the given Thread ID doesn't return anything, delete it.

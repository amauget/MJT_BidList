# A Bidder Contact List Automated for my Current Employer
NOTE: All actual contact info have been removed for the sake of confidentiality. The rows remain to represent the sheer size of this bid list.
### Background:
The general contractor that I work for sends personalized invitations to bid to dozens if not hundreds of subcontractors. The point of this sheet to to record whether companies have received an invitation to bid, whether they're planning to bid, and when they were contacted. 
### Why Excel/VBA?
The company is set on only using excel, as it's a software that the team members are comfortable using. This is a template that is copied every time a project comes out to bid.


## The major problems:
### Time Consuming:
The original spreadsheet was completely manual. If a user wanted to send a mass email, each email was to be copied 1 by 1, and pasted into an outlook email. An announcement email would easily take 15 minutes of copying and pasting before being ready to send. Copying multiple emails often lead to the copying of hidden cells.

### Inconsistent:
The manual method also lacked consistency. People would be overlooked during the copy/paste process and not receive vital announcements. Users often wouldn’t record the date of contact, or indicate whether an invite had been sent (highlight the row in <ins>***yellow***</ins>)

## The Fix:
### Add selector, and an Output Page:
I added an “X” toggle hyperlink to each contact row. When the “X” is selected, the row is automatically highlighted in yellow (indicating invite has been sent), and add a time stamp to the row. All selected rows contacts are output on the “Output” sheet. From there, a user can drag to select all emails, and copy at once.

### “Output” Sheet Compatible for Personalized Form Emails Or Manual Use
Some team members prefer to send personalized emails that use the contact’s name. The “Output” sheet can be targeted through Microsoft Word “Mailings” feature to add a contact’s first name before sending them an email.

Select All & De-Select All Buttons:
If all desired parties have been selected before, the “Re-Select-Subs” button can be selected to target all subs that need an email. If someone is highlighted <ins>***red***</ins>, they’re not interested in bidding, and will not be sent an email during the select all process. Time stamps will update when select all is used.


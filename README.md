# Mail-Merge-Email
Will take a mailing list from a google sheet and will email them based on a template

# Dependencies

* credentials and authorization for both the Sheets and Gmail API, and can be recieved by completing the following:
    * https://developers.google.com/gmail/api/quickstart/python
    * https://developers.google.com/sheets/api/quickstart/python
* python 2.6 or greater


# Instructions
First you must take the sheet ID by going onto the sheet document, and look into its url. From there you will see that it is formatted as such:
```
https://docs.google.com/spreadsheets/d/SHEET_ID/edit#gid=0
```

Copy the sheet ID, and assign it to the `SHEETS_FILE_ID` variable as a string. You are now able to pull from the spreadsheet! Congratulations.

Create a text document called 'message.html' in the working directory and write your message in HTML. Wherever you would like the program to insert data, simply put in a variable with the format `{{Variable Name}}` where variable name is the case sensitive name of the column in the google sheets. For example, if you have a table like so:
| First Name  | Last Name  |  Email |
|:--|:-:|--:|
| Jonn  | Doe  | john@doe.com  |
|  Jane | Who  |  jane@who.com |
| Jake  | Lum  | lummyboi@gmail.com  |

Then wherever in your template you want them to insert the recipients first name you would put `{{First Name}}` in the template. In this case, it will replaced by 'John' in the first email, 'Jane' in the second, and 'Jake' in the third. Any variables that do not corrrespond to a column in the spreadsheet will be left as is, and columns in the sheet that are not called will not be used.

You must set a subject, which you can do by assigning it to the `SUBJECT` variable, and finally you must set your email. This can be done by changing the `SENDER` variable to your current email address. If it does not match your credentials, then it will fail.

You must also set a `PURPOSE`, this is different than `SUBJECT`, as `PURPOSE` will be the title of the new column entered on your spreadsheet and will only appear there. If an email was sent successfully, then it will have a timestamp for that person, if not, then it will have `FALSE`. For example, if `PURPOSE = 'Mandatory Advising 4/20'` then your spreadsheet after sending will look like this:
| First Name  | Last Name  |  Email | Mandatory Advising 4/20 |
|:--|:-:|:-:|--:|
| Jonn  | Doe  | john@doe.com  | Mon Apr 27 13:58:57 2020 |
|  Jane | Who  |  jane@who.com | Mon Apr 27 13:58:57 2020 |
| Jake  | Lum  | lummyboi@gmail.com  | Mon Apr 27 13:58:57 2020 |


You can then run the program by typing in the commandline
```
python email-sender.py
```

And if successful, it will print out the email IDs in the console. NOTE: It will send out 5 emails every 5 seconds to prevent spam alarms.

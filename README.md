# Mail-Merge-Email
Will take a mailing list from a google sheet and will email them based on a template


# Instructions
First you must take the sheet ID by going onto the sheet document, and look into its url. From there you will see that it is formatted as such:
```
https://docs.google.com/spreadsheets/d/SHEET_ID/edit#gid=0
```

Copy the sheet ID, and paste into the `SHEETS_FILE_ID` variable as a string. You are now able to pull from the spreadsheet! Congratulations.

Create a text document called 'template.txt' in the working directory and write your message. Wherever you would like the program to insert data, simply put in a variable with the format `{{Variable Name}}` where variable name is the case sensitive name of the column in the google sheets. For example, if you have a table like so:
| First Name  | Last Name  |  Email |
|:--|:-:|--:|
| Jonn  | Doe  | john@doe.com  |
|  Jane | Who  |  jane@who.com |
| Jake  | Lum  | lummyboi@gmail.com  |

Then wherever in your template you want them to insert the recipients first name you would put `{{First Name}}` in the template. Any variables that do not corrrespond to a column in the spreadsheet will be left as is, and columns in the sheet that are not called will not be used.

You must set a subject, which you can do by assigning it to the `SUBJECT` variable, and finally you must set your email. This can be done by changing the `SENDER` variable to your current email address. If it does not match your credentials, then it will fail.

You can then run the program by typing
```
python email-sender.py
```

And if successful, it will print out the email IDs in the console.
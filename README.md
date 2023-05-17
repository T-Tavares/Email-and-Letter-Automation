That's an Google Scripts automation project to send Emails and Create Letters from Data extracted from an Google Sheets table.

On this fictional project I used a fictional Event Invitation Email and Letter to examplify how the script works
However, It can be used with any letter template. As long as the variables names on the templates match with the script.
We'll go through how to set up or edit it in a bit.

For most part of the code I've tried to make the functions names clear and you can follow the comments to understand/edit for your
own templates.

Here I highlighet the parts I believed to be more crucial to understand and make everything work.

// ------------- FOLDER STRUCTURE ------------- //

I Advice you to use the same folder structure as me to follow along with the script. But that's up to you.

On my Google Drive I've created this folders/files structure:

├── Google Sheets.xls
├── Letters Output
│   ├── letter-1.doc
│   ├── letter-2.doc
│   └── letter-3.doc
└── Templates
│ ├── Email Template.doc
│ └── Letter Template.doc
└── Files
├── file-1.pdf
├── file-2.pdf
└── file-3.pdf

Where the outputs of letters will depend on how many destinataries you have on your list
And you can have as many files as supported by Gmail.
I believe there is a 25mb limit for each sent email.

// ---------- WRITING YOUR TEMPLATES ---------- //

The vital thing about writing your templates is that the variables to be changed/personalized on each email or letter,
MUST match the ones on the script. In my case I used two curly braces to define them
ex:.

{{Destinatary}}
{{City}}

The second thing is that your paragraphs must have two <br><br> tags on it. You can play around adding more or taking one out.

In the end You should have something like that.

<details> 
<summary> Click to see how it should look like</summary>

Dear {{Destinatary}},
<br><br>
We are thrilled to invite you to the upcoming Tech Innovators Summit, where technology enthusiasts, industry leaders, and visionaries come together to explore the latest trends and innovations in the IT world. This premier event aims to foster networking opportunities, share knowledge, and inspire breakthrough ideas.
<br><br>
Event Details:
Date: {{Event Date}}
Time: {{Event Time}}
<br><br>

// ----------- GOOGLE SHEET LAYOUT ------------ //

The sheet layout is pretty straightfoward, you need a header on the first row to guide you and that's it.

`---------- ----------------- ----------
|   Name   |      Email      |  Phone   | <- HEADER
|----------|-----------------|----------|
| Jonh Doe | doe@example.com | 9-99-999 | <- DATA
 ---------- ----------------- ----------`

The script will run on your active Sheet.

</details>

// ---- SETTING UP THE INITIAL VARIABLES ----- //

<details>
<summary> Click to see how does the section look like</summary>

```// ---------------------------------------------------------------- //
// ----------------- EMAILS - MUST SET VARIABLES ------------------ //
// ---------------------------------------------------------------- //

const EMAILS_TEMPLATE_ID = 'ID OF YOUR EMAIL TEMPLATE';
const EMAILS_SHEET_NAME = 'YOUR EMAILS SHEET NAME'; // Name have to match your Sheet name on google sheets.
const EMAILS_SUBJECT = 'YOUR EMAIL SUBJECT PREFIX OR SUFIX';
const ATTACHMENT_FILES_IDS = ['ARRAY', 'LIST', 'OF', 'FILES', 'IDS']; // File must be uploaded on Google Drive.
```

</details>

You need to update the variables with the ID's on your own documents.

The ID's on Google look like that:
https://docs.google.com/document/d/**_2VJpe8ZAbR70-\_d1qr7DYPvZqy87Phevm_rTcmJeu348_**/edit

Where:
Document ID == 2VJpe8ZAbR70-\_d1qr7DYPvZqy87Phevm_rTcmJeu348

Replace the respective ID's on the variables on the beginning of the script as strings.

`const EMAILS_TEMPLATE_ID = '_2VJpe8ZAbR70-\_d1qr7DYPvZqy87Phevm_rTcmJeu348_';`

// ---- UPDATING YOUR OWN VARIABLES LINKS ----- //

On lines 75 and 132 of the script you'll find the code block that will link your data to the created letters/emails
replace those with the exact name you wrote on your templates. I've created variables and then called them on the replaceText() functions, but you can make it one in one line if you feel confortable with that.

```
const city = row[0];
const destinatary = row[1];
const destinataryEmail = row[2];

body.replaceText('{{Destinatary}}', destinatary);
body.replaceText('{{City}}', city);
```

Here the row[number] will depends on the order of the data on your sheet starting from zero.
Although Google Sheets Columns are read starting by one, on the script they are converted to a js array. That's why we count from zero.

     0              1             2

`---------- ----------------- ----------
|   Name   |      Email      |  Phone   | <- HEADER
|----------|-----------------|----------|
| Jonh Doe | doe@example.com | 9-99-999 | <- DATA
 ---------- ----------------- ----------`

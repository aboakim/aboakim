# What this work does

October 22, 2026

This note is for someone who does not need to read the program. It explains the email-import work that was built for AE Techron, and what is safe to hand to the client.

## The job

Larry Shank has an old Outlook mailbox saved as one file:

`C:\EmailArchive\Larry_Shank.pst`

A PST file is a copy of an Outlook mailbox. The goal is:

1. Open that file.
2. Read each email.
3. Create one row in a SharePoint list for each email.
4. Keep enough detail that a later step can classify the mail (customer, product, problem, and so on).

That program now exists in the folder `pst-sharepoint-import`.

## What the program saves for each email

- Subject
- Date
- Sender name and email address
- Recipients
- Folder inside the mailbox
- The text of the email
- The names of attachments

It also creates empty columns for later classification: technical relevance, customer, product, application, problem, resolution, and keywords. This version leaves those columns blank. It does not decide what an email is about.

## What happens if it is run twice

Each email gets an ID. The program writes that ID to a local list after a row is created. The next run skips emails already copied. It can also look at the SharePoint list and skip rows that are already there.

If one email fails, the rest continue. The failed email is tried again on the next run.

## What was checked

32 automated tests passed. Those tests use stand-in emails and a stand-in SharePoint connection. They check the rules: duplicate skipping, retries when SharePoint asks the program to slow down, sign-in refresh, column names, and reading the email fields.

The real mailbox file was not opened. The AE Techron SharePoint site was not called. A full import is still unproven.

## What you can give the client

Give them the `pst-sharepoint-import` folder and the README, as a zip or from a private project. Tell them it is a first version for a small trial.

It is not a finished import. Do not tell them the archive has already been copied into SharePoint.

Hold this link. It is a draft on a public personal GitHub profile, which is the wrong package to send as the delivery:

https://github.com/aboakim/aboakim/pull/1

## What the client still has to set up

1. Install the Outlook-file reader (`pypff`) on the Windows computer that has the PST. Installing the Python packages in `requirements.txt` does not install that reader. This is the step most likely to fail.
2. Create a Microsoft app sign-in and get admin approval to write to SharePoint.
3. Fill in the company tenant id, app id, SharePoint site id, and list id. The sample settings file still says `YOUR_...`.
4. Try 10 emails without writing to SharePoint, then write 1 real row:

```bash
python -m pst_sharepoint_import --config config.json --dry-run --no-sync-existing --limit 10
python -m pst_sharepoint_import --config config.json --ensure-columns --limit 1
```

5. Open that one SharePoint row. If the subject, sender, date, and text look right, run the rest of the mailbox.

## Limits to say in the same breath

- Attachment files are not uploaded. Only the file names are saved.
- The classification columns stay empty.
- Some old emails store the body as RTF. Those are marked and the text is not converted.
- Long emails are shortened so SharePoint will accept the row.

## Do not send these

Do not send `config.json`, `.env`, the `state` folder, or the PST file. They can contain the mailbox and the Microsoft sign-in.

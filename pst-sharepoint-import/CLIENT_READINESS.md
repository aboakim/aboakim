# Can this go to the client?

October 22, 2026

The importer is in good shape as a **pilot tool**. It is not ready to send the client as a finished import.

32 unit tests passed, using stand-in messages and a stand-in SharePoint connection. Nothing has been run against `Larry_Shank.pst` or the AE Techron tenant, so a full archive run is still unproven.

## What to send

Send the `pst-sharepoint-import` folder and its README, from a private project repo or a zip.

Hold this GitHub link. The work is a **draft** pull request on a **public** profile repo:

https://github.com/aboakim/aboakim/pull/1

That repository is the personal GitHub profile site. A client opening it would see public source and a draft pull request, which is the wrong package for a delivery.

## Before the client runs it

1. Install libpff and `pypff` on the Windows machine that has the PST. That step is separate from `pip install -r requirements.txt` and is the most likely setup failure.
2. Create an Entra app and grant admin consent for `Sites.ReadWrite.All`, or `Sites.Selected` on this one site.
3. Fill in the tenant id, client id, SharePoint site id, and list id. The sample config still has placeholders.
4. Run a 10-message preview, then one real row:

```bash
python -m pst_sharepoint_import --config config.json --dry-run --no-sync-existing --limit 10
python -m pst_sharepoint_import --config config.json --ensure-columns --limit 1
```

5. Open that one SharePoint row and confirm the subject, sender, date, and body look right. Then run the rest.

## Say this up front

The tool stores attachment **names** only, leaves the AI columns blank, and does not convert RTF bodies to text.

## Leave these out of the handoff

Do not send `config.json`, `.env`, `state/`, or the PST itself. Those can contain the mailbox and sign-in tokens.

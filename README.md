# Thunderbird → Exchange Online Migration

Tools and documentation for a staged migration from a Thunderbird/legacy-hosted
mailbox setup to Microsoft 365 / Exchange Online, done one mailbox at a time
rather than as a single company-wide cutover.

## What's in here

- **`scripts/migrate_archive.py`** — uploads a Thunderbird mbox folder tree
  (the classic `Foldername` + `Foldername.sbd` pattern) straight into an
  Exchange Online mailbox over IMAP, recreating the folder hierarchy as it
  goes. Written to work around Thunderbird's own drag-and-drop copy, which
  failed (`TRYCREATE` errors) on deeply nested folders. Standard library
  only — nothing to install. See the docstring at the top of the script for
  full usage details.
- **`scripts/config.example.json`** — template for the script's
  configuration (tenant ID, app registration client ID, mailbox username,
  local archive path). Copy it to `private/config.json` and fill in your own
  values — see "Configuration" below.
- **`docs/CHANGELOG.md`** — a record of the Exchange Online / mail-flow
  configuration changes made during this migration (mailbox aliasing, DKIM,
  SPF, accepted-domain type, etc.), with the reasoning behind each change and
  what it did and didn't fix. Specifics that don't belong in a public repo
  (exact addresses, internal IPs, raw config values) are kept in `private/`
  instead — see below.

## Configuration

The script needs a few values specific to your own tenant and mailbox. These
are **not** committed to this repo:

1. Copy `scripts/config.example.json` to `private/config.json`.
2. Fill in your `tenant_id`, `client_id` (your own "device code" app
   registration), `username`, and `source_dir` (the local path to your
   Thunderbird `Archives.sbd` folder).
3. Run the script as usual from `scripts/`.

## The `private/` folder

`private/` is git-ignored (see `.gitignore`) and never gets pushed. It's
where the real, un-redacted versions of anything sensitive live locally:

- `private/config.json` — the real tenant ID, app client ID, mailbox
  username, and local archive path used for this migration.
- `private/CHANGELOG-full.md` — the full, un-redacted version of
  `docs/CHANGELOG.md`, with real names, addresses, and infrastructure
  details, kept as the authoritative internal record.

## Background

This migration is being done in stages: one mailbox is fully cut over to
Exchange Online at a time (mail flow verified, DMARC/SPF/DKIM alignment
fixed, etc.) before moving on to the next, rather than migrating the whole
organization in one go. `docs/CHANGELOG.md` documents the mail-flow-level
changes that came out of migrating the first mailbox; the same steps should
generalize to each mailbox that follows.

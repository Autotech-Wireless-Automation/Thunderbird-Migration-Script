# Thunderbird → Exchange Online Migration

Tools and documentation for a staged migration from a Thunderbird/legacy-hosted
mailbox setup to Microsoft 365 / Exchange Online, done one mailbox at a time
rather than as a single company-wide cutover.

## What's in here

- **`scripts/migrate_archive.py`** — uploads Thunderbird mbox folder trees
  (the classic `Foldername` + `Foldername.sbd` pattern) straight into an
  Exchange Online mailbox over IMAP, recreating the folder hierarchy as it
  goes. Written to work around Thunderbird's own drag-and-drop copy, which
  failed (`TRYCREATE` errors) on deeply nested folders. Supports migrating
  several source accounts/profiles in one run, each landing under its own
  top-level folder in the destination mailbox, and an optional list of
  folder names (e.g. `Trash`) to leave out entirely. Standard library
  only — nothing to install. See the docstring at the top of the script for
  full usage details.
- **`scripts/config.example.json`** — template for the script's
  configuration (tenant ID, app registration client ID, mailbox username,
  and a `sources` list of local-folder → destination-folder mappings, plus
  an optional `skip_folder_names` list). Copy it to `private/config.json`
  and fill in your own values — see "Configuration" below.
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
   registration), and `username` (the destination mailbox).
3. Fill in `sources` — a list of `{"source_dir": ..., "dest_top_folder": ...}`
   pairs. Each `source_dir` is a local directory holding one or more
   Thunderbird top-level mbox files directly (for example a Thunderbird
   `Mail/<host>` account folder, or a `Local Folders` directory) — everything
   found in it, subfolders included, is recreated under `dest_top_folder` in
   the destination mailbox. Add one entry per account/profile you're
   migrating; they all land in the same mailbox, each under its own
   top-level folder, sharing one migration_state.json.
4. Optionally set `skip_folder_names` — folder names (matched
   case-insensitively, at any depth, in any source) to leave out of the
   migration entirely, folder and contents both. Used to skip `Trash`
   folders when migrating a live profile.
5. Run the script as usual from `scripts/`.

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

# Exchange Online / Mail Flow Configuration Changelog

**Microsoft 365 tenant:** this organization's tenant
**Scope:** Changes made to Exchange Online (Microsoft 365 admin center / Exchange admin center) and the legacy mail hosting panel while migrating one mailbox's daily mail to Outlook / Exchange Online, as part of a staged, per-user cutover (the rest of the organization remains on the legacy mail system for now).

> A small number of specifics in this changelog (exact employee names/addresses, internal IP addresses, and the raw config values used by the migration script) have been kept out of this public copy and are recorded instead in the git-ignored `private/` folder for internal reference.

Context predating this record (already complete, not detailed here): Exchange Online tenant provisioned; a cold-storage Thunderbird mail archive was migrated into the primary mailbox via the script in this repo.

---

## 1. Exchange admin center — mailbox configuration

### 1.1 Secondary email alias added to the primary mailbox
- **Where:** Exchange admin center → Mailboxes → [primary user] → Manage email address types
- **Change:** Added a secondary SMTP alias on the tenant's own `onmicrosoft.com` domain to the mailbox being migrated (primary address unchanged).
- **Why:** Needed as the forwarding target for the legacy-side forwarding rule (see §2.1) — the tenant's own default domain address, unambiguous and unaffected by the accepted-domain routing behavior of the custom domain itself.
- **Status:** ✅ Active.

### 1.2 DKIM signing enabled for the custom domain
- **Where:** Microsoft Defender portal → Email & collaboration → Policies & rules → DKIM
- **Change:** Custom-domain DKIM signing turned **On**, using selectors `selector1` / `selector2`, once the corresponding CNAME records existed in DNS (see §3.2).
- **Why:** Exchange Online signs outbound mail with the tenant's default `onmicrosoft.com` domain unless custom-domain DKIM is explicitly enabled. Outbound mail was previously signed under the tenant's default domain, which does not align with the custom `From:` domain under DMARC — this was the direct cause of DMARC-rejection bounces from recipients enforcing `p=reject`.
- **Verification:** DKIM status page showed **Valid / Enabled** after DNS propagation; confirmed by a successful two-way mail test with an external mailbox.
- **Status:** ✅ Active.

### 1.3 Mailbox forwarding — attempted for three other already-provisioned mailboxes (reverted)
- **Where:** Exchange admin center → Mailboxes → [each user] → Email forwarding
- **Change attempted:** "Forward to an external email address" set to each user's own address on the custom domain, with "Deliver message to both forwarding address and mailbox" enabled, intended as a bridge to relay mail landing in their (unused) Exchange mailboxes back out to the legacy system they actually use.
- **Outcome: did not work, and was reverted.** Because the custom domain was (at the time) an **Authoritative** accepted domain and each of these users already has a mailbox object in Exchange Online, Exchange resolves their address to their own Exchange mailbox before consulting the forwarding setting — the "forward" was silently absorbed and mail simply stayed in their Exchange mailbox (confirmed via mailbox-size growth and a live test that never reached the legacy mailbox).
- **Current state:** Forwarding switched back **Off** on all three mailboxes to avoid leaving an ineffective, unnecessary external-forwarding configuration in place.
- **Open item:** These three mailboxes are provisioned in Exchange but not used day-to-day (their owners are still on the legacy system). Mail sent to them from Exchange currently goes unread in those dormant mailboxes. No fix applied yet — options under consideration: (a) have them check Outlook on the web periodically until migrated, (b) build a dedicated relay/sync script, (c) migrate each of them the same way the primary mailbox was migrated, whenever there is time to do so.

### 1.4 Accepted domain type changed: Authoritative → Internal Relay
- **Where:** Exchange admin center → Mail flow → Accepted domains → [custom domain]
- **Before:** `Authoritative` — "Email is delivered to email addresses that are listed for recipients in Microsoft 365. Emails for unknown recipients are rejected."
- **After:** `Internal relay` — "Recipients for this domain can be in Microsoft 365 or your own email servers. Email is delivered to known recipients in Office 365 or is relayed to your own email server if the recipients aren't known to Microsoft 365."
- **Why:** With the domain set to Authoritative, any mail sent from Exchange Online to an address on the custom domain **not** provisioned in Exchange was hard-rejected by Exchange itself with `550 5.1.10 RESOLVER.ADR.RecipientNotFound`, before it ever reached the organization's real mail system. This was confirmed directly by a non-delivery report when mail was sent to a not-yet-enrolled colleague's address.
- **Effect of the change:** Mail to any recipient on the custom domain not known to Exchange Online now relays out via the domain's existing MX record (still pointing at the legacy mail host) instead of bouncing. This is the standard configuration for a domain in a staged/partial cloud migration (some mailboxes in Exchange, everyone else still on the legacy system).
- **Verification:** The same message was re-sent after the change and delivered successfully.
- **Does not affect:** Mail to mailboxes Exchange already knows about — that continues to be delivered inside Exchange as before (see §1.3 open item).
- **Status:** ✅ Active, confirmed working.

---

## 2. Legacy mail panel — mailbox forwarding

### 2.1 Forwarding enabled on the primary legacy mailbox
- **Where:** legacy mail panel → the primary mailbox → Forwarding settings
- **Change:** Mail forwarding enabled, forwarding target set to the tenant's own `onmicrosoft.com` alias (§1.1). Copies are currently still being delivered to the legacy mailbox as well.
- **Why:** The domain's MX record was intentionally left unchanged (still points to the legacy mail host) to avoid affecting any other user's mail while this one mailbox is cut over. This legacy-side forwarding rule is what bridges external mail addressed to the primary mailbox into its Exchange Online counterpart.
- **Open item:** Once fully confident in the Exchange Online mailbox, the "don't deliver a copy to the legacy mailbox" option can be enabled to stop dual storage of incoming mail (deferred future step, not yet done).

---

## 3. DNS records (managed via the domain registrar / hosting panel's DNS zone editor)

### 3.1 SPF (TXT record)
- **Change:** Added Microsoft's SPF include (`include:spf.protection.outlook.com`) to the existing SPF record, alongside the mechanisms already in place for the legacy mail host and other existing senders.
- **Why:** The domain's DMARC policy enforces `p=reject`. Outbound mail sent via Exchange Online was not covered by the existing SPF record, so it failed SPF alignment (on top of the DKIM alignment problem in §1.2), causing DMARC rejections at strict recipients. Adding Microsoft's SPF include allows Exchange Online's sending infrastructure to pass SPF for the domain.
- **Verified** via direct DNS query after the edit.

### 3.2 DKIM (two CNAME records added)
| Name | Type | Points to |
|---|---|---|
| `selector1._domainkey.<your-domain>` | CNAME | `selector1-<tenant>._domainkey.<tenant>.onmicrosoft.com` |
| `selector2._domainkey.<your-domain>` | CNAME | `selector2-<tenant>._domainkey.<tenant>.onmicrosoft.com` |

- **Why:** Required by Microsoft before custom-domain DKIM signing can be enabled (§1.2). The exact CNAME targets are shown on the tenant's DKIM setup page in Microsoft Defender when signing is first turned on.
- **Verified** via direct DNS query after the edit; both selectors resolved correctly, and the Defender DKIM page subsequently showed status **Valid**.

---

## 4. Summary of current state

| Item | Status |
|---|---|
| Primary mailbox mail (send/receive, internal + external) | ✅ Fully working via Outlook / Exchange Online |
| DMARC / SPF / DKIM alignment for the custom domain | ✅ Fixed — outbound mail passes DMARC |
| Mail to non-enrolled addresses on the custom domain | ✅ Fixed — now relays out via MX instead of bouncing |
| Mail to other already-provisioned-but-unused Exchange mailboxes | ⚠️ Open — delivered into their unused Exchange mailboxes, not seen by them; no fix applied yet |
| Domain MX record | Unchanged — still points to the legacy mail host |
| Domain accepted-domain type | Internal Relay (changed from Authoritative) |

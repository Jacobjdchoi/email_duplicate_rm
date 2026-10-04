# Outlook Duplicate Sweeper

Finds and removes duplicate emails that Outlook Classic cannot recognise as duplicates —
in particular the copies you receive when one sender addresses several of your own
addresses in the same message.

- **Runs as a single `.exe`.** No installer, no admin rights, no macros, no add-in registration.
- **Nothing is deleted.** Duplicates are *moved*, and there is a one-click **Undo last run**.
- **Every duplicate group always keeps exactly one copy.** This is enforced in code, not by convention.

---

## Why Outlook misses these

When someone sends one email to three of your addresses, that is **one message** on their side.
Their mail server fans it out into three mailboxes. Every copy therefore carries the **same
RFC 5322 `Message-ID` header**, generated once by the sender before the fan-out.

What differs between the copies is only the local MAPI identity — `EntryID`, `PR_SEARCH_KEY` —
and the `Received:` trace headers. Outlook's Clean Up works on conversation membership and
store-local identity, and never compares across separate accounts at all, so it cannot see these
as related.

This tool matches on `PR_INTERNET_MESSAGE_ID` (MAPI tag `0x1035001F`) — an exact string
comparison, not a fuzzy subject/body heuristic — **and then corroborates it against
`PR_CLIENT_SUBMIT_TIME`**, because a shared Message-ID by itself turned out not to be proof.
See *Message-ID is necessary but not sufficient* below; that correction matters more than
anything else on this page.

**Why exact matching matters here.** Measured on a real pair of cross-account copies:

```
"AW: AW: CoC delivera CW 8/2026"
  Jacob_SamTra_NEW_   size = 159569
  ChoisTechnics       size = 165427     <- same message, different byte size
```

The copies are *not* byte-identical, because each carries its own `Received:` headers.
Any size- or content-hash-based matcher would miss them. The `Message-ID` catches them exactly.

---

## Running it

Double-click **`dist\OutlookDupeSweeper.exe`** with Outlook already open.

> Both must run at the same privilege level. If Outlook is *not* elevated, don't run this
> "as administrator" either, or Windows will refuse the connection between them.

### First run

1. **Folders to scan** (left panel) — a tree of every mail folder in every data file, each
   showing its item count. Ticked by default: every folder of every delivering account except
   Sent, Drafts, Outbox, Junk and Deleted Items. Archive `.pst` files are listed but unticked.
   The **All / None / Inboxes only** buttons re-tick in bulk.
2. **Look back** — start with *Last 12 months*. *All mail* works but is much slower.
3. **Scan for duplicates.**
4. Review the list. Rows in **green are the copy being kept**; ticked rows will be moved.
   The summary line states how many folders and messages were actually read.
5. **Move duplicates...** and confirm.

> **Scan the folders, not just the Inbox.** Outlook's *designated* Inbox is often not where
> the mail is. In the profile this was developed against, one `.pst` held 70,804 messages while
> its designated Inbox was empty, and another kept 25,438 messages in an Inbox subfolder called
> 참조. An Inbox-only scan saw none of it.

### Reading the results

| | Meaning |
|---|---|
| `Group` column | Group number. Every row sharing a number is one message. |
| Green row, `KEEP` | The copy that survives |
| Ticked row, `REMOVE` | Will be moved |
| **Amber row** | The copies disagree on subject. **Never ticked automatically** — right-click to act on one. |
| Orange text | This copy is flagged or categorised and is about to be moved — worth a look |
| `Why grouped` | `several addresses`, `two data files`, `same mailbox`, or `sender reuses IDs` |

The `Sent` column is the one that matters when judging a group: every copy of one message
shares it. If two rows in a group show different `Sent` times, that is a bug — report it.

**Right-click any row:**
- *Keep this copy instead* — makes that copy the survivor
- *Skip this whole group* — leaves all copies alone
- *Open this message in Outlook*

The **Across accounts only** / **Within one account only** filters separate the two distinct
problems (see *Two kinds of duplicate* below).

---

## Settings

**Which copy to keep.** Order your accounts; the copy in the highest-listed account wins.
This decides which inbox ends up holding your mail, and therefore which address you reply from.

Full ranking, in order: account priority → prefer a copy you have flagged or categorised
(so your own annotations survive) → earliest received → EntryID (so repeat runs are identical).

**What to do with the extra copies.**
- *"Duplicates" folder* (default) — a folder per account. Easy to audit; drag anything back.
- *Deleted Items* — ordinary Outlook deletion.
- *Flag only* — changes nothing, just marks them. Good for a trial run.

**Matching.**
- *Also remove duplicates inside a single account* — POP3 re-downloads (see below).
- *Also match messages with no `Message-ID`* — **off by default.** When off, only exact
  `Message-ID` matches are ever touched. On a real 21,625-message scan this option would
  have changed nothing, so leaving it off costs you almost nothing and removes all guesswork.

---

## Two kinds of duplicate

An early **Inbox-only** scan of one real profile (21,625 messages, last 12 months) found
**2,315 duplicate groups**. Now that whole folder trees are scanned the real totals are higher,
but the split between the two causes is the point:

| Kind | Groups | Cause |
|---|---|---|
| **Across accounts** | 1,014 | One sender, several of your addresses — the problem this tool was built for |
| **Within one account** | 1,301 | The *same* mailbox holding two copies of one message |

The second kind is a separate problem: POP3 accounts configured to *leave a copy on the server*
can download the same message twice. Both copies are real, distinct items in the same Inbox with
the same `Message-ID`. If you want only the cross-account cleanup, untick
*Also remove duplicates inside a single account*.

---

## Safety

- Duplicates are **moved, never erased**. The chosen destination is always recoverable.
- A group is **never emptied**. If you tick every copy in a group, that group is refused and
  reported instead of being acted on.
- **Only what the list is showing can be acted on.** With *Across accounts only* selected,
  the within-account groups are not merely hidden — they are outside the action entirely.
- The same item is **never counted twice**. Rows are de-duplicated by `EntryID`, so a message
  reachable through two paths (a search folder, say) cannot be grouped against itself.
- Every run writes a CSV log to `%APPDATA%\OutlookDupeSweeper\logs\`.
- **Undo last run** reads the newest log and moves each message back to the folder it came from.

> `EntryID` changes when an item moves between folders, so the log records the *new* ID
> returned by the move alongside the original folder. That is what makes undo work.

---

## Installing, and giving it to a colleague

Hand them **one file: `Outlook-Duplicate-Sweeper-Setup.exe`** (831 KB). Copy it to a USB stick
or a share. They double-click it and click **Install**. That is the whole procedure.

Everything is inside that file — the application itself and the instructions — so nothing else
has to travel with it. There is no folder to copy, no unzipping, and no second file to lose.

Verified end-to-end by putting the `.exe` alone in an otherwise empty folder and installing
from there:

| Produced by the lone file | |
|---|---|
| `OutlookDupeSweeper.exe` | byte-identical (SHA-256) to the built application |
| `HOW-TO-RUN.txt` | extracted from the embedded copy |
| `Uninstall.exe` | local copy, so removal works after the USB is gone |
| Start menu + desktop shortcuts | correct target, icon and working directory |
| Settings → Apps entry | name, version, publisher, working uninstall path |

Uninstalling then removed all of it and left `%APPDATA%\OutlookDupeSweeper` (settings and undo
logs) untouched.

Colleagues who want to read the instructions *before* installing can click
**"What does this do?"** in the setup window — the full guide is inside the `.exe` too.

The app then lives in the Start menu — press Windows, type "duplicate". It also appears in
**Settings → Apps → Installed apps** with a working uninstaller.

`Setup.exe` contains the application embedded inside it, so it is the only file that needs to
travel. It installs per-user to `%LOCALAPPDATA%\Programs\OutlookDupeSweeper`, needs **no
administrator rights**, and writes nothing outside the user's own profile.

**Do not attach it to an email.** Outlook blocks executable attachments, and most mail gateways
strip them even inside a `.zip`. Use a **USB stick** or a network share — which, as a bonus,
avoids the SmartScreen prompt entirely.

Two things to tell them:

- **A USB stick or network share does not trigger SmartScreen.** Windows only attaches the
  "downloaded from the internet" mark to files that arrive over the web, so copying from a
  stick installs with no warning. If they *do* download it, SmartScreen may show *"Windows
  protected your PC"* — **More info → Run anyway**, once. The only real fix is a paid
  code-signing certificate, not worth buying for internal use.
- **Classic Outlook only.** The tool reaches mail through MAPI/COM, which "New Outlook" does not
  expose at all. If their New Outlook toggle is on, it will find nothing. This is also why a
  modern Office.js add-in could not have done this job.

Settings are per-user in `%APPDATA%\OutlookDupeSweeper\` and survive updates and uninstalls, so
everyone sets their own account priority once.

### Unattended install

For scripted or IT-managed rollout:

```
Setup.exe /silent                    installs with a desktop shortcut, no UI
Setup.exe /uninstall /silent         removes it, no prompts
```

Both set exit code 0 on success, 1 on failure.

### Updating

Run `build.cmd` then `build-setup.cmd`, **bump `Version` in `setup\Setup.cs` first**, and send
the new `Outlook-Duplicate-Sweeper-Setup.exe`. Running it over an existing install updates in
place — the button reads *Update* and existing settings are untouched. The version is what a
colleague sees in Settings → Apps, and the only way to tell two rollouts apart, so a build
shipped without bumping it is a build nobody can identify later.

---

## Rebuilding after a change

```
build.cmd           ->  dist\OutlookDupeSweeper.exe
build-setup.cmd     ->  dist\Setup.exe        (embeds the above; run build.cmd first)
tools\build-icon.cmd->  assets\app.ico        (only if the icon design changes)
```

Everything compiles with `csc.exe` from the .NET Framework, which ships with Windows.
There is no toolchain to install.

Source layout:

| File | Purpose |
|---|---|
| `src\Model.cs` | Data types and persisted settings |
| `src\OutlookAccess.cs` | All Outlook COM interaction (late-bound; no interop assemblies) |
| `src\Engine.cs` | Grouping and keeper selection — no COM, easy to reason about |
| `src\MainForm.cs` | Main window |
| `src\SettingsForm.cs` | Settings dialog |
| `src\Program.cs` | Entry point |
| `setup\Setup.cs` | Installer; embeds the app as a resource |
| `tools\IconGen.cs` | Draws `app.ico` at nine sizes |

**Icon note.** `IconGen` writes every size as a classic DIB rather than PNG. `System.Drawing.Icon`
cannot parse PNG-compressed icon entries — verified: 16/32/48 as DIB load fine while 128/256 as
PNG throw *"Requested range extends past the end of the array"*. Explorer copes with PNG entries,
but the app loading its own window icon does not. That costs a few hundred KB and removes the
failure mode.

`probe\` holds the diagnostics used against a real mailbox. None are needed to run the tool.

| | |
|---|---|
| `Probe.cs`, `Verify.cs` | Validated the approach before any deletion code was written. Read-only. |
| `Diag.cs` | Found the three faults above. Read-only apart from one folder create/delete test. |
| `Suspect.cs` | Audits whether grouped copies actually agree with each other. This is what caught the Message-ID reuse. Read-only. |
| `Validate.cs` | Diffs the old matching rule against the shipping one, in both directions. Read-only. |
| `TestMove.cs` | End-to-end move/undo check — moves one real message and restores it. |
| `Cleanup.cs` | Removes the empty test folders `Diag.cs` creates. |

`Suspect.cs` is the one to re-run if the matching is ever doubted again: it reproduces the
saved scan and prints every group whose members disagree on subject, sender or submit time.
A correct rule produces none.

---

## Message-ID is necessary but not sufficient

The original design matched on `Message-ID` alone, on the reasoning that the sender generates
it once before the message fans out. That reasoning is correct. The assumption that a
*shared* Message-ID therefore implies a shared message is **not** — and it failed badly on
real mail.

Audited against 333,036 messages in this profile, the ID-only rule produced groups like these:

```
one Message-ID  <!&!GAAAAAAAAAA+mdUMKE/qT5NOUDjeizRw4o...>   shared by 73 messages
   2024-05-31   2024년 05월 수주 진행 사항 송부의 건_20240531   3,321,465 b
   2025-08-14   2025년 08월 수주 진행 사항 송부의 건_20250814   2,762,796 b
   ...  two years of weekly reports, every one a different email  ...
   2026-08-21   2026년 08월 수주 진행 사항 송부의 건_20260821     735,001 b
```

```
one Message-ID  <!&!AAAAAAAAAAAYAAAAAAAAABHktpAB...@samtra...>   shared by 2 messages
   2026-02-26   RE: 이엠시스 커넥터 긴급 견적 요청건_260225   1,823,664 b
   2026-03-06   RE: 이엠시스 커넥터 긴급 견적 요청건_260225     523,800 b   <- 8 days later, 1.3 MB smaller
```

The senders responsible emit IDs shaped `<!&!AAAAAAAA…>`, and they are **per conversation,
not per message**. Matching on ID alone would have moved 53 genuine, distinct emails.

**The corroborator is `PR_CLIENT_SUBMIT_TIME`** (`0x00390040`). The sender stamps it once,
before fan-out, so every copy of one message carries the same value — while two replies in a
thread never do. Copies are grouped only when their IDs match *and* their submit times agree
within five seconds. The tolerance exists because two transports can round the timestamp
differently; five seconds cannot merge thread replies, which are minutes or days apart.

**Subject is deliberately not required to match.** Gateways rewrite it — measured here, one
copy of a message arrives as `AW: CoC delivera CW 8/2026` and the other as
`[SPAM] AW: CoC delivera CW 8/2026`. Requiring subject equality would have broken those
genuine pairs. Subject is used only as a *warning*: a group whose copies disagree on subject
after normalisation is shown in amber and never ticked automatically.

Measured effect of the correction on the same 333,036 messages:

| | ID only | ID + submit time |
|---|---|---|
| Groups | 1,967 | 1,993 |
| Copies ticked for removal | 2,081 | 2,032 |
| Genuine emails wrongly targeted | **49** | 0 |
| Groups that agreed on subject *and* submit time but got dropped | — | **0** |

The worst single group went from *one group of 73 → 72 removals* to *26 groups → 26 removals*,
all of them same-second, same-subject pairs.

Verified by `probe\Validate.cs`, which reimplements the old rule locally and diffs both
directions — the prevented removals *and* the false negatives a stricter rule could introduce.
It also measures the near-misses the tolerance exists for: 4 groups here have copies whose
submit times differ by under a second, and an exact-equality rule would have split all four.
None differ by 1–120 s, so the 5 s window is comfortably clear of both failure modes.

---

## Faults found after the first real use, and what they were

Three defects were confirmed against the live profile rather than reasoned about:

| Symptom | Cause | Evidence |
|---|---|---|
| Nothing was ever moved | `Folders.Add(name, 0)` — `Folders.Add` takes an `OlDefaultFolders` value, **not** an `OlItemType`. `0` is not a member of it. The `"Duplicates"` folder could never be created, so every move failed. | *"Value does not fall within the expected range"* on all 12 stores. Both run logs contained a header row and nothing else. |
| Most mail was never scanned | Only `GetDefaultFolder(olFolderInbox)` was read. | One `.pst` holds 70,804 messages with an **empty** designated Inbox; another keeps 25,438 in an Inbox subfolder. Ticking such a store contributed exactly zero rows — which is why the duplicate count did not move when far more mail was selected. |
| A broken data file hid every store after it | One `try`/`catch` wrapped the whole store loop in `GetStores`. | Two `.pst` files listed in the profile are missing from disk; touching either threw and truncated the list. |

The table read loop itself was verified correct at the same time: row counts came back exactly
equal to `Folder.Items.Count` (9,578 and 2,537) with no short reads.

The move and undo path is now exercised end-to-end by `probe\TestMove.cs`, which drives the
shipping `OutlookAccess` / `Engine` classes against one real message and puts it straight back.

---

## Known limits

- **The Exchange store is slow to scan.** In testing, 1,394 recent messages took 23 seconds
  because the date filter is evaluated server-side across the whole 77,000-item folder. The
  POP3 stores returned ~21,000 messages in under 1.1 seconds. This is Outlook/Exchange
  behaviour, not something the tool controls.
- **The To:-vs-Cc: tiebreaker is not available.** It relies on `PR_MESSAGE_TO_ME` /
  `PR_MESSAGE_CC_ME`, which the Exchange transport stamps at delivery. POP3 accounts never
  receive them, so on a POP3-heavy profile they are always false and useless as a signal.
  Account priority and flag-preservation are used instead.
- **About 0.2% of messages carry no `Message-ID`** (45 of 21,625 measured). With fallback
  matching off, those are simply left alone.
- **Only mail items are considered** (`IPM.Note*`). Meeting requests, delivery receipts and
  calendar items are skipped.
- **Sent Items is excluded by default, deliberately.** If you copy yourself on a message, your
  sent copy and the received copy share a `Message-ID`, so including Sent would offer to remove
  one of them. You can tick it in the tree, but do so knowing that.
- **A `.pst` listed in the profile but missing from disk cannot be scanned.** It is reported at
  startup rather than passed over quietly.
- Scanning runs on the UI thread and pumps messages, because Outlook's object model is
  apartment-threaded and driving it from a worker thread invites `RPC_E_SERVERCALL_RETRYLATER`.
  The window stays responsive, but avoid clicking around in Outlook mid-scan.

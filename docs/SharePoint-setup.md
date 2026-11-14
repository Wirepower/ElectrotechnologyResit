# SharePoint setup for Electrotechnology Resit

## Sample data (testing)

On a new install the data source is **Sample data (testing)**. Nothing is read from SharePoint. Bookings, students, teachers, units and a placeholder signature image are stored as JSON under `%LOCALAPPDATA%\ElectrotechnologyResit\SampleData`.

The seed is fictional only: student IDs 9000001–9000014, names that are not real students, and `@example.com` addresses. Unit codes are Frank's resit units from UEE30820 Certificate III in Electrotechnology Electrician: `UEECD0019`, `UEECD0020`, `UEECD0046`, `UEEEL0005`, `UEEEL0012`, `UEEEL0021` and `UEEEL0025`, with the national unit titles. Bookings cover three nights: today, two days ago, and sixteen days ago (so the 14-day delete has something to remove).

**Reset sample data** in Settings (shown only in sample mode) deletes those files and writes the seed again. It does not touch SharePoint.

The main window title and status line say **SAMPLE DATA**. Teacher email still opens Outlook, but only to display the messages. They are not sent, because the addresses are fake.

Switch the data source to **SharePoint**, then Save and Exit. If the site or lists are not set up, the first-run wizard runs. Sign in / Test connection, Run first-time setup, and Import CSV/Excel are on the SharePoint section of Settings. The client ID is a public client. The only token store is the MSAL cache on this PC. No client secret is kept.

Do not put real student names, IDs or emails in the sample files, tests, or this guide.

Owner: Frank Offer, Victoria University, Sunshine.

This desktop app no longer uses SQL Server. Every booking, student lookup, teacher, unit and email-signature image is read and written in SharePoint through Microsoft Graph, as the person who is signed in.

This has **not** been tested against a real Victoria University tenant. The notes below are what Frank needs to try on a site he can write to.

## What the app creates

The first time the app runs, and whenever you open **SharePoint setup wizard** from Settings, it:

1. Asks for the SharePoint site URL. List names are optional and are already filled in.
2. Signs in with a public-client MSAL prompt (Microsoft.Identity.Client). Redirect URI: `http://localhost`.
3. Checks that the signed-in account can open the site.
4. Creates any missing list or column. It does not delete lists, columns or rows, and it is safe to run again.
5. Optionally imports Students, Teachers and Units from CSV or Excel. Import adds or updates rows. It does not remove rows that are absent from the file.

On later startups the app signs in, checks the site, and compares the lists with the schema below. If something is missing it offers to repair it. Repair uses the same non-destructive step.

SharePoint does not need the old VPN prompt. That dialog has been removed.

## App registration

Create an app registration in the Victoria University tenant (or a test tenant).

- Platform: **Mobile and desktop applications**
- Redirect URI: `http://localhost`
- Supported account type: whatever IT will allow for staff. The app sends the directory (tenant) ID from Settings. The default is `common`.
- Public client: allow public client flows.

Delegated Microsoft Graph permissions:

| Permission | When |
| --- | --- |
| `User.Read` | Always. Identifies the signed-in person. |
| `Sites.Manage.All` | Default. Required to create lists and columns (`POST /sites/{id}/lists` and `POST .../columns`). `Sites.ReadWrite.All` cannot do that. |
| `Sites.ReadWrite.All` | Default. Required to read and write booking rows and the signature image. |
| `Sites.Selected` | Optional. Tick **Use Sites.Selected** only after an administrator has granted this app permission to create lists on that one site, not only to edit rows. Without that grant, sign-in can succeed and list creation can still fail. |

The default sign-in requests `Sites.Manage.All`, `Sites.ReadWrite.All` and `User.Read` together. Admin consent is often required. Frank and IT need to choose that pair, or a `Sites.Selected` grant that includes list management. The app does not request application (app-only) permissions.

The application (client) ID and tenant ID are typed into Settings or the wizard. They are not stored in this repository.

Microsoft Bookings is a separate, hidden option. Live Bookings is owned by VU Timetable and needs Frank's sign-off before anyone turns it on. That registration, if it is ever used, needs delegated `Bookings.Read.All` and `User.Read`, with the same `http://localhost` redirect. Do not build a new live Bookings pull in this app.

## Settings that stay on the PC

These are Windows user settings. They are not written to SharePoint.

- SharePoint site URL
- Application (client) ID and directory (tenant) ID
- Whether to request `Sites.Selected` instead of `Sites.Manage.All` and `Sites.ReadWrite.All`
- List names (defaults below)
- Monthly PC/lab password, typed by the operator. There is no default. It is inserted into the Word login sheet only.
- Word template path, and whether to open the document after save
- Bookings client, tenant, business ID, and the checkbox that shows **Pull from Bookings** (off by default)

The monthly password must not be copied into source, docs, or sample files.

## Lists and internal column names

Display names are the defaults. Settings can change the list display names. Column internal names are fixed, because Graph reads and writes those names.

### ResitBookings (generic list)

One row per student ID. A second save updates that row. It is not one row per student per unit per night.

The grid shows these columns, in this order. SharePoint display names match them. There is no TeacherEmailSent column. The allocated teacher email is not edited on its own: a save always copies it from the selected teacher in the Teachers list.

| Internal name | Display name (grid header) | Type | Index | Unique |
| --- | --- | --- | --- | --- |
| StudentID | Student ID | Single line of text | Yes | Yes |
| Unit | Unit | Single line of text | | |
| QualificationCode | QualificationCode | Single line of text | Yes | |
| ResitDate | Resit date | Date only | Yes | |
| Blockgroup | Blockgroup | Single line of text | | |
| AllocatedTeacher | AllocatedTeacher | Single line of text | | |
| Status | status | Single line of text | | |
| AllocatedTeacherEmail | AllocatedTeacher email | Single line of text | | |
| EnergyspaceClassCreated | EnergyspaceClassCreated | Yes/No | | |
| EnergyspaceAssessmentBooked | EnergyspaceAssessmentBooked | Yes/No | | |
| AttemptNo | AttemptNo | Single line of text | | |
| StudentFirstname | Student Firstname | Single line of text | | |
| Surname | Surname | Single line of text | | |
| Email | Email | Single line of text | | |
| UnitName | Unit Name | Single line of text | | |

`ResitDate` is the Australia/Melbourne calendar date. The app writes `yyyy-MM-dd` and reads the date part back without shifting it through UTC.

Student ID and the resit date are the columns used in Graph `$filter`. If SharePoint rejects a filter because a column is not indexed, the app retries that one request with `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`. That header is not sent on the first attempt.

Graph paging (`@odata.nextLink`) and throttling (`429` and `503`, using `Retry-After`) are handled in the shared client.

The SharePoint `Title` column is set to the student ID so the list is readable. The app does not use `Title` as the key.

### Students (generic list, lookup)

| Internal name | Display name | Type | Index | Unique |
| --- | --- | --- | --- | --- |
| StudentID | Student ID | Single line of text | Yes | Yes |
| StudentGivenName | Student Given Name | Single line of text | | |
| StudentFamilyName | Student Family Name | Single line of text | | |
| StudentPersonalEmail | Student Personal Email | Single line of text | | |
| BlockGroupCode | Block Group Code | Single line of text | | |

Search looks up one student with `fields/StudentID eq '...'`. It does not download the whole list. Teacher and unit lists are read once per session and then cached. Student hits are cached for that session too.

### Teachers (generic list, lookup)

| Internal name | Display name | Type | Index |
| --- | --- | --- | --- |
| TeacherFullName | Teacher Full Name | Single line of text | Yes |
| Email | Email | Single line of text | |

### Units (generic list, lookup)

| Internal name | Display name | Type | Index | Unique |
| --- | --- | --- | --- | --- |
| UnitCode | Unit Code | Single line of text | Yes | |
| UnitTitle | Unit Title | Single line of text | | |
| QualificationCode | QualificationCode | Single line of text | Yes | |

Unit Code is indexed and is not unique. The same unit code can belong to a later qualification. A site created by an older build may still have Enforce unique values turned on for Unit Code. Turn that off before adding a second package that reuses a code. The app does not remove that setting.

### Qualifications (generic list)

| Internal name | Display name | Type | Index | Unique |
| --- | --- | --- | --- | --- |
| QualificationCode | QualificationCode | Single line of text | Yes | Yes |
| QualificationTitle | Qualification title | Single line of text | | |
| Version | Version | Single line of text | | |
| Status | Status | Single line of text | Yes | |
| ClassGroupSuffix | ClassGroupSuffix | Single line of text | | |
| AbbreviationRules | AbbreviationRules | Single line of text | | |

Status is Draft, Active or Superseded. Only Active is offered for a new booking. Draft can be edited. Superseded stays readable. Abbreviation rules are one line each, `prefix UEECD=CD` or `override CODE=ABBR`, and must fit in 255 characters. The first-run setup creates this list and, when it is empty, adds UEE30820 as the Active package with its seven units.

`QualificationCode` on a booking is written when the row is created and is not changed by a later package edit. A booking that was saved before this column existed, and so has a blank code, is filled with the default Active qualification (UEE30820 when that package is Active) the next time the bookings are loaded. The filled value is then kept.

### EmailSignature (document library)

Put one current signature image in this library (JPEG, PNG, GIF or BMP). The newest image is embedded in teacher emails. If the library is empty, the email still sends without an image. The message tags the bytes as JPEG, matching the previous app, so a JPEG file is the right choice.

## Energyspace ticks

These are progress checkpoints for a unit on a night, not a per-student mark.

- **Energyspace Class Created** means the Energyspace class was created.
- **Energyspace Assessment Booked** means the assessment was scheduled after students were added.
- The box is ticked when the lists load only if every booking for that unit and that date is true.
- Unticking clears the flag for that unit and that night.
- Export to Excel appears when either Energyspace list has a tick.
- Resit Login Details appears only when at least one unit is ticked under Energyspace Class Created.
- The Word {{CLASS_GROUPS}} block uses only those Class Created ticks. If none are ticked, the sheet says to tick units under Energyspace Class Created.
- Each class-group line is `{abbreviation}-RESIT-{suffix}-{year}`. `RESIT`, the hyphens and the year from the date picker stay fixed. The abbreviation and the suffix (for example CERTIII) come from the qualification record, not from a hard-coded unit list.
- On startup the app deletes bookings whose resit date is 14 or more days before today. There is no button and no prompt. The status line says `Removed N old resits` when N is greater than zero. Only that count is logged.
- Pull from Bookings, if it is ever turned on, updates the booking but leaves both flags as they were.

## Import files

Sample files, with fictional people only, are in `samples/` and are copied next to the program:

- `students.csv` — student IDs 9000001 to 9000004
- `teachers.csv`
- `units.csv`

The wizard accepts `.csv` and `.xlsx`. Headers can be the SharePoint display names or the shorter names in the samples (`Student ID`, `Given Name`, `Family Name`, `Personal Email`, `Block Group Code`, `Full Name`, `Email`, `Unit Code`, `Unit Title`).

Do not put real student data in this repository, in tests, or in logs.

## PowerShell alternative

`tools/Create-SharePointLists.ps1` creates the same lists with PnP PowerShell. It is for IT if the wizard cannot be used. It is also idempotent and does not delete anything. The supported path is the wizard inside the app.

## What Frank needs to test

1. App registration, redirect URI `http://localhost`, and consent for `Sites.Manage.All`, `Sites.ReadWrite.All` and `User.Read` (or a real `Sites.Selected` grant that can create lists, if that option is chosen).
2. First-run wizard against the real site: sign-in, site check, lists and columns created, Student ID unique, Resit date indexed.
3. Run the wizard again and confirm it does not duplicate or delete anything.
4. Add, edit and delete a booking; one student ID stays one row.
5. Tick and untick Energyspace for a unit, close and reopen that date, and confirm the saved state comes back. Confirm a Bookings import does not clear flags (only if that hidden button is deliberately turned on).
6. Export to Excel after a tick on either Energyspace list. The Word login sheet, and its class groups, only after a Class Created tick.
7. Teacher email on Edit saves the address of the teacher who is selected, including when the teacher combo was only loaded and not clicked again.
8. Email mode: exactly one of **Display & Send Manually** and **Send Email - Do Not Display**.
9. Outlook message text, ordinal date (for example `3rd November, 2026`, and `11th`, `12th`, `13th`), sent on behalf of Electrotechnology.admin@vu.edu.au, with the signature image when one is in the library.
10. Excel roster layout, and the Word tokens, including the monthly password typed in Settings.
11. Delete one student, delete a night, and delete records older than 14 Melbourne days.
12. Import the sample CSVs and confirm counts, not a dump of names, in the wizard.
13. A failed sign-in offers Settings and does not pretend the site is available.

## Still for Frank to decide

- Whether IT will approve `Sites.Manage.All` plus `Sites.ReadWrite.All`, or `Sites.Selected` with permission to create lists on the site.
- Whether live Microsoft Bookings should ever be switched on. Until then the button stays hidden.

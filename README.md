# EmailArchiveAutomation

Automates: Fetch -> Filter -> Download -> Decrypt -> Rename -> Save to NAS.

Requirements
- .NET 10 SDK
- Visual Studio 2022+ / Visual Studio Community 2026
- NuGet packages: `MailKit`, `MimeKit`, `SharpCompress`

Quickstart
1. Clone the repository:
   `git clone https://github.com/AhmadAsyraf/EmailArchiveAutomation.git`
2. Open the solution in Visual Studio.
3. Restore NuGet packages.
4. Configure secrets (do NOT keep credentials in source):
   - Set `IMAP` server, `EMAIL_USER`, `EMAIL_PASS`, `NAS_PATH` and `TEMP_FOLDER` using environment variables or a secret manager.
   - Current `Program.cs` uses hard-coded placeholders — replace before running.
5. Build and run the project.

Security
- The sample `Program.cs` contains plaintext credentials and a NAS path. Remove or replace these with environment variables or a secret store before pushing to public repos.

Notes
- The project expects password-protected ZIP attachments and extracts files to the configured NAS.
- Logs are appended to `automation_log.txt` in the configured NAS folder.

License
MIT

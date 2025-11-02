# Internal Tools Portfolio

This repository contains refactored versions of my in-house automation utilities. All tooling has been scrubbed of proprietary servers, credentials, and domain names while keeping the workflows functional. Wherever an environment-specific value is required, the script exposes a configuration block or an obvious placeholder so you can connect it to your own lab, demo tenant, or homelab.

## Toolbelt

| Script | What it Demonstrates |
| --- | --- |
| `AccessManager.ps1` | WPF + PowerShell desktop app that administers LabOps/OpsPortal technician accounts over ODBC with audit logging, guardrails, and transactional updates. |
| `vaxcomAdmin.ps1` | Lightweight single-window variant of the admin console focused on rapid user edits, privilege mirroring, and JSONL auditing. |
| `64to32bitdsn.ps1` | Clones 64-bit system DSNs into their 32-bit equivalents with driver-path rewriting and ShouldProcess support. |
| `dsncheck.ps1` | Menu-driven TLS/Schannel policy helper and ODBC DSN maintainer for legacy SQL clients on Windows 10/11. |
| `odbctest.ps1`, `TestDatabaseQuery.ps1` | Diagnostic probes that validate DSN connectivity, enumerate schema metadata, and exercise parameterised queries. |
| `Set-LegacyCipherSuite.ps1` | Hardens (or intentionally loosens) TLS cipher order for line-of-business SQL clients that require TLS 1.0/1.1. |
| `Add-AdminGroups.ps1` | Bootstrapper that adds central security groups to the local Administrators group with graceful error handling. |
| `liveMigrationTool.ps1`, `RoboCopy.ps1`, `robocopyFrom.ps1` | USMT + Robocopy migration suite with GUI/CLI front ends, telemetry, throttling profiles, and resumable transfers. |
| `usmt.ps1` | Menu-driven USMT capture/restore orchestrator with profile filters, account remapping, and encryption support. |
| `LargeFileTransfer.ps1` | Guided Robocopy wrapper that normalises admin-share paths, performs dry runs, and captures logs. |
| `Collect-BSOD-Diag.ps1` | One-click collector that snaps system info, drivers, services, and installed programs after a crash. |
| `refresh.ps1` | Endpoint imaging post-task that mounts a deployment share, copies provisioning assets, applies cipher fixes, enables BitLocker escrow, and runs gpupdate. |
| `registry-templates/lab-odbc-sample.reg`, `sebia registry` | Sanitised registry templates for ODBC DSNs and serial-instrument software. |
| `samples/bsod.log` | Example output from the diagnostic collector for portfolio demonstrations. |

## Getting Started

1. Clone the repository and open an elevated PowerShell session.
2. Replace the placeholder UNC roots (`\\fileserver01`), DSN names (`LAB_PRIMARY`, `LAB_SECONDARY`), and domain groups (`CONTOSO\\...`) with values from your test environment. Each script keeps its configuration at the top, so a single edit block is usually enough.
3. For GUI experiences (`AccessManager`, `liveMigrationTool`, `RoboCopy`), launch via `powershell.exe -ExecutionPolicy Bypass -File .\\ScriptName.ps1`. Console-only utilities can be dot-sourced or invoked directly.
4. Optional resources:
   - Import the sample registry templates before running DSN-dependent tooling.
   - Review `samples/bsod.log` to understand the data captured by the collector.

## Highlights

- **No dependencies on internal infrastructure** – all secrets, server names, and corporate identifiers were replaced with placeholders.
- **Production-hardened ergonomics** – elevation and STA enforcement, logging, error trapping, and WhatIf support mirror real enterprise practices.
- **Documented flows** – each script prints instructions or exposes menus to guide technicians through migrations, imaging, and database maintenance.

Feel free to fork the repo and adapt the tooling to your own environment. If you have questions about any specific workflow or want to see the original production variants, reach out via the contact details on my résumé.

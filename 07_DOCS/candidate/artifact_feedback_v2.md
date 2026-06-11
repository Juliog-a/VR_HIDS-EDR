# Artifact feedback v2 - P4/P3 candidate iteration

## Executive summary

This candidate iteration addresses the current P4/P3 issues without modifying validated artifacts or original scripts.

The previous campaign is valid as a pipeline smoke test, but not as final detection evidence. The main blockers were:

- TEC-007 did not produce effective encryption because `enumeracion.ps1` and `Scriptransom.ps1` used a different folder than the runner.
- CU-006 did not appear as its own alert.
- CU-008 and CU-009 were emitted from the ransomware ScriptBlock instead of their own behavior.
- CU-001 included router JSONL self-events.
- CU-002 included a Chrome/Adobe Native Messaging false positive.

## Files created

- `artifacts_candidate/Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
- `artifacts_candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v2.yaml`
- `artifacts_candidate/Custom.TFM.HIDS.Router.JSONL.Discord.Final_v3.yaml`
- `scripts_candidate/enumeracion_v2.ps1`
- `scripts_candidate/Scriptransom_v2.ps1`
- `scripts_candidate/exfiltracion_v2.ps1`
- `runner_candidate/TFM_Run_All_TEC_Tests_v2.ps1`
- `docs_candidate/artifact_feedback_v2.md`
- `docs_candidate/PLAN_VALIDACION_P4_P3_v2.md`

## Changes by problem

### TEC-007 not validable

Candidate scripts use `$PSScriptRoot` and explicit parameters so the CSV and password locations are aligned.

The runner counts `.aes` files before restoring the dataset. This gives campaign-level validation that encryption actually happened.

### CU-006 missing

P3 v2 keeps CU-006 as a dedicated contextual detection for `SecurityCenter2`, `AntivirusProduct`, `Get-CimInstance`, `Get-WmiObject` and `wmic`-like discovery.

### CU-008/CU-009 misclassification

P3 v2 excludes crypto ScriptBlocks from CU-008 and CU-009. This preserves semantic priority:

1. CU-007 crypto/encryption.
2. CU-009 staging/archive/network context.
3. CU-008 destructive deletion.
4. CU-006 discovery.
5. CU-001 generic PowerShell.

### CU-001 router self-events

P4 v2 excludes bounded router JSONL self-events by matching `alerts.jsonl`, `FromBase64String`, `Add-Content`, `RouterArtifact` and JSON conversion patterns.

### CU-002 Chrome/Adobe false positive

P4 v2 excludes only the specific Native Messaging pattern:

- parent image is `chrome.exe`;
- command or parent command contains `WCChromeNativeMessagingHost.exe`, `chrome.nativeMessaging`, `Adobe\Acrobat`, `Acrobat\Browser` or `chrome-extension://`.

This does not suppress generic suspicious `cmd.exe`.

## Residual false-positive risk

- CU-001 can still detect administrative PowerShell using suspicious flags.
- CU-002 can still detect legitimate administrative `cmd.exe /c` chains.
- CU-006 can still detect legitimate security inventory activity.
- CU-007 can match legitimate encryption scripts or backup tooling.
- CU-008 can match legitimate bulk cleanup scripts.
- CU-009 can match legitimate archive/copy automation with optional network checks.

## False-negative risk introduced

- CU-008 will not fire for destructive behavior that lacks the selected deletion context terms.
- CU-009 will not fire for staging that only copies or only compresses without both behavior families.
- CU-001 router self-event exclusion could hide a malicious command that exactly mimics the router JSONL writer pattern.
- CU-002 Chrome/Adobe exclusion could hide malicious activity if it abuses the exact Native Messaging pattern.

## Benchmark impact

Expected impact is low. The changes add regex exclusions and conjunctions inside existing event streams. No new heavy collection, no Sysmon ID 26 fan-out, and no P2/P1 activation.

## Router note

`Router.Final_v3` is included because P4/P3 candidate artifact names are versioned. It must not be treated as a detector. Discord remains disabled by default and the JSONL writer uses the validated Base64 `execve` pattern.

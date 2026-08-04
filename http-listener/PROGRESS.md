# Development Progress

Built with the multi-agent loop from `Claude_Code_Development_Instructions.md`: Claude Opus as
orchestrator, three DeepSeek workers, QC gate after every dispatch.

## Roles as actually run

| Agent | Scope | Delivered |
|---|---|---|
| Worker A (DeepSeek) | HTML UI framework | `index.html` — later dropped, see below |
| Worker B (DeepSeek) | Data architecture & validation schemas | `docs/data-contract.md` |
| Worker C (DeepSeek) | PowerShell test suite | `tests/test.ps1` |
| Orchestrator (Opus) | Architecture, integration, QC, arbitration | `listen.ps1`, `launcher.cmd`, all fixes below |

`listen.ps1` was kept as orchestrator work rather than dispatched. The file-locking guarantee is
the entire reason this project exists, and it lives in the accept loop, the handle lifetime and
the retry policy — the parts most likely to come back subtly wrong from a worker and pass a
shallow test anyway.

## Status

| Gate | Result |
|---|---|
| `tests/test.ps1` | **19 / 19 pass**, exit 0 |
| Browser console errors | none |
| External dependencies | none |
| PowerShell | 5.1, verified on 5.1.26100 |
| No file locking | verified by exclusive `[System.IO.File]::Open(..., 'None')` while running |
| Round trip through the real UI | verified: success, 409, 400, batch, lookup |

## Corrections to the instruction documents

**DeepSeek endpoint.** The instructions show `POST https://api.deepseek.com/v1/messages` with
model `deepseek-coder`. Posting there with the OpenAI-compatible shape fails. Two surfaces exist:
`/chat/completions` (OpenAI shape) and `/anthropic/v1/messages` (Anthropic shape). We use the
Anthropic surface with `deepseek-v4-pro[1m]`, matching `Start-DeepSeek.bat`. `deepseek-coder` is
retired and folded into `deepseek-chat`.

**Separate data directory.** The instructions place generated CSVs in the parent tool's
`system/generated/`. That folder already holds an `fg_dimensions.csv` with an unrelated schema
(`fg_id,fg_name,...,conflict_detail`). Writing this listener's schema there would corrupt a file
the main Fit Tool parses, so this project keeps its own `system/generated/`.

**SQLite dropped.** The instructions allow a bundled `sqlite3.exe`; the master prompt asks for a
CSV pipeline. CSV alone satisfies every requirement and removes the only external binary, so
there is no `.exe` in the deliverable.

**Idle timeout, not wall-clock.** The master prompt says "handle requests for 30 seconds". Taken
literally the server dies mid-typing. Implemented as 30 seconds *idle*, with a browser keep-alive
ping. The stated goal — never hold a handle longer than needed — is preserved.

## Iteration log

### Iteration 1 - dispatch

All three workers failed with `content should be a string or a list at line 768 column 18`.

Root cause was ours, not the API's: `Get-Content -Raw` returns a **decorated** string carrying
`PSPath` / `PSParentPath` / `PSDrive` note-properties. `ConvertTo-Json` serialised those, so
`content` went out as a nested object across 768 lines instead of a string. Fixed by casting to
`[string]` and adding `-Encoding UTF8`. Re-dispatched.

### Iteration 2 - QC gates

- **Worker A: FAIL.** Hit the 16k `max_tokens` ceiling; `index.html` truncated mid-file.
  Re-dispatched at 40k. Iteration 2 returned a complete document (85 KB, `<!DOCTYPE>` to
  `</html>`) and passed all 11 constraint checks.
- **Worker B: PASS.** All four deliverables present.
- **Worker C: PASS** on content, three defects on integration (below).

### Iteration 3 - integration defects found and fixed

**In Worker C's test suite:**

1. *Error bodies always empty.* The suite read `$_.Exception.Response.GetResponseStream()`. In
   PowerShell 5.1 `Invoke-WebRequest` has **already drained** that stream into
   `$_.ErrorDetails.Message`, so every 4xx body assertion read `''`. Now prefers `ErrorDetails`
   with the stream as fallback.
2. *Listener never started.* `Start-Process -ArgumentList` joins elements with spaces and does
   not quote them, so `-File C:\...\MasterCarton fit check tool\...` split into separate
   arguments. Paths are now quoted.
3. *`Join-Path` with three arguments* is PowerShell 7+. Nested on 5.1.

**In Worker A's UI — the most serious defect found:**

`setMessage('msgFg', ...)` passed a **bare element id**, but `$` was
`document.querySelector`, which looked for a `<msgFg>` tag, returned `null`, and returned early.
**Every success and error message in the entire application was silently discarded**, along with
all Raw Response panels. Rows were still written, so a shallow test passes and the user simply
never sees confirmation. `$` now resolves a bare identifier as an id before falling back to
`querySelector`.

**In the orchestrator's own `listen.ps1`:**

4. *Batch partial success reported as total failure.* The server set `success = ($failed -eq 0)`,
   so 2-inserted-1-duplicate rendered as "✗ Batch failed (200)" even though the client had a
   correct partial-success branch. The master prompt shows `success: true` alongside a `failed`
   count: success means the batch was processed. Server corrected.
5. *Refresh killed the server.* The UI beaconed `/api/shutdown` on `pagehide`, which also fires
   on F5 — so refreshing the page shut the listener down and the reloaded page could not reach
   it. Added `/api/release` with a 5-second grace window instead.
6. *The fix for 5 did not work at first.* The grace-cancel assignment sat below the `/api/ping`
   and static-file handlers, which `continue` out of the loop body — so a reconnecting page never
   cancelled the countdown. Moved the reset up to where activity is recorded. Verified by
   reloading and confirming the listener is alive 9 seconds later.

### Arbitration - three test failures that were the test's fault

First full run was 16/19. All three were bad assertions, not server bugs:

- **T03** compared the raw header line against unquoted text. The listener writes RFC-4180 quoted
  fields, matching the existing WVN CSVs and what `Export-Csv` produces. The contract is column
  *order*, so the test now compares parsed column names.
- **T08 / T12** matched patterns against `"$($res.Body)"`. Interpolating a `PSCustomObject`
  renders `@{...; details=System.Object[]}` — the nested per-field messages are unreachable, so
  these assertions could never have passed. Now matched against `RawBody`.

Second run: **19/19**.

## Iteration 4 - folded into the real tool, standalone page dropped

The separate data-entry page was the wrong shape: it duplicated a tool that already exists and
would have meant two places to look. The feature moved into
`system/WVN_Master_Carton_Fit_Tool.html` instead, and `index.html` / `launcher.cmd` were deleted.
Worker A's page is still in `tools/worker-output/` as the record of that iteration.

What changed:

- **New Product tab** gained a *Save this product* button that reuses the fields already on the
  form, so nothing new has to be typed.
- **Add Data tab** added for bulk entry: a downloadable blank sheet, drag-and-drop or paste, a
  preview that flags each row as ready / will-replace / has-a-problem, then Save all.
- **Header** gained a small "Saving is on · Turn off" pill that ends the background program.
- Every one of those controls is hidden unless the helper actually answers, so the tool is
  unchanged for anyone without it.
- All wording is plain language in English and Vietnamese: no ports, no servers, no CSV, no HTTP.

**Duplicates now overwrite instead of erroring.** Re-measuring a product is a correction, so the
newest entry wins and the row count does not grow. The 409 path was removed and T05 rewritten to
assert the replacement.

**Saved entries feed back into the tool.** They are merged onto the raw workbook rows inside
`loadFromCSVs`, before any mapping, so no analysis code changed. They are stored outside
`generated/` because the launcher rebuilds that folder from the workbooks on every run and would
erase them. The merge is wrapped so a failure can never block startup.

Three integration bugs found and fixed while testing this:

7. *Referencing `toolUi` too early.* The wiring block sat above `const toolUi = new UiController()`.
   `const` is in the temporal dead zone until its declaration, so even `typeof toolUi` throws
   rather than reading as undefined. Deferred it to run after the controller exists.
8. *The Add Data tab opened to an empty page.* `activateWorkflow` force-hides any panel not in a
   hard-coded `primary` set, so the new panel was hidden the moment its tab was clicked.
9. *Batch result read "Saved. 3 · 1 replaced".* Reworded to "3 rows saved. 1 of them replaced an
   item that was already on the list."

Verified in the real tool against real data: 968 products loaded, three saved entries merged to
971, a saved product is findable in the normal picker with its saved dimensions, re-saving the
same ID replaces it rather than duplicating, bad rows are skipped with a plain reason, and the
Turn off button ends the PowerShell process.

## Known cosmetic gap

Client-side range validation says `Must be 10 – 1000 mm.` (en dash, trailing period); the server
says `Must be 10-1000 mm`. Both are clear and the client message only appears before a request is
sent. Left as is rather than spending an iteration on it.

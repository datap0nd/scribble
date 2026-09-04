# Release notes

## User-supplied domain navigation (extension 1.5.2)

When the user types a bare domain such as `samsungtradein.ae`, Scribble now
explicitly treats it as a valid direct-navigation target and opens it with
HTTPS. The model-facing prompt and tool schema no longer misleadingly require
the user to repeat the same site with a scheme and full path.

## Google result ref hotfix (extension 1.5.1)

Scribble now resolves the exact ref returned by Google before taking any new
full snapshot. Previously, `browser_act` performed a full rescan first; that
created a new document revision and made the requested result ref stale before
the click could occur. The action now uses the already returned snapshot, or a
non-mutating resolve-and-probe fallback, and retains the observed-link fallback
for a verified no-effect click.

## Verifiable browser results (extension 1.5.0)

Scribble now completes price and valuation research only after the extension
validates a structured evidence record against the final live DOM, current work
tab, snapshot revision, and verified action receipts. The result card keeps the
amount, configuration, caveat, source, and observation time together, and
**Open evidence tab** returns to the same surviving work tab. Stale, cross-tab,
invented, Google-snippet-only, or answer-mismatched evidence is rejected.

All Scribble-authored browser progress, errors, clarifications, and completion
copy now use first person. Page quotations, protocol markers, and error codes
remain unchanged where they are data rather than Scribble's voice.

## Calibrated browser policy (extension 1.4.0)

The browser policy now judges the requested operation instead of blocking every
control near a field with words such as “mobile.” Reversible product cards and
HTTP(S) product links—including a link labelled Buy—can be opened. Sensitive
typing, password forms, uploads, payment or identity fields, purchase buttons,
and final purchase/trade-in submission remain hard-denied. The JavaScript and
native policy layers have matching tests for this boundary.

Browser clarification can group one to three related questions, allowing a
trade-in journey to collect market, storage, and condition together. Office
panes retain the existing single-question schema, and both legacy and grouped
answers remain parseable during extension/native-host version skew.

## Reliable browser operation (extension 1.3.0)

Localized Google hosts and `name=q` search fields no longer depend on English
labels. Snapshot labels now recover meaningful visible card, option, group, and
associated-label text when ARIA metadata is blank, generic, duplicated, or
contains `undefined`. Snapshot/native exchange capacity is paired at 24 KiB,
with whole control records prioritized ahead of less useful page text.

The fixed post-click delay is replaced by bounded adaptive stabilization.
Mutations invalidate old refs immediately, one mutating browser tool is allowed
per model round, and every action returns a refreshed snapshot with a verified
`changed`, `no_effect`, `stale_ref`, `incomplete`, or `blocked` outcome. A
no-effect observed anchor may be opened only through its exact HTTPS snapshot
target; general direct navigation remains limited to user-supplied URLs.

## Browser operator (extension 1.2.2)

Scribble for Chrome can now search through Google's visible UI and inspect,
type, select, click, check, press, hover, scroll, and wait in its own inactive
work tabs. Results are tied to observed page refs; direct navigation accepts
only URLs supplied by the user.

Google results-page clicks now bring offscreen targets into view, wait for a
real navigation, retry once through trusted keyboard input, and report failure
instead of claiming success when the result did not open. Browser progress and
assistant copy now speak consistently in Scribble's first person.

This release adds Chrome's required `debugger` permission for trusted input.
Chrome may disable an existing unpacked installation until the user reviews and
accepts the permission increase. Chrome's debugging banner appears during an
action and may flicker because Scribble detaches after every atomic action.
Canceling that banner stops the run cleanly.

Typing is capped at 200 characters and may originate in the user request, a
locally validated public alias, or a clarification answer. An unfamiliar
inferred public term now prompts for exact confirmation and retries instead of
ending the task. Typed text is shown verbatim once beside the Pixel Pal before it
is sent to a page. Browser activity is now translated into plain language; raw
control refs no longer appear as duplicate chat cards. A native policy blocks credential, personal/traveler identity,
booking, payment, messaging, upload/download, and destructive fields, forms,
and controls. CAPTCHA, bot checks, protected pages, inaccessible cross-origin
widgets, and sign-in walls are not bypassed.

The normal browser stopping rule is now progress-based: 20 consecutive calls
without an observed page-state change stop as a loop, while a longer task may
continue as long as it progresses. A 120-round emergency cost/safety fuse
remains. Travel flows verify displayed airports and dates before searching;
Dubai resolves to DXB rather than nearby Sharjah unless the user asks for nearby
airports. Replay keeps clarification answers and six recent snapshots in full
and compacts older browser results.

The browser panel now compares its running extension version with the extension
bundled by installed Scribble. When Chrome is still using stale unpacked code,
the footer shows the available version and offers a one-click reload.

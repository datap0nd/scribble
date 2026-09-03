# Release notes

## Browser operator (extension 1.2.1)

Scribble for Chrome can now search through Google's visible UI and inspect,
type, select, click, check, press, hover, scroll, and wait in its own inactive
work tabs. Results are tied to observed page refs; direct navigation accepts
only URLs supplied by the user.

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

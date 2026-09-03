# Release notes

## Browser operator (extension 1.2.0)

Scribble for Chrome can now search through Google's visible UI and inspect,
type, select, click, check, press, hover, scroll, and wait in its own inactive
work tabs. Results are tied to observed page refs; direct navigation accepts
only URLs supplied by the user.

This release adds Chrome's required `debugger` permission for trusted input.
Chrome may disable an existing unpacked installation until the user reviews and
accepts the permission increase. Chrome's debugging banner appears during an
action and may flicker because Scribble detaches after every atomic action.
Canceling that banner stops the run cleanly.

Typing is capped at 200 characters, must originate in the user request or a
clarification answer, and is shown verbatim in the activity log before it is
sent to a page. A native policy blocks credential, personal/traveler identity,
booking, payment, messaging, upload/download, and destructive fields, forms,
and controls. CAPTCHA, bot checks, protected pages, inaccessible cross-origin
widgets, and sign-in walls are not bypassed.

The browser budget is now 24 chargeable action rounds, up to 12 additional
scroll/wait-only rounds, no more than four consecutive support-only rounds, and
36 total. Replay keeps clarification answers and six recent snapshots in full
and compacts older browser results.

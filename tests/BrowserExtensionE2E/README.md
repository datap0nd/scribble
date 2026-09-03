# Browser extension Playwright tests

This optional, dev-only suite exercises the production page-reader function
against local fixtures for controlled inputs, autocomplete-sensitive fields,
selects, checkboxes, Shadow DOM, same-origin frames, stale refs, hidden and
disabled controls, popup creation, and prompt-injection text.

Run from this directory with `npm install` followed by `npm test`. The Windows
MSBuild gate does not install Node packages or run this suite. Chrome-extension
loading and debugger-banner cancellation remain manual acceptance checks because
browser-level Chrome UI is outside page automation.

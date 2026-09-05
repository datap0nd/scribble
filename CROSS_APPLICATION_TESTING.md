# Cross-application handoffs

Outlook, Excel, PowerPoint, Word, and Chrome have twenty directed handoffs.
Office destinations open a visible draft. Outlook drafts remain unsent, and
document drafts remain unsaved. Office-to-Chrome opens the exact HTTP/HTTPS
URL supplied by the user in a new Chrome window; it does not transfer chat
state or upload the current document.

Chrome exposes `open_outlook_draft`, `open_excel_table`, `send_to_powerpoint`,
and `send_to_word`. Office panes expose the sibling `send_to_*` tools,
`create_email_draft` where appropriate, and `open_in_chrome`.

## Release checks

- The twenty-direction matrix invokes the production draft host and writers
  against recording application doubles. It checks the requested application,
  exactly one new destination, source content, authorization refusal, and
  exclusive-call refusal. Existing document access and document save/send/close
  calls fail the fixture. PowerPoint runs source review, rendering/export calls,
  ownership checks, and rendered review before returning success.
- Browser round trips use real HTTP responses through `BrowserChatService`
  and each of its four Office handlers. They assert that application
  activation occurs on a pumped STA, including after asynchronous review.
- Slide argument and empty endpoint response tests cover the model transport
  boundary independently. Browser page fixtures and installer/COM registration
  checks remain separate required gates.
- CI publishes `CrossApplicationResults.json` as the `CrossApplicationResults`
  artifact. It identifies its verification as simulated application APIs.

These checks do not certify native Office layout or a specific model endpoint.
Native PowerPoint activation on the development machine failed with
`0x80080005` during this change; native Outlook was not registered. A native
acceptance run still needs a working Office installation and configured model.
Never label the simulated matrix as successful native Office testing.

## Native acceptance scenario

Start in Outlook with a selected source email and ask to create a presentation.
Verify that PowerPoint starts, a new visible deck opens, source content survives,
review completes, and additional batches continue that same deck. Repeat from
each source app for each sibling destination, with the destination both stopped
and already running with an unrelated document open. Verify that the unrelated
document remains unchanged. For Chrome destinations, supply an explicit URL.
Repeat cancellation and unavailable-application cases and verify that no
completion is reported without a destination receipt.

# Reading or saving attachments

The `get-attachment` AI tool supports Outlook mail and Teams file attachments.
Use `get-mail` to find an Outlook message's attachment IDs and names.

By default, `get-attachment` returns the bytes as model-visible image or file
content. To save the original file for local processing instead, supply
`savePath`:

```json
{
  "source": "mail",
  "messageId": "<message-id>",
  "attachmentId": "<attachment-id>",
  "savePath": "~/Downloads/mail-attachments/report.xlsx"
}
```

`savePath` must be a full file path, either absolute or starting with `~/`.
Parent directories are created as needed. Existing files are never overwritten;
choose a different path if one already exists. Files are saved on the machine
running the plugin.

A successful save returns `success`, `source`, `messageId`, `name`, `mediaType`,
the exact `bytes` count, and an absolute `path` (`chatId` is also returned for
Teams). It does **not** include `_modelContent` or base64 file contents. The AI
can use local file tools to inspect the saved file, including when Kai's media
or context limits prevent direct delivery. A save failure returns an error;
it does not fall back to sending file contents to the model.

For Teams, use `source: "teams"`, include `chatId` and `messageId`, and select the
file by `name` or zero-based `index`. The same `savePath` option applies.

Omit `savePath` whenever direct model access to the attachment is preferred.

# Sending attachments

Both `send-mail` and `reply-to-mail` accept `filePaths`, an array of absolute or
`~/` paths on the machine running the plugin. The plugin reads the files and
uploads them directly to Microsoft; file bytes are not embedded in the AI's
tool call or returned to the model. Names come from the file names and MIME
types are detected from the contents or extensions.

Example `send-mail` input:

```json
{
  "to": ["recipient@example.com"],
  "subject": "Runner inventory",
  "body": "Attached are the inventory and summary.",
  "filePaths": ["~/Downloads/inventory.csv", "~/Downloads/summary.md"]
}
```

Example `reply-to-mail` input:

```json
{
  "messageId": "<original-message-id>",
  "mode": "replyAll",
  "body": "Attached is the completed inventory.",
  "filePaths": ["~/Downloads/inventory.csv"]
}
```

Attachments work with `reply`, `replyAll`, and `forward`. Forwarding also requires
`to`. Replies preserve the recipients selected by Outlook. Both tools also accept
`attachments` entries with `name`, optional `contentType`, and base64
`contentBytes`; these can be combined with `filePaths`.

When attachments are present, the plugin creates a draft, adds each attachment,
and sends the draft only after all uploads succeed. Files under 3 MiB use the
ordinary attachment endpoint; larger files (including a 5.6 MB CSV) use
[Microsoft's upload sessions](https://learn.microsoft.com/en-us/graph/outlook-large-attachments)
with sequential chunks smaller than 4 MiB. The plugin limits each file to 150 MiB;
the mailbox's total message-size policy can impose a lower limit. This draft
flow requires `Mail.ReadWrite` as well as `Mail.Send`.

If an upload fails, the error identifies the retained draft and the plugin does
not send it. An uncertain final send is not automatically repeated; check Sent
Items and the draft before retrying to avoid duplicate emails. Messages without
attachments continue to use the direct send/reply endpoints.

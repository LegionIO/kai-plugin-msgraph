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

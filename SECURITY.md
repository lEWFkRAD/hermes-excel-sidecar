# Security Policy

Use GitHub private vulnerability reporting. Never publish workbook data, bridge tokens, Hermes keys, attachments, client information, or exploit details.

The supported bridge is localhost-only, token-authenticated, checks Host and Origin, and accepts structured workbook actions only. Never expose it publicly or enable file, terminal, or code-execution toolsets for its Hermes `api_server` channel. Rotate compromised bridge credentials through reinstall/configuration and restart the tracked supervisor; rotate Hermes credentials in Hermes configuration and restart both services.

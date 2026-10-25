# Password management

Reset a user's password — to a value you choose or one the server generates — and read the user's password methods.

## 📖 Overview

- Set a new password and, by default, make the user change it at next sign-in.
- Ask the server to generate the password instead: 8 to 256 characters, 12 by default, always with a digit, an upper-case and a lower-case letter and a symbol. It comes back in the result so you can hand it over.
- The password never appears in an error message: it is replaced by `***`.
- List a user's password methods, or read one.

Called by an AI client on behalf of an Entra ID admin. The reset is marked destructive, so a client can ask before it runs; the two reads are read-only.

## 🏢 Business domain

Credential reset. Which methods count as MFA belongs to [MFA](../mfa/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Password profile** | The password and its change-at-next-sign-in flag on the user | `PasswordProfile` in `password_auth.reset_user_password_direct` |
| **Generated password** | A random password from the operating system's secure source | `generate_secure_password` in `utils/password_generator.py` |
| **Password method** | The user's registered password authentication method | `/users/{id}/authentication/passwordMethods` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| A reset needs a password or `generate_password: true` | the check in the `reset_user_password_direct` tool | Error result `Password must be provided or generate_password must be set to True` |
| A generated password is 8 to 256 characters | `MIN_PASSWORD_LENGTH`, `MAX_PASSWORD_LENGTH` in `password_generator` | Error result `The password length must be 8 to 256 characters, got <n>` |
| A generated password has a digit, an upper- and a lower-case letter and a symbol | `generate_secure_password` | — |
| No error text carries the password | the masking `except` in the `reset_user_password_direct` tool | The error with `***` in place of the password |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "reset_user_password_direct", "arguments": {"user_id": "aoife@contoso.com", "generate_password": true, "password_length": 16}}}
```

The result's `structuredContent` (the generated password is shown here as `***`; the real result carries it):

```json
{
  "status": "success",
  "userId": "aoife@contoso.com",
  "passwordResetRequired": true,
  "message": "Password has been reset using the direct method.",
  "generated_password": "***"
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls reset_user_password_direct; the tool generates a password when asked or refuses when none is given, then PATCHes the user's passwordProfile in Graph and returns the result with any generated password; a Graph refusal comes back with the password masked.](diagrams/password-reset.svg)

One `PATCH /users/{id}` does the reset; nothing is stored by the server. With `generate_password`, a given `password` is ignored. The generated password goes back to the AI client in the result — treat that conversation as holding a secret. When Graph refuses (for example a password its policy rejects), the error text is rebuilt with the password masked and without the original exception chained, so no log line or result carries it.

## 📚 Pages

- [Endpoints](endpoints.md) — the 3 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [MFA](../mfa/index.md) — reach for it to check the user's other authentication methods.
- [Users](../users/index.md) — reach for it to find the user id.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Which Graph permission does the reset need: `User-PasswordProfile.ReadWrite.All` (the old README) or `User.ReadWrite.All` (the permissions helper's suggestion)? | Admins grant the wrong permission and the reset is denied, or grant more than needed. | The code sends `PATCH /users/{id}` with `passwordProfile`; `COMMON_PERMISSION_MAPPINGS` lists `User.ReadWrite.All` and `Directory.ReadWrite.All`; the old README listed `User-PasswordProfile.ReadWrite.All`. | Maintainer (baoduy), against Microsoft's Graph permission reference |

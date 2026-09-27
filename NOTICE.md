# NOTICE

## License

This project is released under the **MIT License**. See [LICENSE](LICENSE) for
the full text.

## Commercial use

The MIT License grants anyone the right to use, copy, modify, merge, publish,
distribute, sublicense and sell copies of this software, including for
commercial purposes. **No permission is required, and none is withheld.**

This NOTICE cannot add conditions to the MIT License, and does not attempt to.
If you are reading this hoping it sets rules, it does not: the terms in
[LICENSE](LICENSE) are the terms.

**A courtesy request, not a restriction:** if you use this commercially, or build
on it in a way you make money from, please **let the author know** — an issue or
a note is welcome. This is a request out of interest in the project, not a
condition of use. Nobody can enforce it, and no licence condition depends on it.

## Attribution

The copyright notice and the MIT permission notice must be retained in all
copies or substantial portions of the Software. Keeping the author's name in the
files is the one real obligation MIT does impose, and the reason the author
field is populated in `package.json` for every package in this organisation.

## Paid features

If a paid or hosted version of this tool ever appears, it will be paid for as a
**service** — hosting, support, or convenience — never as a licence condition.
The code in this repository stays MIT for everyone, permanently. A licence
cannot be both permissively open and conditional, so the open-source grant is
not something that will ever be withdrawn or moved behind a paywall.

## Origin

Extracted from [MSOneNote Exporter](https://github.com/enoola/Microsoft-OneNote-Exporter).
Sibling packages: [microsoft-onenote-list-notebooks](https://github.com/Ms-OneNote-Exporter/microsoft-onenote-list-notebooks),
[microsoft-onenote-export-notebook](https://github.com/Ms-OneNote-Exporter/microsoft-onenote-export-notebook),
[microsoft-outlook-list-emails](https://github.com/Ms-OneNote-Exporter/microsoft-outlook-list-emails).

## Why this project exists

Microsoft does not provide a convenient way to export a whole OneNote notebook,
and the Graph API path is not a usable substitute: it caps page retrieval and
requires Entra admin rights. This tool produces the authenticated session state
those other tools consume, using Playwright instead of the Graph API.

It deliberately does **not** rely on Graph, and it deliberately does not ask you
to hand it your password in a hosted web form. Authentication state is written
to a local file (`auth-file.json`, mode `600`) that you own, can inspect, and can
delete at any time.

## Handling of credentials

`auth-file.json` contains a Playwright `storageState` — live session cookies.
Anyone who reads that file can act as you until the session expires. Treat it
like a password file: do not commit it, do not upload it to a service you do not
control, and delete it when you are done. The default path is under
`~/.microsoft-webauth/`, outside any repository.

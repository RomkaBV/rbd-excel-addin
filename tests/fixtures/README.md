# Parser fixture

`test2.msg.base64` is the base64 representation of the upstream `test/test2.msg` fixture from https://github.com/HiraokaHyperTools/msgreader (Apache-2.0). It contains test mail and an A.txt attachment, never production correspondence. Decode with `Buffer.from(text, 'base64')` in tests only. The project license is included in `vendor/THIRD_PARTY_LICENSES.md`.

All EML messages and Excel request data in the browser tests are synthetic test inputs. Tests inject them into an isolated Playwright browser; the application does not load them.

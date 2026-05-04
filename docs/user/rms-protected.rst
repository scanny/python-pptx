.. _rms-protected:

RMS / AIP-protected files
=========================

Microsoft **Rights Management Services** (marketed under several names —
*Azure Information Protection*, *Microsoft Purview Information Protection*,
*Information Rights Management*, *Sensitivity labels with encryption*) wraps
an ordinary OOXML ``.pptx`` in a second layer of encryption that is tied to
the user's Azure AD / Microsoft 365 identity rather than to a password.

python-pptx **cannot open or produce RMS-protected files.** Attempting to
open one raises :class:`pptx.exc.RmsProtectedPackageError` (a subclass of
:class:`pptx.exc.EncryptedPackageError`) with a message pointing at this page.

This limitation is intrinsic, not an oversight. A proper implementation
requires:

* The **Microsoft Information Protection (MIP) SDK** — which is
  C#/.NET-only (with limited C++ and Java wrappers). There is no official
  Python binding.
* An **authenticated Azure AD session** — MIP cannot decrypt an
  RMS-protected file without contacting the RMS server and presenting a
  valid user or service-principal token.

Detecting an RMS-protected file
-------------------------------

An RMS-protected ``.pptx`` is an OLE2 Compound File Binary Format (CFBF)
container whose first eight bytes are ``D0 CF 11 E0 A1 B1 1A E1`` rather
than the plain-zip ``PK\x03\x04`` signature. Inside the CFBF it has a
``DRMEncryptedTransform`` DataSpaces descriptor and a ``DRMContent`` stream,
instead of the ``EncryptionInfo`` + ``EncryptedPackage`` pair used by
ECMA-376 Agile Encryption (password-protected files).

python-pptx tells these apart automatically::

    from pptx import Presentation
    from pptx.exc import EncryptedPackageError, RmsProtectedPackageError

    try:
        prs = Presentation("protected.pptx", password="hunter2")
    except RmsProtectedPackageError:
        # RMS / AIP-protected — cannot be opened with a password
        ...
    except EncryptedPackageError:
        # ECMA-376 password-encrypted — try a different password
        ...

The ``RmsProtectedPackageError`` inherits from ``EncryptedPackageError``, so
callers that only care "this file is encrypted" can keep their existing
``except EncryptedPackageError`` handler and still catch both kinds.

Workarounds
-----------

If you must process RMS-protected decks from Python, the decryption has to
happen *outside* python-pptx. Each workaround below produces a plain
``.pptx`` that python-pptx can then read normally.

1. **Delegate to Microsoft Office automation.** Drive PowerPoint (Windows
   or macOS) or the Office COM automation layer to open the protected
   deck — Office will resolve the RMS credentials transparently because
   the user is signed in — then *Save As* an unprotected copy. Works for
   interactive workflows and small batches.

2. **Call the MIP SDK from a sidecar service.** Write a small C# /
   .NET service using the `Microsoft Information Protection SDK
   <https://learn.microsoft.com/information-protection/develop/>`_ that
   loads an RMS-protected file, verifies the caller's Azure AD token,
   and emits a plain ``.pptx``. python-pptx picks it up from there.
   This is the Microsoft-supported path for server-side RMS decryption.

3. **Use the AIP Unified Labeling client.** On a Windows box where the
   `Microsoft Purview Information Protection client
   <https://learn.microsoft.com/purview/deploy-scc-ms-info-protect>`_
   is installed and logged in, a PowerShell ``Unprotect-RMSFile``
   cmdlet can strip RMS from a file batch. Handy for one-off
   remediation on a workstation.

4. **Ask the owner to re-publish without RMS.** If the protection was
   applied by a sensitivity label, removing or downgrading the label
   (in Word/PowerPoint/Excel, or via the `Set-AIPFileLabel` cmdlet)
   removes the RMS wrapper entirely.

Password-protected (ECMA-376 Agile Encryption) files are a different
mechanism and *are* supported — see the note about ``password=`` in
:doc:`presentations` and the ``msoffcrypto-tool`` optional dependency.

See also
--------

* :class:`pptx.exc.RmsProtectedPackageError`
* :class:`pptx.exc.EncryptedPackageError`
* The `[MS-OFFCRYPTO] §2.2.10 DRMEncryptedTransform
  <https://learn.microsoft.com/openspecs/office_file_formats/ms-offcrypto/>`_
  specification for the on-disk RMS wrapper format.

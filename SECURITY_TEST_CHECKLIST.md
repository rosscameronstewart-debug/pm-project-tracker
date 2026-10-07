# Local permission test checklist

Use http://127.0.0.1:8765/. All changes in this checklist are local.
Use accounts and records created for testing. Prefix sample records with SECURITY TEST.
Use separate browser profiles or private windows for different accounts; tabs share sign-in cookies.
Record the account, action, expected result, actual result, and any error for each test.

| Test | Steps | Expected result |
| --- | --- | --- |
| Admin normal use | Sign in as Admin. Open Projects, Setup, Vendor Invoices, Customer Billing, NTE, PO Review, Quoting, Texas Ops, and Admin. Create/edit a sample project and open existing authorized PDFs. | Allowed pages load normally; sample changes save; authorized documents open. |
| Read Only | Sign in with a Read Only test account. Open allowed pages and documents. Look for save, upload, and delete controls. | Viewing works. Changes are unavailable or rejected. Direct write-request rejections are covered by automated tests. |
| Field PO ownership | Use two Field PO test accounts. Create one sample PO for each, with different attachment files. Copy each PO's printable-page and attachment addresses. | Each account sees its own POs and can open its own attachment. |
| Direct PO bypass | While signed in as the first Field PO account, paste the second account's copied PO and attachment links into the address bar. | Both return Not found. The second account's data is not shown. |
| Office invoice privacy | As Admin, copy a vendor invoice download link from a PO. Paste it while signed in as Field PO. | Invoice download is blocked, including on the employee's own PO. Own printable PO and list do not expose office invoice details. |
| PDF preview bypass | Copy a restricted PDF's preview link as Admin. Open it as Field PO. Also try its page-image URL: `/pdf-page/FILENAME/0.png`, preserving the filename's URL encoding. | Download, preview, and page-image routes all deny access. |
| Texas-only account | Sign in with a TX/Read Only test account. Open an authorized financial report. Try `/api/projects`, `/api/customer-invoices`, and a copied PO invoice link. | Texas summary and financial report work. Project/customer APIs and unrelated documents are blocked. |
| Feature disabled | In local Admin role permissions, temporarily turn off View and Edit for Bid Tracking on the User role. Sign in as a User test account and open `/api/bids` and `/api/bid-summary` directly. Restore settings afterward. | Both return permission denied. Home alerts do not reveal bid counts or values. Other permitted pages still load. |
| Edit disabled | Leave View enabled but turn off Edit for a test feature on the User role. Try editing a sample record in that feature. Restore settings afterward. | Read access remains. Saves are rejected, even if the old button is still visible in an already-open page. |
| Revoked access | Leave the User test account signed in. In another browser profile, remove its feature access. Retry the direct feature URL without signing out. | The next request is denied. No new sign-in is necessary for revocation to take effect. |
| Disabled account | Disable a signed-in test account in local Admin. Refresh its page or request `/api/me`. Re-enable it afterward. | The session can no longer retrieve protected data. |
| Signed out | Sign out. Reopen copied document links and `/api/bids`. | Documents redirect to sign-in. The API reports Login required. No protected data appears. |
| Calendar regression | Sign in as rstewart@twinpeakselectrical.net. Add/edit a sample recurring calendar event. Try the calendar API as another Admin. | Popup and recurrence continue working for Ross. The other Admin is blocked. |

## Automated checks

Run `python -m unittest test_access_control test_company_calendar test_session_security -v` from the project folder.
The suites use temporary databases and files, without modifying the local business records.
They check route policy coverage, direct write denials, cross-owner PO uploads, mixed-source import deletion,
document access, feature revocation, disabled accounts, filtered alerts, and calendar regressions.

## Access rules preserved and clarified

- Project-feature access remains company-wide. This change does not introduce project-by-project assignments.
- PO request access covers the requester's own POs, attachments, and pickup documents.
- PO review access covers office PO documents and invoices.
- Other documents require permission for the feature that owns their stored database reference.
- Unreferenced files are denied, including for Admin. A missing ownership reference must be repaired rather than bypassed.
- Permission-denied API responses use 403. Unauthorized document/PO links use 404 to avoid revealing existence.

## Password and session checks

1. Reset a local test account's password as Admin while that account is signed in in another browser profile. Its next protected request should require login.
2. Sign in with the reset password. The password-change prompt should appear. Direct API/document requests should remain blocked until a new password is saved.
3. Change the password successfully. The dashboard should reload and work normally. Older sessions for the account should no longer work.
4. Confirm the existing `http://127.0.0.1:8765` login still works. The live Tailscale address/configuration has not changed.
5. Absolute expiry and future HTTPS/HTTP cookie isolation are verified by automated tests; no public HTTPS address has been configured yet.

These checks do not yet add Microsoft sign-in, public HTTPS deployment,
login throttling, CSRF protection, or upload size limits. Nothing here publishes the application.

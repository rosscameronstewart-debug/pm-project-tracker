# Password and session transition

The current HTTP address and `/api/login` are preserved. The live server is unchanged.
Local password changes are enforced before protected data or documents are available.
Administrator resets revoke every session for the account. A user's successful password change
revokes old sessions and issues a fresh one for the browser completing the change.
Sessions retain the 30-minute idle timeout and now have a 12-hour absolute lifetime.

## Future HTTPS access alongside Tailscale

One app instance can accept both private Tailscale HTTP and a trusted HTTPS reverse proxy.
Both routes use the same account database and permissions, but separate browser cookies
and session transport records:

- Existing HTTP: `pm_session`, preserving current field access.
- Configured HTTPS: `__Host-pm_session`, with Secure, HttpOnly, SameSite=Lax, Path=/ and no Domain.
- An HTTPS-issued token is never accepted as an HTTP session, even if supplied manually.
- An HTTP-issued session cannot authenticate on the configured HTTPS route.

When a real HTTPS hostname and proxy exist, set these on the app service:

```
PM_TRACKER_HTTPS_HOSTS=dashboard.your-company-domain.example
PM_TRACKER_TRUSTED_PROXY_IPS=127.0.0.1
```

The hostname above is an example, not an address created by this change.
Use the actual proxy peer address rather than accepting arbitrary proxy addresses.
The proxy must preserve the configured Host and overwrite X-Forwarded-Proto with `https`
only for requests originally received through HTTPS. It must enforce encrypted browser access.
Requests for a configured HTTPS hostname are rejected unless they arrive from the configured
proxy with the correct HTTPS marker. Forwarded headers from untrusted peers are ignored.

Keep the raw HTTP app port limited to Tailscale and the trusted proxy using firewall rules.
Protect the entire new hostname with the intended Entra/Cloudflare access policy before publishing.
These code settings do not install certificates, create DNS records, configure Cloudflare,
or add Microsoft single sign-on. That migration is a separate deployment.

While HTTP remains enabled, its legacy cookies remain non-Secure for compatibility.
This does not make the old route equivalent to browser HTTPS; the old route must remain private.

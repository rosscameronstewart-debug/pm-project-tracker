"""Transition-safe cookie settings: legacy Tailscale HTTP and trusted HTTPS proxy."""
import os
from urllib.parse import urlparse

ABSOLUTE_HOURS = 12
HTTPS_HOSTS = {h.strip().lower() for h in os.environ.get('PM_TRACKER_HTTPS_HOSTS','').split(',') if h.strip()}
TRUSTED_PROXY_IPS = {h.strip() for h in os.environ.get('PM_TRACKER_TRUSTED_PROXY_IPS','').split(',') if h.strip()}

def https_host(handler):
    return (urlparse('//'+handler.headers.get('Host','')).hostname or '').lower() in HTTPS_HOSTS

def secure_request(handler):
    return (https_host(handler) and getattr(handler,'client_address',('',))[0] in TRUSTED_PROXY_IPS
            and handler.headers.get('X-Forwarded-Proto','').strip().lower() == 'https')

def cookie_name(handler):
    return '__Host-pm_session' if secure_request(handler) else 'pm_session'

def cookie_header(handler, token='', clear=False):
    header = f'{cookie_name(handler)}={token}; Path=/; HttpOnly; SameSite=Lax'
    if secure_request(handler):
        header += '; Secure'
    if clear:
        header += '; Max-Age=0'
    return header

def initialize(con):
    columns = {r[1] for r in con.execute('PRAGMA table_info(user_sessions)')}
    if 'secure_session' not in columns:
        con.execute('ALTER TABLE user_sessions ADD COLUMN secure_session INTEGER NOT NULL DEFAULT 0')

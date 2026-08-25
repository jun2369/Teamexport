import os

import msal

import config

os.makedirs(config.SESSIONS_DIR, exist_ok=True)


def _session_path(sid):
    return os.path.join(config.SESSIONS_DIR, f"{sid}.json")


def _build_msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        config.CLIENT_ID,
        authority=config.AUTHORITY,
        client_credential=config.CLIENT_SECRET,
        token_cache=cache,
    )


def load_cache(sid):
    cache = msal.SerializableTokenCache()
    path = _session_path(sid)
    if os.path.exists(path):
        cache.deserialize(open(path, "r").read())
    return cache


def save_cache(sid, cache):
    if cache.has_state_changed:
        with open(_session_path(sid), "w") as f:
            f.write(cache.serialize())


def build_auth_code_flow():
    return _build_msal_app().initiate_auth_code_flow(
        config.SCOPES, redirect_uri=config.REDIRECT_URI
    )


def acquire_token_by_flow(sid, flow, auth_response):
    cache = load_cache(sid)
    app = _build_msal_app(cache)
    result = app.acquire_token_by_auth_code_flow(flow, auth_response)
    save_cache(sid, cache)
    return result


def get_access_token(sid):
    cache = load_cache(sid)
    app = _build_msal_app(cache)
    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(config.SCOPES, account=accounts[0])
    save_cache(sid, cache)

    if result and "access_token" in result:
        return result["access_token"]
    return None


def current_user(sid):
    cache = load_cache(sid)
    app = _build_msal_app(cache)
    accounts = app.get_accounts()
    return accounts[0] if accounts else None


def logout(sid):
    path = _session_path(sid)
    if os.path.exists(path):
        os.remove(path)

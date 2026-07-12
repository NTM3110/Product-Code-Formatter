> **Recommended Raspberry Pi deployment:** use the ARM-native FastAPI + SQLite server and Velopack update feed documented in [RASPBERRY_PI_LOCAL_LICENSE.md](RASPBERRY_PI_LOCAL_LICENSE.md). The Keygen notes below are retained only for legacy deployments.
# Product Code Formatter License Server

Use this folder on the Windows machine that will host the LAN license server.

This scaffold targets the self-hosted Keygen API project from `keygen-sh/keygen-api`. It is intended for a local LAN server with Docker Desktop installed manually.

## Setup

1. Install Docker Desktop and switch to Linux containers.
2. Create a LAN DNS name for the server, for example `license-server.local`.
   - Keygen expects a host name, not a bare IP address.
   - If you do not have LAN DNS, add the host name to each client machine's Windows `hosts` file.
3. Copy `.env.example` to `.env` and fill in the values.
   - Generate strong secrets before production use.
   - Check the current `keygen-sh/keygen-api` self-hosting docs for any newly required environment variables.
   - `KEYGEN_ACCOUNT_ID` is required by Keygen singleplayer mode and must be a UUID. For local testing you can keep the UUID from `.env.example`, then use that same value in the app's `Account` field.
4. Start the server:

```bat
docker compose up -d
```

5. Create a Product Code Formatter product/policy/license with `run_license_server_admin.bat`, or create it manually in the server admin/API.
6. Put allowed app profiles in the license metadata using one of these keys:
   - `allowed_profiles`
   - `profiles`
   - `company_profiles`
   - `allowed_companies`
   - `companies`

Values can be a JSON array or comma/semicolon/newline separated text. The desktop app matches allowed entries against profile keys or labels, for example `son_phuong`, `cao_thanh`, `quang_thinh`, `vietmax`, or Vietnamese labels like `Cao Thành`.

Vietmax now runs through the unified desktop profile `vietmax`. For new Vietmax customers, include only `vietmax`; that single profile enables both mua-vào and bán-ra workflows. Existing licenses that only contain the older phase keys can still open the unified Vietmax profile in the desktop app, but new licenses should use `vietmax`.

## Desktop Client Fields

The Product Code Formatter activation dialog asks for:

- `License server`: for example `http://license-server.local:3000` for private LAN use, or `https://license-server.local` if you put Keygen behind an HTTPS reverse proxy
- `Account`: your Keygen account id or slug
- `License key`: the key created in Keygen

The app validates the key through Keygen, activates the current Windows machine fingerprint, saves the activation locally, and refreshes validation on each startup.

## Notes

- Do not put Keygen admin, environment, or product tokens inside the desktop app.
- The desktop app only stores the license key and activation result in `product_code_config.json`.
- For simple private LAN use, the desktop app allows HTTP for localhost, private LAN IPs, `.local` names, and single-label LAN hostnames. For production LAN use, put Keygen behind HTTPS. A reverse proxy such as Caddy can terminate TLS and forward to the Keygen API container.
- If you need true air-gapped/offline floating leases, evaluate `keygen-sh/keygen-relay` separately; normal LAN activation should use `keygen-sh/keygen-api`.

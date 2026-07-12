> **Legacy Keygen deployment only.** New Raspberry Pi installations must use [RASPBERRY_PI_LOCAL_LICENSE.md](RASPBERRY_PI_LOCAL_LICENSE.md), which uses the ARM-native license service and Velopack bundles.
# Raspberry Pi license and update server

This folder now contains two services for the LAN:

- `keygen`: the existing Keygen CE license API.
- `admin`: a small web admin and release server on port `8080`.

The admin container stores the Keygen admin token. The Windows EXE only knows the public LAN URL and the license key; it never receives the admin token.

## 1. Prepare the Pi

Use a 64-bit Raspberry Pi OS installation if possible. Install Docker Engine using the official Raspberry Pi instructions, then verify Docker with `docker run hello-world`.

Copy this repository's `license_server` directory to the Pi, then:

```bash
cd license_server
cp .env.example .env
mkdir -p releases
```

Edit `.env`:

```dotenv
KEYGEN_HOST=license-server.local
KEYGEN_DOMAIN=license-server.local
KEYGEN_ACCOUNT_ID=6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001
ADMIN_PORT=8080
ADMIN_TOKEN=generate-a-long-random-private-value
```

Keep the existing Keygen secret values private. The account ID above is the fixed account used by the desktop client.

Start both containers:

```bash
docker compose up -d --build
docker compose ps
```

Open `http://<pi-ip>:8080` in a browser and enter the `ADMIN_TOKEN`. The page can list licenses, edit `allowed_profiles` and `allowed_companies`, and delete licenses.

## 2. Publish a desktop release

Build the EXE on the Windows development machine, then copy it to the Pi. The release script creates the manifest and SHA-256 checksum:

```powershell
python license_server\publish_release.py .\deploy\ProductCodeFormatterWeb_v0.2.0.exe --version 0.2.0 --notes "LAN update test"
```

Copy the generated `license_server/releases` directory to the Pi's `license_server/releases` directory. The desktop app checks:

```text
http://<pi-host>:8080/api/update/manifest?platform=windows-x64
```

The app verifies the downloaded EXE against `sha256`, schedules a PowerShell replacement after the current process closes, then starts the new EXE. A repeated export or config save does not trigger an update; only the explicit update button does.

For the EXE to find the Pi, set `PRODUCT_CODE_FORMATTER_UPDATE_SERVER` at build/runtime or use the default `http://license-server.local:8080`. If there is no LAN DNS, add the Pi hostname to each Windows `hosts` file.

## 3. License changes from the admin page

After editing a license, open the desktop app and press **Tải lại license**. The app validates the existing key against Keygen again and stores the new profile/company metadata locally. If the license has been deleted, suspended, or revoked, refresh returns the app to an unactivated state instead of continuing to use stale permissions.

## 4. Optional HTTPS

For a trusted LAN hostname, put Caddy in front of ports `8080` and `3000` and proxy them to the admin and Keygen containers. For a private LAN, HTTP is supported by the current client, but do not expose these ports to the public internet. Keep the admin page behind the token even on a trusted LAN.

## 5. Operational checks

```bash
curl http://127.0.0.1:8080/health
curl http://127.0.0.1:8080/api/update/manifest
docker compose logs --tail=100 admin
docker compose logs --tail=100 keygen
```

Back up the Docker volumes before upgrades:

```bash
docker compose down
docker run --rm -v license_server_keygen_postgres:/from -v "$PWD/backups:/to" alpine tar czf /to/keygen-postgres.tgz -C /from .
docker compose up -d
```

## References

- [Keygen API](https://keygen.sh/docs/api/)
- [Keygen licenses API](https://keygen.sh/docs/api/licenses/)
- [Keygen self-hosting](https://keygen.sh/docs/self-hosting/)
- [Docker on Raspberry Pi OS](https://docs.docker.com/engine/install/raspberry-pi-os/)
- [Caddy reverse proxy](https://caddyserver.com/docs/caddyfile/directives/reverse_proxy)

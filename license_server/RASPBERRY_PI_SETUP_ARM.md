# Raspberry Pi local license and update server

The current ARM-native deployment uses FastAPI, SQLite, and a Velopack update feed. It does not run Keygen.

Use the complete current guide:

- [RASPBERRY_PI_LOCAL_LICENSE.md](RASPBERRY_PI_LOCAL_LICENSE.md)

Important changes from older builds:

- upload the ZIP created by build_release.ps1, not a standalone EXE;
- install ProductCodeFormatter-Setup.exe once on each Windows PC;
- publish every update with a higher version;
- test releases install separately as ProductCodeFormatter Test;
- license data and update packages persist in license-data and releases.

Quick Pi restart after copying updated server files:

~~~bash
cd /home/admin/license_server
chmod +x setup_pi.sh
./setup_pi.sh
curl http://127.0.0.1:8080/health
~~~
# Raspberry Pi license and update server

This is the supported LAN deployment for ProductCodeFormatter. It runs directly on Raspberry Pi with FastAPI, SQLite, and Docker. It does not require Keygen.

Default client order:

1. http://192.168.1.210:8080
2. http://192.168.101.13:8080

The account ID is **6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001**.

## 1. Copy the server to Raspberry Pi

If an older folder is owned by root, fix it once:

~~~bash
ssh admin@192.168.1.210
sudo mkdir -p /home/admin/license_server
sudo chown -R admin:admin /home/admin/license_server
exit
~~~

Then copy from the Windows project root:

~~~powershell
scp -r license_server admin@192.168.1.210:/home/admin/
~~~

On the Pi:

~~~bash
cd /home/admin/license_server
chmod +x setup_pi.sh
./setup_pi.sh
curl http://127.0.0.1:8080/health
~~~

setup_pi.sh automatically backs up an old Keygen .env, creates the local-pi-v1 .env with ADMIN_TOKEN 310902, and starts the container. No manual copy from an example file is required. Keep the default account ID unless the client configuration is changed at the same time.

Open the admin UI at **http://192.168.1.210:8080**.

Persistent data:

- license-data/licenses.sqlite3: licenses and allowed profiles.
- releases/product-code-formatter: Velopack feeds, packages, and Setup files.

Back up both folders.

## 2. License behavior

The first successful activation binds a license to the Windows machine fingerprint. allowed_profiles controls which company profiles appear in the app.

The Pi admin UI can:

- create, edit, suspend, and delete licenses;
- edit allowed profiles with checkboxes;
- add profile codes to the admin catalog;
- publish test or stable update bundles.

After allowed profiles change, the user clicks **Tải lại license** in the app. Adding a new profile to a license does not add its source code to an old client; the client must first update to an app version that contains that profile.

## 3. Development testing

Normal development does not create a release bundle.

Run from source:

~~~powershell
.\run_react_native_app.bat
~~~

Or build a local packaged app without installing/updating:

~~~powershell
.\build_app.ps1 -Version 0.4.0 -Notes "Local test" -Channel dev
.\dist\ProductCodeFormatter\ProductCodeFormatter.exe
~~~

The update button intentionally works only in an app installed by Velopack Setup. Source and direct dist builds remain easy to test and cannot overwrite the installed app.

## 4. Test the updater safely

The test channel installs as **ProductCodeFormatter Test**, separate from the stable app.

Create the first test release:

~~~powershell
.\build_release.ps1 -Version 0.4.0 -Notes "Updater base" -Channel test -SkipRemoteSync
~~~

Install the base version once from **deploy\ProductCodeFormatter-Test-Setup-v0.4.0.exe**.

Create a newer test release:

~~~powershell
.\build_release.ps1 -Version 0.4.1 -Notes "Updater validation" -Channel test
~~~

Upload **deploy\ProductCodeFormatter_test_v0.4.1_bundle.zip** in the Pi admin page.

Open ProductCodeFormatter Test, click **Kiểm tra cập nhật**, then **Cập nhật**. Velopack downloads the full or delta package, closes the app, applies the update, and reopens the installed app.

Do not run stable and test at the same time because both intentionally share the same user configuration in %LOCALAPPDATA%\ProductCodeFormatter.

## 5. Publish a stable release

Every published update needs a higher semantic version. Velopack rejects overwriting an existing version.

~~~powershell
.\build_release.ps1 -Version 0.4.0 -Notes "Velopack installer and Raspberry Pi updates" -Channel stable
~~~

Artifacts:

- deploy\ProductCodeFormatter-Setup.exe (latest installer alias)
- deploy\ProductCodeFormatter-Setup-v0.4.0.exe (versioned installer)
- deploy\ProductCodeFormatter_stable_v0.4.0_bundle.zip

Upload the stable bundle ZIP through the Pi admin UI. New PCs install ProductCodeFormatter-Setup.exe once. Existing Velopack installations update from inside the app.

The server validates before publishing:

- exactly one test or stable feed;
- matching ProductCodeFormatter package ID and channel;
- required full package and Setup;
- package size and SHA-256;
- no path traversal, encrypted files, or unrelated files.

The feed manifest is replaced last, so a rejected or interrupted upload does not replace the working release.

## 6. One-time migration from old portable EXE

Old portable builds cannot safely convert themselves into an installed updater. Run ProductCodeFormatter-Setup.exe once.

Existing configuration, license data, templates, and mappings remain in **%LOCALAPPDATA%\ProductCodeFormatter**.

On the first Velopack launch, the app removes old shortcut aliases and the old hidden portable EXE. The stable installed shortcut remains ProductCodeFormatter.

## 7. Useful checks

Pi:

~~~bash
curl http://127.0.0.1:8080/health
ls -lh releases/product-code-formatter
docker compose -f docker-compose.pi.yml logs --tail=100 admin
~~~

Windows:

~~~powershell
curl.exe http://192.168.1.210:8080/health
curl.exe http://192.168.1.210:8080/updates/product-code-formatter/releases.stable.json
~~~
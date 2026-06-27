from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


OUTPUT = Path("ProductCodeFormatter_License_Setup_Guide.pdf")


SECTIONS = [
    (
        "Server PC Setup",
        [
            "Install Docker Desktop on the Windows PC that will host the license server.",
            "Copy the license_server folder to the server PC.",
            "Open PowerShell in the license_server folder.",
            "Edit .env. For local testing use KEYGEN_HOST=localhost and KEYGEN_DOMAIN=localhost. For LAN use KEYGEN_HOST=license-server.local and KEYGEN_DOMAIN=license-server.local.",
        ],
    ),
    (
        "Start Keygen",
        [
            "docker compose up -d postgres redis",
            "docker compose run --rm -T -e DISABLE_DATABASE_ENVIRONMENT_CHECK=1 -e KEYGEN_EDITION=CE -e KEYGEN_MODE=singleplayer -e KEYGEN_ADMIN_EMAIL=admin@example.local -e KEYGEN_ADMIN_PASSWORD=Admin123456 keygen bundle exec rake keygen:setup",
            "docker compose up -d keygen",
            "docker compose ps",
            "Expected result: keygen shows 0.0.0.0:3000->3000/tcp.",
        ],
    ),
    (
        "Basic License Server UI",
        [
            "The admin UI is separate from the client application and should be installed only on the license server PC.",
            "Run it from the project folder with: .venv\\Scripts\\python.exe license_server_admin.py",
            "Click Kiem tra server to confirm Docker containers are running.",
            "Select the allowed app profiles for the customer.",
            "Click Tao license and copy LICENSE_KEY from the output box.",
        ],
    ),
    (
        "Client PC Activation",
        [
            "If using license-server.local, add this line to C:\\Windows\\System32\\drivers\\etc\\hosts on each client PC: SERVER_IP license-server.local",
            "Open ProductCodeFormatter.exe on the client PC.",
            "License server: enter the license server IP/URL, for example http://SERVER_IP:3000. The app sends the Keygen host header automatically for LAN IPs.",
            "Account: not required in the client app. APP_ACCOUNT_ID is shown only for admin/debug use; do not use LICENSE_ID.",
            "License key: paste APP_LICENSE_KEY / LICENSE_KEY from the license admin output.",
            "Click Kich hoat.",
        ],
    ),
    (
        "Important Notes",
        [
            "Each client machine activates one time. Later app launches revalidate automatically.",
            "Allowed profiles are controlled by license metadata on the server, not typed by the client.",
            "To change allowed profiles, update metadata/create a new license on the server and restart the client app.",
            "For simple private LAN use, Product Code Formatter allows HTTP for localhost, private LAN IPs, .local names, and single-label LAN hostnames. For production LAN use, put Keygen behind HTTPS.",
        ],
    ),
]


def build_pdf():
    doc = SimpleDocTemplate(str(OUTPUT), pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [Paragraph("Product Code Formatter - License Setup Guide", styles["Title"]), Spacer(1, 16)]

    for title, items in SECTIONS:
        story.append(Paragraph(title, styles["Heading2"]))
        for item in items:
            story.append(Paragraph(f"- {item}", styles["BodyText"]))
            story.append(Spacer(1, 4))
        story.append(Spacer(1, 10))

    doc.build(story)


if __name__ == "__main__":
    build_pdf()
    print(OUTPUT.resolve())

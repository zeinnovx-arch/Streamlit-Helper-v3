---
name: Streamlit preview routing
description: Port mapping behavior for non-artifact Streamlit workflows
---

For a standalone Streamlit workflow, map `localPort = 5000` to `externalPort = 5000` in `.replit` so the Replit dev-domain URL with `:5000` reaches the app. Apply the mapping after workflow configuration if the workflow manager removes it during setup.

**Why:** A healthy Streamlit process can still show “couldn't reach this app” when the dev-domain port is not publicly mapped; root-domain traffic may instead land on another service.

**How to apply:** Keep Streamlit bound to `0.0.0.0`, disable proxy-blocking CORS/XSRF checks, and verify the external `:5000` URL returns HTTP 200 before delivery.
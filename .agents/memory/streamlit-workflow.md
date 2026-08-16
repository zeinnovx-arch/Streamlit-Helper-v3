---
name: Streamlit workflow startup
description: Startup flags needed for Streamlit previews in this workspace
---

Streamlit workflows need headless mode, an explicit `0.0.0.0` bind address, and usage-stat collection disabled so the first-run email prompt cannot block port detection.

**Why:** The default onboarding prompt can keep a workflow from opening its configured port even when the Streamlit code is valid.

**How to apply:** Use `streamlit run <entrypoint> --server.port <PORT> --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false` for long-running preview workflows.
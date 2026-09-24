# Casino Ops — User Guide

This guide walks an operator through a casino site's workflow, from setting up a project to reviewing camera events and keeping records. The app can inventory camera equipment and prepare an event-review workflow. **Live CCTV video processing and automatic person identification are not connected in this starter.** The Overview labels the camera event feed as waiting until an approved on-site connector supplies events.

## 1. Start the app and sign in

On the computer that will host this local instance, install Python 3.10 or newer and run:

```sh
python3 -m pip install -r requirements.txt
python3 server.py
```

Open `http://127.0.0.1:8080` on that same computer. On first run, create the administrator account. Use a unique password with at least 14 characters.

The server binds only to the local computer. It is not accessible to other LAN computers or an online main office. Do not expose camera ports or this development app directly to the internet.

## 2. Create a project for the site

Use **＋ Project** to create a separate project for each casino/site. Choose the active site from the selector in the top bar. Cameras, people records, events, and logs are scoped to the selected project.

## 3. Register the camera equipment

Open **Cameras → Add source**. Add each DVR/NVR as a recorder, or add a standalone camera (such as an entrance camera) separately. Register mixed vendors independently. Record the model, IP/host, port, protocol, channel label, and coverage point (General, Entrance, or Exit).

Saving a source only stores its setup details; it does not connect to the device, test RTSP, open video, or prove that the credentials work. Camera features vary by model and firmware.

### Optional ONVIF discovery

An administrator can open **Settings → ONVIF device discovery**, enter only the private CCTV subnet(s) provided by the site's IT team, enable discovery, and select **Find ONVIF devices**. Run this only from a computer with a route to that CCTV network. Discovery sends a scoped ONVIF probe; it is not a general IP/port scan. Results are unverified candidates and must be classified by an administrator. Some vendor-specific devices will not answer ONVIF discovery and need manual registration.

## 4. Add staff and other authorized records

Open **People** and choose **Download import template**. Fill the supplied XLSX columns and use **Import staff / manager data**. The import validates every row before saving; if a row is invalid, correct the reported row and import again. Up to 10,000 rows and 20 MB are accepted. Keep reference numbers as text and dates in `YYYY-MM-DD` format. You can also add individual records manually.

Records can include Staff, Manager, or Watchlist category, position, internal reference, owner, purpose, review date, and status. A photo can be added to a record for authorized staff's manual reference. In this starter, photos are **not compared to camera footage** and do not establish that a person was seen. Do not enter national ID numbers, face templates, or real watchlist data into this starter.

## 5. Use the Overview

The **Overview** is the site's operating summary:

- **Camera sources** shows registered devices; their connection may still be unverified.
- **People register** shows imported or manually added people records and their review status.
- **Camera entry / exit events** is the prepared feed area for a future on-site connector. It is separate from manual staff check-out/return records. When connected, it is designed to show an event reference, Entry/Exit/Unknown direction, source cameras, timestamps, and review state.
- **Waiting for camera movement events** means that no camera events have been received. IP addresses and uploaded photos alone cannot provide movement records.

Current check-out and return actions in People are entered by an operator and timestamped by the app. They are not camera observations and do not prove a boundary crossing. Camera-generated events will remain empty until a compatible, authorized connector is installed and configured on site.

## 6. Review activity and alerts

Open **Movement** to review project-scoped camera movement events when a connector eventually supplies them. Open **Reviews** to manage human-entered review items. The app does not currently ingest live video or determine a person's identity from a face. Do not treat an event reference or a photo on a roster as a confirmed identity match.

The **AI drafts** page can draft neutral staff-alert wording from a reviewer-submitted incident summary when the server has an OpenAI API key configured. A person must review and send any message outside the app; the assistant does not deliver alerts to authorities or staff.

## 7. Connect future automation

The app records minimal event envelopes in a local automation outbox so a future n8n integration can route selected events. The outbox is not a live connection: no event is sent to n8n until an on-site HTTPS dispatcher and server-side secret are configured. Keep webhook credentials out of browser code. Names, photos, video, face data, and government identifiers should not be sent in automation payloads.

## 8. Manage access, logs, and backups

- **Accounts**: create an account for each operator and grant only the permissions they need. The server enforces permissions on API routes.
- **Audit log**: review administrative changes and account actions.
- **Backups**: create and download a snapshot of project records, people, events, audit history, and queued automation output. Store backups on approved encrypted storage; the app does not encrypt backup files.
- **People export**: use **Export project data (.xlsx)** to download the selected project's roster data.

## Current capability boundary

The app stores local project records, inventory, people entries, manually recorded staff check-out/return activity, audit history, backups, and queued automation events. ONVIF discovery is an optional, administrator-run probe. The following require additional deployment and integration work: RTSP/video ingest, camera-generated Entry/Exit event delivery, compatible vendor connectors for mixed camera brands, any identity review workflow, main-office synchronization, and n8n delivery. No face recognition, biometric matching, continuous identity tracking, or automated authority messaging is active.

For a production deployment, confirm the exact DVR/NVR and camera models, CCTV network design, on-site connector, operator permissions, retention rules, secure credential storage, TLS/access controls, encrypted data and backups, and site approval requirements. The Python service deliberately refuses non-loopback network binding in this starter.

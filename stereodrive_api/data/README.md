# Captured data archive

87 completed capture files, filtered to authorized bus 2 devices 3 and 7. Each folder contains capture.pcapng, decoded.json, metadata.json and metadata.md, plus available scrubbed request/result evidence.

manifest.json indexes all files with SHA-256. Metadata records observed command packets, endpoint counts, discoveries and known incident/passive labels. Native requested physical units are only claimed when retained request/result evidence supports them. Filtered archive frame numbers can differ from older notes; correlate UTC and bytes.

Open capture.pcapng in Wireshark. Rebuild with `py -3 stereodrive_api/scripts/build_archive.py --source PATH_TO_LOGS`; this reads saved files only and needs TShark. Logs from the original native experiment predate this API. Archive inclusion is not evidence that the new API was hardware-tested.

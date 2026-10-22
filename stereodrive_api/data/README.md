# Captured data archive

88 captures total: 87 historical captures filtered to bus 2 devices 3 and 7, plus one 2026-10-22 passive capture filtered to bus 2 device 3. Each archived folder contains capture.pcapng, decoded.json and metadata.json; most also have metadata.md and available scrubbed request/result evidence. The new resistance-candidate capture records an unverified varying AF 45 field; it does not establish units or correlation with the displayed value.

manifest.json indexes all files with SHA-256. Metadata records observed command packets, endpoint counts, discoveries and known incident/passive labels. Native requested physical units are only claimed when retained request/result evidence supports them. Filtered archive frame numbers can differ from older notes; correlate UTC and bytes.

Open capture.pcapng in Wireshark. Rebuild with `py -3 stereodrive_api/scripts/build_archive.py --source PATH_TO_LOGS`; this reads saved files only and needs TShark. Logs from the original native experiment predate this API. Archive inclusion is not evidence that the new API was hardware-tested.

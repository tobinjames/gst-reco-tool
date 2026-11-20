# Official GSTN sample files

Sample payloads published on the GST developer portal (developer.gst.gov.in, Taxpayer API → Returns), downloaded 20 Nov 2026.
They contain every section of the format with dummy values, so `test/gstn.test.ts` checks that the Loader reads each section.

| File | API | Version |
|---|---|---|
| gstr2b_api_v4.5_sample.json | GSTR2B - Get All Details (response) | v4.5, Jul 2026 |
| gstr1_api_v5.0_sample.json | GSTR1 - Save GSTR1 data (request) | v5.0, Feb 2026: advances split into B2B / B2C |
| gstr1_api_v4.1_sample.json | GSTR1 - Save GSTR1 data (request) | v4.1, Feb 2025 |

To refresh: POST https://developer.gst.gov.in/apiportal/getAPIVersion with {"apiId": 3185 (2B) or 2001 (GSTR-1)} lists versions;
the file is at /pages/apiportal/data/Returns/<API name>/<version>/response_payload.json (or request_payload.json).

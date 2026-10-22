# CARS24 GMB Delivery Photos (buy-side)

This is a separate test page and Apps Script receiver for customer delivery photos at C1 buy hubs. It does not change the existing QR scanner, review flow, or review-upload receiver.

## What it does

- Uses only C1 listings marked as buy-side in the current QR scanner data.
- Can use the browser's current location to suggest the nearest hub, or staff can search/select a hub manually.
- Accepts up to 10 JPG, PNG, or WebP photos per upload, up to 5 MB each.
- Saves each image in the connected Drive folder `GMB Delivery Pics / StoreCode_HubName` (one hub folder, no date subfolders) and names it `HubName_YYYY-MM-DD_HHmmss_ID.ext`.
- Adds one row per photo to the supplied `GS HUB - CX Pics` spreadsheet, including timestamp, staff ID/name, hub details, filename, Drive URL, type, and size.

## Setup

1. Open a new, separate Google Apps Script project. Do not paste this receiver into the existing review tracking project.
2. Paste `DeliveryPicsReceiver.gs` into the new project and save it.
3. Run `setupDeliveryPhotoLogging()` once from the Apps Script editor and authorize Google Sheets and Drive access. It initializes the blank `Sheet1` header row. The deploying account must have Editor access to the spreadsheet and access to the `GMB Delivery Pics` Drive folder.
4. Deploy it as a web app, executing as the authorized account. To receive uploads from a static GitHub Pages page, the web app must be accessible to anyone with the link. Anyone who obtains the endpoint could submit files, so treat the link as an upload credential; staff ID/name is attribution, not secure sign-in.
5. The current `/exec` URL is configured in `index.template.html`. If you redeploy Apps Script and receive a different URL, replace it there and regenerate `index.html` with the command below.
6. Publish this folder using the existing GitHub Pages setup. Its URL will be `https://ampcars24-cmyk.github.io/GMB_QR/GMB%20Delivery%20Pics/` if the repo publishes its root as the site.

The `GMB Delivery Pics` Drive root folder was created in the connected Drive. On first upload, the receiver creates the relevant hub folder directly under it. Files remain private unless the Drive owner changes sharing. Browser location is requested only after staff presses the location button; permission cannot be silently granted or permanently forced by a website.

## Refresh the embedded hub list

If the C1 buy listings change in `review-router/index.html`, or after configuring the upload URL in `index.template.html`, regenerate the standalone page from the repo root with:

```sh
node "GMB Delivery Pics/build-index.cjs"
```

This writes `GMB Delivery Pics/index.html` from the template and embeds the current C1 buy listing names and coordinates.

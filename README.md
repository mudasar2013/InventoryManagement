# Parts Inventory

Mobile-first Next.js app for looking up appliance parts, bin locations, and stock levels against open service jobs.

## What's included

- Standalone parts search by part number — no job has to be selected first
- Each part shows part number, bin location, quantity on hand, and a stock status badge, derived from quantity so it can never drift out of sync (see `lib/status.ts`)
- Attach a part to a job from the part page or from a job card
- Linking a part never deducts quantity on hand; stock only changes if you receive or consume inventory separately
- Reads parts from multiple sources and merges them by part number (see `lib/getInventory.ts`, `lib/sources/`):
  - a bundled local catalog (always on, used for anything not covered by another source)
  - a SharePoint-hosted Excel workbook (optional — enabled once it's configured, see below)
- Every technician signs in with their own Microsoft account; the SharePoint source reads the workbook using that person's own delegated permissions, not a shared service account

## Run locally

```bash
npm install
cp .env.example .env.local   # fill in real values, see "Sign-in setup" below
npm run dev -- --port 43127
```

Open [http://localhost:43127](http://localhost:43127). The whole app requires signing in with Microsoft — see below before your first run.

## Sign-in setup (required)

The app is gated behind Microsoft sign-in (`proxy.ts`) because the SharePoint source needs a signed-in user's own Graph token to read anything. You need an Azure AD (Microsoft Entra ID) app registration:

1. In the Azure Portal, go to **Microsoft Entra ID → App registrations → New registration**.
2. Add a redirect URI (platform: **Web**): `{NEXTAUTH_URL}/api/auth/callback/azure-ad` — for local dev that's `http://localhost:43127/api/auth/callback/azure-ad`.
3. Under **Certificates & secrets**, create a client secret.
4. Under **API permissions**, add these **delegated** Microsoft Graph permissions and grant admin consent: `openid`, `profile`, `email`, `offline_access`, `Sites.Read.All`.
   - If your admin would rather not grant read access to every site, use `Sites.Selected` instead and grant it access to just the one site with the inventory workbook (see the [Graph docs on Sites.Selected](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)) — swap the scope in `lib/auth/options.ts` if you go this route.
5. Copy the **Application (client) ID**, **Directory (tenant) ID**, and the client secret you created into `.env.local` as `AZURE_AD_CLIENT_ID`, `AZURE_AD_TENANT_ID`, and `AZURE_AD_CLIENT_SECRET`.
6. Set `NEXTAUTH_SECRET` to a random value (`openssl rand -base64 32`) and `NEXTAUTH_URL` to the URL the app runs at.

Without this, the app has nowhere to redirect a signed-out visitor and will 500 on every page — this step isn't optional.

## Connecting the SharePoint inventory source (optional)

Once sign-in works, point the app at the real workbook by setting these in `.env.local`:

```
SHAREPOINT_SITE_HOSTNAME=contoso.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/ServiceOps
SHAREPOINT_FILE_PATH=Shared Documents/Inventory.xlsx
SHAREPOINT_TABLE_NAME=Inventory
```

The workbook needs the inventory data formatted as a named **Excel Table** (select the range → Insert → Table → give it a name in the Table Design tab), not just a plain range — that's what makes reads robust to inserted rows and columns.

The column headers this reads are defined in `lib/sources/sharepoint-excel-source.ts` (`COLUMN_MAP`, near the top of the file) — edit that one object to match the real workbook's header text. `part_number` and `quantity_on_hand` are required; the source throws a clear error naming the headers it looked for if either is missing from the sheet, so a typo there is loud, not a silent empty catalog.

If the SharePoint source is unreachable (network issue, expired session, a renamed column) it's dropped for that request and a banner explains why — the rest of the catalog (local, and anything from other sources) still renders. It doesn't take the whole app down.

## Adding another source

Implement the `InventorySource` interface in `lib/sources/types.ts` (`fetchParts`/`fetchJobs`/`fetchJobParts`), add it to `buildSources()` in `lib/getInventory.ts`, and give it a priority position in that function's `priority` array — that's the order conflicting part numbers get resolved in. `lib/sources/local-source.ts` is the simplest reference implementation; `lib/sources/sharepoint-excel-source.ts` is a fuller one with error handling worth copying from.

## Run tests

```bash
npm run test
```

Covers status derivation, the cross-source merge/conflict logic, and the SharePoint row-mapping logic — all pure functions that don't need a live SharePoint connection.

## Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- Lucide React
- NextAuth.js (Microsoft Entra ID provider) for sign-in
- Microsoft Graph for reading the SharePoint-hosted workbook
